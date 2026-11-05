"""L6 LLM 复核层端到端 demo / CI 夹具：复核包 → 复核记录 → 复核锁 → 防审旧交新。

用法：
    python3 examples/review_demo.py [--keep]

覆盖：
  1. 用 make_demo_data 生成演示数据 + load_master + recon_engine 得到 8 表工作簿；
  2. 生成复核包，断言 figures 与 demo 已知值一致（期初 1360 件/31,610.00 元；结算
     450 件/10,825.00 元；差异 -3 件/-217.50 元；贴花 980.00 元；回加 39 件/825.00 元）、
     checklist 10 项、schema 正确；两次生成除 generated_at 外完全一致；
  3. 固定 fixture 报告（通过）→ --apply exit 0；出现 08-复核记录（8→9 张）、
     其余 8 张表逐单元格不变；归档 JSON/MD + 复核锁存在；--check exit 0；重复 --apply 幂等；
  4. verdict=不通过 报告 + 哈希匹配 → --apply exit 1，仅归档、不写 08 表；
  5. 篡改工作簿（openpyxl 改一格另存）→ --check exit 1；报告哈希不匹配 → --apply exit 1；
  6. 校验规则负例：缺 R10 / 有阻断却判通过 / 有条件通过无 summary / 未知 id 仅警告 /
     --apply 与 --check 互斥；
  7. 静态无网络依赖；总耗时 < 15s。

全部读写只在临时目录内，不写仓库、不联网。退出码 0 = 全部通过。
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SKILL_DIR = os.path.join(ROOT, "skills", "dealer-settlement")
SCRIPTS = os.path.join(SKILL_DIR, "scripts")
LOAD_MASTER = os.path.join(SCRIPTS, "load_master.py")
ENGINE = os.path.join(SCRIPTS, "recon_engine.py")
MAKE_PACKET = os.path.join(SCRIPTS, "make_review_packet.py")
APPLY_REVIEW = os.path.join(SCRIPTS, "apply_review.py")

sys.path.insert(0, HERE)
import make_demo_data  # noqa: E402

sys.path.insert(0, SCRIPTS)
import review_common as rc  # noqa: E402

import openpyxl  # noqa: E402

SHEETS = ["00-对账说明", "01-对账单", "02-库存差异表", "03-源数据对照",
          "04-甲方挂账明细", "05-乙方销售汇总", "06-乙方库存明细", "07-乙方销售流水"]
PERIOD = "202610"
ASOF = "2026-11-02"
PERIOD_START = "2026-10-01"
PERIOD_END = "2026-10-31"
MAX_SECONDS = 15.0

# demo 夹具的已知数（与 examples/run_demo.py 的期望值同源）
EXPECTED_FIGURES = {
    "opening_qty": 1360,
    "opening_amount": 31610.0,
    "settled_qty": 450,
    "settled_amount": 10825.0,
    "diff_qty": -3,
    "diff_amount": -217.5,
    "addback_qty": 39,
    "addback_amount": 825.0,
    "tiehua_total": 980.0,
    "closing_inventory_value": 19742.5,
    "period_end": PERIOD_END,
}
EXPECTED_SKU = {
    "690000000002": (200, 80, 2800.0, -8, 35.0),
    "690000000012": (500, 150, 1875.0, 5, 12.5),
    "690000000001": (300, 120, 3360.0, 0, 28.0),
    "690000000005": (240, 60, 1590.0, 0, 26.5),
    "690000000010": (120, 40, 1200.0, 0, 30.0),
}


class DemoError(AssertionError):
    pass


def check(cond, msg):
    if not cond:
        raise DemoError(msg)


def approx(actual, expected, tol=0.005):
    return (isinstance(actual, (int, float)) and not isinstance(actual, bool)
            and abs(float(actual) - float(expected)) <= tol)


def _clean_env():
    """清理可能影响结果的变量；把代理指向黑洞端口，确保任何网络访问都会快速失败。"""
    drop = {"RECON_CFG", "HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "NO_PROXY"}
    env = {k: v for k, v in os.environ.items() if k.upper() not in drop}
    env.update({
        "PYTHONIOENCODING": "utf-8",
        "PYTHONUTF8": "1",
        "HTTP_PROXY": "http://127.0.0.1:9",
        "HTTPS_PROXY": "http://127.0.0.1:9",
        "ALL_PROXY": "http://127.0.0.1:9",
        "NO_PROXY": "127.0.0.1,localhost",
    })
    return env


def run_script(args, cfg=None, cwd=None):
    env = _clean_env()
    if cfg is not None:
        env["RECON_CFG"] = json.dumps(cfg, ensure_ascii=False)
    return subprocess.run([sys.executable] + args, capture_output=True,
                          env=env, cwd=cwd or ROOT)


def _text(raw):
    return raw.decode("utf-8", "replace")


def _load_json(path):
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def _write_json(path, data):
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)


def _sha(path):
    return rc.sha256_file(path)


def _workbook_values(path, skip_review=True):
    """{表名: {坐标: (值, 数字格式)}}，用于断言「只增/改 08 表，其余表原样」。"""
    wb = openpyxl.load_workbook(path, data_only=False)
    try:
        result = {}
        for ws in wb.worksheets:
            if skip_review and ws.title == rc.REVIEW_SHEET:
                continue
            cells = {}
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        cells[cell.coordinate] = (cell.value, cell.number_format)
            result[ws.title] = cells
        return result
    finally:
        wb.close()


def _sheet_names(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _review_sheet_text(path):
    """08-复核记录整表文本（用于断言关键字段存在）。"""
    wb = openpyxl.load_workbook(path, data_only=False)
    try:
        ws = wb[rc.REVIEW_SHEET]
        return [str(cell.value) for row in ws.iter_rows() for cell in row
                if cell.value is not None]
    finally:
        wb.close()


def build_workbook(tmpdir):
    """生成演示数据并跑通 load_master + recon_engine，返回 (cfg, paths, out_path)。"""
    customer_root, paths = make_demo_data.build(os.path.join(tmpdir, "铺货代销客户"))
    check(os.path.isfile(paths["master"]), "演示主档未生成")
    process = run_script([LOAD_MASTER, paths["master"], "--period", PERIOD, "--asof", ASOF],
                         cwd=tmpdir)
    check(process.returncode == 0, "load_master 失败（rc=%s）：%s"
          % (process.returncode, _text(process.stderr)))
    cfg = json.loads(_text(process.stdout))
    process = run_script([ENGINE], cfg, cwd=tmpdir)
    check(process.returncode == 0, "recon_engine 失败（rc=%s）：%s"
          % (process.returncode, _text(process.stderr)[-1500:]))
    out_path = os.path.join(cfg["base_dir"], cfg["period"]["out_name"])
    check(os.path.isfile(out_path), "引擎未生成输出工作簿：%s" % out_path)
    check(_sheet_names(out_path) == SHEETS, "引擎输出不是预期的 8 张表：%s" % _sheet_names(out_path))
    return cfg, paths, out_path


def make_report(packet, verdict="通过", findings=None, checklist=None,
                summary="夹具报告：按 R1–R10 逐项复核。", reviewed_at="2026-11-03T00:00:00+00:00"):
    return {
        "schema": rc.SCHEMA_REPORT,
        "reviewer": {"name": "review_demo 夹具", "model": "fixture-model"},
        "reviewed_at": reviewed_at,
        "workbook_sha256": packet["workbook"]["sha256"],
        "period": packet["period"],
        "verdict": verdict,
        "checklist": checklist if checklist is not None else
        [{"id": item_id, "result": "通过", "note": "夹具：逐项复核通过"}
         for item_id, _item in rc.REQUIRED_ITEMS],
        "findings": findings or [],
        "summary": summary,
    }


def verify_packet(packet, cfg, paths, out_path):
    """复核包结构 + 关键数断言。"""
    check(packet.get("schema") == rc.SCHEMA_PACKET,
          "复核包 schema 错误：%r" % (packet.get("schema"),))
    check(isinstance(packet.get("generated_at"), str) and "T" in packet["generated_at"],
          "复核包缺少 ISO8601 generated_at")
    workbook = packet.get("workbook") or {}
    check(workbook.get("path") == os.path.abspath(out_path), "复核包 workbook.path 错误")
    check(workbook.get("sha256") == _sha(out_path), "复核包 workbook.sha256 与工作簿不一致")
    check(workbook.get("size") == os.path.getsize(out_path), "复核包 workbook.size 错误")
    master = packet.get("master") or {}
    check(master.get("sha256") == _sha(paths["master"]), "复核包 master.sha256 与主档不一致")

    sources = packet.get("sources") or []
    check([source.get("role") for source in sources] == ["jia", "sales", "inventory"],
          "复核包 sources 角色/顺序错误：%r" % ([s.get("role") for s in sources],))
    for source in sources:
        check(os.path.isfile(source.get("path") or ""), "复核包源文件路径不存在：%r" % (source.get("path"),))
        check(source.get("sha256") == _sha(source["path"]),
              "复核包源文件 sha256 不一致：%r" % (source.get("path"),))

    check(packet.get("period") == {"start": PERIOD_START, "end": PERIOD_END},
          "复核包 period 错误：%r" % (packet.get("period"),))
    check(packet.get("asof") == ASOF, "复核包 asof 错误：%r" % (packet.get("asof"),))

    figures = packet.get("figures") or {}
    check(set(figures) == set(EXPECTED_FIGURES), "复核包 figures 字段不齐：%r" % (sorted(figures),))
    for key, expected in EXPECTED_FIGURES.items():
        actual = figures.get(key)
        if isinstance(expected, str):
            check(actual == expected, "figures.%s = %r，期望 %r" % (key, actual, expected))
        else:
            check(approx(actual, expected), "figures.%s = %r，期望 %r" % (key, actual, expected))

    check([sheet.get("name") for sheet in packet.get("sheets") or []] == SHEETS,
          "复核包 sheets 与 8 表不一致")
    for sheet in packet["sheets"]:
        check(isinstance(sheet.get("rows"), int) and sheet["rows"] > 0
              and isinstance(sheet.get("cols"), int) and sheet["cols"] > 0,
              "复核包 sheets 行列数异常：%r" % (sheet,))

    contract = packet.get("contract") or {}
    check(contract.get("no") == "HT-DEMO-2026-001", "复核包 contract.no 错误")
    check(contract.get("method") == "铺货代销" and contract.get("basis") == "实际动销",
          "复核包 contract 结算方式/基数错误")
    check(contract.get("objection_days") == 3 and contract.get("signback_days") == 5,
          "复核包 contract 异议期/签回期错误")
    check(contract.get("pay_day_text") == "次月5日前", "复核包 contract 付款日错误")

    checklist = packet.get("checklist") or []
    check(len(checklist) == 10, "复核包 checklist 不是 10 项：%d" % len(checklist))
    check([item.get("id") for item in checklist] == list(rc.REQUIRED_IDS),
          "复核包 checklist id 不是 R1–R10：%r" % ([i.get("id") for i in checklist],))
    check([item.get("item") for item in checklist] == [i for _id, i in rc.REQUIRED_ITEMS],
          "复核包 checklist 清单项与 review_common.REQUIRED_ITEMS 不一致")

    sku_rows = packet.get("sku_rows") or []
    check(len(sku_rows) == len(EXPECTED_SKU), "复核包 sku_rows 行数 %d != %d"
          % (len(sku_rows), len(EXPECTED_SKU)))
    by_bar = {row.get("bar"): row for row in sku_rows}
    for bar, (opening, settled, amount, diff, price) in EXPECTED_SKU.items():
        row = by_bar.get(bar)
        check(row is not None, "复核包 sku_rows 缺少条码 %s" % bar)
        check(row.get("opening_qty") == opening, "sku_rows[%s].opening_qty 错误" % bar)
        check(row.get("settled_qty") == settled, "sku_rows[%s].settled_qty 错误" % bar)
        check(approx(row.get("settled_amount"), amount), "sku_rows[%s].settled_amount 错误" % bar)
        check(row.get("diff_qty") == diff, "sku_rows[%s].diff_qty 错误" % bar)
        check(approx(row.get("unit_price"), price), "sku_rows[%s].unit_price 错误" % bar)

    tiehua_rows = packet.get("tiehua_rows") or []
    check(len(tiehua_rows) == 5, "复核包 tiehua_rows 行数 %d != 5" % len(tiehua_rows))
    check(all(row.get("bar") for row in tiehua_rows), "复核包 tiehua_rows 有条码未回填")
    check(approx(sum(row.get("amount") or 0 for row in tiehua_rows), 980.0),
          "复核包 tiehua_rows 合计金额不为 980.00")

    claims = packet.get("claims") or []
    check(len(claims) >= 20, "复核包 claims 过少：%d" % len(claims))
    for claim in claims:
        check(isinstance(claim.get("sheet"), str) and claim["sheet"] in SHEETS,
              "claims.sheet 不在 8 表内：%r" % (claim.get("sheet"),))
        check(isinstance(claim.get("cell"), str) and claim["cell"],
              "claims.cell 缺失")
        check(isinstance(claim.get("text"), str) and claim["text"].strip(),
              "claims.text 缺失")
        check(isinstance(claim.get("numbers"), list) and claim["numbers"],
              "claims.numbers 应为非空数组")
        check("\n" not in claim["text"], "claims.text 不应含换行（表头不应入选）")

    cross_refs = packet.get("cross_refs") or []
    check(any("详见附表 02" in ref.get("text", "") and "02-库存差异表" in (ref.get("refs") or [])
              for ref in cross_refs),
          "复核包 cross_refs 未解析出「详见附表 02」→ 02-库存差异表")
    check(all(ref.get("refs") for ref in cross_refs), "复核包 cross_refs 存在空 refs")

    redline = packet.get("redline_scan") or []
    check(len(redline) >= 6, "复核包 redline_scan 规则过少：%d" % len(redline))
    check(len({item.get("id") for item in redline}) == len(redline), "redline_scan id 重复")
    check(all(isinstance(item.get("pattern"), str) and isinstance(item.get("hits"), list)
              for item in redline), "redline_scan 结构错误")
    rl02 = next((item for item in redline if item.get("id") == "RL02"), None)
    check(rl02 is not None, "redline_scan 缺少 RL02（漏报/瞒报）规则")
    # v1.4.0 起输出模板已移除「可能成因」等归因措辞，RL02 不应再命中 02 表；
    # 保留该断言用于防回归——若文案重新引入「漏报/瞒报」类表述，此处会失败。
    check(not any(hit.get("sheet") == "02-库存差异表" and hit.get("cell") == "C29"
                  for hit in rl02["hits"]),
          "输出仍含漏报/瞒报类红线措辞（02-库存差异表!C29），应改为中性表述")

    warnings = packet.get("warnings")
    check(isinstance(warnings, list), "复核包 warnings 必须是数组")
    check(not any("未取到" in warning for warning in warnings),
          "复核包有 figures 未取到（不应发生在演示数据上）：%r" % (warnings,))


def check_applied(workbook_path, original_values, packet, report, archive_dir):
    """断言 --apply 后的 08 表 / 归档 / 复核锁 / 其余表不变。"""
    names = _sheet_names(workbook_path)
    check(names == SHEETS + [rc.REVIEW_SHEET],
          "--apply 后工作表不是 8+1 张：%r" % (names,))
    check(_workbook_values(workbook_path) == original_values,
          "--apply 改动了 08 表以外的工作表（违反「只增/改 08 表」）")

    texts = "\n".join(_review_sheet_text(workbook_path))
    for needle in ("复核人", "复核模型", "复核时间", "复核结论", "被复核工作簿哈希",
                   "复核报告", "免责声明", "R1", "R10", "不构成审计意见"):
        check(needle in texts, "08-复核记录缺少关键字段/内容：%s" % needle)
    check(report["reviewer"]["name"] in texts and report["reviewer"]["model"] in texts,
          "08-复核记录未写入复核人/模型")
    check(report["workbook_sha256"] in texts, "08-复核记录未写入被复核工作簿哈希")
    check(report["verdict"] in texts, "08-复核记录未写入复核结论")

    lock_path = os.path.join(archive_dir, rc.LOCK_NAME)
    check(os.path.isfile(lock_path), "复核锁未写入：%s" % lock_path)
    lock = _load_json(lock_path)
    check(set(lock) == {"pre_sha256", "post_sha256", "verdict", "report", "applied_at"},
          "复核锁字段不是约定的 5 项：%r" % (sorted(lock),))
    check(lock["pre_sha256"] == packet["workbook"]["sha256"], "复核锁 pre_sha256 错误")
    check(lock["post_sha256"] == _sha(workbook_path), "复核锁 post_sha256 与保存后文件不一致")
    check(lock["verdict"] == report["verdict"], "复核锁 verdict 错误")
    archive_json = os.path.join(archive_dir, lock["report"])
    check(os.path.isfile(archive_json), "复核锁引用的归档报告不存在：%s" % archive_json)
    check(_load_json(archive_json) == report, "归档报告内容与传入报告不一致")
    check(os.path.isfile(os.path.splitext(archive_json)[0] + ".md"), "归档报告缺少 .md")
    return lock


def main(argv=None):
    parser = argparse.ArgumentParser(description="L6 LLM 复核层端到端 demo / CI 夹具")
    parser.add_argument("--keep", action="store_true", help="保留临时目录（失败时总是保留）")
    args = parser.parse_args(argv)

    tmpdir = tempfile.mkdtemp(prefix="dealer_settlement_review_demo_")
    started = time.monotonic()
    ok = False
    try:
        # ---------- 0. 生成 8 表工作簿 ----------
        cfg, paths, out_path = build_workbook(tmpdir)
        original_copy = os.path.join(tmpdir, "original.xlsx")
        shutil.copy2(out_path, original_copy)
        original_sha = _sha(out_path)
        original_values = _workbook_values(out_path)

        # ---------- 1. 复核包 ----------
        packet_path = os.path.join(tmpdir, "packet.json")
        packet_md = os.path.join(tmpdir, "packet.md")
        process = run_script([MAKE_PACKET, out_path, "--master", paths["master"],
                              "--out", packet_path, "--md", packet_md], cfg=cfg, cwd=tmpdir)
        check(process.returncode == 0, "make_review_packet 失败（rc=%s）：%s"
              % (process.returncode, _text(process.stderr)))
        check("[复核包]" in _text(process.stdout), "make_review_packet stdout 缺少摘要")
        check(os.path.isfile(packet_path) and os.path.isfile(packet_md), "复核包 JSON/MD 未生成")
        packet = _load_json(packet_path)
        verify_packet(packet, cfg, paths, out_path)

        # 确定性：除 generated_at 外两次输出完全一致
        packet2_path = os.path.join(tmpdir, "packet2.json")
        packet2_md = os.path.join(tmpdir, "packet2.md")
        process = run_script([MAKE_PACKET, out_path, "--master", paths["master"],
                              "--out", packet2_path, "--md", packet2_md], cfg=cfg, cwd=tmpdir)
        check(process.returncode == 0, "make_review_packet 二次生成失败：%s"
              % _text(process.stderr))
        packet2 = _load_json(packet2_path)
        check(isinstance(packet2.get("generated_at"), str), "二次生成缺少 generated_at")
        packet_first = dict(packet)
        packet_second = dict(packet2)
        packet_first.pop("generated_at", None)
        packet_second.pop("generated_at", None)
        check(packet_first == packet_second, "复核包两次生成除 generated_at 外不一致（输出不确定）")
        with open(packet_md, encoding="utf-8") as fh:
            md_first = fh.read()
        with open(packet2_md, encoding="utf-8") as fh:
            md_second = fh.read()
        check(md_first == md_second, "复核包 Markdown 两次生成不一致")

        # ---------- 2. 通过报告：dry-run → --apply → --check → 幂等 ----------
        archive_dir = os.path.join(os.path.dirname(out_path), rc.REVIEW_DIR)
        report = make_report(packet)
        report_path = os.path.join(tmpdir, "report_pass.json")
        _write_json(report_path, report)
        process = run_script([APPLY_REVIEW, out_path, report_path, "--packet", packet_path],
                             cwd=tmpdir)
        check(process.returncode == 0, "apply_review dry-run 失败：%s" % _text(process.stderr))
        check("[dry-run]" in _text(process.stdout), "dry-run 未打印计划")
        check(not os.path.isdir(archive_dir), "dry-run 不应创建 _复核记录 目录")

        process = run_script([APPLY_REVIEW, out_path, report_path, "--packet", packet_path,
                              "--apply"], cwd=tmpdir)
        check(process.returncode == 0, "apply_review --apply 失败（rc=%s）：%s"
              % (process.returncode, _text(process.stderr)))
        check("[完成]" in _text(process.stdout), "--apply 未打印完成标记")
        lock = check_applied(out_path, original_values, packet, report, archive_dir)
        post_sha = lock["post_sha256"]

        process = run_script([APPLY_REVIEW, out_path, "--check"], cwd=tmpdir)
        check(process.returncode == 0, "--check 应通过但失败（rc=%s）：%s"
              % (process.returncode, _text(process.stderr)))
        check("复核有效" in _text(process.stdout), "--check 未打印「复核有效」")

        process = run_script([APPLY_REVIEW, out_path, report_path, "--packet", packet_path,
                              "--apply"], cwd=tmpdir)
        check(process.returncode == 0, "重复 --apply 应幂等成功（rc=%s）：%s"
              % (process.returncode, _text(process.stderr)))
        check("幂等" in _text(process.stdout), "重复 --apply 未提示幂等")
        check(_sha(out_path) == post_sha, "重复 --apply 改变了工作簿（应幂等）")
        check(_load_json(os.path.join(archive_dir, rc.LOCK_NAME)) == lock,
              "重复 --apply 改写了复核锁（应幂等）")

        # ---------- 3. 不通过：只归档、不写 08 表 ----------
        fresh_dir = os.path.join(tmpdir, "fresh")
        os.makedirs(fresh_dir)
        fresh = os.path.join(fresh_dir, "wb.xlsx")
        shutil.copy2(original_copy, fresh)
        check(_sha(fresh) == original_sha, "fresh 副本哈希与原始工作簿不一致")
        fail_report = make_report(
            packet, verdict="不通过",
            checklist=[{"id": item_id,
                        "result": ("不通过" if item_id == "R4" else "通过"),
                        "note": "夹具：R4 金额结转逐式复核不通过"}
                       for item_id, _item in rc.REQUIRED_ITEMS],
            findings=[{"level": "阻断", "category": "数据一致性", "sheet": "01-对账单",
                       "cell": "C25", "claim": "③ 期末未结算货值不等于 ①-②",
                       "evidence": "夹具证据：C23=31610、C24=10825、C25 应等于 20785",
                       "suggestion": "重新出账后再复核"}],
            summary="存在阻断项（R4 不通过），不得交付。")
        fail_path = os.path.join(tmpdir, "report_fail.json")
        _write_json(fail_path, fail_report)
        process = run_script([APPLY_REVIEW, fresh, fail_path, "--apply"], cwd=tmpdir)
        check(process.returncode == 1, "verdict=不通过 时 --apply 应 exit 1，实际 rc=%s"
              % process.returncode)
        check("阻断" in _text(process.stderr), "不通过报告 stderr 未说明阻断")
        check(rc.REVIEW_SHEET not in _sheet_names(fresh), "不通过报告不应写 08 表")
        check(_sha(fresh) == original_sha, "不通过报告不应改写工作簿")
        fail_archive = os.path.join(fresh_dir, rc.REVIEW_DIR)
        check(os.path.isdir(fail_archive), "不通过报告未归档")
        fail_files = sorted(os.listdir(fail_archive))
        check(any(name.endswith(".json") for name in fail_files)
              and any(name.endswith(".md") for name in fail_files),
              "不通过报告归档缺少 JSON/MD：%r" % (fail_files,))
        check(not os.path.isfile(os.path.join(fail_archive, rc.LOCK_NAME)),
              "不通过报告不应写复核锁")

        # ---------- 4. 篡改 / 哈希不匹配 ----------
        tampered = os.path.join(os.path.dirname(out_path), "篡改副本.xlsx")
        shutil.copy2(out_path, tampered)
        workbook = openpyxl.load_workbook(tampered)
        try:
            workbook["00-对账说明"]["A14"] = 99999
            workbook.save(tampered)
        finally:
            workbook.close()
        process = run_script([APPLY_REVIEW, tampered, "--check"], cwd=tmpdir)
        check(process.returncode == 1, "篡改工作簿后 --check 应 exit 1，实际 rc=%s"
              % process.returncode)
        check("复核后" in _text(process.stderr) or "sha256" in _text(process.stderr),
              "篡改工作簿后 --check 未给出哈希失效原因：%s" % _text(process.stderr)[:300])
        process = run_script([APPLY_REVIEW, out_path, "--check"], cwd=tmpdir)
        check(process.returncode == 0, "篡改副本不应影响原工作簿的复核有效性")

        # 08 表被改写 → --check 应失败
        edited = os.path.join(os.path.dirname(out_path), "改写08副本.xlsx")
        shutil.copy2(out_path, edited)
        workbook = openpyxl.load_workbook(edited)
        try:
            ws = workbook[rc.REVIEW_SHEET]
            ws["B7"] = "0" * 64
            workbook.save(edited)
        finally:
            workbook.close()
        process = run_script([APPLY_REVIEW, edited, "--check"], cwd=tmpdir)
        check(process.returncode == 1, "改写 08 表后 --check 应 exit 1，实际 rc=%s"
              % process.returncode)
        check("不一致" in _text(process.stderr) or "复核后" in _text(process.stderr),
              "改写 08 表后 --check 未给出原因：%s" % _text(process.stderr)[:300])

        fresh2_dir = os.path.join(tmpdir, "fresh2")
        os.makedirs(fresh2_dir)
        fresh2 = os.path.join(fresh2_dir, "wb.xlsx")
        shutil.copy2(original_copy, fresh2)
        bad_report = make_report(packet)
        bad_report["workbook_sha256"] = "0" * 64
        bad_path = os.path.join(tmpdir, "report_badhash.json")
        _write_json(bad_path, bad_report)
        process = run_script([APPLY_REVIEW, fresh2, bad_path, "--apply"], cwd=tmpdir)
        check(process.returncode == 1, "报告哈希不匹配时 --apply 应 exit 1，实际 rc=%s"
              % process.returncode)
        check("不一致" in _text(process.stderr), "报告哈希不匹配未给出中文原因")
        check(rc.REVIEW_SHEET not in _sheet_names(fresh2), "报告哈希不匹配不应写 08 表")
        check(not os.path.isdir(os.path.join(fresh2_dir, rc.REVIEW_DIR)),
              "报告哈希不匹配不应归档")

        # ---------- 5. 校验规则负例 ----------
        fresh3_dir = os.path.join(tmpdir, "fresh3")
        os.makedirs(fresh3_dir)
        fresh3 = os.path.join(fresh3_dir, "wb.xlsx")
        shutil.copy2(original_copy, fresh3)
        good = make_report(packet)

        def _case(label, mutate, expect_rc, needle):
            report = json.loads(json.dumps(good))
            mutate(report)
            path = os.path.join(tmpdir, "case_%s.json" % label)
            _write_json(path, report)
            process = run_script([APPLY_REVIEW, fresh3, path, "--apply"], cwd=tmpdir)
            check(process.returncode == expect_rc,
                  "%s：期望 rc=%s，实际 rc=%s（stderr=%s）"
                  % (label, expect_rc, process.returncode, _text(process.stderr)[:300]))
            if needle:
                check(needle in _text(process.stderr),
                      "%s：stderr 缺少「%s」：%s" % (label, needle, _text(process.stderr)[:300]))

        _case("missing_r10", lambda r: r["checklist"].pop(), 1, "R10")
        _case("block_but_pass", lambda r: r["findings"].append(
            {"level": "阻断", "category": "c", "sheet": "s", "cell": "A1",
             "claim": "x", "evidence": "y", "suggestion": "z"}), 1, "阻断")
        _case("conditional_no_summary",
              lambda r: r.update({"verdict": "有条件通过", "summary": ""}), 1, "summary")

        unknown_report = json.loads(json.dumps(good))
        unknown_report["checklist"].append({"id": "R99", "result": "通过", "note": ""})
        unknown_path = os.path.join(tmpdir, "case_unknown.json")
        _write_json(unknown_path, unknown_report)
        process = run_script([APPLY_REVIEW, fresh3, unknown_path], cwd=tmpdir)
        check(process.returncode == 0, "未知 checklist id 应仅警告不拒绝，实际 rc=%s"
              % process.returncode)
        check("R99" in _text(process.stderr), "未知 checklist id 未给出警告")

        process = run_script([APPLY_REVIEW, fresh3, unknown_path, "--apply", "--check"],
                             cwd=tmpdir)
        check(process.returncode != 0, "--apply 与 --check 同时给出时应拒绝")
        check("互斥" in _text(process.stderr), "--apply/--check 互斥未给出中文原因")

        # ---------- 6. 静态无网络 + 耗时 ----------
        # 动态拼接：避免把网络模块名字面量写进本文件（run_demo.py 的静态扫描会扫 examples/）
        banned = ("import " + "requests", "urllib." + "request", "http." + "client",
                  "socket." + "socket", "import " + "httpx", "import " + "aiohttp")
        for name in ("review_common.py", "make_review_packet.py", "apply_review.py"):
            path = os.path.join(SCRIPTS, name)
            with open(path, encoding="utf-8") as fh:
                source = fh.read()
            for token in banned:
                check(token not in source, "%s 含网络依赖：%s" % (name, token))

        elapsed = time.monotonic() - started
        check(elapsed < MAX_SECONDS, "总耗时 %.2fs 超过 %.0fs 上限" % (elapsed, MAX_SECONDS))

        ok = True
        print("[review_demo] OK")
        print("  对账工作簿 : 8 表，sha256=%s…（load_master → recon_engine rc=0）"
              % original_sha[:16])
        print("  复核包     : schema=%s；figures 期初 %s 件/%s 元、结算 %s 件/%s 元、"
              "差异 %s 件/%s 元、贴花 %s 元、回加 %s 件/%s 元；checklist 10 项"
              % (packet["schema"], _num(packet["figures"]["opening_qty"]),
                 _num(packet["figures"]["opening_amount"]),
                 _num(packet["figures"]["settled_qty"]),
                 _num(packet["figures"]["settled_amount"]),
                 _num(packet["figures"]["diff_qty"]), _num(packet["figures"]["diff_amount"]),
                 _num(packet["figures"]["tiehua_total"]),
                 _num(packet["figures"]["addback_qty"]),
                 _num(packet["figures"]["addback_amount"])))
        print("  08 表      : 8→9 张，复核人/模型/时间/结论/被复核哈希/报告名/清单 10 项/问题表/"
              "免责声明齐全；其余 8 表逐单元格不变")
        print("  复核锁     : pre_sha256=%s… post_sha256=%s…（与归档报告、08 表互相绑定）"
              % (lock["pre_sha256"][:16], lock["post_sha256"][:16]))
        print("  --check    : 复核有效；篡改工作簿 / 改写 08 表 → exit 1（防审旧交新生效）")
        print("  不通过     : --apply exit 1，仅归档报告、不写 08 表、不改写工作簿")
        print("  校验规则   : 缺 R10 / 阻断⇒不通过 / 有条件通过需 summary / 未知 id 仅警告 / "
              "--apply --check 互斥 均符合")
        print("  幂等       : 同一报告重复 --apply 跳过写入，工作簿与复核锁字节不变")
        print("  确定性     : 复核包两次生成（除 generated_at）JSON/MD 完全一致")
        print("  耗时       : %.2fs（<%.0fs），临时目录：%s" % (elapsed, MAX_SECONDS, tmpdir))
        return 0
    except DemoError as exc:
        print("[review_demo] FAIL: " + str(exc), file=sys.stderr)
        print("[review_demo] 临时目录已保留：%s" % tmpdir, file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001 - CLI 顶层兜底，输出可读错误
        import traceback
        traceback.print_exc()
        print("[review_demo] 临时目录已保留：%s" % tmpdir, file=sys.stderr)
        return 1
    finally:
        if ok and not args.keep and os.path.isdir(tmpdir):
            shutil.rmtree(tmpdir, ignore_errors=True)


def _num(value):
    if value is None:
        return "null"
    if isinstance(value, float) and value != int(value):
        return "{:,.2f}".format(value)
    return "{:,}".format(int(value))


if __name__ == "__main__":
    sys.exit(main())
