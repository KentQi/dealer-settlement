"""端到端演示 / CI 夹具：结构漂移 → 自动识别 → 用户确认 → 引擎按映射重跑 → 主档修复。

用法：
    python3 examples/schema_demo.py [--keep]

覆盖的验收链路（全部在系统临时目录，无网络、确定性）：
  1. 用 make_demo_data 生成标准演示数据，跑通基线（load_master → recon_engine → 自检通过）；
  2. 复制演示数据并制造漂移：
     - 挂账导出：工作表改名、列名改错、列顺序打乱、多一列「备注」；
     - 库存导出：主工作表改名、列名变体、同一文件里多一张含数据表（关键词无法定位）；
     - 主档副本：合同表改名、合同编号列改名、结算方式/结算基数换序、工作表顺序打乱；
  3. 断言：无映射时引擎按旧行为中止（中文原因，stdout 0 字节）；
     inspect_schema 识别出 SHEET_NAME_MISMATCH / FIELD_NAME_MISMATCH / FIELD_ORDER_DIFF /
     MULTI_CANDIDATE_SHEET / EXTRA_COLUMN，并给出高置信字段建议（低置信项进 notes 待确认）；
  4. 生成映射 → apply_mapping dry-run → 负例（未确认 / 指向不存在列）→ --write --confirmed-by fixture；
     断言三份源导出 sha256 始终不变（永不修改源导出）；
  5. 引擎带映射重跑：自检通过，输出与基线逐单元格一致（仅「实际工作表名」说明文字按映射变化）；
  6. 主档副本：--fix-master --fix-order dry-run → --write 修复；_备份/ 有毫秒时间戳备份且备份
     内容 == 修复前主档；load_master 恢复合同编号；引擎重跑通过；
  7. 确定性：inspect_schema 连跑两次输出完全一致；总耗时 <20s。

退出码 0 = 全部通过；失败时保留临时目录并打印路径。
"""
import argparse
import copy
import glob
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCRIPTS = os.path.join(ROOT, "skills", "dealer-settlement", "scripts")
LOAD_MASTER = os.path.join(SCRIPTS, "load_master.py")
ENGINE = os.path.join(SCRIPTS, "recon_engine.py")
INSPECT = os.path.join(SCRIPTS, "inspect_schema.py")
APPLY = os.path.join(SCRIPTS, "apply_mapping.py")

if HERE not in sys.path:
    sys.path.insert(0, HERE)
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)

import make_demo_data  # noqa: E402
from make_master import SHEETS  # noqa: E402  （主档规范表头/顺序的单一事实来源）
from openpyxl.utils import get_column_letter  # noqa: E402
import openpyxl  # noqa: E402

PERIOD = make_demo_data.PERIOD
ASOF = make_demo_data.ASOF
JIA_SHEET_DRIFT = "未结算明细"
INV_SHEET_DRIFT = "结存明细"
# 引擎报表里如实写入「实际工作表名」；比对基线时把这些说明文字归一化回规范名
SHEET_TEXT_RENAMES = {INV_SHEET_DRIFT: "库存明细", JIA_SHEET_DRIFT: "挂账"}


class SchemaDemoError(AssertionError):
    pass


def check(cond, msg):
    if not cond:
        raise SchemaDemoError(msg)


def _clean_env():
    """清理可能影响结果的变量；代理指向黑洞端口，确保任何网络访问快速失败。"""
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


def _text(data):
    return data.decode("utf-8", "replace")


def sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def read_json(path):
    with open(path, "r", encoding="utf-8-sig") as fh:
        return json.load(fh)


def write_json(path, data):
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
        fh.write("\n")


def _cell_map(path):
    workbook = openpyxl.load_workbook(path, data_only=True)
    try:
        return {(ws.title, cell.row, cell.column): cell.value
                for ws in workbook.worksheets
                for row in ws.iter_rows()
                for cell in row}
    finally:
        workbook.close()


def _normalized_cell_map(path):
    mapping = _cell_map(path)
    normalized = {}
    for key, value in mapping.items():
        if isinstance(value, str):
            for actual, canonical in SHEET_TEXT_RENAMES.items():
                value = value.replace(actual, canonical)
        normalized[key] = value
    return normalized


def load_cfg(master_path):
    processed = run_script([LOAD_MASTER, master_path, "--period", PERIOD, "--asof", ASOF],
                           cwd=os.path.dirname(master_path))
    check(processed.returncode == 0,
          "load_master 失败（rc=%d）：%s" % (processed.returncode, _text(processed.stderr)))
    check(processed.stdout.strip(), "load_master 未输出 JSON")
    return json.loads(_text(processed.stdout))


def run_engine(cfg):
    return run_script([ENGINE], cfg=cfg)


# ---------------- 制造漂移（只改副本，不动基线） ----------------
def drift_jia(path):
    """挂账导出：工作表改名 + 列名改错 + 列顺序打乱 + 多一列备注。"""
    workbook = openpyxl.load_workbook(path)
    worksheet = workbook.active
    rows = [[cell.value for cell in row] for row in worksheet.iter_rows()]
    workbook.close()
    header = rows[1] if len(rows) > 1 else []
    check(header[:5] == ["商品名称", "单价", "未结算金额", "未结算数量", "初始数量"],
          "演示挂账表头与预期不符：%r" % (header[:5],))
    title, data = rows[0][0], rows[2:]
    source_order = [0, 3, 1, 2, 4]  # 商品、数量、供货价、金额、期初数量
    new_headers = ["商品", "数量", "供货价", "金额", "期初数量", "备注"]
    drifted = openpyxl.Workbook()
    sheet = drifted.active
    sheet.title = JIA_SHEET_DRIFT
    sheet.cell(row=1, column=1, value=title)
    for index, name in enumerate(new_headers, 1):
        sheet.cell(row=2, column=index, value=name)
    for row_index, row in enumerate(data, 3):
        for column, source in enumerate(source_order, 1):
            sheet.cell(row=row_index, column=column,
                       value=row[source] if source < len(row) else None)
        sheet.cell(row=row_index, column=6, value="")
    drifted.save(path)


def drift_inventory(path):
    """库存导出：主表改名 + 列名变体 + 同文件多一张含数据表（引擎无法按关键词定位）。"""
    workbook = openpyxl.load_workbook(path)
    main_name = next(name for name in workbook.sheetnames if "库存" in name)
    worksheet = workbook[main_name]
    worksheet.title = INV_SHEET_DRIFT
    renames = {"商品条码": "条码", "库存数量": "库存", "销售价": "零售价"}
    for cell in worksheet[1]:
        if cell.value in renames:
            cell.value = renames[cell.value]
    decoy = workbook.create_sheet("商品档案")
    decoy.append(["档案编号", "商品名称", "规格", "创建时间"])
    decoy.append(["DA001", "示例面膜A（盒装5片）", "盒装5片", "2026-10-01"])
    decoy.append(["DA002", "示例颈膜（盒装5片）", "盒装5片", "2026-10-01"])
    workbook.save(path)


def drift_master(path):
    """主档副本：合同表改名 + 合同编号列改名 + 结算方式/基数换序 + 工作表顺序打乱。"""
    workbook = openpyxl.load_workbook(path)
    worksheet = workbook["合同"]
    worksheet.title = "合同信息"
    header = [cell.value for cell in worksheet[1]]
    worksheet.cell(row=1, column=header.index("合同编号") + 1, value="合同号")
    first = header.index("结算方式") + 1
    second = header.index("结算基数") + 1
    for row in range(1, worksheet.max_row + 1):
        left = worksheet.cell(row=row, column=first).value
        right = worksheet.cell(row=row, column=second).value
        worksheet.cell(row=row, column=first, value=right)
        worksheet.cell(row=row, column=second, value=left)
    workbook.move_sheet("合同信息",
                        offset=len(workbook.sheetnames) - 1
                        - workbook.sheetnames.index("合同信息"))
    workbook.save(path)


# ---------------- 断言辅助 ----------------
def issue_codes(file_report):
    return [issue["code"] for issue in file_report["issues"]]


def find_file(report, role):
    for item in report["files"]:
        if item["role"] == role:
            return item
    raise SchemaDemoError("探查报告缺少 %s 文件" % role)


def expect_codes(file_report, codes, label):
    found = issue_codes(file_report)
    missing = [code for code in codes if code not in found]
    check(not missing, "%s 未识别出 %s（实际：%s）" % (label, missing, found))


def main(argv=None):
    parser = argparse.ArgumentParser(description="结构漂移识别 → 映射确认 → 主档修复 端到端演示")
    parser.add_argument("--keep", action="store_true", help="保留临时目录（失败时总是保留）")
    args = parser.parse_args(argv)

    started = os.times()
    tmpdir = tempfile.mkdtemp(prefix="dealer_schema_demo_")
    ok = False
    try:
        # ---------- 1. 基线数据：无映射时必须完全按旧行为跑通 ----------
        baseline_root, baseline_paths = make_demo_data.build(os.path.join(tmpdir, "铺货代销客户"))
        cfg_base = load_cfg(baseline_paths["master"])
        check("mapping" not in cfg_base, "基线数据不应带 _数据映射.json")
        processed = run_engine(cfg_base)
        check(processed.returncode == 0,
              "基线引擎失败（rc=%d）：%s" % (processed.returncode, _text(processed.stderr)[-1500:]))
        check("自检通过（残差 0.00）" in _text(processed.stderr), "基线引擎未自检通过")
        baseline_out = os.path.join(cfg_base["base_dir"], cfg_base["period"]["out_name"])
        check(os.path.isfile(baseline_out), "基线引擎未生成输出文件")

        # ---------- 2. 复制并制造三份导出漂移 ----------
        drift_root = os.path.join(tmpdir, "导出漂移客户")
        shutil.copytree(baseline_root, drift_root)
        drift_paths = {
            "master": os.path.join(drift_root, "_主档.xlsx"),
            "jia": os.path.join(drift_root, PERIOD, os.path.basename(baseline_paths["jia"])),
            "sales": os.path.join(drift_root, PERIOD, os.path.basename(baseline_paths["sales"])),
            "inventory": os.path.join(drift_root, PERIOD,
                                      os.path.basename(baseline_paths["inventory"])),
        }
        export_shas_before = {key: sha256_file(drift_paths[key])
                              for key in ("jia", "sales", "inventory")}
        drift_jia(drift_paths["jia"])
        drift_inventory(drift_paths["inventory"])
        check(sha256_file(drift_paths["jia"]) != export_shas_before["jia"], "挂账文件漂移未生效")
        check(sha256_file(drift_paths["inventory"]) != export_shas_before["inventory"],
              "库存文件漂移未生效")
        drift_shas = {key: sha256_file(drift_paths[key])
                      for key in ("jia", "sales", "inventory")}

        # ---------- 3. 无映射：引擎按旧行为中止（中文原因 + stdout 0 字节）----------
        cfg_drift = load_cfg(drift_paths["master"])
        check("mapping" not in cfg_drift, "无 _数据映射.json 时不应有 mapping 配置")
        processed = run_engine(cfg_drift)
        stderr = _text(processed.stderr)
        check(processed.returncode != 0, "无映射时引擎应中止，实际 rc=0")
        check(processed.stdout == b"", "无映射失败时 stdout 必须 0 字节，实际 %r"
              % processed.stdout[:120])
        check("数据表未找到" in stderr or "未找到" in stderr,
              "引擎中止原因不是预期的中文结构错误：%s" % stderr[-500:])
        check("甲方挂账" in stderr or "库存" in stderr, "引擎中止未指明是哪张表/哪份文件")

        # ---------- 4. inspect_schema：识别漂移并给出带置信度的映射建议 ----------
        inspect_json = os.path.join(tmpdir, "探查.json")
        inspect_md = os.path.join(tmpdir, "探查.md")
        inspect_json_2 = os.path.join(tmpdir, "探查_重跑.json")
        processed = run_script([INSPECT, os.path.join(drift_root, PERIOD),
                                "--out", inspect_json, "--md", inspect_md])
        check(processed.returncode == 0,
              "inspect_schema 失败（rc=%d）：%s" % (processed.returncode, _text(processed.stderr)))
        check(os.path.isfile(inspect_json) and os.path.isfile(inspect_md),
              "inspect_schema 未生成 JSON/MD 报告")
        drift_report = read_json(inspect_json)
        jia_report = find_file(drift_report, "jia")
        inventory_report = find_file(drift_report, "inventory")
        expect_codes(jia_report,
                     ["SHEET_NAME_MISMATCH", "FIELD_NAME_MISMATCH", "FIELD_ORDER_DIFF",
                      "EXTRA_COLUMN"],
                     "挂账文件")
        expect_codes(inventory_report,
                     ["SHEET_NAME_MISMATCH", "FIELD_NAME_MISMATCH", "MULTI_CANDIDATE_SHEET"],
                     "库存文件")
        jia_fields = jia_report["fields"]
        check(jia_fields["price"]["found"] == "供货价"
              and jia_fields["price"]["confidence"] == "高",
              "挂账·单价未给出高置信建议：%r" % (jia_fields["price"],))
        check(jia_fields["amount"]["found"] == "金额"
              and jia_fields["amount"]["confidence"] == "高",
              "挂账·未结算金额未给出高置信建议：%r" % (jia_fields["amount"],))
        check(jia_fields["init"]["confidence"] in ("中", "低")
              and jia_fields["init"]["found"] == "期初数量",
              "挂账·初始数量应按中/低置信交用户确认：%r" % (jia_fields["init"],))
        proposed = drift_report["proposed_mapping"]
        check(proposed.get("schema") == "dealer-settlement/data-mapping@1",
              "proposed_mapping.schema 不规范")
        check(proposed["sources"]["jia"]["sheet"] == JIA_SHEET_DRIFT,
              "proposed_mapping 未建议挂账工作表名")
        check(proposed["sources"]["jia"]["header_row"] == 2,
              "proposed_mapping 未记录挂账表头行（应为第 2 行）")
        check(proposed["sources"]["inventory"]["sheet"] == INV_SHEET_DRIFT,
              "proposed_mapping 未建议库存工作表名")
        check(proposed["sources"]["inventory"].get("next_month", {}).get("sheet") == "11月销售",
              "proposed_mapping 未建议「次月已销」工作表")
        check(any("需人工确认" in note for note in proposed["notes"]),
              "proposed_mapping.notes 未标出需人工确认字段")
        # 确定性：同样输入连跑两次，探查 JSON 完全一致
        processed = run_script([INSPECT, os.path.join(drift_root, PERIOD),
                                "--out", inspect_json_2])
        check(processed.returncode == 0, "inspect_schema 第二次运行失败")
        with open(inspect_json, "rb") as fh:
            first_bytes = fh.read()
        with open(inspect_json_2, "rb") as fh:
            second_bytes = fh.read()
        check(first_bytes == second_bytes, "inspect_schema 两次输出不一致（不确定）")

        # ---------- 5. 映射确认：dry-run → 负例 → --write --confirmed-by ----------
        mapping_path = os.path.join(tmpdir, "映射建议.json")
        write_json(mapping_path, proposed)
        processed = run_script([APPLY, drift_root, mapping_path])
        check(processed.returncode == 0,
              "apply_mapping dry-run 失败：%s" % _text(processed.stderr))
        check("dry-run" in _text(processed.stdout) and "未写任何文件" in _text(processed.stdout),
              "apply_mapping dry-run 未声明未写文件")
        check(not os.path.isfile(os.path.join(drift_root, "_数据映射.json")),
              "dry-run 不应写 _数据映射.json")

        processed = run_script([APPLY, drift_root, mapping_path, "--write"])
        check(processed.returncode != 0, "未 --confirmed-by 时 --write 应失败")
        check("必须经用户确认" in _text(processed.stderr),
              "未确认时报错未提示「必须经用户确认」：%s" % _text(processed.stderr))
        check(not os.path.isfile(os.path.join(drift_root, "_数据映射.json")),
              "未确认时不应写 _数据映射.json")

        bad_mapping = copy.deepcopy(proposed)
        bad_mapping["sources"]["jia"]["columns"]["price"] = "不存在的列XYZ"
        bad_path = os.path.join(tmpdir, "映射建议_坏.json")
        write_json(bad_path, bad_mapping)
        processed = run_script([APPLY, drift_root, bad_path])
        check(processed.returncode != 0, "映射指向不存在列时 apply_mapping 应失败")
        check("不存在" in _text(processed.stderr),
              "不存在列时报错未说明原因：%s" % _text(processed.stderr))
        check(not os.path.isfile(os.path.join(drift_root, "_数据映射.json")),
              "校验失败时不应写 _数据映射.json")

        processed = run_script([APPLY, drift_root, mapping_path, "--write",
                                "--confirmed-by", "fixture"])
        check(processed.returncode == 0,
              "apply_mapping --write 失败：%s" % _text(processed.stderr))
        mapping_file = os.path.join(drift_root, "_数据映射.json")
        check(os.path.isfile(mapping_file), "--write 未生成 _数据映射.json")
        confirmed = read_json(mapping_file)
        check(confirmed.get("schema") == "dealer-settlement/data-mapping@1"
              and confirmed.get("confirmed_by") == "fixture"
              and confirmed.get("confirmed_at"),
              "落盘映射缺少 schema/confirmed_by/confirmed_at：%r" % confirmed)
        archives = glob.glob(os.path.join(drift_root, "_确认记录", "映射_*.json"))
        check(len(archives) == 1, "_确认记录/ 未归档映射副本：%r" % archives)
        for key in ("jia", "sales", "inventory"):
            check(sha256_file(drift_paths[key]) == drift_shas[key],
                  "确认落盘阶段修改了源导出 %s（违反只读约束）" % key)

        # ---------- 6. 引擎带映射重跑：自检通过，输出与基线一致 ----------
        cfg_mapped = load_cfg(drift_paths["master"])
        check((cfg_mapped.get("mapping") or {}).get("confirmed_by") == "fixture",
              "load_master 未把 _数据映射.json 读入 RECON_CFG['mapping']")
        cfg_mapped = json.loads(json.dumps(cfg_mapped))
        cfg_mapped["out_file"] = os.path.join(cfg_mapped["base_dir"], "映射后对账.xlsx")
        processed = run_engine(cfg_mapped)
        check(processed.returncode == 0,
              "带映射重跑失败（rc=%d）：%s" % (processed.returncode,
                                              _text(processed.stderr)[-2000:]))
        check("自检通过（残差 0.00）" in _text(processed.stderr), "带映射重跑未自检通过")
        mapped_out = cfg_mapped["out_file"]
        check(os.path.isfile(mapped_out), "带映射重跑未生成输出文件")
        check(_normalized_cell_map(mapped_out) == _cell_map(baseline_out),
              "带映射输出与基线不一致（除实际工作表名说明外应逐单元格相同）")

        # 源导出在引擎读取后仍未被修改
        for key in ("jia", "sales", "inventory"):
            check(sha256_file(drift_paths[key]) == drift_shas[key],
                  "引擎/映射流程修改了源导出 %s（违反只读约束）" % key)

        # ---------- 6b. 映射指向不存在表/列：引擎必须中文中止并指明是哪一项 ----------
        bad_root = os.path.join(tmpdir, "映射不匹配客户")
        shutil.copytree(drift_root, bad_root)
        bad_master = os.path.join(bad_root, "_主档.xlsx")
        bad_mapping_file = os.path.join(bad_root, "_数据映射.json")
        good_mapping = read_json(os.path.join(drift_root, "_数据映射.json"))
        column_bad = copy.deepcopy(good_mapping)
        column_bad["sources"]["jia"]["columns"]["price"] = "不存在的列XYZ"
        write_json(bad_mapping_file, column_bad)
        cfg_bad = load_cfg(bad_master)
        check(cfg_bad.get("mapping"), "坏映射未进入 RECON_CFG['mapping']")
        processed = run_engine(cfg_bad)
        check(processed.returncode != 0, "映射列不存在时引擎应中止")
        check(processed.stdout == b"", "映射列错误时 stdout 必须 0 字节")
        check("[映射错误]" in _text(processed.stderr)
              and "sources.jia.columns.price" in _text(processed.stderr)
              and "不存在的列XYZ" in _text(processed.stderr),
              "引擎未指明不存在的映射列：%s" % _text(processed.stderr)[-500:])

        sheet_bad = copy.deepcopy(good_mapping)
        sheet_bad["sources"]["inventory"]["sheet"] = "不存在的表"
        write_json(bad_mapping_file, sheet_bad)
        cfg_bad = load_cfg(bad_master)
        processed = run_engine(cfg_bad)
        check(processed.returncode != 0, "映射工作表不存在时引擎应中止")
        check(processed.stdout == b"", "映射表错误时 stdout 必须 0 字节")
        check("[映射错误]" in _text(processed.stderr)
              and "sources.inventory.sheet" in _text(processed.stderr)
              and "不存在的表" in _text(processed.stderr),
              "引擎未指明不存在的映射工作表：%s" % _text(processed.stderr)[-500:])

        # ---------- 7. 主档漂移风险：不修复则引擎中止；修复后恢复 ----------
        master_root = os.path.join(tmpdir, "主档漂移客户")
        shutil.copytree(baseline_root, master_root)
        master_path = os.path.join(master_root, "_主档.xlsx")
        master_sha_before = sha256_file(master_path)
        drift_master(master_path)
        drifted_master_sha = sha256_file(master_path)
        check(drifted_master_sha != master_sha_before, "主档漂移未生效")
        # 记录漂移后「合同信息」到的列序/列宽：修复应把值/格式/列宽随列搬回规范顺序
        workbook = openpyxl.load_workbook(master_path)
        try:
            drifted_contract_header = [cell.value for cell in workbook["合同信息"][1]]
            drifted_contract_widths = [
                workbook["合同信息"].column_dimensions[get_column_letter(index)].width
                for index in range(1, len(drifted_contract_header) + 1)]
        finally:
            workbook.close()

        cfg_master = load_cfg(master_path)
        check(cfg_master["contract"]["no"] == "",
              "主档合同表被改名后 load_master 仍读到合同编号（漂移未生效）")
        processed = run_engine(cfg_master)
        check(processed.returncode != 0, "主档漂移但未修复时引擎应中止")
        check(processed.stdout == b"", "主档漂移失败时 stdout 必须 0 字节")
        check("合同编号" in _text(processed.stderr),
              "主档漂移中止原因未指明合同编号：%s" % _text(processed.stderr)[-400:])

        master_report_path = os.path.join(tmpdir, "主档探查.json")
        processed = run_script([INSPECT, master_root, "--out", master_report_path])
        check(processed.returncode == 0,
              "主档 inspect_schema 失败：%s" % _text(processed.stderr))
        master_report = read_json(master_report_path)
        master_file = find_file(master_report, "master")
        expect_codes(master_file,
                     ["SHEET_NAME_MISMATCH", "FIELD_NAME_MISMATCH", "FIELD_ORDER_DIFF"],
                     "主档")
        master_proposed = master_report["proposed_mapping"]
        check(master_proposed["master"]["sheets"].get("合同信息") == "合同",
              "主档 proposed_mapping 未建议表名映射：%r"
              % (master_proposed["master"]["sheets"],))
        master_mapping_path = os.path.join(tmpdir, "主档映射.json")
        write_json(master_mapping_path, master_proposed)

        backup_dir = os.path.join(master_root, "_备份")
        processed = run_script([APPLY, master_root, master_mapping_path,
                                "--fix-master", "--fix-order"])
        check(processed.returncode == 0,
              "主档修复 dry-run 失败：%s" % _text(processed.stderr))
        plan_text = _text(processed.stdout)
        check("主档修复·dry-run" in plan_text and "表改名" in plan_text,
              "主档修复 dry-run 未打印改名计划：%r" % plan_text[-800:])
        check(not os.path.isdir(backup_dir), "dry-run 不应创建 _备份/")
        check(sha256_file(master_path) == drifted_master_sha, "dry-run 不应修改主档")

        processed = run_script([APPLY, master_root, master_mapping_path,
                                "--fix-master", "--fix-order", "--write"])
        check(processed.returncode != 0, "主档修复未确认时应失败")
        check("必须经用户确认" in _text(processed.stderr), "主档修复未确认时报错缺失确认提示")

        processed = run_script([APPLY, master_root, master_mapping_path,
                                "--fix-master", "--fix-order", "--write",
                                "--confirmed-by", "fixture"])
        check(processed.returncode == 0,
              "主档修复失败：%s" % _text(processed.stderr))
        check("已修复" in _text(processed.stdout), "主档修复未报告完成")
        backups = glob.glob(os.path.join(backup_dir, "*.xlsx"))
        check(len(backups) == 1, "_备份/ 未生成备份：%r" % backups)
        check(sha256_file(backups[0]) == drifted_master_sha,
              "备份内容与修复前主档不一致（备份不可回滚）")
        check(sha256_file(master_path) != drifted_master_sha, "主档修复后文件未变化")
        check(not os.path.isfile(os.path.join(master_root, "_数据映射.json")),
              "--fix-master 只修主档，不应写 _数据映射.json")

        workbook = openpyxl.load_workbook(master_path)
        try:
            check(workbook.sheetnames == list(SHEETS.keys()),
                  "主档修复后工作表顺序/名称不符合规范：%s" % workbook.sheetnames)
            contract_header = [cell.value for cell in workbook["合同"][1]]
            check(contract_header == SHEETS["合同"][0],
                  "主档修复后合同表头不符合规范：%r" % contract_header)
            widths = [workbook["合同"].column_dimensions[get_column_letter(index)].width
                      for index in range(1, len(SHEETS["合同"][0]) + 1)]
            expected_widths = []
            for canonical in SHEETS["合同"][0]:
                old_name = "合同号" if canonical == "合同编号" else canonical
                expected_widths.append(drifted_contract_widths[
                    drifted_contract_header.index(old_name)])
            check(all((left is None and right is None)
                      or (left is not None and right is not None
                          and abs(float(left) - float(right)) < 0.01)
                      for left, right in zip(widths, expected_widths)),
                  "主档修复未保留列宽（应随列一起重排）：%r != %r"
                  % (widths, expected_widths))
        finally:
            workbook.close()

        cfg_repaired = load_cfg(master_path)
        check(cfg_repaired["contract"]["no"] == "HT-DEMO-2026-001"
              and cfg_repaired["contract"]["method"] == "铺货代销"
              and cfg_repaired["contract"]["basis"] == "实际动销",
              "主档修复后 load_master 未恢复合同条款：%r" % cfg_repaired["contract"])
        processed = run_engine(cfg_repaired)
        check(processed.returncode == 0,
              "主档修复后引擎失败：%s" % _text(processed.stderr)[-1500:])
        check("自检通过（残差 0.00）" in _text(processed.stderr), "主档修复后未自检通过")

        elapsed = os.times()
        cpu = (elapsed.user - started.user) + (elapsed.system - started.system)
        ok = True
        print("[schema_demo] OK（CPU %.1fs）" % cpu)
        print("  漂移识别 : 挂账 表名/5 列名/列序/多余列；库存 表名/列名/多候选表；"
              "主档 表名/列名/列序/表序")
        print("  issue    : SHEET_NAME_MISMATCH / FIELD_NAME_MISMATCH / FIELD_ORDER_DIFF / "
              "MULTI_CANDIDATE_SHEET / EXTRA_COLUMN")
        print("  无映射   : 引擎 rc≠0 且 stdout 0 字节，中文原因「数据表未找到含…」")
        print("  映射确认 : _数据映射.json（confirmed_by=fixture）+ "
              "_确认记录/映射_YYYYmmdd_HHMMSS_mmm.json；三份源导出只读")
        print("  带映射   : 引擎 rc=0 自检通过，输出与基线逐单元格一致（仅实际表名说明变化）")
        print("  主档修复 : 表名/列名/列序/表序 → 规范；load_master 恢复合同编号；"
              "_备份/ 1 份毫秒时间戳备份")
        print("  负例     : 映射指向不存在列 exit 1；未经 --confirmed-by exit 1；"
              "引擎遇到映射错配中文中止并指明是哪一项")
        return 0
    except SchemaDemoError as exc:
        print("[schema_demo] FAIL: " + str(exc), file=sys.stderr)
        print("[schema_demo] 临时目录已保留：%s" % tmpdir, file=sys.stderr)
        return 1
    except Exception:  # noqa: BLE001 - CLI 顶层兜底，输出可读错误
        import traceback
        traceback.print_exc()
        print("[schema_demo] 临时目录已保留：%s" % tmpdir, file=sys.stderr)
        return 1
    finally:
        if ok and not args.keep and os.path.isdir(tmpdir):
            shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
