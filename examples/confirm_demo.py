"""端到端演示 / CI 夹具：确认单零文件操作（生成确认页 → 提交 → 自动定位 → 写主档）。

用法：
    python3 examples/confirm_demo.py [--keep]

覆盖的验收链路（全部在系统临时目录，无网络、确定性、总耗时 <15s）：
  1. make_master 生成演示主档 → extract_contract 抽取 22 项候选 → make_confirm_html 生成确认页；
     断言：按钮「提交并保存确认单」、弹窗「确认并保存」；成功弹窗主文案「回到对话说『好了』」
     与副文案「助手会自动找到…不需要你移动任何文件」；文件名前缀「合同确认单_<客户名>_」；
     含 File System Access API（showSaveFilePicker）特性检测 + createWritable + 取消/失败回退；
     旧引导词（放进客户目录 / 把该 JSON / 下载 JSON …）已清除；用户可见区域不出现「JSON」；
     单文件、无 http(s) 外链、占位符全部替换。
  2. 临时「下载」目录：一份合法确认单（旧版「_YYYY-MM-DD」命名，验证向后兼容）＋一份缺确认人
     的坏文件＋一份无关 JSON → locate_confirm.py 退出 0 且选中合法文件；--json 字段齐全。
  3. 端到端：定位到的确认单 → apply_contract dry-run（主档不变）→ --write 到主档副本
     → 断言字段写入、_确认记录/ 归档确认单、_备份/ 备份为写入前主档；重复 --write 幂等。
  4. 再放一份更新的有效确认单 → locate 退出 2、ambiguous=true、candidates=2、chosen=最新。
  5. 空目录 → locate 退出 1，stdout 说明查过哪些目录、可能原因与下一步建议。
  6. 其它入口：--file 指定文件；--customer 客户目录顶层与 _确认记录/；平台「下载」目录回退。
     测试统一把 HOME/USERPROFILE 指向临时目录，避免读到本机下载/桌面，保证确定性。

退出码 0 = 全部通过；失败时保留临时目录并打印路径。
"""
import argparse
import contextlib
import glob
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCRIPTS = os.path.join(ROOT, "skills", "dealer-settlement", "scripts")
MAKE_CONFIRM = os.path.join(SCRIPTS, "make_confirm_html.py")
LOCATE = os.path.join(SCRIPTS, "locate_confirm.py")
APPLY = os.path.join(SCRIPTS, "apply_contract.py")
EXTRACT = os.path.join(SCRIPTS, "extract_contract.py")

if HERE not in sys.path:
    sys.path.insert(0, HERE)
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)

import make_master  # noqa: E402  （生成演示主档）
from make_confirm_html import safe_filename_part  # noqa: E402  （文件名安全化单测）

CUSTOMER = "示例经销商有限公司"
CONTRACT_NO = "HT-DEMO-2026-001"
CONFIRMED_AT = "2026-11-03T09:45:12Z"

# 演示合同（虚构）：覆盖 extract_contract.py 的 22 项候选字段
CONTRACT_LINES = [
    "示例品牌方有限公司 与 示例经销商有限公司",
    "铺货代销合同（演示数据，人物与公司均为虚构）",
    "",
    "第一条 合同编号",
    "合同编号：HT-DEMO-2026-001",
    "",
    "第二条 合同主体",
    "甲方：示例品牌方有限公司（以下简称甲方）",
    "统一社会信用代码：91440100DEMO00001A",
    "地址：示例市示例区示例路100号",
    "联系人：张示例",
    "联系电话：4008000000",
    "",
    "乙方：示例经销商有限公司（以下简称乙方）",
    "统一社会信用代码：91310000DEMO00002B",
    "地址：示例市示例区示例路200号",
    "联系人：王示例",
    "联系电话：13900000000",
    "",
    "第三条 合同期限",
    "合同期限：自 2026年10月19日 至 2026年12月31日",
    "",
    "第四条 结算约定",
    "结算方式：铺货代销（按实际动销结算）",
    "结算基数：实际动销",
    "付款期限：每月 10 日前",
    "收到甲方对账单后 5 个工作日内签字确认并寄回。",
    "异议期为 3 个工作日内提出。",
    "",
    "第五条 收款账户",
    "户名：示例品牌方有限公司",
    "账号：44056601040012345",
    "开户行：示例银行股份有限公司示例支行",
    "",
    "第六条 价格",
    "供货价格：详见附件二 示例政策表v2。",
]


class ConfirmDemoError(AssertionError):
    pass


def check(cond, msg):
    if not cond:
        raise ConfirmDemoError(msg)


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
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
        fh.write("\n")


def _clean_env(home):
    """清理可能影响结果的变量；HOME/USERPROFILE 指向临时目录，保证 locate 的确定性。"""
    drop = {"RECON_CFG", "HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "NO_PROXY"}
    env = {key: value for key, value in os.environ.items() if key.upper() not in drop}
    env.update({
        "PYTHONIOENCODING": "utf-8",
        "PYTHONUTF8": "1",
        "HOME": home,
        "USERPROFILE": home,
    })
    return env


def run_script(args, home):
    return subprocess.run([sys.executable] + args, capture_output=True,
                          env=_clean_env(home), cwd=ROOT)


def _set_mtime(path, epoch):
    os.utime(path, (epoch, epoch))


def _without_blocks(html, tag):
    return re.sub(r"<%s\b[^>]*>.*?</%s>" % (tag, tag), "", html, flags=re.S | re.I)


def _visible_text(html):
    return re.sub(r"<[^>]+>", "", _without_blocks(_without_blocks(html, "style"), "script"))


def check_confirm_html(html):
    """断言确认页满足零文件操作交互要求。"""
    for placeholder in ("__DATA__", "__META__", "__TITLE__", "__SUBTITLE__", "__SRC__"):
        check(placeholder not in html, "确认页占位符未替换：%s" % placeholder)

    check("提交并保存确认单" in html, "提交按钮文案应为「提交并保存确认单」")
    check("确认并保存" in html, "确认弹窗文案应为「确认并保存」")
    check("已提交成功" in html, "成功弹窗缺少「已提交成功」")
    check("回到对话，对助手说一句" in html and "好了" in html,
          "成功弹窗缺少「回到对话，对助手说一句『好了』」主文案")
    check("助手会自动找到" in html, "成功弹窗缺少「助手会自动找到…」副文案")
    check("不需要你移动任何文件" in html, "成功弹窗未强调「不需要你移动任何文件」")
    check("存到客户目录" in html, "成功弹窗缺少次要按钮「存到客户目录…」")
    check("复制确认单内容" in html, "高级回退缺少「复制确认单内容」")

    check("typeof window.showSaveFilePicker" in html,
          "缺少 File System Access API 特性检测（showSaveFilePicker）")
    check("createWritable" in html, "缺少直接写入实现（createWritable）")
    check("AbortError" in html, "缺少取消/失败回退处理（AbortError）")
    check("YYYYMMDD_HHMMSS_mmm" in html, "页面未说明确定性文件名的时间戳格式")

    for old in ("放进客户目录", "把该 JSON", "下载 JSON", "提交并下载", "确认并下载", "复制 JSON"):
        check(old not in html, "确认页仍含旧引导：%s" % old)

    visible = _visible_text(html)
    check("JSON" not in visible, "用户可见文案出现「JSON」字样（主流程应为普通中文）")

    check("http://" not in html and "https://" not in html, "确认页含 http(s) 外链")
    check(_without_blocks(html, "script").count("src=") == 0, "确认页引用了外部图片/脚本")
    check("<link" not in html, "确认页引用了外部样式")
    check(html.count("<script") == 1, "确认页应只有一个内联脚本")
    check(html.count("<style") == 1, "确认页应只有一个内联样式")

    match = re.search(r"const META = (\{.*?\});", html, re.S)
    check(match is not None, "确认页缺少 META（无法核对文件名前缀）")
    meta = json.loads(match.group(1))
    expected_prefix = "合同确认单_%s_" % CUSTOMER
    check(meta.get("download_prefix") == expected_prefix,
          "META.download_prefix 应为 %r，实际 %r" % (expected_prefix, meta.get("download_prefix")))
    check(meta.get("customer") == CUSTOMER, "META.customer 与主档不一致：%r" % meta.get("customer"))
    check(meta.get("contract") == CONTRACT_NO, "META.contract 与主档不一致：%r" % meta.get("contract"))
    check(meta.get("target_dir"), "META.target_dir 缺失（客户目录信息）")


def build_master(path):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with contextlib.redirect_stdout(io.StringIO()):
        make_master.build(path)
    check(os.path.isfile(path), "make_master 未生成主档：%s" % path)


def write_confirm(path, fields, operator="王示例", confirmed_at=CONFIRMED_AT):
    """构造一份与确认页 buildPayload() 同构的确认单。"""
    payload = {
        "schema": "contract-confirm/v1",
        "meta": {"customer": CUSTOMER, "contract": CONTRACT_NO, "operator": operator,
                 "memo": "", "confirmed_at": confirmed_at,
                 "source_file": "示例合同.txt", "page": "2026-11-03 17:00"},
        "fields": fields,
        "skipped": [],
        "attachments": [],
    }
    write_json(path, payload)
    return path


def _cell_value(master_path, target):
    """读取主档单行字段现值（演示只用到 主体.* 与 合同.* 两类 target）。"""
    workbook = openpyxl.load_workbook(master_path, data_only=True)
    try:
        if target.startswith("主体."):
            _, role, column = target.split(".", 2)
            sheet = workbook["主体"]
            headers = [cell.value for cell in sheet[1]]
            role_index = headers.index("角色") + 1
            column_index = headers.index(column) + 1
            for row in range(2, sheet.max_row + 1):
                if str(sheet.cell(row=row, column=role_index).value or "").strip() == role:
                    return sheet.cell(row=row, column=column_index).value
            return None
        if target.startswith("合同."):
            sheet = workbook["合同"]
            headers = [cell.value for cell in sheet[1]]
            return sheet.cell(row=2, column=headers.index(target.split(".", 1)[1]) + 1).value
        return None
    finally:
        workbook.close()


def main(argv=None):
    parser = argparse.ArgumentParser(description="确认单零文件操作端到端演示 / CI 夹具")
    parser.add_argument("--keep", action="store_true", help="保留临时目录（失败时总是保留）")
    args = parser.parse_args(argv)

    started = os.times()
    tmpdir = tempfile.mkdtemp(prefix="dealer_confirm_demo_")
    home = os.path.join(tmpdir, "home")
    os.makedirs(home, exist_ok=True)
    ok = False
    try:
        # ---------- 1. 主档 → 22 项候选 → 确认页 ----------
        master = os.path.join(tmpdir, "铺货代销客户", CUSTOMER, "_主档.xlsx")
        build_master(master)

        contract_path = os.path.join(tmpdir, "示例合同.txt")
        with open(contract_path, "w", encoding="utf-8") as fh:
            fh.write("\n".join(CONTRACT_LINES) + "\n")
        candidate_path = os.path.join(tmpdir, "合同候选.json")
        processed = run_script([EXTRACT, contract_path, "-o", candidate_path], home)
        check(processed.returncode == 0,
              "extract_contract 失败（rc=%d）：%s" % (processed.returncode, _text(processed.stderr)))
        candidate = read_json(candidate_path)
        fields = candidate.get("fields") or []
        check(len(fields) == 22, "候选字段应为 22 项，实际 %d 项" % len(fields))

        page = os.path.join(tmpdir, "确认页.html")
        processed = run_script([MAKE_CONFIRM, candidate_path, master, "-o", page], home)
        check(processed.returncode == 0,
              "make_confirm_html 失败（rc=%d）：%s"
              % (processed.returncode, _text(processed.stderr)))
        check(os.path.isfile(page), "确认页未生成：%s" % page)
        with open(page, encoding="utf-8") as fh:
            html = fh.read()
        check_confirm_html(html)
        check(safe_filename_part("示例/经销商:有限公司") == "示例_经销商_有限公司",
              "客户名安全化未替换保留字符：%r" % safe_filename_part("示例/经销商:有限公司"))
        check(safe_filename_part("   ") == "客户", "空客户名应回落为「客户」")

        confirm_fields = []
        for field in fields:
            value = field.get("value")
            if value is None or str(value).strip() == "":
                continue
            confirm_fields.append({"field": field.get("field", ""),
                                   "target": field.get("target", ""),
                                   "value": value, "from": "新增"})
        check(len(confirm_fields) >= 20,
              "演示候选有效字段过少（%d/22），确认页无法覆盖主要字段" % len(confirm_fields))

        # ---------- 2. 模拟用户在确认页提交：确认单落在临时「下载」目录 ----------
        now = time.time()
        downloads = os.path.join(tmpdir, "下载")
        os.makedirs(downloads, exist_ok=True)
        legacy = write_confirm(
            os.path.join(downloads, "合同确认单_%s_2026-11-03.json" % CUSTOMER),
            confirm_fields)
        _set_mtime(legacy, now - 180)          # 旧版「_YYYY-MM-DD」命名，验证向后兼容
        bad = write_confirm(
            os.path.join(downloads, "合同确认单_缺少确认人_20261103_120000_000.json"),
            [], operator="")
        _set_mtime(bad, now - 170)
        unrelated = os.path.join(downloads, "会议纪要.json")
        write_json(unrelated, {"title": "与确认单无关的文件"})
        _set_mtime(unrelated, now - 160)

        processed = run_script([LOCATE, "--dir", downloads,
                                "--max-age-minutes", "60", "--json"], home)
        check(processed.returncode == 0,
              "locate_confirm 单份确认单应 exit 0，实际 %d：%s"
              % (processed.returncode, _text(processed.stdout)[-800:] + _text(processed.stderr)[-800:]))
        located = json.loads(_text(processed.stdout))
        check(located["found"] is True and located["ambiguous"] is False,
              "locate --json 状态错误：%r" % ({k: located[k] for k in ("found", "ambiguous")}))
        check(located["chosen"] is not None
              and located["chosen"]["path"] == os.path.abspath(legacy),
              "locate 未选中合法确认单：%r" % (located["chosen"],))
        check(len(located["candidates"]) == 1,
              "坏文件/无关文件不应进入候选：%r" % ([c["path"] for c in located["candidates"]],))
        chosen = located["candidates"][0]
        for key in ("path", "sha256", "size", "mtime", "operator",
                    "confirmed_at", "customer", "fields_count", "source"):
            check(key in chosen, "locate --json candidate 缺字段：%s" % key)
        check(chosen["sha256"] == sha256_file(legacy), "locate 计算的 sha256 与文件不一致")
        check(chosen["operator"] == "王示例", "locate 未解析确认人：%r" % chosen["operator"])
        check(chosen["confirmed_at"] == CONFIRMED_AT, "locate 未解析确认时间")
        check(chosen["customer"] == CUSTOMER, "locate 未解析客户名")
        check(chosen["fields_count"] == len(confirm_fields),
              "locate 项数错误：%d != %d" % (chosen["fields_count"], len(confirm_fields)))
        check(chosen["source"] == "explicit_dir", "locate 来源标记错误：%r" % chosen["source"])
        check(os.path.abspath(downloads) in located["searched_dirs"],
              "locate --json 未列出已查目录：%r" % located["searched_dirs"])
        check(any("缺少确认人" in warning for warning in located["warnings"]),
              "坏确认单未记入 warnings：%r" % located["warnings"])

        processed = run_script([LOCATE, "--dir", downloads, "--max-age-minutes", "60"], home)
        check(processed.returncode == 0, "locate 人话模式应 exit 0")
        human = _text(processed.stdout)
        check("确认人" in human and "王示例" in human, "locate 人话摘要未显示确认人")
        check(os.path.basename(legacy) in human, "locate 人话摘要未显示文件名")
        check("下一步" in human and "apply_contract.py" in human,
              "locate 人话摘要缺少下一步写库命令")

        # ---------- 3. 端到端：定位到的确认单 → dry-run → --write → 幂等 ----------
        write_root = os.path.join(tmpdir, "写库", "铺货代销客户", CUSTOMER)
        os.makedirs(write_root, exist_ok=True)
        master_copy = os.path.join(write_root, "_主档.xlsx")
        shutil.copy2(master, master_copy)
        master_sha = sha256_file(master_copy)

        expected_changes = 0
        for item in confirm_fields:
            current = _cell_value(master_copy, item["target"])
            if str(current if current is not None else "").strip() != str(item["value"]).strip():
                expected_changes += 1
        check(expected_changes > 0, "演示候选与主档完全一致，无法验证写入")

        chosen_path = located["chosen"]["path"]
        processed = run_script([APPLY, master_copy, chosen_path], home)
        check(processed.returncode == 0,
              "apply_contract dry-run 失败（rc=%d）：%s"
              % (processed.returncode, _text(processed.stderr)[-1500:]))
        dry_run = _text(processed.stdout)
        check("[dry-run] 未写入" in dry_run, "dry-run 未打印「未写入」")
        check("可写入 %d 项" % expected_changes in dry_run,
              "dry-run 可写入项数不符（预期 %d）：%s" % (expected_changes, dry_run[-600:]))
        check("拒绝 0 项" in dry_run, "演示确认单不应有被类型校验拒绝的字段：%s" % dry_run[-600:])
        check(sha256_file(master_copy) == master_sha, "dry-run 不应修改主档")

        processed = run_script([APPLY, master_copy, chosen_path, "--write"], home)
        check(processed.returncode == 0,
              "apply_contract --write 失败（rc=%d）：%s"
              % (processed.returncode, _text(processed.stderr)[-1500:]))
        written = _text(processed.stdout)
        check("已写入" in written, "写库未报告「已写入」：%s" % written[-600:])
        check("确认单已归档" in written, "写库未归档确认单：%s" % written[-600:])
        for item in confirm_fields:
            current = _cell_value(master_copy, item["target"])
            check(str(current if current is not None else "").strip()
                  == str(item["value"]).strip(),
                  "字段未按确认值写入：%s（主档=%r，确认=%r）"
                  % (item["target"], current, item["value"]))

        archives = sorted(glob.glob(os.path.join(write_root, "_确认记录", "*.json")))
        check(len(archives) == 1, "_确认记录/ 应有 1 份归档，实际 %r" % archives)
        archived = read_json(archives[0])
        check(archived.get("schema") == "contract-confirm/v1", "归档确认单 schema 错误")
        check(str((archived.get("meta") or {}).get("operator")) == "王示例", "归档确认单缺确认人")
        check(len(archived.get("fields") or []) == len(confirm_fields), "归档确认单字段数不一致")

        backups = sorted(glob.glob(os.path.join(write_root, "_备份", "*.xlsx")))
        check(len(backups) == 1, "_备份/ 应有 1 份备份，实际 %r" % backups)
        check(sha256_file(backups[0]) == master_sha, "备份内容应等于写入前的主档")

        after_write_sha = sha256_file(master_copy)
        processed = run_script([APPLY, master_copy, chosen_path, "--write"], home)
        check(processed.returncode == 0, "重复写入应成功返回")
        check("[无变化]" in _text(processed.stdout), "重复写入未命中幂等分支（[无变化]）")
        check(sha256_file(master_copy) == after_write_sha, "重复写入不应改动主档")
        check(len(glob.glob(os.path.join(write_root, "_备份", "*.xlsx"))) == 1,
              "重复写入不应再生成备份")
        check(len(glob.glob(os.path.join(write_root, "_确认记录", "*.json"))) == 1,
              "重复写入不应重复归档")

        # ---------- 4. 出现第二份有效确认单：歧义 → exit 2，chosen 为最新 ----------
        newer = write_confirm(
            os.path.join(downloads, "合同确认单_%s_20261103_180512_456.json" % CUSTOMER),
            confirm_fields, operator="张复核", confirmed_at="2026-11-03T10:05:12Z")
        _set_mtime(newer, now - 60)
        processed = run_script([LOCATE, "--dir", downloads,
                                "--max-age-minutes", "60", "--json"], home)
        check(processed.returncode == 2,
              "多份确认单应 exit 2，实际 %d" % processed.returncode)
        ambiguous = json.loads(_text(processed.stdout))
        check(ambiguous["ambiguous"] is True, "多份确认单 ambiguous 应为 true")
        check(len(ambiguous["candidates"]) == 2,
              "多份确认单 candidates 应为 2：%r" % ([c["path"] for c in ambiguous["candidates"]],))
        check(ambiguous["chosen"]["path"] == os.path.abspath(newer), "chosen 应为最新的一份")
        check([item["path"] for item in ambiguous["candidates"]]
              == [os.path.abspath(newer), os.path.abspath(legacy)],
              "候选应按修改时间降序排列")
        processed = run_script([LOCATE, "--dir", downloads, "--max-age-minutes", "60"], home)
        check(processed.returncode == 2, "多份确认单人话模式应 exit 2")
        human = _text(processed.stdout)
        check("确认单有多个候选" in human and "--file" in human,
              "歧义时未提示用 --file 指定")

        # ---------- 5. 空目录：exit 1，输出已查目录、原因与建议 ----------
        empty_dir = os.path.join(tmpdir, "空下载")
        os.makedirs(empty_dir, exist_ok=True)
        processed = run_script([LOCATE, "--dir", empty_dir,
                                "--max-age-minutes", "60", "--json"], home)
        check(processed.returncode == 1, "空目录应 exit 1，实际 %d" % processed.returncode)
        not_found = json.loads(_text(processed.stdout))
        check(not_found["found"] is False and not_found["ambiguous"] is False
              and not_found["chosen"] is None and not_found["candidates"] == [],
              "空目录 --json 状态错误：%r" % not_found)
        check(os.path.abspath(empty_dir) in not_found["searched_dirs"],
              "空目录 --json 未列出已查目录：%r" % not_found["searched_dirs"])
        processed = run_script([LOCATE, "--dir", empty_dir, "--max-age-minutes", "60"], home)
        check(processed.returncode == 1, "空目录人话模式应 exit 1")
        human = _text(processed.stdout)
        check(os.path.abspath(empty_dir) in human, "未找到时未列出查过的目录")
        check("可能原因" in human and "下一步" in human, "未找到时缺少原因与下一步建议")

        # ---------- 6. 其它入口：--file / --customer / 平台「下载」目录 ----------
        processed = run_script([LOCATE, "--file", legacy,
                                "--max-age-minutes", "60", "--json"], home)
        check(processed.returncode == 0, "--file 指定有效确认单应 exit 0")
        by_file = json.loads(_text(processed.stdout))
        check(by_file["chosen"]["path"] == os.path.abspath(legacy)
              and by_file["chosen"]["source"] == "explicit_file",
              "--file 未直接采用指定文件")

        customer_root = os.path.join(tmpdir, "扫描", "客户A")
        os.makedirs(customer_root, exist_ok=True)
        home2 = os.path.join(tmpdir, "home2")
        os.makedirs(home2, exist_ok=True)
        archive_dir = os.path.join(customer_root, "_确认记录")
        os.makedirs(archive_dir, exist_ok=True)
        in_archive = os.path.join(archive_dir, "20261103_174000_000_合同确认.json")
        write_confirm(in_archive, confirm_fields)
        _set_mtime(in_archive, now - 40)
        processed = run_script([LOCATE, "--customer", customer_root,
                                "--max-age-minutes", "60", "--json"], home2)
        check(processed.returncode == 0, "--customer 应找到 _确认记录/ 内的确认单")
        by_customer = json.loads(_text(processed.stdout))
        check(by_customer["chosen"]["source"] == "customer_confirm_dir",
              "_确认记录/ 内的确认单 source 标记错误：%r" % by_customer["chosen"]["source"])
        in_root = os.path.join(customer_root, "合同确认单_%s_20261103_174500_123.json" % CUSTOMER)
        write_confirm(in_root, confirm_fields, operator="李复核")
        _set_mtime(in_root, now - 20)
        processed = run_script([LOCATE, "--customer", customer_root,
                                "--max-age-minutes", "60", "--json"], home2)
        check(processed.returncode == 0, "--customer 应找到客户目录顶层的确认单")
        by_customer = json.loads(_text(processed.stdout))
        check(by_customer["chosen"]["path"] == os.path.abspath(in_root)
              and by_customer["chosen"]["source"] == "customer_root",
              "--customer 未优先采用客户目录顶层的确认单")

        platform_downloads = os.path.join(home, "Downloads")
        os.makedirs(platform_downloads, exist_ok=True)
        platform_confirm = write_confirm(
            os.path.join(platform_downloads,
                         "合同确认单_%s_20261103_181000_789.json" % CUSTOMER),
            confirm_fields)
        _set_mtime(platform_confirm, now - 30)
        processed = run_script([LOCATE, "--max-age-minutes", "60", "--json"], home)
        check(processed.returncode == 0, "未显式指定目录时应回退到平台「下载」目录")
        by_platform = json.loads(_text(processed.stdout))
        check(by_platform["chosen"]["path"] == os.path.abspath(platform_confirm),
              "平台「下载」目录回退未选中确认单")
        check(by_platform["chosen"]["source"] == "downloads",
              "平台「下载」目录回退 source 标记错误")

        elapsed = os.times()
        cpu = (elapsed.user - started.user) + (elapsed.system - started.system)
        ok = True
        print("[confirm_demo] OK（CPU %.1fs）" % cpu)
        print("  确认页    : 单文件内联；按钮「提交并保存确认单」，弹窗「确认并保存」，"
              "成功弹窗「回到对话说『好了』」")
        print("  文件名    : 合同确认单_%s_YYYYMMDD_HHMMSS_mmm.json"
              "（META.download_prefix 由 Python 注入，含客户名安全化）" % CUSTOMER)
        print("  零文件操作: 成功弹窗有「存到客户目录…」（showSaveFilePicker 特性检测 + "
              "createWritable；取消/不支持回退默认下载）")
        print("  自动定位  : 旧版日期命名与新版确定性命名均命中；坏文件/无关 JSON 只记 warnings")
        print("  退出码    : 单份=0，多份=2（ambiguous，chosen=最新），空目录=1（列目录/原因/建议）")
        print("  写库链路  : locate → dry-run → --write（%d 项写入）；_确认记录/ 归档、"
              "_备份/ 1 份；重复写幂等" % expected_changes)
        print("  入口覆盖  : --dir / --file / --customer（顶层 + _确认记录/）/ 平台「下载」目录")
        return 0
    except ConfirmDemoError as exc:
        print("[confirm_demo] FAIL: %s" % exc, file=sys.stderr)
        print("[confirm_demo] 临时目录已保留：%s" % tmpdir, file=sys.stderr)
        return 1
    except Exception:  # noqa: BLE001 - CLI 顶层兜底，输出可读错误
        import traceback
        traceback.print_exc()
        print("[confirm_demo] 临时目录已保留：%s" % tmpdir, file=sys.stderr)
        return 1
    finally:
        if ok and not args.keep and os.path.isdir(tmpdir):
            shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
