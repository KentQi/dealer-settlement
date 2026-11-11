"""L6 复核结论校验与落盘（交付门禁）。

用法：
    python3 apply_review.py <工作簿.xlsx> <报告.json> [--packet packet.json] [--report-dir DIR] [--apply|--check]

行为：
- 默认 dry-run：只校验报告并打印将执行的动作，不写任何文件；
- --apply：结论 通过 / 有条件通过 → 写/更新 08-复核记录、归档报告（JSON+MD）、
  保存后计算新 sha256 并写 _复核记录/复核锁.json，退出码 0；
  结论 不通过 → 只归档报告、不写 08 表、stderr 说明阻断、退出码 1（不得交付）；
- --check：交付前校验（08 表 + 复核锁 + 当前文件哈希 + 结论），有效打印「复核有效」退出 0，
  否则逐条给中文原因退出 1。复核后任何保存都会使复核锁失效，必须重新复核。

校验规则（不满足即拒绝，退出码 1）：
1. schema 必须为 dealer-settlement/review-report@1；reviewer.name/model、reviewed_at(ISO8601)、
   verdict、checklist、findings、summary 必须齐全且类型正确；
2. workbook_sha256 必须等于当前工作簿 sha256（复核后文件已变更 → 必须重评）；
3. checklist 必须覆盖 R1–R10（缺 → 失败；未知 id → 警告；重复 → 警告）；
4. 一致性：有 level=阻断 的 finding ⇒ verdict 必须=不通过；
   无阻断但有 警告 finding ⇒ 不得为 通过；不通过/有条件通过 的 summary 必须非空；
5. 若提供 --packet：复核包 schema、workbook.sha256 必须与当前工作簿一致，period 不得冲突。

幂等：同一份报告重复 --apply 不会重复改写（检测到「报告已归档 + 锁 post_sha256 == 当前文件」时跳过）。
"""
import argparse
import json
import os
import re
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:  # 脚本目录即 sys.path[0]；作为模块被导入时兜底
    from review_common import (
        DELIVERABLE_VERDICTS, LEGACY_REVIEW_SHEETS, LEVELS, LOCK_NAME,
        REQUIRED_IDS, REQUIRED_ITEMS, RESULTS, REVIEW_DIR, REVIEW_SHEET,
        SCHEMA_PACKET, SCHEMA_REPORT, VERDICTS,
        iso_now, load_json, parse_iso8601, sha256_file,
    )
except ImportError:  # pragma: no cover - 仅异常运行方式触发
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from review_common import (
        DELIVERABLE_VERDICTS, LEGACY_REVIEW_SHEETS, LEVELS, LOCK_NAME,
        REQUIRED_IDS, REQUIRED_ITEMS, RESULTS, REVIEW_DIR, REVIEW_SHEET,
        SCHEMA_PACKET, SCHEMA_REPORT, VERDICTS,
        iso_now, load_json, parse_iso8601, sha256_file,
    )

CODE = chr(96)  # Markdown 行内代码标记；用常量而非字面反引号，便于源码审阅

DISCLAIMER = (
    "本复核记录由当前 harness 的 LLM 依据复核包（dealer-settlement/review-packet@1）按 R1–R10 "
    "清单逐项执行，仅代表对该工作簿数据一致性的程序性复核，不构成审计意见或会计师意见；"
    "差异归因、税率、赔偿/责任类条款及贴花兑付方式均须由双方书面确认。"
    "复核锁与工作簿哈希绑定：复核后任何保存（含用 Excel/WPS 打开另存）都会使复核失效，须重新复核。"
)

SHEET_WIDTHS = (12, 52, 14, 22, 12, 12, 52, 40)
NAVY = "1F3864"
LIGHT = "D9E2F3"
GREY = "F2F2F2"
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


class ReviewError(Exception):
    """复核门禁失败（中文原因由 main 打印到 stderr）。"""


def _code(text):
    return CODE + str(text) + CODE


def _norm(value):
    if value is None:
        return ""
    return re.sub(r"[\s\u3000]+", "", str(value))


def _right_value(ws, row, col):
    """同一行右侧第一个非空单元格值（合并单元格取左上角）。"""
    for column in range(col + 1, ws.max_column + 1):
        cell = ws.cell(row=row, column=column)
        if cell.value is not None and str(cell.value).strip() != "":
            return cell.value
    return None


# ---------------------------------------------------------------- 报告 Markdown
def report_to_md(report):
    """复核报告 JSON → Markdown（归档用，确定性输出）。"""
    lines = ["# L6 复核报告", ""]
    lines.append("- schema：%s" % _code(report.get("schema")))
    reviewer = report.get("reviewer") or {}
    lines.append("- 复核人：%s（模型：%s）"
                 % (reviewer.get("name") or "—", reviewer.get("model") or "—"))
    lines.append("- 复核时间：%s" % (report.get("reviewed_at") or "—"))
    lines.append("- 复核结论：%s" % (report.get("verdict") or "—"))
    lines.append("- 被复核工作簿哈希（SHA-256）：%s" % (report.get("workbook_sha256") or "—"))
    period = report.get("period")
    if isinstance(period, dict) and period.get("start") and period.get("end"):
        lines.append("- 期间：%s 至 %s" % (period["start"], period["end"]))
    elif isinstance(period, str) and period.strip():
        lines.append("- 期间：%s" % period.strip())
    else:
        lines.append("- 期间：未标注")
    lines.append("")

    lines.append("## 一、R1–R10 复核清单")
    lines.append("")
    lines.append("| 序号 | 清单项 | 结果 | 备注 |")
    lines.append("|---|---|---|---|")
    for index, item in enumerate(report.get("checklist") or [], 1):
        if not isinstance(item, dict):
            continue
        lines.append("| %d | %s | %s | %s |"
                     % (index, item.get("id") or "—", item.get("result") or "—",
                        item.get("note") or "—"))
    lines.append("")

    lines.append("## 二、问题清单（findings，%d 条）" % len(report.get("findings") or []))
    lines.append("")
    findings = report.get("findings") or []
    if findings:
        lines.append("| 序号 | 级别 | 类别 | 工作表 | 单元格 | 问题描述 | 证据 | 建议 |")
        lines.append("|---|---|---|---|---|---|---|---|")
        for index, finding in enumerate(findings, 1):
            if not isinstance(finding, dict):
                continue
            lines.append("| %d | %s | %s | %s | %s | %s | %s | %s |"
                         % (index, finding.get("level") or "—",
                            finding.get("category") or "—", finding.get("sheet") or "—",
                            finding.get("cell") or "—", finding.get("claim") or "—",
                            finding.get("evidence") or "—", finding.get("suggestion") or "—"))
    else:
        lines.append("（无）")
    lines.append("")

    lines.append("## 三、复核摘要")
    lines.append("")
    summary = report.get("summary")
    lines.append(summary.strip() if isinstance(summary, str) and summary.strip() else "（无）")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append(DISCLAIMER)
    lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------- 08-复核记录
def _style(cell, bold=False, size=10, color="000000", fill=None, horizontal="left"):
    cell.font = Font(name="微软雅黑", size=size, bold=bold, color=color)
    cell.alignment = Alignment(horizontal=horizontal, vertical="center", wrap_text=True)
    cell.border = BORDER
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)


def _put(ws, row, col1, col2, value, bold=False, size=10, color="000000",
         fill=None, horizontal="left"):
    """写一个单元格；col1 != col2 时先合并再加边框。"""
    if col2 > col1:
        ws.merge_cells(start_row=row, start_column=col1, end_row=row, end_column=col2)
    for column in range(col1, col2 + 1):
        _style(ws.cell(row=row, column=column), bold=bold, size=size, color=color,
               fill=fill, horizontal=horizontal)
    cell = ws.cell(row=row, column=col1, value=value)
    return cell


def _section(ws, row, text):
    _put(ws, row, 1, len(SHEET_WIDTHS), text, bold=True, size=11, color=NAVY, fill=LIGHT)
    return row + 1


def write_review_sheet(wb, report, pre_sha, archive_name, packet_name, workbook_name):
    """写入/更新「10-复核记录」工作表（只动这一张表，其余表原样）。"""
    for _name in (REVIEW_SHEET,) + LEGACY_REVIEW_SHEETS:
        if _name in wb.sheetnames:
            del wb[_name]
    ws = wb.create_sheet(REVIEW_SHEET)
    for index, width in enumerate(SHEET_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.sheet_view.showGridLines = False

    reviewer = report.get("reviewer") or {}
    findings = report.get("findings") or []
    levels = [item.get("level") for item in findings if isinstance(item, dict)]
    summary = report.get("summary") if isinstance(report.get("summary"), str) else ""

    row = 1
    _put(ws, row, 1, len(SHEET_WIDTHS), "08 复核记录（L6 LLM 语义复核）",
         bold=True, size=14, color="FFFFFF", fill=NAVY, horizontal="center")
    ws.row_dimensions[row].height = 30
    row += 1
    _put(ws, row, 1, len(SHEET_WIDTHS),
         "被复核工作簿：%s　·　复核协议：%s　·　本记录由 apply_review.py 写入，复核后任何保存都会使复核锁失效"
         % (workbook_name or "—", SCHEMA_REPORT),
         size=9, color="595959", fill=GREY)
    ws.row_dimensions[row].height = 30
    row += 2

    row = _section(ws, row, "一、复核基本信息")
    _put(ws, row, 1, 1, "复核人", bold=True)
    _put(ws, row, 2, 3, reviewer.get("name") or "—")
    _put(ws, row, 4, 4, "复核模型", bold=True)
    _put(ws, row, 5, 8, reviewer.get("model") or "—")
    row += 1
    _put(ws, row, 1, 1, "复核时间", bold=True)
    _put(ws, row, 2, 3, report.get("reviewed_at") or "—")
    _put(ws, row, 4, 4, "复核结论", bold=True)
    _put(ws, row, 5, 8, report.get("verdict") or "—", bold=True)
    row += 1
    _put(ws, row, 1, 1, "被复核工作簿哈希（SHA-256）", bold=True)
    _put(ws, row, 2, 8, pre_sha)
    row += 1
    _put(ws, row, 1, 1, "复核报告", bold=True)
    _put(ws, row, 2, 3, archive_name)
    _put(ws, row, 4, 4, "复核包", bold=True)
    _put(ws, row, 5, 8, packet_name or "—")
    row += 1
    _put(ws, row, 1, 1, "复核范围", bold=True)
    _put(ws, row, 2, 8,
         "R1–R10 共 10 项清单逐项判定；findings 共 %d 条（阻断 %d / 警告 %d / 提示 %d）"
         % (len(findings), levels.count("阻断"), levels.count("警告"), levels.count("提示")))
    row += 2

    row = _section(ws, row, "二、R1–R10 复核清单")
    _put(ws, row, 1, 1, "序号", bold=True, fill=GREY, horizontal="center")
    _put(ws, row, 2, 5, "清单项", bold=True, fill=GREY, horizontal="center")
    _put(ws, row, 6, 6, "结果", bold=True, fill=GREY, horizontal="center")
    _put(ws, row, 7, 8, "备注", bold=True, fill=GREY, horizontal="center")
    row += 1
    checklist = [item for item in (report.get("checklist") or []) if isinstance(item, dict)]
    item_texts = dict(REQUIRED_ITEMS)
    for item in checklist:
        item_id = str(item.get("id") or "").strip()
        _put(ws, row, 1, 1, item_id or "—", horizontal="center")
        _put(ws, row, 2, 5, item_texts.get(item_id) or "—")
        _put(ws, row, 6, 6, item.get("result") or "—", horizontal="center")
        _put(ws, row, 7, 8, item.get("note") or "")
        row += 1
    row += 1

    row = _section(ws, row, "三、问题清单（findings）")
    _put(ws, row, 1, 1, "序号", bold=True, fill=GREY, horizontal="center")
    _put(ws, row, 2, 2, "问题描述", bold=True, fill=GREY, horizontal="center")
    _put(ws, row, 3, 3, "级别", bold=True, fill=GREY, horizontal="center")
    _put(ws, row, 4, 4, "类别", bold=True, fill=GREY, horizontal="center")
    _put(ws, row, 5, 5, "工作表", bold=True, fill=GREY, horizontal="center")
    _put(ws, row, 6, 6, "单元格", bold=True, fill=GREY, horizontal="center")
    _put(ws, row, 7, 7, "证据", bold=True, fill=GREY, horizontal="center")
    _put(ws, row, 8, 8, "建议", bold=True, fill=GREY, horizontal="center")
    row += 1
    if findings:
        for index, finding in enumerate(findings, 1):
            if not isinstance(finding, dict):
                continue
            _put(ws, row, 1, 1, index, horizontal="center")
            _put(ws, row, 2, 2, finding.get("claim") or "—")
            _put(ws, row, 3, 3, finding.get("level") or "—", horizontal="center")
            _put(ws, row, 4, 4, finding.get("category") or "—")
            _put(ws, row, 5, 5, finding.get("sheet") or "—", horizontal="center")
            _put(ws, row, 6, 6, finding.get("cell") or "—", horizontal="center")
            _put(ws, row, 7, 7, finding.get("evidence") or "—")
            _put(ws, row, 8, 8, finding.get("suggestion") or "—")
            row += 1
    else:
        _put(ws, row, 1, 8, "（无：未发现阻断/警告/提示项）", horizontal="center")
        row += 1
    row += 1

    row = _section(ws, row, "四、复核结论与免责声明")
    _put(ws, row, 1, 1, "复核摘要", bold=True)
    _put(ws, row, 2, 8, summary.strip() if summary.strip() else "（无）")
    ws.row_dimensions[row].height = 48
    row += 1
    _put(ws, row, 1, 1, "免责声明", bold=True)
    _put(ws, row, 2, 8, DISCLAIMER)
    ws.row_dimensions[row].height = 66
    row += 1
    _put(ws, row, 1, 1, "复核锁", bold=True)
    _put(ws, row, 2, 8,
         "_复核记录/复核锁.json（pre_sha256 / post_sha256 与文件哈希绑定；复核后任何保存都会使复核失效）")
    row += 1
    _put(ws, row, 1, 1, "复核依据", bold=True)
    _put(ws, row, 2, 8, "R1–R10 清单；复核包 %s；报告 JSON 归档于 _复核记录/"
         % (packet_name or "（未提供）"))
    return ws


def _review_sheet_of(wb):
    """复核记录工作表名：优先新名，兼容旧工作簿的 08-复核记录。"""
    for name in (REVIEW_SHEET,) + LEGACY_REVIEW_SHEETS:
        if name in wb.sheetnames:
            return name
    return ""


def _read_review_sheet(wb):
    """从复核记录表读取被复核哈希与复核结论（--check 用）。"""
    info = {}
    _name = _review_sheet_of(wb)
    if not _name:
        return info
    ws = wb[_name]
    for row in ws.iter_rows():
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            text = _norm(cell.value)
            if "pre_sha256" not in info and "被复核" in text and "哈希" in text:
                value = _right_value(ws, cell.row, cell.column)
                if value is not None:
                    info["pre_sha256"] = str(value).strip()
            elif "verdict" not in info and text == "复核结论":
                value = _right_value(ws, cell.row, cell.column)
                if value is not None:
                    info["verdict"] = str(value).strip()
    return info


# ---------------------------------------------------------------- 报告校验
def _period_pair(value):
    """把 period 归一成 (start, end) 或字符串；无法归一返回 (None, None)。"""
    if isinstance(value, dict):
        start, end = value.get("start"), value.get("end")
        if isinstance(start, str) and isinstance(end, str) and start.strip() and end.strip():
            return start.strip(), end.strip()
        return None, None
    if isinstance(value, str) and value.strip():
        return value.strip(), None
    return None, None


def validate_report(report, workbook_path, current_sha, packet=None, packet_path=None):
    """校验报告是否可应用；返回 (errors, warnings)，errors 非空即拒绝。"""
    errors = []
    warnings = []
    if not isinstance(report, dict):
        return ["报告 JSON 顶层必须是对象"], warnings

    if report.get("schema") != SCHEMA_REPORT:
        errors.append("报告 schema 必须为 %s，收到 %r" % (SCHEMA_REPORT, report.get("schema")))

    reviewer = report.get("reviewer")
    if not isinstance(reviewer, dict):
        errors.append("reviewer 必须是含 name / model 的对象")
    else:
        if not str(reviewer.get("name") or "").strip():
            errors.append("reviewer.name 不得为空（复核人必须留痕）")
        if not str(reviewer.get("model") or "").strip():
            errors.append("reviewer.model 不得为空")

    reviewed_at = report.get("reviewed_at")
    if parse_iso8601(reviewed_at) is None:
        errors.append("reviewed_at 须为带时区的 ISO8601（如 2026-11-03T12:00:00+08:00），"
                      "收到 %r" % (reviewed_at,))

    workbook_sha = report.get("workbook_sha256")
    if not isinstance(workbook_sha, str) or not re.fullmatch(r"[0-9a-fA-F]{64}", workbook_sha.strip()):
        errors.append("workbook_sha256 须为 64 位十六进制字符串，收到 %r" % (workbook_sha,))
    elif workbook_sha.strip().lower() != current_sha:
        errors.append("报告记录的 workbook_sha256 与当前工作簿不一致：报告=%s / 当前=%s。"
                      "工作簿在复核后已变更，须重新复核。" % (workbook_sha.strip().lower(), current_sha))

    verdict = report.get("verdict")
    if verdict not in VERDICTS:
        errors.append("verdict 须为 通过 / 有条件通过 / 不通过，收到 %r" % (verdict,))

    checklist = report.get("checklist")
    covered = set()
    checklist_bad = []
    if not isinstance(checklist, list):
        errors.append("checklist 必须是数组")
        checklist = []
    for index, item in enumerate(checklist, 1):
        if not isinstance(item, dict):
            errors.append("checklist 第 %d 项必须是对象" % index)
            continue
        item_id = str(item.get("id") or "").strip()
        result = item.get("result")
        if not item_id:
            errors.append("checklist 第 %d 项缺少 id" % index)
        elif item_id not in REQUIRED_IDS:
            warnings.append("checklist 出现未知清单项 %s（R1–R10 之外，不计入覆盖）" % item_id)
        else:
            if item_id in covered:
                warnings.append("checklist 重复出现 %s" % item_id)
            covered.add(item_id)
        if result not in RESULTS:
            errors.append("checklist %s 的 result 须为 通过/不通过/不适用，收到 %r"
                          % (item_id or ("第 %d 项" % index), result))
        elif result == "不通过":
            checklist_bad.append(item_id or ("第 %d 项" % index))
    missing = [item_id for item_id in REQUIRED_IDS if item_id not in covered]
    if missing:
        errors.append("checklist 未覆盖 R1–R10，缺少：" + "、".join(missing))

    findings = report.get("findings")
    levels = []
    if not isinstance(findings, list):
        errors.append("findings 必须是数组（无问题写 []）")
        findings = []
    for index, finding in enumerate(findings, 1):
        if not isinstance(finding, dict):
            errors.append("findings 第 %d 项必须是对象" % index)
            continue
        level = finding.get("level")
        if level not in LEVELS:
            errors.append("findings 第 %d 项的 level 须为 阻断/警告/提示，收到 %r"
                          % (index, level))
        else:
            levels.append(level)
        for key, label in (("category", "类别"), ("sheet", "工作表"), ("cell", "单元格"),
                           ("claim", "问题描述"), ("evidence", "证据"), ("suggestion", "建议")):
            if not str(finding.get(key) or "").strip():
                warnings.append("findings 第 %d 项缺少「%s」" % (index, label))

    summary = report.get("summary")
    if not isinstance(summary, str):
        errors.append("summary 必须是字符串（无内容写空字符串）")
        summary = ""

    if verdict in VERDICTS:
        if "阻断" in levels and verdict != "不通过":
            errors.append("存在 level=阻断 的 finding，复核结论必须为「不通过」，收到「%s」" % verdict)
        if "阻断" not in levels and "警告" in levels and verdict == "通过":
            errors.append("存在 level=警告 的 finding，复核结论不得为「通过」"
                          "（应为「有条件通过」或「不通过」）")
        if verdict in ("不通过", "有条件通过") and not summary.strip():
            errors.append("复核结论为「%s」时 summary 不得为空" % verdict)
        if verdict == "通过" and checklist_bad:
            warnings.append("清单项 " + "、".join(checklist_bad)
                            + " 结果为「不通过」，但复核结论为「通过」，请复核判定是否一致")

    if "period" not in report:
        warnings.append("报告缺少 period 字段（按 null 处理）")

    if packet is not None:
        if not isinstance(packet, dict):
            errors.append("复核包 JSON 顶层必须是对象")
        else:
            if packet.get("schema") != SCHEMA_PACKET:
                errors.append("复核包 schema 必须为 %s，收到 %r"
                              % (SCHEMA_PACKET, packet.get("schema")))
            packet_sha = (packet.get("workbook") or {}).get("sha256") \
                if isinstance(packet.get("workbook"), dict) else None
            if packet_sha != current_sha:
                errors.append("复核包记录的工作簿哈希与当前工作簿不一致：复核包=%s / 当前=%s"
                              % (packet_sha, current_sha))
            report_start, report_end = _period_pair(report.get("period"))
            packet_start, packet_end = _period_pair(packet.get("period"))
            if report_start and packet_start and report_end and packet_end:
                if (report_start, report_end) != (packet_start, packet_end):
                    errors.append("报告 period 与复核包不一致：报告=%s 至 %s / 复核包=%s 至 %s"
                                  % (report_start, report_end, packet_start, packet_end))
            elif report.get("period") is not None and packet.get("period") is not None:
                warnings.append("报告 period 与复核包形态不同，未做一致性比对"
                                "（建议统一写 {\"start\": \"YYYY-MM-DD\", \"end\": \"YYYY-MM-DD\"}）")
    return errors, warnings


# ---------------------------------------------------------------- 归档 / 锁
def _archive_dir(workbook_path, report_dir):
    if report_dir:
        return os.path.abspath(report_dir)
    return os.path.join(os.path.dirname(os.path.abspath(workbook_path)), REVIEW_DIR)


def _archive_base_name(workbook_path, reviewed_at):
    stem = os.path.splitext(os.path.basename(workbook_path))[0]
    parsed = parse_iso8601(reviewed_at)
    stamp = parsed.strftime("%Y%m%dT%H%M%S") if parsed else "未标注时间"
    return "%s_复核报告_%s" % (stem, stamp)


def _archive_report(report, archive_dir, workbook_path):
    """把报告归档为 JSON+MD（同名同内容不重复写，保持幂等）；返回归档 JSON 文件名。"""
    base = _archive_base_name(workbook_path, report.get("reviewed_at"))
    json_text = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    md_text = report_to_md(report) + "\n"
    try:
        os.makedirs(archive_dir, exist_ok=True)
    except OSError as exc:
        raise ReviewError("[归档失败] 无法创建目录 %s\n原因：%s" % (archive_dir, exc))
    suffix = 1
    while True:
        name = base if suffix == 1 else "%s_%d" % (base, suffix)
        json_path = os.path.join(archive_dir, name + ".json")
        md_path = os.path.join(archive_dir, name + ".md")
        if not os.path.isfile(json_path):
            break
        try:
            with open(json_path, encoding="utf-8") as fh:
                existing = fh.read()
        except OSError:
            existing = None
        if existing == json_text:
            break
        suffix += 1
        if suffix > 99:
            raise ReviewError("[归档失败] %s 同名报告过多（>99），请清理 _复核记录/ 后重试" % archive_dir)
    try:
        with open(json_path, "w", encoding="utf-8") as fh:
            fh.write(json_text)
        with open(md_path, "w", encoding="utf-8") as fh:
            fh.write(md_text)
    except OSError as exc:
        raise ReviewError("[归档失败] %s\n原因：%s" % (json_path, exc))
    return name + ".json", json_path


def _write_lock(archive_dir, pre_sha, post_sha, verdict, report_name):
    lock = {
        "pre_sha256": pre_sha,
        "post_sha256": post_sha,
        "verdict": verdict,
        "report": report_name,
        "applied_at": iso_now(),
    }
    path = os.path.join(archive_dir, LOCK_NAME)
    try:
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(json.dumps(lock, ensure_ascii=False, indent=2) + "\n")
    except OSError as exc:
        raise ReviewError("[复核锁写入失败] %s\n原因：%s" % (path, exc))
    return lock, path


# ---------------------------------------------------------------- 幂等检测
def _is_idempotent(workbook_path, archive_dir, report, current_sha):
    """同一份报告已应用且工作簿自复核后未再变更 → 可安全跳过。"""
    if not isinstance(report, dict):
        return False
    lock_path = os.path.join(archive_dir, LOCK_NAME)
    if not os.path.isfile(lock_path):
        return False
    try:
        lock = load_json(lock_path)
    except (OSError, ValueError):
        return False
    if not isinstance(lock, dict):
        return False
    if lock.get("post_sha256") != current_sha:
        return False
    if lock.get("pre_sha256") != report.get("workbook_sha256"):
        return False
    if lock.get("verdict") != report.get("verdict"):
        return False
    archive_name = lock.get("report")
    if not isinstance(archive_name, str) or not archive_name.strip():
        return False
    archive_path = os.path.join(archive_dir, archive_name)
    if not os.path.isfile(archive_path):
        return False
    try:
        archived = load_json(archive_path)
    except (OSError, ValueError):
        return False
    if archived != report:
        return False
    try:
        wb = openpyxl.load_workbook(workbook_path, read_only=True)
    except Exception:  # noqa: BLE001
        return False
    try:
        return bool(_review_sheet_of(wb))
    finally:
        wb.close()


# ---------------------------------------------------------------- --apply 落盘
def _apply_to_workbook(workbook_path, report, pre_sha, archive_name, packet_name):
    """写/更新 08 表并原子替换工作簿；返回 post_sha256。"""
    workbook_path = os.path.abspath(workbook_path)
    try:
        wb = openpyxl.load_workbook(workbook_path, data_only=False)
    except Exception as exc:  # noqa: BLE001
        raise ReviewError("[无法读取工作簿] %s\n原因：%s" % (workbook_path, exc))
    directory = os.path.dirname(workbook_path)
    stem = os.path.splitext(os.path.basename(workbook_path))[0]
    tmp_path = os.path.join(directory, "." + stem + ".review-tmp.xlsx")
    try:
        write_review_sheet(wb, report, pre_sha, archive_name, packet_name,
                           os.path.basename(workbook_path))
        wb.save(tmp_path)
    except Exception as exc:  # noqa: BLE001 - 保存失败一律转成中文原因
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise ReviewError("[写入失败] 无法写入工作簿目录 %s\n原因：%s\n"
                          "请关闭 Excel/WPS 后重试；原工作簿未受影响。" % (directory, exc))
    finally:
        try:
            wb.close()
        except Exception:  # noqa: BLE001
            pass
    try:
        os.replace(tmp_path, workbook_path)
    except OSError as exc:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise ReviewError("[写入失败] 目标工作簿被占用或不可写：%s\n原因：%s\n"
                          "请关闭 Excel/WPS 后重试；原工作簿未受影响。" % (workbook_path, exc))
    try:
        return sha256_file(workbook_path)
    except OSError as exc:
        raise ReviewError("[无法读取工作簿] 保存后计算哈希失败：%s（%s）" % (workbook_path, exc))


# ---------------------------------------------------------------- --check
def do_check(workbook_path, archive_dir):
    """交付前校验；有效返回 0，否则打印中文原因返回 1。"""
    errors = []
    warnings = []
    lock_path = os.path.join(archive_dir, LOCK_NAME)
    lock = None
    if not os.path.isfile(lock_path):
        errors.append("缺少复核锁：%s（未执行 apply_review.py --apply，或复核记录被清理/移动）"
                      % lock_path)
    else:
        try:
            lock = load_json(lock_path)
        except (OSError, ValueError) as exc:
            errors.append("复核锁不是合法 JSON：%s（%s）" % (lock_path, exc))
            lock = None
        if lock is not None:
            if not isinstance(lock, dict):
                errors.append("复核锁顶层必须是 JSON 对象")
                lock = None
            else:
                missing = [key for key in ("pre_sha256", "post_sha256", "verdict", "report", "applied_at")
                           if key not in lock]
                if missing:
                    errors.append("复核锁缺少字段：" + "、".join(missing))
                    lock = None

    info = {}
    try:
        wb = openpyxl.load_workbook(workbook_path, data_only=False)
    except Exception as exc:  # noqa: BLE001
        errors.append("[无法读取工作簿] %s\n原因：%s" % (workbook_path, exc))
        wb = None
    if wb is not None:
        try:
            if not _review_sheet_of(wb):
                errors.append("工作簿缺少「%s」工作表（未执行 --apply，或该表已被删除）" % REVIEW_SHEET)
            else:
                info = _read_review_sheet(wb)
                if not info.get("pre_sha256"):
                    errors.append("08-复核记录缺少「被复核工作簿哈希（SHA-256）」记录")
                if not info.get("verdict"):
                    errors.append("08-复核记录缺少「复核结论」记录")
        finally:
            wb.close()

    try:
        current_sha = sha256_file(workbook_path)
    except OSError as exc:
        errors.append("[无法读取工作簿] %s（%s）" % (workbook_path, exc))
        current_sha = None

    if lock is not None:
        if current_sha is not None and current_sha != lock.get("post_sha256"):
            errors.append("工作簿在复核后被修改：当前 sha256=%s ≠ 复核锁 post_sha256=%s。"
                          "复核后任何保存（含用 Excel/WPS 打开另存）都会使复核失效，须重新复核。"
                          % (current_sha, lock.get("post_sha256")))
        if lock.get("verdict") not in DELIVERABLE_VERDICTS:
            errors.append("复核结论为「%s」，不得交付（只允许 通过 / 有条件通过）。"
                          % lock.get("verdict"))
        if info.get("pre_sha256") and lock.get("pre_sha256") \
                and info["pre_sha256"] != lock.get("pre_sha256"):
            errors.append("08-复核记录的被复核哈希与复核锁不一致：08=%s / 锁=%s（复核记录可能被改写）"
                          % (info["pre_sha256"], lock.get("pre_sha256")))
        if info.get("verdict") and lock.get("verdict") and info["verdict"] != lock.get("verdict"):
            errors.append("08-复核记录结论（%s）与复核锁（%s）不一致"
                          % (info["verdict"], lock.get("verdict")))
        report_name = lock.get("report")
        if isinstance(report_name, str) and report_name.strip():
            if not os.path.isfile(os.path.join(archive_dir, report_name)):
                warnings.append("复核锁引用的归档报告不存在：%s"
                                % os.path.join(archive_dir, report_name))

    for warning in warnings:
        print("[警告] " + warning, file=sys.stderr)
    if errors:
        print("[复核无效] 共 %d 项问题：" % len(errors), file=sys.stderr)
        for error in errors:
            print("  - " + error, file=sys.stderr)
        return 1
    print("复核有效：%s 复核结论「%s」，工作簿自复核后未被修改"
          "（post_sha256=%s，报告=%s）。"
          % (os.path.basename(workbook_path), lock["verdict"],
             str(lock["post_sha256"])[:16] + "…", lock["report"]))
    return 0


# ---------------------------------------------------------------- CLI
def _count_sheets(workbook_path):
    try:
        wb = openpyxl.load_workbook(workbook_path, read_only=True)
    except Exception:  # noqa: BLE001
        return None
    try:
        return len(wb.sheetnames)
    finally:
        wb.close()


def _print_dry_run(report, workbook_path, archive_dir, packet_name):
    verdict = report.get("verdict")
    findings = [item for item in (report.get("findings") or []) if isinstance(item, dict)]
    levels = [item.get("level") for item in findings]
    checklist = [item for item in (report.get("checklist") or []) if isinstance(item, dict)]
    print("[校验通过] 报告合法：verdict=%s，checklist=%d/10，findings=%d 条"
          "（阻断 %d / 警告 %d / 提示 %d）"
          % (verdict, len(checklist), len(findings),
             levels.count("阻断"), levels.count("警告"), levels.count("提示")))
    if verdict == "不通过":
        print("[dry-run] 结论为「不通过」：--apply 只归档报告（JSON+MD），不写 08 表，"
              "退出码 1（不得交付）。")
        print("[dry-run] 未写入任何文件。")
        return
    sheet_count = _count_sheets(workbook_path)
    after = "（%d 张 → %d 张）" % (sheet_count, sheet_count + 1) if sheet_count is not None else ""
    print("[dry-run] 将通过：写/更新「%s」%s、归档报告到 %s、保存后写复核锁 %s。"
          % (REVIEW_SHEET, after, archive_dir, os.path.join(archive_dir, LOCK_NAME)))
    if packet_name:
        print("[dry-run] 复核包：%s" % packet_name)
    print("[dry-run] 未写入任何文件；确认无误后加 --apply。")


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="L6 复核结论校验与落盘（交付门禁，dealer-settlement/review-report@1）")
    parser.add_argument("workbook", help="被复核的对账工作簿（.xlsx）")
    parser.add_argument("report", nargs="?", help="复核报告 JSON（由 LLM 按复核包产出；--check 时可省略）")
    parser.add_argument("--packet", help="复核包 JSON（提供时校验复核包与被复核文件哈希一致）")
    parser.add_argument("--report-dir", help="报告归档目录（默认 <工作簿目录>/_复核记录）")
    parser.add_argument("--apply", action="store_true", help="写入 08-复核记录并落盘复核锁")
    parser.add_argument("--check", action="store_true", help="交付前校验复核记录与复核锁是否有效")
    args = parser.parse_args(argv)

    if args.apply and args.check:
        print("[参数错误] --apply 与 --check 互斥，只能二选一。", file=sys.stderr)
        return 2
    if not args.check and not args.report:
        print("[参数错误] 缺少复核报告 JSON 路径（--check 模式可省略）。", file=sys.stderr)
        return 2

    workbook_path = os.path.abspath(args.workbook)
    if not os.path.isfile(workbook_path):
        print("[找不到工作簿] %s\n请确认路径（本脚本只处理 .xlsx，.xls 请先另存为 .xlsx）。"
              % workbook_path, file=sys.stderr)
        return 1
    try:
        current_sha = sha256_file(workbook_path)
    except OSError as exc:
        print("[无法读取工作簿] %s\n原因：%s" % (workbook_path, exc), file=sys.stderr)
        return 1

    archive_dir = _archive_dir(workbook_path, args.report_dir)
    if args.check:
        return do_check(workbook_path, archive_dir)

    try:
        report = load_json(os.path.abspath(args.report))
    except OSError as exc:
        print("[找不到报告] %s\n原因：%s" % (args.report, exc), file=sys.stderr)
        return 1
    except ValueError as exc:
        print("[JSON 解析失败] %s\n原因：%s\n请确认报告文件未被截断或手工改坏。"
              % (args.report, exc), file=sys.stderr)
        return 1

    packet = None
    packet_name = None
    if args.packet:
        packet_path = os.path.abspath(args.packet)
        try:
            packet = load_json(packet_path)
        except OSError as exc:
            print("[找不到复核包] %s\n原因：%s" % (args.packet, exc), file=sys.stderr)
            return 1
        except ValueError as exc:
            print("[JSON 解析失败] 复核包 %s\n原因：%s" % (args.packet, exc), file=sys.stderr)
            return 1
        packet_name = os.path.basename(packet_path)

    if _is_idempotent(workbook_path, archive_dir, report, current_sha):
        lock_path = os.path.join(archive_dir, LOCK_NAME)
        print("[幂等] 该复核报告已应用且工作簿自复核后未再变更"
              "（post_sha256=%s，复核锁=%s），跳过重复写入。" % (current_sha[:16] + "…", lock_path))
        return 0

    errors, warnings = validate_report(report, workbook_path, current_sha,
                                       packet=packet, packet_path=packet_name)
    for warning in warnings:
        print("[警告] " + warning, file=sys.stderr)
    if errors:
        print("[复核未通过校验] 报告不能应用，共 %d 项问题：" % len(errors), file=sys.stderr)
        for error in errors:
            print("  - " + error, file=sys.stderr)
        return 1

    verdict = report.get("verdict")
    if not args.apply:
        _print_dry_run(report, workbook_path, archive_dir, packet_name)
        return 0

    archive_name, archive_path = _archive_report(report, archive_dir, workbook_path)
    if verdict == "不通过":
        print("[已归档] %s" % archive_path)
        print("[阻断] 复核结论为「不通过」，不得交付：请修正工作簿后重新出账并重新复核。", file=sys.stderr)
        return 1

    post_sha = _apply_to_workbook(workbook_path, report, current_sha, archive_name, packet_name)
    lock, lock_path = _write_lock(archive_dir, current_sha, post_sha, verdict, archive_name)
    sheet_count = _count_sheets(workbook_path)
    print("[复核记录] 已写入「%s」（工作表 %s 张）"
          % (REVIEW_SHEET, sheet_count if sheet_count is not None else "?"))
    print("[归档] %s" % archive_path)
    print("[复核锁] %s（pre_sha256=%s post_sha256=%s）"
          % (lock_path, lock["pre_sha256"][:16] + "…", lock["post_sha256"][:16] + "…"))
    print("[完成] 复核有效（%s），可交付；复核后任何保存都会使复核锁失效。" % verdict)
    return 0


if __name__ == "__main__":
    sys.exit(main())
