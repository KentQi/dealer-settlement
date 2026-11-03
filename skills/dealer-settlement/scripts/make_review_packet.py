"""生成 L6 LLM 复核包（结构化底稿，token 友好）。

用法：
    python3 make_review_packet.py <工作簿.xlsx> [--master _主档.xlsx] [--out packet.json] [--md packet.md]

设计原则：
- 只做确定性取数（openpyxl 读单元格 + 标签关键词扫描），不重算、不联网、不调用外部模型；
- 取不到的数字一律写 null 并追加 warnings，绝不臆造 0；
- 若环境变量 RECON_CFG 有效，period / asof / sources / contract 直接取自配置，不重复推断；
- 除 generated_at 外，同一输入的输出完全一致（JSON 与 Markdown 均如此）。

复核包字段（schema=dealer-settlement/review-packet@1）：
    schema, generated_at, workbook{path,sha256,size}, master{path,sha256}|null,
    sources[{role,path,sha256}], period|null, asof|null, sheets[{name,rows,cols}],
    figures{opening_qty,opening_amount,settled_qty,settled_amount,diff_qty,diff_amount,
            addback_qty,addback_amount,tiehua_total,closing_inventory_value,period_end|null},
    sku_rows[{bar,name,opening_qty,settled_qty,settled_amount,diff_qty,unit_price}],
    tiehua_rows[{bar,name,settled_qty,standard,amount}],
    claims[{sheet,cell,text,numbers[]}], cross_refs[{sheet,cell,text,refs[]}],
    contract{no,method,basis,objection_days,signback_days,pay_day_text}|null,
    redline_scan[{id,pattern,hits[{sheet,cell,text}]}],
    checklist[{id,item}], warnings[]

失败：中文原因写 stderr，退出码 1（stdout 不输出半成品）。
"""
import argparse
import datetime
import json
import math
import os
import re
import sys

import openpyxl

try:  # 脚本目录即 sys.path[0]；作为模块被导入时兜底
    from review_common import (
        REVIEW_SHEET, SCHEMA_PACKET, REQUIRED_ITEMS, REVIEW_DIR,
        iso_now, sha256_file,
    )
except ImportError:  # pragma: no cover - 仅异常运行方式触发
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from review_common import (
        REVIEW_SHEET, SCHEMA_PACKET, REQUIRED_ITEMS, REVIEW_DIR,
        iso_now, sha256_file,
    )

CODE = chr(96)  # Markdown 行内代码标记；用常量而非字面反引号，便于源码审阅


def _code(text):
    return CODE + str(text) + CODE


# 取数只扫这几张表（与文档口径一致：00/01/02/05/06）
SCAN_SHEETS = ("00-对账说明", "01-对账单", "02-库存差异表", "05-乙方销售汇总", "06-乙方库存明细")
CLAIM_SHEETS = SCAN_SHEETS
CLAIM_MIN_LEN = 15          # 短于此长度的字符串多为标签/表头，不作为“文字断言”
MAX_CLAIMS = 200            # 超过时截断并告警，避免复核包无限膨胀

# 红线正则：命中只报告、不下结论（LLM 依据命中自行判断是否违规）
REDLINE_PATTERNS = (
    ("RL01", "初步判断|初步认定|初步分析|初步结论"),
    ("RL02", "漏报|瞒报|少报|隐瞒"),
    ("RL03", "故意|恶意|蓄意"),
    ("RL04", "责任|归责|追责|赔偿"),
    ("RL05", "实盘|盘点(?:结果|表|日)"),
    ("RL06", r"税率\s*(?:按|为|是)\s*[0-9]|[0-9]+\s*%\s*(?:税率|征收率)"),
    ("RL07", r"贴花[^。；\n]{0,6}(?:从货款中?(?:直接)?(?:扣除|抵扣)|直接抵扣)"),
    ("RL08", "必然|一定是|肯定是|毫无疑问"),
    ("RL09", "乙方(?:应|须)(?:承担|负责|赔偿)"),
    ("RL10", r"(?:非结算|陈列装)[^。；\n]{0,12}(?:计入|并入)[^。；\n]{0,6}结算"),
)

CROSS_REF_RE = re.compile(r"(?:详见)?附表\s*([0-9]{2}(?:\s*[、,，]\s*[0-9]{2})*)")
RANGE_RE = re.compile(r"([0-9]{4}-[0-9]{2}-[0-9]{2})\s*至\s*([0-9]{4}-[0-9]{2}-[0-9]{2})")
ISO_DATE_RE = re.compile(r"([0-9]{4})-([0-9]{2})-([0-9]{2})")
CN_DATE_RE = re.compile(r"([0-9]{4})\s*年\s*([0-9]{1,2})\s*月\s*([0-9]{1,2})\s*日")
NUMBER_RE = re.compile(r"(?<![0-9A-Za-z])[+-]?[0-9][0-9,]*(?:\.[0-9]+)?")
# 单号/编码里的数字不是业务数字，提取前剔除（先剔除含连字符的单号，再剔除含字母的编码）
CODE_STRIP_PATTERNS = (
    re.compile(r"[0-9A-Za-z]+(?:[-_/][0-9A-Za-z]+)+"),
    re.compile(r"[0-9A-Za-z]*[A-Za-z][0-9A-Za-z]*"),
)
BAR_RE = re.compile(r"[0-9]{8,14}")

# 提取数字前先剔除日期，避免把 2026-10-01 拆成 2026/-10/-01 混进 numbers[]
DATE_STRIP_PATTERNS = (
    re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}"),
    re.compile(r"[0-9]{4}/[0-9]{1,2}/[0-9]{1,2}"),
    re.compile(r"[0-9]{1,2}/[0-9]{1,2}"),
    re.compile(r"[0-9]{4}\s*年\s*[0-9]{1,2}\s*月(?:\s*[0-9]{1,2}\s*日)?"),
    re.compile(r"[0-9]{1,2}\s*月\s*[0-9]{1,2}\s*日"),
)


class PacketError(Exception):
    """复核包生成失败（中文原因由 main 打印到 stderr）。"""


# ---------------------------------------------------------------- 基础工具
def _norm(value):
    """去掉全部空白（含全角空格）后的文本，用于标签/表头匹配。"""
    if value is None:
        return ""
    return re.sub(r"[\s\u3000]+", "", str(value))


def _is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _to_num(value):
    """单元格值 → int/float；无法确定时返回 None（绝不默认 0）。"""
    if _is_number(value):
        number = float(value)
        if not math.isfinite(number):
            return None
        return int(number) if number == int(number) else number
    if isinstance(value, str):
        text = value.strip().replace(",", "").replace("，", "")
        if re.fullmatch(r"[+-]?[0-9]+(?:\.[0-9]+)?", text):
            number = float(text)
            return int(number) if number == int(number) else number
    return None


def _bar_text(value):
    """条码单元格 → 文本（Excel 可能把条码存成数字）。"""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, int):
        text = str(value)
    elif isinstance(value, float):
        if not math.isfinite(value) or value != int(value):
            return None
        text = str(int(value))
    else:
        text = str(value).strip()
    return text if BAR_RE.fullmatch(text) else None


def _first_date(value):
    """从文本里取第一个日期（ISO 或「YYYY 年 M 月 D 日」），规范成 YYYY-MM-DD。"""
    if not isinstance(value, str):
        return None
    match = ISO_DATE_RE.search(value)
    if match:
        try:
            return datetime.date(int(match.group(1)), int(match.group(2)),
                                 int(match.group(3))).isoformat()
        except ValueError:
            return None
    match = CN_DATE_RE.search(value)
    if match:
        try:
            return datetime.date(int(match.group(1)), int(match.group(2)),
                                 int(match.group(3))).isoformat()
        except ValueError:
            return None
    return None


def numbers_in(text):
    """从一句话里提取数字（剔除日期/单号/编码与列表序号；千分位/符号/小数都支持）。

    单号（HT-DEMO-2026-001）、税号（91440100DEMO00001A）里的数字不是文字断言要核对的
    业务数字，先剔除，避免 numbers[] 里出现 -2026、-1 之类的噪声。
    """
    if not isinstance(text, str):
        return []
    stripped = re.sub(r"^\s*[0-9]+[.、．]\s*", "", text)
    for pattern in DATE_STRIP_PATTERNS:
        stripped = pattern.sub(" ", stripped)
    for pattern in CODE_STRIP_PATTERNS:
        stripped = pattern.sub(" ", stripped)
    numbers = []
    for match in NUMBER_RE.finditer(stripped):
        raw = match.group(0).replace(",", "")
        try:
            number = float(raw)
        except ValueError:
            continue
        if not math.isfinite(number):
            continue
        numbers.append(int(number) if number == int(number) else number)
    return numbers


def _num_after(text, keyword):
    """取关键词之后出现的第一个数字（如「铺货数量 1,360 件」→ 1360）。"""
    if not isinstance(text, str):
        return None
    norm_text = _norm(text)
    norm_key = _norm(keyword)
    index = norm_text.find(norm_key)
    if index < 0:
        return None
    match = NUMBER_RE.search(norm_text[index + len(norm_key):])
    return _to_num(match.group(0)) if match else None


# ---------------------------------------------------------------- 工作表定位
def _sheet(wb, name):
    return wb[name] if name in wb.sheetnames else None


def _find_cell(ws, *keywords):
    """找第一个文本同时包含全部关键词的单元格，返回 (row, col, value)。"""
    norm_keys = [_norm(key) for key in keywords]
    for row in ws.iter_rows():
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            text = _norm(cell.value)
            if all(key in text for key in norm_keys):
                return cell.row, cell.column, cell.value
    return None


def _right(ws, row, col):
    """同一行右侧第一个非空单元格，返回 (value, coordinate)。"""
    for column in range(col + 1, ws.max_column + 1):
        cell = ws.cell(row=row, column=column)
        if cell.value is not None and str(cell.value).strip() != "":
            return cell.value, cell.coordinate
    return None, None


def _below(ws, row, col):
    """同一列下方第一个非空单元格，返回 (value, coordinate)。"""
    for row_index in range(row + 1, ws.max_row + 1):
        cell = ws.cell(row=row_index, column=col)
        if cell.value is not None and str(cell.value).strip() != "":
            return cell.value, cell.coordinate
    return None, None


def _header_cols(ws, row, keywords):
    """判断第 row 行是否是表头：每个关键词都能在某列找到；返回 {关键词: 列号}。"""
    cols = {}
    for column in range(1, ws.max_column + 1):
        value = ws.cell(row=row, column=column).value
        if not isinstance(value, str):
            continue
        text = _norm(value)
        for key in keywords:
            if key not in cols and _norm(key) in text:
                cols[key] = column
    return cols if len(cols) == len(keywords) else None


def _find_header(ws, keywords, max_scan=120):
    """向下扫描找表头行，返回 (row, {关键词: 列号})。"""
    for row in range(1, min(ws.max_row, max_scan) + 1):
        cols = _header_cols(ws, row, keywords)
        if cols:
            return row, cols
    return None, None


def _find_total(ws, start_row, first_col=1):
    """从 start_row 起找第一行「合计」行（首列文本以 合计 开头）。"""
    for row in range(start_row + 1, ws.max_row + 1):
        value = ws.cell(row=row, column=first_col).value
        if isinstance(value, str) and _norm(value).startswith("合计"):
            return row
    return None


def _cell_num(ws, row, col):
    return _to_num(ws.cell(row=row, column=col).value)


# ---------------------------------------------------------------- figures 取数
def _fig_opening_amount(wb):
    ws = _sheet(wb, "01-对账单")
    if ws is not None:
        hit = _find_cell(ws, "①", "期初未结算货值")
        if hit:
            value, coord = _right(ws, hit[0], hit[1])
            number = _to_num(value)
            if number is not None:
                return number, {"sheet": "01-对账单", "cell": coord}
    ws = _sheet(wb, "00-对账说明")
    if ws is not None:
        hit = _find_cell(ws, "期初未结算货值")
        if hit:
            value, coord = _below(ws, hit[0], hit[1])
            number = _to_num(value)
            if number is not None:
                return number, {"sheet": "00-对账说明", "cell": coord}
    return None, None


def _fig_opening_qty(wb):
    ws = _sheet(wb, "02-库存差异表")
    if ws is not None:
        header_row, cols = _find_header(ws, ("甲方期初", "未结算数量"))
        if header_row is not None:
            total_row = _find_total(ws, header_row)
            if total_row is not None:
                number = _cell_num(ws, total_row, cols["甲方期初"])
                if number is not None:
                    return number, {"sheet": "02-库存差异表",
                                    "cell": ws.cell(row=total_row, column=cols["甲方期初"]).coordinate}
    ws = _sheet(wb, "00-对账说明")
    if ws is not None:
        hit = _find_cell(ws, "铺货数量")
        if hit:
            number = _num_after(hit[2], "铺货数量")
            if number is not None:
                return number, {"sheet": "00-对账说明", "cell": ws.cell(row=hit[0], column=hit[1]).coordinate}
    return None, None


def _fig_settled_qty(wb):
    ws = _sheet(wb, "01-对账单")
    if ws is not None:
        header_row, cols = _find_header(ws, ("本期动销数量",))
        if header_row is not None:
            total_row = _find_total(ws, header_row)
            if total_row is not None:
                number = _cell_num(ws, total_row, cols["本期动销数量"])
                if number is not None:
                    return number, {"sheet": "01-对账单",
                                    "cell": ws.cell(row=total_row, column=cols["本期动销数量"]).coordinate}
    ws = _sheet(wb, "00-对账说明")
    if ws is not None:
        hit = _find_cell(ws, "结算数量")
        if hit:
            number = _num_after(hit[2], "结算数量")
            if number is not None:
                return number, {"sheet": "00-对账说明", "cell": ws.cell(row=hit[0], column=hit[1]).coordinate}
    return None, None


def _fig_settled_amount(wb):
    ws = _sheet(wb, "01-对账单")
    if ws is not None:
        header_row, cols = _find_header(ws, ("本期动销数量", "结算金额"))
        if header_row is not None:
            total_row = _find_total(ws, header_row)
            if total_row is not None:
                number = _cell_num(ws, total_row, cols["结算金额"])
                if number is not None:
                    return number, {"sheet": "01-对账单",
                                    "cell": ws.cell(row=total_row, column=cols["结算金额"]).coordinate}
    ws = _sheet(wb, "00-对账说明")
    if ws is not None:
        hit = _find_cell(ws, "本月应结算金额")
        if hit:
            value, coord = _below(ws, hit[0], hit[1])
            number = _to_num(value)
            if number is not None:
                return number, {"sheet": "00-对账说明", "cell": coord}
    return None, None


def _fig_diff_qty(wb):
    ws = _sheet(wb, "02-库存差异表")
    if ws is not None:
        header_row, cols = _find_header(ws, ("差异数量",))
        if header_row is not None:
            total_row = _find_total(ws, header_row)
            if total_row is not None:
                number = _cell_num(ws, total_row, cols["差异数量"])
                if number is not None:
                    return number, {"sheet": "02-库存差异表",
                                    "cell": ws.cell(row=total_row, column=cols["差异数量"]).coordinate}
    ws = _sheet(wb, "00-对账说明")
    if ws is not None:
        hit = _find_cell(ws, "占理论结存")
        if hit:
            numbers = numbers_in(hit[2])
            if numbers:
                return numbers[0], {"sheet": "00-对账说明",
                                    "cell": ws.cell(row=hit[0], column=hit[1]).coordinate}
    return None, None


def _fig_diff_amount(wb):
    ws = _sheet(wb, "01-对账单")
    if ws is not None:
        hit = _find_cell(ws, "⑥", "库存差异")
        if hit:
            value, coord = _right(ws, hit[0], hit[1])
            number = _to_num(value)
            if number is not None:
                return number, {"sheet": "01-对账单", "cell": coord}
    ws = _sheet(wb, "02-库存差异表")
    if ws is not None:
        header_row, cols = _find_header(ws, ("差异金额",))
        if header_row is not None:
            total_row = _find_total(ws, header_row)
            if total_row is not None:
                number = _cell_num(ws, total_row, cols["差异金额"])
                if number is not None:
                    return number, {"sheet": "02-库存差异表",
                                    "cell": ws.cell(row=total_row, column=cols["差异金额"]).coordinate}
    return None, None


def _fig_addback_qty(wb):
    ws = _sheet(wb, "02-库存差异表")
    if ws is not None:
        header_row, cols = _find_header(ws, ("加：11月", "销售"))
        if header_row is not None:
            total_row = _find_total(ws, header_row)
            if total_row is not None:
                number = _cell_num(ws, total_row, cols["加：11月"])
                if number is not None:
                    return number, {"sheet": "02-库存差异表",
                                    "cell": ws.cell(row=total_row, column=cols["加：11月"]).coordinate}
    ws = _sheet(wb, "00-对账说明")
    if ws is not None:
        hit = _find_cell(ws, "11月已销")
        if hit:
            number = _num_after(hit[2], "11月已销")
            if number is not None:
                return number, {"sheet": "00-对账说明", "cell": ws.cell(row=hit[0], column=hit[1]).coordinate}
    return None, None


def _fig_addback_amount(wb):
    ws = _sheet(wb, "01-对账单")
    if ws is not None:
        hit = _find_cell(ws, "④", "导出日前已销售")
        if hit:
            value, coord = _right(ws, hit[0], hit[1])
            number = _to_num(value)
            if number is not None:
                return number, {"sheet": "01-对账单", "cell": coord}
    ws = _sheet(wb, "02-库存差异表")
    if ws is not None:
        hit = _find_cell(ws, "共", "件", "元")
        if hit:
            match = re.search(r"共\s*([0-9][0-9,]*)\s*件\s*[/／]\s*([0-9][0-9,.]*)\s*元", hit[2])
            if match:
                number = _to_num(match.group(2))
                if number is not None:
                    return number, {"sheet": "02-库存差异表",
                                    "cell": ws.cell(row=hit[0], column=hit[1]).coordinate}
    return None, None


def _fig_tiehua_total(wb):
    ws = _sheet(wb, "01-对账单")
    if ws is not None:
        header_row, cols = _find_header(ws, ("贴花标准", "动销数量", "贴花金额"))
        if header_row is not None:
            total_row = _find_total(ws, header_row)
            if total_row is not None:
                number = _cell_num(ws, total_row, cols["贴花金额"])
                if number is not None:
                    return number, {"sheet": "01-对账单",
                                    "cell": ws.cell(row=total_row, column=cols["贴花金额"]).coordinate}
    ws = _sheet(wb, "00-对账说明")
    if ws is not None:
        hit = _find_cell(ws, "贴花费用")
        if hit:
            match = re.search(r"合计\s*([0-9][0-9,.]*)\s*元", hit[2])
            if match:
                number = _to_num(match.group(1))
                if number is not None:
                    return number, {"sheet": "00-对账说明",
                                    "cell": ws.cell(row=hit[0], column=hit[1]).coordinate}
    return None, None


def _fig_closing_inventory_value(wb):
    ws = _sheet(wb, "01-对账单")
    if ws is not None:
        hit = _find_cell(ws, "⑦", "期末乙方库存货值")
        if hit:
            value, coord = _right(ws, hit[0], hit[1])
            number = _to_num(value)
            if number is not None:
                return number, {"sheet": "01-对账单", "cell": coord}
    ws = _sheet(wb, "06-乙方库存明细")
    if ws is not None:
        header_row, cols = _find_header(ws, ("库存货值",))
        if header_row is not None:
            for row in range(header_row + 1, ws.max_row + 1):
                first = ws.cell(row=row, column=1).value
                if isinstance(first, str) and _norm(first).startswith("结算类小计"):
                    number = _cell_num(ws, row, cols["库存货值"])
                    if number is not None:
                        return number, {"sheet": "06-乙方库存明细",
                                        "cell": ws.cell(row=row, column=cols["库存货值"]).coordinate}
    return None, None


FIGURE_EXTRACTORS = (
    ("opening_qty", _fig_opening_qty),
    ("opening_amount", _fig_opening_amount),
    ("settled_qty", _fig_settled_qty),
    ("settled_amount", _fig_settled_amount),
    ("diff_qty", _fig_diff_qty),
    ("diff_amount", _fig_diff_amount),
    ("addback_qty", _fig_addback_qty),
    ("addback_amount", _fig_addback_amount),
    ("tiehua_total", _fig_tiehua_total),
    ("closing_inventory_value", _fig_closing_inventory_value),
)


# ---------------------------------------------------------------- 逐 SKU / 贴花
def _scan_sku_rows(wb, warnings):
    """从 02-库存差异表「一、差异明细」+「二、核对相符明细」取逐 SKU 数量。

    settled_amount 在 02 表没有列，按品名从 01 表结算明细回填（找不到写 null）。
    """
    rows = []
    ws = _sheet(wb, "02-库存差异表")
    if ws is None:
        warnings.append("未找到「02-库存差异表」，sku_rows 置空")
        return rows
    settled_amounts = _scan_settled_amounts(wb)
    header_rows = []
    for row in range(1, min(ws.max_row, 200) + 1):
        cols = _header_cols(ws, row, ("商品条码", "商品名称", "结算单价",
                                      "甲方期初", "申报动销", "差异数量"))
        if cols:
            header_rows.append((row, cols))
    if not header_rows:
        warnings.append("02-库存差异表未识别到逐 SKU 明细表头，sku_rows 置空")
        return rows
    for header_row, cols in header_rows:
        row = header_row + 1
        while row <= ws.max_row:
            serial = ws.cell(row=row, column=1).value
            if not (_is_number(serial) and float(serial) == int(float(serial))):
                break
            bar = _bar_text(ws.cell(row=row, column=cols["商品条码"]).value)
            name = ws.cell(row=row, column=cols["商品名称"]).value
            if not isinstance(name, str) or not name.strip():
                name = None
            else:
                name = name.strip()
            if bar is None:
                warnings.append("02-库存差异表第 %d 行商品条码缺失或不是 8–14 位数字（sku_rows 中 bar 置 null）" % row)
            opening_qty = _cell_num(ws, row, cols["甲方期初"])
            settled_raw = _cell_num(ws, row, cols["申报动销"])
            settled_qty = None if settled_raw is None else -settled_raw
            diff_qty = _cell_num(ws, row, cols["差异数量"])
            unit_price = _cell_num(ws, row, cols["结算单价"])
            settled_amount = None
            if name is not None:
                settled_amount = settled_amounts.get(_norm(name))
                if settled_amount is None:
                    warnings.append("02-库存差异表第 %d 行（%s）在 01-对账单结算明细中未匹配到金额，"
                                    "settled_amount 置 null" % (row, name))
            rows.append({
                "bar": bar,
                "name": name,
                "opening_qty": opening_qty,
                "settled_qty": settled_qty,
                "settled_amount": settled_amount,
                "diff_qty": diff_qty,
                "unit_price": unit_price,
            })
            row += 1
    if not rows:
        warnings.append("02-库存差异表未取到任何逐 SKU 行，sku_rows 置空")
    return rows


def _scan_settled_amounts(wb):
    """01-对账单「一、本期结算明细」→ {规范品名: 结算金额}。"""
    amounts = {}
    ws = _sheet(wb, "01-对账单")
    if ws is None:
        return amounts
    header_row, cols = _find_header(ws, ("商品名称", "本期动销数量", "结算金额"))
    if header_row is None:
        return amounts
    total_row = _find_total(ws, header_row)
    end_row = total_row if total_row is not None else ws.max_row + 1
    for row in range(header_row + 1, end_row):
        serial = ws.cell(row=row, column=1).value
        if not _is_number(serial):
            continue
        name = ws.cell(row=row, column=cols["商品名称"]).value
        amount = _cell_num(ws, row, cols["结算金额"])
        if isinstance(name, str) and name.strip() and amount is not None:
            amounts[_norm(name)] = amount
    return amounts


def _scan_name_bars(wb):
    """从 02 / 06 表建立 {规范品名: 条码}，用于贴花行回填条码。"""
    mapping = {}
    for sheet_name in ("02-库存差异表", "06-乙方库存明细"):
        ws = _sheet(wb, sheet_name)
        if ws is None:
            continue
        header_row, cols = _find_header(ws, ("商品条码", "商品名称"))
        if header_row is None:
            continue
        for row in range(header_row + 1, ws.max_row + 1):
            bar = _bar_text(ws.cell(row=row, column=cols["商品条码"]).value)
            name = ws.cell(row=row, column=cols["商品名称"]).value
            if bar and isinstance(name, str) and name.strip():
                mapping.setdefault(_norm(name), bar)
    return mapping


def _scan_tiehua_rows(wb, warnings):
    rows = []
    ws = _sheet(wb, "01-对账单")
    if ws is None:
        warnings.append("未找到「01-对账单」，tiehua_rows 置空")
        return rows
    header_row, cols = _find_header(ws, ("商品名称", "动销数量", "贴花标准", "贴花金额"))
    if header_row is None:
        warnings.append("01-对账单未识别到贴花明细表头（三、贴花费用），tiehua_rows 置空")
        return rows
    name_bars = _scan_name_bars(wb)
    total_row = _find_total(ws, header_row)
    end_row = total_row if total_row is not None else ws.max_row + 1
    for row in range(header_row + 1, end_row):
        serial = ws.cell(row=row, column=1).value
        if not _is_number(serial):
            continue
        name = ws.cell(row=row, column=cols["商品名称"]).value
        if isinstance(name, str):
            name = name.strip() or None
        else:
            name = None
        bar = name_bars.get(_norm(name)) if name else None
        if name and bar is None:
            warnings.append("01-对账单贴花明细第 %d 行（%s）未匹配到商品条码，bar 置 null" % (row, name))
        rows.append({
            "bar": bar,
            "name": name,
            "settled_qty": _cell_num(ws, row, cols["动销数量"]),
            "standard": _cell_num(ws, row, cols["贴花标准"]),
            "amount": _cell_num(ws, row, cols["贴花金额"]),
        })
    if not rows:
        warnings.append("01-对账单未取到任何贴花明细行，tiehua_rows 置空")
    return rows


# ---------------------------------------------------------------- 文字断言 / 交叉引用 / 红线
def _scan_claims(wb, skip_texts, warnings):
    claims = []
    truncated = False
    for sheet_name in CLAIM_SHEETS:
        ws = _sheet(wb, sheet_name)
        if ws is None:
            continue
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if not isinstance(value, str):
                    continue
                text = value.strip()
                if not text or "\n" in text or text.startswith("="):
                    continue
                if len(text) < CLAIM_MIN_LEN:
                    continue
                if _norm(text) in skip_texts:
                    continue
                numbers = numbers_in(text)
                if not numbers:
                    continue
                if len(claims) >= MAX_CLAIMS:
                    truncated = True
                    break
                claims.append({"sheet": sheet_name, "cell": cell.coordinate,
                               "text": text, "numbers": numbers})
            if truncated:
                break
        if truncated:
            break
    if truncated:
        warnings.append("文字断言超过 %d 条，已截断（其余请直接读工作簿）" % MAX_CLAIMS)
    if not claims:
        warnings.append("未扫描到任何含数字的文字断言（claims 置空）")
    return claims


def _scan_cross_refs(wb, warnings):
    refs = []
    prefix_map = {}
    for name in wb.sheetnames:
        match = re.match(r"([0-9]{2})-", name)
        if match:
            prefix_map[match.group(1)] = name
    missing = []
    for ws in wb.worksheets:
        if ws.title == REVIEW_SHEET:
            continue
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                found = CROSS_REF_RE.findall(cell.value)
                if not found:
                    continue
                numbers = []
                for group in found:
                    for number in re.findall(r"[0-9]{2}", group):
                        if number not in numbers:
                            numbers.append(number)
                targets = []
                for number in numbers:
                    target = prefix_map.get(number)
                    if target is None:
                        target = "附表 %s" % number
                        missing.append(number)
                    if target not in targets:
                        targets.append(target)
                refs.append({"sheet": ws.title, "cell": cell.coordinate,
                             "text": cell.value.strip(), "refs": targets})
    if missing:
        warnings.append("交叉引用指向的工作表不存在：附表 " + "、".join(sorted(set(missing))))
    return refs


def _scan_redline(wb):
    scan = []
    for redline_id, pattern in REDLINE_PATTERNS:
        hits = []
        regex = re.compile(pattern)
        for ws in wb.worksheets:
            if ws.title == REVIEW_SHEET:
                continue
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and regex.search(cell.value):
                        hits.append({"sheet": ws.title, "cell": cell.coordinate,
                                     "text": cell.value.strip()})
        scan.append({"id": redline_id, "pattern": pattern, "hits": hits})
    return scan


# ---------------------------------------------------------------- 期间 / 合同 / 来源
def _scan_period(wb):
    for sheet_name in ("00-对账说明", "01-对账单", "02-库存差异表"):
        ws = _sheet(wb, sheet_name)
        if ws is None:
            continue
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    match = RANGE_RE.search(cell.value)
                    if match:
                        return {"start": match.group(1), "end": match.group(2)}
    return None


def _scan_asof(wb):
    ws = _sheet(wb, "00-对账说明")
    if ws is not None:
        hit = _find_cell(ws, "数据基准日")
        if hit:
            value, _coord = _right(ws, hit[0], hit[1])
            date_text = _first_date(value if isinstance(value, str) else None)
            if date_text:
                return date_text
    for sheet_name in ("00-对账说明", "01-对账单", "02-库存差异表"):
        ws = _sheet(wb, sheet_name)
        if ws is None:
            continue
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                if "基准日" in cell.value or "制单日期" in cell.value:
                    date_text = _first_date(cell.value)
                    if date_text:
                        return date_text
    return None


def _read_cfg(warnings):
    raw = os.environ.get("RECON_CFG")
    if not raw or not raw.strip():
        return None
    try:
        cfg = json.loads(raw)
    except (ValueError, TypeError):
        warnings.append("环境变量 RECON_CFG 不是合法 JSON，已忽略；period/asof 改为从工作簿扫描，"
                        "sources/contract 置空")
        return None
    if not isinstance(cfg, dict):
        warnings.append("环境变量 RECON_CFG 不是 JSON 对象，已忽略")
        return None
    return cfg


def _cfg_date(value):
    if not isinstance(value, str):
        return None
    return _first_date(value) if ISO_DATE_RE.fullmatch(value.strip()) else None


def _resolve_period(cfg, wb, warnings):
    if isinstance(cfg, dict) and isinstance(cfg.get("period"), dict):
        start = _cfg_date(cfg["period"].get("start"))
        end = _cfg_date(cfg["period"].get("end"))
        if start and end:
            return {"start": start, "end": end}
        warnings.append("RECON_CFG.period 缺少合法的 start/end，period 改为从工作簿扫描")
    period = _scan_period(wb)
    if period is not None:
        return period
    warnings.append("未找到对账期间（00/01/02 均无「YYYY-MM-DD 至 YYYY-MM-DD」），period 置 null")
    return None


def _resolve_asof(cfg, wb, warnings):
    if isinstance(cfg, dict) and isinstance(cfg.get("period"), dict):
        asof = _cfg_date(cfg["period"].get("asof"))
        if asof:
            return asof
        warnings.append("RECON_CFG.period.asof 缺失或非法，asof 改为从工作簿扫描")
    asof = _scan_asof(wb)
    if asof is not None:
        return asof
    warnings.append("未找到数据基准日（asof 置 null）")
    return None


def _resolve_contract(cfg, warnings):
    if not isinstance(cfg, dict):
        warnings.append("未提供有效 RECON_CFG，contract 置 null（合同条款须人工核对）")
        return None
    contract = cfg.get("contract")
    if not isinstance(contract, dict) or not any(contract.get(key) for key in
                                                ("no", "method", "basis", "objection_days",
                                                 "signback_days", "pay_day_text")):
        warnings.append("RECON_CFG 缺少 contract 信息，contract 置 null")
        return None
    return {key: contract.get(key) for key in
            ("no", "method", "basis", "objection_days", "signback_days", "pay_day_text")}


def _resolve_sources(cfg, warnings):
    sources = []
    if not isinstance(cfg, dict):
        warnings.append("未提供有效 RECON_CFG，sources 置空（无法绑定源文件哈希）")
        return sources
    base_dir = cfg.get("base_dir")
    source_names = cfg.get("sources")
    if not isinstance(base_dir, str) or not isinstance(source_names, dict):
        warnings.append("RECON_CFG 缺少 base_dir/sources，sources 置空")
        return sources
    for role in ("jia", "sales", "inventory"):
        name = source_names.get(role)
        if not isinstance(name, str) or not name.strip():
            warnings.append("RECON_CFG.sources 缺少 role=%s，sources 中该角色 path/sha256 置 null" % role)
            sources.append({"role": role, "path": None, "sha256": None})
            continue
        path = os.path.abspath(os.path.join(base_dir, name.strip()))
        digest = None
        if os.path.isfile(path):
            try:
                digest = sha256_file(path)
            except OSError as exc:
                warnings.append("源文件无法读取（role=%s）：%s（%s）" % (role, path, exc))
        else:
            warnings.append("源文件不存在（role=%s）：%s" % (role, path))
        sources.append({"role": role, "path": path, "sha256": digest})
    return sources


def _resolve_master(explicit_master, cfg, warnings):
    master_path = explicit_master
    if not master_path and isinstance(cfg, dict):
        master_path = cfg.get("master_path")
    if not master_path:
        warnings.append("未提供主档（--master 未指定且 RECON_CFG 无 master_path），master 置 null")
        return None
    path = os.path.abspath(master_path)
    if not os.path.isfile(path):
        if explicit_master:
            raise PacketError("[找不到主档] %s\n请确认 --master 路径，或先跑 load_master.py 生成配置。"
                              % path)
        warnings.append("主档不存在（RECON_CFG.master_path 已失效）：%s，master 置 null" % path)
        return None
    try:
        digest = sha256_file(path)
    except OSError as exc:
        raise PacketError("[无法读取主档] %s\n原因：%s" % (path, exc))
    return {"path": path, "sha256": digest}


def _count_formula_cells(path):
    try:
        wb = openpyxl.load_workbook(path, data_only=False)
    except Exception:  # noqa: BLE001 - 只是辅助告警，读不到就跳过
        return 0
    count = 0
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        count += 1
    finally:
        wb.close()
    return count


# ---------------------------------------------------------------- 组装
def build_packet(workbook_path, master_path=None):
    """读取工作簿与配置，返回复核包 dict（含 warnings）。"""
    workbook_abs = os.path.abspath(workbook_path)
    if not os.path.isfile(workbook_abs):
        raise PacketError("[找不到工作簿] %s\n请确认路径（本脚本只读 .xlsx，.xls 请先另存为 .xlsx）。"
                          % workbook_abs)
    try:
        workbook_sha = sha256_file(workbook_abs)
        workbook_size = os.path.getsize(workbook_abs)
    except OSError as exc:
        raise PacketError("[无法读取工作簿] %s\n原因：%s" % (workbook_abs, exc))
    try:
        wb = openpyxl.load_workbook(workbook_abs, data_only=True)
    except Exception as exc:  # noqa: BLE001 - openpyxl 抛出的异常类型不稳定
        raise PacketError("[无法读取工作簿] %s\n原因：%s\n工作簿须为 .xlsx（.xls 请先另存为 .xlsx）。"
                          % (workbook_abs, exc))

    warnings = []
    try:
        cfg = _read_cfg(warnings)
        period = _resolve_period(cfg, wb, warnings)
        asof = _resolve_asof(cfg, wb, warnings)
        master = _resolve_master(master_path, cfg, warnings)
        sources = _resolve_sources(cfg, warnings)
        contract = _resolve_contract(cfg, warnings)

        figures = {}
        for key, extractor in FIGURE_EXTRACTORS:
            value, _source = extractor(wb)
            if value is None:
                warnings.append("未取到 figures.%s（按标签扫描 00/01/02/06 未命中），已置 null" % key)
            figures[key] = value
        figures["period_end"] = period["end"] if period else None

        sku_rows = _scan_sku_rows(wb, warnings)
        tiehua_rows = _scan_tiehua_rows(wb, warnings)

        skip_texts = set()
        for row in sku_rows:
            if row.get("name"):
                skip_texts.add(_norm(row["name"]))
        for row in tiehua_rows:
            if row.get("name"):
                skip_texts.add(_norm(row["name"]))
        claims = _scan_claims(wb, skip_texts, warnings)
        cross_refs = _scan_cross_refs(wb, warnings)
        redline_scan = _scan_redline(wb)

        formula_count = _count_formula_cells(workbook_abs)
        if formula_count:
            warnings.append("工作簿含 %d 个公式单元格；复核包按缓存值取数，无缓存值的公式会表现为 null"
                            "（请用 Excel 打开并保存一次后重新生成复核包）" % formula_count)

        if REVIEW_SHEET in wb.sheetnames:
            warnings.append("工作簿已含「%s」表；复核包基于当前文件生成，重新复核请确认被复核版本"
                            % REVIEW_SHEET)

        sheets = [{"name": ws.title, "rows": ws.max_row, "cols": ws.max_column}
                  for ws in wb.worksheets]
    finally:
        wb.close()

    packet = {
        "schema": SCHEMA_PACKET,
        "generated_at": iso_now(),
        "workbook": {"path": workbook_abs, "sha256": workbook_sha, "size": workbook_size},
        "master": master,
        "sources": sources,
        "period": period,
        "asof": asof,
        "sheets": sheets,
        "figures": figures,
        "sku_rows": sku_rows,
        "tiehua_rows": tiehua_rows,
        "claims": claims,
        "cross_refs": cross_refs,
        "contract": contract,
        "redline_scan": redline_scan,
        "checklist": [{"id": item_id, "item": item} for item_id, item in REQUIRED_ITEMS],
        "warnings": warnings,
    }
    return packet


# ---------------------------------------------------------------- Markdown
def _fmt(value):
    if value is None:
        return "—"
    if isinstance(value, float) and value != int(value):
        return "{:,.2f}".format(value)
    if isinstance(value, int):
        return "{:,}".format(value)
    return str(value)


def _fmt_numbers(numbers):
    return "、".join(_fmt(number) for number in numbers)


def render_md(packet):
    """复核包 → Markdown（确定性：不含 generated_at）。"""
    lines = []
    workbook = packet["workbook"]
    lines.append("# 复核包 · %s" % os.path.basename(workbook["path"]))
    lines.append("")
    lines.append("- schema：%s" % _code(packet["schema"]))
    lines.append("- 工作簿：%s（sha256 %s，%s 字节）"
                 % (_code(workbook["path"]), _code(workbook["sha256"]), _fmt(workbook["size"])))
    master = packet["master"]
    if master:
        lines.append("- 主档：%s（sha256 %s）" % (_code(master["path"]), _code(master["sha256"])))
    else:
        lines.append("- 主档：未提供")
    period = packet["period"]
    if period:
        lines.append("- 期间：%s 至 %s；数据基准日：%s"
                     % (period["start"], period["end"], packet["asof"] or "未识别"))
    else:
        lines.append("- 期间：未识别；数据基准日：%s" % (packet["asof"] or "未识别"))
    contract = packet["contract"]
    if contract:
        lines.append("- 合同：编号 %s；结算方式 %s；结算基数 %s；异议期 %s 个工作日；"
                     "签回期 %s 个工作日；付款日 %s"
                     % (contract.get("no") or "—", contract.get("method") or "—",
                        contract.get("basis") or "—", _fmt(contract.get("objection_days")),
                        _fmt(contract.get("signback_days")), contract.get("pay_day_text") or "—"))
    else:
        lines.append("- 合同：未从 RECON_CFG 取到（置 null）")
    lines.append("- 源文件：")
    if packet["sources"]:
        for source in packet["sources"]:
            lines.append("  - %s：%s（sha256 %s）"
                         % (source["role"], _code(source["path"] or "—"),
                            _code(source["sha256"] or "—")))
    else:
        lines.append("  - （无）")
    lines.append("- 工作表（%d）：%s" % (len(packet["sheets"]), "；".join(
        "%s %d×%d" % (sheet["name"], sheet["rows"], sheet["cols"]) for sheet in packet["sheets"])))
    lines.append("")

    lines.append("## 一、关键数（figures）")
    lines.append("")
    lines.append("| 字段 | 值 |")
    lines.append("|---|---|")
    for key, _extractor in FIGURE_EXTRACTORS:
        lines.append("| %s | %s |" % (key, _fmt(packet["figures"].get(key))))
    lines.append("| period_end | %s |" % (packet["figures"].get("period_end") or "—"))
    lines.append("")

    lines.append("## 二、逐 SKU（sku_rows，%d 行）" % len(packet["sku_rows"]))
    lines.append("")
    if packet["sku_rows"]:
        lines.append("| 条码 | 名称 | 期初数量 | 结算数量 | 结算金额 | 差异数量 | 单价 |")
        lines.append("|---|---|---|---|---|---|---|")
        for row in packet["sku_rows"]:
            lines.append("| %s | %s | %s | %s | %s | %s | %s |"
                         % (row["bar"] or "—", row["name"] or "—",
                            _fmt(row["opening_qty"]), _fmt(row["settled_qty"]),
                            _fmt(row["settled_amount"]), _fmt(row["diff_qty"]),
                            _fmt(row["unit_price"])))
    else:
        lines.append("（未取到）")
    lines.append("")

    lines.append("## 三、贴花明细（tiehua_rows，%d 行）" % len(packet["tiehua_rows"]))
    lines.append("")
    if packet["tiehua_rows"]:
        lines.append("| 条码 | 名称 | 结算数量 | 标准 | 金额 |")
        lines.append("|---|---|---|---|---|")
        for row in packet["tiehua_rows"]:
            lines.append("| %s | %s | %s | %s | %s |"
                         % (row["bar"] or "—", row["name"] or "—",
                            _fmt(row["settled_qty"]), _fmt(row["standard"]), _fmt(row["amount"])))
    else:
        lines.append("（未取到）")
    lines.append("")

    lines.append("## 四、文字断言（claims，%d 条；供 R5 核对）" % len(packet["claims"]))
    lines.append("")
    for index, claim in enumerate(packet["claims"], 1):
        lines.append("%d. %s：%s（数字：%s）"
                     % (index, _code(claim["sheet"] + "!" + claim["cell"]), claim["text"],
                        _fmt_numbers(claim["numbers"])))
    if not packet["claims"]:
        lines.append("（未取到）")
    lines.append("")

    lines.append("## 五、交叉引用（cross_refs，%d 条；供 R6 核对）" % len(packet["cross_refs"]))
    lines.append("")
    for ref in packet["cross_refs"]:
        lines.append("- %s：%s → %s"
                     % (_code(ref["sheet"] + "!" + ref["cell"]), ref["text"], "、".join(ref["refs"])))
    if not packet["cross_refs"]:
        lines.append("（未取到）")
    lines.append("")

    hit_count = sum(1 for item in packet["redline_scan"] if item["hits"])
    lines.append("## 六、红线扫描（redline_scan，%d 类命中；命中只报告不下结论）" % hit_count)
    lines.append("")
    for item in packet["redline_scan"]:
        if not item["hits"]:
            continue
        lines.append("- %s /%s/：%d 处" % (_code(item["id"]), item["pattern"], len(item["hits"])))
        for hit in item["hits"]:
            lines.append("  - %s：%s"
                         % (_code(hit["sheet"] + "!" + hit["cell"]), hit["text"]))
    if not hit_count:
        lines.append("（无命中）")
    lines.append("")

    lines.append("## 七、R1–R10 复核清单")
    lines.append("")
    for item in packet["checklist"]:
        lines.append("- %s %s" % (item["id"], item["item"]))
    lines.append("")

    lines.append("## 八、warnings（%d 条）" % len(packet["warnings"]))
    lines.append("")
    for warning in packet["warnings"]:
        lines.append("- %s" % warning)
    if not packet["warnings"]:
        lines.append("（无）")
    lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------- CLI
def _default_paths(workbook_abs, out, md):
    directory = os.path.join(os.path.dirname(workbook_abs), REVIEW_DIR)
    stem = os.path.splitext(os.path.basename(workbook_abs))[0]
    out_path = os.path.abspath(out) if out else os.path.join(directory, stem + "_复核包.json")
    if md:
        md_path = os.path.abspath(md)
    elif out:
        md_path = os.path.splitext(out_path)[0] + ".md"
    else:
        md_path = os.path.join(directory, stem + "_复核包.md")
    return out_path, md_path


def _write_text(path, text):
    try:
        directory = os.path.dirname(path)
        if directory:
            os.makedirs(directory, exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(text)
    except OSError as exc:
        raise PacketError("[写入失败] %s\n原因：%s" % (path, exc))


def _print_summary(packet, json_path, md_path):
    figures = packet["figures"]

    def _show(value):
        return "null" if value is None else _fmt(value)

    period = packet["period"]
    period_text = ("%s 至 %s（基准日 %s）" % (period["start"], period["end"],
                                            packet["asof"] or "未识别")) if period else "未识别"
    redline_hits = sum(1 for item in packet["redline_scan"] if item["hits"])
    print("[复核包] %s" % packet["schema"])
    print("  工作簿  : %s（sha256=%s…，%s 字节）"
          % (packet["workbook"]["path"], packet["workbook"]["sha256"][:16],
             _fmt(packet["workbook"]["size"])))
    print("  期间    : %s" % period_text)
    print("  关键数  : 期初 %s 件 / %s 元；结算 %s 件 / %s 元；差异 %s 件 / %s 元；"
          "回加 %s 件 / %s 元；贴花 %s 元；期末库存 %s 元"
          % (_show(figures["opening_qty"]), _show(figures["opening_amount"]),
             _show(figures["settled_qty"]), _show(figures["settled_amount"]),
             _show(figures["diff_qty"]), _show(figures["diff_amount"]),
             _show(figures["addback_qty"]), _show(figures["addback_amount"]),
             _show(figures["tiehua_total"]), _show(figures["closing_inventory_value"])))
    print("  明细    : sku_rows %d 行 / tiehua_rows %d 行 / claims %d 条 / cross_refs %d 条 / 红线命中 %d 类"
          % (len(packet["sku_rows"]), len(packet["tiehua_rows"]),
             len(packet["claims"]), len(packet["cross_refs"]), redline_hits))
    print("  警告    : %d 条" % len(packet["warnings"]))
    print("  JSON    : %s" % json_path)
    print("  MD      : %s" % md_path)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="生成 L6 LLM 复核包（dealer-settlement/review-packet@1）")
    parser.add_argument("workbook", help="被复核的对账工作簿（.xlsx）")
    parser.add_argument("--master", help="客户主档 _主档.xlsx（默认取 RECON_CFG.master_path）")
    parser.add_argument("--out", help="复核包 JSON 输出路径（默认 <工作簿目录>/_复核记录/<工作簿名>_复核包.json）")
    parser.add_argument("--md", help="复核包 Markdown 输出路径（默认与 JSON 同名 .md）")
    args = parser.parse_args(argv)
    try:
        packet = build_packet(args.workbook, args.master)
        out_path, md_path = _default_paths(packet["workbook"]["path"], args.out, args.md)
        _write_text(out_path, json.dumps(packet, ensure_ascii=False, indent=2) + "\n")
        _write_text(md_path, render_md(packet) + "\n")
    except PacketError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    _print_summary(packet, out_path, md_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
