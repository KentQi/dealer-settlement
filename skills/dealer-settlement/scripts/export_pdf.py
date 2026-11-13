#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把指定的对账工作表合成一份 PDF，供乙方打印盖章。

- 默认导出 00 / 01 / 02 / 08 四张打印表，一张表从新页开始
- 表头重复、签章区保留、金额与日期按原格式呈现
- 中文字体自动发现；找不到时明确报错并要求 --font 指定
- 纯本地处理，不上传任何数据
"""
import argparse
import datetime
import os
import sys

import openpyxl
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, Frame, PageBreak, PageTemplate,
                                Paragraph, Spacer, Table, TableStyle)

FONT_CANDIDATES = [
    ("/Library/Fonts/Arial Unicode.ttf", 0),
    ("/System/Library/Fonts/STHeiti Medium.ttc", 0),
    ("/System/Library/Fonts/Supplemental/Songti.ttc", 0),
    ("/System/Library/Fonts/PingFang.ttc", 0),
    ("C:/Windows/Fonts/msyh.ttc", 0),
    ("C:/Windows/Fonts/simsun.ttc", 0),
    ("C:/Windows/Fonts/simhei.ttf", 0),
    ("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc", 0),
    ("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc", 0),
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 0),
]

INK = colors.HexColor("#1F3864")
GREY = colors.HexColor("#A6A6A6")
LIGHT = colors.HexColor("#F2F2F2")


def register_font(explicit=""):
    cands = [(explicit, 0)] if explicit else FONT_CANDIDATES
    for path, idx in cands:
        if not path or not os.path.exists(path):
            continue
        try:
            pdfmetrics.registerFont(TTFont("DS_CN", path, subfontIndex=idx))
            return path
        except Exception:
            continue
    sys.stderr.write(
        "[中止] 找不到可用的中文字体，无法生成 PDF。\n"
        "请把本机一个 TTF/TTC 中文字体的完整路径传给 --font，例如：\n"
        "  macOS  : --font /Library/Fonts/Arial\\ Unicode.ttf\n"
        "  Windows: --font C:/Windows/Fonts/msyh.ttc\n")
    raise SystemExit(1)


def _fmt(value, cell):
    if value is None:
        return ""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, (int, float)):
        nf = cell.number_format or ""
        if ".00" in nf or "0.0" in nf:
            return format(float(value), ",.2f")
        if isinstance(value, float) and not value.is_integer():
            return format(value, ",.2f")
        return format(int(value), ",")
    return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _width(text):
    return sum(2 if ord(ch) > 127 else 1 for ch in str(text))


def _col_widths(ws, avail):
    caps = []
    for c in range(1, ws.max_column + 1):
        longest = 0
        for r in range(1, min(ws.max_row, 80) + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value is None:
                continue
            longest = max(longest, min(_width(_fmt(cell.value, cell)), 26))
        caps.append(max(4.0, float(longest)))
    total = sum(caps) or 1
    raw = [avail * x / total for x in caps]
    lo, hi = 26.0, 150.0
    out = [min(hi, max(lo, x)) for x in raw]
    scale = avail / sum(out)
    return [x * scale for x in out]


def _sheet_table(ws, styles, avail):
    merged = set()
    for rg in ws.merged_cells.ranges:
        for r in range(rg.min_row, rg.max_row + 1):
            for c in range(rg.min_col, rg.max_col + 1):
                if (r, c) != (rg.min_row, rg.min_col):
                    merged.add((r, c))
    data, cmds = [], []
    header_font = styles["hdr"]
    for r in range(1, ws.max_row + 1):
        row = []
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            txt = "" if (r, c) in merged else _fmt(cell.value, cell)
            row.append(Paragraph(txt or "&nbsp;", header_font if r <= 5 else styles["cell"]))
        data.append(row)
        fill = LIGHT if r % 2 == 0 else colors.white
        cmds.append(("BACKGROUND", (0, r - 1), (-1, r - 1), fill))
        if r <= 5:
            cmds.append(("BACKGROUND", (0, r - 1), (-1, r - 1), INK))
            cmds.append(("TEXTCOLOR", (0, r - 1), (-1, r - 1), colors.white))
        if ws.row_dimensions[r].height and ws.row_dimensions[r].height >= 40:
            cmds.append(("BOTTOMPADDING", (0, r - 1), (-1, r - 1), 9))
    cmds += [
        ("GRID", (0, 0), (-1, -1), 0.25, GREY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]
    t = Table(data, colWidths=_col_widths(ws, avail), repeatRows=1)
    t.setStyle(TableStyle(cmds))
    return t


def export(book, out, sheets, font_path="", period_cn=""):
    register_font(font_path)
    wb = openpyxl.load_workbook(book, data_only=True)
    picked = []
    for no in sheets:
        name = next((n for n in wb.sheetnames if n.startswith(str(no) + "-")), None)
        if not name:
            sys.stderr.write("[中止] 工作簿缺少第 %s 张表：%s\n" % (no, wb.sheetnames))
            raise SystemExit(1)
        picked.append((no, name))
    page_w, page_h = landscape(A4)
    margin = 12 * mm
    avail = page_w - 2 * margin

    styles = {
        "title": ParagraphStyle("t", fontName="DS_CN", fontSize=14, leading=20,
                                textColor=INK, alignment=TA_CENTER, spaceAfter=6),
        "hdr": ParagraphStyle("h", fontName="DS_CN", fontSize=6.8, leading=8.6,
                              textColor=colors.white, wordWrap="CJK"),
        "cell": ParagraphStyle("c", fontName="DS_CN", fontSize=6.8, leading=8.6,
                               textColor=colors.black, wordWrap="CJK"),
    }

    def _decorate(canvas, doc):
        canvas.saveState()
        canvas.setFont("DS_CN", 7.5)
        canvas.setFillColor(GREY)
        canvas.drawString(margin, page_h - margin + 4,
                          ("%s 铺货代销对账 · 签章件" % period_cn) if period_cn else "铺货代销对账 · 签章件")
        canvas.drawRightString(page_w - margin, 8 * mm, "第 %d 页" % doc.page)
        canvas.restoreState()

    doc = BaseDocTemplate(out, pagesize=landscape(A4),
                          leftMargin=margin, rightMargin=margin,
                          topMargin=margin, bottomMargin=margin,
                          title="铺货代销对账签章件", author="")
    frame = Frame(margin, margin, avail, page_h - 2 * margin, id="body")
    doc.addPageTemplates([PageTemplate(id="main", frames=[frame], onPage=_decorate)])

    story = []
    for idx, (no, name) in enumerate(picked):
        if idx:
            story.append(PageBreak())
        story.append(Paragraph(name, styles["title"]))
        story.append(_sheet_table(wb[name], styles, avail))
        story.append(Spacer(1, 4))
    doc.build(story)
    wb.close()
    return out, len(picked)


def main():
    ap = argparse.ArgumentParser(description="导出对账签章 PDF")
    ap.add_argument("--book", required=True, help="对账工作簿路径")
    ap.add_argument("--out", default="", help="输出 PDF；缺省为同目录 <期间>对账-签章件.pdf")
    ap.add_argument("--sheets", default="00,01,02,08", help="要导出的表号，逗号分隔")
    ap.add_argument("--font", default="", help="中文字体 TTF/TTC 完整路径")
    ap.add_argument("--period-cn", default="", help="页眉显示的期间，如 2026 年 10 月")
    a = ap.parse_args()
    if not os.path.exists(a.book):
        sys.stderr.write("[中止] 找不到工作簿：%s\n" % a.book)
        return 2
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.book)), "对账-签章件.pdf")
    sheets = [x.strip() for x in a.sheets.split(",") if x.strip()]
    try:
        path, n = export(a.book, out, sheets, a.font, a.period_cn)
    except ImportError:
        sys.stderr.write("[中止] 缺少 reportlab，请先 pip install -r requirements.txt\n")
        return 2
    print(json_line(path, n))
    return 0


def json_line(path, n):
    import json
    return json.dumps({"pdf": os.path.abspath(path), "sheets": n,
                       "bytes": os.path.getsize(path)}, ensure_ascii=False)


if __name__ == "__main__":
    raise SystemExit(main())
