#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把指定工作表的「活动区域」各截成一张图，一表一页 A4，合成签章 PDF。

每张表按单元格真实版式（列宽 / 行高 / 合并 / 底色 / 边框 / 字体 / 对齐）
**直接按最终 A4 像素尺寸绘制**（不先画大图再缩小，避免二次重采样把字糊掉），
再居中放到一整页 A4 上。

- 默认导出 00 / 01 / 02 / 08 四张打印表，输出页数 == 表数
- 默认 300 dpi（打印级）；--dpi 可调，越大越清晰、文件越大
- 页面方向 auto（按表格长宽比选横/纵）/ landscape / portrait
- 中文字体自动发现；也可用 --font 指定 TTF/TTC
- 纯本地处理；不依赖 Excel / LibreOffice
"""
import argparse
import datetime
import json
import os
import re
import sys

import openpyxl
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as pdf_canvas

FONT_CANDIDATES = [
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/System/Library/Fonts/Supplemental/Songti.ttc",
    "/System/Library/Fonts/PingFang.ttc",
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/simsun.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
]

GREY_LINE = (217, 217, 217)
INK = (31, 56, 100)
FOOT = (140, 140, 140)
A4_PT = (595.28, 841.89)
PAGE_MARGIN_IN = 0.30      # 页面左右留白（英寸）
HEAD_IN, FOOT_IN = 0.36, 0.30


def find_font(explicit=""):
    for p in ([explicit] if explicit else FONT_CANDIDATES):
        if p and os.path.exists(p):
            try:
                ImageFont.truetype(p, 12, index=0)
                return p
            except Exception:
                continue
    sys.stderr.write(
        "[中止] 找不到可用的中文字体，无法绘制表格快照。\n"
        "请把本机 TTF/TTC 中文字体的完整路径传给 --font。\n")
    raise SystemExit(1)


def _rgb(color, default=None):
    if color is None:
        return default
    raw = getattr(color, "rgb", None)
    if isinstance(raw, str) and len(raw) >= 6:
        s = raw[-6:]
        try:
            return tuple(int(s[i:i + 2], 16) for i in (0, 2, 4))
        except ValueError:
            pass
    return default


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
    return str(value)


def _used_range(ws):
    lo_r = lo_c = None
    hi_r = hi_c = 0
    for row in ws.iter_rows():
        for c in row:
            if c.value not in (None, ""):
                r, col = c.row, c.column
                lo_r = r if lo_r is None else min(lo_r, r)
                lo_c = col if lo_c is None else min(lo_c, col)
                hi_r, hi_c = max(hi_r, r), max(hi_c, col)
    for rg in ws.merged_cells.ranges:
        lo_r = rg.min_row if lo_r is None else min(lo_r, rg.min_row)
        lo_c = rg.min_col if lo_c is None else min(lo_c, rg.min_col)
        hi_r, hi_c = max(hi_r, rg.max_row), max(hi_c, rg.max_col)
    if lo_r is None:
        return 1, 1, 1, 1
    return lo_r, lo_c, hi_r, hi_c


def _logical_size(ws):
    """按 1x（Excel 默认 96dpi 视口）估算活动区域的逻辑像素尺寸。"""
    r0, c0, r1, c1 = _used_range(ws)
    lw = 0.0
    for c in range(c0, c1 + 1):
        dim = ws.column_dimensions.get(get_column_letter(c))
        lw += (dim.width if dim and dim.width else 8.43) * 7.0
    lh = 0.0
    for r in range(r0, r1 + 1):
        dim = ws.row_dimensions.get(r)
        lh += (dim.height if dim and dim.height else 15.0) * (96.0 / 72.0)
    return max(1.0, lw), max(1.0, lh)


def _wrap(draw, text, font, max_w):
    lines = []
    for para in str(text).split("\n"):
        if not para:
            lines.append("")
            continue
        cur = ""
        for ch in para:
            if draw.textlength(cur + ch, font=font) > max_w and cur:
                lines.append(cur)
                cur = ch
            else:
                cur += ch
        lines.append(cur)
    return lines


def _warn_section_merges(ws):
    """章节标题行若未合并，快照右侧会出现一排竖线；提示而不是静默。"""
    bad = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if not isinstance(v, str):
            continue
        t = v.strip()
        if not (re.match(r"^[一二三四五六七八九十]+、", t) or t.startswith("附 表")):
            continue
        if not [m for m in ws.merged_cells.ranges if m.min_row == r and m.min_col == 1]:
            bad.append("A%d" % r)
    if bad:
        sys.stderr.write("[提示] %s 的章节标题行未合并（%s），快照会出现竖线；"
                         "增删行后请同步移动合并区。\n" % (ws.title, "、".join(bad)))


def render_sheet(ws, font_path, target_w, target_h):
    """把活动区域**按目标像素尺寸直接绘制**（字号/线宽同步缩放，不再二次重采样）。"""
    _warn_section_merges(ws)
    r0, c0, r1, c1 = _used_range(ws)
    lw, lh = _logical_size(ws)
    s = min(float(target_w) / lw, float(target_h) / lh)

    col_w, row_h = {}, {}
    for c in range(c0, c1 + 1):
        dim = ws.column_dimensions.get(get_column_letter(c))
        col_w[c] = max(1.0, (dim.width if dim and dim.width else 8.43) * 7.0 * s)
    for r in range(r0, r1 + 1):
        dim = ws.row_dimensions.get(r)
        row_h[r] = max(1.0, (dim.height if dim and dim.height else 15.0) * (96.0 / 72.0) * s)
    xs, ys = {c0: 0.0}, {r0: 0.0}
    for c in range(c0, c1 + 1):
        xs[c + 1] = xs[c] + col_w[c]
    for r in range(r0, r1 + 1):
        ys[r + 1] = ys[r] + row_h[r]
    W, H = max(1, int(round(xs[c1 + 1]))), max(1, int(round(ys[r1 + 1])))
    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    merged_at, hidden = {}, set()
    for rg in ws.merged_cells.ranges:
        merged_at[(rg.min_row, rg.min_col)] = rg
        for r in range(rg.min_row, rg.max_row + 1):
            for c in range(rg.min_col, rg.max_col + 1):
                if (r, c) != (rg.min_row, rg.min_col):
                    hidden.add((r, c))

    font_cache = {}

    def _font(size_px, bold):
        key = (int(size_px), bool(bold))
        if key not in font_cache:
            try:
                font_cache[key] = ImageFont.truetype(font_path, max(1, int(size_px)), index=0)
            except Exception:
                font_cache[key] = ImageFont.load_default()
        return font_cache[key]

    line_w = max(1, int(round(s)))
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            if (r, c) in hidden:
                continue
            rg = merged_at.get((r, c))
            er = rg.max_row if rg else r
            ec = rg.max_col if rg else c
            x0, y0 = xs[c], ys[r]
            x1, y1 = xs[ec + 1], ys[er + 1]
            cell = ws.cell(row=r, column=c)
            bg = _rgb(cell.fill.fgColor if cell.fill else None)
            if bg and bg != (255, 255, 255) and getattr(cell.fill, "patternType", None):
                draw.rectangle([x0, y0, x1 - 1, y1 - 1], fill=bg)
            bc = _rgb(getattr(getattr(cell.border, "bottom", None), "color", None), GREY_LINE)
            draw.rectangle([x0, y0, x1 - 1, y1 - 1], outline=bc, width=line_w)

            txt = _fmt(cell.value, cell)
            if not txt:
                continue
            f = cell.font
            size = max(5.0, float(f.sz or 10)) * (96.0 / 72.0) * s
            font = _font(size, f.b)
            fg = _rgb(f.color, (0, 0, 0))
            if fg == (0, 0, 0) and f.color and getattr(f.color, "theme", None) is not None:
                fg = INK
            pad = max(1.0, 3.0 * s)
            lines = _wrap(draw, txt, font, max(1.0, x1 - x0 - 2 * pad))
            lh_px = size * 1.22
            total = lh_px * len(lines)
            ty = y0 + max(pad, (y1 - y0 - total) / 2)
            if not (cell.alignment and cell.alignment.wrap_text) and len(lines) > 1:
                if (y1 - y0) < total * 1.4:
                    lines, total = lines[:1], lh_px
                    ty = y0 + max(pad, (y1 - y0 - total) / 2)
            for ln in lines:
                tw = draw.textlength(ln, font=font)
                hz = (cell.alignment.horizontal if cell.alignment else None) or "center"
                if hz == "left":
                    tx = x0 + pad
                elif hz == "right":
                    tx = x1 - tw - pad
                else:
                    tx = x0 + max(pad, (x1 - x0 - tw) / 2)
                kw = {"stroke_width": max(1, line_w // 2), "stroke_fill": fg} if f.b else {}
                draw.text((tx, ty), ln, font=font, fill=fg, **kw)
                ty += lh_px
    return img


def _compose(sheet_img, label, period_cn, page_no, total, land, dpi, font_path):
    wpt, hpt = (A4_PT[1], A4_PT[0]) if land else A4_PT
    pw, ph = int(wpt / 72 * dpi), int(hpt / 72 * dpi)
    page = Image.new("RGB", (pw, ph), "white")
    d = ImageDraw.Draw(page)
    head_h, foot_h = int(HEAD_IN * dpi), int(FOOT_IN * dpi)
    mx = int(PAGE_MARGIN_IN * dpi)
    try:
        hf = ImageFont.truetype(font_path, max(9, int(0.105 * dpi)), index=0)
        sf = ImageFont.truetype(font_path, max(8, int(0.09 * dpi)), index=0)
    except Exception:
        hf = sf = ImageFont.load_default()
    d.text((mx, int(0.10 * dpi)),
           (period_cn + " 铺货代销对账 · 签章件  |  " + label)
           if period_cn else ("铺货代销对账 · 签章件  |  " + label),
           font=hf, fill=INK)
    avail_w, avail_h = pw - 2 * mx, ph - head_h - foot_h
    nw, nh = min(sheet_img.width, avail_w), min(sheet_img.height, avail_h)
    page.paste(sheet_img.crop((0, 0, nw, nh)),
               (mx + (avail_w - nw) // 2, head_h + (avail_h - nh) // 2))
    d.text((pw - int(1.30 * dpi), ph - int(0.24 * dpi)),
           "第 %d 页 / 共 %d 页" % (page_no, total), font=sf, fill=FOOT)
    return page, (wpt, hpt)


def export(book, out, sheets, font_path="", period_cn="",
           orientation="auto", dpi=300, scale=None):
    font_path = find_font(font_path)
    wb = openpyxl.load_workbook(book, data_only=True)
    picked = []
    for no in sheets:
        name = next((n for n in wb.sheetnames if n.startswith(str(no) + "-")), None)
        if not name:
            sys.stderr.write("[中止] 工作簿缺少第 %s 张表：%s\n" % (no, wb.sheetnames))
            raise SystemExit(1)
        picked.append((no, name))

    pages, sizes, dirs = [], [], []
    for i, (_, name) in enumerate(picked, 1):
        ws = wb[name]
        lw, lh = _logical_size(ws)
        land = (orientation == "landscape" or
                (orientation == "auto" and lw / lh > 1.15))
        wpt, hpt = (A4_PT[1], A4_PT[0]) if land else A4_PT
        pw, ph = int(wpt / 72 * dpi), int(hpt / 72 * dpi)
        avail_w = pw - 2 * int(PAGE_MARGIN_IN * dpi)
        avail_h = ph - int(HEAD_IN * dpi) - int(FOOT_IN * dpi)
        img = render_sheet(ws, font_path, avail_w, avail_h)
        page, size = _compose(img, name, period_cn, i, len(picked),
                              land, dpi, font_path)
        pages.append(page)
        sizes.append(size)
        dirs.append("横向" if land else "纵向")
    wb.close()

    c = pdf_canvas.Canvas(out, pagesize=sizes[0])
    c.setTitle("铺货代销对账签章件")
    c.setAuthor("")
    for i, (page, size) in enumerate(zip(pages, sizes)):
        if i and size != sizes[0]:
            c.setPageSize(size)
        c.drawImage(ImageReader(page), 0, 0, width=size[0], height=size[1])
        c.showPage()
    c.save()
    return out, len(pages), dirs


def main():
    ap = argparse.ArgumentParser(description="活动区域快照 → 一表一页 A4 签章 PDF")
    ap.add_argument("--book", required=True)
    ap.add_argument("--out", default="")
    ap.add_argument("--sheets", default="00,01,02,08")
    ap.add_argument("--font", default="")
    ap.add_argument("--period-cn", default="")
    ap.add_argument("--orientation", default="auto",
                    choices=("auto", "landscape", "portrait"))
    ap.add_argument("--dpi", type=int, default=300,
                    help="输出分辨率，默认 300（打印级）；越大越清晰、文件越大")
    a = ap.parse_args()
    if not os.path.exists(a.book):
        sys.stderr.write("[中止] 找不到工作簿：%s\n" % a.book)
        return 2
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.book)), "对账-签章件.pdf")
    sheets = [x.strip() for x in a.sheets.split(",") if x.strip()]
    path, n, dirs = export(a.book, out, sheets, a.font, a.period_cn,
                           a.orientation, a.dpi)
    print(json.dumps({"pdf": os.path.abspath(path), "pages": n,
                      "sheets": sheets, "orientation": dirs, "dpi": a.dpi,
                      "bytes": os.path.getsize(path)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
