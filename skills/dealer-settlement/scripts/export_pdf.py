#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把指定工作表的「活动区域」各截成一张图，一表一页 A4，合成签章 PDF。

与 v1.7 的流式排版不同：这里不重新排表格，也不让一张表跨页。
每张表按单元格真实版式（列宽 / 行高 / 合并 / 底色 / 边框 / 字体 / 对齐）
绘制成一张位图，再等比缩放到一整页 A4 的可用区域，居中放置。

- 默认导出 00 / 01 / 02 / 08 四张打印表，输出页数 == 表数
- 页面方向 auto（按图片宽高比选横/纵）/ landscape / portrait
- 中文字体自动发现；也可用 --font 指定 TTF/TTC
- 纯本地处理，不上传任何数据；不依赖 Excel / LibreOffice
"""
import argparse
import datetime
import json
import os
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
    """openpyxl 颜色 -> (r,g,b)；主题色/索引色取默认值。"""
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
    import re as _re
    bad = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if not isinstance(v, str):
            continue
        t = v.strip()
        if not (_re.match(r"^[一二三四五六七八九十]+、", t) or t.startswith("附 表")):
            continue
        if not [m for m in ws.merged_cells.ranges if m.min_row == r and m.min_col == 1]:
            bad.append("A%d" % r)
    if bad:
        sys.stderr.write("[提示] %s 的章节标题行未合并（%s），快照会出现竖线；"
                         "增删行后请同步移动合并区。\n" % (ws.title, "、".join(bad)))


def render_sheet(ws, font_path, scale=2.0):
    """把工作表活动区域绘成 PIL 图片（保留合并/底色/边框/字体/对齐）。"""
    _warn_section_merges(ws)
    r0, c0, r1, c1 = _used_range(ws)
    col_w, row_h = {}, {}
    for c in range(c0, c1 + 1):
        dim = ws.column_dimensions.get(get_column_letter(c))
        w = (dim.width if dim and dim.width else 8.43) * 7.0 * scale
        col_w[c] = max(28.0 * scale, w)
    for r in range(r0, r1 + 1):
        dim = ws.row_dimensions.get(r)
        h = (dim.height if dim and dim.height else 15.0) * (96.0 / 72.0) * scale
        row_h[r] = max(14.0 * scale, h)
    xs, ys = {c0: 0.0}, {r0: 0.0}
    for c in range(c0, c1 + 1):
        xs[c + 1] = xs[c] + col_w[c]
    for r in range(r0, r1 + 1):
        ys[r + 1] = ys[r] + row_h[r]
    W, H = int(xs[c1 + 1]) + 2, int(ys[r1 + 1]) + 2
    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    merged_at, hidden = {}, set()
    for rg in ws.merged_cells.ranges:
        merged_at[(rg.min_row, rg.min_col)] = rg
        for r in range(rg.min_row, rg.max_row + 1):
            for c in range(rg.min_col, rg.max_col + 1):
                if (r, c) != (rg.min_row, rg.min_col):
                    hidden.add((r, c))

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
            draw.rectangle([x0, y0, x1 - 1, y1 - 1], outline=bc, width=1)

            txt = _fmt(cell.value, cell)
            if not txt:
                continue
            f = cell.font
            size = max(7.0, min(26.0, float(f.sz or 10))) * (96.0 / 72.0) * scale
            try:
                font = ImageFont.truetype(font_path, int(size), index=0)
            except Exception:
                font = ImageFont.load_default()
            fg = _rgb(f.color, (0, 0, 0))
            if fg == (0, 0, 0) and f.color and getattr(f.color, "theme", None) is not None:
                fg = INK
            pad = 3 * scale
            lines = _wrap(draw, txt, font, max(8.0, x1 - x0 - 2 * pad))
            lh = size * 1.22
            total = lh * len(lines)
            ty = y0 + max(pad, (y1 - y0 - total) / 2)
            if not (cell.alignment and cell.alignment.wrap_text) and len(lines) > 1:
                if (y1 - y0) < total * 1.4:
                    lines, total = lines[:1], lh
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
                kw = {"stroke_width": 1, "stroke_fill": fg} if f.b else {}
                draw.text((tx, ty), ln, font=font, fill=fg, **kw)
                ty += lh
    return img


def _page_image(sheet_img, label, period_cn, page_no, total, orientation, dpi, font_path):
    auto_landscape = sheet_img.width / sheet_img.height > 1.15
    land = (orientation == "landscape" or
            (orientation == "auto" and auto_landscape))
    wpt, hpt = (A4_PT[1], A4_PT[0]) if land else A4_PT
    pw, ph = int(wpt / 72 * dpi), int(hpt / 72 * dpi)
    page = Image.new("RGB", (pw, ph), "white")
    d = ImageDraw.Draw(page)
    head_h, foot_h = int(0.34 * dpi), int(0.28 * dpi)
    try:
        hf = ImageFont.truetype(font_path, int(0.11 * dpi), index=0)
        sf = ImageFont.truetype(font_path, int(0.095 * dpi), index=0)
    except Exception:
        hf = sf = ImageFont.load_default()
    d.text((int(0.32 * dpi), int(0.09 * dpi)),
           (period_cn + " 铺货代销对账 · 签章件  |  " + label)
           if period_cn else ("铺货代销对账 · 签章件  |  " + label),
           font=hf, fill=INK)
    avail_w, avail_h = pw - int(0.64 * dpi), ph - head_h - foot_h
    ratio = min(avail_w / sheet_img.width, avail_h / sheet_img.height)
    nw, nh = max(1, int(sheet_img.width * ratio)), max(1, int(sheet_img.height * ratio))
    thumb = sheet_img.resize((nw, nh), Image.LANCZOS)
    page.paste(thumb, ((pw - nw) // 2, head_h + (avail_h - nh) // 2))
    d.text((pw - int(1.35 * dpi), ph - int(0.22 * dpi)),
           "第 %d 页 / 共 %d 页" % (page_no, total), font=sf, fill=FOOT)
    return page, (wpt, hpt)


def export(book, out, sheets, font_path="", period_cn="",
           orientation="auto", dpi=150, scale=2.0):
    font_path = find_font(font_path)
    wb = openpyxl.load_workbook(book, data_only=True)
    picked = []
    for no in sheets:
        name = next((n for n in wb.sheetnames if n.startswith(str(no) + "-")), None)
        if not name:
            sys.stderr.write("[中止] 工作簿缺少第 %s 张表：%s\n" % (no, wb.sheetnames))
            raise SystemExit(1)
        picked.append((no, name))
    pages, sizes = [], []
    for i, (_, name) in enumerate(picked, 1):
        img = render_sheet(wb[name], font_path, scale)
        page, size = _page_image(img, name, period_cn, i, len(picked),
                                 orientation, dpi, font_path)
        pages.append(page)
        sizes.append(size)
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
    return out, len(pages), [("横向" if s[0] > s[1] else "纵向") for s in sizes]


def main():
    ap = argparse.ArgumentParser(description="活动区域快照 → 一表一页 A4 签章 PDF")
    ap.add_argument("--book", required=True)
    ap.add_argument("--out", default="")
    ap.add_argument("--sheets", default="00,01,02,08")
    ap.add_argument("--font", default="")
    ap.add_argument("--period-cn", default="")
    ap.add_argument("--orientation", default="auto",
                    choices=("auto", "landscape", "portrait"))
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--scale", type=float, default=2.0)
    a = ap.parse_args()
    if not os.path.exists(a.book):
        sys.stderr.write("[中止] 找不到工作簿：%s\n" % a.book)
        return 2
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.book)), "对账-签章件.pdf")
    sheets = [x.strip() for x in a.sheets.split(",") if x.strip()]
    path, n, dirs = export(a.book, out, sheets, a.font, a.period_cn,
                           a.orientation, a.dpi, a.scale)
    print(json.dumps({"pdf": os.path.abspath(path), "pages": n,
                      "sheets": sheets, "orientation": dirs,
                      "bytes": os.path.getsize(path)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
