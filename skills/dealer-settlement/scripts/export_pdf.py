#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把工作表活动区域按真实版式输出为 A4 签章 PDF。

版式规则（v1.9）
- **宽度填满**：表格宽度缩放到 A4 可用宽度，不再为了「一表一页」压缩字；
- **高度切页**：内容超出可用高度就分页，页脚标注总页数；
- 方向：全簿统一，保证所有页宽度一致；auto 时按「最小缩放比更大」的方向整体应用；
- 字号不放大（缩放比上限 1.0），避免小表被撑得过大。

渲染后端
- ``vector``（默认）：reportlab 直接绘制文字、边框、底色，文字是矢量（可选中、
  可搜索、放大不虚），字体按子集嵌入。
- ``raster``：PIL 按像素重绘再贴图（兜底，仍为单表一页）。

注意：字号大小由「表格宽度 ÷ A4 宽度」决定。若某表列特别多（如 16 列），
即使切页也受宽度限制；要再大只能横向、精简列或改纸张。
"""
import argparse
import datetime
import math
import json
import os
import re
import sys

import openpyxl
from openpyxl.utils import get_column_letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
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
BREAK_AFTER = "·；，。、）】》：!?！？ "
INK = (31, 56, 100)
FOOT = (140, 140, 140)
A4_PT = (595.28, 841.89)
PAGE_MARGIN_IN = 0.30
HEAD_IN, FOOT_IN = 0.36, 0.30


def find_font(explicit=""):
    for p in ([explicit] if explicit else FONT_CANDIDATES):
        if p and os.path.exists(p):
            try:
                TTFont("_probe", p, subfontIndex=0)
                return p
            except Exception:
                try:
                    TTFont("_probe2", p)
                    return p
                except Exception:
                    continue
    sys.stderr.write(
        "[中止] 找不到可用的中文字体，无法绘制表格。\n"
        "请把本机 TTF/TTC 中文字体的完整路径传给 --font。\n")
    raise SystemExit(1)


def register_font(path):
    name = "CNBody"
    try:
        pdfmetrics.registerFont(TTFont(name, path, subfontIndex=0))
    except Exception:
        pdfmetrics.registerFont(TTFont(name, path))
    return name


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


def _grid(ws):
    """活动区域几何。隐藏行/列宽度置 0（不印进签章件），并回传集合供绘制跳过。"""
    r0, c0, r1, c1 = _used_range(ws)
    hidden_cols, hidden_rows = set(), set()
    xs, ys = {c0: 0.0}, {r0: 0.0}
    for c in range(c0, c1 + 1):
        dim = ws.column_dimensions.get(get_column_letter(c))
        if dim is not None and getattr(dim, "hidden", False):
            hidden_cols.add(c)
            w = 0.0
        else:
            w_chars = dim.width if dim and dim.width else 8.43
            w = max(1.0, w_chars * 7.0 * 0.75)
        xs[c + 1] = xs[c] + w
    for r in range(r0, r1 + 1):
        dim = ws.row_dimensions.get(r)
        if dim is not None and getattr(dim, "hidden", False):
            hidden_rows.add(r)
            h = 0.0
        else:
            h = dim.height if dim and dim.height else 15.0
            h = max(1.0, h)
        ys[r + 1] = ys[r] + h
    return (r0, c0, r1, c1), xs, ys, xs[c1 + 1], ys[r1 + 1], hidden_cols, hidden_rows


def _main_font_size(ws):
    from collections import Counter
    r0, c0, r1, c1 = _used_range(ws)
    sizes = Counter()
    for row in ws.iter_rows(min_row=r0, max_row=r1, min_col=c0, max_col=c1):
        for cell in row:
            if cell.value not in (None, "") and cell.font and cell.font.sz:
                sizes[float(cell.font.sz)] += 1
    return sizes.most_common(1)[0][0] if sizes else 9.0


def _wrap(text, fontname, size, max_w):
    def w(t):
        return pdfmetrics.stringWidth(t, fontname, size)

    lines = []
    for para in str(text).split("\n"):
        if not para:
            lines.append("")
            continue
        cur = ""
        for ch in para:
            if w(cur + ch) <= max_w or not cur:
                cur += ch
                continue
            cut = -1
            for i in range(len(cur) - 1, max(0, int(len(cur) * 0.55)) - 1, -1):
                if cur[i] in BREAK_AFTER:
                    cut = i + 1
                    break
            if cut > 0:
                lines.append(cur[:cut])
                cur = cur[cut:] + ch
            else:
                lines.append(cur)
                cur = ch
        lines.append(cur)
    return lines


def _warn_section_merges(ws):
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


HEADER_KEYS = ("序号", "商品名称", "项目", "类型", "名称", "数量", "金额",
               "日期", "规格", "单位", "门店")
SKIP_KEYS = ("合计", "小计", "总计", "累计", "本页", "备注：")


def _header_rows(ws, r0, c0, r1, c1):
    """识别表头行：优先 print_title_rows，否则「过半加粗 + 含列名关键词」。

    合计/小计行即使加粗也不算表头，否则续页会把小计行当表头重复。
    """
    ptr = str(getattr(ws, "print_title_rows", "") or "").strip()
    m = re.match(r"^\$?(\d+):\$?(\d+)$", ptr)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        rows = [r for r in range(max(a, r0), min(b, r1) + 1)]
        # print_title_rows 也可能被设成文档标题行（如 2:4），必须含表头关键词才采信
        if rows:
            blob = "".join(re.sub(r"\s+", "", str(ws.cell(rr, c).value or ""))
                           for rr in rows for c in range(c0, c1 + 1))
            if any(k in blob for k in HEADER_KEYS):
                return rows
    out = []
    for r in range(r0, r1 + 1):
        a_val = ws.cell(row=r, column=c0).value
        # 明细行的 A 列是序号数字，不可能是表头
        if isinstance(a_val, (int, float)):
            continue
        if isinstance(a_val, str) and a_val.strip().isdigit():
            continue
        cells = [ws.cell(row=r, column=col) for col in range(c0, c1 + 1)]
        vals = [c.value for c in cells if c.value not in (None, "")]
        if len(vals) < 2:
            continue
        text = " ".join(str(v) for v in vals)
        flat = re.sub(r"\s+", "", text)      # 「合　计」中间是全角空格，先归一化
        if any(k in flat for k in SKIP_KEYS):
            continue
        # 只按「有值单元格」算加粗比例：合并表头在 12 列里可能只有 4 格有值
        nval = len(vals)
        nb = sum(1 for c in cells
                 if c.value not in (None, "") and c.font is not None and c.font.b)
        if nb >= max(2, int(nval * 0.6)) and any(k in flat for k in HEADER_KEYS):
            out.append(r)
    return out


def _is_section_row(ws, r):
    v = ws.cell(r, 1).value
    if not isinstance(v, str):
        return False
    t = v.strip()
    return bool(re.match(r"^[一二三四五六七八九十]+、", t)) or t.startswith("附 表")


def _plan(ws, orientation="auto"):
    """决定方向、缩放比与分页：宽度填满，高度切页，字号不放大。"""
    (r0, c0, r1, c1), xs, ys, W, H, hidden_cols, hidden_rows = _grid(ws)

    if r0 == r1 == c0 == c1 == 1 and ws.cell(1, 1).value in (None, ""):
        return {"land": False, "wpt": A4_PT[0], "hpt": A4_PT[1],
                "aw": A4_PT[0] - 2 * PAGE_MARGIN_IN * 72,
                "ah": A4_PT[1] - HEAD_IN * 72 - FOOT_IN * 72,
                "s": 1.0, "pages": [], "hdr": [], "xs": xs, "ys": ys, "W": 0.0, "H": 0.0,
                "r0": r0, "c0": c0, "r1": r1, "c1": c1, "main_pt": 0.0,
                "hidden_cols": hidden_cols, "hidden_rows": hidden_rows, "empty": True}

    def cand(land):
        wpt, hpt = (A4_PT[1], A4_PT[0]) if land else A4_PT
        aw = wpt - 2 * PAGE_MARGIN_IN * 72
        ah = hpt - HEAD_IN * 72 - FOOT_IN * 72
        return {"land": land, "wpt": wpt, "hpt": hpt, "aw": aw, "ah": ah,
                "s": min(1.0, aw / W)}

    p, l = cand(False), cand(True)
    if orientation == "portrait":
        c = p
    elif orientation == "landscape":
        c = l
    else:
        c = p if (p["s"] >= 0.85 or l["s"] <= p["s"] * 1.10) else l

    # 平衡分页：先算需要几页，再按平均高度切；续页重复表头并从可用高度扣除
    headers = _header_rows(ws, r0, c0, r1, c1)
    header_set = set(headers)

    def need_header(first):
        # 本页首行是表头或新章节标题时都不重复上一节表头
        if first in header_set or _is_section_row(ws, first):
            return None
        prev = [h for h in headers if h < first]
        if not prev:
            return None
        h = prev[-1]
        # 表头与首行之间夹着章节标题（哪怕标题在上一页页末）说明已换节，不重复
        if any(_is_section_row(ws, r) for r in range(h + 1, first)):
            return None
        return h

    def hh(h):
        return (ys[h + 1] - ys[h]) * c["s"] if h is not None else 0.0

    total_h = sum((ys[r + 1] - ys[r]) * c["s"] for r in range(r0, r1 + 1))
    n = max(1, int(math.ceil(total_h / c["ah"] - 1e-9)))
    target = total_h / n
    pages, metas, cur, cur_h, cur_hdr = [], [], [], 0.0, None
    for r in range(r0, r1 + 1):
        rh = (ys[r + 1] - ys[r]) * c["s"]
        if rh > c["ah"] + 1e-6:
            sys.stderr.write("[提示] %s 第 %d 行高 %.1fpt 超过单页可用高度 %.1fpt，"
                             "将在本页裁断；请把该行拆成多行\n"
                             % (ws.title, r, rh, c["ah"]))
        if not cur:
            cur_hdr = need_header(r)
        limit = c["ah"] - hh(cur_hdr)
        if cur and (cur_h + rh > limit + 1e-6 or
                    (len(pages) < n - 1 and cur_h >= target - 1e-9)):
            pages.append(cur)
            metas.append(cur_hdr)
            cur, cur_h = [], 0.0
            cur_hdr = need_header(r)
            limit = c["ah"] - hh(cur_hdr)
        cur.append(r)
        cur_h += rh
    if cur:
        pages.append(cur)
        metas.append(cur_hdr)

    # 跨页合并区：只按左上角所在页绘制，提示核对
    row_page = {}
    for idx, rows in enumerate(pages):
        for r in rows:
            row_page[r] = idx
    for rg in ws.merged_cells.ranges:
        a, b = row_page.get(rg.min_row), row_page.get(rg.max_row)
        if a is not None and b is not None and a != b:
            sys.stderr.write("[提示] %s 的合并区 %s 跨页，已按首页裁断，请核对版式\n"
                             % (ws.title, rg.coord))

    # 孤行保护：章节标题若落在页末，且下一页放得下，就整行下移
    def _page_h(rows_):
        return sum((ys[r + 1] - ys[r]) * c["s"] for r in rows_)

    for i in range(len(pages) - 1):
        if len(pages[i]) <= 1:
            continue
        last = pages[i][-1]
        if not _is_section_row(ws, last):
            continue
        nh = hh(metas[i + 1])
        if _page_h(pages[i + 1]) + (ys[last + 1] - ys[last]) * c["s"] <= c["ah"] - nh:
            pages[i + 1].insert(0, pages[i].pop())

    c.update({"pages": pages, "hdr": metas, "xs": xs, "ys": ys, "W": W, "H": H,
              "hidden_cols": hidden_cols, "hidden_rows": hidden_rows, "empty": False,
              "r0": r0, "c0": c0, "r1": r1, "c1": c1,
              "main_pt": _main_font_size(ws) * c["s"]})
    return c


def _decide_global(wb, picked, orientation):
    """全簿统一页面方向，保证所有页宽度一致。

    auto 时分别按纵向 / 横向计算每张表的缩放比，取「最小缩放比更大」的方向：
    最小缩放比决定全簿最丑的那张表，优先把它抬高。
    """
    if orientation in ("portrait", "landscape") or not picked:
        return orientation if orientation in ("portrait", "landscape") else "portrait"
    score = {"portrait": [], "landscape": []}
    for _, name in picked:
        ws = wb[name]
        _, _, _, W, _, _, _ = _grid(ws)
        for key, land in (("portrait", False), ("landscape", True)):
            wpt = A4_PT[1] if land else A4_PT[0]
            aw = wpt - 2 * PAGE_MARGIN_IN * 72
            score[key].append(min(1.0, aw / W))
    return max(score, key=lambda k: min(score[k]))


def draw_rows(c, ws, fontname, plan, rows, page_no, total, label, period_cn,
              header_row=None, is_cont=False):
    """把 rows 画到当前页；header_row 非空时在页顶重复表头；is_cont 标注续页。"""
    _warn_section_merges(ws)
    s = plan["s"]
    wpt, hpt = plan["wpt"], plan["hpt"]
    c.setPageSize((wpt, hpt))
    mx = PAGE_MARGIN_IN * 72
    head_h = HEAD_IN * 72
    aw = plan["aw"]

    c._code.append("0 Tr")          # 清掉上一页可能残留的描边文字模式
    hs = 7.5
    c.setFont(fontname, hs)
    c.setFillColorRGB(*(v / 255.0 for v in INK))
    head = ((period_cn + " 铺货代销对账 · 签章件  |  " + label) if period_cn
            else ("铺货代销对账 · 签章件  |  " + label))
    if is_cont:
        head += "（续）"
    c.drawString(mx, hpt - 0.10 * 72 - pdfmetrics.getAscent(fontname, hs), head)

    xs, ys = plan["xs"], plan["ys"]
    hid_c = plan.get("hidden_cols") or set()
    hid_r = plan.get("hidden_rows") or set()
    paint_rows = ([header_row] if header_row is not None else []) + list(rows)
    row_top, y = {}, 0.0
    for r in paint_rows:
        row_top[r] = y
        y += (ys[r + 1] - ys[r]) * s
    ox = mx + max(0.0, (aw - plan["W"] * s) / 2)
    base_y = hpt - head_h

    row_set = set(paint_rows)
    last_row = rows[-1]
    first_row = rows[0]
    filled_rows = set()
    for r in paint_rows:
        if r in hid_r:
            continue
        for col in range(plan["c0"], plan["c1"] + 1):
            if col in hid_c:
                continue
            if ws.cell(row=r, column=col).value not in (None, ""):
                filled_rows.add(r)
                break
    for rg in ws.merged_cells.ranges:
        if rg.min_row in row_set:
            for r in range(rg.min_row, rg.max_row + 1):
                filled_rows.add(r)

    merged_at, hidden, cover_at = {}, set(), {}
    for rg in ws.merged_cells.ranges:
        merged_at[(rg.min_row, rg.min_col)] = rg
        for r in range(rg.min_row, rg.max_row + 1):
            for col in range(rg.min_col, rg.max_col + 1):
                cover_at[(r, col)] = rg
                if (r, col) != (rg.min_row, rg.min_col):
                    hidden.add((r, col))

    line_w = max(0.3, 0.5 * s)
    for r in paint_rows:
        if r not in filled_rows or r in hid_r:
            continue
        for col in range(plan["c0"], plan["c1"] + 1):
            if col in hid_c:
                continue
            cov = cover_at.get((r, col))
            if cov is not None and cov.min_row < first_row:
                # 续页：合并区从上一页延续，只补边框、不重复文字
                x0 = ox + xs[col] * s
                x1 = ox + xs[col + 1] * s
                y0 = row_top.get(r, 0.0)
                y1 = min(y0 + (ys[r + 1] - ys[r]) * s, base_y)
                if y1 > y0:
                    c.setStrokeColorRGB(*(v / 255.0 for v in GREY_LINE))
                    c.setLineWidth(line_w)
                    c.rect(x0, base_y - y1, x1 - x0, y1 - y0, fill=0, stroke=1)
                continue
            if (r, col) in hidden:
                continue
            rg = merged_at.get((r, col))
            er = min(rg.max_row, last_row) if rg else r
            ec = rg.max_col if rg else col
            x0 = ox + xs[col] * s
            x1 = ox + xs[ec + 1] * s
            y0 = row_top[r]
            y1 = row_top.get(er + 1, y0 + (ys[er + 1] - ys[r]) * s)
            y1 = min(y1, base_y)          # 超高行不越出页脚
            if y1 <= y0:
                continue
            ry = base_y - y1
            w, h = x1 - x0, y1 - y0
            cell = ws.cell(row=r, column=col)

            bg = _rgb(cell.fill.fgColor if cell.fill else None)
            if bg and bg != (255, 255, 255) and getattr(cell.fill, "patternType", None):
                c.setFillColorRGB(*(v / 255.0 for v in bg))
                c.rect(x0, ry, w, h, fill=1, stroke=0)

            bc = _rgb(getattr(getattr(cell.border, "bottom", None), "color", None), GREY_LINE)
            c.setStrokeColorRGB(*(v / 255.0 for v in bc))
            c.setLineWidth(line_w)
            c.rect(x0, ry, w, h, fill=0, stroke=1)

            txt = _fmt(cell.value, cell)
            if not txt:
                continue
            f = cell.font
            size = max(4.0, float(f.sz or 10) * s)
            pad = max(0.8, 1.5 * s)
            avail_w = max(1.0, w - 2 * pad)
            lh = size * 1.22
            max_lines = max(1, int((h - 2 * pad) / lh))

            lines = _wrap(txt, fontname, size, avail_w)
            if "\n" not in txt:
                one_w = pdfmetrics.stringWidth(txt, fontname, size)
                if one_w > avail_w:
                    fit = (avail_w * 0.97) / one_w
                    if fit >= 0.82:
                        size *= fit
                        lh = size * 1.22
                        lines = [txt]
            if len(lines) > max_lines:
                for k in (0.95, 0.9, 0.85, 0.8, 0.75, 0.7):
                    nsize = size * k
                    nlh = nsize * 1.22
                    nmax = max(1, int((h - 2 * pad) / nlh))
                    l2 = _wrap(txt, fontname, nsize, avail_w)
                    if len(l2) <= nmax:
                        size, lines, lh = nsize, l2, nlh
                        break
                else:
                    lines = lines[:max_lines]
                    sys.stderr.write("[提示] %s!%s 文本超出行高，已截断\n"
                                     % (ws.title, cell.coordinate))
            if not (cell.alignment and cell.alignment.wrap_text) and len(lines) > 1:
                if h < lh * len(lines) * 1.4:
                    lines = lines[:1]

            total_h = lh * len(lines)
            ty = y0 + max(pad, (h - total_h) / 2)
            fg = _rgb(f.color, (0, 0, 0))
            if fg == (0, 0, 0) and f.color and getattr(f.color, "theme", None) is not None:
                fg = INK
            c.setFillColorRGB(*(v / 255.0 for v in fg))
            ascent = pdfmetrics.getAscent(fontname, size)
            hz = (cell.alignment.horizontal if cell.alignment else None) or "center"
            for ln in lines:
                baseline = base_y - ty - ascent
                tw = pdfmetrics.stringWidth(ln, fontname, size)
                if hz == "left":
                    tx = x0 + pad
                elif hz == "right":
                    tx = x1 - tw - pad
                else:
                    tx = x0 + max(pad, (w - tw) / 2)
                if f.b:
                    # 用 PDF 文本渲染模式加粗（fill+stroke），只写一份文字；
                    # 必须立即复位，否则同页后续正文都会被描边
                    c.setLineWidth(max(0.12, size * 0.022))
                    t = c.beginText(tx, baseline)
                    t.setFont(fontname, size)
                    t.setTextRenderMode(2)
                    t.textOut(ln)
                    c.drawText(t)
                    c._code.append("0 Tr")
                    c.setLineWidth(line_w)
                else:
                    c.setFont(fontname, size)
                    c.drawString(tx, baseline, ln)
                ty += lh

    fs = 6.5
    c.setFont(fontname, fs)
    c.setFillColorRGB(*(v / 255.0 for v in FOOT))
    c.drawString(wpt - 1.30 * 72, 0.24 * 72,
                 "第 %d 页 / 共 %d 页" % (page_no, total))


# ----------------------------------------------------------------- 位图兜底
def _render_sheet_raster(ws, font_path, target_w, target_h):
    from PIL import Image, ImageDraw, ImageFont

    _warn_section_merges(ws)
    (r0, c0, r1, c1), xs0, ys0, W0, H0 = _grid(ws)
    lw = W0 / 0.75
    lh = H0 * (96.0 / 72.0)
    s = min(float(target_w) / lw, float(target_h) / lh)
    col_w = {c: (xs0[c + 1] - xs0[c]) / 0.75 * s for c in range(c0, c1 + 1)}
    row_h = {r: (ys0[r + 1] - ys0[r]) * (96.0 / 72.0) * s for r in range(r0, r1 + 1)}
    xs, ys = {c0: 0.0}, {r0: 0.0}
    for c in range(c0, c1 + 1):
        xs[c + 1] = xs[c] + col_w[c]
    for r in range(r0, r1 + 1):
        ys[r + 1] = ys[r] + row_h[r]
    W, H = max(1, int(round(xs[c1 + 1]))), max(1, int(round(ys[r1 + 1])))
    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    filled_rows = set()
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            if ws.cell(row=r, column=c).value not in (None, ""):
                filled_rows.add(r)
    for rg in ws.merged_cells.ranges:
        for r in range(rg.min_row, rg.max_row + 1):
            filled_rows.add(r)

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

    def _wl(t, ft, avail_w):
        out, cur = [], ""
        for ch in str(t):
            if draw.textlength(cur + ch, font=ft) <= avail_w or not cur:
                cur += ch
                continue
            cut = -1
            for i in range(len(cur) - 1, max(0, int(len(cur) * 0.55)) - 1, -1):
                if cur[i] in BREAK_AFTER:
                    cut = i + 1
                    break
            if cut > 0:
                out.append(cur[:cut])
                cur = cur[cut:] + ch
            else:
                out.append(cur)
                cur = ch
        out.append(cur)
        return out

    line_w = max(1, int(round(s)))
    for r in range(r0, r1 + 1):
        if r not in filled_rows:
            continue
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
            pad = max(1.0, 2.0 * s)
            avail_w = max(1.0, x1 - x0 - 2 * pad)
            lh_px = size * 1.22
            max_lines = max(1, int((y1 - y0 - 2 * pad) / lh_px))
            lines = _wl(txt, font, avail_w)
            if "\n" not in txt:
                one_w = draw.textlength(txt, font=font)
                if one_w > avail_w:
                    fit = (avail_w * 0.97) / one_w
                    if fit >= 0.82:
                        font = _font(size * fit, f.b)
                        size *= fit
                        lh_px = size * 1.22
                        lines = [txt]
            if len(lines) > max_lines:
                for k in (0.95, 0.9, 0.85, 0.8, 0.75, 0.7):
                    f2 = _font(size * k, f.b)
                    l2 = _wl(txt, f2, avail_w)
                    if len(l2) <= max_lines:
                        font, lines = f2, l2
                        size *= k
                        lh_px = size * 1.22
                        break
                else:
                    lines = lines[:max_lines]
                    sys.stderr.write("[提示] %s!%s 文本超出行高，已截断\n"
                                     % (ws.title, cell.coordinate))
            total_h = lh_px * len(lines)
            ty = y0 + max(pad, (y1 - y0 - total_h) / 2)
            if not (cell.alignment and cell.alignment.wrap_text) and len(lines) > 1:
                if (y1 - y0) < total_h * 1.4:
                    lines, total_h = lines[:1], lh_px
                    ty = y0 + max(pad, (y1 - y0 - total_h) / 2)
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


def _compose_raster(sheet_img, label, period_cn, page_no, total, land, dpi, font_path):
    from PIL import Image, ImageDraw, ImageFont

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


# ----------------------------------------------------------------- 导出
def export(book, out, sheets, font_path="", period_cn="",
           orientation="auto", dpi=600, scale=None, engine="vector"):
    font_path = find_font(font_path)
    fontname = register_font(font_path)
    wb = openpyxl.load_workbook(book, data_only=True)
    picked = []
    for no in sheets:
        name = next((n for n in wb.sheetnames if n.startswith(str(no) + "-")), None)
        if not name:
            sys.stderr.write("[中止] 工作簿缺少第 %s 张表：%s\n" % (no, wb.sheetnames))
            raise SystemExit(1)
        picked.append((no, name))

    if not picked:
        sys.stderr.write("[中止] 没有可导出的工作表（--sheets 为空）\n")
        raise SystemExit(1)

    orient = _decide_global(wb, picked, orientation)

    # 公式单元格若没有缓存值，data_only=True 读出来是 None，PDF 会静默空白
    try:
        wb_raw = openpyxl.load_workbook(book, data_only=False)
        for no, name in picked:
            nf = sum(1 for row in wb_raw[name].iter_rows() for cell in row
                     if cell.data_type == "f")
            if nf:
                sys.stderr.write("[提示] %s 有 %d 个公式单元格；若未用 Excel 保存过缓存值，"
                                 "PDF 中会显示为空白\n" % (name, nf))
        wb_raw.close()
    except Exception:
        pass

    if engine == "raster":
        from reportlab.lib.utils import ImageReader
        c, prev = None, None
        total = len(picked)
        meta = []
        sys.stderr.write("[提示] raster 后端仍为「一表一页」，不参与分页；"
                         "需要分页请用默认 vector\n")
        for i, (no, name) in enumerate(picked, 1):
            ws = wb[name]
            plan = _plan(ws, orient)
            land = plan["land"]
            wpt, hpt = plan["wpt"], plan["hpt"]
            pw, ph = int(wpt / 72 * dpi), int(hpt / 72 * dpi)
            avail_w = pw - 2 * int(PAGE_MARGIN_IN * dpi)
            avail_h = ph - int(HEAD_IN * dpi) - int(FOOT_IN * dpi)
            img = _render_sheet_raster(ws, font_path, avail_w, avail_h)
            page, size = _compose_raster(img, name, period_cn, i, total, land, dpi, font_path)
            if c is None:
                c = pdf_canvas.Canvas(out, pagesize=size)
                c.setTitle("铺货代销对账签章件")
                c.setAuthor("")
            elif size != prev:
                c.setPageSize(size)
            c.drawImage(ImageReader(page), 0, 0, width=size[0], height=size[1])
            c.showPage()
            del img, page
            prev = size
            meta.append({"sheet": no, "pages": 1,
                         "orientation": "横向" if land else "纵向",
                         "scale": round(plan["s"], 3),
                         "font_pt": round(plan["main_pt"], 2)})
        wb.close()
        c.save()
        return out, total, meta

    plans = [(no, name, _plan(wb[name], orient)) for no, name in picked]
    for no, name, plan in plans:
        if plan.get("hidden_cols") or plan.get("hidden_rows"):
            sys.stderr.write("[提示] %s 跳过隐藏列 %d 个、隐藏行 %d 行（不印进签章件）\n"
                             % (name, len(plan.get("hidden_cols") or ()),
                                len(plan.get("hidden_rows") or ())))
        if plan.get("empty"):
            sys.stderr.write("[提示] %s 无内容，已跳过\n" % name)
    total = sum(len(p["pages"]) for _, _, p in plans)
    if total == 0:
        sys.stderr.write("[中止] 所选工作表均为空，没有可导出的页\n")
        raise SystemExit(1)
    c, page_no = None, 0
    for no, name, plan in plans:
        ws = wb[name]
        for idx, rows in enumerate(plan["pages"]):
            page_no += 1
            if c is None:
                c = pdf_canvas.Canvas(out, pagesize=(plan["wpt"], plan["hpt"]))
                c.setTitle("铺货代销对账签章件")
                c.setAuthor("")
            draw_rows(c, ws, fontname, plan, rows, page_no, total, name,
                      period_cn, plan["hdr"][idx], is_cont=idx > 0)
            c.showPage()
    wb.close()
    if c is None:
        raise SystemExit("[中止] 没有可导出的工作表")
    c.save()
    meta = [{"sheet": no, "pages": len(p["pages"]),
             "orientation": "横向" if p["land"] else "纵向",
             "scale": round(p["s"], 3), "font_pt": round(p["main_pt"], 2)}
            for no, _, p in plans]
    return out, total, meta


def main():
    ap = argparse.ArgumentParser(description="活动区域 → A4 签章 PDF（宽度填满、高度切页）")
    ap.add_argument("--book", required=True)
    ap.add_argument("--out", default="")
    ap.add_argument("--sheets", default="00,01,02,08")
    ap.add_argument("--font", default="")
    ap.add_argument("--period-cn", default="")
    ap.add_argument("--orientation", default="auto",
                    choices=("auto", "landscape", "portrait"))
    ap.add_argument("--engine", default="vector", choices=("vector", "raster"),
                    help="vector=矢量文字（默认，宽度填满+切页）；raster=位图兜底（一表一页）")
    ap.add_argument("--dpi", type=int, default=600, help="仅 raster 后端使用")
    a = ap.parse_args()
    if not os.path.exists(a.book):
        sys.stderr.write("[中止] 找不到工作簿：%s\n" % a.book)
        return 2
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.book)), "对账-签章件.pdf")
    outdir = os.path.dirname(os.path.abspath(out))
    if not os.path.isdir(outdir):
        sys.stderr.write("[中止] 输出目录不存在：%s\n" % outdir)
        return 2
    if not os.access(outdir, os.W_OK):
        sys.stderr.write("[中止] 输出目录不可写：%s\n" % outdir)
        return 2
    sheets = [x.strip() for x in a.sheets.split(",") if x.strip()]
    path, n, meta = export(a.book, out, sheets, a.font, a.period_cn,
                           a.orientation, a.dpi, None, a.engine)
    print(json.dumps({"pdf": os.path.abspath(path), "pages": n,
                      "sheets": meta, "engine": a.engine,
                      "bytes": os.path.getsize(path)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
