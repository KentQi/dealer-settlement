#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""对外文案与增量冻结门禁（保存后、交付前）。

机械断言，不依赖 LLM 语义判断：
  A1 禁止词 l1 全簿零命中
  A2 internal_only 内容零泄漏
  A3 关闭效期时全簿零效期字样
  A4 文档属性无作者/工具痕迹
  A5 增量模式 frozen_sheets 逐格零差异
  A6 附表编号与交叉引用一致
"""
import argparse
import json
import os
import re
import sys
import zipfile

import openpyxl

_SK = os.path.dirname(os.path.abspath(__file__))
_COPY_PATH = os.path.join(_SK, "..", "assets", "copy", "对外文案.json")


def load_copy():
    with open(os.path.abspath(_COPY_PATH), "r", encoding="utf-8") as fh:
        return json.load(fh)


def clean_properties(path):
    """zip 成员级清除作者痕迹。openpyxl 保存时会回填 creator，必须在最后一次保存后调用。"""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                s = data.decode("utf-8")
                s = re.sub(r"<dc:creator>.*?</dc:creator>", "<dc:creator></dc:creator>", s)
                s = re.sub(r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>",
                           "<cp:lastModifiedBy></cp:lastModifiedBy>", s)
                s = re.sub(r"<dcterms:(created|modified)[^>]*>.*?</dcterms:\1>", "", s)
                data = s.encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, path)


def _all_strings(wb):
    """产出 (位置, 文本)，覆盖单元格、页眉页脚、表名、定义名称、批注、打印定义。"""
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    yield "%s!%s" % (ws.title, c.coordinate), c.value
        for part in ("oddHeader", "oddFooter", "evenHeader", "evenFooter",
                     "firstHeader", "firstFooter"):
            obj = getattr(ws, part, None)
            if obj is None:
                continue
            for side in ("left", "center", "right"):
                txt = getattr(getattr(obj, side, None), "text", None)
                if txt:
                    yield "%s!%s.%s" % (ws.title, part, side), txt
        yield "%s!表名" % ws.title, ws.title
        if ws.print_area:
            yield "%s!打印区" % ws.title, str(ws.print_area)
        for row in ws.iter_rows():
            for c in row:
                if c.comment and c.comment.text:
                    yield "%s!%s!批注" % (ws.title, c.coordinate), c.comment.text
    for name in getattr(wb, "defined_names", {}):
        yield "定义名称", str(name)


def _core_props(path):
    out = []
    try:
        with zipfile.ZipFile(path) as zf:
            for member in ("docProps/core.xml", "docProps/app.xml"):
                if member in zf.namelist():
                    out.append((member, zf.read(member).decode("utf-8", "ignore")))
    except (zipfile.BadZipFile, OSError):
        pass
    return out


def _fingerprint(ws):
    """冻结表的结构指纹：值、公式、格式、合并区、行高列宽、打印定义。"""
    cells = {}
    for row in ws.iter_rows():
        for c in row:
            cells[c.coordinate] = (
                c.value, c.data_type, c.number_format,
                c.font.name, c.font.sz, c.font.b,
                str(c.font.color.rgb) if c.font.color else None,
                str(c.fill.fgColor.rgb) if c.fill and c.fill.fgColor else None,
                c.alignment.horizontal, c.alignment.wrap_text,
            )
    return {
        "cells": cells,
        "merged": sorted(str(x) for x in ws.merged_cells.ranges),
        "rows": {r: d.height for r, d in ws.row_dimensions.items()},
        "cols": {k: d.width for k, d in ws.column_dimensions.items()},
        "print_area": str(ws.print_area),
        "print_titles": str(ws.print_title_rows),
        "freeze": str(ws.freeze_panes),
        "tab": str(ws.sheet_properties.tabColor.rgb) if ws.sheet_properties.tabColor else None,
    }


def scan(path, cfg=None, baseline=None, strict=True):
    """返回 findings；strict=True 时有问题直接 SystemExit(1)。"""
    cfg = cfg or {}
    dec = (cfg.get("koujing") or {}).get("decisions") or {}
    expiry = dec.get("expiry") or {}
    expiry_on = bool(expiry.get("enabled"))
    internal = [str(x) for x in (cfg.get("koujing") or {}).get("internal_only", []) if x]
    delta = dec.get("delta") or {}
    frozen = delta.get("frozen_sheets") or []
    copy = load_copy()

    problems = []
    wb = openpyxl.load_workbook(path, data_only=False)
    strings = list(_all_strings(wb))
    for member, raw in _core_props(path):
        strings.append((member, raw))

    for item in copy["l1_banned"]:
        pat = item["pattern"]
        hits = [(loc, txt) for loc, txt in strings if pat in txt]
        if hits:
            loc, txt = hits[0]
            problems.append("A1 禁止词「%s」(%s) 命中 %d 处，首处 %s：%s" %
                            (pat, item["label"], len(hits), loc, txt[:70]))

    for val in internal:
        hits = [(loc, txt) for loc, txt in strings if val and val in txt]
        if hits:
            problems.append("A2 内部内容泄漏「%s」，首处 %s" % (val[:40], hits[0][0]))

    if not expiry_on:
        for term in copy["expiry_terms"]:
            hits = [(loc, txt) for loc, txt in strings if term in txt]
            if hits:
                problems.append("A3 本期关闭效期，但仍出现「%s」%d 处，首处 %s" %
                                (term, len(hits), hits[0][0]))
                break

    for member, raw in _core_props(path):
        for tag in ("dc:creator", "cp:lastModifiedBy"):
            m = re.search(r"<%s>(.*?)</%s>" % (tag, tag), raw)
            if m and m.group(1).strip() not in ("", "None"):
                problems.append("A4 文档属性 %s 残留「%s」" % (tag, m.group(1)[:40]))

    sm = dec.get("sheet_map") or {}
    if sm:
        names = wb.sheetnames
        for no, title in sm.items():
            target = "%s-%s" % (no, title)
            if no == "09" and not expiry_on:
                if any(n.endswith("-效期预警") for n in names):
                    problems.append("A6 本期关闭效期，却仍存在效期工作表")
            elif expiry_on or no == "08":
                if target not in names:
                    problems.append("A6 缺少按口径确认的表：%s" % target)
        actual_nos = {n[:2] for n in names if re.match(r"^[0-9]{2}-", n)}
        for loc, txt in strings:
            for m in re.finditer(r"附表\s*([0-9]{2})", txt):
                no = m.group(1)
                if no not in actual_nos:
                    problems.append("A6 悬空交叉引用「附表 %s」于 %s" % (no, loc))
                    break
    wb.close()

    if baseline and os.path.exists(baseline):
        wa = openpyxl.load_workbook(baseline, data_only=False)
        wnew = openpyxl.load_workbook(path, data_only=False)
        for no in frozen:
            name = next((n for n in wa.sheetnames if n.startswith(str(no) + "-")), None)
            if not name or name not in wnew.sheetnames:
                problems.append("A5 基线缺少冻结表 %s" % no)
                continue
            fa, fb = _fingerprint(wa[name]), _fingerprint(wnew[name])
            for key in fa:
                if fa[key] != fb[key]:
                    if key == "cells":
                        diff = [k for k in fa[key] if fa[key].get(k) != fb[key].get(k)]
                        problems.append("A5 冻结表 %s 有 %d 个单元格变化：%s" %
                                        (name, len(diff), ", ".join(diff[:5])))
                    else:
                        problems.append("A5 冻结表 %s 的 %s 发生变化" % (name, key))
        wa.close()
        wnew.close()

    result = {
        "ok": not problems,
        "book": os.path.abspath(path),
        "expiry_enabled": expiry_on,
        "frozen_sheets": frozen,
        "scanned_strings": len(strings),
        "problems": problems,
    }
    if problems and strict:
        sys.stderr.write("[对外门禁未通过] 不交付。\n")
        for p in problems:
            sys.stderr.write("  · " + p + "\n")
        raise SystemExit(1)
    return result


def main():
    ap = argparse.ArgumentParser(description="对外文案与增量冻结门禁")
    ap.add_argument("--book", required=True)
    ap.add_argument("--cfg", default="", help="RECON_CFG JSON 文件路径，或直接给 JSON 字符串")
    ap.add_argument("--baseline", default="")
    ap.add_argument("--clean-props", action="store_true", help="清除文档属性后退出")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    if a.clean_props:
        clean_properties(a.book)
        print("OK")
        return 0
    cfg = {}
    if a.cfg:
        try:
            cfg = json.loads(a.cfg) if a.cfg.strip().startswith("{") else json.load(
                open(a.cfg, "r", encoding="utf-8"))
        except Exception as e:
            sys.stderr.write("[中止] --cfg 不是合法 JSON：%s\n" % e)
            return 2
    res = scan(a.book, cfg=cfg, baseline=a.baseline, strict=not a.json)
    print(json.dumps(res, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
