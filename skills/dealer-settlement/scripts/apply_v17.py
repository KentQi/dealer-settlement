#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v1.7 附表渲染器：按 _口径确认.json 生成 08 / 09，并做交付门禁。

工作流位置：load_master.py -> recon_engine.py -> apply_v17.py -> L6 复核。

- 引擎只负责 00–07 的结算计算与六条恒等式；附表由本脚本按已确认口径渲染。
- 08 试用装与中小样：甲方期初 → 期间发货 → 乙方使用 → 甲方期末，期末结转下期。
- 09 效期预警：能力完整保留，由 expiry.enabled 开关控制；关闭时不建表、
  且 00 表由引擎按同一开关裁剪，全簿零效期字样。
- incremental 模式：以已有工作簿为底，01–07 原样保留，00 表按本期口径重生成。
- 最后跑对外文案门禁，并按配置导出签章 PDF。
"""
import argparse
import copy
import datetime
import json
import os
import shutil
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

_SK = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _SK)
from check_public_copy import clean_properties, scan  # noqa: E402
import koujing  # noqa: E402

NAVY, BLUE, WHITE = "1F3864", "2E5C9A", "FFFFFF"
DGREY, GREY, VLBLUE, LBLUE = "595959", "F2F2F2", "D9E2F3", "EAF0F8"
RED, LRED, AMBER, LAMBER = "C00000", "FCE4E4", "BF8F00", "FFF2CC"
GREEN, LGREEN = "2E7D32", "E8F3E8"
THIN = Side(style="thin", color="BFBFBF")
MED = Side(style="medium", color="1F3864")


def F(sz=9, b=False, c="000000"):
    return Font(name="微软雅黑", size=sz, bold=b, color=c)


def fill(c):
    return PatternFill("solid", fgColor=c)


def A(h="center", v="center", wrap=True, ind=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=ind)


def mput(ws, r1, c1, r2, c2, v, font=None, al=None, fl=None, border=THIN):
    if r2 > r1 or c2 > c1:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    cell = ws.cell(row=r1, column=c1, value=v)
    cell.font = font or F()
    cell.alignment = al or A()
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            x = ws.cell(row=r, column=c)
            x.border = Border(left=border, right=border, top=border, bottom=border)
            if fl:
                x.fill = fill(fl)
    return cell


def num(cell, v, fmt="#,##0"):
    cell.value = v
    if isinstance(v, (int, float)):
        cell.number_format = fmt
    cell.font = F()
    cell.alignment = A()
    return cell


def sec(ws, r, c1, c2, title):
    mput(ws, r, c1, r, c2, title, F(10, True, WHITE), A("left", "center", True, 1), NAVY)
    ws.row_dimensions[r].height = 24


def hdr(ws, r, c1, labels, height=26):
    for i, t in enumerate(labels):
        mput(ws, r, c1 + i, r, c1 + i, t, F(8.5, True, WHITE), A(), NAVY, MED)
    ws.row_dimensions[r].height = height


def setup(ws, widths, orientation="landscape"):
    ws.sheet_view.showGridLines = False
    ws.print_options.gridLines = False
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.page_setup.orientation = orientation
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.oddFooter.center.text = "第 &P 页 / 共 &N 页"
    ws.oddFooter.center.size = 8


def fit_h(txt, cpl=42.0):
    return max(20, 15 * (1 + len(str(txt)) / cpl))


def months_between(base, exp):
    m = (exp.year - base.year) * 12 + (exp.month - base.month)
    if exp.day < base.day:
        m -= 1
    return m


def parse_day(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if not v:
        return None
    s = str(v)[:10]
    for f in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return datetime.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def copy_sheet(src_ws, dst_wb, name):
    """跨工作簿整表复制（值 + 样式 + 合并区 + 行高列宽 + 打印定义）。"""
    if name in dst_wb.sheetnames:
        idx = dst_wb.sheetnames.index(name)
        del dst_wb[name]
    else:
        idx = len(dst_wb.sheetnames)
    ws = dst_wb.create_sheet(name, idx)
    for row in src_ws.iter_rows():
        for c in row:
            d = ws.cell(row=c.row, column=c.column, value=c.value)
            if c.has_style:
                d.font = copy.copy(c.font)
                d.fill = copy.copy(c.fill)
                d.border = copy.copy(c.border)
                d.alignment = copy.copy(c.alignment)
                d.number_format = c.number_format
    for rg in src_ws.merged_cells.ranges:
        ws.merge_cells(str(rg))
    for k, dim in src_ws.row_dimensions.items():
        ws.row_dimensions[k].height = dim.height
    for k, dim in src_ws.column_dimensions.items():
        ws.column_dimensions[k].width = dim.width
    ws.freeze_panes = src_ws.freeze_panes
    ws.print_area = src_ws.print_area
    ws.print_title_rows = src_ws.print_title_rows
    ws.page_setup.orientation = src_ws.page_setup.orientation
    ws.page_setup.fitToWidth = src_ws.page_setup.fitToWidth
    if src_ws.sheet_properties.tabColor:
        ws.sheet_properties.tabColor = src_ws.sheet_properties.tabColor
    ws.oddFooter.center.text = src_ws.oddFooter.center.text
    ws.oddFooter.right.text = src_ws.oddFooter.right.text
    return ws


def ledger_rows(kj):
    out = []
    for x in (kj.get("decisions", {}).get("usage", {}).get("ledger") or []):
        y = dict(x)
        for k in ("opening", "shipped", "used", "sales"):
            y[k] = int(y.get(k) or 0)
        y["closing"] = y["opening"] + y["shipped"] - y["used"]
        out.append(y)
    return out


def build_08(wb, kj, sidecar):
    name = "08-" + (kj["decisions"].get("sheet_map", {}).get("08") or "试用装与中小样")
    ws = wb.create_sheet(name)
    setup(ws, [6, 34, 10, 30, 12, 11, 11, 11, 11, 14])
    rows = ledger_rows(kj)
    mput(ws, 2, 1, 2, 10, "附 表 08　试用装 / 中小样 数量确认", F(15, True, NAVY), A("center"))
    ws.row_dimensions[2].height = 32
    mput(ws, 3, 1, 3, 10,
         "本表所列商品不参与货款结算　·　请乙方核对期初、期间发货、使用与期末数量后签认",
         F(8.5, c=DGREY), A("center"))
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 6

    notes = [
        "1. 试用装、中小样不参与货款结算，本期单独列示，供双方核对数量。",
        "2. 「甲方期初」为上期结转数量，本期为首期，期初为 0；「期间发货」为本期发出数量；「乙方使用」为本期使用数量，由乙方核对确认。",
        "3. 「甲方期末」＝期初 ＋ 期间发货 － 乙方使用，结转下期。",
        "4. 请乙方核对下列数量并在签章回执中书面确认；如有差异，请附明细说明。",
    ]
    r = 5
    sec(ws, r, 1, 10, "一、说明"); r += 1
    for t in notes:
        mput(ws, r, 1, r, 10, t, F(8.5), A("left", "center", True, 1))
        ws.row_dimensions[r].height = fit_h(t, 60); r += 1
    r += 1

    sec(ws, r, 1, 10, "二、数量汇总"); r += 1
    hdr(ws, r, 1, ["类型", "", "甲方期初", "期间发货", "乙方使用", "甲方期末", "", "", "", "备注"])
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=9)
    r += 1
    totals = {}
    kinds = []
    for _x in rows:
        if _x.get("kind") and _x["kind"] not in kinds:
            kinds.append(_x["kind"])
    for kind in kinds:
        items = [x for x in rows if x.get("kind") == kind]
        if not items:
            continue
        vals = [sum(x[k] for x in items) for k in ("opening", "shipped", "used", "closing")]
        totals[kind] = vals
        mput(ws, r, 1, r, 2, kind, F(9, True), A("center"), VLBLUE)
        for c, v in zip((3, 4, 5, 6), vals):
            num(ws.cell(row=r, column=c), v)
        for c in (7, 8, 9, 10):
            ws.cell(row=r, column=c, value="—" if c == 10 else None)
            ws.cell(row=r, column=c).alignment = A()
        ws.row_dimensions[r].height = 22; r += 1
    allv = [sum(v[i] for v in totals.values()) for i in range(4)]
    mput(ws, r, 1, r, 2, "合　计", F(9, True, WHITE), A("center"), NAVY)
    for c, v in zip((3, 4, 5, 6), allv):
        cell = num(ws.cell(row=r, column=c), v)
        cell.font = F(9, True, WHITE); cell.fill = fill(NAVY)
    for c in (7, 8, 9, 10):
        ws.cell(row=r, column=c).fill = fill(NAVY)
    ws.row_dimensions[r].height = 24; r += 2

    sec(ws, r, 1, 10, "三、逐项明细"); r += 1
    hdr(ws, r, 1, ["序号", "商品名称", "类型", "对应正装", "本期正装销量",
                   "甲方期初", "期间发货", "乙方使用", "甲方期末", "备注"], 30)
    r += 1
    for i, x in enumerate(rows, 1):
        bg = WHITE if i % 2 else GREY
        vals = [i, x.get("name", ""), x.get("kind", ""), x.get("link", "") or "—",
                x.get("sales", 0), x["opening"], x["shipped"], x["used"], x["closing"], "—"]
        for c, v in enumerate(vals, 1):
            cell = mput(ws, r, c, r, c, v,
                        F(8.5) if c != 2 else F(8.5),
                        A("left", "center", True, 0) if c in (2, 4) else A(), bg)
            if isinstance(v, int):
                cell.number_format = "#,##0"
        ws.row_dimensions[r].height = 24; r += 1
    mput(ws, r, 1, r, 2, "合　计", F(9, True), A("center"), LBLUE)
    for c in (3, 4, 5):
        ws.cell(row=r, column=c, value="—").alignment = A()
    for c, v in zip((6, 7, 8, 9), allv):
        num(ws.cell(row=r, column=c), v).font = F(9, True, NAVY)
    ws.cell(row=r, column=10, value="—")
    r += 2

    sec(ws, r, 1, 10, "四、乙方确认"); r += 1
    mput(ws, r, 1, r, 10,
         "本单位已核对上述期初、期间发货、使用与期末数量，如有差异另附说明；无异议请签章确认。",
         F(9), A("left", "center", True, 1))
    ws.row_dimensions[r].height = 26; r += 1
    mput(ws, r, 1, r, 5, "乙方（盖章）：", F(9), A("left", "center", False, 1), WHITE, THIN)
    mput(ws, r, 6, r, 10, "日　期：　　　年　　月　　日", F(9), A("left", "center", False, 1), WHITE, THIN)
    ws.row_dimensions[r].height = 85
    ws.freeze_panes = "A19"
    ws.print_title_rows = "18:18"
    ws.print_area = "A1:J%d" % r
    ws.sheet_properties.tabColor = "FFC00000"
    return ws, allv


def build_09(wb, kj, sidecar):
    name = "09-" + (kj["decisions"].get("sheet_map", {}).get("09") or "效期预警")
    ws = wb.create_sheet(name)
    setup(ws, [6, 34, 11, 14, 14, 12, 11, 14, 13, 14])
    ex = kj["decisions"].get("expiry") or {}
    tiers = sorted([int(x) for x in (ex.get("tiers") or [32, 28, 24])], reverse=True)
    base = parse_day(ex.get("base") if ex.get("base") not in (None, "asof") else None) \
        or parse_day(sidecar.get("asof")) or datetime.date.today()
    labels = (["已过期"] + ["低于 %d 个月" % t for t in reversed(tiers)]
              + ["正常（≥%d 个月）" % tiers[0]])

    ledger = ledger_rows(kj)
    sample_closing = {x["name"]: x["closing"] for x in ledger if x.get("kind") == "中小样"}
    items = []
    for b in sidecar.get("batch") or []:
        d = parse_day(b.get("exp"))
        if d is None:
            continue
        q = int(b.get("qty") or 0)
        items.append(dict(name=b.get("name", ""), bar=b.get("bar", ""),
                          lot=b.get("lot", "") or "—", exp=d, qty=q, kind="正装"))
    covered = set()
    for b in items:
        for nm, q in sample_closing.items():
            if nm and (nm in b["name"] or b["name"] in nm):
                covered.add(nm)
    if ex.get("sample_qty_basis") == "台账期末":
        for nm, q in sample_closing.items():
            if nm not in covered:
                items.append(dict(name=nm, bar="", lot="—", exp=None, qty=q, kind="中小样"))
    for it in items:
        if it["exp"] is None:
            it["months"], it["tier"], it["days"] = None, "无有效期数据", None
            continue
        m = months_between(base, it["exp"])
        it["months"] = m
        it["days"] = (it["exp"] - base).days
        if m < 0 or it["days"] < 0:
            it["tier"] = "已过期"
        else:
            it["tier"] = "正常（≥%d 个月）" % tiers[0]
            for t in reversed(tiers):
                if m < t:
                    it["tier"] = "低于 %d 个月" % t
                    break

    tier_fill = {"已过期": LRED, "无有效期数据": LRED}
    for i, t in enumerate(reversed(tiers)):
        tier_fill["低于 %d 个月" % t] = (LRED if i == 0 else LAMBER if i == 1 else "FFF7E0")
    tier_fill["正常（≥%d 个月）" % tiers[0]] = WHITE

    mput(ws, 2, 1, 2, 10, "附 表 09　商品批次效期提示", F(15, True, NAVY), A("center"))
    ws.row_dimensions[2].height = 32
    mput(ws, 3, 1, 3, 10,
         f"对账期间：{sidecar.get('period', '')}　·　基准日：{base}　·　"
         "批次与到期日取自甲方批次记录，件数为截至基准日的在库数量",
         F(8.5, c=DGREY), A("center"))
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 6

    r = 5
    sec(ws, r, 1, 10, "一、预警汇总（按剩余效期月数分档）"); r += 1
    hdr(ws, r, 1, ["预警档", "", "", "商品数", "件数", "", "", "", "", ""])
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=10)
    r += 1
    for t in labels:
        sub = [x for x in items if x["tier"] == t]
        fl = tier_fill.get(t, WHITE)
        mput(ws, r, 1, r, 3, t, F(9, True), A("center"), fl)
        num(ws.cell(row=r, column=4), len(sub)).fill = fill(fl)
        num(ws.cell(row=r, column=5), sum(x["qty"] for x in sub)).fill = fill(fl)
        for c in range(6, 11):
            ws.cell(row=r, column=c).fill = fill(fl)
        ws.row_dimensions[r].height = 22; r += 1
    alert = [x for x in items if x["tier"] not in ("正常（≥%d 个月）" % tiers[0],)]
    mput(ws, r, 1, r, 10,
         f"合计：{len(items)} 个商品 / {sum(x['qty'] for x in items):,} 件；"
         f"低于 {tiers[0]} 个月 {len(alert)} 个商品、{sum(x['qty'] for x in alert):,} 件。",
         F(9, True, NAVY), A("left", "center", True, 1), LBLUE)
    ws.row_dimensions[r].height = 24; r += 2

    sec(ws, r, 1, 10, "二、商品批次明细（按剩余月数升序）"); r += 1
    hdr(ws, r, 1, ["序号", "商品名称", "类型", "规格", "批次号", "有效期至",
                   "剩余月数", "件数", "预警档", "商品条码"], 30)
    r += 1
    for i, x in enumerate(sorted(items, key=lambda z: (z["months"] if z["months"] is not None else 9999,
                                                       z["exp"] or datetime.date.max, z["name"])), 1):
        bg = WHITE if i % 2 else GREY
        vals = [i, x["name"], x["kind"], "—", x["lot"], x["exp"],
                x["months"] if x["months"] is not None else "—", x["qty"],
                x["tier"], x["bar"] or "—"]
        for c, v in enumerate(vals, 1):
            cell = mput(ws, r, c, r, c, v, F(8.5),
                        A("left", "center", True, 0) if c == 2 else A(), bg)
            if isinstance(v, int):
                cell.number_format = "#,##0"
        cell = ws.cell(row=r, column=6)
        cell.number_format = "yyyy-mm-dd"
        ws.cell(row=r, column=9).fill = fill(tier_fill.get(x["tier"], WHITE))
        ws.row_dimensions[r].height = 24; r += 1
    r += 1

    sec(ws, r, 1, 10, "三、口径说明"); r += 1
    tier_txt = " / ".join("低于 %d 个月" % t for t in reversed(tiers))
    notes = [
        f"1. 件数口径：本表件数为截至 {base} 的在库数量。",
        f"2. 效期口径：批次号与失效日期取自甲方批次记录（{sidecar.get('asof', base)}）。",
        f"3. 分档口径：基准日为 {base}，按剩余效期月数分档——{tier_txt} / 正常；月数为日历月差，不足一个月向下取整。",
        "4. 本表仅提示剩余效期，供双方安排动销，不构成对乙方仓储作业的要求；具体处理方式由双方另行协商。",
    ]
    for t in notes:
        mput(ws, r, 1, r, 10, t, F(8.5), A("left", "center", True, 1))
        ws.row_dimensions[r].height = fit_h(t, 58); r += 1
    ws.freeze_panes = "A16"
    ws.print_area = "A1:J%d" % (r - 1)
    ws.sheet_properties.tabColor = "FFA6A6A6"
    return ws, len(items)


def main():
    ap = argparse.ArgumentParser(description="v1.7 附表渲染与交付门禁")
    ap.add_argument("--cfg", required=True, help="RECON_CFG JSON 文件路径")
    ap.add_argument("--book", default="", help="引擎输出工作簿；缺省取 CFG.out_file/base_dir")
    ap.add_argument("--sidecar", default="", help="引擎产出 _v17_data.json 路径")
    ap.add_argument("--no-gate", action="store_true", help="跳过对外门禁（仅供调试）")
    a = ap.parse_args()

    with open(a.cfg, "r", encoding="utf-8") as fh:
        cfg = json.load(fh)
    kj = cfg.get("koujing") or {}
    if not kj:
        sys.stderr.write("[中止] RECON_CFG 缺少 koujing，请先完成阶段 G。\n")
        return 2
    errors, _ = koujing.validate(kj, kj.get("scene"))
    if errors:
        sys.stderr.write("[中止] 口径确认记录无效：" + "；".join(errors) + "\n")
        return 2
    v17 = cfg.get("v17") or {}
    expiry_on_cfg = bool((kj["decisions"].get("expiry") or {}).get("enabled"))
    base = cfg.get("base_dir") or ""
    sc_path = a.sidecar or os.path.join(base, "_v17_data.json")
    with open(sc_path, "r", encoding="utf-8") as fh:
        sidecar = json.load(fh)
    expiry_on_sc = bool((sidecar.get("expiry") or {}).get("enabled"))
    if expiry_on_sc != expiry_on_cfg:
        sys.stderr.write("[中止] 口径记录与引擎产物的效期开关不一致 "
                         "（口径 %s / 引擎 %s）。口径变更后必须重跑 load_master 与 recon_engine。\n"
                         % (expiry_on_cfg, expiry_on_sc))
        return 2
    out = a.book or cfg.get("out_file") or os.path.join(base, cfg["period"]["out_name"])
    out = sidecar.get("out") or out
    src_book = out
    baseline = v17.get("input_workbook") or ""
    mode = v17.get("mode") or "full"
    if not baseline and sidecar.get("cfg_v17", {}).get("input_workbook"):
        baseline = sidecar["cfg_v17"]["input_workbook"]

    if mode == "incremental":
        kj.setdefault("decisions", {}).setdefault("delta", {})["frozen_sheets"] = \
            ["01", "02", "03", "04", "05", "06", "07"]
        cfg["koujing"] = kj
    if mode == "incremental" and baseline:
        if not os.path.exists(baseline):
            sys.stderr.write("[中止] 找不到已有工作簿：%s\n" % baseline)
            return 2
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        bak = os.path.join(os.path.dirname(baseline), "_备份",
                           os.path.basename(baseline).replace(".xlsx", "_增量前_%s.xlsx" % stamp))
        os.makedirs(os.path.dirname(bak), exist_ok=True)
        shutil.copy2(baseline, bak)
        eng = openpyxl.load_workbook(src_book)
        wb = openpyxl.load_workbook(baseline)
        copy_sheet(eng["00-对账说明"], wb, "00-对账说明")
        eng.close()
        out = baseline
        gate_baseline = bak
    else:
        if os.path.exists(out):
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            bk = os.path.join(os.path.dirname(out), "_备份",
                              os.path.basename(out).replace(".xlsx", "_出账前_%s.xlsx" % stamp))
            os.makedirs(os.path.dirname(bk), exist_ok=True)
            shutil.copy2(out, bk)
        wb = openpyxl.load_workbook(src_book)
        gate_baseline = ""

    for n in list(wb.sheetnames):
        if n[:2] in ("08", "09", "10") and ("效期" in n or "试用装" in n or "配赠" in n or "复核记录" in n):
            if "复核记录" not in n:
                del wb[n]

    build_08(wb, kj, sidecar)
    expiry_on = bool((kj["decisions"].get("expiry") or {}).get("enabled"))
    n09 = 0
    if expiry_on:
        _, n09 = build_09(wb, kj, sidecar)

    red = set((kj["decisions"].get("print") or {}).get("red_sheets") or ["00", "01", "02", "08"])
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = ("FFC00000" if ws.title[:2] in red else "FFA6A6A6")
    for no in ("00", "02"):
        ws = next((w for w in wb.worksheets if w.title.startswith(no + "-")), None)
        if ws and no == "00":
            ws.print_area = "A1:L%d" % ws.max_row
        elif ws:
            ws.print_title_rows = "10:10"
    wb.properties.creator = ""
    wb.properties.lastModifiedBy = ""
    wb.save(out)
    wb.close()
    clean_properties(out)

    if not a.no_gate:
        try:
            result = scan(out, cfg=cfg, baseline=gate_baseline, strict=False)
        except SystemExit:
            return 1
        if result["problems"]:
            sys.stderr.write("[对外门禁未通过] 不交付。\n")
            for p in result["problems"]:
                sys.stderr.write("  · " + p + "\n")
            return 1
        print(json.dumps({"ok": True, "book": os.path.abspath(out),
                          "mode": mode, "expiry_enabled": expiry_on,
                          "expiry_items": n09, "gate": result["scanned_strings"]},
                         ensure_ascii=False))
    else:
        print(json.dumps({"ok": True, "book": os.path.abspath(out), "mode": mode},
                         ensure_ascii=False))

    if v17.get("pdf"):
        try:
            import export_pdf
        except Exception as exc:
            print(json.dumps({"pdf_skipped": True,
                              "reason": "reportlab 不可用：%s" % exc},
                             ensure_ascii=False))
            return 0
        sheets = v17.get("pdf_sheets") or ["00", "01", "02", "08"]
        tag = cfg["period"].get("period") or str(cfg["period"].get("out_name", ""))[:6]
        pdf = os.path.join(os.path.dirname(out), "%s对账-签章件.pdf" % tag)
        export_pdf.export(out, pdf, sheets, period_cn=cfg["period"].get("period_cn", ""))
        print(json.dumps({"pdf": os.path.abspath(pdf), "bytes": os.path.getsize(pdf)},
                         ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
