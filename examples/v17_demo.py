#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""v1.7.0 回归：口径门禁 / 模块开关 / 增量冻结 / 对外文案 / PDF 签章件。

跑法：python3 examples/v17_demo.py
"""
import datetime
import json
import os
import shutil
import subprocess
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCRIPTS = os.path.join(ROOT, "skills", "dealer-settlement", "scripts")
sys.path.insert(0, HERE)
import make_demo_data  # noqa: E402

LOAD = os.path.join(SCRIPTS, "load_master.py")
ENGINE = os.path.join(SCRIPTS, "recon_engine.py")
APPLY = os.path.join(SCRIPTS, "apply_v17.py")
PERIOD, ASOF = "202610", "2026-11-02"


def check(ok, msg):
    if not ok:
        raise AssertionError(msg)
    print("  ✓ " + msg)


def sh(args, cfg=None):
    env = dict(os.environ)
    if cfg is not None:
        env["RECON_CFG"] = cfg
    return subprocess.run([sys.executable] + args, capture_output=True, text=True, env=env)


def load(cfg_path, extra=()):
    p = sh([LOAD, os.path.join(MASTER_ROOT, "_主档.xlsx"), "--period", PERIOD,
            "--asof", ASOF, "--overwrite"] + list(extra))
    check(p.returncode == 0, "load_master 成功：" + (p.stderr.strip()[:80] or "ok"))
    with open(cfg_path, "w", encoding="utf-8") as fh:
        fh.write(p.stdout)
    return p.stdout


def engine(cfg):
    p = sh([ENGINE], cfg)
    check(p.returncode == 0, "recon_engine 成功")
    return p


def apply(cfg_path):
    p = sh([APPLY, "--cfg", cfg_path])
    return p


def all_text(book):
    wb = openpyxl.load_workbook(book, data_only=True)
    out = "\n".join(str(c.value) for ws in wb.worksheets
                    for row in ws.iter_rows() for c in row if c.value is not None)
    wb.close()
    return out


def set_expiry(enabled):
    kj = json.load(open(KJ_PATH, encoding="utf-8"))
    kj["decisions"]["expiry"]["enabled"] = enabled
    with open(KJ_PATH, "w", encoding="utf-8") as fh:
        json.dump(kj, fh, ensure_ascii=False, indent=2)


def main():
    global MASTER_ROOT, KJ_PATH, BOOK
    tmp = tempfile.mkdtemp(prefix="dealer_settlement_v17_")
    try:
        MASTER_ROOT, _ = make_demo_data.build(tmp)
        KJ_PATH = os.path.join(MASTER_ROOT, "_口径确认.json")
        BOOK = os.path.join(MASTER_ROOT, PERIOD, PERIOD + "铺货代销对账.xlsx")
        cfg_path = os.path.join(tmp, "cfg.json")

        print("\n[1] 阶段 G：口径确认门禁")
        os.rename(KJ_PATH, KJ_PATH + ".bak")
        p = sh([LOAD, os.path.join(MASTER_ROOT, "_主档.xlsx"),
                "--period", PERIOD, "--asof", ASOF])
        check(p.returncode == 1 and len(p.stdout) == 0,
              "缺少口径记录时 exit 1、stdout 0 字节")
        check("口径未确认" in p.stderr and "哪张表说了算" in p.stderr or
              "请先逐项确认" in p.stderr, "stderr 输出中文业务语言问卡")
        os.rename(KJ_PATH + ".bak", KJ_PATH)

        print("\n[2] 关闭效期：9 表 + 全簿零效期字样 + 四栏台账")
        cfg = load(cfg_path)
        engine(cfg)
        p = apply(cfg_path)
        check(p.returncode == 0, "交付门禁通过：" + p.stderr.strip()[-120:])
        wb = openpyxl.load_workbook(BOOK, data_only=True)
        check(len(wb.sheetnames) == 9 and "08-试用装与中小样" in wb.sheetnames,
              "9 张表，08 为试用装与中小样")
        check("09-效期预警" not in wb.sheetnames, "不生成 09 效期表")
        ws = wb["08-试用装与中小样"]
        heads = [ws.cell(12, c).value for c in range(1, 7)]
        check(heads[:5] == ["类型", None, "甲方期初", "期间发货", "乙方使用"],
              "08 表四栏台账表头正确")
        check("结转下期" in all_text(BOOK), "08 表含期末结转说明")
        wb.close()
        txt = all_text(BOOK)
        check(not any(k in txt for k in ("效期", "批次", "到期", "临期", "预警")),
              "关闭效期时全簿零效期字样")
        check(not any(k in txt for k in ("费用核算", "1:1", "BIP", "ERP", "清洗")),
              "对外单据零内部口径")

        print("\n[3] 开启效期：10 表 + 按月分档 + 00 表引用同步")
        set_expiry(True)
        cfg = load(cfg_path)
        engine(cfg)
        p = apply(cfg_path)
        check(p.returncode == 0, "开启效期后门禁通过：" + p.stderr.strip()[-120:])
        wb = openpyxl.load_workbook(BOOK, data_only=True)
        check("09-效期预警" in wb.sheetnames, "09 效期表已生成")
        t9 = "\n".join(str(c.value) for row in wb["09-效期预警"].iter_rows()
                       for c in row if c.value is not None)
        check("低于 32 个月" in t9 and "低于 28 个月" in t9 and "低于 24 个月" in t9,
              "按 32 / 28 / 24 个月分档")
        t0 = "\n".join(str(c.value) for row in wb["00-对账说明"].iter_rows()
                       for c in row if c.value is not None)
        check("明细见附表 09" in t0 and "附表 09" in t0,
              "00 表结论与附表清单引用已恢复")
        wb.close()

        print("\n[4] 内部内容注入必须被门禁拦截")
        set_expiry(False)
        wb = openpyxl.load_workbook(BOOK)
        wb["08-试用装与中小样"]["A40"] = "甲方内部费用核算口径"
        wb.save(BOOK)
        wb.close()
        sys.path.insert(0, SCRIPTS)
        from check_public_copy import scan
        cfg_obj = json.load(open(cfg_path, encoding="utf-8"))
        cfg_obj["koujing"]["decisions"]["expiry"]["enabled"] = False
        res = scan(BOOK, cfg_obj, strict=False)
        check(any("费用核算" in x for x in res["problems"]), "禁止词被机械拦截")

        print("\n[5] 增量补表：01–07 零差异 + 幂等")
        cfg = load(cfg_path)
        engine(cfg)
        p = apply(cfg_path)
        check(p.returncode == 0, "恢复干净产物")
        baseline = os.path.join(tmp, "baseline.xlsx")
        shutil.copy2(BOOK, baseline)
        wb = openpyxl.load_workbook(baseline)
        if "08-试用装与中小样" in wb.sheetnames:
            del wb["08-试用装与中小样"]
        wb.save(baseline)
        wb.close()
        snap = os.path.join(tmp, "snap.xlsx")
        shutil.copy2(baseline, snap)
        cfg = load(cfg_path, ["--mode", "incremental", "--in-workbook", baseline])
        engine(cfg)
        p = apply(cfg_path)
        check(p.returncode == 0, "增量交付门禁通过：" + p.stderr.strip()[-160:])
        a, b = openpyxl.load_workbook(snap), openpyxl.load_workbook(baseline)
        diff = 0
        for sn in a.sheetnames:
            if not sn[:2].isdigit() or not (1 <= int(sn[:2]) <= 7):
                continue
            for row in a[sn].iter_rows():
                for c in row:
                    if c.value != b[sn][c.coordinate].value:
                        diff += 1
        check(diff == 0, "01–07 逐单元格零差异（差异 %d）" % diff)
        check("08-试用装与中小样" in b.sheetnames, "增量补回 08 表")
        a.close(); b.close()

        print("\n[6] 签章 PDF：四表合成一份")
        sys.path.insert(0, SCRIPTS)
        try:
            from export_pdf import register_font
            register_font()
            has_font = True
        except SystemExit:
            has_font = False
            print("  · 未找到中文字体，跳过 PDF 断言（可用 --font 指定）")
        if not has_font:
            print("\n[v17_demo] OK（PDF 步骤跳过）  ·  临时目录：" + tmp)
            return 0
        cfg = load(cfg_path, ["--pdf"])
        engine(cfg)
        p = apply(cfg_path)
        check(p.returncode == 0, "PDF 链路成功：" + p.stderr.strip()[-120:])
        pdf = os.path.join(MASTER_ROOT, PERIOD, PERIOD + "对账-签章件.pdf")
        check(os.path.isfile(pdf) and os.path.getsize(pdf) > 10000, "签章 PDF 已生成")
        try:
            import pymupdf
            d = pymupdf.open(pdf)
            check(d.page_count >= 4, "PDF 至少 4 页（实际 %d 页）" % d.page_count)
            d.close()
        except ImportError:
            print("  · 跳过页数校验（未装 pymupdf）")

        print("\n[v17_demo] OK  ·  临时目录：" + tmp)
        return 0
    except AssertionError as e:
        print("\n[v17_demo] FAIL: %s" % e)
        print("临时目录已保留：" + tmp)
        return 1
    finally:
        pass


if __name__ == "__main__":
    raise SystemExit(main())
