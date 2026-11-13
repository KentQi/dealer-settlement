#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""业务口径确认门禁（阶段 G）。

本模块只做三件事：定义问卡、校验确认记录、计算口径哈希。
它不生成任何对外单据，也不参与数字计算。

设计要点：
- 口径不写死在 SKILL.md 里，必须由用户逐项确认后落在 _口径确认.json；
- 任何必填项缺失即阻断（exit 1、stdout 0 字节、stderr 给业务语言问卡）；
- internal_only 与 disclosure_whitelist 是两条互斥通道；
- 口径哈希进复核锁，改口径必须重走门禁。
"""
import argparse
import hashlib
import json
import os
import sys

SCHEMA = "dealer-settlement/koujing@1"
FILENAME = "_口径确认.json"
SCENES = ("首次建档", "每月出账", "增量补表", "换源表结构")

# 业务语言问卡：面向人，不出现字段名、脚本名、schema 术语。
CARDS = {
    "首次建档": [
        "以后每个月算客户还欠多少货、该结多少钱，以哪张表说了算？哪张表里的数量是准的？",
        "有没有哪张表只是给我们自己看的，里面的数量不能拿来当结存？",
        "试用装、中小样、赠品要不要单独做一张表让客户签字确认数量？它们算不算钱？",
        "哪些话是我们自己内部定的规矩，绝对不能印到给客户的单子上？",
        "效期提醒这期要不要做？按还剩几个月提醒？从哪天算起？",
        "单子上第 08、09 张表分别放什么？哪几张要标红打印、导出成一份 PDF 给客户盖章？",
        "挂账表上只有物料号、没有条码时，有没有一张对照表能对上？",
    ],
    "每月出账": [
        "这个月按哪天截止？上期剩下来的数转过来是多少？",
        "这个月发出去多少、客户实际用了多少？正装和赠品分别说。",
        "赠品使用量这次按什么数出？还按上次那个规矩吗？",
        "效期这期开不开？分档和起算日跟上次一样吗？",
        "有没有新表、新列、新物料号，或者换过导出系统？",
        "对客户的说法跟上期一样吗？有没有哪句话这次不能写？",
    ],
    "增量补表": [
        "这份工作簿是哪一期？哪几张表一个字都不能动？",
        "要补哪几张表、编号排第几？补完一共几张？",
        "补的表里每个数字从哪张源表、哪一列来？跟前面几张表要对得上吗？",
        "补的这张表要不要客户签字确认？哪些表要一起导成 PDF？",
        "上次定过的口径这次还作数吗？有哪条要改（比如效期这次做不做）？",
        "有没有只能我们内部知道、不能印到补的表上的规矩？",
    ],
    "换源表结构": [
        "这张新表是谁发的、管什么的？跟老表是不是一回事？",
        "哪一列是数量、哪一列是编码或条码？跟老表对不上时以谁为准？",
        "原来定的哪张表说了算，还成立吗？要不要换？",
        "条码和物料号的对照还全吗？有没有新商品没登记？",
        "结构变了之后，之前定过的口径有没有哪条要跟着改？",
    ],
}

# 每个场景的必填项：(decisions 里的键, 未答时回给用户的业务语言追问)
REQUIRED = {
    "首次建档": [
        ("settlement_basis", "这批货哪张表说了算？请指出表名和那一列。"),
        ("category_source", "试用装、中小样各发了多少？正装数量和挂账表对得上吗？"),
        ("promo", "这些赠品算不算钱？要不要单独做一张表让客户签字？"),
        ("usage", "客户这次用了多少试用装和中小样？请给数，规矩我们不印。"),
        ("expiry", "效期这期做不做？按几个月提醒、从哪天算起？"),
        ("sheet_map", "第 08、09 张表分别放什么？"),
        ("internal_only", "哪些规矩只能内部用、不能印给客户？"),
    ],
    "每月出账": [
        ("settlement_basis", "这批货哪张表说了算？跟上期一样吗？"),
        ("category_source", "本期发出去多少正装、试用装、中小样？"),
        ("rolling", "上期期末转过来是多少？本期按四栏台账结转吗？"),
        ("usage", "客户这次用了多少试用装和中小样？请给数。"),
        ("expiry", "效期这期开不开？分档和起算日跟上次一样吗？"),
        ("sheet_map", "第 08、09 张表放什么、这期出哪几张？"),
        ("internal_only", "哪些规矩只能内部用、不能印给客户？"),
    ],
    "增量补表": [
        ("settlement_basis", "补的表里每个数字从哪张源表哪一列来？"),
        ("delta", "哪几张表一个字都不能动？补的表排第几、补完一共几张？"),
        ("usage", "客户这次用了多少试用装和中小样？请给数。"),
        ("expiry", "效期这次做不做？"),
        ("sheet_map", "补完之后的附表编号怎么排？"),
        ("internal_only", "有没有只能内部知道、不能印到补的表上的规矩？"),
    ],
    "换源表结构": [
        ("settlement_basis", "换表之后哪张表说了算？以谁为准？"),
        ("category_source", "新表里正装和赠品的数量还从同一张表取吗？"),
        ("barcode_bridge", "条码和物料号的对照还全吗？"),
        ("sheet_map", "表号要不要跟着改？"),
        ("internal_only", "有没有只能内部知道、不能印给客户的规矩？"),
    ],
}

DEFAULT_DECISIONS = {
    "settlement_basis": {
        "source": "甲方挂账表", "field": "未结算数量",
        "excluded": ["效期参考表的结存数量"],
        "rule": "正装结存与结算唯一依据",
    },
    "category_source": {
        "source": "商品分类", "settle": True,
        "正装": 0, "试用装": 0, "中小样": 0,
    },
    "promo": {"settle": False, "separate_sheet": True, "confirm_by_party_b": True},
    "rolling": {"opening_source": "上期甲方期末", "carry_forward": True,
                "first_period_opening": 0},
    "usage": {
        "values": {},
        "source": "甲方口径",
        "private": True,
        "note": "ledger 每行是一门已确认的既成数字；closing 由脚本按 opening+shipped-used 计算，不写规则",
        "ledger": [],
    },
    "expiry": {"enabled": False, "unit": "month", "tiers": [32, 28, 24],
               "base": "asof", "fifo": "display_only",
               "sample_qty_basis": "台账期末", "next_period": True},
    "barcode_bridge": {"required": True},
    "sheet_map": {"08": "试用装与中小样", "09": "效期预警"},
    "print": {"red_sheets": ["00", "01", "02", "08"],
              "pdf_sheets": ["00", "01", "02", "08"]},
    "delta": {"mode": "full", "frozen_sheets": [], "note_only_sheets": [],
              "rebuild_sheets": []},
}


def blank(scene="每月出账", period="", asof="", customer=""):
    """生成一份待填写的确认记录（不含 confirmed_by，故必然被门禁阻断）。"""
    import copy
    import datetime
    return {
        "schema": SCHEMA,
        "customer": customer,
        "period": period,
        "asof": asof,
        "scene": scene,
        "confirmed_by": "",
        "confirmed_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "decisions": copy.deepcopy(DEFAULT_DECISIONS),
        "internal_only": [],
        "disclosure_whitelist": [],
        "notes": [],
    }


def digest(cfg):
    """口径哈希：排除哈希字段自身与留痕，只覆盖真正的口径内容。"""
    d = {k: v for k, v in cfg.items()
         if k not in ("config_sha256", "prev_config_sha256", "notes", "confirmed_at")}
    raw = json.dumps(d, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def validate(cfg, scene=None, require_hash=True):
    """返回 (错误列表, 应重发的问卡列表)。空列表表示通过。"""
    errors, questions = [], []
    if not isinstance(cfg, dict):
        return ["确认记录不是 JSON 对象"], list(CARDS.get(scene or "每月出账", []))
    if cfg.get("schema") != SCHEMA:
        errors.append("确认记录版本不对，应为 " + SCHEMA)
    sc = scene or cfg.get("scene") or "每月出账"
    if sc not in SCENES:
        errors.append("场景取值不在四种之内：" + str(sc))
        sc = "每月出账"
    if not cfg.get("confirmed_by"):
        errors.append("还没有确认人")
        questions.append("这份口径由谁确认？请留一个名字或代号。")
    if not cfg.get("confirmed_at"):
        errors.append("还没有确认时间")
    if not cfg.get("period"):
        errors.append("还没有写对账期间")
    if not cfg.get("asof"):
        errors.append("还没有写数据基准日")

    dec = cfg.get("decisions") or {}
    for key, ask in REQUIRED.get(sc, []):
        val = cfg.get(key) if key == "internal_only" else dec.get(key)
        if val is None or val == {} or val == "" or val == []:
            errors.append("口径未确认：" + key)
            questions.append(ask)
    usage = dec.get("usage") or {}
    if "usage" in dec and not usage.get("values"):
        errors.append("使用量还没有给数")
        questions.append("客户这次用了多少试用装和中小样？请给数，规矩我们不印。")
    if usage.get("values") and not usage.get("ledger"):
        errors.append("还没有逐项数量台账")
        questions.append("请逐项给出每种试用装、中小样的期初、本期发出与客户使用数量。")
    for row in usage.get("ledger") or []:
        if not isinstance(row, dict) or not row.get("name") or not row.get("kind"):
            errors.append("台账有行的品名或类型为空")
            break
        for key in ("opening", "shipped", "used"):
            if not isinstance(row.get(key), (int, float)):
                errors.append("台账「%s」的 %s 不是数字" % (row.get("name"), key))
                break
    if "expiry" in dec:
        ex = dec["expiry"] or {}
        if not isinstance(ex.get("tiers"), list) or not ex.get("tiers"):
            errors.append("效期分档没填")
            questions.append("效期按还剩几个月提醒？")
        if ex.get("sample_qty_basis") not in ("台账期末", "期间发货"):
            errors.append("效期表里中小样件数按哪个口径没确认")
            questions.append("效期表里的中小样件数，按台账期末还是按本期发货量？")

    internal = cfg.get("internal_only") or []
    if not internal:
        errors.append("还没有列出只能内部使用、不能印给客户的内容")
        questions.append("哪些规矩只能我们内部用、不能印到给客户的单子上？")
    white = cfg.get("disclosure_whitelist") or []
    overlap = set(map(str, internal)) & set(map(str, white))
    if overlap:
        errors.append("内部内容与对外白名单重叠：" + "、".join(sorted(overlap))[:120])
    sm = dec.get("sheet_map") or {}
    if sm.get("08") != "试用装与中小样" or sm.get("09") != "效期预警":
        errors.append("附表编号与定稿不一致：08 应为试用装与中小样、09 应为效期预警")
        questions.append("第 08、09 张表分别放什么？")
    if require_hash and cfg.get("config_sha256") and cfg["config_sha256"] != digest(cfg):
        errors.append("口径哈希不匹配：记录被改过，请重发问卡重新确认")
    return errors, questions


def load(path, scene=None):
    """读确认记录；失败时打印中文原因与业务语言问卡，并以 exit 1 结束。"""
    def _fail(msgs, asks):
        sys.stderr.write("[口径未确认] 本次不生成任何文件。\n")
        for m in msgs:
            sys.stderr.write("  · " + m + "\n")
        if asks:
            sys.stderr.write("\n请先逐项确认（用业务语言回答即可）：\n")
            for i, q in enumerate(dict.fromkeys(asks), 1):
                sys.stderr.write("  %d. %s\n" % (i, q))
        sys.stderr.write("\n确认后写入 %s 再重跑。\n" % FILENAME)
        raise SystemExit(1)

    if not path or not os.path.exists(path):
        _fail(["找不到 " + FILENAME], list(CARDS.get(scene or "每月出账", [])))
    try:
        with open(path, "r", encoding="utf-8") as fh:
            cfg = json.load(fh)
    except Exception as e:
        _fail([FILENAME + " 不是合法 JSON：" + str(e)], [])
    errs, asks = validate(cfg, scene=scene)
    if errs:
        _fail(errs, asks)
    cfg["config_sha256"] = digest(cfg)
    return cfg


def main():
    ap = argparse.ArgumentParser(description="口径确认门禁")
    ap.add_argument("--base", required=True, help="客户根目录")
    ap.add_argument("--scene", default="每月出账", choices=list(SCENES))
    ap.add_argument("--check", action="store_true", help="只校验，不打印哈希")
    ap.add_argument("--template", action="store_true", help="生成待填模板")
    ap.add_argument("--period", default="")
    ap.add_argument("--asof", default="")
    a = ap.parse_args()
    path = os.path.join(a.base, FILENAME)
    if a.template:
        if os.path.exists(path):
            sys.stderr.write("[中止] %s 已存在，不覆盖。\n" % path)
            return 2
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(blank(a.scene, a.period, a.asof), fh, ensure_ascii=False, indent=2)
        sys.stderr.write("已生成待填模板：%s\n" % path)
        return 0
    cfg = load(path, a.scene)
    print(cfg["config_sha256"] if not a.check else "OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
