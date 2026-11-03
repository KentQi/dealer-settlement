"""读取 _主档.xlsx，输出 recon_engine.py 需要的 RECON_CFG JSON。

用法：
    export RECON_CFG=$(python3 load_master.py /path/_主档.xlsx --period 202610 --asof 2026-11-02)
    python3 recon_engine.py

自动从 base_dir 下按关键词发现三份源文件：
    含「挂账」→ jia    含「销售」→ sales    含「库存」→ inventory
"""
import sys, os, json, glob, re, argparse, datetime
import openpyxl


def read_sheet(wb, name, key=None):
    if name not in wb.sheetnames:
        return []
    ws = wb[name]
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in r):
            continue
        rows.append(dict(zip(headers, r)))
    return rows


def month_bounds(period):
    y, m = int(period[:4]), int(period[4:6])
    start = datetime.date(y, m, 1)
    nxt = datetime.date(y + (m == 12), 1 if m == 12 else m + 1, 1)
    end = nxt - datetime.timedelta(days=1)
    return start, end, nxt


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("master")
    ap.add_argument("--period", required=True, help="YYYYMM，如 202610")
    ap.add_argument("--asof", help="库存/挂账导出日，默认次月 2 日")
    ap.add_argument("--pay", help="付款日，默认次月 5 日")
    ap.add_argument("--base", help="月份目录，默认主档同级的 {period} 或主档所在目录")
    ap.add_argument("--out", help="输出文件全路径，默认 {月份目录}/{period}铺货代销对账.xlsx")
    ap.add_argument("--strict", action="store_true", help="主档关键信息不全时直接退出")
    a = ap.parse_args()

    wb = openpyxl.load_workbook(a.master, data_only=True)
    parties = read_sheet(wb, "主体")
    contracts = read_sheet(wb, "合同")

    pa = next((p for p in parties if str(p.get("角色", "")).strip() == "甲方"), {})
    pb = next((p for p in parties if str(p.get("角色", "")).strip() == "乙方"), {})
    ct = contracts[0] if contracts else {}

    if not re.fullmatch(r"\d{6}", a.period or ""):
        sys.exit(f"[参数错误] --period 需为 6 位年月，如 202610；收到：{a.period!r}")
    if not 1 <= int(a.period[4:6]) <= 12:
        sys.exit(f"[参数错误] --period 的月份不合法：{a.period!r}")

    start, end, nxt = month_bounds(a.period)
    asof = a.asof or nxt.strftime("%Y-%m-%d")
    pay = a.pay or (nxt.replace(day=5)).strftime("%Y-%m-%d")
    ad = datetime.date.fromisoformat(asof)

    root = os.path.dirname(os.path.abspath(a.master))

    def has_src(d):
        try:
            return any("挂账" in f and f.endswith((".xls", ".xlsx")) and not f.startswith("~$")
                       for f in os.listdir(d))
        except OSError:
            return False

    if a.base:
        base = a.base
    else:
        subs = sorted((os.path.join(root, d) for d in os.listdir(root)
                       if os.path.isdir(os.path.join(root, d)) and d.isdigit()), reverse=True)
        cands = [os.path.join(root, a.period)] + subs
        base = next((d for d in cands if os.path.isdir(d) and has_src(d)), None)
        if base is None:
            if has_src(root):
                sys.exit(
                    "[中止] 客户根目录下有挂账文件，但没有可识别的月份子目录（如 202611/）。\n"
                    "       根目录混放旧导出会导致取错期间，请把三份源文件放进对应月份目录，\n"
                    "       或用 --base 明确指定本次使用的目录。")
            base = root
        elif base != subs[0] if subs else False:
            pass

    def pick(rules, default):
        """rules: [(必须含, 必须不含或None)]，按顺序取第一个命中。"""
        files = [os.path.basename(f) for f in glob.glob(os.path.join(base, "*.xls*"))
                 if not os.path.basename(f).startswith("~$")]
        for must, deny in rules:
            hits = [f for f in files if must in f and (deny is None or deny not in f)]
            if hits:
                return sorted(hits, key=len)[0]
        return default

    # ---- 预检：关键字段缺失时明确告警（写 stderr，不污染 stdout 的 JSON）----
    problems = []
    if not pa.get("名称"):
        problems.append("主体表缺少「甲方」行或名称为空")
    if not pb.get("名称"):
        problems.append("主体表缺少「乙方」行或名称为空")
    if not ct.get("合同编号"):
        problems.append("合同表缺少「合同编号」或表为空")
    if not (pa.get("开户行") and pa.get("账号")):
        problems.append("甲方开户行或账号为空（对账单付款栏将留白）")
    for _p in problems:
        print(f"[主档告警] {_p}", file=sys.stderr)
    if problems and a.strict:
        sys.exit("[中止] 主档关键信息不全，已按 --strict 退出。请先补全主档。")
    if problems:
        print(f"[提示] 共 {len(problems)} 项主档信息不全，仍继续生成；请核对输出文件抬头。", file=sys.stderr)

    cfg = {
        "base_dir": base,
        "master_path": os.path.abspath(a.master),
        "party_a": {"name": pa.get("名称", ""), "tax": pa.get("统一社会信用代码", ""),
                    "addr": pa.get("地址", ""), "contact": pa.get("联系人", ""),
                    "phone": pa.get("电话", "")},
        "party_b": {"name": pb.get("名称", ""), "tax": pb.get("统一社会信用代码", ""),
                    "addr": pb.get("地址", ""), "contact": pb.get("联系人", ""),
                    "phone": pb.get("电话", "")},
        "contract": {"no": ct.get("合同编号", "")},
        "bank": {"account_name": pa.get("户名") or pa.get("名称", ""),
                 "account_no": pa.get("账号", ""),
                 "bank_name": pa.get("开户行", "")},
        "period": {
            "start": start.strftime("%Y-%m-%d"),
            "end": end.strftime("%Y-%m-%d"),
            "period_cn": f"{start.year} 年 {start.month} 月",
            "asof": asof,
            "asof_cn": f"{ad.year} 年 {ad.month:02d} 月 {ad.day:02d} 日",
            "pay_cn": f"{datetime.date.fromisoformat(pay).year} 年 "
                      f"{datetime.date.fromisoformat(pay).month:02d} 月 "
                      f"{datetime.date.fromisoformat(pay).day:02d} 日",
            "next_month": ad.month,
            "out_name": f"{a.period}铺货代销对账.xlsx",
        },
        "sources": {
            "jia": pick([("挂账", None)], "甲方-挂账.xls"),
            "sales": pick([("销售数据", None), ("销售", "库存")], "乙方-销售数据.xlsx"),
            "inventory": pick([("库存", None)], "乙方-库存.xlsx"),
        },
    }
    if a.out:
        cfg["out_file"] = a.out
    print(json.dumps(cfg, ensure_ascii=False))


if __name__ == "__main__":
    main()
