"""读取 _主档.xlsx，输出 recon_engine.py 需要的 RECON_CFG JSON。

用法：
    export RECON_CFG=$(python3 load_master.py /path/_主档.xlsx --period 202610 --asof 2026-11-02)
    python3 recon_engine.py

从主档同级的 {period} 月份目录（找不到时取最近一个含挂账文件的数字目录）自动发现
三份源文件：
    文件名含「挂账」→ jia    含「销售数据」（或含「销售」且不含「库存」）→ sales
    文件名含「库存」→ inventory

约定：
- stdout 只输出一行 JSON 配置；所有告警/错误走 stderr（失败时 stdout 保持 0 字节）。
- 合同条款（结算方式/结算基数/异议期/签回期/付款日）从主档「合同」表读取并传给引擎；
  读不到时引擎会显式标注「以合同为准」，不会静默套用固定文案。
- 客户根目录存在 _数据映射.json（结构漂移经用户确认后的映射）时，读入并放进
  RECON_CFG["mapping"] 交给引擎按映射取数；不存在时行为与旧版完全一致，映射损坏则中止。
"""
import argparse
import datetime
import json
import os
import re
import sys

import openpyxl

CFG_SCHEMA_HINT = "RECON_CFG"


def read_sheet(wb, name):
    """读取工作表为 [{表头: 值}]；空行跳过。"""
    if name not in wb.sheetnames:
        return []
    ws = wb[name]
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        rows.append(dict(zip(headers, r)))
    return rows


def month_bounds(period):
    y, m = int(period[:4]), int(period[4:6])
    start = datetime.date(y, m, 1)
    nxt = datetime.date(y + (m == 12), 1 if m == 12 else m + 1, 1)
    end = nxt - datetime.timedelta(days=1)
    return start, end, nxt


def _date_text(value):
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.strftime("%Y-%m-%d")
    return str(value).strip() if value is not None else ""


def _as_int(value):
    try:
        f = float(str(value).strip())
    except (TypeError, ValueError):
        return None
    return int(f) if f == int(f) else None


def _pay_day_from_text(text, default=5):
    """「次月5日前」→ 5；解析不到用默认 5 日。"""
    m = re.search(r"([0-9]{1,2})\s*日", str(text or ""))
    if m:
        day = int(m.group(1))
        if 1 <= day <= 28:
            return day
    return default


def _excel_files(directory):
    try:
        return sorted(f for f in os.listdir(directory)
                      if f.lower().endswith((".xls", ".xlsx", ".xlsm"))
                      and not f.startswith("~$"))
    except OSError:
        return []


def _fail(msg):
    sys.exit(msg)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("master")
    ap.add_argument("--period", required=True, help="YYYYMM，如 202610")
    ap.add_argument("--asof", help="库存/挂账导出日，默认次月 2 日")
    ap.add_argument("--pay", help="付款日，默认取主档「合同·付款日」中的日期")
    ap.add_argument("--base", help="月份目录，默认主档同级的 {period} 或主档所在目录")
    ap.add_argument("--out", help="输出文件全路径，默认 {月份目录}/{period}铺货代销对账.xlsx")
    ap.add_argument("--strict", action="store_true", help="主档关键信息不全时直接退出")
    a = ap.parse_args()

    # ---- 参数校验放在读文件之前，错误一律走 stderr、stdout 0 字节 ----
    if not re.fullmatch(r"[0-9]{6}", a.period or ""):
        _fail(f"[参数错误] --period 需为 6 位年月，如 202610；收到：{a.period!r}")
    year, month = int(a.period[:4]), int(a.period[4:6])
    if not 1 <= month <= 12:
        _fail(f"[参数错误] --period 的月份不合法：{a.period!r}")
    if not 1900 <= year <= 2999:
        _fail(f"[参数错误] --period 的年份不合法（应为 1900–2999）：{a.period!r}")

    start, end, nxt = month_bounds(a.period)
    asof = a.asof or nxt.strftime("%Y-%m-%d")
    try:
        ad = datetime.date.fromisoformat(asof)
    except ValueError:
        _fail(f"[参数错误] --asof 需为 YYYY-MM-DD（如 2026-11-02）；收到：{asof!r}")
    if a.pay:
        try:
            datetime.date.fromisoformat(a.pay)
        except ValueError:
            _fail(f"[参数错误] --pay 需为 YYYY-MM-DD（如 2026-11-05）；收到：{a.pay!r}")

    # ---- 读主档 ----
    try:
        wb = openpyxl.load_workbook(a.master, data_only=True)
    except FileNotFoundError:
        _fail(f"[找不到主档] {a.master}\n请确认路径，或先用 make_master.py 生成主档模板。")
    except Exception as e:
        _fail(f"[无法读取主档] {a.master}\n原因：{e}\n主档须为 .xlsx（.xls 请先另存为 .xlsx）。")
    try:
        parties = read_sheet(wb, "主体")
        contracts = read_sheet(wb, "合同")
    finally:
        wb.close()

    pa = next((p for p in parties if str(p.get("角色", "")).strip() == "甲方"), {})
    pb = next((p for p in parties if str(p.get("角色", "")).strip() == "乙方"), {})
    ct = contracts[0] if contracts else {}

    # ---- 合同条款（缺项交给引擎显式标注，不在这里补历史演示值）----
    contract_cfg = {
        "no": str(ct.get("合同编号") or "").strip(),
        "start": _date_text(ct.get("起始")),
        "end": _date_text(ct.get("终止")),
        "method": str(ct.get("结算方式") or "").strip(),
        "basis": str(ct.get("结算基数") or "").strip(),
        "pay_day_text": str(ct.get("付款日") or "").strip(),
        "objection_days": _as_int(ct.get("异议期(工作日)")),
        "signback_days": _as_int(ct.get("签回期(工作日)")),
        "price_annex": str(ct.get("价格附件") or "").strip(),
    }
    pay_day = _pay_day_from_text(contract_cfg["pay_day_text"])
    pay = a.pay or (nxt.replace(day=pay_day)).strftime("%Y-%m-%d")

    root = os.path.dirname(os.path.abspath(a.master))

    def has_src(d):
        return any("挂账" in f for f in _excel_files(d))

    if a.base:
        base = a.base
    else:
        try:
            subs = sorted((os.path.join(root, d) for d in os.listdir(root)
                           if os.path.isdir(os.path.join(root, d)) and d.isdigit()), reverse=True)
        except OSError:
            subs = []
        cands = [os.path.join(root, a.period)] + subs
        base = next((d for d in cands if os.path.isdir(d) and has_src(d)), None)
        if base is None:
            if has_src(root):
                _fail(
                    "[中止] 客户根目录下有挂账文件，但没有可识别的月份子目录（如 202611/）。\n"
                    "       根目录混放旧导出会导致取错期间，请把三份源文件放进对应月份目录，\n"
                    "       或用 --base 明确指定本次使用的目录。")
            base = root
    base = os.path.abspath(base)
    if not os.path.isdir(base):
        _fail(f"[目录不存在] {base}\n请检查 --base 参数或月份目录是否已创建。")

    def pick(rules, default):
        """rules: [(必须含, 必须不含或 None)]，按顺序取第一个命中；多个候选取文件名最短者。"""
        files = _excel_files(base)
        for must, deny in rules:
            hits = [f for f in files if must in f and (deny is None or deny not in f)]
            if hits:
                return sorted(hits, key=lambda x: (len(x), x))[0]
        return default

    # ---- 预检：关键字段缺失时明确告警（stderr，不污染 stdout 的 JSON）----
    problems = []
    if not pa.get("名称"):
        problems.append("主体表缺少「甲方」行或名称为空")
    if not pb.get("名称"):
        problems.append("主体表缺少「乙方」行或名称为空")
    if not contract_cfg["no"]:
        problems.append("合同表缺少「合同编号」或表为空")
    if not (pa.get("开户行") and pa.get("账号")):
        problems.append("甲方开户行或账号为空（对账单付款栏将留白）")
    for _p in problems:
        print(f"[主档告警] {_p}", file=sys.stderr)
    if problems and a.strict:
        _fail("[中止] 主档关键信息不全，已按 --strict 退出。请先补全主档。")
    if problems:
        print(f"[提示] 共 {len(problems)} 项主档信息不全，仍继续生成；请核对输出文件抬头。", file=sys.stderr)

    # ---- 结构映射（可选）：客户根目录存在 _数据映射.json 时随 CFG 下传引擎 ----
    # 不存在时行为与旧版完全一致；文件损坏时中文报错并 exit 1（stdout 保持 0 字节）。
    mapping = None
    mapping_file = os.path.join(root, "_数据映射.json")
    if os.path.exists(mapping_file):
        try:
            with open(mapping_file, "r", encoding="utf-8-sig") as fh:
                mapping = json.load(fh)
        except (OSError, ValueError) as e:
            _fail(f"[映射文件损坏] {mapping_file}\n原因：{e}\n"
                  f"请修复该 JSON，或删除后按旧流程运行；未生成任何文件。")
        if not isinstance(mapping, dict):
            _fail(f"[映射文件损坏] {mapping_file}\n"
                  f"顶层必须是 JSON 对象（schema 应为 dealer-settlement/data-mapping@1）。")
        if mapping.get("schema") != "dealer-settlement/data-mapping@1":
            _fail(f"[映射 schema 不支持] {mapping_file}\n"
                  f"当前 {mapping.get('schema')!r}；要求 'dealer-settlement/data-mapping@1'。\n"
                  f"请用 apply_mapping.py 重新生成/确认映射后重试。")

    cfg = {
        "base_dir": base,
        "master_path": os.path.abspath(a.master),
        "party_a": {"name": pa.get("名称", ""), "tax": pa.get("统一社会信用代码", ""),
                    "addr": pa.get("地址", ""), "contact": pa.get("联系人", ""),
                    "phone": pa.get("电话", "")},
        "party_b": {"name": pb.get("名称", ""), "tax": pb.get("统一社会信用代码", ""),
                    "addr": pb.get("地址", ""), "contact": pb.get("联系人", ""),
                    "phone": pb.get("电话", "")},
        "contract": contract_cfg,
        "bank": {"account_name": pa.get("户名") or pa.get("名称", ""),
                 "account_no": pa.get("账号", ""),
                 "bank_name": pa.get("开户行", "")},
        "period": {
            "start": start.strftime("%Y-%m-%d"),
            "end": end.strftime("%Y-%m-%d"),
            "period_cn": f"{start.year} 年 {start.month} 月",
            "asof": ad.isoformat(),
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
    if mapping is not None:
        cfg["mapping"] = mapping
    if a.out:
        cfg["out_file"] = os.path.abspath(a.out)
    print(json.dumps(cfg, ensure_ascii=False))


if __name__ == "__main__":
    main()
