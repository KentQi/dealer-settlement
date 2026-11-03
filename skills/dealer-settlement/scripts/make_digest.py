"""从 _主档.xlsx 生成《合同关键信息摘要》，供用户确认与存档。

输出：
  1) Markdown 文件（默认与主档同目录 _合同摘要.md）
  2) 控制台打印，便于直接阅读确认

用法：
    python3 make_digest.py _主档.xlsx [-o 输出.md]
"""
import sys, os, argparse, datetime
import openpyxl


def rows(wb, name):
    if name not in wb.sheetnames:
        return [], []
    ws = wb[name]
    headers = [c.value for c in ws[1]]
    data = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        data.append(dict(zip(headers, r)))
    return headers, data


def fmt_date(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime("%Y-%m-%d")
    return str(v) if v is not None else ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("master")
    ap.add_argument("-o", "--out")
    a = ap.parse_args()

    try:
        wb = openpyxl.load_workbook(a.master, data_only=True)
    except FileNotFoundError:
        sys.exit(f"[找不到主档] {a.master}\n请确认路径，或先用 make_master.py 生成主档模板。")
    except Exception as e:
        sys.exit(f"[无法读取主档] {a.master}\n原因：{e}\n主档须为 .xlsx（.xls 请先另存为 .xlsx）。")
    _, parties = rows(wb, "主体")
    _, contracts = rows(wb, "合同")
    _, prices = rows(wb, "价格")
    _, skus = rows(wb, "SKU")
    _, periods = rows(wb, "期间")

    pa = next((p for p in parties if str(p.get("角色", "")).strip() == "甲方"), {})
    pb = next((p for p in parties if str(p.get("角色", "")).strip() == "乙方"), {})
    ct = contracts[0] if contracts else {}

    today = datetime.date.today().strftime("%Y-%m-%d")
    L = []
    L.append(f"# 合同关键信息摘要\n")
    L.append(f"> 来源：`{os.path.basename(a.master)}`　生成日期：{today}　本摘要为对账作业的取数依据\n")

    L.append("## 一、缔约主体\n")
    L.append("| 项目 | 甲方（供货方） | 乙方（代销方） |")
    L.append("|---|---|---|")
    for key in ["名称", "统一社会信用代码", "地址", "联系人", "电话"]:
        L.append(f"| {key} | {pa.get(key) or '—'} | {pb.get(key) or '—'} |")
    L.append("")

    L.append("## 二、结算安排\n")
    L.append("| 项目 | 内容 |")
    L.append("|---|---|")
    for key in ["合同编号", "起始", "终止", "结算方式", "结算基数", "付款日", "异议期(工作日)", "签回期(工作日)", "价格附件", "备注"]:
        v = ct.get(key)
        if key in ("起始", "终止"):
            v = fmt_date(v)
        L.append(f"| {key} | {v if v not in (None, '') else '—'} |")
    if pa.get("开户行") or pa.get("账号"):
        L.append(f"| 收款账户 | {pa.get('开户行') or ''}　{pa.get('账号') or ''}　户名：{pa.get('户名') or pa.get('名称') or ''} |")
    L.append("")

    L.append("## 三、价格版本\n")
    L.append("| 条码 | 商品名称 | 供货价 | 生效起 | 生效止 | 零售价 |")
    L.append("|---|---|---|---|---|---|")
    for p in prices:
        L.append(f"| {p.get('商品条码','')} | {p.get('商品名称','')} | {p.get('供货价','')} | "
                 f"{fmt_date(p.get('生效起'))} | {fmt_date(p.get('生效止')) or '至今'} | {p.get('零售价','')} |")
    L.append("")

    settle = [s for s in skus if str(s.get("结算属性", "")).strip() == "结算类"]
    nonsettle = [s for s in skus if str(s.get("结算属性", "")).strip() == "非结算类"]
    L.append("## 四、SKU 概况\n")
    L.append(f"- 结算类 {len(settle)} 个，非结算类 {len(nonsettle)} 个，合计 {len(skus)} 个。")
    if nonsettle:
        names = "、".join(str(s.get("标准品名", "")) for s in nonsettle)
        L.append(f"- 非结算类（陈列/试用，不参与结算）：{names}。")
    L.append("")

    if periods:
        L.append("## 五、已出账期间\n")
        L.append("| 期间 | 基准日 | 期初金额 | 结算金额 | 差异金额 | 状态 |")
        L.append("|---|---|---|---|---|---|")
        for p in periods:
            L.append(f"| {p.get('期间','')} | {fmt_date(p.get('数据基准日'))} | {p.get('期初金额','')} | "
                     f"{p.get('结算金额','')} | {p.get('差异金额','')} | {p.get('状态','')} |")
        L.append("")

    gaps = []
    if not pa.get("电话"):
        gaps.append("甲方联系电话")
    if not pa.get("开户行") or not pa.get("账号"):
        gaps.append("甲方收款账户")
    if not ct.get("价格附件"):
        gaps.append("价格附件编号")
    L.append("## 六、待补项\n")
    L.append("- " + ("；".join(gaps) if gaps else "无，主档要素齐全。"))
    L.append("")
    L.append("---")
    L.append("确认要点：① 主体名称与税号是否与营业执照一致；② 付款日、异议期、签回期是否与合同原文一致；"
             "③ 价格版本是否为本期适用；④ 非结算类清单是否完整。")

    md = "\n".join(L)
    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.master)), "_合同摘要.md")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(md + "\n")
    print(md)
    print(f"\n[已写入] {out}")


if __name__ == "__main__":
    main()
