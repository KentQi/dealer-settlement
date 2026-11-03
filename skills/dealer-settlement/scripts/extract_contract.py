"""从合同文件抽取主档所需字段，输出候选 JSON（不写库）。

支持 .docx / .pdf / .xlsx / .xlsm / .txt / .md。仅做「候选提取 + 证据留痕」，最终取值由人确认。

用法：
    python3 extract_contract.py <合同文件> [-o 输出.json]
"""
import sys, os, re, json, zipfile, argparse
from xml.etree import ElementTree as ET

NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def text_of_docx(path):
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml").decode("utf-8", "ignore")
    except (zipfile.BadZipFile, KeyError) as e:
        raise SystemExit(f"[无法读取] {path}\n不是有效的 .docx（缺少 word/document.xml 或已损坏）：{e}\n"
                         f"请用 Word 另存为 .docx 后重试。")
    root = ET.fromstring(xml)
    lines = []
    for p in root.iter(NS + "p"):
        buf = "".join(t.text or "" for t in p.iter(NS + "t"))
        if buf.strip():
            lines.append(buf.strip())
    # 表格内的文字也并入
    for tbl in root.iter(NS + "tbl"):
        for tr in tbl.iter(NS + "tr"):
            cells = ["".join(t.text or "" for t in tc.iter(NS + "t")).strip() for tc in tr.iter(NS + "tc")]
            cells = [c for c in cells if c]
            if cells:
                lines.append(" | ".join(cells))
    return lines


PDF_PAGES = None


def text_of_pdf(path):
    import pdfplumber
    global PDF_PAGES
    lines = []
    with pdfplumber.open(path) as pdf:
        PDF_PAGES = len(pdf.pages)
        for page in pdf.pages:
            for ln in (page.extract_text() or "").split("\n"):
                if ln.strip():
                    lines.append(ln.strip())
    return lines


def text_of_xlsx(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        lines = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                cells = [str(v).strip() for v in row if v is not None and str(v).strip()]
                if cells:
                    lines.append(" | ".join(cells))
        return lines
    finally:
        wb.close()


def load_lines(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return text_of_docx(path)
    if ext == ".pdf":
        return text_of_pdf(path)
    if ext in (".xlsx", ".xlsm"):
        return text_of_xlsx(path)
    if ext in (".txt", ".md"):
        for enc in ("utf-8-sig", "gb18030"):
            try:
                with open(path, encoding=enc) as fh:
                    return [line.strip() for line in fh if line.strip()]
            except UnicodeDecodeError:
                continue
        raise SystemExit(f"[编码无法识别] {path}\n既不是 UTF-8 也不是 GB18030，请转码后重试。")
    raise SystemExit(f"不支持的格式 {ext}；.doc/.wps 请先另存为 .docx")


def find(lines, pattern, flags=0):
    """返回 (值, 证据行, 行号)。"""
    rx = re.compile(pattern, flags)
    for i, ln in enumerate(lines):
        m = rx.search(ln)
        if m:
            return (m.group(1).strip() if m.groups() else m.group(0).strip()), ln, i
    return None, None, -1


def block_of(joined, start_kw, end_kws, fallback_len=800):
    a = joined.find(start_kw)
    if a < 0:
        return ""
    b = len(joined)
    for kw in end_kws:
        p = joined.find(kw, a + len(start_kw))
        if p > 0:
            b = min(b, p)
    return joined[a:min(b, a + fallback_len)]


def extract(path):
    global PDF_PAGES
    PDF_PAGES = None          # 避免上一次解析 PDF 的页数污染本次 needs_ocr 判断
    if not os.path.exists(path):
        raise SystemExit(f"[找不到合同文件] {path}\n请确认路径（支持 .docx/.pdf/.xlsx/.xlsm/.txt/.md）。")
    lines = load_lines(path)
    joined = "\n".join(lines)
    chars = len(joined.strip())
    if len(lines) < 15 or chars < 300:
        ext = os.path.splitext(path)[1].lower()
        if ext in (".pdf", ".docx"):
            msg = (f"\n[很可能不是文本版] 仅解析到 {len(lines)} 行 / {chars} 字。\n"
                   f"多为扫描件或图片型 PDF，没有文本层，正则无法提取。\n"
                   f"建议：① 取得可复制的电子版合同；② 或对 PDF 做 OCR 后再跑；\n"
                   f"      ③ 或直接按确认页手工填写主档。\n")
        else:
            msg = (f"\n[提示] 仅解析到 {len(lines)} 行 / {chars} 字，内容可能不完整；"
                   f"请核对文本格式（表格型合同建议提供 .docx/.xlsx 原件）。\n")
        print(msg, file=sys.stderr)
    out = []

    def add(field, value, evidence, conf="高", target=""):
        out.append({"field": field, "value": value, "evidence": evidence,
                    "confidence": conf if value else "待核", "target": target})

    # 合同编号
    v, ev, _ = find(lines, r"(?:主合同编号|合同编号)[：:\s]*([A-Za-z0-9\-－—_]{4,})")
    if not v:
        v, ev, _ = find(lines, r"(JL-[A-Z]+-\d{4}-\d+)")
    add("合同编号", v, ev, target="合同.合同编号")

    # 双方名称
    v, ev, _ = find(lines, r"(?:供货商|供货方|甲方)[：:\s]*([^\s，,\n]{2,40}?(?:有限公司|股份有限公司|合伙企业|个体工商户))")
    add("甲方名称", v, ev, target="主体.甲方.名称")
    v2, ev2, _ = find(lines, r"(?:经销商|代销方|乙方)[：:\s]*([^\s，,\n]{2,40}?(?:有限公司|股份有限公司|合伙企业|个体工商户))")
    if not v2:
        v2, ev2, _ = find(lines, r"(经销商[：:\s]*[^\n]{2,40})")
    add("乙方名称", v2, ev2, target="主体.乙方.名称")

    # 按甲方/乙方分段取税号
    pa_block = block_of(joined, "以下简称甲方", ["以下简称乙方", "一、"])
    pb_block = block_of(joined, "以下简称乙方", ["一、", "二、"])
    if not pa_block:
        pa_block = joined[:len(joined) // 2]
        pb_block = pb_block or joined[len(joined) // 2:]

    v, ev, _ = find(pa_block.split("\n"), r"统一社会信用代码[：:\s]*([0-9A-Z]{18})")
    add("甲方税号", v, ev, target="主体.甲方.统一社会信用代码")
    v, ev, _ = find(pb_block.split("\n"), r"统一社会信用代码[：:\s]*([0-9A-Z]{18})")
    add("乙方税号", v, ev, target="主体.乙方.统一社会信用代码")

    for party, blk in (("甲方", pa_block), ("乙方", pb_block)):
        ls = blk.split("\n")
        v, ev, _ = find(ls, r"联系人[：:\s]*([^\s，,；;]{2,12})")
        add(f"{party}联系人", v, ev, target=f"主体.{party}.联系人")
        v, ev, _ = find(ls, r"(?:通讯地址|地址)[：:\s]*([^\n]{4,60})")
        add(f"{party}地址", v, ev, target=f"主体.{party}.地址")
        v, ev, _ = find(ls, r"(?:联系电话|电话|手机)[：:\s]*([0-9\-—]{7,20})")
        if not v and party == "乙方":
            # 乙方电话通常写在「往来对账」条款的对账联系人处，按上下文窗口取，避免误抓甲方
            j = joined.find("对账联系人")
            if j >= 0:
                v, ev, _ = find(joined[j:j + 300].split("\n"),
                                 r"联系电话[：:\s]*([0-9\-—]{7,20})")
        add(f"{party}电话", v, ev, target=f"主体.{party}.电话")

    # 期限
    v, ev, _ = find(lines, r"合同期限[：:\s]*自?\s*(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)\s*至\s*(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)")
    if v:
        m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日\s*至\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", ev)
        if m:
            g = m.groups()
            add("合同起始", f"{g[0]}-{int(g[1]):02d}-{int(g[2]):02d}", ev, target="合同.起始")
            add("合同终止", f"{g[3]}-{int(g[4]):02d}-{int(g[5]):02d}", ev, target="合同.终止")
    else:
        add("合同起始", None, None, target="合同.起始")
        add("合同终止", None, None, target="合同.终止")

    # 结算方式（原文留证，取规范值）
    v, ev, _ = find(lines, r"结算方式[：:\s]*([^\n。；]{4,30})")
    canon = None
    if v or ev:
        full = (ev or "") + (v or "")
        if "铺货代销" in joined:
            canon = "铺货代销"
        elif "买断" in full:
            canon = "买断"
        elif "代销" in full:
            canon = "代销"
    add("结算方式", canon, ev or v, conf="高" if canon else "待核", target="合同.结算方式")

    # 结算基数
    v, ev, _ = find(lines, r"结算基数[：:\s]*([^\n。]{6,80})")
    if not v:
        v, ev, _ = find(lines, r"(已实际销售[^\n。]{4,60})")
    canon = None
    if v or ev:
        t = (v or "") + (ev or "")
        if "实际动销" in t or "已实际销售" in t:
            canon = "实际动销"
        elif "签收" in t:
            canon = "签收"
        elif "发货" in t:
            canon = "发货"
    add("结算基数", canon, v or ev, conf="高" if canon else "待核", target="合同.结算基数")

    # 账期
    v, ev, _ = find(lines, r"(?:结款期限|付款期限)[^\n]{0,40}?每月\s*(\d{1,2})\s*日前")
    add("付款日", f"次月{v}日前" if v else None, ev, target="合同.付款日")

    # 异议期 / 签回期
    v, ev, _ = find(lines, r"收到(?:甲方)?对账单后\s*(\d{1,2})\s*个工作日内[^\n]{0,20}(?:签字|签章|邮寄|寄回)")
    add("签回期(工作日)", int(v) if v and v.isdigit() else None, ev, target="合同.签回期(工作日)")
    v, ev, _ = find(lines, r"异议[^\n]{0,30}?(\d{1,2})\s*个工作日内")
    add("异议期(工作日)", int(v) if v and v.isdigit() else None, ev, target="合同.异议期(工作日)")

    # 收款账户
    v, ev, _ = find(lines, r"户\s*名[：:\s]*([^\n]{2,40})")
    add("账户户名", v, ev, target="主体.甲方.户名")
    v, ev, _ = find(lines, r"账\s*号[：:\s]*([0-9]{6,32})")
    add("银行账号", v, ev, target="主体.甲方.账号")
    v, ev, _ = find(lines, r"开户行[：:\s]*([^\n]{4,50})")
    add("开户行", v, ev, target="主体.甲方.开户行")

    # 价格附件
    v, ev, _ = find(lines, r"(?:供货价格|价格)[^\n]{0,20}?(?:详见|见)[^\n。]{2,40}")
    add("价格依据", v, ev, conf="待核", target="合同.价格附件")

    needs_ocr = bool(os.path.splitext(path)[1].lower() == ".pdf" and PDF_PAGES and len(lines) < 15)
    missing = [f["field"] for f in out if not f["value"]]
    if len(missing) >= len(out) * 0.6:
        print(f"[提示] {len(missing)}/{len(out)} 个字段未识别，可能合同模板差异较大，"
              f"建议人工核对后手工补录。", file=sys.stderr)
    return {"file": os.path.abspath(path), "line_count": len(lines),
            "pages": PDF_PAGES, "needs_ocr": needs_ocr,
            "fields": out, "unresolved": missing,
            "note": "本文件仅为候选值，须逐项人工确认后方可写入主档"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("contract")
    ap.add_argument("-o", "--out")
    a = ap.parse_args()
    res = extract(a.contract)
    js = json.dumps(res, ensure_ascii=False, indent=2)
    if a.out:
        out_dir = os.path.dirname(os.path.abspath(a.out))
        os.makedirs(out_dir, exist_ok=True)
        with open(a.out, "w", encoding="utf-8") as fh:
            fh.write(js + "\n")
        print(f"已写入 {a.out}")
    print(js)


if __name__ == "__main__":
    main()
