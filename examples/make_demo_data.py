"""生成「铺货代销逐月对账」演示数据（全部 .xlsx，无网络依赖，输出确定）。

生成结构（客户根目录按主档设计文档的约定）：
    <输出目录>/铺货代销客户/示例经销商有限公司/
    ├── _主档.xlsx                        # 六表：主体/合同/价格/SKU/期间/贴花
    └── 202610/
        ├── 甲方-20261102ERP挂账.xlsx           # 甲方 ERP 未结算挂账（首行为标题行，验证表头定位）
        ├── 乙方-202610销售数据.xlsx            # 乙方当期动销流水（含同号拆行、日期缺口）
        └── 乙方-20261102库存+当月销售.xlsx     # 乙方库存明细（含同条码拆行）+ 11 月 1—2 日销售

演示数据刻意覆盖以下引擎分支：
    - SCD2 生效区间：价格/贴花均有「旧版 + 2026-11-02 新版」两行，本期取期末（2026-10-31）生效版；
    - 时点回加：库存导出日 2026-11-02 晚于期末，11 月 1—2 日销售需回加；
    - 库存差异：1 个 SKU 短缺、1 个 SKU 盘盈；
    - 多行库存：同一 SKU 拆两行，需按条码/品名合计；
    - 品名变体：半角括号、空格、「（限定版）」后缀，均按条码匹配；
    - 非结算类：陈列装库存单独成桶，不参与结算；
    - 零库存新品：ERP 以 0 数量挂账、乙方库存为 0；
    - 贴花：结算类动销 SKU 全部有当期记录（显式 0 也算），合计 980.00 元。

用法：
    python3 examples/make_demo_data.py <输出目录>
"""
import argparse
import os
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# 复用 make_master.py 的表头/列宽定义，保证演示主档与 skill 自带模板同构
_SCRIPTS = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "skills", "dealer-settlement", "scripts")
if _SCRIPTS not in sys.path:
    sys.path.insert(0, _SCRIPTS)
from make_master import SHEETS  # noqa: E402  （仅取表头与列宽，不写演示行）

PERIOD = "202610"
PERIOD_START = "2026-10-01"
PERIOD_END = "2026-10-31"
ASOF = "2026-11-02"
NEXT_SALES_DAYS = "2026-11-01 至 2026-11-02"

PARTY_A = {
    "角色": "甲方", "名称": "示例品牌方有限公司",
    "统一社会信用代码": "91440100DEMO00001A",
    "地址": "示例市示例区示例路100号", "联系人": "张示例", "电话": "",
    "开户行": "示例银行股份有限公司示例支行", "账号": "44056601040012345",
    "户名": "示例品牌方有限公司",
}
PARTY_B = {
    "角色": "乙方", "名称": "示例经销商有限公司",
    "统一社会信用代码": "91310000DEMO00002B",
    "地址": "示例市示例区示例路200号", "联系人": "李示例", "电话": "13800000000",
    "开户行": "", "账号": "", "户名": "",
}
CONTRACT = {
    "合同编号": "HT-DEMO-2026-001", "起始": "2026-10-19", "终止": "2026-12-31",
    "结算方式": "铺货代销", "结算基数": "实际动销", "付款日": "次月5日前",
    "异议期(工作日)": 3, "签回期(工作日)": 5, "价格附件": "示例政策表v2",
    "备注": "演示数据（虚构）：额度、动销保障比例等条款以合同原文为准",
}

# ---- SKU 主数据：供货价/贴花标准取「2026-10-31 当期生效版本」 ----
# price_rows / std_rows 为 SCD2 行；只有标了 *_new 的 SKU 才有 2026-11-02 起重定价/调标行。
SKUS = [
    {
        "bar": "690000000001", "name": "示例面膜A（盒装5片）", "alias": "",
        "spec": "盒装5片", "unit": "盒", "attr": "结算类",
        "supply_price": 28.00, "std": 4.00,
        "price_rows": [("28.00", "2026-10-19", "")],
        "std_rows": [("4.00", "2026-10-19", "2026-11-01"), ("4.50", "2026-11-02", "")],
        "sales_name": "示例面膜A（盒装5片）", "inv_name": "示例面膜A（盒装5片）",
    },
    {
        "bar": "690000000002", "name": "示例面膜B（盒装5片）", "alias": "示例面膜B(盒装5片)",
        "spec": "盒装5片", "unit": "套", "attr": "结算类",
        "supply_price": 35.00, "std": 2.50,
        "price_rows": [("35.00", "2026-10-19", "")],
        "std_rows": [("2.50", "2026-10-19", "")],
        "sales_name": "示例面膜B(盒装5片)", "inv_name": "示例面膜B（盒装5片）",
    },
    {
        "bar": "690000000005", "name": "示例颈膜（盒装5片）", "alias": "",
        "spec": "盒装5片", "unit": "盒", "attr": "结算类",
        "supply_price": 26.50, "std": 0.00,
        "price_rows": [("26.50", "2026-10-19", "2026-11-01"), ("26.40", "2026-11-02", "")],
        "std_rows": [("0.00", "2026-10-19", "")],
        "sales_name": "示例颈膜（盒装5片）", "inv_name": "示例颈膜（盒装5片）",
    },
    {
        "bar": "690000000012", "name": "示例面膜E（盒装5片）", "alias": "",
        "spec": "盒装5片", "unit": "盒", "attr": "结算类",
        "supply_price": 12.50, "std": 1.20,
        "price_rows": [("12.50", "2026-10-19", "2026-11-01"), ("12.80", "2026-11-02", "")],
        "std_rows": [("1.20", "2026-10-19", "")],
        "sales_name": "示例面膜E（盒装5片）", "inv_name": "示例面膜E（盒装5片）",
    },
    {
        "bar": "690000000010", "name": "示例油敷膜（盒装5片·限定版）", "alias": "示例油敷膜（盒装5片）",
        "spec": "盒装5片", "unit": "盒", "attr": "结算类",
        "supply_price": 30.00, "std": 3.00,
        "price_rows": [("30.00", "2026-10-19", "")],
        "std_rows": [("3.00", "2026-10-19", "")],
        "sales_name": "示例油敷膜（盒装5片）", "inv_name": "示例油敷膜（盒装5片·限定版）",
    },
    {
        "bar": "690000000020", "name": "示例精华蜜15ml", "alias": "",
        "spec": "15ml", "unit": "盒", "attr": "结算类",
        "supply_price": 18.00, "std": 0.50,
        "price_rows": [("18.00", "2026-10-19", "")],
        "std_rows": [("0.50", "2026-10-19", "")],
        "sales_name": "示例精华蜜15ml", "inv_name": "示例精华蜜15ml",
    },
]
# 批次夹具：每行 (批次号, 有效期至, 未结算数量, 初始数量)。同一 SKU 合计必须等于期初数量，
# 否则会改变库存差异。示例面膜E 的第二批已发生扣减、而第一批仍有余额，用于演示 FIFO 校验。
LOTS = {
    "690000000001": [("A1-202612", "2026-12-15", 120, 120),
                     ("A2-202706", "2027-06-30", 180, 180)],
    "690000000002": [("B1-202611", "2026-11-20", 200, 200)],
    "690000000005": [("C1-202610", "2026-10-20", 40, 40),
                     ("C2-202703", "2027-03-31", 200, 200)],
    "690000000012": [("D1-202701", "2027-01-31", 300, 300),
                     ("D2-202709", "2027-09-30", 200, 260)],
    "690000000010": [("E1-202705", "2027-05-31", 120, 120)],
    "690000000020": [("F1-202612", "2026-12-31", 0, 0)],   # 零库存新品：批次数量为 0
}

# 促销物料：只在甲方挂账表出现（乙方不一定入库），不参与货款结算
PROMO = [
    ("示例面膜A 试用装（单片）", "690000000101", "试用装", 600, "SY-A1", "2027-03-31"),
    ("示例精华蜜 中小样 5ml", "690000000102", "中小样", 400, "SY-B1", "2026-12-15"),
    ("示例面膜B 配赠装（2 片）", "690000000103", "配赠", 150, "SY-C1", "2027-01-31"),
]
# 乙方使用数量（可选文件才生成）：用于演示「乙方已提供使用数据」分支
PROMO_USE = [("示例面膜A 试用装（单片）", "690000000101", 500, "示例门店一店", "2026-10-20"),
             ("示例精华蜜 中小样 5ml", "690000000102", 300, "示例门店二店", "2026-10-25"),
             ("示例面膜B 配赠装（2 片）", "690000000103", 150, "示例门店一店", "2026-10-28")]

NONSETTLE = {
    "bar": "690000000016", "name": "示例面膜C（单片）", "alias": "示货示例面膜C 单片",
    "spec": "单片", "unit": "片", "attr": "非结算类",
    "sales_name": "示例面膜C（单片）", "inv_name": "示例面膜C（单片）",
}

# 期初/动销/回加/库存：q0 - q10 = 理论结存；act + q11 = 调整后结存
QTY = {
    "690000000001": {"q0": 300, "q10": 120, "q11": 10, "act": 170},   # 账实相符
    "690000000002": {"q0": 200, "q10": 80, "q11": 5, "act": 107},     # 短缺 8 件
    "690000000005": {"q0": 240, "q10": 60, "q11": 0, "act": 180},     # 账实相符
    "690000000012": {"q0": 500, "q10": 150, "q11": 20, "act": 335},   # 盘盈 5 件
    "690000000010": {"q0": 120, "q10": 40, "q11": 4, "act": 76},      # 账实相符
    "690000000020": {"q0": 0, "q10": 0, "q11": 0, "act": 0},          # 零库存新品
}
EXTRA_ACT = 50  # 非结算类库存件数（甲方挂账未列示）

STORES = ["示例门店一店", "示例门店二店", "示例门店三店"]

# 乙方当期销售流水：(流水号, 销售时间, 门店, 条码, 数量)
# 其中 XS2026100001 同流水号 + 同条码 + 同数量出现两行，用于触发 07 表拆行核对说明。
SALES = [
    ("XS2026100001", "2026-10-01 09:12:00", "示例门店一店", "690000000001", 10),
    ("XS2026100001", "2026-10-01 09:15:00", "示例门店二店", "690000000001", 10),
    ("XS2026100002", "2026-10-05 14:20:00", "示例门店一店", "690000000001", 40),
    ("XS2026100003", "2026-10-12 10:05:00", "示例门店三店", "690000000001", 60),
    ("XS2026100004", "2026-10-02 11:00:00", "示例门店一店", "690000000002", 30),
    ("XS2026100005", "2026-10-08 16:30:00", "示例门店二店", "690000000002", 50),
    ("XS2026100006", "2026-10-06 13:45:00", "示例门店三店", "690000000005", 60),
    ("XS2026100007", "2026-10-03 10:10:00", "示例门店一店", "690000000012", 70),
    ("XS2026100008", "2026-10-10 15:00:00", "示例门店二店", "690000000012", 50),
    ("XS2026100009", "2026-10-20 17:20:00", "示例门店三店", "690000000012", 30),
    ("XS2026100010", "2026-10-07 09:30:00", "示例门店一店", "690000000010", 25),
    ("XS2026100011", "2026-10-18 11:40:00", "示例门店二店", "690000000010", 15),
]

# 库存导出月 1—2 日销售（回加用）：(条码, 数量)
NEXT_SALES = [
    ("690000000001", 6),
    ("690000000001", 4),
    ("690000000002", 5),
    ("690000000012", 20),
    ("690000000010", 4),
]


def _expect(cond, msg):
    if not cond:
        raise SystemExit("[演示数据自检失败] " + msg)


def _check_fixture():
    """保证夹具自身先满足六条恒等式，避免用错误数据去测引擎。"""
    per_bar_sales = {}
    for _sid, _t, _st, bar, q in SALES:
        per_bar_sales[bar] = per_bar_sales.get(bar, 0) + q
    per_bar_nov = {}
    for bar, q in NEXT_SALES:
        per_bar_nov[bar] = per_bar_nov.get(bar, 0) + q
    for sku in SKUS:
        bar = sku["bar"]
        qq = QTY[bar]
        _expect(per_bar_sales.get(bar, 0) == qq["q10"],
                "%s 销售流水合计 %s != q10 %s" % (bar, per_bar_sales.get(bar, 0), qq["q10"]))
        _expect(per_bar_nov.get(bar, 0) == qq["q11"],
                "%s 次月销售合计 %s != q11 %s" % (bar, per_bar_nov.get(bar, 0), qq["q11"]))
        _expect(sku["supply_price"] > 0, "%s 必须有供货价" % bar)
        _expect(sku["std"] >= 0, "%s 贴花标准不得为负" % bar)
    _expect(len({s["bar"] for s in SKUS}) == len(SKUS), "条码重复")
    _expect(len({s["bar"] for s in SKUS} | {NONSETTLE["bar"]}) == len(SKUS) + 1, "条码重复")


def _header_style(ws, ncols, widths):
    for i in range(1, ncols + 1):
        if widths and i <= len(widths):
            ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
        cell = ws.cell(row=1, column=i)
        cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin = Side(style="thin", color="BFBFBF")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False


def _fill_sheet(wb, name, rows):
    headers, widths = SHEETS[name][0], SHEETS[name][1]
    ws = wb.create_sheet(name)
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    _header_style(ws, len(headers), widths)
    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name="微软雅黑", size=10)
            cell.alignment = Alignment(vertical="center",
                                       horizontal="right" if isinstance(v, (int, float)) else "left",
                                       wrap_text=True)
            if isinstance(v, float):
                cell.number_format = "#,##0.00"
    return ws


def _master_sheets():
    sku_rows = []
    for s in SKUS + [NONSETTLE]:
        sku_rows.append([s["bar"], s["name"], s["alias"], s["spec"], s["unit"],
                         "24盒/箱" if s["unit"] in ("盒", "套") else "",
                         s["attr"], ""])
    price_rows = []
    for s in SKUS:
        for price, start, end in s["price_rows"]:
            price_rows.append([s["bar"], s["name"], float(price), start, end,
                               169.0, "示例政策表v1" if end else "示例政策表v2"])
    std_rows = []
    for s in SKUS:
        for std, start, end in s["std_rows"]:
            std_rows.append([s["bar"], s["name"], float(std), start, end,
                             "示例价格政策表（原版）" if end else "示例价格政策表20261102更新", "演示行"])
    contract_row = [CONTRACT[h] for h in SHEETS["合同"][0]]
    return {
        "主体": [list(PARTY_A[h] for h in SHEETS["主体"][0]),
                 list(PARTY_B[h] for h in SHEETS["主体"][0])],
        "合同": [contract_row],
        "价格": price_rows,
        "SKU": sku_rows,
        "期间": [],
        "贴花": std_rows,
    }


def build_master(path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    data = _master_sheets()
    for name in ("主体", "合同", "价格", "SKU", "期间", "贴花"):
        _fill_sheet(wb, name, data[name])
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
    return path


def build_bip(path):
    """甲方 ERP 未结算挂账：首行是导出标题，真正的表头在第 2 行（验证表头定位）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "挂账"
    ws["A1"] = "甲方 ERP 系统未结算挂账（演示数据，%s 导出）" % ASOF
    ws["A1"].font = Font(name="微软雅黑", size=11, bold=True)
    headers = ["商品名称", "单价", "未结算金额", "未结算数量", "初始数量",
               "批次号", "有效期至", "物料属性"]
    widths = [40, 10, 14, 12, 12, 14, 14, 12]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=i, value=h)
        cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    r = 3
    for s in SKUS:
        price = s["supply_price"]
        lots = LOTS[s["bar"]]
        _expect(sum(x[2] for x in lots) == QTY[s["bar"]]["q0"],
                "%s 批次数量合计 %s != 期初 %s" % (s["bar"], sum(x[2] for x in lots), QTY[s["bar"]]["q0"]))
        for lot, exp, qty, init in lots:
            ws.cell(row=r, column=1, value=s["name"]).font = Font(name="微软雅黑", size=10)
            ws.cell(row=r, column=2, value=price).number_format = "#,##0.00"
            ws.cell(row=r, column=3, value=round(price * qty, 2)).number_format = "#,##0.00"
            ws.cell(row=r, column=4, value=qty)
            ws.cell(row=r, column=5, value=init)
            ws.cell(row=r, column=6, value=lot)
            ws.cell(row=r, column=7, value=exp)
            ws.cell(row=r, column=8, value="正装")
            r += 1
    for name, bar, kind, qty, lot, exp in PROMO:
        ws.cell(row=r, column=1, value=name).font = Font(name="微软雅黑", size=10)
        ws.cell(row=r, column=2, value=0).number_format = "#,##0.00"
        ws.cell(row=r, column=3, value=0).number_format = "#,##0.00"
        ws.cell(row=r, column=4, value=qty)
        ws.cell(row=r, column=5, value=qty)
        ws.cell(row=r, column=6, value=lot)
        ws.cell(row=r, column=7, value=exp)
        ws.cell(row=r, column=8, value=kind)
        r += 1
    ws.sheet_view.showGridLines = False
    wb.save(path)
    return path


def build_promo_use(path):
    """乙方试用装/中小样/配赠使用明细（可选源，演示「已提供」分支）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "使用明细"
    headers = ["商品名称", "商品条码", "使用数量", "使用门店", "使用日期"]
    widths = [34, 16, 12, 16, 14]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    for r, (name, bar, qty, store, day) in enumerate(PROMO_USE, 2):
        ws.cell(row=r, column=1, value=name).font = Font(name="微软雅黑", size=10)
        cell = ws.cell(row=r, column=2, value=bar)
        cell.number_format = "@"
        ws.cell(row=r, column=3, value=qty)
        ws.cell(row=r, column=4, value=store)
        ws.cell(row=r, column=5, value=day)
    ws.sheet_view.showGridLines = False
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
    return path


def _name_of(bar, kind="sales"):
    for s in SKUS:
        if s["bar"] == bar:
            return s["sales_name"] if kind == "sales" else s["inv_name"]
    if bar == NONSETTLE["bar"]:
        return NONSETTLE["sales_name"] if kind == "sales" else NONSETTLE["inv_name"]
    raise KeyError(bar)


def _unit_spec_of(bar):
    for s in SKUS + [NONSETTLE]:
        if s["bar"] == bar:
            return s["unit"], s["spec"]
    raise KeyError(bar)


def build_sales(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "2026年10月"
    headers = ["流水号", "销售时间", "销售门店", "商品名称", "商品条码", "规格", "单位", "销售数量", "品牌"]
    widths = [18, 20, 16, 34, 16, 12, 6, 10, 10]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    for r, (sid, when, store, bar, qty) in enumerate(SALES, 2):
        unit, spec = _unit_spec_of(bar)
        row = [sid, when, store, _name_of(bar), bar, spec, unit, qty, "示例品牌"]
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name="微软雅黑", size=10)
            if c in (1, 5):
                cell.number_format = "@"
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    wb.save(path)
    return path


def build_inventory(path):
    """乙方库存（含同条码拆行）＋ 11 月 1—2 日销售（回加用）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "库存明细"
    headers = ["商品名称", "商品条码", "规格", "商品分类", "库存数量", "销售价", "品牌"]
    widths = [40, 16, 12, 14, 10, 10, 12]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]

    rows = []
    # 示例面膜A 拆两行（100 + 70 = 170），用于验证库存按条码合计
    rows.append((_name_of("690000000001", "inv"), "690000000001", "盒装5片", "贴片面膜", 100, 169.0, "示例品牌"))
    rows.append((_name_of("690000000001", "inv"), "690000000001", "盒装5片", "贴片面膜", 70, 169.0, "示例品牌"))
    for s in SKUS:
        if s["bar"] == "690000000001":
            continue
        act = QTY[s["bar"]]["act"]
        rows.append((s["inv_name"], s["bar"], s["spec"], "面膜类", act,
                     169.0 if s["bar"] != "690000000020" else 59.0, "示例品牌"))
    rows.append((NONSETTLE["inv_name"], NONSETTLE["bar"], NONSETTLE["spec"], "陈列装",
                 EXTRA_ACT, 9.9, "示例品牌"))
    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name="微软雅黑", size=10)
            if c == 2:
                cell.number_format = "@"
            if c in (5, 6):
                cell.number_format = "#,##0.00" if c == 6 else "#,##0"
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

    ws2 = wb.create_sheet("11月销售")
    headers2 = ["商品名称", "商品条码", "销售数量"]
    for i, h in enumerate(headers2, 1):
        cell = ws2.cell(row=1, column=i, value=h)
        cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws2.column_dimensions[get_column_letter(i)].width = [40, 16, 10][i - 1]
    for r, (bar, qty) in enumerate(NEXT_SALES, 2):
        ws2.cell(row=r, column=1, value=_name_of(bar, "sales")).font = Font(name="微软雅黑", size=10)
        cell = ws2.cell(row=r, column=2, value=bar)
        cell.font = Font(name="微软雅黑", size=10)
        cell.number_format = "@"
        ws2.cell(row=r, column=3, value=qty)
    ws2.freeze_panes = "A2"
    ws2.sheet_view.showGridLines = False
    wb.save(path)
    return path


def build_koujing(root, period="202610", asof="2026-11-02"):
    """演示用的口径确认记录；真实使用必须由阶段 G 问卡逐项确认后生成。"""
    import datetime
    import json
    ledger = []
    for name, bar, kind, qty, lot, exp in PROMO:
        ledger.append(dict(name=name, kind=kind, link=name, sales=0,
                           opening=0, shipped=qty, used=1 if kind == "试用装" else 0))
    kj = {
        "schema": "dealer-settlement/koujing@1",
        "customer": "示例经销商有限公司", "period": period, "asof": asof,
        "scene": "每月出账", "confirmed_by": "演示确认人",
        "confirmed_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "decisions": {
            "settlement_basis": {"source": "甲方挂账表", "field": "未结算数量",
                                 "excluded": ["效期参考表的结存数量"]},
            "category_source": {"source": "商品分类", "正装": 0, "试用装": 600, "中小样": 550},
            "promo": {"settle": False, "separate_sheet": True, "confirm_by_party_b": True,
                      "policy_note": "为甲方政策支持"},
            "rolling": {"opening_source": "上期甲方期末", "carry_forward": True,
                        "first_period_opening": 0},
            "usage": {"values": {"试用装": 1, "中小样": 0, "配赠": 0},
                      "source": "甲方口径", "private": True, "ledger": ledger},
            "expiry": {"enabled": False, "unit": "month", "tiers": [32, 28, 24],
                       "base": "asof", "fifo": "display_only",
                       "sample_qty_basis": "台账期末", "next_period": True},
            "barcode_bridge": {"required": True},
            "sheet_map": {"08": "试用装与中小样", "09": "效期预警"},
            "print": {"red_sheets": ["00", "01", "02", "08"],
                      "pdf_sheets": ["00", "01", "02", "08"]},
            "delta": {"mode": "full", "frozen_sheets": [],
                      "note_only_sheets": [], "rebuild_sheets": []},
        },
        "internal_only": ["演示用出数规则", "演示用内部费用目的"],
        "disclosure_whitelist": ["本表所列商品不参与货款结算"],
        "notes": [],
    }
    path = os.path.join(root, "_口径确认.json")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(kj, fh, ensure_ascii=False, indent=2)
    return path


def build(outdir, with_promo_use=False):
    """在 outdir 下生成完整演示数据，返回 (客户根目录, 文件路径表)。

    with_promo_use=True 时额外生成乙方使用明细，用于演示「乙方已提供使用数据」分支；
    默认不生成，走「乙方未提供」分支。
    """
    _check_fixture()
    root = os.path.join(os.path.abspath(outdir), "铺货代销客户", "示例经销商有限公司")
    month = os.path.join(root, PERIOD)
    os.makedirs(month, exist_ok=True)
    paths = {
        "master": build_master(os.path.join(root, "_主档.xlsx")),
        "jia": build_bip(os.path.join(month, "甲方-20261102ERP挂账.xlsx")),
        "sales": build_sales(os.path.join(month, "乙方-202610销售数据.xlsx")),
        "inventory": build_inventory(os.path.join(month, "乙方-20261102库存+当月销售.xlsx")),
    }
    if with_promo_use:
        paths["promo_use"] = build_promo_use(os.path.join(month, "乙方-202610试用装使用.xlsx"))
    build_koujing(root)
    return root, paths


def main(argv=None):
    ap = argparse.ArgumentParser(description="生成铺货代销月度对账演示数据（.xlsx）")
    ap.add_argument("outdir", help="输出目录（示例：/tmp/dealer_settlement_demo）")
    a = ap.parse_args(argv)
    root, paths = build(a.outdir)
    print("演示客户根目录：", root)
    for k in ("master", "jia", "sales", "inventory"):
        print("  %-9s %s" % (k, os.path.relpath(paths[k], root)))
    print("期间：%s（期末 %s，库存导出日 %s）" % (PERIOD, PERIOD_END, ASOF))
    return 0


if __name__ == "__main__":
    sys.exit(main())
