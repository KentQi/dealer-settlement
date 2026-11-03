"""生成客户主档模板 _主档.xlsx（六张表：主体 / 合同 / 价格 / SKU / 期间 / 贴花）。

用法：
    python3 make_master.py <输出路径>      # 生成到指定路径，如 <客户根目录>/_主档.xlsx
    python3 make_master.py --template      # 更新 skill 自带模板 assets/客户主档-模板.xlsx

不带参数只打印用法并退出（不会生成任何文件），避免无意覆盖 skill 自带模板。
示例行为演示数据（虚构），实际使用请整行替换为本客户信息。
"""
import sys, os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

NAVY, GREY, LBLUE = "1F3864", "F2F2F2", "D9E2F3"
THIN = Side(style="thin", color="BFBFBF")

SHEETS = {
    "主体": (
        ["角色", "名称", "统一社会信用代码", "地址", "联系人", "电话", "开户行", "账号", "户名"],
        [14, 34, 24, 40, 12, 16, 34, 26, 34],
        [
            ["甲方", "示例品牌方有限公司", "91440100DEMO00001A",
             "示例市示例区示例路100号", "张示例", "",
             "示例银行股份有限公司示例支行", "44056601040012345",
             "示例品牌方有限公司"],
            ["乙方", "示例经销商有限公司", "91310000DEMO00002B",
             "示例市示例区示例路200号", "李示例", "13800000000", "", "", ""],
        ],
    ),
    "合同": (
        ["合同编号", "起始", "终止", "结算方式", "结算基数", "付款日", "异议期(工作日)", "签回期(工作日)", "价格附件", "备注"],
        [22, 12, 12, 14, 18, 16, 15, 15, 20, 30],
        [["HT-DEMO-2026-001", "2026-10-19", "2026-12-31", "铺货代销", "实际动销",
          "次月5日前", 3, 5, "示例政策表v2", "额度示例；动销保障比例示例"]],
    ),
    "价格": (
        ["商品条码", "商品名称", "供货价", "生效起", "生效止", "零售价", "依据文件"],
        [16, 40, 10, 12, 12, 10, 26],
        [
            ["690000000001", "示例面膜A（盒装5片）", 28.00, "2026-10-19", "", 169, "示例政策表v1"],
            ["690000000005", "示例颈膜（盒装5片）", 26.50, "2026-10-19", "2026-11-01", 129, "示例政策表v1"],
            ["690000000005", "示例颈膜（盒装5片）", 26.40, "2026-11-02", "", 129, "示例政策表v2"],
            ["690000000012", "示例面膜E（盒装5片）", 12.50, "2026-10-19", "2026-11-01", 79, "示例政策表v1"],
            ["690000000012", "示例面膜E（盒装5片）", 12.80, "2026-11-02", "", 79, "示例政策表v2"],
        ],
    ),
    "SKU": (
        ["商品条码", "标准品名", "乙方品名别名", "规格", "单位", "箱规", "结算属性", "备注"],
        [16, 40, 40, 12, 6, 12, 12, 26],
        [
            ["690000000001", "示例面膜A（盒装5片）", "", "盒装5片", "盒", "24盒/箱", "结算类", ""],
            ["690000000005", "示例颈膜（盒装5片）", "", "盒装5片", "盒", "24盒/箱", "结算类", ""],
            ["690000000002", "示例面膜B（盒装5片）", "示例面膜B(盒装5片)",
             "盒装5片", "套", "24盒/箱", "结算类", "乙方系统括号为半角（示例）"],
            ["690000000010", "示例油敷膜（盒装5片·限定版）",
             "示例油敷膜（盒装5片）", "盒装5片", "盒", "24盒/箱", "结算类", "甲方名称含（限定版）后缀（示例）"],
            ["690000000016", "示例面膜C（单片）", "示货示例面膜C 单片", "单片", "片", "",
             "非结算类", "乙方品名笔误（示例）；陈列装，未挂账"],
        ],
    ),
    "期间": (
        ["期间", "数据基准日", "期初数量", "期初金额", "结算数量", "结算金额", "差异数量", "差异金额",
         "贴花金额", "贴花兑付方式", "对账单编号", "状态", "备注"],
        [10, 12, 10, 14, 10, 14, 10, 14, 12, 14, 18, 10, 34],
        [["202610", "2026-11-02", 1000, 50000.00, 120, 5000.00, -20, -800.00, 200.00, "待定", "", "待确认",
          "示例行为演示数据（虚构）；贴花兑付方式（抵扣货款/分开支付）由双方书面确认后更新"]],
    ),
    "贴花": (
        ["商品条码", "商品名称", "贴花标准(元/件)", "生效起", "生效止", "依据文件", "备注"],
        [16, 40, 16, 12, 12, 30, 40],
        [
            # 贴花＝销售结算时给乙方的奖励，按本期动销数量 × 当期生效标准计算。
            # 与价格表同构（SCD2）：调标时旧行补「生效止」、新行填「生效起」，两行并存。
            # 本期有动销的每个 SKU 必须有当期有效记录（显式 0 也算），缺失会被引擎阻断。
            ["690000000001", "示例面膜A（盒装5片）", 4.00, "2026-10-19", "2026-11-01", "示例价格政策表（原版）", "示例行"],
            ["690000000001", "示例面膜A（盒装5片）", 4.00, "2026-11-02", "", "示例价格政策表20261102更新", "示例行：调标照此两行并存"],
            ["690000000005", "示例颈膜（盒装5片）", 0.00, "2026-10-19", "", "示例价格政策表", "示例行：无贴花也须显式登记为 0"],
        ],
    ),
}



def build(path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, (headers, widths, rows) in SHEETS.items():
        ws = wb.create_sheet(name)
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        ws.row_dimensions[1].height = 30
        for r, row in enumerate(rows, 2):
            for c, v in enumerate(row, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = Font(name="微软雅黑", size=10)
                cell.alignment = Alignment(vertical="center",
                                           horizontal="right" if isinstance(v, (int, float)) else "left",
                                           wrap_text=True)
                cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
                if isinstance(v, float) and headers[c - 1] in ("供货价", "零售价"):
                    cell.number_format = "#,##0.00"
                elif isinstance(v, (int, float)) and "金额" in headers[c - 1]:
                    cell.number_format = "#,##0.00"
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"
        ws.sheet_view.showGridLines = False
    if "SKU" in wb.sheetnames:
        ws = wb["SKU"]
        dv = DataValidation(type="list", formula1='"结算类,非结算类"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"G2:G500")
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
    print("已生成:", path)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("用法：python3 make_master.py <输出路径/_主档.xlsx>\n"
                 "不带参数不会生成任何文件——避免无意覆盖 skill 自带模板。"
                 "模板维护请用 --template 参数。")
    if sys.argv[1] == "--template":
        out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "..", "assets", "客户主档-模板.xlsx")
        print("[模板模式] 更新 skill 自带模板")
    else:
        out = sys.argv[1]
    build(os.path.abspath(out))
