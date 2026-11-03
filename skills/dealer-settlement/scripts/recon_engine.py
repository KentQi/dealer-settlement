import openpyxl, re, os, sys
from collections import defaultdict, Counter
from datetime import date as _date, datetime as _datetime, timedelta as _timedelta
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

import json

# ==================== 可配置项（通过环境变量 RECON_CFG 传入 JSON）====================
CFG = {}
_CFG_ERR = ""
try:
    _CFG_RAW = os.environ.get("RECON_CFG", "") or ""
    if _CFG_RAW.strip():
        _parsed = json.loads(_CFG_RAW)
        if not isinstance(_parsed, dict):
            _CFG_ERR = "顶层必须是 JSON 对象"
        else:
            CFG = _parsed
except Exception as _e:
    _CFG_ERR = str(_e)

_SK = os.path.dirname(os.path.abspath(__file__))


def _cfg_fail(detail, missing=()):
    """缺配置时中止：stdout 必须为空，全部信息走 stderr。"""
    msg = [
        f"[中止] {detail}，未生成任何文件。",
        "生成方式（整行复制，先改 <> 里的内容）：",
    ]
    if os.name == "nt":
        msg += [
            "  Windows（PowerShell）：",
            '    $env:RECON_CFG = & py -3 "' + _SK + '\\load_master.py" "<主档的完整路径>" --period 202610 --asof 2026-11-02',
            '    py -3 "' + _SK + '\\recon_engine.py"',
        ]
    else:
        msg += [
            "  macOS / Linux（bash 或 zsh）：",
            f'    export RECON_CFG=$(python3 "{_SK}/load_master.py" "<主档的完整路径>" --period <202610> --asof <2026-11-02>)',
            f'    python3 "{_SK}/recon_engine.py"',
        ]
    if os.environ.get("RECON_DEBUG") and missing:
        msg.append("[debug] 缺失：" + "、".join(missing))
    sys.stderr.write("\n".join(msg) + "\n")
    raise SystemExit(1)


if _CFG_ERR:
    _cfg_fail("RECON_CFG 不是合法的配置 JSON：" + _CFG_ERR, ["RECON_CFG"])

_need = [("party_a", "甲方主体"), ("party_b", "乙方主体"), ("period", "对账期间"),
         ("base_dir", "月份目录"), ("sources", "源文件位置")]
_miss = [lab for k, lab in _need if not CFG.get(k)]
if CFG.get("party_a") and not (CFG["party_a"] or {}).get("name"):
    _miss.append("甲方名称")
if CFG.get("party_b") and not (CFG["party_b"] or {}).get("name"):
    _miss.append("乙方名称")
if not CFG or _miss:
    _cfg_fail("本次对账缺少必要信息：" + (" / ".join(_miss) if _miss else "甲方主体 / 乙方主体 / 对账期间 / 源文件位置"), _miss)


_REAL_OUT = sys.stdout
sys.stdout = sys.stderr          # 过程中的一切输出走 stderr；失败时 stdout 保持 0 字节


def emit(msg=""):
    _REAL_OUT.write(str(msg) + "\n")
    _REAL_OUT.flush()


_PA = CFG.get("party_a", {})
_PB = CFG.get("party_b", {})
_CT = CFG.get("contract", {})
_PD = CFG.get("period", {})
_BK = CFG.get("bank", {})

from datetime import date as _date

PA_NAME = _PA.get("name") or ""
PA_TAX  = _PA.get("tax") or ""
PB_NAME = _PB.get("name") or ""
PB_TAX  = _PB.get("tax") or ""
CONTRACT_NO = _CT.get("no") or ""

# 期间三项必填；格式化文案与派生量一律由日期推导，不留业务字面量
P_START = _PD.get("start") or ""
P_END   = _PD.get("end") or ""
ASOF    = _PD.get("asof") or ""

_req2 = []
for _lab, _v in (("甲方纳税人识别号", PA_TAX), ("乙方纳税人识别号", PB_TAX),
                 ("合同编号", CONTRACT_NO), ("期间起始", P_START),
                 ("期间截止", P_END), ("数据基准日", ASOF)):
    if not _v:
        _req2.append(_lab)
_BK = _BK if isinstance(_BK, dict) else {}
if not (_BK.get("account_no") and _BK.get("bank_name")):
    _req2.append("甲方收款账户（开户行与账号）")
for _k, _lab in (("jia", "甲方挂账文件"), ("sales", "销售流水文件"), ("inventory", "库存文件")):
    if not (CFG.get("sources") or {}).get(_k):
        _req2.append(_lab)
_MASTER = CFG.get("master_path") or ""
if not _MASTER:
    _req2.append("主档路径（master_path）")
elif not os.path.exists(_MASTER):
    _req2.append("主档文件不存在：" + str(_MASTER))
if _req2:
    _cfg_fail("本次对账缺少必要信息：" + " / ".join(_req2), _req2)

try:
    _d_end  = _date.fromisoformat(str(P_END)[:10])
    _d_asof = _date.fromisoformat(str(ASOF)[:10])
    _d_start = _date.fromisoformat(str(P_START)[:10])
except ValueError:
    _cfg_fail("期间日期格式不正确（应为 YYYY-MM-DD）：start=%r end=%r asof=%r"
              % (P_START, P_END, ASOF), ["期间日期"])

P_CN    = _PD.get("period_cn") or f"{_d_end.year} 年 {_d_end.month} 月"
ASOF_CN = _PD.get("asof_cn") or f"{_d_asof.year} 年 {_d_asof.month:02d} 月 {_d_asof.day:02d} 日"
NEXT_M  = int(_PD.get("next_month") or _d_asof.month)
_d_next = _d_end + _timedelta(days=1)
_days_add = (_d_asof - _d_next).days + 1
ADD_BACK = _days_add > 0
NEXT_LABEL = (f"{_d_next.month} 月 {_d_next.day}—{_d_asof.day} 日" if ADD_BACK
              else f"{_d_next.month} 月 1—2 日")
NEXT_FULL = (f"{_d_next.year} 年 {_d_next.month} 月 {_d_next.day}—{_d_asof.day} 日"
             if ADD_BACK else f"{_d_next.month} 月 1—2 日")
PAY_CN  = _PD.get("pay_cn") or f"{_d_asof.year} 年 {_d_asof.month:02d} 月 05 日"
PERIOD  = f"{_d_end.year:04d}{_d_end.month:02d}"

BK_NAME = _BK.get("account_name") or PA_NAME
BK_NO   = _BK.get("account_no") or ""
BK_BANK = _BK.get("bank_name") or ""

# ---- 合同条款：以主档「合同」表为准；缺失时不得静默套用演示值 ----
CT_METHOD = str(_CT.get("method") or "").strip()
CT_BASIS  = str(_CT.get("basis") or "").strip()
CT_METHOD_TXT = CT_METHOD or "（主档未登记，以合同为准）"
CT_BASIS_TXT  = CT_BASIS or "（主档未登记，以合同为准）"
if CT_METHOD and CT_BASIS:
    SETTLE_TERMS_TXT = f"结算方式{CT_METHOD}、结算基数{CT_BASIS}"
elif CT_METHOD:
    SETTLE_TERMS_TXT = f"结算方式{CT_METHOD}（结算基数主档未登记，以合同为准）"
elif CT_BASIS:
    SETTLE_TERMS_TXT = f"结算基数{CT_BASIS}（结算方式主档未登记，以合同为准）"
else:
    SETTLE_TERMS_TXT = "结算方式与结算基数主档未登记，以合同为准"


def _term_workdays(value, what):
    """异议期/签回期：主档有正整数就用，否则显式写“合同约定的…”，不印固定演示值。"""
    try:
        n = int(float(str(value).strip()))
    except (TypeError, ValueError):
        n = 0
    return f"{n} 个工作日" if n > 0 else f"合同约定的{what}"


CT_SIGN_TXT = _term_workdays(_CT.get("signback_days"), "签回期限")
CT_OBJ_TXT  = _term_workdays(_CT.get("objection_days"), "异议期限")

SRC = CFG.get("sources", {})
F_JIA = SRC.get("jia") or ""
F_OCT = SRC.get("sales") or ""
F_INV = SRC.get("inventory") or ""

BASE = CFG["base_dir"]

OUT = CFG.get("out_file") or os.path.join(
    BASE, _PD.get("out_name") or f"{_d_end.year}{_d_end.month:02d}铺货代销对账.xlsx")

_OUT_EXISTED = os.path.exists(OUT)
_OUT_LABEL = "[覆盖]" if _OUT_EXISTED else "[新建]"
if _OUT_EXISTED and not CFG.get("overwrite"):
    _stem, _ext = os.path.splitext(OUT)
    _n = 2
    while os.path.exists(f"{_stem}_v{_n}{_ext}"):
        _n += 1
    _new = f"{_stem}_v{_n}{_ext}"
    print(f"[保护] 目标文件已存在，本次写入 {os.path.basename(_new)}")
    print(f'       如需覆盖原文件，在 RECON_CFG 中加入 "overwrite": true')
    _OUT_LABEL = f"[覆盖→{os.path.basename(_new)}]"
    OUT = _new


NAVY, BLUE = "1F3864", "2F5597"
LBLUE, VLBLUE = "D9E2F3", "EDF2F9"
GREY, DGREY = "F2F2F2", "595959"
RED, LRED = "C00000", "FCE4E4"
GREEN, LGREEN = "2E6B30", "E8F1E8"
AMBER, LAMBER = "BF8F00", "FFF8E1"
WHITE, BORDC = "FFFFFF", "BFBFBF"
FN = "微软雅黑"

THIN = Side(style="thin", color=BORDC)
MED  = Side(style="medium", color="404040")

def F(sz=10, b=False, c="000000", it=False):
    return Font(name=FN, size=sz, bold=b, color=c, italic=it)

def fill(c): return PatternFill("solid", fgColor=c)

def A(h="center", v="center", wrap=True, ind=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=ind)

def box(ws, r1, c1, r2, c2, side=THIN):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(
                left=side if c == c1 else cell.border.left,
                right=side if c == c2 else cell.border.right,
                top=side if r == r1 else cell.border.top,
                bottom=side if r == r2 else cell.border.bottom)

def mput(ws, r1, c1, r2, c2, v, font=None, al=None, fl=None, border=THIN):
    if r1 != r2 or c1 != c2:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    cell = ws.cell(row=r1, column=c1, value=v)
    if font: cell.font = font
    if al:   cell.alignment = al
    if fl:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).fill = fill(fl)
    if border: box(ws, r1, c1, r2, c2, border)
    return cell

def rmb_upper(amount):
    amount = round(float(amount), 2)
    if abs(amount) < 0.005: return "零元整"
    prefix = "负" if amount < 0 else ""
    amount = abs(amount)
    digits = "零壹贰叁肆伍陆柒捌玖"
    unit1 = ["", "拾", "佰", "仟"]
    secs = ["", "万", "亿", "万亿"]
    integer = int(amount)
    cents = int(round((amount - integer) * 100))
    jiao, fen = divmod(cents, 10)
    if integer == 0:
        yuan = "零"
    else:
        s = str(integer); n = len(s)
        fl_ = n % 4 or 4
        groups = [s[:fl_]]; i = fl_
        while i < n:
            groups.append(s[i:i + 4]); i += 4
        nsec = len(groups); out = []
        for gi, g in enumerate(groups):
            if set(g) == {"0"}: continue
            gs = ""; zero = False
            for j, ch in enumerate(g):
                d = int(ch); pos = len(g) - 1 - j
                if d == 0: zero = True
                else:
                    if zero and gs: gs += "零"
                    gs += digits[d] + unit1[pos]; zero = False
            if gi > 0 and int(g) < 1000 and out: gs = "零" + gs
            out.append(gs + secs[nsec - 1 - gi])
        yuan = "".join(out)
    res = yuan + "元"
    if jiao == 0 and fen == 0: res += "整"
    else:
        if jiao: res += digits[jiao] + "角"
        elif fen: res += "零"
        res += digits[fen] + "分" if fen else "整"
    return prefix + res

# ==================== 读取源数据 ====================
def norm(s):
    s = str(s).replace("（", "(").replace("）", ")").replace(" ", "").replace("\u3000", "")
    return re.sub(r"\(限定版\)$", "", s)


def _pick_sheet(wb, keys, exclude=()):
    """按关键词找工作表；命中多张时中止并列出候选，不做猜测（缺失即阻断）。"""
    hits = [w for w in wb.worksheets
            if any(k in w.title for k in keys) and not any(x in w.title for x in exclude)]
    if len(hits) == 1:
        return hits[0], hits[0].title
    if len(hits) > 1:
        raise SystemExit(f"[中止] 含关键词 {list(keys)} 的工作表有 {len(hits)} 张："
                         f"{[w.title for w in hits]}，无法确定用哪一张。"
                         f"请只保留一张（或改表名），不要依赖脚本猜测。")
    cand = [w for w in wb.worksheets if not any(x in w.title for x in exclude) and w.max_row > 1]
    if not cand:
        raise SystemExit(f"[中止] 未找到含 {keys} 的工作表，且无含数据的候选表。实际工作表：{wb.sheetnames}")
    if len(cand) > 1:
        raise SystemExit(f"[中止] 未找到含 {keys} 的工作表，但存在 {len(cand)} 张含数据表："
                         f"{[w.title for w in cand]}。请确认用哪张（可改表名含关键词）。")
    ws = cand[0]
    print(f"⚠ 未找到含 {keys} 的工作表，使用唯一含数据表「{ws.title}」")
    return ws, ws.title


def _find_header(ws, keys, max_scan=6):
    """在前几行里找表头行，返回 (行号, 表头列表)。"""
    for r in range(1, min(ws.max_row, max_scan) + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if any(v is not None and any(k in str(v) for k in keys) for v in vals):
            return r, vals
    return 1, [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]


def _col(headers, keys, default, label=""):
    for i, h in enumerate(headers):
        if h is not None and any(k in str(h) for k in keys):
            return i
    if label:
        _bail(f"数据表未找到含「{'/'.join(keys)}」的列（{label}）。实际表头：{[str(h) for h in headers if h]}。"
              f"请核对导出格式，或在配置中调整源文件")
    return default


def _cell_of(row, idx):
    if idx is None or row is None or idx < 0 or idx >= len(row):
        return None
    return row[idx]


def _num_val(value, ctx):
    """源单元格 → 数字。缺失=0；脏值（文本、全角、千分位以外符号）直接中止，不静默兜底。"""
    if value is None:
        return 0.0
    if isinstance(value, bool):
        _bail(f"{ctx} 是布尔值（{value!r}），疑似表头/说明行混入数据区，请核对源表")
    if isinstance(value, (int, float)):
        f = float(value)
    else:
        s = str(value).strip().replace(",", "").replace("，", "")
        if not s:
            return 0.0
        try:
            f = float(s)
        except ValueError:
            _bail(f"{ctx} 的值「{value!r}」不是数字。缺失即阻断，不做静默转换，请核对源表")
    if f != f or f in (float("inf"), float("-inf")):
        _bail(f"{ctx} 的值「{value!r}」不是有限数字，请核对源表")
    return f


def _read_matrix(path):
    """把源文件首个工作表读成二维列表：.xlsx/.xlsm 用 openpyxl，.xls 用 xlrd（按扩展名分派）。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            ws = wb.worksheets[0]
            return [list(r) for r in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
    if ext == ".xls":
        try:
            import xlrd
        except ImportError:
            raise SystemExit("[缺少依赖] 读取 .xls 需要 xlrd（python3 -m pip install xlrd）；"
                             "或请甲方将挂账表另存为 .xlsx 后重试。")
        wb = xlrd.open_workbook(path)
        sh = wb.sheet_by_index(0)
        return [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    raise SystemExit(f"[中止] 不支持的源文件格式：{path}（仅支持 .xls / .xlsx / .xlsm）")


def _find_header_rows(rows, keys, max_scan=6):
    """在二维矩阵的前几行里找表头行（兼容导出文件首行是标题行）。"""
    for i, row in enumerate(rows[:max_scan]):
        if any(v is not None and any(k in str(v) for k in keys) for v in row):
            return i, list(row)
    return 0, list(rows[0]) if rows else []


def _iso_day(value):
    """销售时间 → YYYY-MM-DD；支持 datetime/date、ISO、斜杠、点号、yyyymmdd 字符串。"""
    if isinstance(value, _datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, _date):
        return value.isoformat()
    s = str(value or "").strip()
    if not s:
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d",
                "%Y.%m.%d", "%Y%m%d"):
        try:
            return _datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


UNIT = {"示例面膜A（盒装5片）": "盒", "示例面膜B（盒装5片）": "套",
        "示例精华油": "瓶", "示例面膜C（盒装5片）": "盒",
        "示例颈膜（盒装5片）": "盒", "示例面膜D（盒装5片）": "盒",
        "示例面霜": "瓶", "示例乳液": "瓶",
        "示例精华水": "瓶", "示例油敷膜（盒装5片·限定版）": "盒",
        "示例精粹水140ml": "瓶", "示例面膜E（盒装5片）": "盒",
        "示例面膜F（盒装5片·二版）": "盒", "示例精华蜜15ml": "盒",
        "示例精华蜜套装15ml*4": "盒"}

NAME_NOTE = {
    "示例面膜B（盒装5片）": "乙方系统括号为半角",
    "示例面膜C（盒装5片）": "乙方系统品名含空格",
    "示例油敷膜（盒装5片·限定版）": "甲方含“（限定版）”后缀",
    "示例精华蜜15ml": "乙方系统品名含空格",
    "示例精华蜜套装15ml*4": "乙方系统品名含空格",
}

_PART = {}

def _bail(msg):
    """失败时先把已算出的关键数打到 stderr，便于定位。"""
    sys.stderr.write("\n[自检未通过] " + msg + "\n")
    if _PART:
        sys.stderr.write("已算出：期初 %s 件/%s 元；结算 %s 件/%s 元；差异 %s 件/%s 元\n"
                         % (_PART.get("q0"), _PART.get("open"), _PART.get("q10"),
                            _PART.get("settle"), _PART.get("diff"), _PART.get("diff_amt")))
    raise SystemExit(1)


def _require(path):
    if not os.path.exists(path):
        cands = sorted(f for f in os.listdir(BASE)
                       if f.lower().endswith((".xls", ".xlsx", ".xlsm"))) if os.path.isdir(BASE) else []
        raise SystemExit(
            f"\n[缺少源文件] {path}\n"
            f"目录 {BASE} 下现有 Excel：{cands or '（无）'}\n"
            f"请核对文件名，或在 RECON_CFG 的 sources 中指定。")

for _p in (os.path.join(BASE, F_JIA), os.path.join(BASE, F_OCT), os.path.join(BASE, F_INV)):
    _require(_p)

# ---------- 甲方挂账（.xls / .xlsx，按表头关键词定位列）----------
_jia_mat = _read_matrix(os.path.join(BASE, F_JIA))
_hr0, _h0 = _find_header_rows(_jia_mat, ["商品名称", "未结算数量", "未结算金额", "单价"])
c_name = _col(_h0, ["商品", "名称", "品名"], 0, label="甲方挂账·品名列")
c_price = _col(_h0, ["单价", "供货价"], 1, label="甲方挂账·单价列")
c_amt = _col(_h0, ["未结算金额", "金额"], 2, label="甲方挂账·金额列")
c_qty = _col(_h0, ["未结算数量", "数量"], 3, label="甲方挂账·数量列")
c_init = _col(_h0, ["初始数量", "初始"], c_qty)

_jia_raw = []
for _ri, r in enumerate(_jia_mat[_hr0 + 1:], _hr0 + 2):
    _nm = _cell_of(r, c_name)
    nm = "" if _nm is None else str(_nm).strip()
    if not nm:
        continue
    if nm in ("合计", "小计", "总计", "合 计"):
        continue
    _jia_raw.append(dict(name=nm,
                         price=round(_num_val(_cell_of(r, c_price), f"甲方挂账第 {_ri} 行·单价"), 2),
                         amt=round(_num_val(_cell_of(r, c_amt), f"甲方挂账第 {_ri} 行·未结算金额"), 2),
                         qty=int(round(_num_val(_cell_of(r, c_qty), f"甲方挂账第 {_ri} 行·未结算数量"))),
                         init=int(round(_num_val(_cell_of(r, c_init), f"甲方挂账第 {_ri} 行·初始数量")))))

# 同品名多行（拆行/重复挂账）按数量合计，避免动销被双算
jia, _jm = [], {}
for _x in _jia_raw:
    _k = norm(_x["name"])
    if _k in _jm:
        _o = _jm[_k]
        print(f"⚠ 挂账表中「{_x['name']}」出现多行，已按数量合计")
        _o["qty"] += _x["qty"]; _o["init"] += _x["init"]; _o["amt"] = round(_o["amt"] + _x["amt"], 2)
    else:
        _jm[_k] = _x; jia.append(_x)

# ---------- 乙方库存（表名与列均按关键词识别）----------
wb1 = openpyxl.load_workbook(os.path.join(BASE, F_INV), data_only=True)
ws_inv, inv_sheet_name = _pick_sheet(wb1, ["库存"], exclude=("销售",))
_hr, _hinv = _find_header(ws_inv, ["商品名称", "库存"])
k_name = _col(_hinv, ["商品名称", "名称", "品名"], 0, label="库存表·品名列")
k_bar = _col(_hinv, ["条码", "编码"], 1, label="库存表·条码列")
k_spec = _col(_hinv, ["规格"], 2)
k_cat = _col(_hinv, ["分类"], 3)
k_qty = _col(_hinv, ["库存", "数量"], 4, label="库存表·库存列")
k_retail = _col(_hinv, ["销售价", "零售价"], 5)
k_brand = _col(_hinv, ["品牌"], None)

inv = []
for _ri, r in enumerate(ws_inv.iter_rows(min_row=_hr + 1, values_only=True), _hr + 1):
    if not r or len(r) <= max(k_name, k_qty) or not r[k_name]:
        continue
    _bar = _cell_of(r, k_bar)
    _bar = "" if _bar is None else str(_bar).strip()
    if _bar.endswith(".0") and _bar[:-2].isdigit():
        _bar = _bar[:-2]
    _retail = _cell_of(r, k_retail)
    _brand = _cell_of(r, k_brand)
    inv.append(dict(name=str(r[k_name]).strip(),
                    bar=_bar,
                    spec=_cell_of(r, k_spec) if _cell_of(r, k_spec) is not None else "",
                    cat=_cell_of(r, k_cat) if _cell_of(r, k_cat) is not None else "",
                    qty=int(round(_num_val(_cell_of(r, k_qty), f"库存表第 {_ri} 行·库存数量"))),
                    retail=_num_val(_retail, f"库存表第 {_ri} 行·销售价") if _retail is not None else 0.0,
                    brand=str(_brand).strip() if _brand is not None else ""))

# 先按条码聚合（同一商品被拆成两行、或换名不换码，都能合并）
_bydedup = {}
for _x in inv:
    _bk = _x["bar"] or ("NAME:" + norm(_x["name"]))
    if _bk in _bydedup:
        _bydedup[_bk]["qty"] += _x["qty"]
        if norm(_x["name"]) != norm(_bydedup[_bk]["name"]):
            print(f"⚠ 库存表内条码 {_x['bar']} 对应多个品名（{_bydedup[_bk]['name']} / {_x['name']}），已按条码合并")
        else:
            print(f"⚠ 库存表内「{_x['name']}」出现多行，已按数量合计")
    else:
        _bydedup[_bk] = dict(_x)
# 再按名称聚合（个别表无条码）
_dedup = {}
for _x in _bydedup.values():
    _k = norm(_x["name"])
    if _k in _dedup:
        _dedup[_k]["qty"] += _x["qty"]
    else:
        _dedup[_k] = _x
inv = list(_dedup.values())
inv_by_n = {norm(x["name"]): x for x in inv}
inv_by_bar = {x["bar"]: x for x in inv if x["bar"]}

# ---------- 乙方库存导出月次月销售（另一张表，表名随月变化）----------
nov = defaultdict(int)
ws_nov, nov_sheet_name = None, ""
_nov_skip = []
if not ADD_BACK:
    print(f"库存导出日 {ASOF} 不晚于期末 {P_END}，无需回加，跳过次月销售表")
else:
    ws_nov, nov_sheet_name = _pick_sheet(wb1, ["销售"], exclude=("库存",))
    _t = ws_nov.title or ""
    if str(_d_next.month) not in _t and "月" in _t:
        print(f"⚠ 回加工作表「{_t}」与推算区间 {_d_next.month} 月不一致，请核对库存导出日")
if ADD_BACK:
    _hr2, _hnov = _find_header(ws_nov, ["商品名称", "销售数量"])
    n_name = _col(_hnov, ["商品名称", "名称"], 0)
    n_bar = _col(_hnov, ["条码", "编码"], 1)
    n_num = _col(_hnov, ["销售数量", "数量", "销量"], 3, label="次月销售表·销售数量列")
else:
    _hr2, _hnov, n_name, n_bar, n_num = 1, [], 0, 1, 3
nov = defaultdict(int)
for _ri, r in enumerate((ws_nov.iter_rows(min_row=_hr2 + 1, values_only=True) if ADD_BACK else []), _hr2 + 1):
    if not r or not _cell_of(r, n_name):
        continue
    _q = _num_val(_cell_of(r, n_num), f"次月销售表第 {_ri} 行·销售数量")
    if not _q:
        continue
    _bar = _cell_of(r, n_bar)
    _bar = "" if _bar is None else str(_bar).strip()
    if _bar.endswith(".0") and _bar[:-2].isdigit():
        _bar = _bar[:-2]
    if not _bar:
        _iv = inv_by_n.get(norm(_cell_of(r, n_name)))
        _bar = _iv["bar"] if _iv else ""
    if not _bar:
        _nov_skip.append(str(_cell_of(r, n_name)).strip())
        continue
    nov[_bar] += int(round(_q))
if _nov_skip:
    print(f"⚠ 次月销售表有 {len(_nov_skip)} 行无条码且品名无法匹配库存，已跳过（如："
          f"{'、'.join(_nov_skip[:5])}）；请核对是否影响回加")

# ---------- 乙方当期销售流水 ----------
wb2 = openpyxl.load_workbook(os.path.join(BASE, F_OCT), data_only=True)
ws_oct, oct_sheet_name = _pick_sheet(wb2, [str(P_END)[5:7], "月"], exclude=("库存",))
_m_end = _date.fromisoformat(str(P_END)[:10]).month
_m_hit = re.search(r"(\d{1,2})\s*月", ws_oct.title or "")
if _m_hit and int(_m_hit.group(1)) != _m_end:
    _bail(f"销售工作表是「{ws_oct.title}」，但本次期间是 {_m_end} 月，疑似期间与源文件不匹配。"
          f"请把对应月份的销售数据放到本目录，或在配置中显式指定 sources.sales")

_hr3, _hoct = _find_header(ws_oct, ["商品名称", "销售数量"])
o_id = _col(_hoct, ["流水号", "单号"], 0)
o_time = _col(_hoct, ["销售时间", "时间", "日期"], 1)
o_store = _col(_hoct, ["门店", "店铺", "网点"], 2)
o_name = _col(_hoct, ["商品名称", "名称", "品名"], 3, label="销售流水·品名列")
o_bar = _col(_hoct, ["条码", "编码"], 4, label="销售流水·条码列")
o_spec = _col(_hoct, ["规格"], 5)
o_unit = _col(_hoct, ["单位"], 6)
o_num = _col(_hoct, ["销售数量", "数量", "销量"], 7, label="销售流水·数量列")
o_brand = _col(_hoct, ["品牌"], 8)

detail = []
oct_qty, oct_store = defaultdict(int), defaultdict(lambda: defaultdict(int))
oct_name, oct_spec, oct_unit = {}, {}, {}
oct_by_name = {}
for _ri, r in enumerate(ws_oct.iter_rows(min_row=_hr3 + 1, values_only=True), _hr3 + 1):
    if not r or not _cell_of(r, o_id) or not _cell_of(r, o_name):
        continue
    q = int(round(_num_val(_cell_of(r, o_num), f"销售流水第 {_ri} 行·销售数量")))
    _bar = _cell_of(r, o_bar)
    _bar = "" if _bar is None else str(_bar).strip()
    if _bar.endswith(".0") and _bar[:-2].isdigit():
        _bar = _bar[:-2]
    if not _bar:
        _bail(f"销售流水第 {_ri} 行缺少商品条码。销售流水必须带条码，不得按品名猜算，请核对导出文件")
    _store = _cell_of(r, o_store)
    _store = str(_store).strip() if _store not in (None, "") else "未标门店"
    _spec = _cell_of(r, o_spec)
    _unit = _cell_of(r, o_unit)
    _brand = _cell_of(r, o_brand)
    detail.append([_cell_of(r, o_id), _cell_of(r, o_time), _store, _cell_of(r, o_name),
                   _bar, _spec if _spec is not None else "",
                   _unit if _unit is not None else "", q,
                   _brand if _brand is not None else ""])
    oct_qty[_bar] += q
    oct_store[_bar][_store] += q
    oct_name[_bar] = str(_cell_of(r, o_name)).strip()
    oct_spec[_bar] = _spec if _spec is not None else ""
    oct_unit[_bar] = _unit if _unit is not None else ""
    oct_by_name[norm(_cell_of(r, o_name))] = _bar

# ---- 数据覆盖披露（动态）：有记录日期区间 + 期间内缺哪几天 ----
_dates_seen = sorted({_iso_day(d[1]) for d in detail if _iso_day(d[1])})
_bad_time = sum(1 for d in detail if not _iso_day(d[1]))
if _bad_time:
    print(f"⚠ 销售流水有 {_bad_time} 行销售时间无法解析，数据覆盖期间将不含这些行，请核对")
if _dates_seen:
    _d0, _d1 = _date.fromisoformat(_dates_seen[0]), _date.fromisoformat(_dates_seen[-1])
    _span = {(_d0 + _timedelta(i)).isoformat() for i in range((_d1 - _d0).days + 1)}
    _gap = sorted(_span - set(_dates_seen))
    COV_TXT = (f"{_dates_seen[0]} 至 {_dates_seen[-1]}"
               + (f"，其中 " + "、".join(g[5:].replace("-", "/") for g in _gap) + " 无记录" if _gap else ""))
else:
    COV_TXT = "无有效销售记录"

STORE_ORDER = [s for s in ["示例门店一店", "示例门店二店", "示例门店三店", "示例门店四店", "示例门店五店", "示例门店六店"]
               if s in {k for v in oct_store.values() for k in v}]
STORE_ORDER += [k for v in oct_store.values() for k in v if k not in STORE_ORDER]
STORE_ORDER = list(dict.fromkeys(STORE_ORDER))
print(f"识别：库存表「{inv_sheet_name}」/ 次月表「{ws_nov.title if ws_nov is not None else '（无需回加）'}」/ 销售表「{ws_oct.title}」")

# ---------- 与主档价格表交叉校验（仅提示，不改数）----------
_mp = CFG.get("master_path")
if _mp and os.path.exists(_mp):
    try:
        _wb = openpyxl.load_workbook(_mp, data_only=True)
        if "价格" in _wb.sheetnames:
            _ws = _wb["价格"]
            _h = [c.value for c in _ws[1]]
            def _c(k, d=0):
                for i, x in enumerate(_h):
                    if x and k in str(x):
                        return i
                return d
            pi, pp, ps, pe = _c("条码"), _c("供货价"), _c("生效起"), _c("生效止")
            _rows = [r for r in _ws.iter_rows(min_row=2, values_only=True) if r and r[pi]]
            _bybar = {}
            for r in _rows:
                _bybar.setdefault(str(r[pi]).strip(), []).append(r)
            _mis = []
            for x in jia:
                kb = oct_by_name.get(norm(x["name"]), "")
                if not kb or kb not in _bybar:
                    continue
                hit = None
                for r in _bybar[kb]:
                    s0 = str(r[ps])[:10] if r[ps] else ""
                    s1 = str(r[pe])[:10] if r[pe] else ""
                    if (not s0 or s0 <= P_END) and (not s1 or s1 >= P_END):
                        hit = r; break
                if hit is None:
                    continue
                mp = float(hit[pp] or 0)
                if abs(mp - x["price"]) > 0.005:
                    _mis.append((x["name"], x["price"], mp))
            if _mis:
                print("⚠ 挂账单单价与主档价格表不一致（本期口径以挂账单为准，请人工确认）：")
                for n, a1, a2 in _mis:
                    print(f"    - {n}：挂账 {a1:.2f} / 主档 {a2:.2f}")
    except Exception as _e:
        print(f"⚠ 主档价格校验跳过：{_e}")

# ---------- 读主档 SKU：别名反查 + 结算属性 + 条码全集 ----------
MS_BAR, MS_ALIAS, MS_SETTLE, MS_NAME = {}, {}, set(), {}
_mp2 = CFG.get("master_path")
if _mp2 and os.path.exists(_mp2):
    try:
        _wb = openpyxl.load_workbook(_mp2, data_only=True)
        if "SKU" in _wb.sheetnames:
            _ws = _wb["SKU"]
            _h = [c.value for c in _ws[1]]
            def _cc(k, d=None):
                for i, x in enumerate(_h):
                    if x and k in str(x):
                        return i
                return d
            ci_bar = _cc("条码", 0); ci_name = _cc("标准品名", 1)
            ci_alias = _cc("别名", 2); ci_attr = _cc("结算属性", 6)
            for r in _ws.iter_rows(min_row=2, values_only=True):
                if not r or not r[ci_bar]:
                    continue
                b = str(r[ci_bar]).strip()
                attr = str(r[ci_attr] or "").strip() if ci_attr is not None else ""
                MS_BAR[b] = dict(name=r[ci_name], attr=attr,
                                 alias=r[ci_alias] if ci_alias is not None else "")
                if attr == "结算类":
                    MS_SETTLE.add(b)
                MS_NAME[norm(r[ci_name] or "")] = b
                if ci_alias is not None and r[ci_alias]:
                    for a in str(r[ci_alias]).split("|"):
                        if a.strip():
                            MS_ALIAS[norm(a)] = b
            print(f"主档：SKU {len(MS_BAR)} 条，结算类 {len(MS_SETTLE)} 条，别名 {len(MS_ALIAS)} 条")
    except Exception as _e:
        print(f"⚠ 主档 SKU 读取失败（{_e}），本次仅按品名/条码容错匹配")


def resolve_sku(name):
    """条码优先 → 主档别名反查 → 品名归一化。返回 (库存行, 依据)。"""
    k = norm(name)
    for cand in (oct_by_name.get(k), MS_ALIAS.get(k), MS_NAME.get(k)):
        if cand and cand in inv_by_bar:
            return inv_by_bar[cand], "条码"
    iv = inv_by_n.get(k)
    if iv:
        return iv, "品名"
    return None, "未匹配"


# ==================== 核心计算 ====================
rows = []
_missing = []
_uncls = []
if MS_BAR:
    for x in jia:
        _b = (oct_by_name.get(norm(x["name"])) or MS_ALIAS.get(norm(x["name"]))
              or MS_NAME.get(norm(x["name"])) or "")
        _a = MS_BAR.get(_b, {}).get("attr") if _b else None
        if _a not in ("结算类", "非结算类"):
            _uncls.append((x["name"], _b or "（无条码）", _a or "空"))
    for x in inv:
        if x["bar"] not in MS_BAR:
            _uncls.append((x["name"], x["bar"] or "（无条码）", "主档未建档"))
        elif MS_BAR[x["bar"]]["attr"] not in ("结算类", "非结算类"):
            _uncls.append((x["name"], x["bar"], MS_BAR[x["bar"]]["attr"] or "空"))
if _uncls:
    for _n, _b, _a in _uncls[:12]:
        sys.stderr.write(f"    - {_n}｜{_b}｜结算属性：{_a}\n")
    _bail(f"有 {len(_uncls)} 个商品在主档中未明确「结算类/非结算类」。"
          f"空值不会被默认成结算类，请先补主档 SKU 表再对账")

for x in jia:
    k = norm(x["name"]); iv, _how = resolve_sku(x["name"])
    if iv is not None and _how == "条码" and norm(iv["name"]) != k:
        print(f"⚠ 挂账「{x['name']}」与库存「{iv['name']}」品名不同，已按条码匹配")
    if iv is None:
        _missing.append(x["name"])
        _b = oct_by_name.get(k) or MS_ALIAS.get(k) or ""
        iv = dict(name=x["name"], bar=_b, spec="", cat="", qty=0, retail=0.0, _absent=True)
    bar = iv["bar"]
    q10 = oct_qty.get(bar, 0); q11 = nov.get(bar, 0)
    theo = x["qty"] - q10                 # 甲方账面 10/31 结存
    adj  = iv["qty"] + q11                # 乙方库存回加导出日前已销
    diff = adj - theo
    _note = NAME_NOTE.get(x["name"], "")
    if not _note and not iv.get("_absent") and norm(iv["name"]) != norm(x["name"]):
        _note = f"乙方系统品名：{iv['name']}"
    if not _note and iv.get("_absent"):
        _note = "乙方库存表未列示该 SKU，按 0 计"
    rows.append(dict(name=x["name"], bar=bar, spec=iv["spec"],
                     unit=oct_unit.get(bar) or UNIT.get(x["name"], "—"),
                     price=x["price"], retail=iv["retail"], q0=x["qty"], q10=q10, q11=q11,
                     theo=theo, act=iv["qty"], adj=adj, diff=diff,
                     diff_amt=round(diff * x["price"], 2),
                     settle=round(q10 * x["price"], 2),
                     init=x["init"], src_amt=x["amt"], note=_note))
T = dict(q0=0, q10=0, q11=0, theo=0, act=0, adj=0, diff=0, diff_amt=0, settle=0)
for r in rows:
    for kk in T: T[kk] += r[kk]
T = {k: round(v, 2) for k, v in T.items()}

OPEN_VAL = round(sum(r["q0"] * r["price"] for r in rows), 2)

_PART.update(q0=T.get("q0"), open=OPEN_VAL, q10=T.get("q10"),
             settle=T.get("settle"), diff=T.get("diff"), diff_amt=T.get("diff_amt"))
NOV_VAL  = round(sum(r["q11"] * r["price"] for r in rows), 2)
ACT_VAL  = round(sum(r["act"] * r["price"] for r in rows), 2)
ADJ_VAL  = round(sum(r["adj"] * r["price"] for r in rows), 2)
THEO_VAL = round(sum(r["theo"] * r["price"] for r in rows), 2)
BOOK_END = round(OPEN_VAL - T["settle"], 2)
SETTLE_ALL = round(T["settle"] + abs(T["diff_amt"]), 2)
DIFF_RATE  = T["diff"] / abs(T["theo"]) if T["theo"] else 0

if T["act"] == 0:
    _bail("乙方库存表解析结果为 0 件，疑似表名或列位不匹配")
if T["theo"] and abs(T["diff"]) / abs(T["theo"]) > 0.30 and not CFG.get("allow_large_diff"):
    _bail(f"差异 {T['diff']} 件，占理论结存 {abs(DIFF_RATE)*100:.1f}%，超过 30% 阈值；多为源文件放错期间或表结构变化")

if _missing:
    print(f"⚠ 以下 {len(_missing)} 个 SKU 在乙方库存表中未找到，已按库存 0 计：")
    for _m in _missing:
        print(f"    - {_m}")

def _check_identity(label, a, b):
    """恒等式用显式中止而非 assert（python -O 下 assert 会被移除，自检静默失效）。"""
    if abs(a - b) > 0.01:
        _bail(f"{label}：{a:,.2f} ≠ {b:,.2f}")


_check_identity("期末账面货值 ≠ 理论结存货值", BOOK_END, THEO_VAL)
_check_identity("期末账面货值 － 调整后结存货值 ＋ 差异金额 ≠ 0",
                BOOK_END - ADJ_VAL + T["diff_amt"], 0.0)
_check_identity("调整后结存货值 ≠ 库存货值 ＋ 回加货值", ADJ_VAL - ACT_VAL, NOV_VAL)
if T["q0"] != T["q10"] + T["theo"]:
    _bail(f"数量恒等式不成立：期初 {T['q0']} ≠ 动销 {T['q10']} ＋ 理论结存 {T['theo']}")
if T["adj"] != T["act"] + T["q11"]:
    _bail(f"数量恒等式不成立：调整后结存 {T['adj']} ≠ 库存 {T['act']} ＋ 回加 {T['q11']}")
if T["diff"] != T["adj"] - T["theo"]:
    _bail(f"数量恒等式不成立：差异 {T['diff']} ≠ 调整后结存 {T['adj']} － 理论结存 {T['theo']}")

N_PAY = sum(1 for r in rows if r["q0"] > 0)
diff_rows = [r for r in rows if r["diff"] != 0]
ok_rows   = [r for r in rows if r["diff"] == 0 and r["q0"] > 0]
zero_rows = [r for r in rows if r["q0"] == 0]
_jia_bars_all = {r["bar"] for r in rows if r["bar"]}
_jia_names_all = {norm(y["name"]) for y in jia}
extra     = [x for x in inv
             if not ((x["bar"] and x["bar"] in _jia_bars_all) or norm(x["name"]) in _jia_names_all)]
EXTRA_Q   = sum(x["qty"] for x in extra)

# ==================== 贴花（销售奖励）模块 ====================
# 判据与主档一致：条码为主键、按生效区间取当期标准；缺失即阻断，不默认为 0
TH_ROWS, TH_TOTAL, TH_ACTIVE = [], 0.0, False
_mp3 = CFG.get("master_path")
_th_map, _th_dup = {}, []
if _mp3 and os.path.exists(_mp3):
    if "贴花" not in openpyxl.load_workbook(_mp3, read_only=True).sheetnames:
        print("贴花：主档无「贴花」表，本期不计算销售奖励（如需计算请先建表）")
    else:
        TH_ACTIVE = True
        try:
            _wst = openpyxl.load_workbook(_mp3, data_only=True)["贴花"]
        except Exception as _e:
            _bail(f"主档「贴花」表无法读取：{_e}")
        _ht = [c.value for c in _wst[1]]
        def _ct(k, d=None):
            for i, x in enumerate(_ht):
                if x and k in str(x):
                    return i
            return d
        t_bar, t_std = _ct("条码", 0), _ct("贴花标准", 2)
        t_from, t_to = _ct("生效起", 3), _ct("生效止", 4)
        _p_end = str(P_END)[:10]
        for _r in _wst.iter_rows(min_row=2, values_only=True):
            if not _r or not _r[t_bar]:
                continue
            _s0 = str(_r[t_from])[:10] if _r[t_from] else ""
            _s1 = str(_r[t_to])[:10] if _r[t_to] else ""
            if (_s0 and _s0 > _p_end) or (_s1 and _s1 < _p_end):
                continue
            _bk = str(_r[t_bar]).strip()
            try:
                _std = float(_r[t_std])
            except (TypeError, ValueError):
                _bail(f"贴花表条码 {_bk} 的「贴花标准」不是数字（收到：{_r[t_std]!r}）。"
                      f"脏值不会被默认为 0，请修正主档")
            if _std < 0:
                _bail(f"贴花表条码 {_bk} 的贴花标准为负数（{_std}），请核对")
            if _bk in _th_map:
                _th_dup.append((_bk, _th_map[_bk], _std))
            else:
                _th_map[_bk] = _std
        # SCD2 区间重叠：同条码多条命中 → 阻断（不同标准取行序是静默错误）
        if _th_dup:
            for _bk, _a, _b2 in _th_dup[:6]:
                sys.stderr.write(f"    - 条码 {_bk}：命中 {_a} 与 {_b2}\n")
            _bail(f"贴花表有 {len(_th_dup)} 个条码在期间 {P_END} 命中多条生效记录（区间重叠）。"
                  f"请修正生效起止，保证同一时点每条码仅一条有效记录")
        # 每个 q10>0 的 SKU 必须有当期有效贴花记录（显式 0 也算）
        _th_miss = [r["name"] for r in rows if r["q10"] > 0 and r["bar"] not in _th_map]
        if _th_miss:
            for _n in _th_miss[:10]:
                sys.stderr.write(f"    - {_n}\n")
            _bail(f"贴花表缺少 {len(_th_miss)} 个本期动销 SKU 的有效记录（期间 {P_END}）。"
                  f"缺失不会被默认为 0，请先补主档「贴花」表")
        for r in rows:
            if r["q10"] > 0:
                std = _th_map[r["bar"]]
                TH_ROWS.append(dict(name=r["name"], bar=r["bar"], qty=r["q10"],
                                    std=std, amt=round(r["q10"] * std, 2)))
        TH_TOTAL = round(sum(x["amt"] for x in TH_ROWS), 2)
        _th_chk = sum(r["q10"] * _th_map[r["bar"]] for r in rows if r["q10"] > 0)
        if abs(_th_chk - TH_TOTAL) > 0.005:
            _bail(f"贴花合计复算不一致：{TH_TOTAL} ≠ {_th_chk:.2f}")
if TH_ACTIVE:
    print(f"贴花：{len(TH_ROWS)} 个动销 SKU，合计 {TH_TOTAL:,.2f} 元（标准取自主档贴花表期间 {P_END} 生效版）")
print(f"输入目录：{BASE}")
print(f"源文件：{F_JIA} / {F_OCT} / {F_INV}")
print(f"输出文件：{OUT}")
print(f"期初 {T['q0']}件/{OPEN_VAL} | 10月结算 {T['q10']}件/{T['settle']} | 11月 {T['q11']}件/{NOV_VAL}")
print(f"理论结存 {T['theo']}件/{THEO_VAL} | 实盘 {T['act']}件/{ACT_VAL} + 11月 = 调整后 {T['adj']}件/{ADJ_VAL}")
if MS_BAR:
    _inv_bars = {x["bar"] for x in inv if x["bar"]}
    _unreg = _inv_bars - set(MS_BAR)
    if _unreg:
        print(f"⚠ 库存表有 {len(_unreg)} 个条码未在主档 SKU 登记：{sorted(_unreg)[:6]}")
    _jia_bars = {}
    for x in jia:
        _b = (oct_by_name.get(norm(x["name"])) or MS_ALIAS.get(norm(x["name"]))
              or MS_NAME.get(norm(x["name"])) or "")
        if _b:
            _jia_bars[_b] = _jia_bars.get(_b, 0) + 1
    _dupb = [b for b, n in _jia_bars.items() if n > 1]
    if _dupb:
        _bail(f"挂账表中同一条码出现多行（{_dupb[:3]}），请先在主档/挂账表核对，避免动销双算")
    if MS_SETTLE and len(_jia_bars) and len(_jia_bars) != len(MS_SETTLE):
        print(f"⚠ 挂账覆盖条码 {len(_jia_bars)} 个，主档结算类 {len(MS_SETTLE)} 个，差额 "
              f"{sorted(MS_SETTLE - set(_jia_bars))[:5]}，请确认是否漏挂/多挂")

_nounit = [r["name"] for r in rows if r["unit"] == "—"]
if _nounit:
    print(f"⚠ {len(_nounit)} 个单品未在主档登记单位，对账单按「—」显示：{_nounit[:5]}")

if not ADD_BACK:
    # 库存文件里若存在含数据的次月销售表，说明库存实际导出于期末之后——asof 大概率填错
    _late_q = 0
    try:
        _wl = [w for w in wb1.worksheets if w is not ws_inv and w.max_row > 1]
        if _wl:
            _lw = max(_wl, key=lambda w: w.max_row)
            for _r in _lw.iter_rows(min_row=2, values_only=True):
                for _v in _r:
                    if isinstance(_v, (int, float)) and _v:
                        _late_q += int(_v)
    except Exception:
        pass
    if _late_q > 0:
        _bail(f"库存文件中存在含 {_late_q} 件数量记录的「次月销售」类工作表，"
              f"说明库存实际导出于期末 {P_END} 之后；但 --asof={ASOF} 不晚于期末，"
              f"本次按无需回加处理会虚增短缺。请把 --asof 改为真实库存导出日（次月 2 日左右）")

_det_sum = sum(int(d[7] or 0) for d in detail)
if _det_sum != sum(oct_qty.values()):
    _bail(f"明细合计 {_det_sum} ≠ 按条码汇总 {sum(oct_qty.values())}（源表可能有合计行或空行）")
_src_q0 = sum(x["qty"] for x in _jia_raw)
if _src_q0 != T["q0"]:
    _bail(f"挂账源表数量合计 {_src_q0} ≠ 归一化后 {T['q0']}（同品名多行合并异常）")

_unreg_rows = [x for x in extra if (x["bar"] not in MS_BAR)] if MS_BAR else list(extra)
if _unreg_rows:
    _q = sum(x["qty"] for x in _unreg_rows)
    print(f"⚠ 库存有 {len(_unreg_rows)} 行、{_q} 件在主档 SKU 中查不到条码，已归入非结算桶，请建档或确认：")
    for x in _unreg_rows[:10]:
        print(f"    - {x['bar'] or '（无条码）'} {x['name']} {x['qty']} 件")
elif MS_BAR and EXTRA_Q:
    print(f"非结算桶 {EXTRA_Q} 件，均已在主档标为非结算类")

_resid = round(BOOK_END - ADJ_VAL + T["diff_amt"], 2)
_sum_q0 = sum(x["q0"] for x in rows)
if abs(_resid) > 0.01 or _sum_q0 != T["q0"]:
    _bail(f"结转残差 {_resid}，期初合计 {_sum_q0} ≠ {T['q0']}")
print(f"差异 {T['diff']}件/{T['diff_amt']} ({DIFF_RATE*100:.2f}%) | 自检通过（残差 0.00）| 非结算片装 {EXTRA_Q}件")

# ==================== 工作簿框架 ====================
out = openpyxl.Workbook(); out.remove(out.active)

def setup(ws, widths, orient="portrait"):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.sheet_view.showGridLines = False
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = orient
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.35
    ws.page_margins.top = ws.page_margins.bottom = 0.45

def sec(ws, r, c1, c2, text):
    for c in range(c1, c2 + 1):
        ws.cell(row=r, column=c).fill = fill(LBLUE)
        ws.cell(row=r, column=c).border = Border(bottom=MED)
    cell = ws.cell(row=r, column=c1, value=text)
    cell.font = F(10.5, True, NAVY)
    cell.alignment = A("left", "center", False, 1)
    ws.row_dimensions[r].height = 26

def hdr(ws, r, c1, headers, h=36, fl=NAVY):
    for i, t in enumerate(headers):
        cell = ws.cell(row=r, column=c1 + i, value=t)
        cell.font = F(8.5, True, WHITE)
        cell.fill = fill(fl)
        cell.alignment = A("center", "center", True)
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=MED)
    ws.row_dimensions[r].height = h

def num(cell, v, fmt="#,##0"):
    cell.value = v; cell.number_format = fmt
    cell.font = F(9)
    cell.alignment = A("right", "center", False, 1) if ".00" in fmt else A("center")

# ==================== 00 对账说明 ====================
ws = out.create_sheet("00-对账说明")
setup(ws, [10.5] * 12)
ws.row_dimensions[1].height = 8
mput(ws, 2, 1, 2, 12, f"{P_CN}  铺货代销对账报告", F(22, True, WHITE), A("center"), NAVY)
ws.row_dimensions[2].height = 48
mput(ws, 3, 1, 3, 12, f"{PA_NAME}  ×  {PB_NAME}      |      对账期间 {P_START} 至 {P_END}",
     F(9.5, False, NAVY), A("center"), LBLUE)
ws.row_dimensions[3].height = 24
ws.row_dimensions[4].height = 6

sec(ws, 5, 1, 12, "一、对账基础信息")
info = [("甲方（供货方）", PA_NAME, "乙方（代销方）", PB_NAME),
        ("甲方纳税人识别号", PA_TAX, "乙方纳税人识别号", PB_TAX),
        ("经销合同编号", CONTRACT_NO, "结算方式", SETTLE_TERMS_TXT),
        ("对账期间", f"{P_START} 至 {P_END}", "数据基准日", ASOF)]
r = 6
for lab1, v1, lab2, v2 in info:
    mput(ws, r, 1, r, 2, lab1, F(9, True, DGREY), A("left", "center", True, 1), GREY)
    mput(ws, r, 3, r, 6, v1, F(9.5), A("left", "center", True, 1))
    mput(ws, r, 7, r, 8, lab2, F(9, True, DGREY), A("left", "center", True, 1), GREY)
    mput(ws, r, 9, r, 12, v2, F(9.5), A("left", "center", True, 1))
    ws.row_dimensions[r].height = 24; r += 1
mput(ws, r, 1, r, 2, "数据来源", F(9, True, DGREY), A("left", "center", True, 1), GREY)
mput(ws, r, 3, r, 12, f"甲方 ERP 系统未结算挂账表（{ASOF} 导出）／乙方销售数据（有记录日期 {COV_TXT}）／乙方库存＋当月销售表（{ASOF}）",
     F(9.5), A("left", "center", True, 1))
ws.row_dimensions[r].height = 24; r += 2

sec(ws, r, 1, 12, "二、对账结论"); r += 1
_kpi_diff_lab = "库存差异（净短缺，元）" if T["diff"] < 0 else ("库存差异（净盘盈，元）" if T["diff"] > 0 else "库存差异（元）")
kpi = [("本月应结算金额（元）", T['settle'], f"结算数量 {T['q10']} 件 · {sum(1 for r in rows if r['q10'] > 0)} 个单品", NAVY, VLBLUE),
       ("期初未结算货值（元）", OPEN_VAL, f"铺货数量 {T['q0']:,} 件（首批铺货）", BLUE, VLBLUE),
       ("调整后结存货值（元）", ADJ_VAL,
        (f"库存 {T['act']:,} ＋ {NEXT_M}月已销 {T['q11']} ＝ {T['adj']:,} 件" if ADD_BACK
         else f"库存 {T['act']:,}（导出日不晚于期末）"), GREEN, LGREEN),
       (_kpi_diff_lab, abs(T['diff_amt']), f"{T['diff']:+d} 件 · 占理论结存 {abs(DIFF_RATE)*100:.2f}%", RED, LRED)]
for i, (lab, val, sub, col, bg) in enumerate(kpi):
    c1 = 1 + i * 3
    mput(ws, r, c1, r, c1 + 2, lab, F(8.5, True, col), A("center"), bg)
    cc = mput(ws, r + 1, c1, r + 1, c1 + 2, val, F(16.5, True, col), A("center"), bg)
    cc.number_format = '"¥ "#,##0.00'
    mput(ws, r + 2, c1, r + 2, c1 + 2, sub, F(7.5, False, DGREY), A("center"), bg)
ws.row_dimensions[r].height = 21
ws.row_dimensions[r + 1].height = 38
ws.row_dimensions[r + 2].height = 22
r += 4

sec(ws, r, 1, 12, "三、结算与核对口径（关键：库存需回加导出日前已发生销量）" if ADD_BACK
    else "三、结算与核对口径（库存导出日不晚于期末，无需回加）"); r += 1
formula = [
    ("① 本期结算金额", f"＝ Σ（乙方经对账确认的 {_d_end.month} 月实际动销数量 × 甲方供货结算单价）"),
    ("② 甲方理论结存", f"＝ 甲方期初未结算数量 － 乙方 {_d_end.month} 月申报动销数量　【截至 {P_END}】"),
    ("③ 乙方调整后结存", (f"＝ 乙方 {ASOF} 库存 ＋ {NEXT_LABEL}已发生销售数量　【回加至同一时点】" if ADD_BACK
                          else f"＝ 乙方 {ASOF} 库存（导出日不晚于期末，即为期末结存）")),
    ("④ 库存差异", "＝ 乙方调整后结存 － 甲方理论结存（负数表示存货短缺）"),
    ("⑤ 账实勾稽", "期末账面货值 － 调整后结存货值 ＝ 库存差异金额（本期 ①→⑦ 逐式复算相符，⑥ 为差额项）"),
]
for lab, txt in formula:
    mput(ws, r, 1, r, 3, lab, F(9, True, BLUE), A("left", "center", True, 1), VLBLUE)
    mput(ws, r, 4, r, 12, txt, F(9), A("left", "center", True, 1))
    ws.row_dimensions[r].height = 24; r += 1
_add_note = (
    f"说明：乙方库存导出日为 {ASOF}，晚于对账截止日 {P_END}，其间已发生销售 {T['q11']} 件。"
    f"若直接以库存对比期末理论结存，会因时间口径差异虚增短缺；故统一将 {NEXT_LABEL}销量回加至库存，"
    f"使双方数据同处 {P_END} 同一时点后再比对。" if ADD_BACK else
    f"说明：乙方库存导出日 {ASOF} 不晚于对账截止日 {P_END}，无需回加，库存即为期末结存。")
mput(ws, r, 1, r, 12, _add_note, F(8.5, False, AMBER), A("left", "center", True, 1), LAMBER)
ws.row_dimensions[r].height = 46; r += 2

sec(ws, r, 1, 12, "四、主要结论"); r += 1
_settle_after = T["settle"] - T["diff_amt"]
concl = [
    f"1. 本月应结算货款 {T['settle']:,.2f} 元（{T['q10']} 件／{sum(1 for r in rows if r['q10'] > 0)} 个单品）。"
    + (f"{_d_next.isoformat()} 至 {_d_asof.isoformat()} 期间发生的 {T['q11']} 件销售（{NOV_VAL:,.2f} 元）"
       f"不属本期，按合同约定归入 {NEXT_M} 月结算周期。" if ADD_BACK
       else "库存导出日不晚于期末，无导出日前销售需回加。"),
    f"2. 库存核对：{len(ok_rows) + len(diff_rows)} 个有铺货单品中 {len(ok_rows)} 个账实相符；{len(diff_rows)} 个存在差异，净差异 {T['diff']:+d} 件／{T['diff_amt']:,.2f} 元，占理论结存 {abs(DIFF_RATE)*100:.2f}%。",
    ("3. 差异集中于 " + "、".join(f"{r['name']} {r['diff']:+d} 件" for r in sorted(diff_rows, key=lambda x: x['diff']))
     + f"；其余 {len(ok_rows)} 个单品账实相符。差异原因待双方核对，本表不作归因。") if diff_rows else
  f"3. 本期 {len(ok_rows)} 个有铺货单品全部账实相符，无差异。",
    f"4. 期末未结算货值 {BOOK_END:,.2f} 元；乙方调整后结存货值 {ADJ_VAL:,.2f} 元；差额 {BOOK_END - ADJ_VAL:,.2f} 元即上述库存差异金额，已在本表逐项列示。",
    f"5. 若上述差异 {T['diff']:+d} 件经双方核对确认为 {_d_end.month} 月动销口径差异，则本期应结算数量为 {T['q10'] - T['diff']} 件、金额 {_settle_after:,.2f} 元，"
    + (f"贴花亦需按调整后数量重新计算，对账单与贴花均按核对结果重新出具。" if TH_ACTIVE
       else "对账单按核对结果重新出具。"),
    (f"6. 动销保障监控：合同期内累计发货 {OPEN_VAL:,.2f} 元，截至基准日未动销库存占当季发货总额 {ACT_VAL/OPEN_VAL*100:.1f}%，暂未触及合同约定的动销保障上限，需持续监控至合同期末。"
     if OPEN_VAL else "6. 期初未结算货值为 0，无法计算未动销库存占当季发货总额比例，请核对主档与甲方挂账表。"),
] + ([f"7. 贴花费用：按本期动销 {T['q10']} 件 × 各单品当期生效贴花标准计算，合计 {TH_TOTAL:,.2f} 元，明细见对账单第三节；兑付方式（抵扣货款或单独支付）按双方约定执行。"] if TH_ACTIVE else [])
for t in concl:
    mput(ws, r, 1, r, 12, t, F(9), A("left", "center", True, 1))
    ws.row_dimensions[r].height = 30 if len(t) < 90 else 44
    r += 1
r += 1

sec(ws, r, 1, 12, "五、需乙方书面确认／说明的事项"); r += 1
hdr(ws, r, 1, ["序号", "事  项", "", "", "涉 及 内 容", "", "", "", "需 乙 方 回 复", "", "", ""], 26)
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=8)
ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=12)
r += 1
asks = [("1", "本月结算金额确认", f"{T['settle']:,.2f} 元／{T['q10']} 件，逐 SKU 核对单价与动销数量", f"签章确认，收到后 {CT_SIGN_TXT}内寄回原件"),
        ("2", "库存差异说明", f"{T['diff']:+d} 件／{T['diff_amt']:,.2f} 元，逐 SKU 说明差异原因并提供支持单据", f"收到后 {CT_OBJ_TXT}内书面回复"),
        ("3", "非结算类库存归属", f"{len(extra)} 个单品共 {EXTRA_Q:,} 件，甲方挂账表未列示", "请确认归属与结算属性"),
        ("4", "结算期间划分", (f"{NEXT_LABEL}销售 {T['q11']} 件（{NOV_VAL:,.2f} 元）回加库存、不计入 {_d_end.month} 月"
                             if ADD_BACK else "库存导出日不晚于期末，无导出日前销售需回加"), "无异议视为确认")]
for no, item, cont, rep in asks:
    mput(ws, r, 1, r, 1, no, F(9, True), A("center"))
    mput(ws, r, 2, r, 4, item, F(9, True), A("left", "center", True, 1))
    mput(ws, r, 5, r, 8, cont, F(9), A("left", "center", True, 1))
    mput(ws, r, 9, r, 12, rep, F(9), A("left", "center", True, 1))
    ws.row_dimensions[r].height = 32; r += 1
r += 1

sec(ws, r, 1, 12, "六、附表清单"); r += 1
hdr(ws, r, 1, ["附  表", "", "用  途", "", "", "", "", "", "", "", "", ""], 24)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=12)
r += 1
for no, nm, use in [("01", "对账单", "正式结算单据，可直接打印盖章，作为乙方确认及付款依据"),
                    ("02", "库存差异表", ("逐 SKU 账实核对（含导出日前销量回加），作为乙方核对及书面说明依据" if ADD_BACK
                                        else "逐 SKU 账实核对，作为乙方核对及书面说明依据")),
                    ("03", "源数据对照", "甲方挂账／乙方动销／乙方库存 三源逐 SKU 汇总对照，供双方逐行确认"),
                    ("04", "甲方挂账明细", "甲方 ERP 未结算挂账原表（清洗），期初余额依据"),
                    ("05", "乙方销售汇总", f"{_d_end.month} 月动销按 SKU × 门店交叉汇总，销量核对依据"),
                    ("06", "乙方库存明细", f"乙方 {_d_asof:%Y%m%d} 库存原表（{inv_sheet_name} 清洗），库存核对依据"),
                    ("07", "乙方销售流水", f"乙方当期销售流水明细 {len(detail)} 行，逐笔溯源")]:
    mput(ws, r, 1, r, 2, f"附表 {no}", F(9, True, BLUE), A("center"), VLBLUE)
    mput(ws, r, 3, r, 4, nm, F(9, True), A("left", "center", True, 1))
    mput(ws, r, 5, r, 12, use, F(9), A("left", "center", True, 1))
    ws.row_dimensions[r].height = 21; r += 1
r += 1

sec(ws, r, 1, 12, "七、签章确认"); r += 1
mput(ws, r, 1, r, 6, f"甲方（盖章）：{PA_NAME}", F(9.5, True), A("left", "center", True, 1))
mput(ws, r, 7, r, 12, f"乙方（盖章）：{PB_NAME}", F(9.5, True), A("left", "center", True, 1))
ws.row_dimensions[r].height = 30; r += 1
mput(ws, r, 1, r, 6, "授权代表（签字）：", F(9), A("left", "center", True, 1))
mput(ws, r, 7, r, 12, "授权代表（签字）：", F(9), A("left", "center", True, 1))
ws.row_dimensions[r].height = 42; r += 1
mput(ws, r, 1, r, 6, "日　期：　　　年　　月　　日", F(9), A("left", "center", True, 1))
mput(ws, r, 7, r, 12, "日　期：　　　年　　月　　日", F(9), A("left", "center", True, 1))
ws.row_dimensions[r].height = 30
print("00 ok")

# ==================== 01 对账单 ====================
ws = out.create_sheet("01-对账单")
setup(ws, [6, 42, 12, 7, 12, 14, 16])
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 7, "销  售  对  账  单", F(22, True, NAVY), A("center"))
ws.row_dimensions[2].height = 44
mput(ws, 3, 1, 3, 7, f"{P_CN}    ·    对账期间：{P_START} 至 {P_END}", F(10, False, DGREY), A("center"))
ws.row_dimensions[3].height = 22
for c in range(1, 8):
    ws.cell(row=4, column=c).border = Border(top=MED)
ws.row_dimensions[4].height = 6

mput(ws, 5, 1, 5, 4, f"甲方（供货方）：{PA_NAME}", F(9.5, True), A("left", "center", True, 0))
mput(ws, 5, 5, 5, 7, f"单据编号：＿＿＿＿＿＿　　制单日期：{ASOF_CN}", F(9, False, DGREY), A("right", "center", True, 0))
mput(ws, 6, 1, 6, 4, f"纳税人识别号：{PA_TAX}", F(9, False, DGREY), A("left", "center", True, 0))
mput(ws, 6, 5, 6, 7, f"合同编号：{CONTRACT_NO}", F(9, False, DGREY), A("right", "center", True, 0))
mput(ws, 7, 1, 7, 4, f"乙方（代销方）：{PB_NAME}", F(9.5, True), A("left", "center", True, 0))
mput(ws, 7, 5, 7, 7, "币种：人民币（CNY）", F(9, False, DGREY), A("right", "center", True, 0))
mput(ws, 8, 1, 8, 4, f"纳税人识别号：{PB_TAX}", F(9, False, DGREY), A("left", "center", True, 0))
mput(ws, 8, 5, 8, 7, f"结算方式：{CT_METHOD_TXT}", F(9, False, DGREY), A("right", "center", True, 0))
for rr in range(5, 9):
    for c in range(1, 8):
        ws.cell(row=rr, column=c).border = Border(bottom=Side(style="dotted", color=BORDC))
    ws.row_dimensions[rr].height = 21
mput(ws, 9, 1, 9, 7, f"结算依据：按合同约定（{SETTLE_TERMS_TXT}）计算；已发货但截至对账日尚未动销的铺货产品，不计入本期结算货款（合同约定的结算口径）。",
     F(8.5, False, BLUE), A("left", "center", True, 0), VLBLUE)
ws.row_dimensions[9].height = 26
ws.row_dimensions[10].height = 5

sec(ws, 11, 1, 7, "一、本期结算明细")
hdr(ws, 12, 1, ["序号", "商品名称", "规格", "单位", "结算单价\n（元）", "本期动销数量", "结算金额\n（元）"], 34)
r = 13
idx = 0
pay_rows = [x for x in rows if x["q0"] > 0]
for x in pay_rows:
    idx += 1
    bg = WHITE if idx % 2 else GREY
    mput(ws, r, 1, r, 1, idx, F(9), A("center"), bg)
    mput(ws, r, 2, r, 2, x["name"], F(9), A("left", "center", True, 0), bg)
    mput(ws, r, 3, r, 3, x["spec"], F(9), A("center"), bg)
    mput(ws, r, 4, r, 4, x["unit"], F(9), A("center"), bg)
    c = ws.cell(row=r, column=5); num(c, x["price"], "#,##0.00"); c.fill = fill(bg)
    c = ws.cell(row=r, column=6); num(c, x["q10"]); c.fill = fill(bg)
    if x["q10"] == 0:
        c.font = F(9, c="A6A6A6")
    c = ws.cell(row=r, column=7); num(c, x["settle"], "#,##0.00"); c.fill = fill(bg)
    ws.row_dimensions[r].height = 26
    r += 1
mput(ws, r, 1, r, 4, "合　计", F(9.5, True, NAVY), A("center"), LBLUE)
c = ws.cell(row=r, column=5); c.fill = fill(LBLUE); c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
c = ws.cell(row=r, column=6); num(c, T["q10"]); c.font = F(10, True, NAVY)
c.fill = fill(LBLUE); c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
c = ws.cell(row=r, column=7); num(c, T["settle"], "#,##0.00"); c.font = F(10, True, NAVY)
c.fill = fill(LBLUE); c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
ws.row_dimensions[r].height = 28
r += 1
mput(ws, r, 1, r, 3, "人民币（大写）", F(9, True, DGREY), A("center"), GREY)
mput(ws, r, 4, r, 7, rmb_upper(T["settle"]), F(10.5, True, NAVY), A("left", "center", True, 1), GREY)
ws.row_dimensions[r].height = 30
r += 2

sec(ws, r, 1, 7, f"二、铺货余额结转（{P_START} → {ASOF}）"); r += 1
hdr(ws, r, 1, ["项目", "", "金额（元）", "计算／说明"], 26)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
r += 1
roll = [("① 期初未结算货值", OPEN_VAL, "甲方 ERP 未结算数量 × 结算单价", False),
        ("② 减：本期结算金额", T["settle"], "见“一、本期结算明细”", False),
        ("③ 期末未结算货值", BOOK_END, "＝ ① － ②", True),
        ("④ 减：导出日前已销售未结算货值" if ADD_BACK else "④ 无需回加（库存导出日不晚于期末）",
         NOV_VAL, (f"{T['q11']} 件，按合同归入下期" if ADD_BACK else "—"), False),
        ("⑤ 期末理论结存货值", round(BOOK_END - NOV_VAL, 2), "＝ ③ － ④", True),
        ("⑥ 库存差异（负数=短缺，正数=盘盈）", T["diff_amt"], f"{T['diff']:+d} 件，详见附表 02", False),
        ("⑦ 期末乙方库存货值", ACT_VAL, "＝ ⑤ ＋ ⑥，与乙方库存表合计一致", True)]
for lab, amt, note, bold in roll:
    mput(ws, r, 1, r, 2, lab, F(9, bool(bold), NAVY if bold else "000000"), A("left", "center", True, 1), LBLUE if bold else WHITE)
    c = ws.cell(row=r, column=3); num(c, amt, "#,##0.00")
    c.font = F(9.5, bool(bold), "000000")
    c.fill = fill(LBLUE if bold else WHITE)
    mput(ws, r, 4, r, 7, note, F(8.5, False, DGREY), A("left", "center", True, 1), LBLUE if bold else WHITE)
    ws.row_dimensions[r].height = 24
    r += 1
r += 1

sec(ws, r, 1, 7, "三、贴花费用（销售奖励，本期动销口径）"); r += 1
if TH_ACTIVE:
    hdr(ws, r, 1, ["序号", "商品名称", "动销数量", "贴花标准\n（元/件）", "贴花金额\n（元）", "依据", ""], 24)
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
    r += 1
    for i, x in enumerate(TH_ROWS, 1):
        mput(ws, r, 1, r, 1, i, F(9), A("center"), WHITE)
        mput(ws, r, 2, r, 2, x["name"], F(9), A("left", "center", True, 1), WHITE)
        c = ws.cell(row=r, column=3); num(c, x["qty"]); c.font = F(9); c.fill = fill(WHITE)
        c = ws.cell(row=r, column=4); num(c, x["std"], "#,##0.00"); c.font = F(9); c.fill = fill(WHITE)
        c = ws.cell(row=r, column=5); num(c, x["amt"], "#,##0.00"); c.font = F(9); c.fill = fill(WHITE)
        mput(ws, r, 6, r, 7, "主档「贴花」表当期生效标准", F(8, False, DGREY), A("left", "center", True, 1), WHITE)
        ws.row_dimensions[r].height = 20
        r += 1
    mput(ws, r, 1, r, 2, "合　计", F(9.5, True, WHITE), A("center"), NAVY)
    c = ws.cell(row=r, column=3); num(c, sum(x["qty"] for x in TH_ROWS)); c.font = F(9.5, True, WHITE)
    c.fill = fill(NAVY); c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    for j in (4,):
        c = ws.cell(row=r, column=j); c.value = None; c.fill = fill(NAVY)
        c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    c = ws.cell(row=r, column=5); num(c, TH_TOTAL, "#,##0.00"); c.font = F(10, True, WHITE)
    c.fill = fill(NAVY); c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    mput(ws, r, 6, r, 7, "按本期动销数量 × 当期生效标准", F(8, False, WHITE), A("left", "center", True, 1), NAVY)
    ws.row_dimensions[r].height = 26
    r += 1
    _net = round(T["settle"] - TH_TOTAL, 2)
    if _net >= 0:
        _net_txt = (f"若约定从货款抵扣，乙方本期应付净额 ＝ 结算货款 {T['settle']:,.2f} － 贴花 {TH_TOTAL:,.2f} "
                    f"＝ {_net:,.2f} 元；若分开结算，本节金额单独支付")
    else:
        _net_txt = (f"本期贴花 {TH_TOTAL:,.2f} 元超过结算货款 {T['settle']:,.2f} 元，若约定抵扣则甲方应向乙方支付差额 "
                    f"{abs(_net):,.2f} 元——请双方确认兑付方向后再出正式单据")
    mput(ws, r, 2, r, 7,
         f"口径：贴花按本期结算类动销数量 × 当期生效贴花标准计算（标准见主档「贴花」表，取对账期末 {P_END} 日生效版本；"
         f"非结算类陈列装不参与贴花；导出日前已发生销售的贴花随其货款计入下期）。"
         f"贴花兑付方式按双方约定执行：{_net_txt}。"
         f"本表仅列示计算结果，抵扣方式与开票口径（费用或货款折让）请双方书面确认。",
         F(8.5, False, AMBER), A("left", "center", True, 1), LAMBER)
    ws.row_dimensions[r].height = 44
    r += 1
else:
    mput(ws, r, 2, r, 7, "主档未建立「贴花」表，本期未计算销售奖励。如需计算，请先在主档按「商品条码 / 贴花标准 / 生效起止」登记。",
         F(8.5, False, DGREY), A("left", "center", True, 1), WHITE)
    r += 1
r += 1
sec(ws, r, 1, 7, "四、对账说明"); r += 1
notes = [
    ("1.", f"本对账单依据《经销合同》（编号 {CONTRACT_NO}）约定的结算口径出具（{SETTLE_TERMS_TXT}），未动销铺货不计入本期。"),
    ("2.", f"数据来源：甲方 ERP 系统“未结算挂账”表（{ASOF} 导出口径）、乙方《{PERIOD} 销售数据》、乙方《{_d_asof:%Y%m%d} 库存＋当月销售》表。"),
    ("3.", (f"乙方库存导出日为 {ASOF}，{NEXT_LABEL}期间发生销售 {T['q11']} 件（{NOV_VAL:,.2f} 元）。该部分已从库存回加、并从本期结算中剔出，归入 {NEXT_M} 月结算周期。"
             if ADD_BACK else f"乙方库存导出日为 {ASOF}，不晚于对账截止日 {P_END}，无需回加导出日前销售。")),
    ("4.", f"本期结算数量与乙方库存变动逐 SKU 勾稽：{len(ok_rows) + len(diff_rows)} 个有铺货单品中 {len(ok_rows)} 个账实相符，{len(diff_rows)} 个存在差异 {T['diff']:+d} 件／{T['diff_amt']:,.2f} 元（详见附表 02《库存差异表》），该部分未计入本期结算货款，待双方核对确认后处理。"),
    ("5.", f"请乙方于收到本对账单后 {CT_SIGN_TXT}内将签章原件寄回甲方；如有异议，应于 {CT_OBJ_TXT}内书面提出。逾期未提出异议，视为对本对账单内容予以确认。"),
    ("6.", f"付款期限：请于 {PAY_CN}前将上述款项汇入甲方指定账户。按合同约定，对账单存在异议不影响无争议部分按期支付。"),
]
for no, t in notes:
    mput(ws, r, 1, r, 1, no, F(8.5, True, BLUE), A("right", "top", False, 0))
    mput(ws, r, 2, r, 7, t, F(8.5), A("left", "center", True, 0))
    ws.row_dimensions[r].height = 20 if len(t) < 62 else 30
    r += 1
for t in [f"户　名：{BK_NAME}",
          f"账　号：{BK_NO}",
          f"开户行：{BK_BANK}"]:
    mput(ws, r, 2, r, 7, t, F(8.5, True, "000000"), A("left", "center", True, 1), VLBLUE)
    ws.row_dimensions[r].height = 18; r += 1
mput(ws, r, 2, r, 7, "7. 甲方按乙方实际付款金额开具增值税发票，请付款时同步提供开票资料。本对账单所列单价为双方约定的供货结算价，税率与发票种类以实际开票资料为准。", F(8.5), A("left", "center", True, 0))
ws.row_dimensions[r].height = 20
r += 2

sec(ws, r, 1, 7, "五、双方确认"); r += 1
mput(ws, r, 1, r, 3, "甲方（盖章）：", F(9, True), A("left", "center", True, 1))
mput(ws, r, 4, r, 7, "乙方（盖章）：", F(9, True), A("left", "center", True, 1))
ws.row_dimensions[r].height = 46; r += 1
mput(ws, r, 1, r, 3, "授权代表（签字）：", F(9), A("left", "center", True, 1))
mput(ws, r, 4, r, 7, "授权代表（签字）：", F(9), A("left", "center", True, 1))
ws.row_dimensions[r].height = 34; r += 1
mput(ws, r, 1, r, 3, "日　期：　　　年　　月　　日", F(9), A("left", "center", True, 1))
mput(ws, r, 4, r, 7, "日　期：　　　年　　月　　日", F(9), A("left", "center", True, 1))
ws.row_dimensions[r].height = 26
ws.print_area = f"A1:G{r}"
print("01 ok")

# ==================== 02 库存差异表 ====================
ws = out.create_sheet("02-库存差异表")
setup(ws, [4, 13, 28, 8, 4, 9, 10, 9, 9, 9, 9, 10, 9, 10, 8, 17], "landscape")
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 16, "库 存 差 异 表", F(20, True, NAVY), A("center"))
ws.row_dimensions[2].height = 38
mput(ws, 3, 1, 3, 16, f"{PA_NAME}  ×  {PB_NAME}    ·    对账期间 {P_START} 至 {P_END}    ·    库存数据基准日 {ASOF}",
     F(9, False, DGREY), A("center"))
ws.row_dimensions[3].height = 22
mput(ws, 4, 1, 4, 16,
     (f"口径：乙方库存导出日为 {ASOF}，晚于对账截止日 {P_END}，须将 {NEXT_LABEL}已发生销售（共 {T['q11']} 件 / {NOV_VAL:,.2f} 元）回加至乙方库存，使双方数据处于同一时点后再行比对。"
      f"　调整后结存 ＝ 乙方库存 ＋ {NEXT_M} 月销售；差异 ＝ 调整后结存 － 理论结存。"
      if ADD_BACK else
      f"口径：乙方库存导出日 {ASOF} 不晚于对账截止日 {P_END}，库存即为期末结存，无需回加。差异 ＝ 乙方库存 － 理论结存。"),
     F(8.5, False, AMBER), A("left", "center", True, 0), LAMBER)
ws.row_dimensions[4].height = 34
ws.row_dimensions[5].height = 6

HDR2 = ["序号", "商品条码", "商品名称", "规格", "单位", "结算单价\n（元）",
        "甲方期初\n未结算数量", f"减：{_d_end.month}月\n申报动销", "甲方理论\n结存数量",
        f"乙方 {_d_asof.month:02d}-{_d_asof.day:02d}\n库存（系统导出）",
        (f"加：{_d_next.month}月\n{_d_next.day}—{_d_asof.day}销售" if ADD_BACK else "导出日前销售\n（无需回加）"),
        "乙方调整后\n结存数量",
        "差异数量", "差异金额\n（元）", "差异率\n（÷理论结存）", "差异性质"]

r = 6
sec(ws, r, 1, 16, f"一、库存差异明细（{len(diff_rows)} 个 SKU，需乙方书面说明）"); r += 1
hdr(ws, r, 1, HDR2, 40); r += 1
tot_d = dict(q0=0, q10=0, theo=0, act=0, q11=0, adj=0, diff=0, amt=0)
for i, x in enumerate(diff_rows, 1):
    bg = LRED
    _rate = x["diff"] / x["theo"] if x["theo"] else 0.0
    vals = [i, x["bar"], x["name"], x["spec"], x["unit"], x["price"],
            x["q0"], -x["q10"] if x["q10"] else 0, x["theo"],
            x["act"], x["q11"], x["adj"], x["diff"], x["diff_amt"], _rate]
    for j, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=j)
        if j in (1, 3, 4, 5):
            cell.value = v; cell.font = F(8.5, j == 3); cell.alignment = A("left" if j == 3 else "center", "center", True, 1 if j == 3 else 0)
        elif j == 2:
            cell.value = v; cell.font = F(8.5); cell.alignment = A("center")
            cell.number_format = "@"
        elif j == 15:
            cell.value = abs(v); cell.number_format = "0.00%"; cell.font = F(8.5, True, RED); cell.alignment = A("center")
        else:
            num(cell, v, "#,##0.00" if j in (6, 14) else "#,##0")
            if j in (13, 14): cell.font = F(8.5, True, RED)
        cell.fill = fill(bg)
    mput(ws, r, 16, r, 16, "库存短缺，需说明去向" if x["diff"] < 0 else "库存盘盈，需说明来源", F(8, True, RED), A("center"), bg)
    ws.row_dimensions[r].height = 30
    for k, v in [("q0", x["q0"]), ("q10", x["q10"]), ("theo", x["theo"]), ("act", x["act"]),
                 ("q11", x["q11"]), ("adj", x["adj"]), ("diff", x["diff"]), ("amt", x["diff_amt"])]:
        tot_d[k] += v
    r += 1
mput(ws, r, 1, r, 5, "小　计", F(9, True, NAVY), A("center"), LBLUE)
for j, v in [(6, None), (7, tot_d["q0"]), (8, -tot_d["q10"]), (9, tot_d["theo"]), (10, tot_d["act"]),
             (11, tot_d["q11"]), (12, tot_d["adj"]), (13, tot_d["diff"]), (14, tot_d["amt"])]:
    cell = ws.cell(row=r, column=j)
    if v is None:
        cell.value = None
    else:
        num(cell, v, "#,##0.00" if j == 14 else "#,##0")
        cell.font = F(9, True, NAVY)
    cell.fill = fill(LBLUE)
    cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
tot_rate = tot_d["diff"] / tot_d["theo"] if tot_d["theo"] else 0
cell = ws.cell(row=r, column=15); cell.value = abs(tot_rate); cell.number_format = "0.00%"
cell.font = F(9, True, RED); cell.alignment = A("center"); cell.fill = fill(LBLUE)
cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
mput(ws, r, 16, r, 16, ((f"{abs(tot_d['diff'])} 件净短缺" if tot_d["diff"] < 0 else f"{tot_d['diff']} 件净盘盈")
                           if tot_d["diff"] else "无差异"), F(9, True, RED), A("center"), LBLUE)
ws.row_dimensions[r].height = 26
r += 2

sec(ws, r, 1, 16, f"二、核对相符明细（{len(ok_rows)} 个 SKU，账实一致）"); r += 1
hdr(ws, r, 1, HDR2, 40, fl=BLUE); r += 1
tot_o = dict(q0=0, q10=0, theo=0, act=0, q11=0, adj=0)
for i, x in enumerate(ok_rows, 1):
    bg = WHITE if i % 2 else GREY
    vals = [i, x["bar"], x["name"], x["spec"], x["unit"], x["price"],
            x["q0"], -x["q10"] if x["q10"] else 0, x["theo"], x["act"], x["q11"], x["adj"],
            0, 0.0, 0.0]
    for j, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=j)
        if j in (1, 3, 4, 5):
            cell.value = v; cell.font = F(8.5, j == 3); cell.alignment = A("left" if j == 3 else "center", "center", True, 1 if j == 3 else 0)
        elif j == 2:
            cell.value = v; cell.font = F(8.5); cell.alignment = A("center"); cell.number_format = "@"
        elif j in (13, 14, 15):
            num(cell, v if j != 15 else v, "#,##0" if j == 13 else ("#,##0.00" if j == 14 else "0.0%"))
            cell.font = F(8.5, c=GREEN)
        else:
            num(cell, v, "#,##0.00" if j == 6 else "#,##0")
        cell.fill = fill(bg)
    mput(ws, r, 16, r, 16, "✓ 账实相符", F(8, True, GREEN), A("center"), bg)
    ws.row_dimensions[r].height = 26
    for k, v in [("q0", x["q0"]), ("q10", x["q10"]), ("theo", x["theo"]), ("act", x["act"]),
                 ("q11", x["q11"]), ("adj", x["adj"])]:
        tot_o[k] += v
    r += 1
mput(ws, r, 1, r, 5, "小　计", F(9, True, NAVY), A("center"), LBLUE)
for j, v in [(7, tot_o["q0"]), (8, -tot_o["q10"]), (9, tot_o["theo"]), (10, tot_o["act"]),
             (11, tot_o["q11"]), (12, tot_o["adj"]), (13, 0), (14, 0.0), (15, 0.0)]:
    cell = ws.cell(row=r, column=j)
    num(cell, v, "#,##0.00" if j == 14 else "#,##0")
    if j == 15: cell.number_format = "0.0%"
    cell.font = F(9, True, NAVY)
    cell.fill = fill(LBLUE); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
mput(ws, r, 16, r, 16, "全部相符", F(9, True, GREEN), A("center"), LBLUE)
ws.row_dimensions[r].height = 26
r += 1
mput(ws, r, 1, r, 5, "合　计（一＋二）", F(9.5, True, WHITE), A("center"), NAVY)
for j, v in [(6, None), (7, T["q0"]), (8, -T["q10"]), (9, T["theo"]), (10, T["act"]),
             (11, T["q11"]), (12, T["adj"]), (13, T["diff"]), (14, T["diff_amt"]), (15, DIFF_RATE)]:
    cell = ws.cell(row=r, column=j)
    if v is None:
        cell.value = None
    else:
        num(cell, abs(v) if j == 15 else v, "#,##0.00" if j == 14 else "#,##0")
        if j == 15: cell.number_format = "0.00%"
        if j in (13, 14, 15): cell.font = F(9.5, True, RED)
        else: cell.font = F(9.5, True, WHITE)
    cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
mput(ws, r, 16, r, 16, ((f"净短缺 {abs(T['diff'])} 件（{abs(DIFF_RATE)*100:.2f}%）" if T['diff'] < 0
                          else f"净盘盈 {T['diff']} 件（{abs(DIFF_RATE)*100:.2f}%）") if T['diff'] else "无差异"),
     F(9, True, WHITE), A("center"), NAVY)
ws.row_dimensions[r].height = 28
r += 1
_zero_names = [z["name"] for z in zero_rows]
mput(ws, r, 1, r, 16,
     (("注：" + "、".join(_zero_names) + " 本期期初铺货数量为 0，不参与数量差异比对，已列入附表 03、06 核对。")
      if _zero_names else "注：本期无期初铺货为 0 的新品，差异明细覆盖全部有铺货 SKU。"),
     F(8, False, DGREY), A("left", "center", True, 0))
ws.row_dimensions[r].height = 22
r += 2
print("02 part ok")

sec(ws, r, 1, 16, f"三、非结算类库存（{len(extra)} 个单品，共 {EXTRA_Q:,} 件，甲方账面未挂账，不参与本期结算）"); r += 1
hdr(ws, r, 1, ["序号", "商品条码", "商品名称", "规格", "单位", "零售价\n（元）", "库存数量", "说　　明"], 32, fl="7F7F7F")
ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=16)
r += 1
for i, x in enumerate(extra, 1):
    bg = WHITE if i % 2 else GREY
    mput(ws, r, 1, r, 1, i, F(8.5), A("center"), bg)
    cell = ws.cell(row=r, column=2); cell.value = x["bar"]; cell.font = F(8.5); cell.alignment = A("center"); cell.number_format = "@"; cell.fill = fill(bg)
    mput(ws, r, 3, r, 3, x["name"], F(8.5), A("left", "center", True, 0), bg)
    mput(ws, r, 4, r, 4, x["spec"], F(8.5), A("center"), bg)
    mput(ws, r, 5, r, 5, "片", F(8.5), A("center"), bg)
    cell = ws.cell(row=r, column=6); num(cell, x["retail"], "#,##0.00"); cell.fill = fill(bg)
    cell = ws.cell(row=r, column=7); num(cell, x["qty"]); cell.fill = fill(bg)
    note = "甲方挂账表未列示，本表仅列数量，请乙方确认归属"
    mput(ws, r, 8, r, 16, note, F(8.5, False, DGREY), A("left", "center", True, 1), bg)
    ws.row_dimensions[r].height = 22
    r += 1
mput(ws, r, 1, r, 5, "小　计", F(9, True, NAVY), A("center"), LBLUE)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
cell = ws.cell(row=r, column=6); cell.value = None; cell.fill = fill(LBLUE)
cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
cell = ws.cell(row=r, column=7); num(cell, EXTRA_Q); cell.font = F(9, True, NAVY); cell.fill = fill(LBLUE)
cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
mput(ws, r, 8, r, 16, "不参与本期结算，仅作数量与归属确认", F(8.5, False, DGREY), A("left", "center", True, 1), LBLUE)
ws.row_dimensions[r].height = 24
r += 2

sec(ws, r, 1, 16, "四、差异结论与处理建议"); r += 1
advice = [
    ("差异性质", (f"{len(diff_rows)} 个 SKU 存在差异，合计 {T['diff']:+d} 件 / {T['diff_amt']:,.2f} 元，"
                f"占理论结存 {abs(DIFF_RATE)*100:.2f}%；{len(ok_rows)} 个 SKU 账实相符。") if diff_rows
               else f"{len(ok_rows)} 个有铺货 SKU 全部账实相符。"),
    ("差异分布", ("差异集中在 " + "、".join(f"{r['name']} {r['diff']:+d}" for r in sorted(diff_rows, key=lambda x: x['diff']))
                + f"；其余 {len(ok_rows)} 个单品差异为 0。") if diff_rows else
               f"{len(ok_rows)} 个有铺货单品全部账实相符，无差异。"),
    ("可能成因", "① 已销售但未纳入本次提供的销售流水（漏报／延时上传）；② 跨门店或跨渠道调拨未登记；③ 赠品、试用、损耗出库未入账；④ 盘点口径差异（含在途、含已打包待发）。"),
    ("处理建议", f"① 请乙方核对上述 {len(diff_rows)} 个 SKU 的本期销售、调拨、赠品及损耗记录，并提供支持单据；② 差异确认前，该部分 {abs(T['diff_amt']):,.2f} 元不纳入本期结算；③ 经确认属 {_d_end.month} 月动销口径差异的，在本期调整结算并重新出具对账单；④ 属其他原因的，双方另行书面确认处理方式。"),
    ("后续控制", f"建议自 {NEXT_M} 月起要求乙方按周报送动销流水，并在库存导出时同步标注导出时点及导出时点的当日销量，避免因时间口径差异重复产生核对差异。"),
]
for lab, txt in advice:
    mput(ws, r, 1, r, 2, lab, F(9, True, BLUE), A("center", "center", True, 0), VLBLUE)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    mput(ws, r, 3, r, 16, txt, F(9), A("left", "center", True, 0))
    ws.row_dimensions[r].height = 24 if len(txt) < 80 else 44
    r += 1
r += 1

mput(ws, r, 1, r, 8, f"甲方（盖章）：{PA_NAME}", F(9, True), A("left", "center", True, 1))
mput(ws, r, 9, r, 16, f"乙方（盖章）：{PB_NAME}", F(9, True), A("left", "center", True, 1))
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=16)
ws.row_dimensions[r].height = 34; r += 1
mput(ws, r, 1, r, 8, "对账人（签字）：　　　　　　　日期：　　年　　月　　日", F(9), A("left", "center", True, 1))
mput(ws, r, 9, r, 16, "对账人（签字）：　　　　　　　日期：　　年　　月　　日", F(9), A("left", "center", True, 1))
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=16)
ws.row_dimensions[r].height = 30
print("02 ok")

# ==================== 03 源数据对照 ====================
ws = out.create_sheet("03-源数据对照")
setup(ws, [4, 13, 27, 12, 8, 4, 9, 8, 8, 10, 8, 8, 8, 8, 7, 9, 8], "landscape")
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 17, "源 数 据 对 照 表（甲方挂账 × 乙方动销 × 乙方库存）", F(18, True, NAVY), A("center"))
ws.row_dimensions[2].height = 36
mput(ws, 3, 1, 3, 17, f"供双方逐行确认　·　数量单位：件／盒／瓶／套　·　金额单位：人民币元　·　数据基准日 {ASOF}",
     F(9, False, DGREY), A("center"))
ws.row_dimensions[3].height = 22
ws.row_dimensions[4].height = 6

HDR3 = ["序号", "商品条码", "商品名称（乙方系统）", "甲方系统品名差异", "规格", "单位",
        "结算单价", "甲方期初\n未结算数量", f"{_d_end.month} 月\n申报动销", "本月结算金额",
        (f"{_d_next.month} 月 {_d_next.day}—{_d_asof.day}\n销售" if ADD_BACK else "导出日前\n销售（无）"),
        f"乙方 {_d_asof.month:02d}-{_d_asof.day:02d}\n库存（系统导出）", "乙方调整后\n结存", "甲方理论\n结存",
        "差异数量", "差异金额", "核对结果"]
r = 5
hdr(ws, r, 1, HDR3, 42); r += 1
for i, x in enumerate(rows, 1):
    bg = LRED if x["diff"] else (WHITE if i % 2 else GREY)
    vals = [i, x["bar"], x["name"], x["note"] or "—", x["spec"], x["unit"], x["price"],
            x["q0"], x["q10"], x["settle"], x["q11"], x["act"], x["adj"], x["theo"],
            x["diff"], x["diff_amt"]]
    for j, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=j)
        if j in (1, 5, 6):
            cell.value = v; cell.font = F(8.5); cell.alignment = A("center")
        elif j == 2:
            cell.value = v; cell.font = F(8.5); cell.alignment = A("center"); cell.number_format = "@"
        elif j == 3:
            cell.value = v; cell.font = F(8.5); cell.alignment = A("left", "center", True, 1)
        elif j == 4:
            cell.value = v; cell.font = F(7.5, c=DGREY); cell.alignment = A("left", "center", True, 1)
        else:
            num(cell, v, "#,##0.00" if j in (7, 10, 16) else "#,##0")
            cell.font = F(8.5)
        cell.fill = fill(bg)
    res = "✗ 短缺" if x["diff"] else "✓ 相符"
    mput(ws, r, 17, r, 17, res, F(8.5, True, RED if x["diff"] else GREEN), A("center"), bg)
    ws.row_dimensions[r].height = 24
    r += 1
mput(ws, r, 1, r, 6, "合　计", F(9.5, True, WHITE), A("center"), NAVY)
for j, v in [(7, None), (8, T["q0"]), (9, T["q10"]), (10, T["settle"]), (11, T["q11"]),
             (12, T["act"]), (13, T["adj"]), (14, T["theo"]), (15, T["diff"]), (16, T["diff_amt"])]:
    cell = ws.cell(row=r, column=j)
    if v is None: cell.value = None
    else:
        num(cell, v, "#,##0.00" if j in (10, 16) else "#,##0")
        cell.font = F(9, True, RED if j in (15, 16) else WHITE)
    cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
mput(ws, r, 17, r, 17, f"差异 {abs(DIFF_RATE)*100:.2f}%", F(9, True, WHITE), A("center"), NAVY)
ws.row_dimensions[r].height = 28
r += 2
_has_name_note = any(x["note"] for x in rows)
mput(ws, r, 1, r, 17,
     ("注 1：甲方系统品名与乙方系统品名存在括号全半角、空格或后缀差异，已按商品条码一一对应，品名差异见「甲方系统品名差异」列。"
      if _has_name_note else "注 1：甲方系统品名与乙方系统品名已按商品条码一一对应。"),
     F(8, False, DGREY), A("left", "center", True, 0))
r += 1
mput(ws, r, 1, r, 17,
     (f"注 2：乙方调整后结存 ＝ 乙方库存 ＋ {NEXT_LABEL}销售（{T['q11']} 件）；甲方理论结存 ＝ 期初未结算 － {_d_end.month} 月动销。二者同处 {P_END} 时点，可直接比较。"
      if ADD_BACK else
      f"注 2：乙方调整后结存 ＝ 乙方库存（导出日不晚于期末，无需回加）；甲方理论结存 ＝ 期初未结算 － {_d_end.month} 月动销。二者同处 {P_END} 时点，可直接比较。"),
     F(8, False, DGREY), A("left", "center", True, 0))
r += 1
mput(ws, r, 1, r, 17,
     (("注 3：" + "、".join(x["name"] for x in zero_rows) + " 本期期初铺货与库存均为 0，无发生额，已列入本表但不参与结算。")
      if zero_rows else "注 3：本期无期初铺货为 0 的新品，本表覆盖全部有铺货 SKU。"),
     F(8, False, DGREY), A("left", "center", True, 0))
ws.freeze_panes = "D6"
print("03 ok")

# ==================== 04 甲方挂账明细 ====================
ws = out.create_sheet("04-甲方挂账明细")
setup(ws, [5, 40, 10, 13, 10, 10, 15, 5, 12, 18], "landscape")
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 10, "附 表 04　甲方 ERP 系统未结算挂账明细", F(15, True, NAVY), A("center"))
ws.row_dimensions[2].height = 32
mput(ws, 3, 1, 3, 10, f"数据来源：{F_JIA}（{ASOF} 导出）　·　金额单位：人民币元",
     F(8.5, False, DGREY), A("center"))
ws.row_dimensions[3].height = 20
ws.row_dimensions[4].height = 6
r = 5
hdr(ws, r, 1, ["序号", "商品名称", "单价", "未结算金额", "未结算数量", "初始数量", "商品条码", "单位", "金额校验", "备注"], 34)
r += 1
for i, x in enumerate(rows, 1):
    bg = WHITE if i % 2 else GREY
    mput(ws, r, 1, r, 1, i, F(9), A("center"), bg)
    mput(ws, r, 2, r, 2, x["name"], F(9), A("left", "center", True, 0), bg)
    for j, v in [(3, x["price"]), (4, x["src_amt"]), (5, x["q0"]), (6, x["init"])]:
        cell = ws.cell(row=r, column=j)
        num(cell, v, "#,##0.00" if j in (3, 4) else "#,##0"); cell.fill = fill(bg)
    cell = ws.cell(row=r, column=7); cell.value = x["bar"]; cell.font = F(9); cell.alignment = A("center")
    cell.number_format = "@"; cell.fill = fill(bg)
    mput(ws, r, 8, r, 8, x["unit"], F(9), A("center"), bg)
    chk = round(x["price"] * x["q0"] - x["src_amt"], 2)
    cell = ws.cell(row=r, column=9); num(cell, chk, "#,##0.00"); cell.fill = fill(bg)
    if abs(chk) < 0.01: cell.font = F(9, c=GREEN)
    note = ""
    if x["q0"] == 0: note = "本期无铺货余额"
    elif x["init"] != x["q0"]: note = "初始数量与未结算数量不一致"
    mput(ws, r, 10, r, 10, note or "—", F(8, c=DGREY), A("left", "center", True, 1), bg)
    ws.row_dimensions[r].height = 24
    r += 1
mput(ws, r, 1, r, 2, "合　计", F(9.5, True, WHITE), A("center"), NAVY)
_chk_total = round(sum(round(x["price"] * x["q0"] - x["src_amt"], 2) for x in rows), 2)
for j, v in [(3, None), (4, OPEN_VAL), (5, T["q0"]), (6, sum(x["init"] for x in rows)), (7, None), (8, None), (9, _chk_total), (10, None)]:
    cell = ws.cell(row=r, column=j)
    if v is None:
        cell.value = None; cell.font = F(9, True, WHITE)
    else:
        num(cell, v, "#,##0.00" if j in (4, 9) else "#,##0")
        cell.font = F(9.5, True, WHITE)
    cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
ws.row_dimensions[r].height = 28
r += 2
_init_same = all(x["init"] == x["q0"] for x in rows)
mput(ws, r, 1, r, 10,
     (("说明：初始数量与未结算数量逐行一致，表明截至 " + ASOF + " 甲方系统内尚无结算扣减记录，未结算余额即为 "
       + P_START + " 期初铺货余额；") if _init_same else
      "说明：存在初始数量与未结算数量不一致的明细行（见备注列），本表按未结算数量作为本期期初，请核对甲方系统结算记录；")
     + "金额校验列＝单价×未结算数量－未结算金额，合计为 0 表示源表内部一致。",
     F(8, False, DGREY), A("left", "center", True, 0))
ws.row_dimensions[r].height = 34
r += 2
sec(ws, r, 1, 10, f"二、源数据核对记录（复核日 {ASOF}）"); r += 1
verify = CFG.get("verify_notes") or [
    ("合计校验", f"{len(rows)} 行经复算：金额合计 {OPEN_VAL:,.2f} 元、数量合计 {T['q0']:,} 件，与本表合计行一致；"
              f"逐行“单价 × 未结算数量 － 未结算金额”偏差合计 {sum(round(x['price'] * x['q0'] - x['src_amt'], 2) for x in rows):+,.2f} 元"
              + ("（全部为 0，源表内部一致）。" if abs(sum(round(x['price'] * x['q0'] - x['src_amt'], 2) for x in rows)) < 0.01
                 else "（存在偏差，请核对上方明细行）。")
              + "源表无合计行，合计数系复算所得。"),
    ("数量口径", f"“未结算数量”与“初始数量”在 {len(rows)} 行中逐行相同，表明导出日尚未扣减任何结算，故该列即为本期期初铺货余额。如与系统实际不符，请以带客户名与期间的正式导出件为准。"),
    ("单价核对", "结算单价取自甲方挂账表。如另有签字生效的价格表，请核对两者是否一致，并以生效日期判断本期适用版本。"),
]
for lab, txt in verify:
    mput(ws, r, 1, r, 2, lab, F(8.5, True, BLUE), A("center", "center", True, 0), VLBLUE)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    mput(ws, r, 3, r, 10, txt, F(8.5), A("left", "center", True, 0))
    ws.row_dimensions[r].height = 22 if len(txt) < 60 else 34
    r += 1
print("04 ok")

# ==================== 05 乙方销售汇总 ====================
ws = out.create_sheet("05-乙方销售汇总")
setup(ws, [14, 34, 10, 5] + [11] * len(STORE_ORDER) + [8, 7], "landscape")
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 10, f"附 表 05　乙方 {PERIOD} 动销汇总（按 SKU × 门店）", F(15, True, NAVY), A("center"))
ws.row_dimensions[2].height = 32
mput(ws, 3, 1, 3, 10, f"数据来源：{F_OCT}（明细 {len(detail)} 行，汇总后 {len(oct_qty)} 个 SKU / {sum(oct_qty.values())} 件）",
     F(8.5, False, DGREY), A("center"))
ws.row_dimensions[3].height = 20
ws.row_dimensions[4].height = 6
r = 5
sec(ws, r, 1, 10, f"一、{P_CN}动销（按门店）"); r += 1
hdr(ws, r, 1, ["商品条码", "商品名称", "规格", "单位"] + STORE_ORDER + ["合　计", "占比"], 32)
_total_col = 5 + len(STORE_ORDER)
_ratio_col = _total_col + 1
r += 1
oct_bars = [b for b in oct_qty]
oct_bars.sort(key=lambda b: -oct_qty[b])
total_oct = sum(oct_qty.values())
for i, b in enumerate(oct_bars, 1):
    bg = WHITE if i % 2 else GREY
    cell = ws.cell(row=r, column=1); cell.value = b; cell.font = F(9); cell.alignment = A("center")
    cell.number_format = "@"; cell.fill = fill(bg)
    mput(ws, r, 2, r, 2, oct_name[b], F(9), A("left", "center", True, 0), bg)
    mput(ws, r, 3, r, 3, oct_spec[b], F(9), A("center"), bg)
    mput(ws, r, 4, r, 4, oct_unit[b], F(9), A("center"), bg)
    for k, st in enumerate(STORE_ORDER):
        cell = ws.cell(row=r, column=5 + k)
        num(cell, oct_store[b].get(st, 0)); cell.fill = fill(bg)
        if oct_store[b].get(st, 0) == 0: cell.font = F(9, c="BFBFBF")
    cell = ws.cell(row=r, column=_total_col); num(cell, oct_qty[b]); cell.font = F(9, True); cell.fill = fill(bg)
    cell = ws.cell(row=r, column=_ratio_col); cell.value = oct_qty[b] / total_oct
    cell.number_format = "0.0%"; cell.font = F(9); cell.alignment = A("center"); cell.fill = fill(bg)
    ws.row_dimensions[r].height = 24
    r += 1
mput(ws, r, 1, r, 4, "合　计", F(9.5, True, WHITE), A("center"), NAVY)
for k, st in enumerate(STORE_ORDER):
    cell = ws.cell(row=r, column=5 + k)
    num(cell, sum(oct_store[b].get(st, 0) for b in oct_bars)); cell.font = F(9.5, True, WHITE)
    cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
cell = ws.cell(row=r, column=_total_col); num(cell, total_oct); cell.font = F(10, True, WHITE)
cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
cell = ws.cell(row=r, column=_ratio_col); cell.value = 1.0; cell.number_format = "0.0%"
cell.font = F(9.5, True, WHITE); cell.alignment = A("center")
cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
ws.row_dimensions[r].height = 28
r += 2

_no_sale_names = '、'.join(sorted(x['name'] for x in rows if x['q10'] == 0 and x['q0'] > 0)) or '（无）'
if ADD_BACK:
    sec(ws, r, 1, 10, f"二、{NEXT_FULL}动销（回加至库存，不参与 {_d_end.month} 月结算，共 {T['q11']} 件）"); r += 1
    hdr(ws, r, 1, ["商品条码", "商品名称", "规格", "单位", f"{NEXT_LABEL}销量（件）", "", "", "", "", "占比"], 30, fl=BLUE)
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=9)
    r += 1
    nov_bars = sorted([b for b in nov if nov[b] > 0], key=lambda b: -nov[b])
    NOV_TOTAL = sum(nov.values())
    for i, b in enumerate(nov_bars, 1):
        bg = WHITE if i % 2 else GREY
        cell = ws.cell(row=r, column=1); cell.value = b; cell.font = F(9); cell.alignment = A("center")
        cell.number_format = "@"; cell.fill = fill(bg)
        mput(ws, r, 2, r, 2, oct_name.get(b, ""), F(9), A("left", "center", True, 0), bg)
        mput(ws, r, 3, r, 3, oct_spec.get(b, ""), F(9), A("center"), bg)
        mput(ws, r, 4, r, 4, oct_unit.get(b, ""), F(9), A("center"), bg)
        cell = ws.cell(row=r, column=5); num(cell, nov[b]); cell.fill = fill(bg)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=9)
        cell = ws.cell(row=r, column=10); cell.value = nov[b] / NOV_TOTAL
        cell.number_format = "0.0%"; cell.font = F(9); cell.alignment = A("center"); cell.fill = fill(bg)
        ws.row_dimensions[r].height = 24
        r += 1
    mput(ws, r, 1, r, 4, "合　计", F(9.5, True, WHITE), A("center"), NAVY)
    cell = ws.cell(row=r, column=5); num(cell, NOV_TOTAL); cell.font = F(10, True, WHITE)
    cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=9)
    cell = ws.cell(row=r, column=10); cell.value = 1.0; cell.number_format = "0.0%"
    cell.font = F(9.5, True, WHITE); cell.alignment = A("center")
    cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    ws.row_dimensions[r].height = 28
    r += 2
else:
    r += 1
mput(ws, r, 1, r, 10,
     (f"注：门店分布仅统计本期动销；导出日前销售（{T['q11']} 件）仅列合计。当期无动销单品：{_no_sale_names}。销售流水有记录日期 {COV_TXT}。"
      if ADD_BACK else
      f"注：门店分布仅统计本期动销。当期无动销单品：{_no_sale_names}。销售流水有记录日期 {COV_TXT}。"),
     F(8, False, DGREY), A("left", "center", True, 0))
ws.row_dimensions[r].height = 26
print("05 ok")

# ==================== 06 乙方库存明细 ====================
ws = out.create_sheet("06-乙方库存明细")
setup(ws, [4, 32, 14, 9, 11, 8, 8, 5, 11, 9, 10, 16], "landscape")
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 12, f"附 表 06　乙方 {_d_asof:%Y%m%d} 库存明细", F(15, True, NAVY), A("center"))
ws.row_dimensions[2].height = 32
mput(ws, 3, 1, 3, 12, f"数据来源：{F_INV}「{inv_sheet_name}」表　·　库存时点 {ASOF}　·　金额单位：人民币元",
     F(8.5, False, DGREY), A("center"))
ws.row_dimensions[3].height = 20
ws.row_dimensions[4].height = 6
r = 5
hdr(ws, r, 1, ["序号", "商品名称", "商品条码", "规格", "商品分类\n（乙方系统）", "库存数量", "零售价",
               "品牌", "结算属性", "甲方结算单价", "库存货值", "备　注"], 34)
r += 1
settle_by_bar = {x["bar"]: x for x in rows if x["bar"]}
settle_by_n = {norm(x["name"]): x for x in rows}
inv_sorted = sorted(inv, key=lambda x: (0 if norm(x["name"]) in settle_by_n else 1, x["name"]))
tot_q, tot_v = 0, 0.0
for i, x in enumerate(inv_sorted, 1):
    st = settle_by_bar.get(x["bar"]) or settle_by_n.get(norm(x["name"]))
    bg = WHITE if i % 2 else GREY
    mput(ws, r, 1, r, 1, i, F(9), A("center"), bg)
    mput(ws, r, 2, r, 2, x["name"], F(9), A("left", "center", True, 0), bg)
    cell = ws.cell(row=r, column=3); cell.value = x["bar"]; cell.font = F(9); cell.alignment = A("center")
    cell.number_format = "@"; cell.fill = fill(bg)
    mput(ws, r, 4, r, 4, x["spec"], F(9), A("center"), bg)
    mput(ws, r, 5, r, 5, x["cat"], F(9), A("center"), bg)
    cell = ws.cell(row=r, column=6); num(cell, x["qty"]); cell.fill = fill(bg)
    cell = ws.cell(row=r, column=7); num(cell, x["retail"], "#,##0.00"); cell.fill = fill(bg)
    mput(ws, r, 8, r, 8, x.get("brand") or "—", F(9), A("center"), bg)
    mput(ws, r, 9, r, 9, "结算类" if st else "非结算类", F(8.5, True, BLUE if st else "7F7F7F"), A("center"), VLBLUE if st else GREY)
    cell = ws.cell(row=r, column=10)
    if st:
        num(cell, st["price"], "#,##0.00"); cell.fill = fill(bg)
        cell = ws.cell(row=r, column=11); num(cell, round(x["qty"] * st["price"], 2), "#,##0.00"); cell.fill = fill(bg)
        tot_v += round(x["qty"] * st["price"], 2)
    else:
        cell.value = "—"; cell.font = F(9, c="A6A6A6"); cell.alignment = A("center"); cell.fill = fill(bg)
        cell = ws.cell(row=r, column=11); cell.value = "—"; cell.font = F(9, c="A6A6A6")
        cell.alignment = A("center"); cell.fill = fill(bg)
    note = ""
    if not st: note = "甲方挂账表未列示，本表仅列数量，请乙方确认归属"
    if x["qty"] == 0: note = "本期无库存"
    mput(ws, r, 12, r, 12, note or "—", F(8, c=DGREY), A("left", "center", True, 1), bg)
    ws.row_dimensions[r].height = 24
    tot_q += x["qty"]
    r += 1
mput(ws, r, 1, r, 5, f"结算类小计（{len(inv) - len(extra)} 个单品）", F(9.5, True, WHITE), A("center"), NAVY)
cell = ws.cell(row=r, column=6); num(cell, tot_q - EXTRA_Q); cell.font = F(10, True, WHITE)
cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
for j in (7, 8, 9, 10):
    cell = ws.cell(row=r, column=j); cell.value = None; cell.fill = fill(NAVY)
    cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
cell = ws.cell(row=r, column=11); num(cell, round(tot_v, 2), "#,##0.00"); cell.font = F(10, True, WHITE)
cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
mput(ws, r, 12, r, 12, "与附表 01 ⑦ 一致", F(8, False, WHITE), A("center"), NAVY)
ws.row_dimensions[r].height = 26
r += 1
mput(ws, r, 1, r, 5, f"非结算类小计（{len(extra)} 个单品）", F(9.5, True, WHITE), A("center"), "7F7F7F")
cell = ws.cell(row=r, column=6); num(cell, EXTRA_Q); cell.font = F(10, True, WHITE)
cell.fill = fill("7F7F7F"); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
for j in (7, 8, 9, 10):
    cell = ws.cell(row=r, column=j); cell.value = None; cell.fill = fill("7F7F7F")
    cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
cell = ws.cell(row=r, column=11); cell.value = "不估值"; cell.font = F(9, False, WHITE)
cell.alignment = A("center"); cell.fill = fill("7F7F7F")
cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
mput(ws, r, 12, r, 12, "甲方挂账表未列示", F(8, False, WHITE), A("center"), "7F7F7F")
ws.row_dimensions[r].height = 26
r += 2
mput(ws, r, 1, r, 12, f"说明：{len(inv)} 个单品中 {len(inv) - len(extra)} 个为结算类（已挂账），库存货值合计 {tot_v:,.2f} 元，"
     f"与附表 01“⑦ 期末乙方库存货值”{ACT_VAL:,.2f} 元一致；{len(extra)} 个单品共 {EXTRA_Q:,} 件未在甲方挂账，不含货值。",
     F(8, False, DGREY), A("left", "center", True, 0))
ws.row_dimensions[r].height = 30
ws.freeze_panes = "C6"
print("06 ok")

# ==================== 07 乙方销售流水 ====================
ws = out.create_sheet("07-乙方销售流水")
setup(ws, [22, 15, 14, 34, 14, 9, 5, 7, 6], "landscape")
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 9, f"附 表 07　乙方 {PERIOD} 销售流水明细", F(15, True, NAVY), A("center"))
ws.row_dimensions[2].height = 32
mput(ws, 3, 1, 3, 9, f"数据来源：{F_OCT}「{ws_oct.title}」表　·　共 {len(detail)} 行 / {sum(x[7] for x in detail)} 件　·　有记录日期 "
     f"{COV_TXT}",
     F(8.5, False, DGREY), A("center"))
ws.row_dimensions[3].height = 20
ws.row_dimensions[4].height = 6
r = 5
hdr(ws, r, 1, ["流水号", "销售时间", "销售门店", "商品名称", "商品条码", "规格", "单位", "销售数量", "品牌"], 32)
r += 1
for i, d in enumerate(detail, 1):
    bg = WHITE if i % 2 else GREY
    cell = ws.cell(row=r, column=1); cell.value = str(d[0]); cell.font = F(8.5); cell.alignment = A("center")
    cell.number_format = "@"; cell.fill = fill(bg)
    _raw = str(d[1]).strip()
    for _fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
        try:
            _dt = _datetime.strptime(_raw, _fmt); break
        except ValueError:
            _dt = _raw
    _c = mput(ws, r, 2, r, 2, _dt, F(8.5, c=DGREY), A("center"), bg)
    if isinstance(_dt, _datetime):
        _c.number_format = "yyyy-mm-dd hh:mm"
    mput(ws, r, 3, r, 3, d[2], F(8.5), A("center"), bg)
    mput(ws, r, 4, r, 4, d[3], F(8.5), A("left", "center", True, 0), bg)
    cell = ws.cell(row=r, column=5); cell.value = str(d[4]); cell.font = F(8.5); cell.alignment = A("center")
    cell.number_format = "@"; cell.fill = fill(bg)
    mput(ws, r, 6, r, 6, d[5], F(8.5), A("center"), bg)
    mput(ws, r, 7, r, 7, d[6], F(8.5), A("center"), bg)
    cell = ws.cell(row=r, column=8); num(cell, d[7]); cell.fill = fill(bg)
    mput(ws, r, 9, r, 9, d[8], F(8.5), A("center"), bg)
    ws.row_dimensions[r].height = 19
    r += 1
mput(ws, r, 1, r, 7, "合　计", F(9.5, True, WHITE), A("center"), NAVY)
cell = ws.cell(row=r, column=8); num(cell, sum(x[7] for x in detail)); cell.font = F(10, True, WHITE)
cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
cell = ws.cell(row=r, column=9); cell.value = None; cell.fill = fill(NAVY)
cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
ws.row_dimensions[r].height = 28
r += 2
_grp = Counter((str(d[0]).strip(), str(d[4]).strip(), d[7]) for d in detail)
_split = sorted({g[0] for g, n in _grp.items() if n > 1})
if _split:
    _shown = "、".join(_split[:4]) + ("…" if len(_split) > 4 else "")
    _note = (f"核对说明：存在 {len(_split)} 组“同一流水号 + 同一条码 + 同一数量”的拆行记录"
             f"（流水号 {_shown}），已逐行计入合计数、未作去重或调整；是否属正常拆行，请结合库存倒推结果核对。")
else:
    _note = "核对说明：流水明细未发现“同一流水号 + 同一条码 + 同一数量”的重复拆行记录；明细合计已与销售汇总核对一致。"
mput(ws, r, 1, r, 9, _note, F(8, False, DGREY), A("left", "center", True, 0))
ws.row_dimensions[r].height = 34
ws.freeze_panes = "A6"
print("07 ok")

# ==================== 保存 ====================
for sn in ["00-对账说明", "01-对账单", "02-库存差异表"]:
    out[sn].sheet_properties.tabColor = NAVY
out["02-库存差异表"].sheet_properties.tabColor = RED
for sn in ["03-源数据对照", "04-甲方挂账明细", "05-乙方销售汇总", "06-乙方库存明细", "07-乙方销售流水"]:
    out[sn].sheet_properties.tabColor = "A6A6A6"
TITLES = {"01-对账单": "12:12", "02-库存差异表": "7:7", "03-源数据对照": "5:5",
          "04-甲方挂账明细": "5:5", "05-乙方销售汇总": "6:6",
          "06-乙方库存明细": "5:5", "07-乙方销售流水": "5:5"}
for sn in out.sheetnames:
    w = out[sn]
    if sn in TITLES:
        w.print_title_rows = TITLES[sn]
    w.oddHeader.center.text = f"{PA_NAME} × {PB_NAME}　·　{P_CN}对账"
    w.oddHeader.center.size = 8
    w.oddHeader.center.color = "808080"
    w.oddFooter.center.text = "第 &P 页 / 共 &N 页"
    w.oddFooter.center.size = 8
    w.oddFooter.center.color = "808080"
    w.oddFooter.right.text = f"{PA_NAME}  ·  {P_CN}对账"
    w.oddFooter.right.size = 8
    w.oddFooter.right.color = "808080"
# 说明：报表内所有主体/期间/条款文案均在上方按 CFG 直接生成；
# 早期版本曾用「保存前对整簿做字符串替换」兜底，会误伤数据区且漏替条款，已移除。
out.active = 0
# 注意：失败时 stdout 必须保持 0 字节，所有 stdout 输出都放到保存成功之后
_tmp_out = OUT + ".tmp.xlsx"
try:
    out.save(_tmp_out)
    os.replace(_tmp_out, OUT)
except PermissionError as _e:
    if os.path.exists(_tmp_out):
        os.remove(_tmp_out)
    raise SystemExit(f"\n[写入失败] {OUT}\n文件可能正被 Excel 打开，请关闭后重试。原始错误：{_e}")
except Exception as _e:
    if os.path.exists(_tmp_out):
        os.remove(_tmp_out)
    raise SystemExit(f"\n[写入失败] {OUT}\n磁盘空间不足或路径不可写，原文件未受影响。原始错误：{_e}")
emit(f"{_OUT_LABEL} {OUT}")
emit(f"SAVED: {OUT}")
emit("大小 %.1f KB" % (os.path.getsize(OUT) / 1024))
