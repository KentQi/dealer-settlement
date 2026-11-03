"""结构漂移识别与数据映射的公共定义（规范字段 / 归一化评分 / Excel 只读读取）。

被 inspect_schema.py、apply_mapping.py 与 recon_engine.py 复用，保证三处对「规范字段」
和「列匹配」的理解完全一致：
- 主档六表的规范表头直接复用 make_master.SHEETS（单一事实来源，不复制表头）；
- 三份导出按业务字段定义规范名与同义词，用于模糊匹配、置信度分级与映射建议；
- 只读工具：sha256、JSON 读取、时间戳、Excel 懒加载读取器（.xlsx/.xlsm → openpyxl，
  .xls → xlrd，仅在真正读取时才 import 与打开文件）。

本模块不写任何文件、不访问网络；被 CLI 与引擎共享。Python 3.9+。
"""
import hashlib
import json
import os
import re
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher

# 复用主档模板定义（make_master.py 与本文件同在 scripts/ 目录）
from make_master import SHEETS

# ---------------- 固定命名与路径 ----------------
MAPPING_SCHEMA = "dealer-settlement/data-mapping@1"
INSPECT_SCHEMA = "dealer-settlement/schema-inspect@1"
MAPPING_FILENAME = "_数据映射.json"
CONFIRM_DIR = "_确认记录"
BACKUP_DIR = "_备份"
MASTER_NAME_HINT = "主档"
EXCEL_EXTS = (".xls", ".xlsx", ".xlsm")

# 主档：规范表头 / 列宽 / 表顺序全部来自 make_master.SHEETS
MASTER_HEADERS = {name: list(spec[0]) for name, spec in SHEETS.items()}
MASTER_WIDTHS = {name: list(spec[1]) for name, spec in SHEETS.items()}
MASTER_SHEET_ORDER = list(SHEETS.keys())

# 置信度分级：>=0.85 高；>=0.6 中；否则低
CONF_HIGH, CONF_MEDIUM, CONF_LOW = "高", "中", "低"


def confidence_of(value):
    """分数 → 置信度标签。"""
    value = float(value or 0.0)
    if value >= 0.85:
        return CONF_HIGH
    if value >= 0.6:
        return CONF_MEDIUM
    return CONF_LOW


# ---------------- 三份导出的规范字段 ----------------
# 每个字段：label=规范名（报告展示用）；synonyms=表头同义词；required=缺失即阻断字段。
# synonyms 在任务书给定集合之外，补充了现有引擎 _col() 识别的关键词（如「销售时间」），
# 保证「干净导出」探查结果不被误报为漂移。
SOURCE_ROLES = {
    "jia": {
        "label": "甲方挂账",
        "sheet_keywords": ["挂账"],
        "fields": {
            "name": {"label": "商品名称", "synonyms": ["商品名称", "商品", "品名"], "required": True},
            "price": {"label": "单价", "synonyms": ["单价", "供货价"], "required": True},
            "amount": {"label": "未结算金额", "synonyms": ["未结算金额", "金额"], "required": True},
            "qty": {"label": "未结算数量", "synonyms": ["未结算数量", "数量"], "required": True},
            "init": {"label": "初始数量", "synonyms": ["初始数量", "初始"], "required": False},
        },
    },
    "sales": {
        "label": "乙方销售流水",
        "sheet_keywords": ["月"],
        "fields": {
            "bar": {"label": "商品条码", "synonyms": ["商品条码", "条码"], "required": True},
            "name": {"label": "商品名称", "synonyms": ["商品名称", "品名", "商品"], "required": True},
            "qty": {"label": "销售数量", "synonyms": ["数量", "销售数量"], "required": True},
            "date": {"label": "销售时间", "synonyms": ["销售日期", "日期", "单据日期", "销售时间", "时间"],
                     "required": False},
            "store": {"label": "销售门店", "synonyms": ["门店", "门店名称", "客户", "销售门店"], "required": False},
            "serial": {"label": "流水号", "synonyms": ["流水号", "单号"], "required": True},
            "unit": {"label": "单位", "synonyms": ["单位"], "required": False},
            "spec": {"label": "规格", "synonyms": ["规格"], "required": False},
            "brand": {"label": "品牌", "synonyms": ["品牌"], "required": False},
        },
    },
    "inventory": {
        "label": "乙方库存",
        "sheet_keywords": ["库存"],
        "fields": {
            "bar": {"label": "商品条码", "synonyms": ["商品条码", "条码"], "required": True},
            "name": {"label": "商品名称", "synonyms": ["商品名称", "品名"], "required": True},
            "spec": {"label": "规格", "synonyms": ["规格"], "required": False},
            "cat": {"label": "商品分类", "synonyms": ["分类", "商品分类"], "required": False},
            "qty": {"label": "库存数量", "synonyms": ["库存", "库存数量", "数量"], "required": True},
            "retail": {"label": "销售价", "synonyms": ["销售价", "零售价"], "required": False},
            "brand": {"label": "品牌", "synonyms": ["品牌"], "required": False},
        },
    },
    "next_month": {
        "label": "乙方次月已销",
        "sheet_keywords": ["销售"],
        "fields": {
            "bar": {"label": "商品条码", "synonyms": ["商品条码", "条码"], "required": True},
            "name": {"label": "商品名称", "synonyms": ["商品名称", "品名", "商品"], "required": True},
            "qty": {"label": "销售数量", "synonyms": ["数量", "销售数量"], "required": True},
            "date": {"label": "销售时间", "synonyms": ["销售日期", "日期", "单据日期", "销售时间", "时间"],
                     "required": False},
            "store": {"label": "销售门店", "synonyms": ["门店", "门店名称", "客户", "销售门店"], "required": False},
            "serial": {"label": "流水号", "synonyms": ["流水号", "单号"], "required": False},
            "unit": {"label": "单位", "synonyms": ["单位"], "required": False},
            "spec": {"label": "规格", "synonyms": ["规格"], "required": False},
            "brand": {"label": "品牌", "synonyms": ["品牌"], "required": False},
        },
    },
}

# 兼容现有引擎默认列序的「参考列序」：干净导出不误报 FIELD_ORDER_DIFF，
# 也只有按同一参考列序才能判断「列顺序被打乱」。
SOURCE_COLUMN_ORDER = {
    "jia": ["name", "price", "amount", "qty", "init"],
    "sales": ["serial", "date", "store", "name", "bar", "spec", "unit", "qty", "brand"],
    "inventory": ["name", "bar", "spec", "cat", "qty", "retail", "brand"],
    "next_month": ["name", "bar", "qty"],
}

SOURCE_ROLE_ORDER = ["jia", "sales", "inventory"]
VALID_SOURCE_ROLES = SOURCE_ROLE_ORDER + ["next_month"]

# 目录模式的文件→角色规则：与 load_master.py 的发现顺序一致。
# (角色, 必须包含, 必须不包含)；命中顺序即优先级。
ROLE_FILE_RULES = (
    ("master", (MASTER_NAME_HINT,), ()),
    ("jia", ("挂账",), ()),
    ("sales", ("销售数据",), ()),
    ("sales", ("销售",), ("库存",)),
    ("inventory", ("库存",), ()),
)

_PUNCT_RE = re.compile(
    r"[\s\u3000()\[\]{}<>:：,，.。;；'\"“”‘’_\-—/\\|·*、~！!？?]")
_SUFFIX_RE = re.compile(r"\(限定版\)$")


def normalize(text):
    """表头归一化：全角→半角、去空白与标点括号、转小写、去「（限定版）」后缀。"""
    if text is None:
        return ""
    value = unicodedata.normalize("NFKC", str(text))
    value = value.strip().lower()
    value = _SUFFIX_RE.sub("", value)   # NFKC 已把全角括号转半角，先去后缀再删标点
    return _PUNCT_RE.sub("", value)


def score(text, field):
    """表头与规范字段的相似度（0~1）。

    同义词归一化后完全命中 → 1.0；否则用 difflib.SequenceMatcher，并给「包含关系」兜底
    （如「销售门店」包含「门店」）0.9，避免同义扩展词被判成低置信。
    """
    synonyms = field["synonyms"] if isinstance(field, dict) else list(field or [])
    head = normalize(text)
    if not head:
        return 0.0
    best = 0.0
    for synonym in synonyms:
        syn = normalize(synonym)
        if not syn:
            continue
        if head == syn:
            return 1.0
        ratio = SequenceMatcher(None, head, syn).ratio()
        if len(syn) >= 2 and (syn in head or head in syn):
            ratio = max(ratio, 0.9)
        if ratio > best:
            best = ratio
    return round(best, 4)


def find_column(headers, spec):
    """在表头列表中按映射值找列序号（0 基）；找不到返回 None。

    spec 允许两种形态：列名（字符串，先精确后归一化匹配）或列序号（整数，Excel 列号，1 基）。
    inspect_schema 与 apply_mapping 用同一函数校验，recon_engine 也复用它按映射取数。
    """
    if isinstance(spec, bool) or spec is None:
        return None
    if isinstance(spec, int):
        index = spec - 1
        return index if 0 <= index < len(headers) else None
    if not isinstance(spec, str):
        return None
    want = spec.strip()
    if not want:
        return None
    for index, head in enumerate(headers):
        if head is not None and str(head).strip() == want:
            return index
    want_norm = normalize(want)
    for index, head in enumerate(headers):
        if head is not None and normalize(head) == want_norm:
            return index
    return None


def column_label(headers, index):
    """列序号（0 基）→ 可回写的列标识：有表头用列名，空表头用 Excel 列号（1 基）。"""
    if index is None or not 0 <= index < len(headers):
        return None
    head = headers[index]
    if head is None or str(head).strip() == "":
        return index + 1
    return str(head).strip()


def match_columns(header_row, wanted_headers, min_score=0.34):
    """把规范列名一对一匹配到实际表头列：返回 {规范列名: (列序号, 分数)}。

    贪心按分数从高到低分配，精确同名优先，用于主档修复（列名不规范 → 规范列名）。
    """
    pairs = []
    for order, wanted in enumerate(wanted_headers):
        for index, value in enumerate(header_row):
            if value is None or str(value).strip() == "":
                continue
            current = score(value, {"label": wanted, "synonyms": [wanted]})
            if current <= 0:
                continue
            pairs.append((-current, order, index, wanted, current))
    pairs.sort()
    matched, used = {}, set()
    for _neg, _order, index, wanted, current in pairs:
        if wanted in matched or index in used or current < min_score:
            continue
        matched[wanted] = (index, current)
        used.add(index)
    return matched


# ---------------- 目录角色识别 ----------------
def guess_role(filename):
    """按文件名关键词判角色（规则与 load_master.py 相同）；识别不到返回 None。"""
    base = os.path.basename(str(filename or ""))
    for role, must, deny in ROLE_FILE_RULES:
        if any(key in base for key in must) and not any(key in base for key in deny):
            return role
    return None


# ---------------- 只读小工具 ----------------
def sha256_file(path, chunk=1 << 20):
    """文件 sha256（用于「永不修改源导出」的可复核证明）。"""
    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        while True:
            block = fh.read(chunk)
            if not block:
                break
            digest.update(block)
    return digest.hexdigest()


def load_json(path):
    """读取 UTF-8（含 BOM）JSON 文件；失败时抛 ValueError/OSError 由调用方转中文报错。"""
    with open(path, "r", encoding="utf-8-sig") as fh:
        return json.load(fh)


def iso_now():
    """本地时区 ISO 时间（秒级），用于 confirmed_at。"""
    return datetime.now().astimezone().isoformat(timespec="seconds")


def ms_stamp(moment=None):
    """毫秒级时间戳文件名片段：YYYYmmdd_HHMMSS_mmm（备份/确认记录用）。"""
    moment = moment or datetime.now()
    return moment.strftime("%Y%m%d_%H%M%S_") + "%03d" % (moment.microsecond // 1000)


def mapping_path(root):
    return os.path.join(root, MAPPING_FILENAME)


def excel_files(directory):
    """目录下的 Excel 文件（排除 Office 临时文件），按名称排序，保持确定性。"""
    try:
        return sorted(name for name in os.listdir(directory)
                      if name.lower().endswith(EXCEL_EXTS) and not name.startswith("~$"))
    except OSError:
        return []


class ExcelError(RuntimeError):
    """Excel 读取/结构错误：消息为中文，直接面向用户。"""


class ExcelReader:
    """按扩展名懒加载的只读 Excel 读取器。

    .xlsx/.xlsm → openpyxl(read_only=True, data_only=True)；.xls → xlrd。
    openpyxl / xlrd 只在首次真正读取时才 import 与打开文件。
    """

    def __init__(self, path):
        self.path = os.path.abspath(path)
        self.ext = os.path.splitext(self.path)[1].lower()
        self._book = None
        self._kind = None

    # -- 懒加载 --
    def _load(self):
        if self._book is not None:
            return
        if self.ext in (".xlsx", ".xlsm"):
            import openpyxl
            try:
                self._book = openpyxl.load_workbook(self.path, data_only=True, read_only=True)
            except Exception as exc:  # noqa: BLE001 - 统一转中文错误
                raise ExcelError("[无法读取] %s\n原因：%s\n（.xlsx/.xlsm 需可被 openpyxl 打开）"
                                 % (self.path, exc))
            self._kind = "openpyxl"
            return
        if self.ext == ".xls":
            try:
                import xlrd
            except ImportError:
                raise ExcelError("[缺少依赖] 读取 .xls 需要 xlrd（python3 -m pip install xlrd），"
                                 "或请将文件另存为 .xlsx 后重试。")
            try:
                self._book = xlrd.open_workbook(self.path)
            except Exception as exc:  # noqa: BLE001
                raise ExcelError("[无法读取] %s\n原因：%s\n（.xls 需可被 xlrd 打开）" % (self.path, exc))
            self._kind = "xlrd"
            return
        raise ExcelError("[不支持的文件格式] %s\n仅支持 .xls / .xlsx / .xlsm。" % self.path)

    def _sheet_names_loaded(self):
        if self._kind == "openpyxl":
            return list(self._book.sheetnames)
        return [self._book.sheet_by_index(i).name for i in range(self._book.nsheets)]

    @property
    def sheetnames(self):
        self._load()
        return self._sheet_names_loaded()

    def _pick(self, sheet):
        names = self.sheetnames
        if not names:
            raise ExcelError("[空工作簿] %s 没有任何工作表。" % self.path)
        if sheet is None or sheet == "":
            return names[0]
        if isinstance(sheet, int) and not isinstance(sheet, bool):
            if not 1 <= sheet <= len(names):
                raise ExcelError("[工作表不存在] %s 只有 %d 张表，映射指定的序号 %r 越界。"
                                 % (self.path, len(names), sheet))
            return names[sheet - 1]
        if str(sheet) not in names:
            raise ExcelError("[工作表不存在] %s 中没有工作表「%s」；实际工作表：%s"
                             % (self.path, sheet, names))
        return str(sheet)

    def matrix(self, sheet=None):
        """返回指定工作表（默认第一张）的二维列表，值为单元格真实值。"""
        self._load()
        name = self._pick(sheet)
        if self._kind == "openpyxl":
            worksheet = self._book[name]
            return [list(row) for row in worksheet.iter_rows(values_only=True)]
        sh = self._book.sheet_by_name(name)
        return [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]

    def dimensions(self, sheet=None):
        """返回 (行数, 列数)；openpyxl 只读模式拿不到维度时退回遍历。"""
        self._load()
        name = self._pick(sheet)
        if self._kind == "openpyxl":
            worksheet = self._book[name]
            rows, cols = worksheet.max_row, worksheet.max_column
            if rows is None or cols is None:
                data = self.matrix(name)
                return len(data), max((len(row) for row in data), default=0)
            return int(rows), int(cols)
        sh = self._book.sheet_by_name(name)
        return int(sh.nrows), int(sh.ncols)

    def close(self):
        if self._book is not None and self._kind == "openpyxl":
            try:
                self._book.close()
            except Exception:  # noqa: BLE001 - 关闭失败不掩盖主流程
                pass
        self._book = None
        self._kind = None

    def __enter__(self):
        return self

    def __exit__(self, *exc_info):
        self.close()
        return False


def excel_reader(path):
    """便捷入口：返回未立即打开的 ExcelReader，调用方用 with 或 close()。"""
    return ExcelReader(path)


# 兼容内部旧名（与 excel_reader 等价）
read_workbook = excel_reader


# ---------------- 映射结构口径 ----------------
def required_fields(role):
    """角色下缺失即阻断的规范字段列表。"""
    return [name for name, cfg in SOURCE_ROLES[role]["fields"].items() if cfg["required"]]
