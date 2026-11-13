"""L6 LLM 复核层的共享常量与工具（review_common.py）。

被 make_review_packet.py / apply_review.py 共同引用，避免两份脚本各写一套常量。
本模块只依赖标准库，不联网、不写文件、不调用任何外部模型。

约定：
- schema 值固定，报告/复核包跨脚本、跨版本可识别；
- R1–R10 清单项与文档逐字一致，脚本不得自行改写（改清单 = 改交付门禁）；
- 时间统一用带时区的 ISO8601（UTC，秒级），保证 JSON 可比较。
"""
import datetime
import hashlib
import json

# ---- schema ----
SCHEMA_PACKET = "dealer-settlement/review-packet@1"
SCHEMA_REPORT = "dealer-settlement/review-report@1"

# ---- 枚举（与文档逐字一致，顺序即展示顺序）----
VERDICTS = ("通过", "有条件通过", "不通过")
LEVELS = ("阻断", "警告", "提示")
RESULTS = ("通过", "不通过", "不适用")

# 交付门禁：只有这两种结论允许进入交付
DELIVERABLE_VERDICTS = ("通过", "有条件通过")

# ---- 复核记录：表号随模块开关变化（9 表 / 10 表 / 11 表都可能）----
REVIEW_SHEET = "10-复核记录"          # 完整 10 表时的规范名，也是历史默认
LEGACY_REVIEW_SHEETS = ("08-复核记录", "09-复核记录")


def review_sheet_name(n_sheets):
    """按实际工作表数量命名（10 张表 -> 10-复核记录），不写死编号。"""
    return "%02d-复核记录" % n_sheets
REVIEW_DIR = "_复核记录"
LOCK_NAME = "复核锁.json"

# ---- R1–R10 复核清单（与文档逐字一致）----
REQUIRED_ITEMS = (
    ("R1", "六条恒等式逐条复算（数量/金额）"),
    ("R2", "合计=明细之和（00/01/02/05/06 跨表一致）"),
    ("R3", "贴花合计独立复算（数量×标准，SCD2 期间正确）"),
    ("R4", "对账单 ①–⑦ 金额结转逐式成立"),
    ("R5", "文字断言与数据一致（结论句中的数字、SKU 数、件数）"),
    ("R6", "交叉引用真实存在（「详见附表 NN」确有该数）"),
    ("R7", "口径披露完整（数据覆盖期间、期初口径、差异率分母、非结算桶）"),
    ("R8", "红线检查（不作归因、不称实盘、不臆断税率、贴花兑付不替用户决定等）"),
    ("R9", "合同条款与主档/合同原文一致（结算方式/基数/异议期/签回期/付款日）"),
    ("R10", "异常与差异解释充分（差异集中在哪些 SKU、是否需人工核对，未编造原因）"),
)
REQUIRED_IDS = tuple(item_id for item_id, _ in REQUIRED_ITEMS)


def sha256_file(path):
    """分块计算文件 SHA-256（十六进制小写）；文件不存在/不可读时抛 OSError。"""
    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_json(path):
    """读取 UTF-8 JSON 文件；解析失败抛 json.JSONDecodeError，文件不存在抛 OSError。"""
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def iso_now():
    """当前时间（UTC、秒级、带时区偏移），形如 2026-11-03T04:13:00+00:00。"""
    return datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")


def parse_iso8601(value):
    """解析 ISO8601 时间字符串；成功返回 datetime（带时区），失败返回 None。

    接受 'Z' 结尾（统一成 +00:00）；不接受纯日期，避免把 '2026-11-02' 当成复核时间。
    """
    if not isinstance(value, str) or not value.strip():
        return None
    text = value.strip()
    if text.endswith(("Z", "z")):
        text = text[:-1] + "+00:00"
    try:
        parsed = datetime.datetime.fromisoformat(text)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        return None
    return parsed


def read_sheet_rows(ws, min_row=1, max_col=None):
    """把工作表读成二维列表（按行、按列取值，保留 None 占位）。

    min_row: 起始行（1 起）；max_col: 最多读取到第几列（None = 工作表已用列数）。
    只读值，不读样式；公式单元格返回公式文本（调用方自行判断是否可用）。
    """
    if ws is None:
        return []
    last_col = max_col or ws.max_column or 0
    rows = []
    for row in ws.iter_rows(min_row=min_row, max_row=ws.max_row, max_col=last_col):
        rows.append([cell.value for cell in row])
    return rows
