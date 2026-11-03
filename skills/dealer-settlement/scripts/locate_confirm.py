"""自动定位用户刚提交的确认单（只读：不修改、不移动任何文件）。

背景：确认页提交后把确认单存进浏览器默认的「下载」文件夹（文件名以
「合同确认单_」开头、带提交时间）。用户回来说一句「好了」后，助手用本脚本
自动找到它，再交给 apply_contract.py 预览并写入主档——用户不需要接触文件。

用法：
    python3 locate_confirm.py [--customer 客户根目录] [--dir DIR ...] [--file 指定路径]
                             [--max-age-minutes 120] [--json]

搜索顺序（先精确、后兜底）：
    1. --file 指定的文件（校验通过即作为唯一结论，不再扫描其它目录）；
    2. --dir 指定的目录（可重复传，如 --dir ~/Downloads --dir ~/Desktop）；
    3. --customer 客户根目录顶层 + 其下的 _确认记录/；
    4. 平台常见目录：~/Downloads、~/下载、~/Desktop、~/桌面（Windows 由
       USERPROFILE 展开，即 %USERPROFILE%\\Downloads、%USERPROFILE%\\Desktop）。

候选规则：
    - 文件名以「合同确认单_」开头、以 .json 结尾（兼容旧版「_YYYY-MM-DD」命名），
      修改时间在 --max-age-minutes 窗口内；
    - 若一个有效候选都没找到，再回退扫描上述目录内修改时间在窗口内、≤5MB 的
      *.json，并逐个做内容校验（避免把其它 JSON 误认成确认单）；
    - 内容校验：schema == "contract-confirm/v1"、meta.operator 非空、fields 是数组；
      不合格的跳过并记入 warnings（不报错、不中断）；
    - 按 sha256 去重（同一份内容只算一份），按修改时间降序（最新在前）。

退出码（凡是要给人/助手看的结果都打印到 stdout）：
    0  恰好 1 份有效：chosen = 该文件；
    2  多份有效：ambiguous = true，chosen = 最新的一份，需回来和用户确认用哪份；
    1  未找到：打印查过哪些目录、可能原因与下一步建议。

--json 输出：
    {"found": bool, "chosen": {...}|null, "ambiguous": bool,
     "candidates": [{"path", "sha256", "size", "mtime", "operator",
                     "confirmed_at", "customer", "fields_count", "source"}],
     "searched_dirs": ["..."], "warnings": ["..."]}
    source 取值：explicit_file / explicit_dir / customer_root /
                customer_confirm_dir / downloads / desktop
纯标准库，Python 3.9+；不依赖任何自定义环境变量。
"""
import argparse
import datetime
import hashlib
import json
import os
import sys
import time

CONFIRM_PREFIX = "合同确认单_"
CONFIRM_SCHEMA = "contract-confirm/v1"
DEFAULT_MAX_AGE_MINUTES = 120
FALLBACK_MAX_BYTES = 5 * 1024 * 1024      # 回退扫描时单个 JSON 的上限
PRIMARY_MAX_BYTES = 30 * 1024 * 1024      # 确认单可能带附件（base64），上限放宽
PLATFORM_DIRS = (
    ("Downloads", "downloads"),
    ("下载", "downloads"),
    ("Desktop", "desktop"),
    ("桌面", "desktop"),
)
SOURCE_LABELS = {
    "explicit_file": "指定文件",
    "explicit_dir": "指定目录",
    "customer_root": "客户目录",
    "customer_confirm_dir": "客户确认记录（_确认记录/）",
    "downloads": "下载文件夹",
    "desktop": "桌面",
}


def _iso_time(epoch):
    return datetime.datetime.fromtimestamp(epoch).isoformat(timespec="seconds")


def _clock_time(epoch):
    return datetime.datetime.fromtimestamp(epoch).strftime("%Y-%m-%d %H:%M:%S")


def _human_size(size):
    if size >= 1024 * 1024:
        return "%.1f MB" % (size / 1048576.0)
    if size >= 1024:
        return "%.0f KB" % (size / 1024.0)
    return "%d B" % size


def validate_confirm(path, size_limit):
    """校验一个文件是否是有效确认单。返回 (candidate, None) 或 (None, 原因)。只读。"""
    try:
        stat = os.stat(path)
    except OSError as exc:
        return None, "无法读取文件信息（%s）" % exc
    if not os.path.isfile(path):
        return None, "不是普通文件"
    if stat.st_size > size_limit:
        return None, "文件过大（%s，上限 %s）" % (_human_size(stat.st_size),
                                                 _human_size(size_limit))
    try:
        with open(path, "rb") as fh:
            raw = fh.read()
        data = json.loads(raw.decode("utf-8-sig"))
    except UnicodeDecodeError:
        return None, "不是 UTF-8 编码的文本"
    except json.JSONDecodeError as exc:
        return None, "不是合法 JSON（第 %d 行第 %d 列：%s）" % (exc.lineno, exc.colno, exc.msg)
    except OSError as exc:
        return None, "读取失败（%s）" % exc
    if not isinstance(data, dict):
        return None, "顶层不是对象"
    if data.get("schema") != CONFIRM_SCHEMA:
        return None, "schema 不是 %s（实际：%r）" % (CONFIRM_SCHEMA, data.get("schema"))
    meta = data.get("meta")
    if not isinstance(meta, dict):
        return None, "缺少 meta 对象"
    operator = str(meta.get("operator") or "").strip()
    if not operator:
        return None, "缺少确认人（meta.operator 为空）"
    fields = data.get("fields")
    if not isinstance(fields, list):
        return None, "fields 不是数组"
    return {
        "path": os.path.abspath(path),
        "sha256": hashlib.sha256(raw).hexdigest(),
        "size": stat.st_size,
        "mtime": _iso_time(stat.st_mtime),
        "operator": operator,
        "confirmed_at": str(meta.get("confirmed_at") or "").strip(),
        "customer": str(meta.get("customer") or "").strip(),
        "fields_count": len(fields),
        "source": "",
        "_mtime": stat.st_mtime,
    }, None


def collect_dirs(args):
    """按优先级返回待搜索目录 [{path, source, warn_missing}]（展开 ~、去重）。"""
    out, seen = [], set()

    def add(path, source, warn_missing):
        resolved = os.path.abspath(os.path.expanduser(path))
        if resolved in seen:
            return
        seen.add(resolved)
        out.append({"path": resolved, "source": source, "warn_missing": warn_missing})

    for directory in args.dir or []:
        add(directory, "explicit_dir", True)
    if args.customer:
        root = os.path.abspath(os.path.expanduser(args.customer))
        add(root, "customer_root", True)
        add(os.path.join(root, "_确认记录"), "customer_confirm_dir", False)
    home = os.path.expanduser("~")
    for name, source in PLATFORM_DIRS:
        add(os.path.join(home, name), source, False)
    return out


def _list_json_files(directory):
    """列出目录顶层的 *.json（不递归）。目录不可读时返回 None。"""
    try:
        names = sorted(os.listdir(directory))
    except OSError:
        return None
    files = []
    for name in names:
        path = os.path.join(directory, name)
        if name.lower().endswith(".json") and os.path.isfile(path):
            files.append(path)
    return files


def scan(args):
    """执行搜索。返回 (result, context)。"""
    now = time.time()
    window = max(0, args.max_age_minutes) * 60.0
    warnings = []
    candidates = []
    searched_dirs = []
    stale = []              # 文件名符合、但修改时间早于时间窗口
    json_per_dir = {}

    def add_warning(message):
        if message not in warnings:
            warnings.append(message)

    # 1) --file：显式指定的文件优先，校验通过即作为唯一结论
    if args.file:
        explicit = os.path.abspath(os.path.expanduser(args.file))
        candidate, why = validate_confirm(explicit, PRIMARY_MAX_BYTES)
        if candidate is not None:
            candidate["source"] = "explicit_file"
            candidates.append(candidate)
            return _build_result(candidates, [], warnings), {
                "stale": [], "json_per_dir": {}, "window": window,
                "max_age_minutes": args.max_age_minutes,
            }
        add_warning("--file 指定的文件不可用：%s（%s）；继续按目录搜索" % (explicit, why))

    # 2)~4) 目录：先存在性过滤，再统一列一次文件清单
    dirs = []
    for item in collect_dirs(args):
        path = item["path"]
        if not os.path.isdir(path):
            if item["warn_missing"]:
                add_warning("目录不存在，已跳过：%s" % path)
            continue
        files = _list_json_files(path)
        if files is None:
            add_warning("目录无法读取，已跳过：%s" % path)
            continue
        dirs.append(item)
        searched_dirs.append(path)
        json_per_dir[path] = len(files)
        item["files"] = files

    if not dirs:
        add_warning("没有任何可搜索的目录")

    def consider(path, source, size_limit):
        candidate, why = validate_confirm(path, size_limit)
        if candidate is None:
            return "invalid", why
        if candidate["sha256"] in {c["sha256"] for c in candidates}:
            return "duplicate", None
        candidate["source"] = source
        candidates.append(candidate)
        return "ok", None

    # 主扫描：只认「合同确认单_*.json」，并记下过期的确认单供提示
    for item in dirs:
        for path in item["files"]:
            if not os.path.basename(path).startswith(CONFIRM_PREFIX):
                continue
            try:
                mtime = os.stat(path).st_mtime
            except OSError:
                continue
            if now - mtime > window:
                stale.append((path, mtime))
                continue
            status, why = consider(path, item["source"], PRIMARY_MAX_BYTES)
            if status == "invalid":
                add_warning("跳过（%s）：%s" % (why, path))

    # 回退扫描：一份有效确认单都没找到时，校验目录内窗口期内的 *.json
    fallback_warned = 0
    if not candidates:
        for item in dirs:
            for path in item["files"]:
                if os.path.basename(path).startswith(CONFIRM_PREFIX):
                    continue        # 主扫描已处理
                try:
                    mtime = os.stat(path).st_mtime
                except OSError:
                    continue
                if now - mtime > window:
                    continue
                candidate, why = validate_confirm(path, FALLBACK_MAX_BYTES)
                if candidate is None:
                    if fallback_warned < 5:
                        add_warning("回退扫描跳过（%s）：%s" % (why, path))
                        fallback_warned += 1
                    continue
                if candidate["sha256"] in {c["sha256"] for c in candidates}:
                    continue
                candidate["source"] = item["source"]
                candidates.append(candidate)
        if fallback_warned == 5:
            add_warning("回退扫描还有更多非确认单 JSON 被跳过（不再逐条列出）")

    candidates.sort(key=lambda c: c["_mtime"], reverse=True)
    return _build_result(candidates, searched_dirs, warnings), {
        "stale": stale, "json_per_dir": json_per_dir, "window": window,
        "max_age_minutes": args.max_age_minutes,
    }


def _build_result(candidates, searched_dirs, warnings):
    public = []
    for candidate in candidates:
        public.append({key: value for key, value in candidate.items()
                       if not key.startswith("_")})
    return {
        "found": bool(public),
        "chosen": public[0] if public else None,
        "ambiguous": len(public) > 1,
        "candidates": public,
        "searched_dirs": searched_dirs,
        "warnings": warnings,
    }


def print_human(result, context):
    max_age = context["max_age_minutes"]
    chosen = result["chosen"]
    if not result["found"]:
        print("[未找到确认单] 时间窗口 %s 分钟内没有找到有效确认单" % max_age)
        if result["searched_dirs"]:
            print("已查目录（%d 个）：" % len(result["searched_dirs"]))
            for directory in result["searched_dirs"]:
                count = context["json_per_dir"].get(directory, 0)
                print("  - %s（%d 个 json 文件）" % (directory, count))
        else:
            print("已查目录：无（没有可搜索的目录）")
        for warning in result["warnings"]:
            print("[提示] " + warning)
        if context["stale"]:
            newest_path, newest_mtime = max(context["stale"], key=lambda item: item[1])
            print("另外：有 %d 份名称符合的确认单早于时间窗口（最新：%s，%s）。"
                  % (len(context["stale"]), newest_path, _clock_time(newest_mtime)))
        print("可能原因：")
        print("  - 确认页还没有点「提交并保存确认单」；")
        print("  - 浏览器把文件存到了其他位置（不是「下载」文件夹）；")
        print("  - 确认单的保存时间早于时间窗口（可用 --max-age-minutes 调大，如 1440）。")
        print("下一步：先请用户确认已在确认页提交；若知道保存位置，用 --dir <目录> 或 "
              "--file <文件> 指定后重跑。")
        return

    if result["ambiguous"]:
        print("[确认单有多个候选] 找到 %d 份有效确认单，需要和用户确认用哪一份"
              "（已按修改时间从新到旧）：" % len(result["candidates"]))
        for index, candidate in enumerate(result["candidates"], 1):
            mark = " [最新]" if index == 1 else ""
            print("  %d.%s %s" % (index, mark, candidate["path"]))
            print("     确认人：%s　确认时间：%s　项数：%d　来源：%s"
                  % (candidate["operator"], candidate["confirmed_at"] or "未记录",
                     candidate["fields_count"],
                     SOURCE_LABELS.get(candidate["source"], candidate["source"])))
        print("提示：最新一份通常就是刚提交的；若用户明确指定，可用 --file <路径> 重跑，例如：")
        print('  python3 "%s" --file "%s"'
              % (os.path.abspath(__file__), result["candidates"][0]["path"]))
        for warning in result["warnings"]:
            print("[提示] " + warning)
        return

    print("[找到确认单] 1 份有效")
    print("  文件    : %s" % chosen["path"])
    print("  确认人  : %s" % chosen["operator"])
    print("  确认时间: %s" % (chosen["confirmed_at"] or "未记录"))
    print("  项数    : %d" % chosen["fields_count"])
    print("  客户    : %s" % (chosen["customer"] or "未记录"))
    print("  来源    : %s" % SOURCE_LABELS.get(chosen["source"], chosen["source"]))
    print("  修改时间: %s" % chosen["mtime"])
    for warning in result["warnings"]:
        print("[提示] " + warning)
    print("下一步：先预览差异，用户确认无误后再加 --write 写入主档：")
    print('  python3 "%s" "<客户根目录>/_主档.xlsx" "%s"'
          % (os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "apply_contract.py"), chosen["path"]))


class _Parser(argparse.ArgumentParser):
    """用法错误统一 exit 1（2 保留给「多份确认单」的歧义结果）。"""

    def error(self, message):
        self.print_usage(sys.stderr)
        sys.stderr.write("%s：错误：%s\n" % (self.prog, message))
        sys.exit(1)


def main(argv=None):
    parser = _Parser(description="自动定位用户刚提交的合同确认单（只读）")
    parser.add_argument("--customer", metavar="客户根目录",
                        help="客户根目录（搜索其顶层与 _确认记录/）")
    parser.add_argument("--dir", action="append", default=[], metavar="DIR",
                        help="额外搜索目录，可重复传")
    parser.add_argument("--file", metavar="路径",
                        help="直接校验指定文件（有效即以它为准）")
    parser.add_argument("--max-age-minutes", type=int, default=DEFAULT_MAX_AGE_MINUTES,
                        metavar="N",
                        help="只看最近 N 分钟内保存的确认单（默认 %d）"
                        % DEFAULT_MAX_AGE_MINUTES)
    parser.add_argument("--json", action="store_true", dest="as_json",
                        help="输出机器可读 JSON（stdout 只有 JSON）")
    args = parser.parse_args(argv)
    if args.max_age_minutes < 0:
        parser.error("--max-age-minutes 不能为负数")

    result, context = scan(args)
    count = len(result["candidates"])
    if args.as_json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print_human(result, context)
    return 0 if count == 1 else (2 if count > 1 else 1)


if __name__ == "__main__":
    sys.exit(main())
