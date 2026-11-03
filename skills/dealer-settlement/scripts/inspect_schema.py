"""结构探查：识别主档 / 三份导出的表名、字段名、字段顺序漂移，并给出可编辑的映射建议。

用法：
    python3 inspect_schema.py <文件或目录> [--role master|jia|sales|inventory|next_month|auto]
                              [--out 探查.json] [--md 探查.md]

约定：
- 只读：不修改任何源文件；纯确定性、无网络；
- 目录模式按文件名关键词判角色（与 load_master.py 同规则），主档取根目录「_主档*.xlsx」，
  三份导出取第一个含「挂账」的数字月份子目录；
- 发现问题不退出非 0（这是探查工具），只有文件不可读 / 参数不合法才 exit 1；
- --out 输出机器可读 JSON（顶层 proposed_mapping 可直接编辑后交给 apply_mapping.py 确认），
  --md 输出人可读报告；两者都省略时仅在终端打印摘要。
"""
import argparse
import json
import os
import sys

import schema_common as sc

MAX_SCAN = 6          # 表头候选只在最前面 6 行里找
MIN_ASSIGN = 0.34     # 低于该分数视为噪声，不建立字段映射
SHEET_SCORE_FLOOR = 0.5


# ---------------- 基础工具 ----------------
def _issue(code, severity, detail):
    return {"code": code, "severity": severity, "detail": detail}


def _blank(value):
    return value is None or str(value).strip() == ""


def _pad(row, width):
    values = list(row or [])
    if len(values) < width:
        values.extend([None] * (width - len(values)))
    return values


def _sheet_width(matrix):
    return max((len(row) for row in matrix), default=0)


def _header_row_candidates(matrix, fields, max_scan=MAX_SCAN):
    """在前几行里给每行打「能匹配多少个规范字段」的分，按分数降序返回。"""
    limit = min(len(matrix), max_scan)
    candidates = []
    for row_index in range(limit):
        row = matrix[row_index]
        if not any(not _blank(value) for value in row):
            continue
        total = 0.0
        for spec in fields.values():
            best = 0.0
            for value in row:
                if _blank(value):
                    continue
                current = sc.score(value, spec)
                if current > best:
                    best = current
            total += best
        candidates.append({"row": row_index + 1,
                          "score": round(total / max(len(fields), 1), 4)})
    candidates.sort(key=lambda item: (-item["score"], item["row"]))
    return candidates


def _choose_header(matrix, fields, max_scan=MAX_SCAN):
    candidates = _header_row_candidates(matrix, fields, max_scan)
    if not candidates:
        return 0, 0.0, []
    return candidates[0]["row"] - 1, candidates[0]["score"], candidates


def _assign_fields(header_row, fields):
    """按相似度贪心做一对一分配：每个规范字段最多占一列，每列最多被一个字段使用。"""
    field_order = {name: index for index, name in enumerate(fields)}
    pairs = []
    for name, spec in fields.items():
        for column, value in enumerate(header_row):
            if _blank(value):
                continue
            current = sc.score(value, spec)
            if current <= 0:
                continue
            pairs.append((-current, field_order[name], column, name, current))
    pairs.sort()
    assigned, used = {}, set()
    for _neg, _order, column, name, current in pairs:
        if name in assigned or column in used or current < MIN_ASSIGN:
            continue
        assigned[name] = {"index": column, "score": current}
        used.add(column)
    return assigned, used


def _field_report(header_row, fields, order=None):
    """生成 fields / unmatched_fields / extra_columns / order_ok 与字段级问题。

    order 为「参考列序」：默认按规范字段定义顺序；导出文件用与引擎默认列序一致的
    SOURCE_COLUMN_ORDER，避免干净导出被误报 FIELD_ORDER_DIFF。
    """
    order = list(order or fields.keys())
    assigned, used = _assign_fields(header_row, fields)
    reports, issues, unmatched = {}, [], []
    for name, spec in fields.items():
        hit = assigned.get(name)
        scored_columns = []
        for column, value in enumerate(header_row):
            if _blank(value):
                continue
            scored_columns.append((sc.score(value, spec), column))
        scored_columns.sort(key=lambda item: (-item[0], item[1]))
        alternatives = [{"column": sc.column_label(header_row, column), "score": round(value, 4)}
                        for value, column in scored_columns
                        if hit is None or column != hit["index"]][:3]
        if hit is None:
            reports[name] = {"label": spec["label"], "found": None, "score": 0.0,
                             "confidence": "", "alternatives": alternatives[:3]}
            if spec["required"]:
                unmatched.append(name)
                issues.append(_issue(
                    "MISSING_FIELD", "error",
                    "缺少必需字段「%s」（同义词：%s）；表头：%s"
                    % (spec["label"], "/".join(spec["synonyms"]),
                       [sc.column_label(header_row, i) for i in range(len(header_row))
                        if not _blank(header_row[i])])))
            continue
        found = sc.column_label(header_row, hit["index"])
        confidence = sc.confidence_of(hit["score"])
        reports[name] = {"label": spec["label"], "found": found,
                         "score": round(hit["score"], 4), "confidence": confidence,
                         "alternatives": alternatives}
        if hit["score"] < 0.6:
            unmatched.append(name)
        if isinstance(found, str) and sc.normalize(found) != sc.normalize(spec["label"]):
            issues.append(_issue(
                "FIELD_NAME_MISMATCH", "warn",
                "规范字段「%s」现列名为「%s」（相似度 %.2f，置信度%s）"
                % (spec["label"], found, hit["score"], confidence)))

    # 顺序：参考列序下已匹配字段的列序号必须严格递增
    matched_columns = [assigned[name]["index"] for name in order if name in assigned]
    order_ok = all(left < right for left, right in zip(matched_columns, matched_columns[1:]))
    if not order_ok and len(matched_columns) > 1:
        expected = "、".join(
            "%s→%s" % (fields[name]["label"], sc.column_label(header_row, assigned[name]["index"]))
            for name in order if name in assigned)
        actual = "、".join(sc.column_label(header_row, column) for column in matched_columns)
        issues.append(_issue(
            "FIELD_ORDER_DIFF", "warn",
            "字段顺序与规范顺序不一致：规范顺序 %s；表头顺序 %s" % (expected, actual)))

    extra_columns = [sc.column_label(header_row, column) for column in range(len(header_row))
                     if not _blank(header_row[column]) and column not in used]
    for column_label in extra_columns:
        issues.append(_issue("EXTRA_COLUMN", "info", "存在规范外列「%s」（不影响取数，已保留）" % column_label))
    return reports, unmatched, extra_columns, order_ok, issues


def _sheet_rows_cols(workbook, name):
    rows, cols = workbook.dimensions(name)
    return int(rows), int(cols)


def _row_average_score(row, fields):
    """一行表头对给定字段集合的平均匹配度（用于衡量「这张表能不能用」）。"""
    total = 0.0
    for spec in fields.values():
        total += max((sc.score(value, spec) for value in row if not _blank(value)), default=0.0)
    return round(total / max(len(fields), 1), 4)


def _select_source_sheet(workbook, role, fields, required, keywords, exclude):
    """选一张工作表：优先「表名含角色关键词且必需字段可匹配」，否则取必需字段匹配度最高者。

    header_candidates 仍按全部规范字段打分（用于定位表头行）；sheet.score 只按必需字段打，
    避免「次月已销」这类本来就只含少数字段的表被整体拉低分。
    """
    scored = []
    for order, name in enumerate(workbook.sheetnames):
        if any(keyword in name for keyword in exclude):
            continue
        matrix = workbook.matrix(name)
        header_index, _score, candidates = _choose_header(matrix, fields)
        width = _sheet_width(matrix)
        header_row = _pad(matrix[header_index] if matrix else [], width)
        score = _row_average_score(header_row, required)
        keyword_hit = any(keyword in name for keyword in keywords)
        scored.append({"name": name, "order": order, "matrix": matrix,
                       "header_index": header_index, "score": score,
                       "keyword": keyword_hit, "candidates": candidates})
    if not scored:
        scored = [{"name": name, "order": order, "matrix": workbook.matrix(name),
                   "header_index": 0, "score": 0.0, "keyword": False,
                   "candidates": []}
                  for order, name in enumerate(workbook.sheetnames)]
    keyword_pool = [item for item in scored
                    if item["keyword"] and item["score"] >= SHEET_SCORE_FLOOR]
    pool = keyword_pool or scored
    chosen = sorted(pool, key=lambda item: (-item["score"], item["order"]))[0]
    alternatives = [{"name": item["name"], "score": round(item["score"], 4)}
                    for item in sorted(scored, key=lambda item: (-item["score"], item["order"]))
                    if item["name"] != chosen["name"]][:3]
    chosen["alternatives"] = alternatives
    chosen["all"] = scored
    return chosen


def _data_sheets(workbook, exclude):
    names = []
    for name in workbook.sheetnames:
        if any(keyword in name for keyword in exclude):
            continue
        rows, _cols = _sheet_rows_cols(workbook, name)
        if rows >= 2:
            names.append(name)
    return names


# ---------------- 导出文件探查 ----------------
def _inspect_source_file(path, role, display_path):
    fields = sc.SOURCE_ROLES[role]["fields"]
    keywords = sc.SOURCE_ROLES[role]["sheet_keywords"]
    exclude = ("库存",) if role in ("sales", "next_month") else ("销售",) if role == "inventory" else ()
    issues = []
    with sc.read_workbook(path) as workbook:
        sheets = []
        for name in workbook.sheetnames:
            rows, cols = _sheet_rows_cols(workbook, name)
            sheets.append({"name": name, "rows": rows, "cols": cols})
        required = {name: spec for name, spec in fields.items() if spec["required"]}
        chosen = _select_source_sheet(workbook, role, fields, required, keywords, exclude)
        sheet_name = chosen["name"]
        matrix = chosen["matrix"]
        width = _sheet_width(matrix)
        header_row = _pad(matrix[chosen["header_index"]] if matrix else [], width)
        header_row_no = chosen["header_index"] + 1
        sheet_score = chosen["score"]

        sheet_confidence = sc.confidence_of(sheet_score)
        if not any(keyword in sheet_name for keyword in keywords):
            issues.append(_issue(
                "SHEET_NAME_MISMATCH", "error",
                "工作表「%s」未含角色关键词「%s」；已按字段匹配度 %.0f%% 建议映射（置信度%s）"
                % (sheet_name, "/".join(keywords), sheet_score * 100, sheet_confidence)))

        ambiguity = _data_sheets(workbook, exclude)
        if len(ambiguity) > 1:
            issues.append(_issue(
                "MULTI_CANDIDATE_SHEET", "error",
                "含数据的工作表有 %d 张：%s；引擎按关键词无法唯一确定，需映射显式指定「%s」"
                % (len(ambiguity), "、".join(ambiguity), sheet_name)))

        reports, unmatched, extra_columns, order_ok, field_issues = _field_report(
            header_row, fields, order=sc.SOURCE_COLUMN_ORDER[role])
        issues.extend(field_issues)
        if header_row_no != 1:
            issues.append(_issue(
                "HEADER_ROW_OFFSET", "info",
                "表头不在第 1 行（识别为第 %d 行），映射将记录 header_row=%d"
                % (header_row_no, header_row_no)))

        columns = {name: info["found"] for name, info in reports.items() if info["found"] is not None}
        mapping_entry = {"file": display_path, "sheet": sheet_name,
                         "header_row": header_row_no, "columns": columns}

        report = {
            "role": role,
            "path": display_path,
            "sha256": sc.sha256_file(path),
            "sheets": sheets,
            "sheet": {"name": sheet_name, "score": round(sheet_score, 4),
                      "confidence": sheet_confidence, "alternatives": chosen["alternatives"]},
            "header_row": header_row_no,
            "header_candidates": chosen["candidates"],
            "fields": reports,
            "unmatched_fields": unmatched,
            "extra_columns": extra_columns,
            "order_ok": order_ok,
            "issues": issues,
            "proposed_mapping": _source_proposed_mapping(role, mapping_entry, reports),
        }

        if role == "inventory":
            next_month = _inspect_next_month(workbook, display_path)
            if next_month is not None:
                report["next_month"] = next_month["report"]
                mapping_entry["next_month"] = next_month["mapping"]
                report["proposed_mapping"]["sources"][role]["next_month"] = next_month["mapping"]
    return report


def _source_proposed_mapping(role, entry, reports):
    fields = sc.SOURCE_ROLES[role]["fields"]
    notes = []
    for name, info in reports.items():
        if info["found"] is None:
            if fields[name]["required"]:
                notes.append("必需字段未匹配：%s（需人工指定实际列）" % info["label"])
        elif info["confidence"] != sc.CONF_HIGH:
            notes.append("非高置信映射：%s → %s（相似度 %.2f，置信度%s，需人工确认）"
                         % (info["label"], info["found"], info["score"], info["confidence"]))
    return {
        "schema": sc.MAPPING_SCHEMA,
        "confirmed_by": "",
        "confirmed_at": "",
        "notes": notes,
        "master": {"sheets": {}},
        "order": [],
        "sources": {role: entry},
    }


def _inspect_next_month(workbook, display_path):
    """库存导出里的「次月已销」工作表（表名随月变化）：单独给一段映射建议。"""
    candidates = [name for name in workbook.sheetnames if "销售" in name and "库存" not in name]
    if not candidates:
        return None
    fields = sc.SOURCE_ROLES["next_month"]["fields"]
    if len(candidates) > 1:
        scored = []
        for name in candidates:
            matrix = workbook.matrix(name)
            header_index, score, _candidates = _choose_header(matrix, fields)
            scored.append((score, name, header_index, matrix))
        scored.sort(key=lambda item: (-item[0], candidates.index(item[1])))
        score, sheet_name, header_index, matrix = scored[0]
        multi_issue = _issue(
            "MULTI_CANDIDATE_SHEET", "error",
            "次月已销候选工作表有 %d 张：%s；已建议「%s」" % (len(candidates), "、".join(candidates), sheet_name))
    else:
        sheet_name = candidates[0]
        matrix = workbook.matrix(sheet_name)
        header_index, score, _candidates = _choose_header(matrix, fields)
        multi_issue = None
    width = _sheet_width(matrix)
    header_row = _pad(matrix[header_index] if matrix else [], width)
    reports, unmatched, extra_columns, order_ok, field_issues = _field_report(
        header_row, fields, order=sc.SOURCE_COLUMN_ORDER["next_month"])
    issues = list(field_issues)
    if multi_issue:
        issues.insert(0, multi_issue)
    if "销售" not in sheet_name:
        issues.insert(0, _issue("SHEET_NAME_MISMATCH", "error",
                                "次月已销工作表「%s」未含关键词「销售」" % sheet_name))
    columns = {name: info["found"] for name, info in reports.items() if info["found"] is not None}
    mapping = {"sheet": sheet_name, "header_row": header_index + 1, "columns": columns}
    report = {
        "role": "next_month",
        "path": display_path,
        "sheet": sheet_name,
        "header_row": header_index + 1,
        "fields": reports,
        "unmatched_fields": unmatched,
        "extra_columns": extra_columns,
        "order_ok": order_ok,
        "issues": issues,
        "proposed_mapping": mapping,
    }
    return {"report": report, "mapping": mapping}


# ---------------- 主档探查 ----------------
def _master_header_field_specs(canonical):
    return {header: {"label": header, "synonyms": [header]}
            for header in sc.MASTER_HEADERS[canonical]}


def _inspect_master_file(path, display_path):
    issues = []
    with sc.read_workbook(path) as workbook:
        sheet_names = workbook.sheetnames
        sheets = []
        signatures = []
        for name in sheet_names:
            rows, cols = _sheet_rows_cols(workbook, name)
            sheets.append({"name": name, "rows": rows, "cols": cols})
            matrix = workbook.matrix(name)
            signatures.append({"name": name, "matrix": matrix, "rows": rows, "cols": cols})

        # 用「表头内容」把实际工作表配到规范表（表名可能已被改掉）
        pairs = []
        for actual_order, signature in enumerate(signatures):
            for canonical_order, canonical in enumerate(sc.MASTER_SHEET_ORDER):
                fields = _master_header_field_specs(canonical)
                header_index, head_score, _cands = _choose_header(signature["matrix"], fields,
                                                                  max_scan=3)
                bonus = 0.001 if signature["name"] == canonical else 0.0
                pairs.append((-(head_score + bonus), actual_order, canonical_order,
                              signature, canonical, header_index, head_score))
        pairs.sort()
        matched, used_canonical = [], set()
        used_actual = set()
        for _neg, _ao, _co, signature, canonical, header_index, head_score in pairs:
            if signature["name"] in used_actual or canonical in used_canonical:
                continue
            if head_score < 0.6:
                continue
            matched.append({"actual": signature["name"], "canonical": canonical,
                            "header_index": header_index, "score": head_score,
                            "matrix": signature["matrix"]})
            used_actual.add(signature["name"])
            used_canonical.add(canonical)

        matched.sort(key=lambda item: sc.MASTER_SHEET_ORDER.index(item["canonical"]))
        for item in matched:
            if item["actual"] != item["canonical"]:
                issues.append(_issue(
                    "SHEET_NAME_MISMATCH", "error",
                    "工作表「%s」应规范为「%s」（表头匹配度 %.2f）；load_master/引擎按规范表名读取"
                    % (item["actual"], item["canonical"], item["score"])))
        for canonical in sc.MASTER_SHEET_ORDER:
            if canonical not in used_canonical:
                issues.append(_issue(
                    "MISSING_SHEET", "error",
                    "主档缺少规范表「%s」（可能被改名且无法按内容识别），需人工确认" % canonical))
        for signature in signatures:
            if signature["name"] not in used_actual:
                issues.append(_issue(
                    "EXTRA_SHEET", "info",
                    "主档存在规范外工作表「%s」（保留不动）" % signature["name"]))

        master_sheets, all_order_ok, name_mapping = [], True, {}
        for item in matched:
            width = _sheet_width(item["matrix"])
            header_row_no = item["header_index"] + 1
            header_row = _pad(item["matrix"][item["header_index"]] if item["matrix"] else [], width)
            fields = _master_header_field_specs(item["canonical"])
            reports, unmatched, extra_columns, order_ok, field_issues = _field_report(header_row, fields)
            sheet_issues = list(field_issues)
            if header_row_no != 1:
                sheet_issues.append(_issue(
                    "HEADER_ROW_OFFSET", "error",
                    "表头不在第 1 行（识别为第 %d 行）；主档修复暂不支持，请人工把表头放回第 1 行"
                    % header_row_no))
            master_sheets.append({
                "actual": item["actual"], "canonical": item["canonical"],
                "header_row": header_row_no,
                "fields": reports, "unmatched_fields": unmatched,
                "extra_columns": extra_columns, "order_ok": order_ok,
                "issues": sheet_issues,
            })
            issues.extend(sheet_issues)
            if item["actual"] != item["canonical"]:
                name_mapping[item["actual"]] = item["canonical"]
            all_order_ok = all_order_ok and order_ok

        current_order = [item["actual"] for item in matched]
        canonical_order = [item["canonical"] for item in matched]
        if current_order != canonical_order:
            issues.append(_issue(
                "SHEET_ORDER_DIFF", "warn",
                "主档工作表顺序与规范顺序不一致：当前 %s；规范 %s"
                % (current_order, canonical_order)))

        notes = []
        for sheet in master_sheets:
            for name, info in sheet["fields"].items():
                if info["found"] is None:
                    notes.append("主档「%s」缺列：%s（需人工指定）"
                                 % (sheet["canonical"], info["label"]))
                elif info["confidence"] != sc.CONF_HIGH:
                    notes.append("主档「%s」列名非高置信：%s → %s（%.2f，需人工确认）"
                                 % (sheet["canonical"], info["label"], info["found"], info["score"]))
        proposed = {
            "schema": sc.MAPPING_SCHEMA,
            "confirmed_by": "",
            "confirmed_at": "",
            "notes": notes,
            "master": {"sheets": name_mapping},
            "order": list(sc.MASTER_SHEET_ORDER),
            "sources": {},
        }
        report = {
            "role": "master",
            "path": display_path,
            "sha256": sc.sha256_file(path),
            "sheets": sheets,
            "sheet": {},
            "header_row": 1,
            "header_candidates": [],
            "fields": {},
            "unmatched_fields": sorted({name for sheet in master_sheets
                                        for name in sheet["unmatched_fields"]}),
            "extra_columns": sorted({str(column) for sheet in master_sheets
                                     for column in sheet["extra_columns"]}),
            "order_ok": all_order_ok,
            "issues": issues,
            "master_sheets": master_sheets,
            "proposed_mapping": proposed,
        }
    return report


# ---------------- 目录扫描 ----------------
def _month_dirs(directory):
    try:
        return sorted((os.path.join(directory, name) for name in os.listdir(directory)
                       if name.isdigit() and os.path.isdir(os.path.join(directory, name))),
                      reverse=True)
    except OSError:
        return []


def _collect_targets(target, explicit_role):
    """返回 [(文件绝对路径, 角色或 None, 展示用相对路径)] 与输入模式。"""
    if os.path.isfile(target):
        role = explicit_role
        if role in (None, "auto"):
            role = sc.guess_role(target)
        return [(os.path.abspath(target), role, os.path.basename(target))], "file"
    if not os.path.isdir(target):
        raise SystemExit("[路径不存在] %s\n请传入主档/导出文件，或客户根目录。" % target)
    if explicit_role not in (None, "auto"):
        raise SystemExit("[参数错误] 目录模式不接受 --role %r（会同时遇到主档与三份导出）；"
                         "请对单个文件使用 --role。" % explicit_role)
    entries = []
    for name in sc.excel_files(target):
        if sc.MASTER_NAME_HINT in name:
            entries.append((os.path.abspath(os.path.join(target, name)), "master", name))
    # 与 load_master 的期间目录选择一致：优先「数字月份子目录里有挂账文件」的那一组导出
    source_entries = []
    for month_dir in _month_dirs(target):
        month_entries = []
        for name in sc.excel_files(month_dir):
            role = sc.guess_role(name)
            if role:
                month_entries.append((os.path.abspath(os.path.join(month_dir, name)), role,
                                      os.path.join(os.path.basename(month_dir), name)))
        if any(role in ("jia", "sales", "inventory") for _p, role, _r in month_entries):
            source_entries = month_entries
            break
    if not source_entries:
        for name in sc.excel_files(target):
            if sc.MASTER_NAME_HINT in name:
                continue
            role = sc.guess_role(name)
            if role:
                source_entries.append((os.path.abspath(os.path.join(target, name)), role, name))
    entries.extend(source_entries)
    known = {item[0] for item in entries}
    for name in sc.excel_files(target):
        path = os.path.abspath(os.path.join(target, name))
        if path in known or sc.guess_role(name) is not None:
            continue
        if "对账" in name:   # 引擎生成的「YYYYMM铺货代销对账.xlsx」等产物，不是客户导出
            continue
        entries.append((path, "unknown", name))
    role_order = {"master": 0, "jia": 1, "sales": 2, "inventory": 3, "next_month": 4, "unknown": 5}
    entries.sort(key=lambda item: (role_order.get(item[1], 9), item[2]))
    return entries, "directory"


# ---------------- 报告合并 / 输出 ----------------
def _merge_mapping(reports):
    merged = {
        "schema": sc.MAPPING_SCHEMA,
        "confirmed_by": "",
        "confirmed_at": "",
        "notes": [],
        "master": {"sheets": {}},
        "order": [],
        "sources": {},
    }
    for report in reports:
        proposed = report.get("proposed_mapping") or {}
        master = proposed.get("master") or {}
        for actual, canonical in (master.get("sheets") or {}).items():
            merged["master"]["sheets"][actual] = canonical
        if proposed.get("order"):
            merged["order"] = list(proposed["order"])
        for role, entry in (proposed.get("sources") or {}).items():
            merged["sources"][role] = entry
        merged["notes"].extend(proposed.get("notes") or [])
    seen, notes = set(), []
    for note in merged["notes"]:
        if note not in seen:
            seen.add(note)
            notes.append(note)
    merged["notes"] = notes
    return merged


def _summary(report):
    counts = {"error": 0, "warn": 0, "info": 0}
    confidence = {sc.CONF_HIGH: 0, sc.CONF_MEDIUM: 0, sc.CONF_LOW: 0}
    for item in report["files"]:
        for issue in item["issues"]:
            counts[issue["severity"]] = counts.get(issue["severity"], 0) + 1
        for sheet in item.get("master_sheets") or []:
            for info in sheet["fields"].values():
                if info["confidence"]:
                    confidence[info["confidence"]] += 1
        for info in item["fields"].values():
            if info["confidence"]:
                confidence[info["confidence"]] += 1
        next_month = item.get("next_month") or {}
        for info in (next_month.get("fields") or {}).values():
            if info["confidence"]:
                confidence[info["confidence"]] += 1
    return {"files": len(report["files"]), "errors": counts.get("error", 0),
            "warnings": counts.get("warn", 0), "infos": counts.get("info", 0),
            "high": confidence[sc.CONF_HIGH], "medium": confidence[sc.CONF_MEDIUM],
            "low": confidence[sc.CONF_LOW]}


def _render_markdown(report):
    fence = chr(96) * 4
    lines = ["# 结构探查报告", ""]
    summary = report["summary"]
    lines.append("- 输入：%s（%s）" % (report["input"]["target"],
                                    "目录" if report["input"]["mode"] == "directory" else "单文件"))
    lines.append("- 文件数：%d；问题：error %d / warn %d / info %d；字段置信度：高 %d / 中 %d / 低 %d"
                 % (summary["files"], summary["errors"], summary["warnings"], summary["infos"],
                    summary["high"], summary["medium"], summary["low"]))
    lines.append("")
    for index, item in enumerate(report["files"], 1):
        lines.append("## %d. %s（%s）" % (index, item["role"], item["path"]))
        lines.append("")
        lines.append("- sha256：%s" % item["sha256"])
        lines.append("- 工作表：%s" % "、".join(
            "%s(%d×%d)" % (sheet["name"], sheet["rows"], sheet["cols"]) for sheet in item["sheets"]))
        sheet = item.get("sheet") or {}
        if sheet:
            lines.append("- 选用工作表：「%s」（匹配度 %.2f，置信度%s）"
                         % (sheet.get("name"), sheet.get("score", 0.0), sheet.get("confidence", "")))
        if item["role"] == "master":
            for sheet in item.get("master_sheets") or []:
                lines.append("")
                lines.append("### 表「%s」→「%s」（表头第 %d 行）"
                             % (sheet["actual"], sheet["canonical"], sheet["header_row"]))
                lines.append("")
                lines.append("| 规范列 | 实际列 | 相似度 | 置信度 |")
                lines.append("|---|---|---|---|")
                for info in sheet["fields"].values():
                    lines.append("| %s | %s | %.2f | %s |"
                                 % (info["label"],
                                    info["found"] if info["found"] is not None else "（未找到）",
                                    info["score"], info["confidence"] or "-"))
        else:
            lines.append("- 表头行：第 %d 行" % item["header_row"])
            lines.append("")
            lines.append("| 规范字段 | 实际列 | 相似度 | 置信度 |")
            lines.append("|---|---|---|---|")
            for info in item["fields"].values():
                lines.append("| %s | %s | %.2f | %s |"
                             % (info["label"],
                                info["found"] if info["found"] is not None else "（未找到）",
                                info["score"], info["confidence"] or "-"))
        if item.get("unmatched_fields"):
            lines.append("")
            lines.append("- 待人工确认字段：%s" % "、".join(item["unmatched_fields"]))
        if item.get("issues"):
            lines.append("")
            lines.append("### 问题")
            for issue in item["issues"]:
                lines.append("- [%s] %s：%s" % (issue["severity"], issue["code"], issue["detail"]))
        next_month = item.get("next_month")
        if next_month:
            lines.append("")
            lines.append("### 次月已销（同文件）")
            lines.append("")
            lines.append("- 工作表：「%s」，表头第 %d 行"
                         % (next_month["sheet"], next_month["header_row"]))
            lines.append("")
            lines.append("| 规范字段 | 实际列 | 相似度 | 置信度 |")
            lines.append("|---|---|---|---|")
            for info in next_month["fields"].values():
                lines.append("| %s | %s | %.2f | %s |"
                             % (info["label"],
                                info["found"] if info["found"] is not None else "（未找到）",
                                info["score"], info["confidence"] or "-"))
            for issue in next_month.get("issues") or []:
                lines.append("- [%s] %s：%s" % (issue["severity"], issue["code"], issue["detail"]))
        lines.append("")
    lines.append("## 映射建议（可直接编辑后交 apply_mapping.py 确认）")
    lines.append("")
    lines.append(fence + "json")
    lines.append(json.dumps(report["proposed_mapping"], ensure_ascii=False, indent=2))
    lines.append(fence)
    return "\n".join(lines) + "\n"


def _print_summary(report, out_path, md_path):
    summary = report["summary"]
    print("[结构探查] %s（%s）" % (report["input"]["target"],
                                 "目录" if report["input"]["mode"] == "directory" else "单文件"))
    print("  文件 %d 个；error %d / warn %d / info %d；字段置信度：高 %d / 中 %d / 低 %d"
          % (summary["files"], summary["errors"], summary["warnings"], summary["infos"],
             summary["high"], summary["medium"], summary["low"]))
    for item in report["files"]:
        sheet = item.get("sheet") or {}
        if item["role"] == "master":
            sheets = item.get("master_sheets") or []
            print("  - master  %s  已识别 %d/%d 张规范表"
                  % (item["path"], len(sheets), len(sc.MASTER_SHEET_ORDER)))
        elif sheet:
            print("  - %-10s %s  工作表「%s」表头行 %d  匹配度 %.2f"
                  % (item["role"], item["path"], sheet.get("name"), item["header_row"],
                     sheet.get("score", 0.0)))
        elif item["role"] == "unknown":
            print("  - unknown    %s  （文件名未含主档/挂账/销售/库存关键词）" % item["path"])
        if item.get("unmatched_fields"):
            print("    待确认字段：%s" % "、".join(item["unmatched_fields"]))
        codes = ["%s(%s)" % (issue["code"], issue["severity"]) for issue in item["issues"]]
        if codes:
            print("    问题：%s" % "、".join(codes))
        next_month = item.get("next_month")
        if next_month:
            print("    次月已销：工作表「%s」表头第 %d 行"
                  % (next_month["sheet"], next_month["header_row"]))
    mapping = report["proposed_mapping"]
    if mapping["notes"]:
        print("  [提示] 以下映射需用户逐项确认：")
        for note in mapping["notes"]:
            print("    - " + note)
    if out_path:
        print("  映射建议 JSON：%s（顶层 proposed_mapping 可直接编辑）" % out_path)
    if md_path:
        print("  人读报告：%s" % md_path)
    if not out_path and not md_path:
        print("  提示：加 --out 探查.json 可导出可直接编辑的映射建议；"
              "加 --md 探查.md 导出人读报告。")


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="结构探查：识别主档/导出文件的表名、字段名、字段顺序漂移，并给出映射建议（只读）")
    parser.add_argument("target", help="文件或目录（客户根目录 / 月份目录）")
    parser.add_argument("--role", default="auto",
                        choices=["auto", "master", "jia", "sales", "inventory", "next_month"],
                        help="单文件时强制指定角色；目录模式仅支持 auto")
    parser.add_argument("--out", help="机器可读探查 JSON 输出路径")
    parser.add_argument("--md", help="人可读探查报告 Markdown 输出路径")
    args = parser.parse_args(argv)

    target = args.target
    try:
        entries, mode = _collect_targets(target, args.role)
    except SystemExit as exc:
        sys.stderr.write(str(exc) + "\n")
        return 1
    if not entries:
        sys.stderr.write("[没有可探查文件] %s 下没有主档或含关键词的导出 .xls/.xlsx/.xlsm 文件。\n"
                         % target)
        return 1
    if mode == "file" and entries[0][1] is None:
        sys.stderr.write("[无法判断角色] %s\n文件名未包含「主档/挂账/销售数据/销售/库存」关键词；"
                         "请用 --role 指定（master|jia|sales|inventory|next_month）。\n"
                         % os.path.basename(target))
        return 1

    reports = []
    for path, role, display in entries:
        try:
            if role == "master":
                report = _inspect_master_file(path, display)
            elif role == "unknown":
                with sc.read_workbook(path) as workbook:
                    sheets = []
                    for name in workbook.sheetnames:
                        rows, cols = _sheet_rows_cols(workbook, name)
                        sheets.append({"name": name, "rows": rows, "cols": cols})
                report = {"role": "unknown", "path": display, "sha256": sc.sha256_file(path),
                          "sheets": sheets, "sheet": {}, "header_row": 1,
                          "header_candidates": [], "fields": {}, "unmatched_fields": [],
                          "extra_columns": [], "order_ok": True,
                          "issues": [_issue("UNKNOWN_ROLE", "warn",
                                            "无法按文件名判断角色，请用 --role 指定后重跑")],
                          "proposed_mapping": {}}
            else:
                report = _inspect_source_file(path, role, display)
        except (sc.ExcelError, OSError) as exc:
            sys.stderr.write("[探查失败] %s\n%s\n" % (path, exc))
            return 1
        except Exception as exc:  # noqa: BLE001 - 单个文件异常按「不可读」处理
            import traceback
            sys.stderr.write("[探查失败] %s\n%s\n" % (path, exc))
            traceback.print_exc(file=sys.stderr)
            return 1
        reports.append(report)

    merged = _merge_mapping(reports)
    report = {
        "schema": sc.INSPECT_SCHEMA,
        "input": {"target": os.path.abspath(target), "mode": mode},
        "files": reports,
        "proposed_mapping": merged,
    }
    report["summary"] = _summary(report)

    if args.out:
        out_path = os.path.abspath(args.out)
        os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
        with open(out_path, "w", encoding="utf-8") as fh:
            json.dump(report, fh, ensure_ascii=False, indent=2)
            fh.write("\n")
    if args.md:
        md_path = os.path.abspath(args.md)
        os.makedirs(os.path.dirname(md_path) or ".", exist_ok=True)
        with open(md_path, "w", encoding="utf-8") as fh:
            fh.write(_render_markdown(report))
    _print_summary(report, args.out, args.md)
    return 0


if __name__ == "__main__":
    sys.exit(main())
