"""映射校验 + 用户确认落盘 + 主档结构修复。

用法：
    python3 apply_mapping.py <客户根目录> <映射.json> [--dry-run | --write]
                             [--confirmed-by 姓名] [--fix-master | --fix-order]

行为：
- 默认 dry-run：逐项核对映射里的文件 / 工作表 / 表头行 / 列在客户目录中真实存在，
  打印将写入或修改的内容与影响；任一项不存在 → 中文报错 exit 1，且不写任何文件。
- --write：必须同时给出 --confirmed-by（用户逐项确认后的姓名/标识），否则 exit 1；
  写 <客户根目录>/_数据映射.json（含 schema / confirmed_by / confirmed_at / notes），
  并把副本归档到 <客户根目录>/_确认记录/映射_YYYYmmdd_HHMMSS_mmm.json。
- --fix-master（可加 --fix-order）：修复主档：表名 → 规范表名、列名 → 规范列名；
  --fix-order 追加：按规范顺序重排列（并在 --fix-order 时按 order 重排工作表）。
  先备份到 _备份/（毫秒时间戳副本），再写临时文件并原子替换；dry-run 只打印计划。
- 永不修改三份源导出：只读打开核对，只写映射文件与主档修复副本。
"""
import argparse
import copy
import json
import os
import re
import shutil
import stat
import sys

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

import schema_common as sc

MASTER_MIN_SCORE = 0.6
SOURCE_FILE_RULES = {
    "jia": ((("挂账", None),), "甲方挂账文件"),
    "sales": ((("销售数据", None), ("销售", "库存")), "销售流水文件"),
    "inventory": ((("库存", None),), "库存文件"),
}


class ApplyError(RuntimeError):
    """可直接展示给用户的中文错误。"""


# ---------------- 通用工具 ----------------
def _blank(value):
    return value is None or str(value).strip() == ""


def _under(root, path):
    """路径展示：位于客户根目录下时用相对路径，保持报告稳定可读。"""
    absolute = os.path.abspath(path)
    root = os.path.abspath(root)
    if absolute == root:
        return "."
    prefix = root + os.sep
    if absolute.startswith(prefix):
        return os.path.relpath(absolute, root)
    return absolute


def _unique_path(directory, stem, ext):
    """在目录下生成不与现有文件冲突的路径（毫秒时间戳 + 极端情况下追加序号）。"""
    base = os.path.join(directory, "%s_%s" % (stem, sc.ms_stamp()))
    candidate = base + ext
    counter = 2
    while os.path.exists(candidate):
        candidate = "%s_%d%s" % (base, counter, ext)
        counter += 1
    return candidate


def _discover_master(root):
    matches = [name for name in sc.excel_files(root) if sc.MASTER_NAME_HINT in name]
    if not matches:
        return None
    if len(matches) > 1:
        raise ApplyError("[主档定位失败] 客户根目录下有多份主档：%s；请只保留一份（如 _主档.xlsx）。"
                         % matches)
    return os.path.join(root, matches[0])


def _discover_sources(root):
    """按 load_master.py 同规则发现三份导出；返回 ({role: 绝对路径}, 月份目录)。"""
    root_files = sc.excel_files(root)

    def has_jia(names):
        return any("挂账" in name for name in names)

    base, files = None, []
    try:
        month_names = sorted((name for name in os.listdir(root)
                              if name.isdigit() and os.path.isdir(os.path.join(root, name))),
                             reverse=True)
    except OSError:
        month_names = []
    for name in month_names:
        directory = os.path.join(root, name)
        candidate_files = sc.excel_files(directory)
        if has_jia(candidate_files):
            base, files = directory, candidate_files
            break
    if base is None and has_jia(root_files):
        base, files = root, root_files
    if base is None:
        return {}, None
    discovered = {}
    for role, (rules, _label) in SOURCE_FILE_RULES.items():
        for must, deny in rules:
            hits = [name for name in files
                    if must in name and (deny is None or deny not in name)]
            if hits:
                discovered[role] = os.path.join(base, sorted(hits, key=lambda x: (len(x), x))[0])
                break
    return discovered, base


def _resolve_sheet_name(names, sheet, path, label):
    if isinstance(sheet, int) and not isinstance(sheet, bool):
        if not 1 <= sheet <= len(names):
            raise ApplyError("[映射校验失败] 来源「%s」（文件 %s）的 sheet 序号 %r 越界；"
                             "实际工作表：%s" % (label, path, sheet, names))
        return names[sheet - 1]
    name = str(sheet)
    if name not in names:
        raise ApplyError("[映射校验失败] 来源「%s」（文件 %s）映射的工作表「%s」不存在；"
                         "实际工作表：%s" % (label, path, name, names))
    return name


def _resolve_source_path(root, role, entry, discovered, label):
    hint = entry.get("file") if isinstance(entry, dict) else None
    if hint:
        candidates = []
        if os.path.isabs(str(hint)):
            candidates.append(str(hint))
        else:
            candidates.append(os.path.join(root, str(hint)))
            candidates.append(os.path.join(root, os.path.basename(str(hint))))
        for candidate in candidates:
            if os.path.isfile(candidate):
                return os.path.abspath(candidate)
        wanted = os.path.basename(str(hint))
        matches = []
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames.sort()
            if wanted in filenames:
                matches.append(os.path.join(dirpath, wanted))
        if len(matches) == 1:
            return os.path.abspath(matches[0])
        if len(matches) > 1:
            raise ApplyError("[映射校验失败] 来源「%s」的文件「%s」在客户根目录下匹配到多份：%s；"
                             "请改用相对路径明确指定。" % (label, hint, matches))
        raise ApplyError("[映射校验失败] 来源「%s」的文件「%s」不存在于客户根目录 %s 下。"
                         % (label, hint, root))
    path = discovered.get(role)
    if path and os.path.isfile(path):
        return os.path.abspath(path)
    raise ApplyError("[映射校验失败] 来源「%s」未给出 file，且客户根目录下未按关键词发现对应文件；"
                     "请补齐 file 字段（相对客户根目录的路径）。" % label)


# ---------------- 来源映射校验 ----------------
def _validate_source_entry(root, role, entry, label, discovered, warnings, default_path=None):
    if not isinstance(entry, dict):
        raise ApplyError("[映射校验失败] 来源「%s」必须是对象（含 file/sheet/header_row/columns）。" % label)
    sheet = entry.get("sheet")
    header_row = entry.get("header_row")
    columns = entry.get("columns")
    if entry.get("file"):
        path = _resolve_source_path(root, role, entry, discovered, label)
    elif default_path:
        path = os.path.abspath(default_path)
    else:
        path = _resolve_source_path(root, role, entry, discovered, label)
    if sheet is None or (isinstance(sheet, str) and not sheet.strip()):
        raise ApplyError("[映射校验失败] 来源「%s」缺少 sheet（工作表名或从 1 开始的序号）。" % label)
    if isinstance(header_row, bool) or not isinstance(header_row, int) or header_row < 1:
        raise ApplyError("[映射校验失败] 来源「%s」的 header_row 必须是大于等于 1 的整数（Excel 行号），"
                         "当前 %r。" % (label, header_row))
    if not isinstance(columns, dict) or not columns:
        raise ApplyError("[映射校验失败] 来源「%s」的 columns 必须是非空对象（规范字段 → 实际列名/列号）。"
                         % label)
    fields_spec = sc.SOURCE_ROLES[role]["fields"]
    with sc.read_workbook(path) as reader:
        names = reader.sheetnames
        sheet_name = _resolve_sheet_name(names, sheet, path, label)
        matrix = reader.matrix(sheet_name)
    if header_row > len(matrix):
        raise ApplyError("[映射校验失败] 来源「%s」header_row=%d 超出工作表「%s」的行数 %d。"
                         % (label, header_row, sheet_name, len(matrix)))
    header = list(matrix[header_row - 1])
    actual_headers = [sc.column_label(header, index) for index in range(len(header))
                      if not _blank(header[index])]
    resolved = []
    for field, spec in columns.items():
        if field not in fields_spec:
            raise ApplyError("[映射校验失败] 来源「%s」含未知规范字段「%s」；%s 支持的字段：%s。"
                             % (label, field, role, list(fields_spec)))
        index = sc.find_column(header, spec)
        if index is None:
            raise ApplyError(
                "[映射校验失败] 来源「%s」（文件 %s，工作表「%s」）的字段「%s」（规范名 %s）指向列 %r，"
                "该列在第 %d 行表头中不存在。\n实际表头：%s"
                % (label, _under(root, path), sheet_name, field, fields_spec[field]["label"],
                   spec, header_row, actual_headers))
        actual = sc.column_label(header, index)
        resolved.append({"field": field, "spec": spec, "index": index, "actual": actual})
        current_score = sc.score(header[index], fields_spec[field])
        if current_score < 0.85:
            warnings.append("[需人工确认] 来源「%s」字段「%s」→ 列「%s」（相似度 %.2f，置信度%s）"
                            % (label, fields_spec[field]["label"], actual, current_score,
                               sc.confidence_of(current_score)))
    missing_required = [field for field in sc.required_fields(role) if field not in columns]
    if missing_required:
        warnings.append("[映射不完整] 来源「%s」未指定必需字段 %s；引擎对缺失字段会回退关键词启发，"
                        "建议补齐后重跑 dry-run。" % (label, missing_required))
    return {"role": role, "path": path, "display": _under(root, path),
            "sheet": sheet_name, "header_row": header_row, "columns": resolved}


def _validate_sources(root, mapping, warnings):
    sources = mapping.get("sources")
    if sources is None:
        return {}
    if not isinstance(sources, dict):
        raise ApplyError("[映射校验失败] sources 必须是对象。")
    unknown = [role for role in sources if role not in sc.VALID_SOURCE_ROLES]
    if unknown:
        raise ApplyError("[映射校验失败] sources 含不支持的来源：%s；支持：%s。"
                         % (unknown, sc.VALID_SOURCE_ROLES))
    discovered, _base = _discover_sources(root)
    plans = {}
    for role in [item for item in sc.VALID_SOURCE_ROLES if item in sources]:
        entry = sources[role]
        label = sc.SOURCE_ROLES[role]["label"]
        plan = _validate_source_entry(root, role, entry, label, discovered, warnings)
        plans[role] = plan
        nested = entry.get("next_month") if isinstance(entry, dict) else None
        if nested is not None:
            if not isinstance(nested, dict):
                raise ApplyError("[映射校验失败] 来源「%s」的 next_month 必须是对象。" % label)
            # 引擎固定从库存导出工作簿读取「次月已销」；next_month 挂在 inventory 或 jia 下都可，
            # 校验也按同一口径（挂在 jia 下时仍核对库存文件的对应工作表）。
            nested_path = plan["path"] if role == "inventory" else discovered.get("inventory")
            nested_plan = _validate_source_entry(
                root, "next_month", nested, label + "·次月已销", discovered, warnings,
                default_path=nested_path or plan["path"])
            plans[role]["next_month"] = nested_plan
            if nested.get("file"):
                warnings.append("[提示] next_month.file 仅作记录：当前引擎固定从库存导出工作簿"
                                "读取「次月已销」工作表。")
            if role != "inventory":
                warnings.append("[提示] next_month 挂在 sources.%s 下；引擎从库存导出工作簿读取"
                                "次月已销（映射仍按该声明生效）。" % role)
    if "next_month" in sources and isinstance(sources["next_month"], dict) \
            and sources["next_month"].get("file"):
        warnings.append("[提示] sources.next_month.file 仅作记录：当前引擎固定从库存导出工作簿"
                        "读取「次月已销」工作表。")
    return plans


# ---------------- 主档映射声明校验 ----------------
def _validate_master_declaration(root, mapping, warnings):
    master = mapping.get("master")
    order = mapping.get("order")
    sheets = {}
    if master is not None:
        if not isinstance(master, dict):
            raise ApplyError("[映射校验失败] master 必须是对象（含 sheets）。")
        sheets = master.get("sheets") or {}
        if not isinstance(sheets, dict):
            raise ApplyError("[映射校验失败] master.sheets 必须是对象（实际表名 → 规范表名）。")
    if order is None:
        order = []
    if not isinstance(order, list) or any(not isinstance(name, str) for name in order):
        raise ApplyError("[映射校验失败] order 必须是规范表名数组，如 %s。" % sc.MASTER_SHEET_ORDER)
    if not sheets and not order:
        return None
    path = _discover_master(root)
    if path is None:
        raise ApplyError("[映射校验失败] 映射声明了 master/order，但客户根目录下未找到主档"
                         "（文件名需含「主档」，如 _主档.xlsx）。")
    with sc.read_workbook(path) as reader:
        names = reader.sheetnames
    seen_canonical = {}
    for actual, canonical in sheets.items():
        if canonical not in sc.MASTER_HEADERS:
            raise ApplyError("[映射校验失败] master.sheets 中「%s」的目标「%s」不是规范表；规范表：%s。"
                             % (actual, canonical, sc.MASTER_SHEET_ORDER))
        if actual not in names:
            raise ApplyError("[映射校验失败] master.sheets 指定的实际工作表「%s」在主档 %s 中不存在；"
                             "实际工作表：%s。" % (actual, path, names))
        seen_canonical.setdefault(canonical, []).append(actual)
    for canonical, actuals in seen_canonical.items():
        if len(actuals) > 1:
            raise ApplyError("[映射校验失败] master.sheets 把多张工作表映射到同一规范表「%s」：%s。"
                             % (canonical, actuals))
        if actuals[0] != canonical and canonical in names and canonical not in sheets:
            raise ApplyError("[映射校验失败] 主档已有规范表「%s」，又要把「%s」也映射为「%s」，会重名；"
                             "请人工确认后再改。" % (canonical, actuals[0], canonical))
    for index, name in enumerate(order):
        if name not in sc.MASTER_HEADERS:
            raise ApplyError("[映射校验失败] order 含非规范表名「%s」；规范表：%s。"
                             % (name, sc.MASTER_SHEET_ORDER))
        if order.count(name) > 1:
            raise ApplyError("[映射校验失败] order 含重复表名「%s」。" % name)
    return {"path": path, "sheets": sheets, "order": order, "actual_names": names}


# ---------------- 主档修复计划 ----------------
def _best_master_header(matrix, wanted_headers):
    best_index, best_score = 0, 0.0
    for index in range(min(len(matrix), 3)):
        values = [value for value in matrix[index] if not _blank(value)]
        if not values:
            continue
        total = 0.0
        for wanted in wanted_headers:
            spec = {"label": wanted, "synonyms": [wanted]}
            total += max((sc.score(value, spec) for value in values), default=0.0)
        average = total / len(wanted_headers)
        if average > best_score:
            best_index, best_score = index, average
    return best_index, best_score


def _build_master_plan(root, mapping, fix_order):
    master_path = _discover_master(root)
    if master_path is None:
        raise ApplyError("[主档修复失败] 客户根目录下未找到主档（文件名需含「主档」，如 _主档.xlsx）。")
    master_cfg = mapping.get("master") or {}
    sheets_map = master_cfg.get("sheets") or {}
    order = mapping.get("order") or list(sc.MASTER_SHEET_ORDER)
    canonical2actual = {}
    for actual, canonical in sheets_map.items():
        canonical2actual[canonical] = actual

    try:
        workbook = openpyxl.load_workbook(master_path)
    except Exception as exc:  # noqa: BLE001 - 统一转中文错误
        raise ApplyError("[主档修复失败] 无法读取主档 %s\n原因：%s\n"
                         "请确认主档为 .xlsx 且未被 Excel 占用。" % (master_path, exc))
    try:
        names = list(workbook.sheetnames)
        resolved = {}
        for canonical in sc.MASTER_SHEET_ORDER:
            actual = canonical2actual.get(canonical, canonical)
            if actual not in names:
                if canonical in names:
                    actual = canonical
                else:
                    raise ApplyError("[主档修复失败] 主档缺少规范表「%s」（映射 master.sheets 指向「%s」，"
                                     "该工作表不存在）；实际工作表：%s。"
                                     % (canonical, actual, names))
            resolved[canonical] = actual
        if len(set(resolved.values())) != len(resolved):
            raise ApplyError("[主档修复失败] 主档中有多张工作表被映射到同一规范表：%s。" % resolved)
        final_names = {actual: canonical for canonical, actual in resolved.items()}
        for canonical in sc.MASTER_SHEET_ORDER:
            actual = resolved[canonical]
            if actual != canonical and canonical in names and canonical not in final_names:
                raise ApplyError("[主档修复失败] 主档已有「%s」表，无法把「%s」改名为「%s」（重名）。"
                                 % (canonical, actual, canonical))

        plan_sheets = []
        for canonical in sc.MASTER_SHEET_ORDER:
            actual = resolved[canonical]
            worksheet = workbook[actual]
            matrix = [[cell.value for cell in row] for row in worksheet.iter_rows()]
            wanted = list(sc.MASTER_HEADERS[canonical])
            header_index, header_score = _best_master_header(matrix, wanted)
            if header_index != 0:
                raise ApplyError("[主档修复失败] 主档「%s」表的表头不在第 1 行（识别为第 %d 行，"
                                 "匹配度 %.2f）；load_master 按第 1 行读取，请人工把表头放回第 1 行。"
                                 % (actual, header_index + 1, header_score))
            header = list(matrix[0]) if matrix else []
            matched = sc.match_columns(header, wanted, min_score=MASTER_MIN_SCORE)
            missing = [name for name in wanted if name not in matched]
            if missing:
                details = []
                for name in missing:
                    best_value, best_value_score = None, 0.0
                    for value in header:
                        if _blank(value):
                            continue
                        current = sc.score(value, {"label": name, "synonyms": [name]})
                        if current > best_value_score:
                            best_value, best_value_score = value, current
                    details.append("「%s」（最佳候选 %r，相似度 %.2f）"
                                   % (name, best_value, best_value_score))
                raise ApplyError("[主档修复失败] 主档「%s」表无法确定以下规范列对应哪一列：%s；"
                                 "请人工核对后再修复。" % (actual, "；".join(details)))
            renames, used = [], set()
            for name in wanted:
                index = matched[name][0]
                used.add(index)
                if sc.normalize(header[index]) != sc.normalize(name):
                    renames.append({"index": index,
                                    "old": sc.column_label(header, index), "new": name})
            extras = [index for index in range(len(header))
                      if index not in used and not _blank(header[index])]
            column_order, order_changed = None, False
            old_headers = new_headers = None
            if fix_order:
                column_order = [matched[name][0] for name in wanted] + extras
                column_order += [index for index in range(len(header)) if index not in column_order]
                order_changed = column_order != list(range(len(header)))
                old_headers = [sc.column_label(header, index) for index in range(len(header))]
                new_headers = [sc.column_label(header, index) for index in column_order]
                if order_changed and list(worksheet.merged_cells.ranges):
                    raise ApplyError("[主档修复失败] 主档「%s」表含合并单元格，按列重排不安全，"
                                     "已中止；可只用 --fix-master 修复表名/列名。" % actual)
                if order_changed and list(worksheet.conditional_formatting):
                    raise ApplyError("[主档修复失败] 主档「%s」表含条件格式，按列重排不安全，"
                                     "已中止；可只用 --fix-master 修复表名/列名。" % actual)
            plan_sheets.append({
                "actual": actual, "canonical": canonical,
                "sheet_rename": actual != canonical,
                "renames": renames, "extras": extras,
                "column_order": column_order, "order_changed": order_changed,
                "old_headers": old_headers, "new_headers": new_headers,
            })

        desired_order = []
        wanted_order = order or list(sc.MASTER_SHEET_ORDER)
        for canonical in wanted_order:
            if canonical in resolved and canonical not in desired_order:
                desired_order.append(canonical)
        for canonical in sc.MASTER_SHEET_ORDER:
            if canonical not in desired_order:
                desired_order.append(canonical)
        extra_sheets = [name for name in names if name not in final_names]
        desired_order += extra_sheets
        current_final_order = [final_names.get(name, name) for name in names]
        sheet_order_changed = current_final_order != desired_order
    finally:
        workbook.close()

    changed = any(sheet["sheet_rename"] or sheet["renames"] for sheet in plan_sheets)
    if fix_order:
        changed = (changed or sheet_order_changed
                   or any(sheet["order_changed"] for sheet in plan_sheets))
    return {"root": root, "master_path": master_path, "sha256": sc.sha256_file(master_path),
            "plan_sheets": plan_sheets, "desired_order": desired_order,
            "current_order": current_final_order, "sheet_order_changed": sheet_order_changed,
            "fix_order": fix_order, "changed": changed}


def _print_master_plan(plan, write, confirmed_by):
    print("[主档修复%s] %s" % ("·执行" if write else "·dry-run", plan["master_path"]))
    print("  主档 sha256：%s" % plan["sha256"])
    print("  备份：%s（毫秒时间戳，先备份后原子替换）"
          % os.path.join(plan["root"], sc.BACKUP_DIR,
                         "%s_YYYYmmdd_HHMMSS_mmm%s"
                         % (os.path.splitext(os.path.basename(plan["master_path"]))[0],
                            os.path.splitext(plan["master_path"])[1])))
    if not plan["changed"]:
        print("  主档结构与规范一致，无需修改。")
        return
    for sheet in plan["plan_sheets"]:
        if sheet["sheet_rename"]:
            print("  表改名：「%s」→「%s」" % (sheet["actual"], sheet["canonical"]))
        for rename in sheet["renames"]:
            print("  列改名「%s」：第 %d 列「%s」→「%s」"
                  % (sheet["canonical"], rename["index"] + 1, rename["old"], rename["new"]))
        if sheet["order_changed"]:
            print("  列重排「%s」：%s → %s"
                  % (sheet["canonical"], "、".join(str(item) for item in sheet["old_headers"]),
                     "、".join(str(item) for item in sheet["new_headers"])))
    if plan["sheet_order_changed"]:
        print("  工作表顺序：%s → %s"
              % (plan["current_order"], plan["desired_order"]))
    print("  影响：%s"
          % ("将备份到 _备份/ 后原子替换主档；三份源导出只读不改。"
             if write else "dry-run 未写任何文件；三份源导出不会被修改。"))
    if not write:
        print("  下一步：确认无误后执行 --fix-master%s --write --confirmed-by <确认人>。"
              % (" --fix-order" if plan["fix_order"] else ""))


def _remap_validations(worksheet, column_order):
    """数据有效性区域随列重排一起平移；解析失败时保留原样（不阻断修复）。"""
    inverse = {source: target for target, source in enumerate(column_order)}
    for validation in list(worksheet.data_validations.dataValidation):
        parts, ok = [], True
        for cell_range in str(validation.sqref).split():
            match = re.fullmatch(r"([A-Z]{1,3})([0-9]+)(?::([A-Z]{1,3})([0-9]+))?", cell_range)
            if not match:
                ok = False
                break
            first = column_index_from_string(match.group(1)) - 1
            last = column_index_from_string(match.group(3)) - 1 if match.group(3) else first
            if first not in inverse or last not in inverse:
                ok = False
                break
            target_first, target_last = inverse[first], inverse[last]
            if target_first > target_last:
                target_first, target_last = target_last, target_first
            part = "%s%s" % (get_column_letter(target_first + 1), match.group(2))
            if match.group(3):
                part += ":%s%s" % (get_column_letter(target_last + 1), match.group(4))
            parts.append(part)
        if ok and parts:
            validation.sqref = " ".join(parts)


def _reorder_columns(worksheet, column_order):
    """按 column_order（目标位置 ← 源列 0 基序号）重排版面：值 + 单元格样式 + 列宽一起搬。"""
    max_row, max_column = worksheet.max_row, worksheet.max_column
    if max_column == 0:
        return
    order = list(column_order) + [index for index in range(max_column)
                                  if index not in column_order]
    snapshot = []
    for row_index in range(1, max_row + 1):
        row = []
        for column in range(1, max_column + 1):
            cell = worksheet.cell(row=row_index, column=column)
            row.append((cell.value, copy.copy(cell._style) if cell.has_style else None))
        snapshot.append(row)
    widths = [worksheet.column_dimensions[get_column_letter(column)].width
              for column in range(1, max_column + 1)]
    for row_index in range(1, max_row + 1):
        for target, source in enumerate(order, start=1):
            cell = worksheet.cell(row=row_index, column=target)
            value, style = snapshot[row_index - 1][source]
            cell.value = value
            if style is not None:
                cell._style = style
    for target, source in enumerate(order, start=1):
        worksheet.column_dimensions[get_column_letter(target)].width = widths[source]
    _remap_validations(worksheet, order)
    if worksheet.auto_filter is not None and worksheet.auto_filter.ref:
        worksheet.auto_filter.ref = "A1:%s%d" % (get_column_letter(max_column), max_row)


def _verify_master_after_fix(plan):
    workbook = openpyxl.load_workbook(plan["master_path"], read_only=True)
    try:
        names = list(workbook.sheetnames)
        for canonical in sc.MASTER_SHEET_ORDER:
            if canonical not in names:
                raise ApplyError("[主档修复失败] 修复后校验未通过：缺少规范表「%s」（当前：%s）。"
                                 "备份文件在 _备份/，请人工恢复。" % (canonical, names))
            worksheet = workbook[canonical]
            header = [cell.value for cell in next(worksheet.iter_rows(min_row=1, max_row=1))]
            wanted = list(sc.MASTER_HEADERS[canonical])
            actual = [sc.column_label(header, index) for index in range(len(header))]
            if actual[:len(wanted)] != wanted:
                raise ApplyError("[主档修复失败] 修复后校验未通过：「%s」表头前 %d 列为 %s，期望 %s。"
                                 "备份文件在 _备份/，请人工恢复。"
                                 % (canonical, len(wanted), actual[:len(wanted)], wanted))
        if plan["fix_order"] and names != plan["desired_order"]:
            raise ApplyError("[主档修复失败] 修复后校验未通过：工作表顺序 %s，期望 %s。"
                             "备份文件在 _备份/，请人工恢复。" % (names, plan["desired_order"]))
    finally:
        workbook.close()


def _execute_master_plan(plan):
    master_path = plan["master_path"]
    root = plan["root"]
    backup_dir = os.path.join(root, sc.BACKUP_DIR)
    os.makedirs(backup_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(master_path))
    backup_path = _unique_path(backup_dir, stem, ext)
    shutil.copy2(master_path, backup_path)
    print("  已备份：%s" % _under(root, backup_path))

    original_stat = os.stat(master_path)
    try:
        workbook = openpyxl.load_workbook(master_path)
    except Exception as exc:  # noqa: BLE001 - 统一转中文错误
        raise ApplyError("[主档修复失败] 无法读取主档 %s\n原因：%s\n备份已保留在 _备份/。"
                         % (master_path, exc))
    try:
        for sheet in plan["plan_sheets"]:
            worksheet = workbook[sheet["actual"]]
            for rename in sheet["renames"]:
                worksheet.cell(row=1, column=rename["index"] + 1, value=rename["new"])
            if plan["fix_order"] and sheet["order_changed"]:
                _reorder_columns(worksheet, sheet["column_order"])
        for sheet in plan["plan_sheets"]:
            if sheet["sheet_rename"]:
                workbook[sheet["actual"]].title = sheet["canonical"]
        if plan["fix_order"] and plan["sheet_order_changed"]:
            ordered = [workbook[name] for name in plan["desired_order"] if name in workbook.sheetnames]
            ordered += [worksheet for worksheet in workbook.worksheets
                        if worksheet.title not in plan["desired_order"]]
            workbook._sheets = ordered
        temp_path = os.path.join(os.path.dirname(master_path),
                                 ".%s.tmp-%d" % (os.path.basename(master_path), os.getpid()))
        try:
            workbook.save(temp_path)
        except Exception:
            if os.path.exists(temp_path):
                os.remove(temp_path)
            raise
    finally:
        workbook.close()
    try:
        os.chmod(temp_path, stat.S_IMODE(original_stat.st_mode))
    except OSError:
        pass
    os.replace(temp_path, master_path)
    _verify_master_after_fix(plan)
    print("  已修复：%s（原子替换完成；修复前副本在 _备份/）" % master_path)


# ---------------- 映射落盘 ----------------
def _write_mapping(root, mapping, confirmed_by):
    final = dict(mapping)
    final["schema"] = sc.MAPPING_SCHEMA
    final["confirmed_by"] = confirmed_by
    final["confirmed_at"] = sc.iso_now()
    final["notes"] = list(mapping.get("notes") or [])
    text = json.dumps(final, ensure_ascii=False, indent=2) + "\n"
    target = sc.mapping_path(root)
    temp_path = target + ".tmp-%d" % os.getpid()
    with open(temp_path, "w", encoding="utf-8") as fh:
        fh.write(text)
    os.replace(temp_path, target)
    print("  已写入：%s" % target)
    confirm_dir = os.path.join(root, sc.CONFIRM_DIR)
    os.makedirs(confirm_dir, exist_ok=True)
    archive_path = _unique_path(confirm_dir, "映射", ".json")
    with open(archive_path, "w", encoding="utf-8") as fh:
        fh.write(text)
    print("  已归档：%s" % archive_path)
    print("  确认信息：confirmed_by=%s，confirmed_at=%s" % (final["confirmed_by"],
                                                          final["confirmed_at"]))
    return target, archive_path


# ---------------- 计划打印 ----------------
def _print_source_plans(plans, master_info, root, dry_run, warnings, notes=(), fix_mode=False):
    print("[映射校验] 客户根目录：%s" % root)
    print("  模式：%s" % ("dry-run（不写任何文件）" if dry_run else "写入"))
    for note in notes:
        print("  映射备注：%s" % note)
    for role in [item for item in sc.VALID_SOURCE_ROLES if item in plans]:
        plan = plans[role]
        print("  来源 %-10s %s" % (role, _under(root, plan["path"])))
        print("      工作表「%s」，表头第 %d 行" % (plan["sheet"], plan["header_row"]))
        for column in plan["columns"]:
            print("      字段 %-7s ← 列「%s」"
                  % (column["field"], column["actual"]))
        next_month = plan.get("next_month")
        if next_month:
            print("      次月已销：工作表「%s」，表头第 %d 行"
                  % (next_month["sheet"], next_month["header_row"]))
            for column in next_month["columns"]:
                print("        字段 %-7s ← 列「%s」" % (column["field"], column["actual"]))
    if master_info:
        print("  主档：%s" % _under(root, master_info["path"]))
        if master_info["sheets"]:
            for actual, canonical in master_info["sheets"].items():
                print("      表名映射：「%s」→「%s」" % (actual, canonical))
        if master_info["order"]:
            print("      规范表顺序：%s" % "、".join(master_info["order"]))
    for warning in warnings:
        print("  " + warning)
    if fix_mode:
        print("  影响：引擎将按该映射确定性取数；三份源导出只读不改；"
              "主档仅在 --write 时按下方修复计划改动。")
    else:
        print("  影响：引擎将按该映射确定性取数；三份源导出只读不改，主档本步骤不动。")


def _print_mapping_write_plan(root, write, confirmed_by):
    target = sc.mapping_path(root)
    print("  将写入：%s（schema=%s，confirmed_by=%s，confirmed_at=写入时生成）"
          % (target, sc.MAPPING_SCHEMA, confirmed_by or "待填"))
    print("  归档副本：%s/映射_YYYYmmdd_HHMMSS_mmm.json" % os.path.join(root, sc.CONFIRM_DIR))
    if not write:
        print("  下一步：用户逐项确认后执行 --write --confirmed-by <确认人> 落盘。")


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="映射校验与确认落盘 / 主档结构修复（默认 dry-run，绝不修改三份源导出）")
    parser.add_argument("root", help="客户根目录（主档与导出所在目录）")
    parser.add_argument("mapping", help="映射 JSON（inspect_schema.py 的 proposed_mapping）")
    parser.add_argument("--write", action="store_true", help="执行写入（需要 --confirmed-by）")
    parser.add_argument("--dry-run", action="store_true", help="只校验并打印计划（默认行为）")
    parser.add_argument("--confirmed-by", default="", help="确认人姓名/标识（--write 必填）")
    parser.add_argument("--fix-master", action="store_true",
                        help="修复主档：表名/列名 → 规范名（先 dry-run，--write 需 --confirmed-by）")
    parser.add_argument("--fix-order", action="store_true",
                        help="修复主档时按规范顺序重排列与工作表（隐含 --fix-master）")
    args = parser.parse_args(argv)

    if args.write and args.dry_run:
        sys.stderr.write("[参数错误] --write 与 --dry-run 只能二选一。\n")
        return 1
    write = bool(args.write)
    dry_run = not write
    fix_mode = bool(args.fix_master or args.fix_order)
    fix_order = bool(args.fix_order)
    confirmed_by = (args.confirmed_by or "").strip()

    root = os.path.abspath(args.root)
    if not os.path.isdir(root):
        sys.stderr.write("[目录不存在] %s\n请传入客户根目录（含 _主档.xlsx 与月份子目录）。\n" % root)
        return 1

    try:
        if write and not confirmed_by:
            raise ApplyError("[必须经用户确认] --write 需要 --confirmed-by <确认人>：\n"
                             "  请先把 dry-run 的校验计划与影响逐项给用户确认，得到明确同意后再执行；\n"
                             "  未写入任何文件。")
        try:
            mapping = sc.load_json(args.mapping)
        except FileNotFoundError:
            raise ApplyError("[找不到映射文件] %s\n请先用 inspect_schema.py --out 生成探查结果，"
                             "取出 proposed_mapping 保存为映射 JSON。" % args.mapping)
        except (ValueError, OSError) as exc:
            raise ApplyError("[映射文件无法解析] %s\n原因：%s" % (args.mapping, exc))
        if not isinstance(mapping, dict):
            raise ApplyError("[映射格式错误] %s 顶层必须是 JSON 对象。" % args.mapping)
        schema = mapping.get("schema")
        if schema is not None and schema != sc.MAPPING_SCHEMA:
            raise ApplyError("[映射 schema 不支持] 当前 %r；要求 %r。"
                             % (schema, sc.MAPPING_SCHEMA))

        warnings = []
        plans = _validate_sources(root, mapping, warnings)
        master_info = _validate_master_declaration(root, mapping, warnings)

        notes = list(mapping.get("notes") or [])
        if fix_mode:
            plan = _build_master_plan(root, mapping, fix_order)
            _print_source_plans(plans, master_info, root, dry_run, warnings, notes,
                                fix_mode=True)
            print("")
            _print_master_plan(plan, write, confirmed_by)
            if not write:
                print("\n[dry-run] 未写任何文件。")
                return 0
            if not plan["changed"]:
                print("\n[无需修复] 主档结构与规范一致，未写任何文件。")
                return 0
            _execute_master_plan(plan)
            print("\n[已修复] 主档结构已按规范修复；三份源导出未修改。")
            return 0

        _print_source_plans(plans, master_info, root, dry_run, warnings, notes)
        _print_mapping_write_plan(root, write, confirmed_by)
        if not write:
            print("\n[dry-run] 未写任何文件。")
            return 0
        _write_mapping(root, mapping, confirmed_by)
        print("\n[已落盘] 引擎下次运行 load_master 时会自动读取 %s；三份源导出未修改。"
              % sc.MAPPING_FILENAME)
        return 0
    except ApplyError as exc:
        sys.stderr.write(str(exc) + "\n")
        return 1
    except (sc.ExcelError, OSError) as exc:
        sys.stderr.write("[执行失败] %s\n" % exc)
        return 1
    except Exception as exc:  # noqa: BLE001 - CLI 顶层兜底，避免把 traceback 直接抛给用户
        import traceback
        sys.stderr.write("[执行失败] 未预期的错误：%s\n" % exc)
        traceback.print_exc(file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
