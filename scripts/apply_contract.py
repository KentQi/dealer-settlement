"""把已确认的合同字段写入 _主档.xlsx。

默认 dry-run：只打印「现值 → 新值」差异，不落盘；确认无误后加 --write 才写入。

用法：
    python3 apply_contract.py _主档.xlsx 合同候选.json           # 预览
    python3 apply_contract.py _主档.xlsx 合同候选.json --write   # 落盘

规则：
- 只写非空值；候选为空一律跳过，绝不把已有内容清空。
- 已有值且与新值不同 → 记为「冲突」，默认不覆盖；需显式 --force。
- 幂等：同一份 JSON 重复执行不产生变化。
"""
import sys, json, argparse, shutil, os, re, base64, hashlib, re
import openpyxl

ROLE_COL = "角色"


# ---- 列类型规则：写入前强校验，脏值拒绝而非转换 ----
_NUM_COLS = {"供货价", "零售价", "贴花标准", "期初数量", "期初金额", "结算数量", "结算金额",
             "差异数量", "差异金额", "异议期(工作日)", "签回期(工作日)"}
_DATE_COLS = {"起始", "终止", "生效起", "生效止", "数据基准日"}
_ENUM_COLS = {"结算属性": ("结算类", "非结算类"),
              "状态": ("待确认", "已确认", "已对账", "已付款"),
              "结算方式": ("铺货代销", "买断", "代销"),
              "结算基数": ("实际动销", "签收", "发货")}
_BAR_COLS = ("商品条码", "条码")
_ENUM_KEYS = tuple(_ENUM_COLS.keys())


def _match_rule(col, table):
    """按「列名包含关键词」匹配类型规则，覆盖「贴花标准(元/件)」等带注记的真实表头。"""
    for k in _NUM_COLS:
        if k in col:
            return "num"
    for k in _DATE_COLS:
        if k in col:
            return "date"
    if table in _ENUM_COLS:
        return ("enum", table)
    if any(k in col for k in _ENUM_KEYS):
        return ("enum", next(k for k in _ENUM_KEYS if k in col))
    if any(k in col for k in _BAR_COLS):
        return "bar"
    return None


def validate_value(col, val):
    """返回 (ok, 规范值, 拒绝原因)。拒绝时不写入、不静默转换。校验器自身异常也按拒绝处理。"""
    s = str(val).strip()
    try:
        rule = _match_rule(col, s if False else col)
        if rule == "num":
            if not re.fullmatch(r"[+-]?[0-9]+(\.[0-9]+)?", s):
                return False, None, f"「{col}」须为半角数字，收到「{s}」（全角/文本一律拒绝）"
            f = float(s)
            import math
            if not math.isfinite(f):
                return False, None, f"「{col}」须为有限数字，收到「{s}」"
            if f < 0:
                return False, None, f"「{col}」不应为负数（{f}）。若属调整/冲减，请与甲方核对后另行登记"
            if abs(f) > 1e9:
                return False, None, f"「{col}」数值异常偏大（{f}），请核对单位"
            f = round(f, 2)
            return True, (int(f) if f == int(f) else f), None
        if rule == "date":
            import datetime as _dt
            if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", s):
                return False, None, f"「{col}」日期须为 YYYY-MM-DD（如 2026-11-02），收到「{s}」"
            try:
                _dt.date.fromisoformat(s)
            except ValueError:
                return False, None, f"「{col}」不是有效日期：{s}"
            return True, s, None
        if isinstance(rule, tuple) and rule[0] == "enum":
            allowed = _ENUM_COLS[rule[1]]
            if s not in allowed:
                return False, None, f"「{col}」仅允许：{' / '.join(allowed)}，收到「{s}」"
            return True, s, None
        if rule == "bar":
            if not (re.fullmatch(r"[0-9]{8,14}", s)):
                return False, None, f"「{col}」须为 8–14 位数字条码，收到「{s}」"
            return True, s, None
        return True, s, None
    except SystemExit:
        raise
    except Exception as e:
        return False, None, f"校验「{col}」时出错（{e}），已拒绝——请人工核对该值"


def _locate_row(ws, headers, keys):
    """按 {列名: 值} 多键定位数据行；列名支持包含匹配（"条码"匹配"商品条码"）。"""
    cols = []
    for k, v in keys.items():
        hit = next((i for i, h in enumerate(headers) if h and k in str(h)), None)
        if hit is None:
            return None
        cols.append((hit + 1, str(v).strip()))
    for r in range(2, ws.max_row + 1):
        if all(str(ws.cell(row=r, column=c).value or "").strip() == v for c, v in cols):
            return r
    return None


def sheet_row(ws, match_col, match_val):
    headers = [c.value for c in ws[1]]
    if match_col not in headers:
        return None, headers
    idx = headers.index(match_col) + 1
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=idx).value).strip() == match_val:
            return r, headers
    return None, headers


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("master")
    ap.add_argument("candidate")
    ap.add_argument("--write", action="store_true")
    ap.add_argument("--force", action="store_true")
    a = ap.parse_args()

    try:
        with open(a.candidate, encoding="utf-8") as fh:
            data = json.load(fh)
    except FileNotFoundError:
        sys.exit(f"[找不到文件] {a.candidate}\n请确认确认单已下载到本机，并传入正确路径。")
    except json.JSONDecodeError as e:
        sys.exit(f"[JSON 解析失败] {a.candidate}\n第 {e.lineno} 行第 {e.colno} 列：{e.msg}\n"
                 f"多为文件被截断或手工编辑出错，请重新从确认页下载。")
    if isinstance(data, list):
        data = {"fields": data}
    confirmed = data.get("schema") == "contract-confirm/v1"
    meta = data.get("meta", {})
    if confirmed and not str(meta.get("operator") or "").strip():
        sys.exit("[中止] 确认单缺少确认人（meta.operator 为空）。留档必须有责任人，"
                 "请在确认页填写姓名后重新提交。")
    fields = data.get("fields", [])
    if not fields:
        print("[提醒] 该文件不含任何字段（fields 为空）。")
        print("       若这是从确认页下载的确认单，说明所有项都被取消了「采用」；")
        print("       若这是原始候选 JSON，请检查提取环节是否失败。")
    wb = openpyxl.load_workbook(a.master)

    changes, conflicts, skipped, rejected = [], [], [], []
    if confirmed:
        print(f"[已确认] 确认人：{meta.get('operator','—')}　时间：{meta.get('confirmed_at','—')}"
              f"　备注：{meta.get('memo') or '—'}\n")
    for f in fields:
        tgt, val = f.get("target", ""), f.get("value")
        if not tgt or "." not in tgt:
            skipped.append((f.get("field"), "无 target"))
            continue
        if val is None or str(val).strip() == "":
            skipped.append((f.get("field"), "候选为空"))
            continue
        sheet, rest = tgt.split(".", 1)
        if sheet not in wb.sheetnames:
            skipped.append((f.get("field"), f"无工作表 {sheet}"))
            continue
        ws = wb[sheet]

        if sheet == "主体":
            role, col = rest.split(".", 1)
            row, headers = sheet_row(ws, ROLE_COL, role)
        else:
            headers = [c.value for c in ws[1]]
            row = 2
            # 合同表按合同编号定位（续签多行）；价格/贴花按 条码+生效起 定位（SCD2 多行）；SKU 按条码定位。
            # 字段自身需带主键说明，格式：列名@条码=xxx 或 列名@条码=xxx@生效起=yyyy-mm-dd
            if sheet == "合同" and "合同编号" in headers:
                ci = headers.index("合同编号") + 1
                want = str(val).strip() if rest == "合同编号" else None
                if want:
                    for r2 in range(2, ws.max_row + 1):
                        if str(ws.cell(row=r2, column=ci).value).strip() == want:
                            row = r2; break
                if row == 2 and rest != "合同编号":
                    for r2 in range(2, ws.max_row + 1):
                        if str(ws.cell(row=r2, column=ci).value or "").strip():
                            row = r2; break
            elif sheet in ("价格", "贴花", "SKU"):
                keys = {}          # 从 target 解析 @主键=值
                base = rest
                while "@" in base:
                    base, kv = base.rsplit("@", 1)
                    if "=" in kv:
                        k2, v2 = kv.split("=", 1)
                        keys[k2.strip()] = v2.strip()
                col = base
                if not keys:
                    rejected.append((f.get("field"), tgt, val,
                                     f"{sheet} 表为多行结构（SCD2/SKU），target 必须带主键，"
                                     f"如：{base}@商品条码=xxx" + ("@生效起=yyyy-mm-dd" if sheet != "SKU" else "")))
                    continue
                _hit = _locate_row(ws, headers, keys)
                if _hit is None:
                    rejected.append((f.get("field"), tgt, val,
                                     f"未找到满足主键 {keys} 的行，请人工定位后录入（不默认写第 2 行）"))
                    continue
                row = _hit
                role, col = "", base     # 主键分支：列名 = 剥离 @主键 后的 base
            if sheet not in ("价格", "贴花", "SKU"):
                role, col = "", rest     # 其余表（含合同）仍用完整 rest 作列名
        if row is None or col not in headers:
            skipped.append((f.get("field"), f"无定位 {tgt}"))
            continue

        cell = ws.cell(row=row, column=headers.index(col) + 1)

        # ---- 类型校验：脏值进不了主档（拒绝，不静默转换）----
        ok, norm_val, why = validate_value(col, val)
        if not ok:
            rejected.append((f.get("field"), tgt, val, why))
            continue
        val = norm_val

        cur = cell.value
        if cur is None or str(cur).strip() == "":
            changes.append((tgt, cur, val, cell))
        elif str(cur).strip() == str(val).strip():
            skipped.append((f.get("field"), "无变化"))
        elif confirmed:
            changes.append((tgt, cur, val, cell))   # 已人工确认，视为批准变更
        else:
            conflicts.append((tgt, cur, val, cell))

    print(f"{'类型':6s}{'字段':26s}{'现值':34s}{'新值'}")
    print("-" * 100)
    for t, c, v, _ in changes:
        print(f"{'新增':6s}{t:26s}{'—':34s}{v}")
    for t, c, v, _ in conflicts:
        print(f"{'冲突':6s}{t:26s}{str(c)[:32]:34s}{v}")
    for n, why in skipped:
        print(f"{'跳过':6s}{n:26s}{why}")
    for n, t, v, why in rejected:
        print(f"{'拒绝':6s}{n:26s}{why}　（候选值：{str(v)[:40]}）")
    print(f"\n合计：可写入 {len(changes)} 项，冲突 {len(conflicts)} 项，跳过 {len(skipped)} 项，拒绝 {len(rejected)} 项")
    if rejected:
        print("[被拒字段未写入] 请修正后重新生成确认单，或修正主档后手工录入。")
    if confirmed:
        print("（本文件为已确认单，冲突项按确认值写入）")

    if not a.write:
        print("\n[dry-run] 未写入。确认无误后加 --write 执行。")
        return
    if not changes and not (conflicts and a.force):
        if confirmed and (data.get("attachments") or []):
            import datetime as _dt0
            _root0 = os.path.dirname(os.path.abspath(a.master))
            _adir0 = os.path.join(_root0, "_确认记录")
            os.makedirs(_adir0, exist_ok=True)
            _stamp0 = _dt0.datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
            _arc0 = _adir0 + f"/{_stamp0}_字段无变化_附件归档.json"
            shutil.copy2(a.candidate, _arc0)
            print(f"[无变化] 主档已是目标状态，未写盘、未备份；但确认单含附件，已归档：{_arc0}")
        else:
            print("\n[无变化] 主档已是目标状态，未写盘、未备份。")
        return
    if rejected and changes:
        print(f"\n[提示] {len(rejected)} 项被类型校验拒绝，本次仅写入其余 {len(changes)} 项。")
    if conflicts and not a.force:
        print("\n存在冲突且未加 --force，已中止写入。请先确认冲突项。")
        sys.exit(1)

    if confirmed:
        import datetime as _dt
        root = os.path.dirname(os.path.abspath(a.master))
        adir = os.path.join(root, "_确认记录")
        os.makedirs(adir, exist_ok=True)
        stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
        arc = os.path.join(adir, f"{stamp}_合同确认.json")
        _n = 2
        while os.path.exists(arc):
            arc = os.path.join(adir, f"{stamp}_{_n}_合同确认.json"); _n += 1
        stamp = os.path.basename(arc).replace("_合同确认.json", "")

        arc_data = json.loads(json.dumps(data))          # 深拷贝，改归档版不影响入参
        atts = arc_data.get("attachments") or []
        if atts:
            sub = os.path.join(adir, f"{stamp}_附件")
            _n2 = 2
            while os.path.exists(sub):
                sub = os.path.join(adir, f"{stamp}_{_n2}_附件"); _n2 += 1
            os.makedirs(sub, exist_ok=True)
            saved = []
            for i, at in enumerate(atts, 1):
                b64 = at.pop("data_b64", "") or ""
                name = re.sub(r'[\\/:*?"<>|]', "_", str(at.get("name") or f"附件{i}"))
                try:
                    raw = base64.b64decode(b64, validate=True) if b64 else b""
                except Exception:
                    print(f"⚠ 附件「{name}」base64 损坏，已跳过"); continue
                if len(raw) > 25 * 1024 * 1024:
                    print(f"⚠ 附件「{name}」超过 25MB，已跳过"); continue
                fp = os.path.join(sub, f"{i:02d}_{name}")
                open(fp, "wb").write(raw)
                at["saved_as"] = os.path.relpath(fp, root)
                at["sha256"] = hashlib.sha256(raw).hexdigest()
                at["bytes"] = len(raw)
                saved.append((at["saved_as"], at["bytes"], at["sha256"][:12]))
            print(f"附件已归档 {len(saved)} 个 → {os.path.relpath(sub, root)}")
            for sp, sz, h in saved:
                print(f"   {os.path.basename(sp)}　{sz:,} bytes　sha256:{h}…")

        json.dump(arc_data, open(arc, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        print(f"确认单已归档：{os.path.relpath(arc, root)}")

    import datetime as _dt2
    bdir = os.path.join(os.path.dirname(os.path.abspath(a.master)), "_备份")
    os.makedirs(bdir, exist_ok=True)
    backup = os.path.join(bdir, f"{os.path.splitext(os.path.basename(a.master))[0]}_"
                                f"{_dt2.datetime.now():%Y%m%d_%H%M%S_%f}".replace("%f", "")[:-1] + ".xlsx")
    _n3 = 2
    while os.path.exists(backup):
        backup = backup[:-5] + f"_{_n3}.xlsx"; _n3 += 1
    shutil.copy2(a.master, backup)
    for t, c, v, cell in changes + (conflicts if a.force else []):
        cell.value = v
    _tmp = a.master + ".tmp.xlsx"
    wb.save(_tmp)
    os.replace(_tmp, a.master)
    print(f"\n已写入 {len(changes) + (len(conflicts) if a.force else 0)} 项；原文件备份于 _备份/{os.path.basename(backup)}")


if __name__ == "__main__":
    main()
