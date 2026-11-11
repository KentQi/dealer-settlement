"""端到端演示 / CI 夹具：生成演示数据 → load_master → recon_engine → 断言 8 表与恒等式。

用法：
    python3 examples/run_demo.py [--keep]

行为：
  1. 在系统临时目录生成演示数据（主档六表 + 202610 三份导出，全部 .xlsx）；
  2. 跑 load_master.py 生成 RECON_CFG，校验主档/合同/源文件发现结果；
  3. 跑 recon_engine.py，校验退出码 0、stderr 自检行、输出文件与 8 张工作表；
  4. 断言六条数量/金额恒等式、贴花独立复算、时点回加、差异（短缺+盘盈）、
     05 表动态门店列、06 表结算/非结算分桶等关键数字；
  5. 反向用例：缺配置 / 缺源文件 / 非法参数时退出码非 0 且 stdout 必须 0 字节；
  6. 泛化用例：换主体名/税号/合同号/收款户名/异议期后重跑，确认输出不再含演示固有文案；
  7. 确定性：同输入重跑一次，逐单元格比对两份输出完全一致；
  8. 静态检查脚本无网络依赖。

全部读写只在临时目录内，不写仓库、不联网。退出码 0 = 全部通过。
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SKILL_DIR = os.path.join(ROOT, "skills", "dealer-settlement")
SCRIPTS = os.path.join(SKILL_DIR, "scripts")
LOAD_MASTER = os.path.join(SCRIPTS, "load_master.py")
ENGINE = os.path.join(SCRIPTS, "recon_engine.py")

sys.path.insert(0, HERE)
import make_demo_data  # noqa: E402

import openpyxl  # noqa: E402

SHEETS = ["00-对账说明", "01-对账单", "02-库存差异表", "03-源数据对照",
          "04-甲方挂账明细", "05-乙方销售汇总", "06-乙方库存明细", "07-乙方销售流水",
          "08-效期预警", "09-试用装与配赠"]
PERIOD = "202610"
ASOF = "2026-11-02"


class DemoError(AssertionError):
    pass


def check(cond, msg):
    if not cond:
        raise DemoError(msg)


def approx(a, b, tol=0.005):
    return isinstance(a, (int, float)) and not isinstance(a, bool) and abs(float(a) - float(b)) <= tol


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def norm_label(v):
    return str(v or "").replace("\u3000", "").replace(" ", "").strip()


def _clean_env():
    """清理可能影响结果的变量；把代理指向黑洞端口，确保任何网络访问都会快速失败。"""
    drop = {"RECON_CFG", "HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "NO_PROXY"}
    env = {k: v for k, v in os.environ.items() if k.upper() not in drop}
    env.update({
        "PYTHONIOENCODING": "utf-8",
        "PYTHONUTF8": "1",
        "HTTP_PROXY": "http://127.0.0.1:9",
        "HTTPS_PROXY": "http://127.0.0.1:9",
        "ALL_PROXY": "http://127.0.0.1:9",
        "NO_PROXY": "127.0.0.1,localhost",
    })
    return env


def run_script(args, cfg=None, cwd=None):
    env = _clean_env()
    if cfg is not None:
        env["RECON_CFG"] = json.dumps(cfg, ensure_ascii=False)
    return subprocess.run([sys.executable] + args, capture_output=True,
                          env=env, cwd=cwd or ROOT)


def _text(b):
    return b.decode("utf-8", "replace")


def _rows(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _find_row(ws, pred):
    for i, row in enumerate(_rows(ws), 1):
        if pred(row):
            return i, row
    return None, None


def _cell_text(wb, needle):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and needle in c.value:
                    return ws.title, c
    return None, None


def _value_below(ws, needle):
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and needle in c.value:
                return ws.cell(row=c.row + 1, column=c.column).value
    return None


def _expected():
    """从演示数据定义推导期望值（单一事实来源）。"""
    price = {s["bar"]: float(s["supply_price"]) for s in make_demo_data.SKUS}
    std = {s["bar"]: float(s["std"]) for s in make_demo_data.SKUS}
    q = make_demo_data.QTY
    bars = [s["bar"] for s in make_demo_data.SKUS]
    exp = {
        "q0": sum(q[b]["q0"] for b in bars),
        "q10": sum(q[b]["q10"] for b in bars),
        "q11": sum(q[b]["q11"] for b in bars),
        "act": sum(q[b]["act"] for b in bars),
        "settle": round(sum(q[b]["q10"] * price[b] for b in bars), 2),
        "open": round(sum(q[b]["q0"] * price[b] for b in bars), 2),
        "nov": round(sum(q[b]["q11"] * price[b] for b in bars), 2),
        "act_val": round(sum(q[b]["act"] * price[b] for b in bars), 2),
        "th": round(sum(q[b]["q10"] * std[b] for b in bars), 2),
        "extra": make_demo_data.EXTRA_ACT,
        "skus": len(bars),
        "skus_sales": sum(1 for b in bars if q[b]["q10"] > 0),
        "stores": len(make_demo_data.STORES),
        "detail_rows": len(make_demo_data.SALES),
    }
    exp["theo"] = exp["q0"] - exp["q10"]
    exp["adj"] = exp["act"] + exp["q11"]
    exp["diff"] = exp["adj"] - exp["theo"]
    exp["diff_amt"] = round(sum(
        ((q[b]["act"] + q[b]["q11"]) - (q[b]["q0"] - q[b]["q10"])) * price[b] for b in bars), 2)
    exp["theo_val"] = round(sum((q[b]["q0"] - q[b]["q10"]) * price[b] for b in bars), 2)
    exp["adj_val"] = round(sum((q[b]["act"] + q[b]["q11"]) * price[b] for b in bars), 2)
    exp["book_end"] = round(exp["open"] - exp["settle"], 2)
    exp["settle_after"] = round(exp["settle"] - exp["diff_amt"], 2)
    return exp


def verify_cfg(cfg, customer_root):
    check(cfg.get("period", {}).get("start") == "2026-10-01", "CFG period.start 不为 2026-10-01")
    check(cfg.get("period", {}).get("end") == "2026-10-31", "CFG period.end 不为 2026-10-31")
    check(cfg.get("period", {}).get("asof") == ASOF, "CFG period.asof 不为 2026-11-02")
    check(cfg.get("master_path") == os.path.join(customer_root, "_主档.xlsx"), "CFG master_path 异常")
    check(os.path.isabs(cfg.get("base_dir", "")), "CFG base_dir 必须是绝对路径")
    check(cfg["base_dir"].endswith(PERIOD), "CFG base_dir 未定位到 202610 月份目录")
    ct = cfg.get("contract", {})
    check(ct.get("no") == "HT-DEMO-2026-001", "CFG 合同编号未从主档「合同」表读取")
    check(ct.get("method") == "铺货代销", "CFG 结算方式未从主档读取")
    check(ct.get("basis") == "实际动销", "CFG 结算基数未从主档读取")
    check(ct.get("objection_days") == 3, "CFG 异议期未从主档读取")
    check(ct.get("signback_days") == 5, "CFG 签回期未从主档读取")
    check(cfg.get("bank", {}).get("account_no") == "44056601040012345", "CFG 甲方账号异常")
    for key, kw in (("jia", "挂账"), ("sales", "销售数据"), ("inventory", "库存")):
        name = cfg.get("sources", {}).get(key, "")
        check(kw in name, f"CFG sources.{key} 未命中关键词：{name!r}")
        check(os.path.isfile(os.path.join(cfg["base_dir"], name)), f"CFG sources.{key} 文件不存在：{name}")


def verify_workbook(path, cfg, exp, generic=False, new_names=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        check(wb.sheetnames == SHEETS, f"工作表不是预期的 10 张：{wb.sheetnames}")
        ws0, ws1, ws2, ws3 = wb["00-对账说明"], wb["01-对账单"], wb["02-库存差异表"], wb["03-源数据对照"]
        ws4, ws5, ws6, ws7 = (wb["04-甲方挂账明细"], wb["05-乙方销售汇总"],
                              wb["06-乙方库存明细"], wb["07-乙方销售流水"])
        ws8, ws9 = wb["08-效期预警"], wb["09-试用装与配赠"]

        # ---------- 00 封面 / KPI ----------
        check(any(isinstance(c.value, str) and f"{cfg['period']['period_cn']}  铺货代销对账报告" in c.value
                  for r in ws0.iter_rows() for c in r), "00 标题未使用 CFG 期间")
        kpi_settle = _value_below(ws0, "本月应结算金额（元）")
        check(approx(kpi_settle, exp["settle"]), f"00 KPI 应结算金额 {kpi_settle} != {exp['settle']}")
        kpi_diff = _value_below(ws0, "库存差异（")
        check(approx(kpi_diff, abs(exp["diff_amt"])), f"00 KPI 差异金额 {kpi_diff} != {abs(exp['diff_amt'])}")
        concl3 = next((c.value for r in ws0.iter_rows() for c in r
                       if isinstance(c.value, str) and "逐 SKU 差异：" in c.value), "")
        check("+5 件" in concl3 and "-8 件" in concl3, f"00 结论 3 差异符号/数值错误：{concl3!r}")
        concl5 = next((c.value for r in ws0.iter_rows() for c in r
                       if isinstance(c.value, str) and "库存净差异" in c.value), "")
        check(f"{exp['diff']:+d} 件" in concl5,
              f"00 结论 5 净差异数值错误：{concl5!r}")

        # ---------- 01 对账单 ----------
        _, total_row = _find_row(ws1, lambda r: norm_label(r[0]) == "合计" and num(r[6]))
        check(total_row is not None, "01 未找到结算明细合计行")
        check(approx(total_row[5], exp["q10"]) and approx(total_row[6], exp["settle"]),
              f"01 结算合计 {total_row[5]}/{total_row[6]} != {exp['q10']}/{exp['settle']}")
        hdr_idx, _ = _find_row(ws1, lambda r: r[0] == "序号" and r[2] == "动销数量")
        check(hdr_idx is not None, "01 未找到贴花明细表头")
        th_detail, th_total = 0.0, None
        for row in _rows(ws1)[hdr_idx:]:
            if norm_label(row[0]) == "合计":
                th_total = row[4]
                break
            if num(row[0]) and num(row[4]):
                th_detail += row[4]
        check(th_total is not None and approx(th_total, exp["th"]),
              f"01 贴花合计 {th_total} != {exp['th']}")
        check(approx(th_detail, exp["th"]), f"01 贴花明细复算 {th_detail:.2f} != {exp['th']}")
        roll = {}
        for row in _rows(ws1):
            key = str(row[0] or "")[:1]
            if key in "①②③④⑤⑥⑦" and num(row[2]):
                roll[key] = row[2]
        check(len(roll) == 7, f"01 结转节未找到 7 行：{sorted(roll)}")
        check(approx(roll["①"], exp["open"]) and approx(roll["②"], exp["settle"]),
              "01 ①/② 与期望不符")
        check(approx(roll["③"], exp["book_end"]), f"01 ③ 期末货值 {roll['③']} != {exp['book_end']}")
        check(approx(roll["④"], exp["nov"]), f"01 ④ 回加货值 {roll['④']} != {exp['nov']}")
        check(approx(roll["⑥"], exp["diff_amt"]), f"01 ⑥ 差异金额 {roll['⑥']} != {exp['diff_amt']}")
        check(approx(roll["⑦"], exp["act_val"]), f"01 ⑦ 库存货值 {roll['⑦']} != {exp['act_val']}")
        # 恒等式：③ = ①-②；⑤ = ③-④；⑦ = ⑤+⑥
        check(abs((roll["①"] - roll["②"]) - roll["③"]) < 0.01, "01 恒等式 ③=①-② 不成立")
        check(abs((roll["③"] - roll["④"]) - roll["⑤"]) < 0.01, "01 恒等式 ⑤=③-④ 不成立")
        check(abs((roll["⑤"] + roll["⑥"]) - roll["⑦"]) < 0.01, "01 恒等式 ⑦=⑤+⑥ 不成立")
        note5 = next((c.value for r in ws1.iter_rows() for c in r
                      if isinstance(c.value, str) and "收到本对账单后" in c.value), "")
        if not generic:
            # 合同条款来自主档：异议期 3 / 签回期 5（泛化用例在下方单独校验新值）
            check("5 个工作日内" in note5 and "3 个工作日内" in note5,
                  f"01 未使用主档异议期/签回期：{note5!r}")
        bank = next((c.value for r in ws1.iter_rows() for c in r
                     if isinstance(c.value, str) and c.value.startswith("户　名：")), "")
        check(bank.endswith(cfg["bank"]["account_name"]), f"01 收款户名错误：{bank!r}")

        # ---------- 02 库存差异表 ----------
        _, d_tot = _find_row(ws2, lambda r: "一＋二" in str(r[0] or ""))
        check(d_tot is not None, "02 未找到一＋二合计行")
        q0, q10, theo = d_tot[6], -d_tot[7], d_tot[8]
        act, q11, adj, diff, diff_amt = d_tot[9], d_tot[10], d_tot[11], d_tot[12], d_tot[13]
        check((q0, q10, theo) == (exp["q0"], exp["q10"], exp["theo"]),
              f"02 数量合计 {(q0, q10, theo)} != {(exp['q0'], exp['q10'], exp['theo'])}")
        check((act, q11, adj, diff) == (exp["act"], exp["q11"], exp["adj"], exp["diff"]),
              f"02 库存/回加/差异 {(act, q11, adj, diff)} != {(exp['act'], exp['q11'], exp['adj'], exp['diff'])}")
        check(approx(diff_amt, exp["diff_amt"]), f"02 差异金额 {diff_amt} != {exp['diff_amt']}")
        check(q0 == q10 + theo and adj == act + q11 and diff == adj - theo, "02 三条数量恒等式不成立")
        diff_rows = [r for r in _rows(ws2)
                     if num(r[0]) and r[0] == int(r[0]) and num(r[12]) and r[12] != 0]
        check(len(diff_rows) == 2, f"02 差异行数 {len(diff_rows)} != 2")
        by_bar = {str(r[1]): r for r in diff_rows}
        check(by_bar["690000000002"][12] == -8 and "短缺" in str(by_bar["690000000002"][15]),
              "02 短缺行性质/数量错误")
        check(by_bar["690000000012"][12] == 5 and "盘盈" in str(by_bar["690000000012"][15]),
              "02 盘盈行性质/数量错误")
        zero_note = " ".join(str(c.value) for r in ws2.iter_rows() for c in r
                             if isinstance(c.value, str) and "本期期初铺货数量为 0" in c.value)
        check("示例精华蜜15ml" in zero_note, f"02 零库存说明未按实际数据生成：{zero_note!r}")

        # ---------- 03 源数据对照 ----------
        _, t3 = _find_row(ws3, lambda r: norm_label(r[0]) == "合计")
        check(t3 is not None, "03 未找到合计行")
        check(approx(t3[7], exp["q0"]) and approx(t3[15], exp["diff_amt"]),
              f"03 合计 {t3[7]}/{t3[15]} 与期望不符")
        n3 = sum(1 for r in _rows(ws3) if num(r[0]) and r[0] == int(r[0]) and r[3] is not None)
        check(n3 == exp["skus"], f"03 明细行数 {n3} != {exp['skus']}")

        # ---------- 04 甲方挂账明细 ----------
        _, t4 = _find_row(ws4, lambda r: norm_label(r[0]) == "合计")
        check(t4 is not None, "04 未找到合计行")
        check(approx(t4[3], exp["open"]) and t4[4] == exp["q0"], f"04 合计 {t4[3]}/{t4[4]} 与期望不符")
        check(abs(t4[8] or 0) < 0.01, f"04 金额校验合计应为 0，实际 {t4[8]}")

        # ---------- 05 乙方销售汇总（动态门店列）----------
        h5, hrow5 = _find_row(ws5, lambda r: r[0] == "商品条码" and any("门店" in str(x) for x in r if x))
        check(hrow5 is not None, "05 未找到按门店汇总表头")
        check(len([x for x in hrow5 if x and "门店" in str(x)]) == exp["stores"],
              f"05 门店列数不为 {exp['stores']}")
        total_idx = next(i for i, x in enumerate(hrow5) if norm_label(x) == "合计")
        ratio_idx = next(i for i, x in enumerate(hrow5) if norm_label(x) == "占比")
        _, t5 = _find_row(ws5, lambda r: norm_label(r[0]) == "合计" and num(r[total_idx]))
        check(t5 is not None and approx(t5[total_idx], exp["q10"]),
              f"05 合计数量 {None if t5 is None else t5[total_idx]} != {exp['q10']}（动态列对齐失败）")
        check(approx(t5[ratio_idx], 1.0), "05 合计占比不为 1.0")
        check(approx(sum(t5[4:4 + exp["stores"]]), exp["q10"]), "05 门店列求和 != 合计")
        n5 = sum(1 for r in _rows(ws5) if num(r[total_idx]) and r[0] != "合　计")
        check(n5 == exp["skus_sales"], f"05 汇总 SKU 行数 {n5} != {exp['skus_sales']}")

        # ---------- 06 乙方库存明细 ----------
        _, s6 = _find_row(ws6, lambda r: str(r[0] or "").startswith("结算类小计"))
        _, n6 = _find_row(ws6, lambda r: str(r[0] or "").startswith("非结算类小计"))
        check(s6 is not None and n6 is not None, "06 未找到结算类/非结算类小计行")
        check(s6[5] == exp["act"], f"06 结算类库存 {s6[5]} != {exp['act']}")
        check(approx(s6[10], exp["act_val"]), f"06 结算类货值 {s6[10]} != {exp['act_val']}")
        check(n6[5] == exp["extra"], f"06 非结算类库存 {n6[5]} != {exp['extra']}")

        # ---------- 07 销售流水 ----------
        _, t7 = _find_row(ws7, lambda r: norm_label(r[0]) == "合计")
        check(t7 is not None and t7[7] == exp["q10"], f"07 流水合计 {None if t7 is None else t7[7]} != {exp['q10']}")
        n7 = sum(1 for r in _rows(ws7) if isinstance(r[0], str) and r[0].startswith("XS"))
        check(n7 == exp["detail_rows"], f"07 流水行数 {n7} != {exp['detail_rows']}")
        note7 = next((c.value for r in ws7.iter_rows() for c in r
                      if isinstance(c.value, str) and "核对说明" in c.value), "")
        check("未作去重" in note7, f"07 拆行说明未声明口径：{note7!r}")

        # ---------- 08 效期预警 ----------
        t8 = "\n".join(str(c.value) for r in ws8.iter_rows() for c in r if isinstance(c.value, str))
        check("附 表 08" in t8, "08 表标题缺失")
        for tier in ("已过期", "30 天内", "90 天内", "180 天内", "正常"):
            check(tier in t8, f"08 表缺少效期分档：{tier}")
        check("FIFO" in t8, "08 表缺少 FIFO 口径说明")
        check("后进先出迹象" in t8, "08 表未识别出后进先出迹象（夹具已构造该场景）")
        check("批次结存与账面逐条码一致" in t8, "08 表未完成批次结存与账面勾稽")

        # ---------- 09 试用装与配赠 ----------
        t9 = "\n".join(str(c.value) for r in ws9.iter_rows() for c in r if isinstance(c.value, str))
        check("附 表 09" in t9, "09 表标题缺失")
        for kind in ("试用装", "中小样", "配赠"):
            check(kind in t9, f"09 表缺少物料属性：{kind}")
        check("不参与货款结算" in t9, "09 表未声明不参与货款结算")

        if generic:
            values = [c.value for ws in wb.worksheets for row in ws.iter_rows() for c in row]
            for old in ("示例品牌方有限公司", "示例经销商有限公司", "HT-DEMO-2026-001"):
                check(old not in values, f"泛化用例仍含演示文案：{old}")
            for name in new_names:
                check(name in values, f"泛化用例未使用新值：{name}")
            note5g = next((c.value for r in ws1.iter_rows() for c in r
                           if isinstance(c.value, str) and "收到本对账单后" in c.value), "")
            check("7 个工作日内" in note5g and "4 个工作日内" in note5g,
                  f"泛化用例未使用新的异议期/签回期：{note5g!r}")
            bank_g = next((c.value for r in ws1.iter_rows() for c in r
                           if isinstance(c.value, str) and c.value.startswith("户　名：")), "")
            check(bank_g.endswith("另示例收款户名有限公司"), f"泛化用例收款户名未按 CFG 输出：{bank_g!r}")
            check(approx(_value_below(ws0, "本月应结算金额（元）"), exp["settle"]), "泛化用例金额变化")
    finally:
        wb.close()


def _value_map(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        return {(ws.title, c.row, c.column): c.value
                for ws in wb.worksheets for row in ws.iter_rows() for c in row}
    finally:
        wb.close()


def static_no_network():
    banned = ("import requests", "urllib.request", "http.client", "socket.socket",
              "import httpx", "import aiohttp")
    self_path = os.path.abspath(__file__)
    for directory in (SCRIPTS, HERE):
        for name in sorted(os.listdir(directory)):
            if not name.endswith(".py"):
                continue
            if os.path.abspath(os.path.join(directory, name)) == self_path:
                continue  # 本文件包含被禁字符串常量本身，跳过自检
            with open(os.path.join(directory, name), encoding="utf-8") as fh:
                src = fh.read()
            for token in banned:
                check(token not in src, f"{name} 含网络依赖：{token}")


def main(argv=None):
    ap = argparse.ArgumentParser(description="铺货代销对账端到端 demo / CI 夹具")
    ap.add_argument("--keep", action="store_true", help="保留临时目录（失败时总是保留）")
    a = ap.parse_args(argv)

    tmpdir = tempfile.mkdtemp(prefix="dealer_settlement_demo_")
    ok = False
    try:
        exp = _expected()
        customer_root, paths = make_demo_data.build(os.path.join(tmpdir, "铺货代销客户"))
        check(os.path.isfile(paths["master"]), "演示主档未生成")

        # ---------- 1. load_master ----------
        p = run_script([LOAD_MASTER, paths["master"], "--period", PERIOD, "--asof", ASOF], cwd=tmpdir)
        check(p.returncode == 0, f"load_master 失败（rc={p.returncode}）：{_text(p.stderr)}")
        check(p.stdout.strip(), "load_master 未输出 JSON")
        cfg = json.loads(_text(p.stdout))
        verify_cfg(cfg, customer_root)
        check("源文件" not in _text(p.stdout), "load_master stdout 混入非 JSON 内容")

        # ---------- 2. recon_engine 正向 ----------
        p = run_script([ENGINE], cfg, cwd=tmpdir)
        check(p.returncode == 0, f"recon_engine 失败（rc={p.returncode}）：{_text(p.stderr)[-2000:]}")
        stderr = _text(p.stderr)
        check("自检通过（残差 0.00）" in stderr, "引擎未打印自检通过行")
        out_path = os.path.join(cfg["base_dir"], cfg["period"]["out_name"])
        check(os.path.isfile(out_path), f"引擎未生成输出文件：{out_path}")
        check(f"SAVED: {out_path}" in _text(p.stdout), "引擎 stdout 未报告 SAVED 路径")
        check("[新建]" in _text(p.stdout), "首次生成应为 [新建] 标记")
        check("出现多行" in stderr, "多行库存聚合告警未触发")

        # ---------- 3. 10 表与恒等式 ----------
        verify_workbook(out_path, cfg, exp)
        wb_tmp = openpyxl.load_workbook(out_path, data_only=True)
        try:
            t8 = "\n".join(str(c.value) for r in wb_tmp["08-效期预警"].iter_rows()
                            for c in r if isinstance(c.value, str))
            check("910 件" in t8, "08 表期末结存合计应为 910 件（= 账面理论结存）")
            t9 = "\n".join(str(c.value) for r in wb_tmp["09-试用装与配赠"].iter_rows()
                            for c in r if isinstance(c.value, str))
            check("1,150 件" in t9, "09 表促销物料出库合计应为 1,150 件")
            check("乙方未提供使用数据" in t9, "未生成使用表时应走「未提供」分支")
        finally:
            wb_tmp.close()

        # ---------- 4. 失败路径：退出码非 0 且 stdout 0 字节 ----------
        for label, cfg_bad, needle in (
                ("缺 RECON_CFG", None, "[中止]"),
                ("缺源文件", {**json.loads(json.dumps(cfg)),
                            "sources": {**cfg["sources"], "jia": "不存在的挂账.xlsx"}}, "[缺少源文件]"),
                ("非法 RECON_CFG", "{not-json", "RECON_CFG")):
            p_bad = run_script([ENGINE], cfg_bad, cwd=tmpdir)
            check(p_bad.returncode != 0, f"{label} 应失败但 rc=0")
            check(p_bad.stdout == b"", f"{label} 失败时 stdout 必须 0 字节，实际 {p_bad.stdout[:120]!r}")
            check(needle in _text(p_bad.stderr), f"{label} stderr 缺少「{needle}」：{_text(p_bad.stderr)[:300]}")
        p_bad = run_script([LOAD_MASTER, paths["master"], "--period", "20261"], cwd=tmpdir)
        check(p_bad.returncode != 0 and p_bad.stdout == b"", "load_master 非法 --period 应失败且 stdout 0 字节")

        # ---------- 5. 泛化：换主体/合同/账户条款 ----------
        cfg2 = json.loads(json.dumps(cfg))
        new_names = ("另示例甲方有限公司", "另示例乙方有限公司")
        cfg2["party_a"].update({"name": new_names[0], "tax": "91000000TEST00001X"})
        cfg2["party_b"].update({"name": new_names[1], "tax": "91000000TEST00002Y"})
        cfg2["contract"].update({"no": "HT-TEST-2026-999", "objection_days": 4, "signback_days": 7})
        cfg2["bank"] = {"account_name": "另示例收款户名有限公司", "account_no": "6222000011112222",
                        "bank_name": "另示例银行"}
        cfg2["out_file"] = os.path.join(cfg["base_dir"], "泛化校验对账.xlsx")
        p = run_script([ENGINE], cfg2, cwd=tmpdir)
        check(p.returncode == 0, f"泛化用例失败：{_text(p.stderr)[-2000:]}")
        verify_workbook(cfg2["out_file"], cfg2, exp, generic=True, new_names=new_names)

        # ---------- 5b. 覆盖保护：目标已存在时写 _v2 且 stdout 标记为 [覆盖→...] ----------
        p = run_script([ENGINE], cfg, cwd=tmpdir)
        check(p.returncode == 0, f"覆盖保护用例失败：{_text(p.stderr)[-1000:]}")
        v2 = os.path.join(cfg["base_dir"], "202610铺货代销对账_v2.xlsx")
        check(os.path.isfile(v2), "目标文件已存在时未生成 _v2 文件")
        check("[覆盖→202610铺货代销对账_v2.xlsx]" in _text(p.stdout),
              f"覆盖保护 stdout 标记错误：{_text(p.stdout)!r}")

        # ---------- 6. 确定性：同输入重跑，逐单元格一致 ----------
        cfg3 = json.loads(json.dumps(cfg))
        cfg3["out_file"] = os.path.join(cfg["base_dir"], "确定性复核对账.xlsx")
        p = run_script([ENGINE], cfg3, cwd=tmpdir)
        check(p.returncode == 0, f"确定性复跑失败：{_text(p.stderr)[-1000:]}")
        check(_value_map(out_path) == _value_map(cfg3["out_file"]),
              "同输入两次运行的单元格值不一致（输出不确定）")

        # ---------- 7. 静态无网络 ----------
        static_no_network()

        # ---------- 8b. 批次字段合写「批次号/有效期」----------
        tmp3 = tempfile.mkdtemp(prefix="dealer_settlement_demo_lot_")
        try:
            root3, paths3 = make_demo_data.build(os.path.join(tmp3, "铺货代销客户"))
            wbj = openpyxl.load_workbook(paths3["jia"])
            wsj = wbj.active
            head3 = [c.value for c in wsj[2]]
            i_lot = head3.index("批次号") + 1
            i_exp = head3.index("有效期至") + 1
            for rr in range(3, wsj.max_row + 1):
                v = wsj.cell(row=rr, column=i_exp).value
                lot_v = wsj.cell(row=rr, column=i_lot).value
                if v and lot_v:
                    d = v.replace("-", "") if isinstance(v, str) else v.strftime("%Y%m%d")
                    wsj.cell(row=rr, column=i_lot, value="%s/%s" % (lot_v, d))
            wsj.delete_cols(i_exp)          # 只留合写的「批次」列
            wbj.save(paths3["jia"])
            wbj.close()

            p = run_script([LOAD_MASTER, os.path.join(root3, "_主档.xlsx"),
                            "--period", PERIOD, "--asof", ASOF], cwd=tmp3)
            check(p.returncode == 0, f"合写批次场景 load_master 失败：{_text(p.stderr)[:400]}")
            cfg_lot = json.loads(_text(p.stdout))
            p = run_script([ENGINE], cfg_lot, cwd=tmp3)
            check(p.returncode == 0, f"合写批次场景引擎失败：{_text(p.stderr)[-1500:]}")
            out3 = os.path.join(cfg_lot["base_dir"], cfg_lot["period"]["out_name"])
            wb3 = openpyxl.load_workbook(out3, data_only=True)
            try:
                # 有效期是日期单元格（datetime），不能只取 str，否则取不到日期文本
                t8c = "\n".join(str(c.value) for r in wb3["08-效期预警"].iter_rows()
                                 for c in r if c.value not in (None, ""))
                check("B1-202611" in t8c, "合写批次未拆出批次号（应为 B1-202611）")
                check("2026-11-20" in t8c, "合写批次未拆出有效期（应为 2026-11-20）")
                check("B1-202611/" not in t8c, "合写字段未被拆分，批次号仍带着日期")
                check("批次结存与账面逐条码一致" in t8c, "合写批次场景批次结存勾稽失败")
            finally:
                wb3.close()
        finally:
            shutil.rmtree(tmp3, ignore_errors=True)

        # ---------- 8. 乙方提供使用数据的分支（试用装/中小样/配赠）----------
        tmp2 = tempfile.mkdtemp(prefix="dealer_settlement_demo_use_")
        try:
            root2, _ = make_demo_data.build(os.path.join(tmp2, "铺货代销客户"), with_promo_use=True)
            p = run_script([LOAD_MASTER, os.path.join(root2, "_主档.xlsx"),
                            "--period", PERIOD, "--asof", ASOF], cwd=tmp2)
            check(p.returncode == 0, f"使用表场景 load_master 失败：{_text(p.stderr)[:400]}")
            cfg_use = json.loads(_text(p.stdout))
            check(cfg_use["sources"].get("promo_use"), "load_master 未识别试用装使用表")
            p = run_script([ENGINE], cfg_use, cwd=tmp2)
            check(p.returncode == 0, f"使用表场景引擎失败：{_text(p.stderr)[-1500:]}")
            out2 = os.path.join(cfg_use["base_dir"], cfg_use["period"]["out_name"])
            wb2 = openpyxl.load_workbook(out2, data_only=True)
            try:
                t9b = "\n".join(str(c.value) for r in wb2["09-试用装与配赠"].iter_rows()
                                 for c in r if isinstance(c.value, str))
                check("乙方使用表" in t9b, "09 表未标注乙方使用数据来源")
                check("乙方已提供使用数据" in t9b, "09 表未进入「已提供」分支")
                check("1,150 件" in t9b, "09 表出库合计异常")
                check("差异不计算" not in t9b, "已提供使用数据时不应出现「差异不计算」")
            finally:
                wb2.close()
        finally:
            shutil.rmtree(tmp2, ignore_errors=True)

        ok = True
        print("[run_demo] OK")
        print(f"  load_master : rc=0，月份目录={os.path.basename(cfg['base_dir'])}，"
              f"源文件={cfg['sources']['jia']} / {cfg['sources']['sales']} / {cfg['sources']['inventory']}")
        print(f"  recon_engine: rc=0，自检通过（残差 0.00），10 表 = {', '.join(SHEETS)}")
        print(f"  关键数      : 期初 {exp['q0']} 件/{exp['open']:.2f} 元；"
              f"结算 {exp['q10']} 件/{exp['settle']:.2f} 元；"
              f"差异 {exp['diff']:+d} 件/{exp['diff_amt']:.2f} 元；"
              f"贴花 {exp['th']:.2f} 元；回加 {exp['q11']} 件/{exp['nov']:.2f} 元")
        print("  断言        : 六条恒等式 / 贴花复算 / 动态门店列 / 结算分桶 / "
              "失败路径 stdout=0 / 泛化 / 确定性 / 无网络 全部通过")
        return 0
    except DemoError as e:
        print("[run_demo] FAIL: " + str(e), file=sys.stderr)
        print(f"[run_demo] 临时目录已保留：{tmpdir}", file=sys.stderr)
        return 1
    except Exception as e:  # noqa: BLE001 - CLI 顶层兜底，输出可读错误
        import traceback
        traceback.print_exc()
        print(f"[run_demo] 临时目录已保留：{tmpdir}", file=sys.stderr)
        return 1
    finally:
        if ok and not a.keep and os.path.isdir(tmpdir):
            shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
