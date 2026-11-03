"""把合同候选字段 + 主档现值渲染成一份可交互确认页（单文件 HTML，无外部依赖）。

用户在浏览器里逐项确认、修改，点「提交」下载 确认单.json；该 JSON 可直接喂给
apply_contract.py 写库，同时页面可打印成 PDF 留档。支持附件随单归档。

用法：
    python3 make_confirm_html.py <候选.json> <_主档.xlsx> [-o 确认页.html]
"""
import sys, os, json, argparse, datetime, html
import openpyxl


def load_master_values(path):
    """返回 {target: 现值}，target 形如 主体.甲方.名称 / 合同.合同编号。"""
    out = {}
    if not os.path.exists(path):
        return out
    wb = openpyxl.load_workbook(path, data_only=True)

    if "主体" in wb.sheetnames:
        ws = wb["主体"]
        headers = [c.value for c in ws[1]]
        if "角色" in headers:
            ri = headers.index("角色") + 1
            for r in range(2, ws.max_row + 1):
                role = ws.cell(row=r, column=ri).value
                if not role:
                    continue
                for ci, h in enumerate(headers, 1):
                    if h and h != "角色":
                        out[f"主体.{role}.{h}"] = ws.cell(row=r, column=ci).value

    for sh in ("合同", "价格", "SKU"):
        if sh not in wb.sheetnames:
            continue
        ws = wb[sh]
        headers = [c.value for c in ws[1]]
        for ci, h in enumerate(headers, 1):
            if h:
                out[f"{sh}.{h}"] = ws.cell(row=2, column=ci).value

    if "合同" in wb.sheetnames:
        ws = wb["合同"]
        headers = [c.value for c in ws[1]]
        out["_合同编号"] = ws.cell(row=2, column=headers.index("合同编号") + 1).value if "合同编号" in headers else ""
    if "主体" in wb.sheetnames:
        ws = wb["主体"]
        headers = [c.value for c in ws[1]]
        if "角色" in headers and "名称" in headers:
            ri, ni = headers.index("角色") + 1, headers.index("名称") + 1
            for r in range(2, ws.max_row + 1):
                if str(ws.cell(row=r, column=ri).value).strip() == "乙方":
                    out["_客户名称"] = ws.cell(row=r, column=ni).value
    return out


def build_rows(cand, master):
    rows = []
    for f in cand.get("fields", []):
        cur = master.get(f.get("target", ""), None)
        new = f.get("value")
        cur_s = "" if cur is None else str(cur).strip()
        new_s = "" if new is None else str(new).strip()
        if not new_s:
            status = "待核"
        elif not cur_s:
            status = "新增"
        elif cur_s == new_s:
            status = "一致"
        else:
            status = "冲突"
        rows.append({
            "field": f.get("field", ""),
            "target": f.get("target", ""),
            "current": cur_s,
            "candidate": new_s,
            "evidence": f.get("evidence") or "",
            "confidence": f.get("confidence", ""),
            "status": status,
        })
    order = {"冲突": 0, "待核": 1, "新增": 2, "一致": 3}
    rows.sort(key=lambda x: (order.get(x["status"], 9), x["field"]))
    return rows


TPL = r"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>合同信息确认单 · __TITLE__</title>
<style>
:root{
  --navy:#1F3864;--blue:#2F5597;--line:#E2E6EE;--soft:#F6F8FB;
  --red:#C0392B;--amber:#B7791F;--green:#2E6B30;--ink:#1A1D24;--muted:#6B7280;
}
*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{background:#EFF2F7;color:var(--ink);
  font:14px/1.62 "PingFang SC","Microsoft YaHei",system-ui,-apple-system,sans-serif;
  -webkit-font-smoothing:antialiased}
a{color:var(--blue)}
.wrap{max-width:1240px;margin:0 auto;padding:0 20px 140px}

/* 头部 */
header.top{background:linear-gradient(135deg,#1F3864,#2E4A7D);color:#fff;padding:20px 24px 18px;border-radius:12px 12px 0 0}
header.top h1{margin:0;font-size:19px;font-weight:600;letter-spacing:.4px;display:flex;align-items:center;gap:10px}
header.top h1 .dot{width:8px;height:8px;border-radius:50%;background:#7FB3FF;display:inline-block}
header.top .sub{opacity:.82;font-size:12.5px;margin-top:7px;line-height:1.7}

/* 统计 */
.chips{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px}
.chip{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.18);color:#fff;
  border-radius:999px;padding:4px 12px;font-size:12.5px;display:flex;align-items:center;gap:6px}
.chip b{font-size:15px;font-weight:700}
.chip.click{cursor:pointer;transition:.12s}
.chip.click:hover{background:rgba(255,255,255,.24)}
.chip.i0 b{color:#FF9A8F} .chip.i1 b{color:#FFD479}
.chip.i2 b{color:#9FC4FF} .chip.i3 b{color:#9BE3A6}

/* 工具条 */
.toolbar{position:sticky;top:0;z-index:30;background:#fff;border:1px solid var(--line);
  border-top:none;padding:10px 16px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;
  box-shadow:0 2px 8px rgba(20,30,60,.04)}
.search{flex:1 1 220px;min-width:180px;position:relative}
.search input{width:100%;padding:8px 12px 8px 32px;border:1px solid var(--line);border-radius:8px;font:inherit;font-size:13px;background:var(--soft)}
.search input:focus{outline:2px solid #C9D6EC;background:#fff;border-color:var(--blue)}
.search::before{content:"⌕";position:absolute;left:11px;top:50%;transform:translateY(-50%);color:#9AA4B2;font-size:15px}
.filters{display:flex;gap:4px;background:var(--soft);padding:3px;border-radius:9px;border:1px solid var(--line)}
.filters button{border:none;background:transparent;padding:6px 11px;border-radius:7px;font:inherit;font-size:12.5px;color:#4B5563;cursor:pointer;white-space:nowrap}
.filters button:hover{background:#fff}
.filters button.on{background:#fff;color:var(--navy);font-weight:600;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.bulk{display:flex;gap:6px}
.bulk button{border:1px solid var(--line);background:#fff;padding:6px 10px;border-radius:7px;font:inherit;font-size:12.5px;cursor:pointer;color:#374151}
.bulk button:hover{border-color:var(--blue);color:var(--blue)}
.stat{margin-left:auto;font-size:12.5px;color:var(--muted);white-space:nowrap}
.stat b{color:var(--navy);font-size:14px}

/* 分组 */
section.grp{background:#fff;border:1px solid var(--line);border-radius:10px;margin-top:14px;overflow:hidden}
.grp>h2{margin:0;padding:11px 16px;font-size:13.5px;background:var(--soft);border-bottom:1px solid var(--line);
  display:flex;align-items:center;gap:9px;cursor:pointer;user-select:none}
.grp>h2 .tw{color:#9AA4B2;font-size:11px;transition:transform .15s}
.grp.collapsed .tw{transform:rotate(-90deg)}
.grp>h2 small{font-weight:400;color:var(--muted);font-size:12px}
.grp>h2 .cnt{margin-left:auto;background:#fff;border:1px solid var(--line);border-radius:999px;padding:1px 9px;font-size:11.5px;color:#4B5563}
.badge{display:inline-flex;align-items:center;gap:4px;font-size:11.5px;padding:1px 8px;border-radius:999px;white-space:nowrap}
.b0{background:#FDECEA;color:var(--red)} .b1{background:#FDF3E2;color:var(--amber)}
.b2{background:#E9F0FA;color:var(--blue)} .b3{background:#E9F5EA;color:var(--green)}

table.t{width:100%;border-collapse:collapse}
table.t thead th{background:#FBFCFE;font-size:11.5px;color:#54606F;font-weight:600;text-align:left;
  padding:9px 14px;border-bottom:1px solid var(--line);letter-spacing:.3px}
table.t td{padding:11px 14px;border-bottom:1px solid #F1F3F7;vertical-align:top;font-size:13px}
table.t tr:last-child td{border-bottom:none}
table.t tr.dirty{background:#FFFBF0}
table.t tr.dirty td:first-child{box-shadow:inset 3px 0 0 #E0A82E}
.fname{font-weight:600;color:#111827}
.fpath{font-weight:400;color:#9AA4B2;font-size:11px;margin-top:2px}
.cur{color:#6B7280;font-size:12.5px;word-break:break-all}
.cur.nil{color:#C3C9D4}
.inp{width:100%;padding:7px 9px;border:1px solid var(--line);border-radius:7px;font:inherit;font-size:13px;background:#fff}
.inp:focus{outline:2px solid #C9D6EC;border-color:var(--blue)}
.inp.changed{border-color:#E0A82E;background:#FFFDF6}
mark{background:#FFE9A8;color:#7A4E00;padding:0 1px;border-radius:2px}
.ev{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:11.5px;line-height:1.55;
  color:#4B5563;background:#F7F8FA;border-left:3px solid #D7DDE8;padding:5px 9px;border-radius:0 4px 4px 0;
  max-height:76px;overflow:auto;word-break:break-all}
.noev{color:#C3C9D4;font-size:12px}
.use label{display:inline-flex;align-items:center;gap:6px;cursor:pointer;font-size:12.5px;color:#374151;white-space:nowrap}
input[type=checkbox]{width:16px;height:16px;accent-color:var(--blue);cursor:pointer}
.undo{border:none;background:none;color:var(--blue);font-size:11.5px;cursor:pointer;padding:0;margin-top:4px;text-decoration:underline dotted}
.undo:hover{color:var(--navy)}
.empty{padding:26px;text-align:center;color:var(--muted);font-size:13px}

/* 附件 */
.attbox{margin-top:16px}
.drop{border:1.5px dashed #B9C6DC;border-radius:10px;padding:20px;text-align:center;color:#5A6B85;
  background:#FBFCFE;cursor:pointer;transition:.15s;font-size:13px}
.drop:hover,.drop.hover{background:#EFF4FC;border-color:var(--blue);color:var(--navy)}
.drop b{color:var(--blue)}
.att{width:100%;border-collapse:collapse;margin-top:10px}
.att td{padding:8px 10px;border-bottom:1px solid #F1F3F7;font-size:12.5px;vertical-align:middle}
.att .nm{font-weight:600;color:#111827}
.att .sz{color:#6B7280;white-space:nowrap}
.att input{width:100%;padding:5px 8px;border:1px solid var(--line);border-radius:6px;font:inherit;font-size:12.5px}
.att .ops{white-space:nowrap;text-align:right}
.att .ops a,.att .ops span{color:var(--blue);cursor:pointer;font-size:12px;margin-left:10px}
.att .ops span.rm{color:var(--red)}
.hint{font-size:12px;color:var(--muted);margin-top:8px}
.warn{color:var(--red);font-weight:600}

/* 底栏 */
footer.bar{position:fixed;left:0;right:0;bottom:0;background:#fff;border-top:1px solid var(--line);
  padding:11px 20px;display:flex;gap:12px;align-items:center;z-index:40;box-shadow:0 -4px 16px rgba(20,30,60,.07)}
footer .who{display:flex;gap:12px;align-items:center;flex-wrap:wrap;font-size:12.5px;color:#374151}
footer input[type=text]{padding:7px 10px;border:1px solid var(--line);border-radius:7px;font:inherit;font-size:13px}
footer .sp{margin-left:auto}
button.btn{padding:9px 18px;border-radius:8px;border:1px solid var(--navy);background:var(--navy);color:#fff;
  font:inherit;font-size:13.5px;cursor:pointer;transition:.12s}
button.btn:hover{background:#2A4A7C}
button.btn.ghost{background:#fff;color:var(--navy)}
button.btn.ghost:hover{background:#F3F6FB}
button.btn:disabled{opacity:.45;cursor:not-allowed}

/* 弹窗 */
.mask{position:fixed;inset:0;background:rgba(17,24,39,.45);display:none;align-items:center;justify-content:center;z-index:60;padding:20px}
.mask.on{display:flex}
.modal{background:#fff;border-radius:12px;max-width:520px;width:100%;padding:22px 24px;box-shadow:0 20px 50px rgba(0,0,0,.2)}
.modal h3{margin:0 0 12px;font-size:16px}
.modal .row{display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px dashed #EEF1F6;font-size:13px}
.modal .row b{color:var(--navy)}
.modal .warnline{margin-top:12px;background:#FDECEA;color:var(--red);padding:9px 12px;border-radius:7px;font-size:12.5px;display:none}
.modal .acts{display:flex;gap:10px;justify-content:flex-end;margin-top:18px}

/* 提示条 */
.toast{position:fixed;left:50%;bottom:86px;transform:translateX(-50%) translateY(10px);opacity:0;
  background:#111827;color:#fff;padding:10px 18px;border-radius:8px;font-size:13px;z-index:70;
  transition:.2s;pointer-events:none;max-width:80vw}
.toast.on{opacity:1;transform:translateX(-50%) translateY(0)}

@media (max-width:860px){
  table.t thead{display:none}
  table.t td{display:block;border:none;padding:6px 14px}
  table.t tr{display:block;border-bottom:1px solid var(--line);padding:8px 0}
  table.t td:before{content:attr(data-l);display:block;font-size:11px;color:#9AA4B2;margin-bottom:3px}
  .stat{width:100%;margin:0}
  footer.bar{flex-wrap:wrap}
  footer .sp{width:100%;margin:8px 0 0}
}
@media print{
  body{background:#fff}
  .wrap{padding:0;max-width:none}
  header.top{border-radius:0;-webkit-print-color-adjust:exact}
  .toolbar,footer.bar,.no-print,.undo,.ops{display:none!important}
  .grp.collapsed table{display:table-row-group}
  .signblock{display:block!important}
  section.grp{break-inside:avoid;border-color:#ccc;margin-top:10px}
  section.grp.collapsed table{display:table!important}
  table.t thead{display:table-header-group}
  table.t tr{break-inside:avoid}
  .inp{border:none!important;border-bottom:1px dotted #888!important;border-radius:0;background:#fff!important;padding:0 2px}
  .ev{max-height:none;overflow:visible;break-inside:avoid}
  .mask,.toast{display:none!important}
}
</style></head><body>
<div class="wrap">

<header class="top">
  <h1><span class="dot"></span>合同信息确认单</h1>
  <div class="sub">__SUBTITLE__</div>
  <div class="chips">
    <span class="chip i0 click" data-f="冲突" title="点击只看冲突项">冲突<b id="c0">0</b></span>
    <span class="chip i1 click" data-f="待核" title="点击只看待核项">待核<b id="c1">0</b></span>
    <span class="chip i2 click" data-f="新增" title="点击只看新增项">新增<b id="c2">0</b></span>
    <span class="chip i3 click" data-f="一致" title="点击只看一致项">一致<b id="c3">0</b></span>
    <span class="chip">来源 <b style="font-weight:400;font-size:12.5px" id="src">__SRC__</b></span>
  </div>
</header>

<div class="toolbar no-print">
  <div class="search"><input id="q" type="text" placeholder="搜索字段名、现值或候选值…"></div>
  <div class="filters" id="filters">
    <button data-f="all" class="on">全部</button>
    <button data-f="冲突">冲突</button>
    <button data-f="待核">待核</button>
    <button data-f="新增">新增</button>
    <button data-f="一致">一致</button>
    <button data-f="dirty">已改动</button>
  </div>
  <div class="bulk">
    <button id="selSame">采纳全部一致项</button>
    <button id="unsAll">全部不采用</button>
    <button id="resetAll">重置改动</button>
  </div>
  <div class="stat" id="stat"></div>
</div>

<main id="app"></main>

<div class="signblock">
  <table style="width:100%;margin-top:18px;border-collapse:collapse">
    <tr>
      <td style="border:1px solid #999;padding:10px 14px;width:33%">确认人：＿＿＿＿＿＿＿＿</td>
      <td style="border:1px solid #999;padding:10px 14px;width:33%">日期：＿＿＿＿＿＿＿＿</td>
      <td style="border:1px solid #999;padding:10px 14px">复核人：＿＿＿＿＿＿＿＿</td>
    </tr>
  </table>
  <p style="font-size:11px;color:#666;margin-top:8px">本页为合同信息确认留档，与确认单 JSON 具同等效力；勾选「采用此值」的项将写入客户主档。</p>
</div>

<section class="grp attbox no-print">
  <h2 style="cursor:default"><span class="tw">▾</span>附件<small>营业执照、价格表、补充协议等，随确认单一并归档</small></h2>
  <div style="padding:14px 16px 18px">
    <div class="drop" id="drop">把文件拖到这里，或 <b>点击选择</b>（可多选）<br>
      <span style="font-size:12px;color:#8A94A6">单文件 ≤5MB，合计 ≤15MB</span>
      <input type="file" id="files" multiple style="display:none">
    </div>
    <table class="att" id="attTable" style="display:none">
      <tbody></tbody>
    </table>
    <div class="hint" id="attHint">尚未添加附件</div>
  </div>
</section>

<div class="hint" style="padding:14px 2px 0">
  勾选「采用」＝把该值写入主档；不勾选＝保留主档现值。冲突项默认不勾选，请逐项判断。
  确认后可「打印 / 存 PDF」留存。
</div>
</div>

<footer class="bar">
  <div class="who">
    <label>确认人 <input id="operator" type="text" placeholder="姓名（必填）" style="width:120px"></label>
    <label>备注 <input id="memo" type="text" placeholder="可选" style="width:240px"></label>
  </div>
  <div class="sp"></div>
  <button class="btn ghost" id="btnPrint">打印 / 存 PDF</button>
  <button class="btn ghost" id="btnCopy">复制 JSON</button>
  <button class="btn" id="btnSubmit">提交并下载确认单</button>
</footer>

<div class="mask" id="mask">
  <div class="modal">
    <h3>确认提交</h3>
    <div class="row"><span>确认人</span><b id="mOp">—</b></div>
    <div class="row"><span>将写入主档</span><b id="mWrite">0 项</b></div>
    <div class="row"><span>保留现值（不写入）</span><b id="mKeep">0 项</b></div>
    <div class="row"><span>附件</span><b id="mAtt">0 个</b></div>
    <div class="row"><span>未处理的冲突 / 待核</span><b id="mLeft">0 项</b></div>
    <div class="warnline" id="mWarn"></div>
    <div class="acts">
      <button class="btn ghost" id="mCancel">返回修改</button>
      <button class="btn" id="mOk">确认并下载</button>
    </div>
  </div>
</div>
<div class="mask" id="doneMask">
  <div class="modal">
    <h3>确认单已下载</h3>
    <p style="font-size:13px;color:#374151;margin:0 0 12px">文件已在你的「下载」目录。接下来三步：</p>
    <ol style="font-size:13px;color:#374151;line-height:1.9;padding-left:20px;margin:0">
      <li>把该 JSON 放进客户目录：<br><code id="dPath" style="background:#F3F5F9;padding:3px 7px;border-radius:5px;font-size:12px;word-break:break-all;display:inline-block;margin-top:4px"></code></li>
      <li>回到对话，告诉助手「确认单已放好」</li>
      <li>助手会写入主档、归档确认单与附件，并生成备份</li>
    </ol>
    <div class="acts">
      <button class="btn ghost" id="dCopy">再复制一份 JSON</button>
      <button class="btn" id="dOk">知道了</button>
    </div>
  </div>
</div>
<div class="toast" id="toast"></div>

<script>
const ROWS = __DATA__;
const META = __META__;
const ORDER = {'冲突':0,'待核':1,'新增':2,'一致':3};
const LABEL = {0:'冲突',1:'待核',2:'新增',3:'一致'};
const DESC = {
 0:'主档现值与合同候选不同，需逐项判断保留哪一个',
 1:'合同未提取到该项，可手工填写或保持空白',
 2:'主档为空，将新增写入',
 3:'与主档现值相同，默认采用'
};
const ICON = {0:'▲',1:'■',2:'＋',3:'✓'};

/* ---------- 状态 ---------- */
const S = ROWS.map((r,i)=>({
  ...r, i,
  g: ORDER[r.status],
  use: (r.status==='新增'||r.status==='一致'),
  val: r.candidate || r.current || '',
  orig: r.candidate || r.current || ''
}));
let ATTS = [];
let filter = 'all';
let q = '';
const collapsed = {0:false,1:false,2:false,3:true};

/* ---------- 工具 ---------- */
const esc = s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const byId = id=>document.getElementById(id);
function toast(msg){const t=byId('toast');t.textContent=msg;t.classList.add('on');clearTimeout(t._h);t._h=setTimeout(()=>t.classList.remove('on'),2200);}
function fmtSize(n){return n>1048576?(n/1048576).toFixed(1)+' MB':(n/1024).toFixed(0)+' KB';}
function isDirty(r){return r.val!==r.orig;}

/* 字符级差异高亮 */
function diff(a,b){
  if(!a) return esc(b);
  if(!b) return '—';
  let p=0; while(p<a.length&&p<b.length&&a[p]===b[p]) p++;
  let s=0; while(s<a.length-p&&s<b.length-p&&a[a.length-1-s]===b[b.length-1-s]) s++;
  const pre=esc(b.slice(0,p)), mid=b.slice(p,b.length-s), post=esc(b.slice(b.length-s));
  return pre + (mid?`<mark>${esc(mid)}</mark>`:'') + post;
}

/* ---------- 渲染 ---------- */
function visible(r){
  if(filter==='dirty'){ if(!isDirty(r)) return false; }
  else if(filter!=='all' && r.status!==filter) return false;
  if(q){
    const hay=(r.field+' '+r.current+' '+r.val+' '+r.target).toLowerCase();
    if(!hay.includes(q.toLowerCase())) return false;
  }
  return true;
}

function render(){
  const buckets={0:[],1:[],2:[],3:[]};
  S.forEach(r=>buckets[r.g].push(r));
  ['c0','c1','c2','c3'].forEach((id,k)=>byId(id).textContent=buckets[k].length);

  let html='';
  let shown=0;
  [0,1,2,3].forEach(g=>{
    const list=buckets[g].filter(visible);
    if(!list.length) return;
    shown+=list.length;
    const col=collapsed[g]?' collapsed':'';
    html+=`<section class="grp${col}" data-g="${g}">
      <h2 data-toggle="${g}"><span class="tw">▾</span>
        <span class="badge b${g}">${ICON[g]} ${LABEL[g]}</span>
        <small>${DESC[g]}</small><span class="cnt">${list.length} 项</span></h2>
      <table class="t"><thead><tr>
        <th style="width:23%">字段</th><th style="width:23%">主档现值</th>
        <th style="width:26%">本次取值</th><th style="width:8%">采用</th><th>合同原文</th>
      </tr></thead><tbody>`;
    list.forEach(r=>{
      const dirty=isDirty(r);
      html+=`<tr class="${dirty?'dirty':''}" data-i="${r.i}">
        <td data-l="字段"><div class="fname">${esc(r.field)}</div><div class="fpath">${esc(r.target)}</div></td>
        <td data-l="主档现值"><div class="cur${r.current?'':' nil'}">${r.current?esc(r.current):'—'}</div></td>
        <td data-l="本次取值"><input class="inp${dirty?' changed':''}" data-in="${r.i}" type="text" value="${esc(r.val)}" placeholder="（空）">
          ${dirty?`<button class="undo" data-undo="${r.i}">恢复为「${esc(r.orig||'空')}」</button>`:''}</td>
        <td data-l="采用"><div class="use"><label><input type="checkbox" data-chk="${r.i}" ${r.use?'checked':''}>采用</label></div></td>
        <td data-l="合同原文">${r.evidence?`<div class="ev">${esc(r.evidence)}</div>`:'<span class="noev">合同未记载</span>'}</td>
      </tr>`;
    });
    html+=`</tbody></table></section>`;
  });
  if(!shown) html=`<div class="grp"><div class="empty">没有符合条件的条目${q?'（搜索：'+esc(q)+'）':''}</div></div>`;
  byId('app').innerHTML=html;

  const dirtyN=S.filter(isDirty).length;
  const conflictPending=S.filter(r=>r.g===0&&!r.use).length + S.filter(r=>r.g===1&&!r.val).length;
  byId('stat').innerHTML=`将写入 <b>${S.filter(r=>r.use).length}</b> 项　·　已改动 <b>${dirtyN}</b> 项　·　待判断 <b>${conflictPending}</b> 项`;
  window.__dirty = dirtyN>0 || ATTS.length>0;
}

/* ---------- 事件 ---------- */
document.addEventListener('input',e=>{
  const i=e.target.dataset.in;
  if(i!==undefined){
    S[+i].val=e.target.value;
    if(!S[+i].use && S[+i].val!==S[+i].orig){S[+i].use=true;toast('已自动勾选「采用」');}
    saveDraft();
    const tr=e.target.closest('tr');
    const dirty=isDirty(S[+i]);
    tr.classList.toggle('dirty',dirty);
    e.target.classList.toggle('changed',dirty);
    render();  // 保持统计与撤销按钮同步
    const again=document.querySelector(`input[data-in="${i}"]`);
    if(again){again.focus();again.setSelectionRange(again.value.length,again.value.length);}
  }
});
document.addEventListener('change',e=>{
  const i=e.target.dataset.chk;
  if(i!==undefined){ S[+i].use=e.target.checked; render(); saveDraft(); }
});
document.addEventListener('click',e=>{
  const t=e.target;
  if(t.dataset.toggle!==undefined){ collapsed[+t.dataset.toggle]=!collapsed[+t.dataset.toggle]; render(); }
  if(t.dataset.undo!==undefined){ const x=S[+t.dataset.undo]; x.val=x.orig; x.use=(x.orig!==null&&x.orig!==''&&x.orig!==undefined); render(); saveDraft(); }
});

byId('q').addEventListener('input',e=>{q=e.target.value.trim();render();});
const GRP_INDEX={'冲突':0,'待核':1,'新增':2,'一致':3};
document.querySelectorAll('.chip.click').forEach(ch=>{
  ch.addEventListener('click',()=>{
    filter=ch.dataset.f;
    [...byId('filters').children].forEach(x=>x.classList.toggle('on',x.dataset.f===filter));
    if(ch.dataset.f in GRP_INDEX) collapsed[GRP_INDEX[ch.dataset.f]]=false;  // 点筛选即展开该组
    render();
    const sec=document.querySelector('.grp[data-g="'+GRP_INDEX[ch.dataset.f]+'"]');
    (sec||byId('app')).scrollIntoView({behavior:'smooth',block:'start'});
  });
});
byId('filters').addEventListener('click',e=>{
  const b=e.target.closest('button'); if(!b)return;
  filter=b.dataset.f;
  [...byId('filters').children].forEach(x=>x.classList.toggle('on',x===b));
  render();
});
byId('selSame').onclick=()=>{S.filter(r=>r.g===3).forEach(r=>r.use=true);render();toast('已采纳全部一致项');};
byId('unsAll').onclick=()=>{S.forEach(r=>r.use=false);render();toast('已全部不采用');};
byId('resetAll').onclick=()=>{S.forEach(r=>{r.val=r.orig;r.use=(r.status==='新增'||r.status==='一致');});render();clearDraft();toast('已重置改动');};

/* ---------- 附件 ---------- */
const MAX_ONE=5*1024*1024, MAX_ALL=15*1024*1024;
function renderAtts(){
  const tb=byId('attTable').querySelector('tbody');
  const hint=byId('attHint');
  if(!ATTS.length){byId('attTable').style.display='none';hint.textContent='尚未添加附件';hint.className='hint';render();return;}
  byId('attTable').style.display='';
  tb.innerHTML=ATTS.map((a,i)=>`<tr>
    <td class="nm">${esc(a.name)}</td>
    <td class="sz">${fmtSize(a.size)}</td>
    <td><input type="text" data-note="${i}" value="${esc(a.note||'')}" placeholder="说明，如：营业执照副本"></td>
    <td class="ops"><a data-open="${i}">预览</a><span class="rm" data-rm="${i}">移除</span></td>
  </tr>`).join('');
  const total=ATTS.reduce((s,a)=>s+a.size,0);
  hint.textContent=`共 ${ATTS.length} 个附件，合计 ${fmtSize(total)}`;
  hint.className=total>MAX_ALL?'hint warn':'hint';
  if(total>MAX_ALL) hint.textContent+=`　已超出 15MB，请精简后再提交`;
  render();
}
function addFiles(list){
  [...list].forEach(f=>{
    if(f.size>MAX_ONE){toast(`「${f.name}」超过 5MB，已跳过`);return;}
    const r=new FileReader();
    r.onload=()=>{
      ATTS.push({name:f.name,size:f.size,mime:f.type||'application/octet-stream',note:'',
                 data_b64:String(r.result).split(',')[1]||'',
                 url:String(r.result)});
      renderAtts();
    };
    r.readAsDataURL(f);
  });
}
(function initDrop(){
  const dz=byId('drop'), inp=byId('files');
  dz.onclick=()=>inp.click();
  inp.onchange=e=>{addFiles(e.target.files);inp.value='';};
  ['dragenter','dragover'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('hover');}));
  ['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('hover');}));
  dz.addEventListener('drop',e=>{if(e.dataTransfer.files.length)addFiles(e.dataTransfer.files);});
})();
byId('attTable').addEventListener('input',e=>{
  const i=e.target.dataset.note;
  if(i!==undefined) ATTS[+i].note=e.target.value;
});
byId('attTable').addEventListener('click',e=>{
  const t=e.target;
  if(t.dataset.rm!==undefined){ATTS.splice(+t.dataset.rm,1);renderAtts();}
  if(t.dataset.open!==undefined){
    const a=ATTS[+t.dataset.open], w=window.open('','_blank');
    if(a.mime.startsWith('image/')) w.document.write(`<img src="${a.url}" style="max-width:100%">`);
    else if(a.mime==='application/pdf') w.location=a.url;
    else w.document.write(`<pre style="white-space:pre-wrap;font:13px/1.6 monospace">${esc(a.url.slice(0,200000))}</pre>`);
  }
});

/* ---------- 草稿 ---------- */
const LS='recon_draft:'+(META.customer||'')+':'+(META.contract||'')+':'+(META.page||'').slice(0,7);
let _lsT=null;
function saveDraft(){
  clearTimeout(_lsT);
  _lsT=setTimeout(()=>{
    try{
      localStorage.setItem(LS,JSON.stringify({
        v:S.map(r=>[r.target,r.val,r.use]),op:byId('operator').value,mm:byId('memo').value,
        att:ATTS.map(a=>({name:a.name,note:a.note})),ts:Date.now()
      }));
    }catch(e){}
  },400);
}
function clearDraft(){try{localStorage.removeItem(LS);}catch(e){}}
function loadDraft(){
  try{
    const raw=localStorage.getItem(LS); if(!raw) return;
    const d=JSON.parse(raw);
    if(!d||!d.v) return;
    let n=0;
    const byT={}; d.v.forEach(([tg,val,use])=>{ if(tg) byT[tg]=[val,use]; });
    S.forEach(r=>{ const d2=byT[r.target]; if(d2){ if(r.val!==d2[0]){r.val=d2[0];n++;} r.use=!!d2[1]; } });
    if(d.op) byId('operator').value=d.op;
    if(d.mm) byId('memo').value=d.mm;
    if(n) { render(); toast(`已恢复上次未提交的修改（${n} 项，含采用状态）`); }
  }catch(e){}
}

/* ---------- 提交 ---------- */
function buildPayload(){
  const op=byId('operator').value.trim();
  const fields=[], skipped=[];
  S.forEach(r=>{
    if(r.use) fields.push({field:r.field,target:r.target,value:r.val,from:r.status});
    else skipped.push({field:r.field,target:r.target,kept:r.current});
  });
  return {
    schema:'contract-confirm/v1',
    meta:{customer:META.customer,contract:META.contract,operator:op,
          memo:byId('memo').value.trim(),confirmed_at:new Date().toISOString(),
          source_file:META.source_file,page:META.page},
    fields, skipped,
    attachments:ATTS.map(a=>({name:a.name,size:a.size,mime:a.mime,note:a.note,data_b64:a.data_b64}))
  };
}
function precheck(){
  const op=byId('operator').value.trim();
  if(!op){toast('请先填写确认人姓名');byId('operator').focus();return false;}
  const total=ATTS.reduce((s,a)=>s+a.size,0);
  if(total>MAX_ALL){toast('附件合计超过 15MB，请精简后重试');return false;}
  return true;
}
byId('btnSubmit').onclick=()=>{
  if(!precheck()) return;
  const p=buildPayload();
  byId('mOp').textContent=p.meta.operator;
  byId('mWrite').textContent=p.fields.length+' 项';
  byId('mKeep').textContent=p.skipped.length+' 项';
  byId('mAtt').textContent=ATTS.length+' 个';
  const left=S.filter(r=>r.g===0&&!r.use).length+S.filter(r=>r.g===1&&!r.val).length;
  byId('mLeft').textContent=left+' 项';
  const w=byId('mWarn');
  if(left>0){w.style.display='block';w.textContent=`仍有 ${left} 项冲突未采纳或待核为空，将按现值保留。`;}
  else w.style.display='none';
  byId('mask').classList.add('on');
};
byId('mCancel').onclick=()=>byId('mask').classList.remove('on');
byId('mOk').onclick=()=>{
  byId('mask').classList.remove('on');
  const p=buildPayload();
  const blob=new Blob([JSON.stringify(p,null,2)],{type:'application/json'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download=`合同确认单_${META.customer||'客户'}_${new Date().toISOString().slice(0,10)}.json`;
  document.body.appendChild(a);a.click();a.remove();
  window.__dirty=false;
  clearDraft();
  byId('dPath').textContent=META.target_dir||'（客户目录）';
  byId('doneMask').classList.add('on');
};
function copyJSON(){
  if(!precheck()) return;
  const txt=JSON.stringify(buildPayload(),null,2);
  (navigator.clipboard?navigator.clipboard.writeText(txt):Promise.reject())
    .then(()=>toast('JSON 已复制到剪贴板'))
    .catch(()=>{
      const ta=document.createElement('textarea');ta.value=txt;document.body.appendChild(ta);
      ta.select();document.execCommand('copy');ta.remove();toast('JSON 已复制到剪贴板');
    });
}
byId('btnCopy').onclick=copyJSON;
byId('dOk').onclick=()=>byId('doneMask').classList.remove('on');
byId('dCopy').onclick=()=>{copyJSON();};
byId('btnPrint').onclick=()=>window.print();

document.addEventListener('keydown',e=>{
  if(e.key==='Escape') byId('mask').classList.remove('on');
  if((e.metaKey||e.ctrlKey)&&e.key==='Enter'){e.preventDefault();byId('btnSubmit').click();}
});
window.addEventListener('beforeunload',e=>{
  if(window.__dirty){e.preventDefault();e.returnValue='有未提交的修改，确认离开？';}
});

render();renderAtts();loadDraft();
</script></body></html>"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("candidate")
    ap.add_argument("master")
    ap.add_argument("-o", "--out")
    a = ap.parse_args()

    cand = json.load(open(a.candidate, encoding="utf-8"))
    master = load_master_values(a.master)
    rows = build_rows(cand, master)

    meta = {
        "customer": master.get("_客户名称") or "",
        "contract": master.get("_合同编号") or "",
        "source_file": os.path.basename(cand.get("file", "")),
        "page": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "target_dir": os.path.dirname(os.path.abspath(a.master)),
    }
    subtitle = (f"客户：{meta['customer'] or '—'}　·　合同编号：{meta['contract'] or '—'}　·　"
                f"来源文件：{meta['source_file']}　·　生成于 {meta['page']}")

    def js_json(o):
        # 防止字段内容含 </script> 提前闭合脚本
        return json.dumps(o, ensure_ascii=False).replace("</", "<\\/")

    page = (TPL.replace("__DATA__", js_json(rows))
               .replace("__META__", js_json(meta))
               .replace("__TITLE__", html.escape(meta["customer"] or "合同信息"))
               .replace("__SUBTITLE__", html.escape(subtitle))
               .replace("__SRC__", html.escape(meta["source_file"] or "—")))

    out = a.out or os.path.join(os.path.dirname(os.path.abspath(a.master)),
                                f"_确认页_{datetime.date.today():%Y%m%d}.html")
    open(out, "w", encoding="utf-8").write(page)
    cnt = {}
    for r in rows:
        cnt[r["status"]] = cnt.get(r["status"], 0) + 1
    print("已生成:", out)
    print("分组:", cnt)


if __name__ == "__main__":
    main()
