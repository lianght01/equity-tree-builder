#!/usr/bin/env python3
"""
股权架构树桌面应用 — pywebview 原生窗口
导入Excel数据 → 软件内直接渲染D3股权树 → 节点可拖拽（父节点带动子节点）
"""
import webview
import json, os, sys, re, threading

TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates')
if hasattr(sys, '_MEIPASS'):
    TEMPLATE_DIR = os.path.join(sys._MEIPASS, 'templates')

# ── 数据解析（复用 equity_tree_gui.py 的逻辑） ──

def find_header(col):
    h = str(col).strip().lower()
    aliases = {
        'parent': ['上级企业名称', '上级企业', '上级公司', '母公司', '股东名称', '母公司名称', '股东', 'parent_name', 'parent', '控股方'],
        'name': ['企业名称', '公司名称', '公司名', 'entity_name', 'company'],
        'ratio': ['持股比例', '投资比例', '股权比例', 'ratio', 'invest_ratio'],
        'status': ['经营状态', '登记状态', '状态', 'status', 'reg_status'],
        'level': ['层级', '级别', 'level', 'depth'],
        'biz_open': ['开户日期', '开户时间', '开户日', 'open_date', 'openDate'],
        'biz_deposit_avg': ['日均存款', '日均', 'deposit_avg', 'depositAvg'],
        'biz_deposit_spot': ['时点存款', '时点', 'deposit_spot', 'depositSpot'],
        'biz_loan': ['授信余额', '贷款余额', '授信', 'loan_balance', 'loanBalance'],
        'biz_income': ['年化收入', '净收入', '净经营收入', '年收入', 'net_income', 'netIncome12m'],
        'biz_scale': ['规模', '企业规模', 'scale'],
        'biz_manager': ['客户经理', '经理', '管户', 'manager'],
    }
    best_key, best_score = None, 0
    for key, names in aliases.items():
        for n in names:
            nl = n.lower()
            if h == nl:
                return key
            if nl in h:
                score = len(nl)
                if score > best_score:
                    best_key, best_score = key, score
    return best_key

def load_xlsx(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    sheets = {}
    for name in wb.sheetnames:
        ws = wb[name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append([str(v).strip() if v is not None else '' for v in row])
        sheets[name] = rows
    wb.close()
    return sheets

def parse_sheet(rows):
    if not rows or len(rows) < 2:
        return []
    header_row_idx = 0
    for idx, row in enumerate(rows):
        cell_text = ''.join(str(c) for c in row if c).lower()
        if any(kw in cell_text for kw in ['企业名称','公司名称','上级企业','名称','company','entity','股东']):
            header_row_idx = idx
            break
    header = rows[header_row_idx]
    col_idx = {}
    for i, h in enumerate(header):
        key = find_header(h)
        if key:
            col_idx[key] = i
    if 'name' not in col_idx:
        non_empty = [i for i, h in enumerate(header) if h]
        if len(non_empty) >= 1:
            col_idx['name'] = non_empty[0]
        if len(non_empty) >= 2:
            col_idx['parent'] = non_empty[1]
    result = []
    for row in rows[header_row_idx + 1:]:
        if not any(v.strip() for v in row if isinstance(v, str)):
            continue
        rec = {}
        for key, idx in col_idx.items():
            if idx < len(row):
                rec[key] = str(row[idx]).strip()
        if rec.get('name'):
            result.append(rec)
    return result

def build_tree(rows, root_name=None, mode='control'):
    all_rels = []
    node_info = {}
    for r in rows:
        name = r.get('name', '').strip()
        parent = r.get('parent', '').strip()
        if not name:
            continue
        if name not in node_info:
            node_info[name] = {'status': r.get('status', ''), 'ratio': ''}
        if parent:
            all_rels.append((name, parent, r.get('ratio', '')))
    if mode == 'invest' and root_name:
        return _build_invest_tree(root_name, all_rels, node_info)
    else:
        return _build_control_tree(all_rels, node_info, root_name)

def _build_control_tree(all_rels, node_info, root_name=None):
    best_parent = {}
    for child, parent, ratio in all_rels:
        if child not in best_parent:
            best_parent[child] = (parent, ratio)
        else:
            try:
                cur = float(best_parent[child][1]) if best_parent[child][1] else 0
                new = float(ratio) if ratio else 0
                if new > cur:
                    best_parent[child] = (parent, ratio)
            except ValueError:
                pass
    children_map = {}
    for child, (parent, ratio) in best_parent.items():
        children_map.setdefault(parent, []).append((child, ratio))
    for p in children_map:
        children_map[p].sort(key=lambda x: float(x[1]) if x[1] else 0, reverse=True)
    all_ps = set(p for p, _ in best_parent.values())
    all_cs = set(best_parent.keys())
    candidates = list(all_ps - all_cs)
    known = [c for c in candidates if c in node_info]
    if root_name:
        roots = [root_name]
    elif known:
        known.sort(key=lambda n: len(children_map.get(n, [])), reverse=True)
        roots = [known[0]]
    else:
        roots = candidates[:1] if candidates else (list(children_map.keys())[:1] if children_map else [])
    if not roots:
        return {'name': all_rels[0][0] if all_rels else '', 'ratio': '', 'status': '', 'children': []}
    return _build_subtree(roots[0], children_map, node_info)

def _build_invest_tree(root_name, all_rels, node_info):
    root_children = set()
    for child, parent, ratio in all_rels:
        if parent == root_name:
            root_children.add(child)
    best_parent = {}
    for child, parent, ratio in all_rels:
        if child not in best_parent:
            best_parent[child] = (parent, ratio)
        else:
            try:
                cur = float(best_parent[child][1]) if best_parent[child][1] else 0
                new = float(ratio) if ratio else 0
                if new > cur:
                    best_parent[child] = (parent, ratio)
            except ValueError:
                pass
    children_map = {}
    for child, (parent, ratio) in best_parent.items():
        children_map.setdefault(parent, []).append((child, ratio))
    for p in children_map:
        children_map[p].sort(key=lambda x: float(x[1]) if x[1] else 0, reverse=True)
    root = {'name': root_name, 'ratio': '', 'status': '个人' if not node_info.get(root_name, {}).get('status') else node_info[root_name]['status'], 'children': []}
    visited = set([root_name])
    def build_subtree(name):
        if name in visited:
            return None
        visited.add(name)
        info = node_info.get(name, {})
        node = {'name': name, 'ratio': '', 'status': info.get('status', ''), 'children': []}
        for child_name, ratio in children_map.get(name, []):
            cn = build_subtree(child_name)
            if cn:
                if ratio:
                    cn['ratio'] = ratio
                node['children'].append(cn)
        visited.discard(name)
        return node
    for child in root_children:
        cn = build_subtree(child)
        if cn:
            root['children'].append(cn)
    return root

def _build_subtree(name, children_map, node_info, visited=None):
    if visited is None:
        visited = set()
    if name in visited:
        return {'name': name, 'ratio': '', 'status': '(循环)', 'children': []}
    visited.add(name)
    info = node_info.get(name, {})
    node = {'name': name, 'ratio': info.get('ratio', ''), 'status': info.get('status', ''), 'children': []}
    for child_name, ratio in children_map.get(name, []):
        cn = _build_subtree(child_name, children_map, node_info, visited)
        if ratio:
            cn['ratio'] = ratio
        node['children'].append(cn)
    visited.discard(name)
    return node

def count_nodes(n):
    c = 1
    for ch in n.get('children', []):
        c += count_nodes(ch)
    return c

def load_biz_data(biz_rows, tree):
    if not biz_rows or len(biz_rows) < 2:
        return tree, 0
    col_map = {}
    for i, h in enumerate(biz_rows[0]):
        kh = str(h).strip().lower()
        for k, aliases in {
            'name': ['企业名称', '公司名称', '名称'],
            'openDate': ['开户日期', '开户时间', '开户日'],
            'depositAvg': ['日均存款', '日均'],
            'depositSpot': ['时点存款', '时点'],
            'loanBalance': ['授信余额', '贷款余额'],
            'netIncome12m': ['年化收入', '净收入'],
            'payrollCount': ['代发个人客户', '代发人数', '代发户数'],
            'payrollAmount': ['发薪量', '代发金额', '代发总额'],
            'pensionCardCount': ['养老金发卡', '养老金卡', '养老发卡'],
            'scale': ['规模'],
            'manager': ['客户经理', '管户'],
        }.items():
            if any(a.lower() in kh for a in aliases):
                col_map[k] = i
                break
    if 'name' not in col_map or 'openDate' not in col_map:
        return tree, 0
    biz_records = []
    for row in biz_rows[1:]:
        if not row or not row[col_map['name']]:
            continue
        biz = {}
        for k, idx in col_map.items():
            if idx < len(row) and row[idx]:
                v = row[idx]
                if k in ['depositAvg','depositSpot','loanBalance','netIncome12m','payrollCount','payrollAmount','pensionCardCount']:
                    try: v = float(str(v).replace(',',''))
                    except: v = 0
                biz[k] = v
        if biz.get('name'):
            biz_records.append(biz)
    lookup = {}
    for b in biz_records:
        name = str(b.get('name','')).strip()
        if name:
            lookup[name] = b
    matched = 0
    def match(n):
        nonlocal matched
        name = n.get('name','')
        if name in lookup:
            b = lookup[name]
            n['_biz'] = {
                'hasAccount': True,
                'openDate': str(b.get('openDate','')),
                'depositAvg': float(b.get('depositAvg',0)),
                'depositSpot': float(b.get('depositSpot',0)),
                'depositTotalAvg': float(b.get('depositAvg',0)),
                'loanBalance': float(b.get('loanBalance',0)),
                'interestIncome': 0,
                'netIncome': float(b.get('netIncome12m',0))/4,
                'netIncome12m': float(b.get('netIncome12m',0)),
                'payrollCount': int(b.get('payrollCount',0)),
                'payrollAmount': float(b.get('payrollAmount',0)),
                'pensionCardCount': int(b.get('pensionCardCount',0)),
                'scale': str(b.get('scale','')),
                'manager': str(b.get('manager','')),
            }
            matched += 1
        for c in n.get('children',[]):
            match(c)
    match(tree)
    return tree, matched

def clean_tree(tree):
    if 'children' in tree:
        kept = []
        for c in tree['children']:
            r = c.get('ratio','')
            if not r or str(r).strip()=='':
                if not c.get('_biz') and not c.get('children'):
                    continue
            kept.append(c)
        tree['children'] = kept
    return tree

def clean_status(tree):
    bad_keywords = ['注销','吊销','已告解散','清算','停业','迁出','撤销','关闭','非正常户','破产','除名','取消']
    removed = []
    def walk(node):
        if not node.get('children'):
            return
        kept = []
        for c in node['children']:
            st = c.get('status', '').strip()
            is_bad = False
            if st:
                if st == '存续' or st == '在营（开业）':
                    is_bad = False
                else:
                    for kw in bad_keywords:
                        if kw in st:
                            is_bad = True
                            break
            if is_bad:
                removed.append(c['name'])
            else:
                walk(c)
                kept.append(c)
        node['children'] = kept
    walk(tree)
    return removed


# ── HTML 生成 ──

def generate_full_html(tree, title):
    """生成完整HTML（head + data + tail）"""
    tree_json = json.dumps(tree, ensure_ascii=False, separators=(',', ':'))
    total = count_nodes(tree)

    head_path = os.path.join(TEMPLATE_DIR, 'head.html')
    tail_path = os.path.join(TEMPLATE_DIR, 'tail.html')

    if os.path.exists(head_path) and os.path.exists(tail_path):
        with open(head_path,'r') as f: head = f.read()
        with open(tail_path,'r') as f: tail_ = f.read()
        head = re.sub(r'<title>[^<]+</title>', f'<title>{title}</title>', head)
        head = re.sub(r'鞍钢集团控股股权关系树', title, head)
        head = head.replace('>1372节点<', f'>{total}节点<')
        head = head.replace('>1372<', f'>{total}<')
        html = head + 'var TD=' + tree_json + ';' + tail_
    else:
        html = _minimal_html(tree_json, title, total)
    return html

def _minimal_html(tree_json, title, total):
    return f'''<!DOCTYPE html>
<html lang="zh"><head><meta charset="UTF-8"><title>{title}</title>
<style>
@media print{{@page{{size:landscape;margin:15mm 10mm}}}}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:-apple-system,"Microsoft YaHei",sans-serif;background:#fff;overflow:hidden;height:100vh}}
#bar{{position:fixed;top:0;left:0;right:0;z-index:10;background:rgba(255,255,255,.95);padding:8px 14px;display:flex;gap:8px;align-items:center;border-bottom:1px solid #d0d7de}}
#bar button{{background:#f6f8fa;border:1px solid #d0d7de;padding:6px 12px;border-radius:5px;cursor:pointer}}
svg{{position:fixed;top:46px;left:0;width:100%;height:calc(100vh-46px)}}
.node rect{{rx:4;ry:4}}.node text{{font-size:13px;fill:#24292f;text-anchor:middle;dominant-baseline:central;pointer-events:none}}
.link{{fill:none;stroke:#d0d7de;stroke-width:2px}}
</style></head><body>
<div id="bar"><span style="color:#0969da;font-weight:bold">{title} · {total}节点</span>
<button onclick="d3.select(this).transition().duration(300).call(zoom.scaleBy,1.3)">🔍+</button>
<button onclick="d3.select(this).transition().duration(300).call(zoom.scaleBy,0.7)">🔍-</button>
<button onclick="root.descendants().forEach(d=>{{if(d._children){{d.children=d._children;d._children=null}}}});upd()">展开</button>
<button onclick="root.descendants().forEach(d=>{{if(d.depth>=1&&d.children){{d._children=d.children;d.children=null}}}});upd()">折叠</button>
</div>
<svg id="svg"></svg>
<script src="https://d3js.org/d3.v7.min.js"></script>
<script>
var DC=["#f0883e","#3fb950","#58a6ff","#d2a850","#bc8cff","#f778ba"];
var TD={tree_json};
var svgEl=document.getElementById("svg"),svg=d3.select(svgEl),g=svg.append("g");
var tl=d3.tree().size([2400,50000]).nodeSize([180,160]).separation(function(a,b){{return a.parent===b.parent?1.3:2}});
var root=d3.hierarchy(TD);tl(root);
root.descendants().forEach(function(d){{if(d.depth>=1){{d._children=d.children;d.children=null}}}});
var zoom=d3.zoom().scaleExtent([0.1,5]).on("zoom",function(e){{g.attr("transform",e.transform)}});
svg.call(zoom);var is=0.9;
svg.call(zoom.transform,d3.zoomIdentity.translate(svgEl.clientWidth/2-root.x*is,60).scale(is));
function upd(){{
  tl(root);var ns=root.descendants(),ls=root.links();
  g.selectAll(".link").data(ls,d=>d.target.data.name).join("path").attr("class","link")
    .attr("d",d=>"M"+d.source.x+","+(d.source.y+22)+"C"+d.source.x+","+((d.source.y+d.target.y)/2)+" "+d.target.x+","+((d.source.y+d.target.y)/2)+" "+d.target.x+","+(d.target.y-22));
  g.selectAll(".node").data(ns,d=>d.data.name).join("g").attr("class","node")
    .attr("transform",d=>"translate("+d.x+","+d.y+")").style("cursor","grab")
    .on("click",function(ev,d){{ev.stopPropagation();if(d.children){{d._children=d.children;d.children=null}}else if(d._children){{d.children=d._children;d._children=null}}upd()}})
    .call(d3.drag().on("start",function(ev,d){{ev.sourceEvent.stopPropagation();d3.select(this).style("cursor","grabbing");d._sx=d.x;d._sy=d.y}})
      .on("drag",function(ev,d){{var dx=ev.x-d._sx,dy=ev.y-d._sy;d._sx=ev.x;d._sy=ev.y;d.x=ev.x;d.y=ev.y;d3.select(this).attr("transform","translate("+d.x+","+d.y+")");d.descendants().forEach(function(c){{if(c!==d){{c.x+=dx;c.y+=dy;c._sx=(c._sx||c.x)+dx;c._sy=(c._sy||c.y)+dy}}}});updLinks()}})
      .on("end",function(ev,d){{d3.select(this).style("cursor","grab")}}))
    .each(function(d){{
      var me=d3.select(this),len=d.data.name.length,w=Math.max(len*14+16,80);
      me.selectAll("*").remove();
      me.append("rect").attr("x",-w/2).attr("y",-18).attr("width",w).attr("height",36)
        .attr("fill",d.data._biz?"#2b7a78":DC[Math.min(d.depth,5)])
        .attr("stroke",d3.color(DC[Math.min(d.depth,5)]).darker(0.3)).attr("stroke-width",1.5);
      me.append("text").attr("y",0).text(len>18?d.data.name.slice(0,16)+"\\u2026":d.data.name);
      var hk=d.children||d._children;if(hk)me.append("text").attr("x",w/2+8).attr("y",0).attr("font-size","10px").text(d.children?"\\u25bc":"\\u25b6");
    }})
}}upd();
function updLinks(){{var ls=root.links();g.selectAll(".link").data(ls,d=>d.target.data.name).join("path").attr("class","link").attr("d",d=>"M"+d.source.x+","+(d.source.y+22)+"C"+d.source.x+","+((d.source.y+d.target.y)/2)+" "+d.target.x+","+((d.source.y+d.target.y)/2)+" "+d.target.x+","+(d.target.y-22));}}
</script></body></html>'''


# ── API（JS 可调用） ──

class Api:
    def __init__(self, app):
        self.app = app

    def open_file_dialog(self):
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False,
            file_types=('Excel文件 (*.xlsx;*.xls)', '所有文件 (*.*)')
        )
        if result:
            return result[0]
        return ''

    def open_biz_dialog(self):
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False,
            file_types=('Excel文件 (*.xlsx;*.xls)', '所有文件 (*.*)')
        )
        if result:
            return result[0]
        return ''

    def load_data(self, data_path, biz_path, root_name, mode, clean_status_flag):
        """在后台线程加载数据，完成后通过JS回调更新页面"""
        def _load():
            try:
                webview.windows[0].evaluate_js('document.getElementById("statusBar").textContent="⏳ 正在加载数据...";')

                sheets = load_xlsx(data_path)
                data_rows = None
                for sn in list(sheets.keys()):
                    if len(sheets[sn]) >= 3:
                        data_rows = sheets[sn]
                        break
                if not data_rows:
                    webview.windows[0].evaluate_js('alert("未找到有效数据");document.getElementById("statusBar").textContent="❌ 加载失败";')
                    return

                rows = parse_sheet(data_rows)
                if not rows:
                    webview.windows[0].evaluate_js('alert("未能解析数据");document.getElementById("statusBar").textContent="❌ 解析失败";')
                    return

                tree = build_tree(rows, root_name, mode)
                tree = clean_tree(tree)

                if clean_status_flag == 'true' or clean_status_flag == True:
                    removed = clean_status(tree)
                    if removed:
                        print(f"已剔除 {len(removed)} 个异常状态节点")

                biz_count = 0
                if biz_path:
                    biz_sheets = load_xlsx(biz_path)
                    for sn in biz_sheets:
                        if len(biz_sheets[sn]) >= 2:
                            tree, biz_count = load_biz_data(biz_sheets[sn], tree)
                            break

                total = count_nodes(tree)
                title = f"股权关系树 ({total}节点)"

                # Generate HTML and load into webview
                html = generate_full_html(tree, title)
                webview.windows[0].load_html(html)

            except Exception as e:
                import traceback
                traceback.print_exc()
                err = str(e).replace('"', "'").replace('\n', ' ')
                webview.windows[0].evaluate_js(f'alert("加载失败: {err}");document.getElementById("statusBar").textContent="❌ 加载失败";')

        threading.Thread(target=_load, daemon=True).start()
        return 'loading'


# ── 启动页 HTML（含完整工具栏） ──

def landing_html():
    return '''<!DOCTYPE html>
<html lang="zh">
<head><meta charset="UTF-8"><title>股权架构树桌面版</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,"Microsoft YaHei",sans-serif;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);height:100vh;display:flex;align-items:center;justify-content:center;color:#fff}
.card{background:rgba(255,255,255,.95);border-radius:16px;padding:40px 50px;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,.3);max-width:520px;color:#24292f}
.card h1{font-size:28px;margin-bottom:8px;color:#0969da}
.card p{color:#656d76;font-size:14px;margin-bottom:24px;line-height:1.6}
.btn{display:inline-block;background:#0969da;color:#fff;border:none;padding:12px 32px;border-radius:8px;font-size:16px;cursor:pointer;margin:6px;transition:all .2s}
.btn:hover{background:#0550ae;transform:translateY(-1px)}
.btn.green{background:#2b7a78}
.btn.green:hover{background:#1e5f5d}
.info{font-size:12px;color:#8b949e;margin-top:20px}
#statusBar{margin-top:16px;font-size:13px;color:#656d76;min-height:20px}
</style>
</head>
<body>
<div class="card">
<h1>🏢 股权架构树桌面版</h1>
<p>导入Excel股东数据 → 自动生成交互式股权架构树<br>节点可拖拽 · 父节点拖动带动子节点 · 搜索 · 筛选 · 打印</p>
<button class="btn" onclick="loadData()">📂 选择数据文件</button>
<div id="statusBar"></div>
<div class="info">支持 .xlsx / .xls 格式 · 需含企业名称和上级企业列</div>
</div>
<script>
var dataPath = '', bizPath = '';

async function loadData() {
  var path = await pywebview.api.open_file_dialog();
  if (!path) return;
  dataPath = path;
  document.getElementById('statusBar').textContent = '已选择: ' + path.split('/').pop();

  var hasBiz = confirm('是否加载交行客户数据？（取消可跳过）');
  if (hasBiz) {
    bizPath = await pywebview.api.open_biz_dialog();
    if (bizPath) document.getElementById('statusBar').textContent += ' + 客户数据';
  }

  var rootName = prompt('指定根节点企业名称（留空自动识别）:', '');
  var mode = confirm('确定=投资穿透树，取消=控制权树') ? 'invest' : 'control';
  var cleanStatus = confirm('是否剔除已注销/吊销的异常状态节点？') ? true : false;

  document.getElementById('statusBar').textContent = '⏳ 正在加载数据...';
  pywebview.api.load_data(dataPath, bizPath||'', rootName||'', mode, cleanStatus);
}
</script>
</body>
</html>'''


# ── 启动 ──

if __name__ == '__main__':
    window = webview.create_window(
        title='股权架构树桌面版',
        width=1400, height=900,
        resizable=True,
        js_api=Api(None),
        html=landing_html(),
    )
    webview.start(debug=True, http_server=True)
