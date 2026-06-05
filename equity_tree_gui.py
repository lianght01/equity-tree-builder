#!/usr/bin/env python3
"""
股权架构树生成器 - GUI版
从Excel数据文件生成交互式股权架构树HTML
打包命令: pyinstaller --onefile --windowed --name "股权树生成器" equity_tree_gui.py
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json, os, sys, re, threading

# ── 核心逻辑（从 equity_tree_builder.py 精简整合） ──

TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates')
# For PyInstaller bundled mode
if hasattr(sys, '_MEIPASS'):
    TEMPLATE_DIR = os.path.join(sys._MEIPASS, 'templates')
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
    # Score each match: prefer longer and more specific
    best_key, best_score = None, 0
    for key, names in aliases.items():
        for n in names:
            nl = n.lower()
            # Exact match (highest priority)
            if h == nl:
                return key  # instant match
            # Alias is a substring of header (e.g. '名称' in '企业名称')
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
    # Find header row: first row where at least one cell matches a known column name
    header_row_idx = 0
    for idx, row in enumerate(rows):
        cell_text = ''.join(str(c) for c in row if c).lower()
        if any(kw in cell_text for kw in ['企业名称','公司名称','上级企业','名称','company','entity','股东']):
            header_row_idx = idx
            break
    
    header = rows[header_row_idx]
    # Log what headers we found
    print(f"  Headers: {[h for h in header if h]}")
    
    col_idx = {}
    for i, h in enumerate(header):
        key = find_header(h)
        if key:
            col_idx[key] = i
    print(f"  Mapped: {col_idx}")
    
    if 'name' not in col_idx:
        print(f"  ⚠️ 未找到'企业名称'列，尝试自动识别...")
        # Fallback: first non-empty column is name, second is parent
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
    
    print(f"  Parsed: {len(result)} records")
    return result

def build_tree(rows, root_name=None, mode='control'):
    """
    从股东关系表构建树。
    
    mode='control': 控制权树 - 每家子公司取持股比例最高的股东作为父节点
    mode='invest':  投资穿透树 - 从根节点出发，包含该节点投资的所有企业，
                     对每个被投企业再取最大股东穿透其下级
    """
    # First pass: build all relationships
    all_rels = []  # (child, parent, ratio)
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
    """控制权树: 取持股比例最高的股东作为父节点"""
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
    
    # Find root
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
    """投资穿透树: 从根节点出发→其投资的企业→这些企业的最大股东下级"""
    # Step 1: Find ALL companies that root_name invests in (any ratio)
    root_children = set()
    for child, parent, ratio in all_rels:
        if parent == root_name:
            root_children.add(child)
    
    # Step 2: For each of those companies, trace their control chain (max parent)
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
    
    # Step 3: Build children_map ONLY for paths descending from root_children
    children_map = {}
    for child, (parent, ratio) in best_parent.items():
        children_map.setdefault(parent, []).append((child, ratio))
    for p in children_map:
        children_map[p].sort(key=lambda x: float(x[1]) if x[1] else 0, reverse=True)
    
    # Step 4: Build the tree
    # Root node
    root = {
        'name': root_name,
        'ratio': '',
        'status': '个人' if not node_info.get(root_name, {}).get('status') else node_info[root_name]['status'],
        'children': []
    }
    
    visited = set([root_name])
    
    def build_subtree(name):
        if name in visited:
            return None
        visited.add(name)
        info = node_info.get(name, {})
        node = {
            'name': name,
            'ratio': '',
            'status': info.get('status', ''),
            'children': []
        }
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
    node = {
        'name': name,
        'ratio': info.get('ratio', ''),
        'status': info.get('status', ''),
        'children': []
    }
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
                if k in ['depositAvg','depositSpot','loanBalance','netIncome12m']:
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
                'scale': str(b.get('scale','')),
                'manager': str(b.get('manager','')),
            }
            matched += 1
        for c in n.get('children',[]):
            match(c)
    match(tree)
    return tree, matched
def clean_tree(tree):
    """基本清洗：去除空ratio的一级节点（疑似误入）"""
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
    """剔除经营状态异常节点，只保留状态为'存续'或'在营（开业）'的企业"""
    bad_keywords = ['注销','吊销','已告解散','清算','停业','迁出','撤销','关闭','非正常户',
                    '破产','除名','迁出','取消']
    removed = []
    
    def walk(node):
        if not node.get('children'):
            return
        kept = []
        for c in node['children']:
            st = c.get('status', '').strip()
            # Also check node's own status from record
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
                # Walk into its children just to log them too
                if c.get('children'):
                    _count_bad(c)
            else:
                walk(c)
                kept.append(c)
        node['children'] = kept
    
    def _count_bad(node):
        for c in node.get('children', []):
            removed.append('  ↳ ' + c['name'])
            _count_bad(c)
    
    walk(tree)
    return removed

def generate_html(tree, title, output_path, log_func=print):
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
        log_func("警告：使用内置简约模板（需联网加载D3）")
        html = _generate_minimal_html(tree_json, title, total)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    return total

def _generate_minimal_html(tree_json, title, total):
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
    .attr("transform",d=>"translate("+d.x+","+d.y+")").style("cursor","pointer")
    .on("click",function(ev,d){{ev.stopPropagation();if(d.children){{d._children=d.children;d.children=null}}else if(d._children){{d.children=d._children;d._children=null}}upd()}})
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
</script></body></html>'''

# ── GUI应用 ──

class EquityTreeApp:
    def __init__(self, root):
        self.root = root
        root.title("股权架构树生成器")
        root.geometry("700x560")
        root.resizable(False, False)
        root.configure(bg="#f0f2f5")
        
        # Try icon
        try:
            root.iconphoto(True, tk.PhotoImage(file=os.path.join(TEMPLATE_DIR, '..', 'icon.png')))
        except:
            pass
        
        # Colors
        self.C = {
            'bg': '#f0f2f5',
            'card': '#ffffff',
            'primary': '#1a73e8',
            'primary_hover': '#1557b0',
            'text': '#202124',
            'text_sec': '#5f6368',
            'border': '#e0e0e0',
            'success': '#1e8e3e',
            'btn_bg': '#f8f9fa',
        }
        
        # Variables
        self.data_path = tk.StringVar()
        self.biz_path = tk.StringVar()
        self.title_var = tk.StringVar(value="股权关系树")
        self.root_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="control")
        self.clean_status_var = tk.BooleanVar(value=True)
        self.output_dir = tk.StringVar(value=os.path.expanduser("~/Desktop"))
        self._last_output = None
        
        self._build_ui()
    
    def _make_card(self, parent):
        """Create a card-style container frame"""
        card = tk.Frame(parent, bg='#ffffff', bd=0, highlightthickness=0,
                       relief=tk.FLAT)
        return card
    
    def _make_row(self, parent, label_text, tooltip=None):
        """Create a form row with label on left, content on right"""
        row = tk.Frame(parent, bg='#ffffff')
        row.pack(fill=tk.X, pady=3)
        lbl = tk.Label(row, text=label_text, width=16, anchor=tk.E,
                      font=("Microsoft YaHei", 10), bg='#ffffff', fg=self.C['text'])
        lbl.pack(side=tk.LEFT, padx=(0, 8))
        content = tk.Frame(row, bg='#ffffff')
        content.pack(side=tk.LEFT, fill=tk.X, expand=True)
        return content, row
    
    def _make_btn(self, parent, text, cmd, style='outline', width=None):
        """Styled button with outline or primary style"""
        if style == 'primary':
            btn = tk.Button(parent, text=text, command=cmd,
                          bg=self.C['primary'], fg='white',
                          font=("Microsoft YaHei", 10, "bold"),
                          padx=16, pady=4, cursor="hand2",
                          border=0, activebackground=self.C['primary_hover'],
                          activeforeground='white', relief=tk.FLAT)
        elif style == 'outline':
            btn = tk.Button(parent, text=text, command=cmd,
                          bg='#ffffff', fg=self.C['primary'],
                          font=("Microsoft YaHei", 10),
                          padx=12, pady=4, cursor="hand2",
                          border=0, activebackground='#e8f0fe',
                          relief=tk.FLAT, highlightthickness=1,
                          highlightcolor=self.C['border'],
                          highlightbackground=self.C['border'])
        else:
            btn = tk.Button(parent, text=text, command=cmd,
                          bg=self.C['btn_bg'], fg=self.C['text'],
                          font=("Microsoft YaHei", 10),
                          padx=12, pady=4, cursor="hand2",
                          border=0, activebackground='#e8eaed',
                          relief=tk.FLAT)
        
        # Add hover effects via bindings
        if style == 'primary':
            btn.bind("<Enter>", lambda e: btn.configure(bg=self.C['primary_hover']))
            btn.bind("<Leave>", lambda e: btn.configure(bg=self.C['primary']))
        
        return btn
    
    def _build_ui(self):
        root = self.root
        
        # ── Header Banner ──
        header = tk.Frame(root, bg='#1a73e8', height=80)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        tk.Label(header, text="🌳  股权架构树生成器", 
                font=("Microsoft YaHei", 20, "bold"), 
                bg='#1a73e8', fg='white').place(x=24, y=14)
        tk.Label(header, text="从Excel数据一键生成交互式股权架构树HTML", 
                font=("Microsoft YaHei", 10), 
                bg='#1a73e8', fg='rgba(255,255,255,0.8)').place(x=24, y=48)
        
        # ── Main Content Area (scrollable form) ──
        body = tk.Frame(root, bg=self.C['bg'])
        body.pack(fill=tk.BOTH, expand=True, padx=20, pady=16)
        
        # ── Card: 数据输入 ──
        card1 = self._make_card(body)
        card1.pack(fill=tk.X, pady=(0, 12))
        
        tk.Label(card1, text="数据输入", font=("Microsoft YaHei", 12, "bold"),
                bg='#ffffff', fg=self.C['text']).pack(anchor=tk.W, padx=16, pady=(12, 2))
        ttk.Separator(card1, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=16)
        
        # Data file
        c1, _ = self._make_row(card1, "数据文件")
        self.data_entry = tk.Entry(c1, textvariable=self.data_path,
                         font=("Microsoft YaHei", 10), bg='#f8f9fa',
                         relief=tk.FLAT, highlightthickness=1,
                         highlightcolor=self.C['border'],
                         highlightbackground=self.C['border'])
        self.data_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=3)
        self._make_btn(c1, "📂 浏览", self._sel_data, 'outline').pack(side=tk.LEFT, padx=(6,0))
        
        # Biz file
        c2, _ = self._make_row(card1, "客户数据")
        tk.Entry(c2, textvariable=self.biz_path,
                font=("Microsoft YaHei", 10), bg='#f8f9fa',
                relief=tk.FLAT, highlightthickness=1,
                highlightcolor=self.C['border'],
                highlightbackground=self.C['border']).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=3)
        self._make_btn(c2, "📂 浏览", self._sel_biz, 'outline').pack(side=tk.LEFT, padx=(6,0))
        tk.Label(c2, text="可选", font=("Microsoft YaHei", 9), bg='#ffffff', fg=self.C['text_sec']).pack(side=tk.LEFT, padx=4)
        
        # ── Card: 参数设置 ──
        card2 = self._make_card(body)
        card2.pack(fill=tk.X, pady=(0, 12))
        
        tk.Label(card2, text="参数设置", font=("Microsoft YaHei", 12, "bold"),
                bg='#ffffff', fg=self.C['text']).pack(anchor=tk.W, padx=16, pady=(12, 2))
        ttk.Separator(card2, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=16)
        
        # Title
        c3, _ = self._make_row(card2, "HTML标题")
        tk.Entry(c3, textvariable=self.title_var,
                font=("Microsoft YaHei", 10), bg='#f8f9fa',
                relief=tk.FLAT, highlightthickness=1,
                highlightcolor=self.C['border'],
                highlightbackground=self.C['border']).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=3)
        
        # Root node
        c4, _ = self._make_row(card2, "指定根节点")
        tk.Entry(c4, textvariable=self.root_var,
                font=("Microsoft YaHei", 10), bg='#f8f9fa',
                relief=tk.FLAT, highlightthickness=1,
                highlightcolor=self.C['border'],
                highlightbackground=self.C['border']).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=3)
        tk.Label(c4, text="留空自动识别", font=("Microsoft YaHei", 9), bg='#ffffff', fg=self.C['text_sec']).pack(side=tk.LEFT, padx=4)
        
        # Output
        c5, _ = self._make_row(card2, "输出位置")
        tk.Entry(c5, textvariable=self.output_dir,
                font=("Microsoft YaHei", 10), bg='#f8f9fa',
                relief=tk.FLAT, highlightthickness=1,
                highlightcolor=self.C['border'],
                highlightbackground=self.C['border']).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=3)
        self._make_btn(c5, "📁 选择", self._sel_dir, 'outline').pack(side=tk.LEFT, padx=(6,0))
        
        # ── Card: 构建选项 ──
        card3 = self._make_card(body)
        card3.pack(fill=tk.X, pady=(0, 12))
        
        # Mode + template on same line inside a flex row
        opt_row = tk.Frame(card3, bg='#ffffff')
        opt_row.pack(fill=tk.X, padx=16, pady=(10, 4))
        
        # Left side: mode selector
        mode_frame = tk.Frame(opt_row, bg='#ffffff')
        mode_frame.pack(side=tk.LEFT)
        tk.Label(mode_frame, text="构建模式:", font=("Microsoft YaHei", 10, "bold"),
                bg='#ffffff', fg=self.C['text']).pack(side=tk.LEFT)
        
        for txt, val in [("控制权树", "control"), ("投资穿透树", "invest")]:
            rb = tk.Radiobutton(mode_frame, text=txt, variable=self.mode_var, value=val,
                              font=("Microsoft YaHei", 10), bg='#ffffff',
                              activebackground='#ffffff',
                              selectcolor='#ffffff', indicatoron=0,
                              padx=12, pady=2, cursor="hand2")
            rb.pack(side=tk.LEFT, padx=(6, 0))
            # Style: selected vs unselected
            def upd_rb():
                for w in mode_frame.winfo_children():
                    if isinstance(w, tk.Radiobutton):
                        if w.cget('value') == self.mode_var.get():
                            w.configure(bg=self.C['primary'], fg='white')
                        else:
                            w.configure(bg='#f1f3f4', fg=self.C['text'])
            self.mode_var.trace_add('write', lambda *a: upd_rb())
            rb.bind("<ButtonRelease-1>", lambda e: upd_rb())
        upd_rb()
        
        # Right side: template + clean check
        right_frame = tk.Frame(opt_row, bg='#ffffff')
        right_frame.pack(side=tk.RIGHT)
        
        self.tpl_btn = self._make_btn(right_frame, "📋 模板", self._gen_template, 'outline')
        self.tpl_btn.pack(side=tk.LEFT)
        
        tk.Checkbutton(right_frame, text="剔除异常企业", variable=self.clean_status_var,
                      font=("Microsoft YaHei", 9), bg='#ffffff',
                      activebackground='#ffffff',
                      selectcolor='#ffffff', fg=self.C['text_sec']).pack(side=tk.LEFT, padx=(10,0))
        
        # Hint below mode
        hint_row = tk.Frame(card3, bg='#ffffff')
        hint_row.pack(fill=tk.X, padx=16, pady=(0, 10))
        mode_hint = tk.Label(hint_row, text="控制权树=取最大股东追溯控制链   |   投资穿透树=从指定人出发展开展示",
                           font=("Microsoft YaHei", 9), bg='#ffffff', fg=self.C['text_sec'])
        mode_hint.pack(side=tk.LEFT)
        
        # ── Generate Button ──
        btn_frame = tk.Frame(body, bg=self.C['bg'])
        btn_frame.pack(fill=tk.X, pady=(4, 0))
        
        self.gen_btn = self._make_btn(btn_frame, "🚀  生成股权树HTML", self._generate, 'primary')
        self.gen_btn.pack(side=tk.LEFT, padx=(0, 8))
        
        self.open_btn = self._make_btn(btn_frame, "📂 打开文件", self._open_output, 'outline')
        self.open_btn.configure(state=tk.DISABLED)
        self.open_btn.pack(side=tk.LEFT)
        
        # ── Log Area ──
        log_frame = tk.Frame(body, bg='#ffffff', bd=0, relief=tk.FLAT,
                            highlightthickness=1, highlightcolor=self.C['border'],
                            highlightbackground=self.C['border'])
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        
        self.log_text = tk.Text(log_frame, height=6,
                               font=("Consolas", 10),
                               wrap=tk.WORD, state=tk.DISABLED,
                               bg='#fafafa', fg=self.C['text'],
                               relief=tk.FLAT, padx=12, pady=8,
                               border=0)
        scroll = tk.Scrollbar(log_frame, command=self.log_text.yview, width=10, bd=0)
        self.log_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)
    
    def _sel_data(self):
        path = filedialog.askopenfilename(title="选择数据文件", 
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("JSON文件", "*.json"), ("所有文件", "*.*")])
        if path:
            self.data_path.set(path)
            # Auto-set title from filename
            basename = os.path.splitext(os.path.basename(path))[0]
            if not self.title_var.get() or self.title_var.get() == "股权关系树":
                self.title_var.set(f"{basename}股权关系树")
            self._log(f"已选择数据文件: {os.path.basename(path)}")
    
    def _sel_biz(self):
        path = filedialog.askopenfilename(title="选择客户数据文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")])
        if path:
            self.biz_path.set(path)
            self._log(f"已选择客户数据: {os.path.basename(path)}")
    
    def _sel_dir(self):
        path = filedialog.askdirectory(title="选择输出目录", initialdir=self.output_dir.get())
        if path:
            self.output_dir.set(path)
    
    def _gen_template(self):
        """Generate an empty Excel template"""
        path = filedialog.asksaveasfilename(
            title="保存模板",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile="股权树数据模板.xlsx"
        )
        if not path:
            return
        try:
            import openpyxl
            wb = openpyxl.Workbook()
            
            ws = wb.active
            ws.title = "股权树数据"
            h = ['企业名称', '上级企业名称', '持股比例', '经营状态']
            ws.append(h)
            ws.append(['鞍钢集团有限公司', '', '', '存续'])
            ws.append(['鞍钢集团矿业有限公司', '鞍钢集团有限公司', '100', '存续'])
            for col in range(1, 5):
                ws.cell(column=col, row=1).font = openpyxl.styles.Font(bold=True, color="0969DA")
            ws.column_dimensions['A'].width = 35
            ws.column_dimensions['B'].width = 30
            
            ws2 = wb.create_sheet("银行客户数据（可选）")
            h2 = ['企业名称', '开户日期', '日均存款', '时点存款', '授信余额', '年化收入']
            ws2.append(h2)
            ws2.append(['鞍钢集团矿业有限公司', '2015-12-23', '6467.92', '6211.11', '0', '120.65'])
            for col in range(1, 7):
                ws2.cell(column=col, row=1).font = openpyxl.styles.Font(bold=True, color="2B7A78")
            ws2.column_dimensions['A'].width = 30
            
            wb.save(path)
            self._log(f"✅ 模板已生成: {path}")
            self._log("  在'股权树数据'sheet中填写企业名称和上下级关系")
            self._log("  第一行是示例数据，可删除或覆盖")
            messagebox.showinfo("完成", f"Excel模板已生成:\n{path}")
        except Exception as e:
            self._log(f"❌ 生成模板失败: {e}")
            messagebox.showerror("错误", f"生成模板失败:\n{e}")
    
    def _log(self, msg):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self.root.update_idletasks()
    
    def _open_output(self):
        if self._last_output and os.path.exists(self._last_output):
            import subprocess
            if sys.platform == 'darwin':
                subprocess.run(['open', self._last_output])
            elif sys.platform == 'win32':
                os.startfile(self._last_output)
    
    def _generate(self):
        if not self.data_path.get():
            messagebox.showwarning("提示", "请先选择数据文件")
            return
        
        self.gen_btn.configure(state=tk.DISABLED, text="⏳ 生成中...")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self._last_output = None
        self.open_btn.configure(state=tk.DISABLED)
        
        # Run in background to keep UI responsive
        thread = threading.Thread(target=self._do_generate, daemon=True)
        thread.start()
    
    def _do_generate(self):
        try:
            ext = os.path.splitext(self.data_path.get())[1].lower()
            title = self.title_var.get().strip() or "股权关系树"
            out_name = f"{title.replace(' ', '_')}.html"
            out_path = os.path.join(self.output_dir.get(), out_name)
            
            # Load data
            self._log(f"📖 加载数据: {os.path.basename(self.data_path.get())}")
            
            if ext == '.json':
                with open(self.data_path.get(), 'r') as f:
                    tree = json.load(f)
                self._log(f"✅ JSON已加载，含 {count_nodes(tree)} 个节点")
            else:
                sheets = load_xlsx(self.data_path.get())
                self._log(f"  发现Sheet: {list(sheets.keys())}")
                data_rows = None
                for name, rows in sheets.items():
                    if len(rows) > 1:
                        data_rows = parse_sheet(rows)
                        if data_rows:
                            self._log(f"  ✅ 使用Sheet: {name}")
                            break
                if not data_rows:
                    self._log("❌ 未找到有效数据")
                    self._log("  检查要点: ①列名是否包含'企业名称' ②数据是否从第1行开始")
                    self._log("  支持的列名: 企业名称/公司名称/上级企业名称/持股比例/经营状态")
                    self._finish(False)
                    return
                
                self._log(f"✅ 解析完成: {len(data_rows)} 条记录")
                root_name = self.root_var.get().strip() or None
                mode = self.mode_var.get()
                mode_label = {'control':'控制权树','invest':'投资穿透树'}.get(mode, mode)
                tree = build_tree(data_rows, root_name, mode)
                if root_name:
                    self._log(f"  模式: {mode_label}，根节点: {root_name}")
                else:
                    self._log(f"  模式: {mode_label}，根节点: {tree['name']}（自动识别）")
                self._log(f"✅ 树构建完成: {count_nodes(tree)} 个节点")
            
            # Clean (status filter + empty ratio)
            removed = []
            if self.clean_status_var.get() and tree.get('children'):
                r = clean_status(tree)
                if r:
                    self._log(f"🧹 剔除经营异常企业: {len(r)}家")
                    for name in r[:10]:
                        self._log(f"   ✗ {name}")
                    if len(r) > 10:
                        self._log(f"   ... 等{len(r)}家")
            tree = clean_tree(tree)
            self._log(f"🧹 清洗完成: {count_nodes(tree)} 个节点")
            
            # Load biz data
            if self.biz_path.get() and os.path.exists(self.biz_path.get()):
                self._log(f"📊 加载客户数据...")
                biz_sheets = load_xlsx(self.biz_path.get())
                biz_rows = None
                for name, rows in biz_sheets.items():
                    if len(rows) > 1:
                        biz_rows = rows
                        break
                if biz_rows:
                    tree, matched = load_biz_data(biz_rows, tree)
                    self._log(f"🏦 客户匹配: {matched} 家")
            
            # Generate HTML
            self._log(f"🎨 生成HTML...")
            total = generate_html(tree, title, out_path, self._log)
            self._log(f"✅ 生成成功!")
            self._log(f"📁 输出: {out_path}")
            self._log(f"📊 总计 {total} 个节点")
            self._last_output = out_path
            self._finish(True)
            
        except Exception as e:
            self._log(f"❌ 错误: {str(e)}")
            import traceback
            self._log(traceback.format_exc())
            self._finish(False)
    
    def _finish(self, success):
        self.gen_btn.configure(state=tk.NORMAL, text="🚀  生成股权树HTML")
        if success:
            self.open_btn.configure(state=tk.NORMAL)
            messagebox.showinfo("完成", f"股权树HTML已生成!\n\n路径: {self._last_output}")

# ── 入口 ──
if __name__ == '__main__':
    root = tk.Tk()
    app = EquityTreeApp(root)
    root.mainloop()