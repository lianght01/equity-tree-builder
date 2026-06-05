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

def build_tree(rows):
    relations = []
    node_map = {}
    for r in rows:
        name = r.get('name', '').strip()
        parent = r.get('parent', '').strip()
        if not name:
            continue
        node_map[name] = {
            'status': r.get('status', ''),
            'ratio': r.get('ratio', ''),
        }
        if parent:
            relations.append({'name': name, 'parent': parent, 'ratio': r.get('ratio', '')})
    
    all_parents = set(r['parent'] for r in relations)
    all_children = set(r['name'] for r in relations)
    roots = list(all_parents - all_children)
    if not roots:
        roots = [max(relations, key=lambda r: r['name'] in all_parents)['parent']] if relations else []
    if not roots:
        return {'name': rows[0]['name'], 'ratio': '', 'status': '', 'children': []}
    
    children_map = {}
    for rel in relations:
        p = rel['parent']
        if p not in children_map:
            children_map[p] = []
        children_map[p].append({'name': rel['name'], 'ratio': str(rel.get('ratio', ''))})
    
    def build_subtree(name):
        info = node_map.get(name, {})
        node = {'name': name, 'ratio': info.get('ratio', ''), 'status': info.get('status', ''), 'children': []}
        for child in children_map.get(name, []):
            cn = build_subtree(child['name'])
            if child.get('ratio', ''):
                cn['ratio'] = child['ratio']
            node['children'].append(cn)
        return node
    
    return build_subtree(roots[0])

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
        root.title("股权架构树生成器 v1.0")
        root.geometry("680x520")
        root.resizable(False, False)
        
        # Set icon if available
        try:
            root.iconphoto(True, tk.PhotoImage(file=os.path.join(TEMPLATE_DIR, '..', 'icon.png')))
        except:
            pass
        
        # Variables
        self.data_path = tk.StringVar()
        self.biz_path = tk.StringVar()
        self.title_var = tk.StringVar(value="股权关系树")
        self.output_dir = tk.StringVar(value=os.path.expanduser("~/Desktop"))
        
        # Build UI
        self._build_ui()
    
    def _build_ui(self):
        main = ttk.Frame(self.root, padding=20)
        main.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_lbl = tk.Label(main, text="股权架构树生成器", font=("Microsoft YaHei", 18, "bold"), fg="#0969da")
        title_lbl.pack(anchor=tk.W, pady=(0, 5))
        tk.Label(main, text="从Excel数据一键生成交互式股权架构树HTML", font=("Microsoft YaHei", 10), fg="#656d76").pack(anchor=tk.W, pady=(0, 15))
        
        # Data file
        f1 = ttk.Frame(main)
        f1.pack(fill=tk.X, pady=4)
        tk.Label(f1, text="数据文件 *", width=18, anchor=tk.W).pack(side=tk.LEFT)
        tk.Entry(f1, textvariable=self.data_path, width=40).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(f1, text="选择文件", command=self._sel_data, bg="#f6f8fa").pack(side=tk.LEFT, padx=2)
        
        # Biz file
        f2 = ttk.Frame(main)
        f2.pack(fill=tk.X, pady=4)
        tk.Label(f2, text="客户数据 (可选)", width=18, anchor=tk.W).pack(side=tk.LEFT)
        tk.Entry(f2, textvariable=self.biz_path, width=40).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(f2, text="选择文件", command=self._sel_biz, bg="#f6f8fa").pack(side=tk.LEFT, padx=2)
        
        # Title
        f3 = ttk.Frame(main)
        f3.pack(fill=tk.X, pady=4)
        tk.Label(f3, text="HTML标题", width=18, anchor=tk.W).pack(side=tk.LEFT)
        tk.Entry(f3, textvariable=self.title_var, width=40).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Template button row
        ft = ttk.Frame(main)
        ft.pack(fill=tk.X, pady=2)
        tk.Label(ft, text="", width=18).pack(side=tk.LEFT)
        self.tpl_btn = tk.Button(ft, text="📋 生成Excel模板（空表）", command=self._gen_template,
                                bg="#f6f8fa", fg="#24292f", font=("Microsoft YaHei", 10),
                                padx=10, pady=2, cursor="hand2", relief=tk.FLAT,
                                activebackground="#eaeef2")
        self.tpl_btn.pack(side=tk.LEFT)
        tk.Label(ft, text="先填模板再导入", font=("Microsoft YaHei", 9), fg="#888").pack(side=tk.LEFT, padx=8)
        
        # Output
        f4 = ttk.Frame(main)
        f4.pack(fill=tk.X, pady=4)
        tk.Label(f4, text="输出位置", width=18, anchor=tk.W).pack(side=tk.LEFT)
        tk.Entry(f4, textvariable=self.output_dir, width=40).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(f4, text="选择目录", command=self._sel_dir, bg="#f6f8fa").pack(side=tk.LEFT, padx=2)
        
        # Generate button
        f5 = ttk.Frame(main)
        f5.pack(fill=tk.X, pady=12)
        self.gen_btn = tk.Button(f5, text="🚀  生成股权树HTML", command=self._generate,
                                bg="#0969da", fg="white", font=("Microsoft YaHei", 12, "bold"),
                                padx=20, pady=6, cursor="hand2")
        self.gen_btn.pack()
        
        # Separator
        ttk.Separator(main, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=5)
        
        # Log area
        tk.Label(main, text="运行日志:", font=("Microsoft YaHei", 9), fg="#656d76").pack(anchor=tk.W)
        
        log_frame = ttk.Frame(main)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(log_frame, height=8, font=("Consolas", 10), 
                                wrap=tk.WORD, state=tk.DISABLED, bg="#f6f8fa", relief=tk.FLAT, padx=8, pady=6)
        scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Open button
        f6 = ttk.Frame(main)
        f6.pack(fill=tk.X, pady=(8,0))
        self.open_btn = tk.Button(f6, text="📂 打开输出文件", command=self._open_output,
                                 state=tk.DISABLED, bg="#f6f8fa", padx=10)
        self.open_btn.pack(side=tk.RIGHT)
        
        self._last_output = None
    
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
                tree = build_tree(data_rows)
                self._log(f"✅ 树构建完成: {count_nodes(tree)} 个节点")
            
            # Clean
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