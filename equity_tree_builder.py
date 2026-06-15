#!/usr/bin/env python3
"""
equity-tree-builder — 股权架构树生成工具

从结构化数据（Excel/JSON）构建交互式股权架构树HTML。

用法:
  # 从Excel生成
  equity-tree-builder nodes.xlsx [--biz 客户数据.xlsx] [--title "企业名称"] -o tree.html

  # 从JSON生成  
  equity-tree-builder tree.json -o tree.html

  # 仅从关系表+节点表生成（两sheet模式）
  equity-tree-builder data.xlsx --sheet-nodes nodes --sheet-rels rels -o tree.html

  # 生成空模板（供填写）
  equity-tree-builder --template template.xlsx

数据格式:
  Excel单sheet模式: 企业名称, 上级企业名称, 持股比例, 经营状态, 层级(可选)
  Excel双sheet模式: Sheet1=节点表, Sheet2=关系表, Sheet3(可选)=客户数据
  JSON格式: 标准递归树结构 {name, ratio, status, children: [...]}

客户数据字段(可选): 企业名称, 开户日期, 日均存款, 时点存款, 授信余额, 年化收入, 规模, 客户经理
"""

import argparse
import json
import os
import re
import sys
from collections import defaultdict

# ── 常量 ──
TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), 'templates')

# ── 辅助函数 ──

def find_header(col, headers):
    """模糊匹配列名，支持中英文别名"""
    aliases = {
        'name': ['企业名称', '公司名称', '公司名', '名称', 'entity_name', 'company'],
        'parent': ['上级企业名称', '上级企业', '母公司', '股东名称', '股东', '母公司名称', 'parent_name', 'parent'],
        'ratio': ['持股比例', '投资比例', '股权比例', 'ratio', 'invest_ratio'],
        'status': ['经营状态', '状态', 'status', 'reg_status'],
        'level': ['层级', '级别', 'level', 'depth'],
        'biz_open': ['开户日期', '开户时间', '开户日', 'open_date', 'openDate'],
        'biz_deposit_avg': ['日均存款', '日均', 'deposit_avg', 'depositAvg'],
        'biz_deposit_spot': ['时点存款', '时点', 'deposit_spot', 'depositSpot'],
        'biz_loan': ['授信余额', '贷款余额', '授信', 'loan_balance', 'loanBalance'],
        'biz_income': ['年化收入', '净收入', '年收入', 'net_income', 'netIncome12m'],
        'biz_scale': ['规模', '企业规模', 'scale'],
        'biz_manager': ['客户经理', '经理', '管户', 'manager'],
    }
    h = str(col).strip().lower()
    for key, names in aliases.items():
        for n in names:
            if n.lower() in h or h in n.lower():
                return key
    return None

def load_xlsx(path):
    """加载xlsx文件，返回所有sheet的dict"""
    try:
        import openpyxl
    except ImportError:
        print("错误: 需要 openpyxl 库。运行: pip3 install openpyxl")
        sys.exit(1)
    
    wb = openpyxl.load_workbook(path, data_only=True)
    sheets = {}
    for name in wb.sheetnames:
        ws = wb[name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append([str(v) if v is not None else '' for v in row])
        sheets[name] = rows
    wb.close()
    return sheets

def parse_sheet(rows, field_map):
    """将sheet的行解析为dict列表，按field_map映射列索引"""
    if not rows or len(rows) < 2:
        return []
    
    header = rows[0]
    # Build column index mapping
    col_idx = {}
    for i, h in enumerate(header):
        key = find_header(h, None)
        if key:
            col_idx[key] = i
    
    result = []
    for row in rows[1:]:
        if not any(row):
            continue
        rec = {}
        for key, idx in col_idx.items():
            if idx < len(row):
                rec[key] = row[idx].strip() if isinstance(row[idx], str) else row[idx]
        if rec.get('name'):
            result.append(rec)
    return result

def build_tree_from_relations(nodes, relations, root_name=None):
    """
    从节点列表和关系列表构建树。
    nodes: [{name, status?, ...}]
    relations: [{name, parent, ratio?}]
    """
    # Build node lookup
    node_info = {}
    for n in nodes:
        name = n.get('name', '').strip()
        if name:
            node_info[name] = {k: v for k, v in n.items() if k != 'name'}
    
    # Build parent->children map
    children_map = defaultdict(list)
    all_children = set()
    for rel in relations:
        child = rel.get('name', '').strip()
        parent = rel.get('parent', '').strip()
        if child and parent:
            ratio = rel.get('ratio', '')
            children_map[parent].append({'name': child, 'ratio': str(ratio)})
            all_children.add(child)
    
    # Find root(s): nodes that are listed as parents but never appear as children
    all_parents = set(children_map.keys())
    if root_name:
        roots = [root_name]
    else:
        roots = list(all_parents - all_children)
        if not roots:
            # Fallback: the node with most children
            roots = [max(children_map, key=lambda k: len(children_map[k]))]
    
    def build_subtree(name):
        info = node_info.get(name, {})
        node = {
            'name': name,
            'ratio': info.get('ratio', ''),
            'status': info.get('status', ''),
            'children': []
        }
        # Add biz data if present
        for bk in ['_biz']:
            if bk in info:
                node[bk] = info[bk]
        
        for child in children_map.get(name, []):
            child_node = build_subtree(child['name'])
            if child_node:
                # The ratio in the relation overwrites if specified
                if child.get('ratio', ''):
                    child_node['ratio'] = child['ratio']
                node['children'].append(child_node)
        
        return node
    
    tree = build_subtree(roots[0])
    return tree

def build_tree_from_flat(rows):
    """
    从扁平表构建树（单sheet模式）。
    rows: [{name, parent, ratio?, status?, ...}]
    """
    relations = []
    node_map = {}
    
    for r in rows:
        name = r.get('name', '').strip()
        parent = r.get('parent', '').strip()
        if not name:
            continue
        
        # Store node info
        node_map[name] = {
            'status': r.get('status', ''),
            'ratio': r.get('ratio', ''),
        }
        
        # Copy biz fields
        biz_fields = ['biz_open', 'biz_deposit_avg', 'biz_deposit_spot', 
                       'biz_loan', 'biz_income', 'biz_scale', 'biz_manager']
        biz_data = {}
        for bf in biz_fields:
            if r.get(bf):
                biz_data[bf] = r[bf]
        
        if biz_data:
            node_map[name]['_biz'] = biz_data
        
        if parent:
            relations.append({'name': name, 'parent': parent, 'ratio': r.get('ratio', '')})
    
    # Find root
    all_parents = set(r['parent'] for r in relations)
    all_children = set(r['name'] for r in relations)
    roots = list(all_parents - all_children)
    
    if not roots:
        # Highest parent by convention
        roots = [max(relations, key=lambda r: r['name'] in all_parents)['parent']]
    
    root_name = roots[0] if roots else ''
    
    # Build nodes list
    nodes = [{'name': k, **v} for k, v in node_map.items()]
    
    return build_tree_from_relations(nodes, relations, root_name)

def load_biz_data(path, tree):
    """
    从Excel加载银行客户数据，匹配到树节点。
    biz格式: 企业名称, 开户日期, 日均存款, 时点存款, 授信余额, 年化收入, 规模, 客户经理
    """
    sheets = load_xlsx(path)
    # Find the first sheet with data
    biz_rows = None
    for name, rows in sheets.items():
        if len(rows) > 1 and rows[0][0] and '企业' in str(rows[0][0]):
            biz_rows = parse_sheet(rows, None)
            break
    
    if not biz_rows:
        print("警告: 未找到客户数据sheet")
        return tree
    
    # Re-parse with proper header matching
    for name, rows in sheets.items():
        if len(rows) > 1:
            biz_records = []
            header = rows[0]
            col_map = {}
            for i, h in enumerate(header):
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
                continue
            
            for row in rows[1:]:
                if not row or not row[col_map['name']]:
                    continue
                biz = {}
                for k, idx in col_map.items():
                    if idx < len(row) and row[idx]:
                        v = row[idx]
                        # Numeric fields
                        if k in ['depositAvg', 'depositSpot', 'loanBalance', 'netIncome12m', 'payrollCount', 'payrollAmount', 'pensionCardCount']:
                            try:
                                v = float(str(v).replace(',', ''))
                            except:
                                v = 0
                        biz[k] = v
                biz_records.append(biz)
            
            if biz_records:
                break
    
    if not biz_records:
        print("警告: 未解析到客户数据")
        return tree
    
    # Build lookup by name
    biz_lookup = {}
    for b in biz_records:
        name = str(b.get('name', '')).strip()
        if name:
            biz_lookup[name] = b
    
    # Match to tree nodes
    matched = 0
    def match_biz(node):
        nonlocal matched
        name = node.get('name', '')
        if name in biz_lookup:
            b = biz_lookup[name]
            node['_biz'] = {
                'hasAccount': True,
                'openDate': str(b.get('openDate', '')),
                'depositAvg': float(b.get('depositAvg', 0)),
                'depositSpot': float(b.get('depositSpot', 0)),
                'depositTotalAvg': float(b.get('depositAvg', 0)),
                'loanBalance': float(b.get('loanBalance', 0)),
                'interestIncome': 0,
                'netIncome': float(b.get('netIncome12m', 0)) / 4,
                'netIncome12m': float(b.get('netIncome12m', 0)),
                'payrollCount': int(b.get('payrollCount', 0)),
                'payrollAmount': float(b.get('payrollAmount', 0)),
                'pensionCardCount': int(b.get('pensionCardCount', 0)),
                'scale': str(b.get('scale', '')),
                'manager': str(b.get('manager', '')),
            }
            matched += 1
        for c in node.get('children', []):
            match_biz(c)
    
    match_biz(tree)
    print(f"客户数据匹配: {matched} 家")
    return tree

def clean_tree(tree):
    """基本清洗：去除空ratio的一级节点（疑似误入）"""
    if 'children' in tree:
        kept = []
        for c in tree['children']:
            r = c.get('ratio', '')
            if not r or str(r).strip() == '':
                # This is suspicious - check if it has any biz data
                if not c.get('_biz') and not c.get('children'):
                    print(f"  清洗: 移除空ratio一级节点 '{c['name']}'")
                    continue
            kept.append(c)
        tree['children'] = kept
    return tree

def generate_html(tree, title='股权关系树', output_path='equity_tree.html'):
    """从模板生成交互式HTML"""
    head_path = os.path.join(TEMPLATE_DIR, 'head.html')
    tail_path = os.path.join(TEMPLATE_DIR, 'tail.html')
    
    if not os.path.exists(head_path) or not os.path.exists(tail_path):
        # Fallback: use the existing equity tree HTML as template
        print("模板文件未找到，使用内嵌模板...")
        return generate_html_fallback(tree, title, output_path)
    
    with open(head_path, 'r') as f:
        head = f.read()
    with open(tail_path, 'r') as f:
        tail = f.read()
    
    # Patch title
    head = re.sub(r'<title>[^<]+</title>', f'<title>{title}</title>', head)
    head = re.sub(r'鞍钢集团控股股权关系树', title, head)
    
    # Serialize tree
    tree_json = json.dumps(tree, ensure_ascii=False, separators=(',', ':'))
    
    # Count nodes
    def count_nodes(n):
        cnt = 1
        for c in n.get('children', []):
            cnt += count_nodes(c)
        return cnt
    total = count_nodes(tree)
    
    # Patch initial node count in the HTML
    head = head.replace('>1372节点<', f'>{total}节点<')
    head = head.replace('>1372<', f'>{total}<')
    
    # Combine
    html = head + 'var TD=' + tree_json + ';' + tail
    
    # Write output
    with open(output_path, 'w') as f:
        f.write(html)
    
    print(f"输出: {output_path}")
    print(f"总节点: {total}")
    return output_path

def generate_html_fallback(tree, title, output_path):
    """Fallback: 从零生成HTML（内嵌D3和完整交互逻辑的迷你版）"""
    # Minimal inline version if template is missing
    # This is a simplified generator that creates a working tree
    tree_json = json.dumps(tree, ensure_ascii=False, separators=(',', ':'))
    
    def count_nodes(n):
        cnt = 1
        for c in n.get('children', []):
            cnt += count_nodes(c)
        return cnt
    total = count_nodes(tree)
    
    html = f'''<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<title>{title}</title>
<style>
@media print{{@page{{size:landscape;margin:15mm 10mm}}body{{overflow:visible;height:auto}}#bar,#legend,.tip,#srch,#pf,#pctInput,#pc,#bizToggle,#zin,#zout,#exp,#col,#rst,#printBtn,#layoutToggle{{display:none!important}}svg{{position:static!important;width:100%!important;height:auto!important}}.node text{{font-size:16px!important}}.link-label{{font-size:12px!important}}}}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:-apple-system,"Microsoft YaHei",sans-serif;background:#fff;overflow:hidden;height:100vh}}
#bar{{position:fixed;top:0;left:0;right:0;z-index:10;background:rgba(255,255,255,.95);padding:8px 14px;display:flex;gap:8px;align-items:center;border-bottom:1px solid #d0d7de;flex-wrap:wrap;font-size:14px}}
#bar button,#bar input{{background:#f6f8fa;border:1px solid #d0d7de;padding:6px 12px;border-radius:5px;font-size:13px;cursor:pointer}}
svg{{position:fixed;top:46px;left:0;width:100%;height:calc(100vh-46px)}}
.node rect{{rx:4;ry:4}}
.node text{{font-size:13px;fill:#24292f;text-anchor:middle;dominant-baseline:central;pointer-events:none}}
.link{{fill:none;stroke:#d0d7de;stroke-width:2px}}
.link-label{{font-size:10px;fill:#656d76;text-anchor:middle}}
#legend{{position:fixed;bottom:0;left:0;right:0;z-index:10;background:rgba(255,255,255,.95);padding:6px 14px;border-top:1px solid #d0d7de;font-size:13px;color:#24292f}}
</style>
</head>
<body>
<div id="bar"><span style="color:#0969da;font-weight:bold">{title} · {total}节点</span>
<button id="zin">🔍+</button><button id="zout">🔍-</button>
<button id="exp">📂展开</button><button id="col">📁折叠</button>
<button id="rst">↺重置</button>
<button id="printBtn" style="background:#0969da;color:#fff">🖨 打印</button>
<input id="srch" placeholder="搜索企业..." style="width:140px">
</div>
<svg id="svg"></svg>
<div id="legend"><span>鼠标悬停查看详情 | 拖拽平移 | 滚轮缩放</span></div>
<script>
// Minimal D3 v7 inlined would be too large; use CDN or instruct user
// Using CDN - user must have internet
document.write('<script src="https://d3js.org/d3.v7.min.js"><\\/script>');
</script>
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
  g.selectAll(".link").data(ls,function(d){{return d.target.data.name}}).join("path").attr("class","link")
    .attr("d",function(d){{return"M"+d.source.x+","+(d.source.y+22)+"C"+d.source.x+","+((d.source.y+d.target.y)/2)+" "+d.target.x+","+((d.source.y+d.target.y)/2)+" "+d.target.x+","+(d.target.y-22)}});
  g.selectAll(".link-label").data(ls,function(d){{return d.target.data.name}}).join("text").attr("class","link-label")
    .attr("x",function(d){{return(d.source.x+d.target.x)/2}}).attr("y",function(d){{return(d.source.y+d.target.y)/2-5}})
    .text(function(d){{var dr=d.data.ratio;return dr&&dr!=""?dr:""}});
  g.selectAll(".node").data(ns,function(d){{return d.data.name}}).join("g").attr("class","node")
    .attr("transform",function(d){{return"translate("+d.x+","+d.y+")"}}).style("cursor","pointer")
    .on("click",function(ev,d){{ev.stopPropagation();if(d.children){{d._children=d.children;d.children=null}}else if(d._children){{d.children=d._children;d._children=null}}upd()}})
    .on("mouseenter",function(ev,d){{var k=d.children||d._children||[];var tip=d.data.name+" ("+(d.depth+1)+"\\u7ea7"+(d.data._biz?" \\ud83c\\udfe6":"")+")\\n\\u4e0b\\u7ea7"+k.length+"\\u5bb6";if(d.data._biz){{tip+="\\n\\u4ea4\\u884c \\u5f00\\u6237"+d.data._biz.openDate}}document.getElementById("tip").innerHTML=tip.replace(/\\n/g,"<br>")}})
    .on("mousemove",function(ev){{var t=document.getElementById("tip");t.style.left=(ev.clientX+15)+"px";t.style.top=(ev.clientY-10)+"px";t.style.opacity="1"}})
    .on("mouseleave",function(){{document.getElementById("tip").style.opacity="0"}})
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
document.getElementById("zin").onclick=function(){{svg.transition().duration(300).call(zoom.scaleBy,1.3)}};
document.getElementById("zout").onclick=function(){{svg.transition().duration(300).call(zoom.scaleBy,0.7)}};
document.getElementById("exp").onclick=function(){{root.descendants().forEach(function(d){{if(d._children){{d.children=d._children;d._children=null}}}});upd()}};
document.getElementById("col").onclick=function(){{root.descendants().forEach(function(d){{if(d.depth>=1&&d.children){{d._children=d.children;d.children=null}}}});upd()}};
document.getElementById("rst").onclick=function(){{svg.transition().duration(500).call(zoom.transform,d3.zoomIdentity.translate(svgEl.clientWidth/2-root.x*is,60).scale(is))}};
document.getElementById("printBtn").onclick=function(){{root.descendants().forEach(function(d){{if(d._children){{d.children=d._children;d._children=null}}}});upd();setTimeout(function(){{window.print()}},300)}};
</script>
</body>
</html>'''
    
    with open(output_path, 'w') as f:
        f.write(html)
    print(f"输出: {output_path} (简约版，需联网加载D3)")
    return output_path

def generate_template_xlsx(path):
    """生成Excel空模板"""
    try:
        import openpyxl
    except ImportError:
        print("错误: 需要 openpyxl")
        sys.exit(1)
    
    wb = openpyxl.Workbook()
    
    # Sheet 1: 节点数据（扁平单表模式）
    ws = wb.active
    ws.title = "股权树数据（单表模式）"
    headers = ['企业名称', '上级企业名称', '持股比例', '经营状态', '层级']
    ws.append(headers)
    ws.append(['鞍钢集团有限公司', '', '', '存续', '1'])
    ws.append(['鞍钢集团矿业有限公司', '鞍钢集团有限公司', '100', '存续', '2'])
    ws.append(['鞍钢集团财务有限责任公司', '鞍钢集团有限公司', '70', '存续', '2'])
    ws.append(['鞍钢联众(广州)不锈钢有限公司', '鞍钢集团有限公司', '60', '存续', '2'])
    for col in range(1, 6):
        ws.cell(column=col, row=1).font = openpyxl.styles.Font(bold=True)
    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 30
    
    # Sheet 2: 客户数据（可选）
    ws2 = wb.create_sheet("银行客户数据（可选）")
    bh = ['企业名称', '开户日期', '日均存款', '时点存款', '授信余额', '年化收入', '规模', '客户经理']
    ws2.append(bh)
    ws2.append(['鞍钢集团矿业有限公司', '2015-12-23', '6467.92', '6211.11', '0', '120.65', '大型', '张小驰'])
    for col in range(1, 9):
        ws2.cell(column=col, row=1).font = openpyxl.styles.Font(bold=True)
    ws2.column_dimensions['A'].width = 30
    
    # Sheet 3: 双表模式的节点表
    ws3 = wb.create_sheet("节点表（双表模式）")
    n3 = ['企业名称', '经营状态']
    ws3.append(n3)
    for col in range(1, 3):
        ws3.cell(column=col, row=1).font = openpyxl.styles.Font(bold=True)
    ws3.column_dimensions['A'].width = 35
    
    # Sheet 4: 双表模式的关系表
    ws4 = wb.create_sheet("关系表（双表模式）")
    r4 = ['子公司名称', '母公司名称', '持股比例']
    ws4.append(r4)
    for col in range(1, 4):
        ws4.cell(column=col, row=1).font = openpyxl.styles.Font(bold=True)
    ws4.column_dimensions['A'].width = 35
    ws4.column_dimensions['B'].width = 30
    
    wb.save(path)
    print(f"模板已生成: {path}")
    print("  Sheet1: 股权树数据（单表模式）— 一表搞定")
    print("  Sheet2: 银行客户数据（可选）— 嵌入业务数据")
    print("  Sheet3-4: 双表模式（节点表+关系表）")

# ── CLI入口 ──

def main():
    parser = argparse.ArgumentParser(
        description='equity-tree-builder — 从数据文件生成股权架构树',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
数据格式说明:
  Excel单表模式: 列名包含 企业名称, 上级企业名称, 持股比例, 经营状态
  Excel双表模式: --sheet-nodes 节点表, --sheet-rels 关系表
  JSON模式: {name, ratio, status, children: [...]}
  
客户数据Excel: 列名包含 企业名称, 开户日期, 日均存款, 时点存款, 授信余额, 年化收入
        """
    )
    parser.add_argument('input', nargs='?', help='输入文件 (Excel/JSON)')
    parser.add_argument('--biz', help='银行客户数据Excel文件')
    parser.add_argument('--title', default='股权关系树', help='HTML标题')
    parser.add_argument('-o', '--output', default='equity_tree.html', help='输出HTML路径')
    parser.add_argument('--template', action='store_true', help='生成Excel模板')
    parser.add_argument('--template-output', default='equity_tree_template.xlsx', help='模板输出路径')
    parser.add_argument('--sheet-nodes', help='双表模式: 节点表sheet名')
    parser.add_argument('--sheet-rels', help='双表模式: 关系表sheet名')
    parser.add_argument('--root', help='指定根节点名称')
    parser.add_argument('--no-clean', action='store_true', help='跳过空ratio清洗')
    parser.add_argument('--json', action='store_true', help='强制JSON模式（输入文件为JSON）')
    
    args = parser.parse_args()
    
    # Generate template
    if args.template:
        generate_template_xlsx(args.template_output)
        return
    
    # Validate input
    if not args.input:
        parser.print_help()
        print("\n错误: 请指定输入文件。先用 --template 生成模板。")
        sys.exit(1)
    
    # Detect file type
    ext = os.path.splitext(args.input)[1].lower()
    is_json = args.json or ext == '.json'
    
    tree = None
    
    if is_json:
        # JSON mode
        with open(args.input, 'r') as f:
            tree = json.load(f)
        print(f"已加载JSON: {args.input}")
    
    else:
        # Excel mode
        print(f"加载Excel: {args.input}")
        sheets = load_xlsx(args.input)
        
        if args.sheet_nodes and args.sheet_rels:
            # Dual sheet mode
            node_rows = sheets.get(args.sheet_nodes, [])
            rel_rows = sheets.get(args.sheet_rels, [])
            nodes = parse_sheet(node_rows)
            relations = parse_sheet(rel_rows)
            tree = build_tree_from_relations(nodes, relations, args.root)
            print(f"双表模式: {len(nodes)}节点, {len(relations)}关系")
        else:
            # Single sheet mode - use first sheet with data
            data_rows = None
            for name, rows in sheets.items():
                if len(rows) > 1 and any('企业' in str(h) for h in rows[0]):
                    data_rows = parse_sheet(rows)
                    print(f"使用sheet: {name}")
                    break
            if not data_rows:
                # Use first non-empty sheet
                for name, rows in sheets.items():
                    if len(rows) > 1:
                        data_rows = parse_sheet(rows)
                        print(f"使用sheet: {name}")
                        break
            
            if not data_rows:
                print("错误: 未找到有效数据")
                sys.exit(1)
            
            tree = build_tree_from_flat(data_rows)
            print(f"单表模式: {len(data_rows)}条记录")
    
    # Clean
    if not args.no_clean:
        print("执行清洗...")
        tree = clean_tree(tree)
    
    # Attach biz data
    if args.biz:
        print(f"加载客户数据: {args.biz}")
        tree = load_biz_data(args.biz, tree)
    
    # Generate HTML
    generate_html(tree, args.title, args.output)

if __name__ == '__main__':
    main()