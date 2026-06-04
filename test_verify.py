#!/usr/bin/env python3
"""
Quick test script for equity-tree-builder.
Validates that we can regenerate the Ansteel tree from JSON.
"""
import sys, json, os

# Add parent dir to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from equity_tree_builder import generate_html, clean_tree

# Load the reference tree
with open('/tmp/angang_tree.json') as f:
    tree = json.load(f)

# Test clean
cleaned = clean_tree(tree)
print(f"Before: {count_nodes(tree)} nodes")
print(f"After: {count_nodes(cleaned)} nodes")

# Generate
out = '/tmp/angang_verified.html'
generate_html(cleaned, '鞍钢集团股权关系树（验证）', out)
print(f"OK: {out}")