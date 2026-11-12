#!/usr/bin/env python3
"""Applies the YourOfficeSuite changes to a freshly built editor engine.

Usage: python3 scripts/patch-engine.py <engine-dir>

Changes made (all listed in NOTICE.md):
  1. The converter wasm is stored uncompressed (Chromium does not decode brotli over plain http).
  2. Font table rows for Calibri, Andale Mono and Comic Sans MS get their bold and italic faces.
  3. The brand colour layer (build/brand/editor-theme.css) is appended to each editor's app.css.
  4. Loading screens use the brand colours.
  5. The spreadsheet cell cursor uses the brand colour.
"""
import glob, json, os, re, subprocess, sys

root = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else 'engine')
here = os.path.dirname(os.path.abspath(__file__))

# 1. converter wasm
wasm = os.path.join(root, 'sdkjs/common/wasm/x2t/x2t.wasm.br')
with open(wasm, 'rb') as f: head = f.read(4)
if head != b'\x00asm':
    subprocess.run(['node', '-e', "const z=require('zlib'),fs=require('fs');const f=process.argv[1];fs.writeFileSync(f,z.brotliDecompressSync(fs.readFileSync(f)))", wasm], check=True)
    print('converter wasm: decoded')

# 2. font rows
p = os.path.join(root, 'sdkjs/common/AllFonts.js'); s = open(p, encoding='utf-8-sig').read()
rows = re.findall(r'\[("[^"\n]+"(?:,-?\d+){8})\]', s); parsed = [json.loads('[' + r + ']') for r in rows]; full = {}
for r in parsed:
    if r[1] >= 0 and r[3] >= 0 and r[5] >= 0 and r[7] >= 0: full.setdefault(r[1], r)
n = 0
for raw, r in zip(rows, parsed):
    if r[1] >= 0 and (r[3] < 0 or r[5] < 0 or r[7] < 0) and r[1] in full:
        new = json.dumps([r[0]] + full[r[1]][1:], separators=(',', ':'), ensure_ascii=False)[1:-1]
        s = s.replace('[' + raw + ']', '[' + new + ']', 1); n += 1
open(p, 'w', encoding='utf-8-sig').write(s); print('font rows fixed:', n)

# 3. brand colour layer
css = open(os.path.join(here, '..', 'build', 'brand', 'editor-theme.css')).read(); MARK = '/* ===== YourOfficeSuite brand layer'
for f in glob.glob(os.path.join(root, 'web-apps/apps/*/main/resources/css/app.css')):
    t = open(f, encoding='utf-8').read(); k = t.find(MARK)
    if k >= 0: t = t[:k].rstrip('\n')
    open(f, 'w', encoding='utf-8').write(t + css)

# 4. loading screens
skel = ('<style id="yos-skel">.loadmask,.theme-night .loadmask,.theme-type-dark .loadmask{--sk-background-toolbar-header-word:#14121c !important;'
        '--sk-background-toolbar-header-cell:#14121c !important;--sk-background-toolbar-header-slide:#14121c !important;--sk-background-toolbar-header-pdf:#14121c !important;'
        '--sk-background-toolbar-header-visio:#14121c !important}.loadmask{--sk-background-toolbar:#fbfaff !important;--sk-canvas-background:#efecf8 !important;'
        '--sk-background-toolbar-tab:#ece7fb !important}</style>')
for f in glob.glob(os.path.join(root, 'web-apps/apps/*/main/index*.html')):
    t = open(f, encoding='utf-8').read(); t = re.sub(r'<style id="yos-skel">.*?</style>', '', t, flags=re.S)
    if '</head>' in t: t = t.replace('</head>', skel + '</head>', 1)
    open(f, 'w', encoding='utf-8').write(t)

# 5. spreadsheet cell cursor
p = os.path.join(root, 'sdkjs/cell/sdk-all.js'); s = open(p, encoding='utf-8').read()
s = s.replace('activeCellBorderColor=new n(72,121,92)', 'activeCellBorderColor=new n(12,160,184)'); open(p, 'w', encoding='utf-8').write(s)
print('engine patched:', root)
