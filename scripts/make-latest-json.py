#!/usr/bin/env python3
"""Writes release/latest.json, the small file the app reads to learn about a new version.

Usage: python3 scripts/make-latest-json.py [path-to-setup.exe] [path-to-release-notes.md]

Upload latest.json to the GitHub release together with the setup file. The app downloads
https://github.com/<repo>/releases/latest/download/latest.json, compares the version with its own,
and checks the downloaded setup file against the sha256 written here before starting it.
"""
import hashlib, json, os, re, sys

root = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
version = json.load(open(os.path.join(root, 'package.json')))['version']
setup = sys.argv[1] if len(sys.argv) > 1 else os.path.join(root, 'release', f'YourOfficeSuite-Setup-{version}.exe')
notes_file = sys.argv[2] if len(sys.argv) > 2 else os.path.join(root, 'docs', f'RELEASE_NOTES_{version}.md')
name = os.path.basename(setup)
if name != f'YourOfficeSuite-Setup-{version}.exe': sys.exit(f'Setup file name must be YourOfficeSuite-Setup-{version}.exe, got {name}')
h = hashlib.sha256()
with open(setup, 'rb') as f:
    for block in iter(lambda: f.read(1 << 20), b''): h.update(block)
notes = ''
if os.path.exists(notes_file):
    m = re.search(r'### New in this version\s*(.*?)(?:\n### |\Z)', open(notes_file, encoding='utf-8').read(), re.S)
    notes = m.group(1).strip() if m else ''
out = {'version': version, 'file': name, 'size': os.path.getsize(setup), 'sha256': h.hexdigest(), 'notes': notes}
target = os.path.join(os.path.dirname(os.path.abspath(setup)), 'latest.json')
json.dump(out, open(target, 'w', encoding='utf-8'), indent=1)
print('wrote', target); print(json.dumps(out, indent=1))
