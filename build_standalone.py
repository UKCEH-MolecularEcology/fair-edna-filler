#!/usr/bin/env python3
"""Rebuild index.html (the no-install browser version) from server.py.

    python3 build_standalone.py

index.html runs server.py itself in the browser (via Pyodide), so after any
change to server.py, run this script so both versions stay identical.
"""
import base64
from pathlib import Path

here = Path(__file__).resolve().parent
shell = (here / 'standalone' / 'shell.html').read_text(encoding='utf-8')
py = (here / 'server.py').read_text(encoding='utf-8')
if '</script' in py.lower():
    raise SystemExit('server.py must not contain "</script" – it is embedded in a <script> block.')
logo = base64.b64encode((here / 'static' / 'ukceh_logo.png').read_bytes()).decode()
html = shell.replace('{{LOGO}}', 'data:image/png;base64,' + logo).replace('{{FAIRE_PY}}', py)
(here / 'index.html').write_text(html, encoding='utf-8')
print(f'index.html written ({len(html) // 1024} KB)')
