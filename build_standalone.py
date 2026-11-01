#!/usr/bin/env python3
"""Rebuild index.html (the no-install browser version) from server.py.

    python3 build_standalone.py            # write index.html
    python3 build_standalone.py --check    # exit 1 if index.html is out of date (used by CI)

index.html runs server.py itself in the browser (via Pyodide), so after any
change to server.py, run this script so both versions stay identical.
"""
import base64, re, sys
from pathlib import Path

here = Path(__file__).resolve().parent


def build() -> str:
    shell = (here / 'standalone' / 'shell.html').read_text(encoding='utf-8')
    py = (here / 'server.py').read_text(encoding='utf-8')
    if '</script' in py.lower():
        raise SystemExit('server.py must not contain "</script" – it is embedded in a <script> block.')
    version = re.search(r"^TEMPLATE_VERSION = '([^']+)'", py, re.M).group(1)
    url = f'https://github.com/FAIR-eDNA/FAIRe_checklist/raw/main/FAIRe_checklist_v{version}_FULLtemplate.xlsx'
    logo = base64.b64encode((here / 'static' / 'ukceh_logo.png').read_bytes()).decode()
    return (shell.replace('{{LOGO}}', 'data:image/png;base64,' + logo)
                 .replace('{{TEMPLATE_URL}}', url).replace('{{TEMPLATE_VERSION}}', version)
                 .replace('{{TEMPLATE_FILE}}', f'FAIRe_checklist_v{version}_FULLtemplate.xlsx')
                 .replace('{{FAIRE_PY}}', py))


if __name__ == '__main__':
    html = build()
    target = here / 'index.html'
    if '--check' in sys.argv:
        if not target.exists() or target.read_text(encoding='utf-8') != html:
            raise SystemExit('index.html is out of date – run: python3 build_standalone.py')
        print('index.html is up to date')
    else:
        target.write_text(html, encoding='utf-8')
        print(f'index.html written ({len(html) // 1024} KB)')
