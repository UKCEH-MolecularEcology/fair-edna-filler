#!/usr/bin/env python3
"""Check the FAIRe download page for a different FULL template version and, if found, switch to it.

    python3 scripts/check_template_update.py            # check and apply
    python3 scripts/check_template_update.py --dry-run  # only report what would change

The FAIRe download page (https://fair-edna.github.io/download.html) is treated as the
source of truth: whatever FULL template it links to is the one this tool should use.
That also covers a withdrawn release (e.g. FAIRe restored v1.0.2 after pulling v1.0.3).

When a change is applied, this script:
  * downloads the new template into template/ and removes the old one,
  * updates TEMPLATE_VERSION in server.py and the template links in README.md,
  * rebuilds index.html.
In GitHub Actions it writes `changed`, `old`, `new`, `url` to $GITHUB_OUTPUT.
"""
import os, re, sys, urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOWNLOAD_PAGE = 'https://fair-edna.github.io/download.html'
UA = {'User-Agent': 'fair-edna-filler-template-check'}


def fetch(url: str) -> bytes:
    with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=60) as r:
        return r.read()


def vkey(v: str):
    return tuple(int(x) for x in v.split('.'))


def output(**kw):
    if os.environ.get('GITHUB_OUTPUT'):
        with open(os.environ['GITHUB_OUTPUT'], 'a') as f:
            for k, v in kw.items(): f.write(f'{k}={v}\n')


def main():
    dry = '--dry-run' in sys.argv
    server_py = ROOT / 'server.py'
    current = re.search(r"^TEMPLATE_VERSION = '([^']+)'", server_py.read_text(encoding='utf-8'), re.M).group(1)

    page = fetch(DOWNLOAD_PAGE).decode('utf-8', errors='replace')
    links = re.findall(r'href="([^"]*FAIRe_checklist_v([\d.]+\d)_FULLtemplate\.xlsx)"', page)
    if not links:
        raise SystemExit(f'Could not find a FULL template link on {DOWNLOAD_PAGE} – the page layout may have changed.')
    href, listed = links[0]
    raw_url = href.replace('/blob/', '/raw/')
    print(f'Tool uses v{current}; FAIRe download page lists v{listed}')

    if listed == current:
        print('Up to date.'); output(changed='false'); return
    direction = 'newer' if vkey(listed) > vkey(current) else 'older (a release may have been withdrawn)'
    print(f'Listed version is {direction}.')
    if dry:
        output(changed='false'); return

    data = fetch(raw_url)
    if data[:2] != b'PK':
        raise SystemExit(f'Downloaded file is not an .xlsx: {raw_url}')
    tdir = ROOT / 'template'
    tdir.mkdir(exist_ok=True)
    for old in tdir.glob('FAIRe_checklist_v*_FULLtemplate.xlsx'): old.unlink()
    new_name = f'FAIRe_checklist_v{listed}_FULLtemplate.xlsx'
    (tdir / new_name).write_bytes(data)

    s = server_py.read_text(encoding='utf-8')
    server_py.write_text(re.sub(r"^TEMPLATE_VERSION = '[^']+'", f"TEMPLATE_VERSION = '{listed}'", s, count=1, flags=re.M), encoding='utf-8')
    readme = ROOT / 'README.md'
    r = readme.read_text(encoding='utf-8')
    r = r.replace(f'FAIRe_checklist_v{current}_FULLtemplate.xlsx', new_name).replace(f'(v{current})', f'(v{listed})')
    readme.write_text(r, encoding='utf-8')

    sys.path.insert(0, str(ROOT))
    import build_standalone
    (ROOT / 'index.html').write_text(build_standalone.build(), encoding='utf-8')
    print(f'Switched to v{listed}: template/{new_name}, server.py, README.md, index.html updated.')
    output(changed='true', old=current, new=listed, url=raw_url, direction=direction)


if __name__ == '__main__':
    main()
