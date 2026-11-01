#!/usr/bin/env python3
"""Benchmark the rule-based extractor on real reports.

    python3 bench/run_benchmark.py [folder_with_reports]     (default: ../benchmark_reports)

For each report (.pdf/.doc/.docx/.txt) it runs the full pipeline: read file ->
extract -> fill the bundled template -> read the filled values back.

* Reports recognised as the source of an official FAIRe example (see
  bench/answer_keys.json) are scored field by field: right / wrong / missed.
* Every report gets a coverage count and a list of values with the sentence
  each came from, so unscored reports can be checked by hand.

Writes bench/results.md and bench/values.csv. Report files are never copied.
"""
import csv, json, re, sys, tempfile, zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT))
import server  # noqa: E402

folder = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT.parent / 'benchmark_reports'
keys = json.loads((HERE / 'answer_keys.json').read_text(encoding='utf-8'))
MISSING = re.compile(r'^(missing:|not applicable)')


def norm(v):
    v = str(v).lower().replace('µ', 'u').replace('μ', 'u')
    v = re.sub(r'\[[a-z]+:\d+\]', '', v)                  # drop ontology IDs
    v = re.sub(r'(\d+)\.0\b', r'\1', v)
    return re.sub(r'[^a-z0-9.]+', ' ', v).strip()


def same(answer, ours):
    a, o = norm(answer), norm(ours)
    if not a or not o: return False
    if a == o: return True
    if re.fullmatch(r'[\d.]+', a) or re.fullmatch(r'[\d.]+', o): return a == o
    short, long_ = sorted((a, o), key=len)
    return len(short) >= 4 and short in long_


def read_filled(path):
    """projectMetadata {term: [values]} and the first 'sample' row of sampleMetadata {term: value}."""
    with zipfile.ZipFile(path) as z:
        pkg = {n: z.read(n) for n in z.namelist()}
    ss = server.shared_strings(pkg); paths = server.sheet_paths(pkg)
    proj = {}
    for n, r in server.parse_rows(pkg[paths['projectMetadata']].decode()).items():
        if n < 2: continue
        v = {c: server.cell_text(x, ss) for c, x in r['cells'].items()}
        vals = [x for c, x in sorted(v.items(), key=lambda kv: server.col_num(kv[0])) if server.col_num(c) >= 4 and x and not MISSING.match(x)]
        if v.get('C'): proj[v['C']] = vals
    rows = server.parse_rows(pkg[paths['sampleMetadata']].decode())
    hdr = {c: server.cell_text(x, ss) for c, x in rows[3]['cells'].items()}
    first = {}
    for n in sorted(rows):
        if n <= 3: continue
        v = {hdr.get(c): server.cell_text(x, ss) for c, x in rows[n]['cells'].items()}
        if v.get('samp_category') == 'sample':
            first = {k: x for k, x in v.items() if k and x and not MISSING.match(x)}; break
    return proj, first


def detect_key(text):
    """The answer key whose source document this is (matched on its title near the start of the text).
    Metabarcoding keys win when one document is the source of both kinds (e.g. NECR534)."""
    hits = [n for n, k in keys.items() if re.search(k['match'], text[:k.get('search_chars', 6000)])]
    hits.sort(key=lambda n: keys[n]['kind'] != 'metabarcoding')
    return hits[0] if hits else None


files = sorted(p for p in folder.glob('*') if p.suffix.lower() in ('.pdf', '.doc', '.docx', '.txt'))
if not files:
    raise SystemExit(f'No reports found in {folder}')

lines = ['# Extraction benchmark', '', f'Reports: {len(files)} (folder `{folder.name}`)', '']
summary_rows, value_rows = [], []
totals = {'right': 0, 'wrong': 0, 'missed': 0}
auto = {'right_ok': 0, 'right_check': 0, 'wrong_ok': 0, 'wrong_check': 0}
for f in files:
    try:
        text = server.extract_report_text(f)
    except Exception as e:
        lines += [f'## {f.name}', f'Could not read: {e}', '']; continue
    lang = server.detect_language(text) if hasattr(server, 'detect_language') else ('en', 'English')
    try:
        meta = server.prepare(text)
    except Exception as e:
        import traceback
        lines += [f'## {f.name}', f'**Extractor crashed:** `{e!r}`', '```', traceback.format_exc()[-1500:], '```', '']
        summary_rows.append({'file': f.name, 'chars': len(text), 'language': lang[1], 'assays': 'CRASH', 'sites': '', 'filled': '', 'sourced': '', 'key': detect_key(text) or '—'})
        continue
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / 'out.xlsx'
        server.fill_workbook(server.TEMPLATE_PATH, meta, f.name, out)
        proj, samp = read_filled(out)
    items = server.review(meta)
    status = {}
    for i in items:
        if i['status'] == 'check' or i['term'] not in status: status[i['term']] = i['status']
    filled = [i for i in items if i['value'] and not MISSING.match(i['value']) and not i['generated']]
    for i in filled:
        value_rows.append([f.name, i['id'], i['value'][:200], (i['evidence'] or '')[:300]])
    key_name = detect_key(text)
    rs = server.review_summary(items)
    row = {'file': f.name, 'auto': rs['ok'], 'check': rs['check'], 'chars': len(text), 'language': lang[1], 'assays': len(meta['assays']),
           'sites': len(meta['sites']), 'filled': len(filled), 'sourced': sum(1 for i in filled if i['evidence']), 'key': key_name or '—'}
    lines += [f'## {f.name}', f'{len(text):,} characters · language: {lang[1]} · assays: {", ".join(a["assay_name"] for a in meta["assays"]) or "none"} · '
              f'sites: {len(meta["sites"])} · rounds: {len(meta["rounds"])} · fields filled: {len(filled)} ({row["sourced"]} with a source sentence)', '']
    if key_name:
        k = keys[key_name]; res = {'right': [], 'wrong': [], 'missed': []}
        for term, answers in k['project'].items():
            ours = proj.get(term, [])
            for ans in answers:
                for part in [p.strip() for p in ans.split('|')] if term not in ('pcr_primer_forward', 'pcr_primer_reverse') else [ans]:
                    if not part: continue
                    flag = 'check' if status.get(term) == 'check' else 'ok'
                    if any(same(part, o) for o in ours): res['right'].append(term); auto['right_' + flag] += 1
                    elif ours: res['wrong'].append(f'{term}: expected "{part[:40]}", got "{ours[0][:40]}"' + (' [flagged]' if flag == 'check' else '')); auto['wrong_' + flag] += 1
                    else: res['missed'].append(f'{term} ("{part[:40]}")')
        for term, answers in k['sample'].items():
            o = samp.get(term)
            ans = answers[0]
            flag = 'check' if status.get(term) == 'check' else 'ok'
            if o and same(ans, o): res['right'].append(term); auto['right_' + flag] += 1
            elif o: res['wrong'].append(f'{term}: expected "{ans[:40]}", got "{o[:40]}"' + (' [flagged]' if flag == 'check' else '')); auto['wrong_' + flag] += 1
            else: res['missed'].append(f'{term} ("{ans[:40]}")')
        n = sum(len(v) for v in res.values())
        for x in res: totals[x] += len(res[x])
        row.update(right=len(res['right']), wrong=len(res['wrong']), missed=len(res['missed']))
        lines += [f'Answer key: **{key_name}** ({k["kind"]}) — right {len(res["right"])}/{n}, wrong {len(res["wrong"])}, missed {len(res["missed"])}', '']
        if res['wrong']: lines += ['Wrong:'] + [f'- {x}' for x in res['wrong']] + ['']
        if res['missed']: lines += ['Missed:'] + [f'- {x}' for x in res['missed']] + ['']
    for note in meta.get('notes') or []:
        lines.append(f'- note: {note}')
    lines.append('')
    summary_rows.append(row)

scored = totals['right'] + totals['wrong'] + totals['missed']
head = ['', '## Summary', '', '| Report | Language | Assays | Sites | Fields filled | Accepted automatically | Needs a look | Answer key | Right | Wrong | Missed |',
        '|---|---|---|---|---|---|---|---|---|---|---|']
for r in summary_rows:
    head.append(f"| {r['file'][:45]} | {r['language']} | {r['assays']} | {r['sites']} | {r['filled']} | {r['auto']} | {r['check']} | {r['key']} | "
                f"{r.get('right', '')} | {r.get('wrong', '')} | {r.get('missed', '')} |")
if scored:
    head += ['', f"**Scored fields overall: right {totals['right']}/{scored} ({100 * totals['right'] // scored}%), "
                 f"wrong {totals['wrong']}, missed {totals['missed']}.**",
             '', f"Automatic review: of the values accepted automatically, {auto['right_ok']} were right and {auto['wrong_ok']} wrong; "
                 f"of the values flagged for a look, {auto['right_check']} were right and {auto['wrong_check']} wrong."]
(HERE / 'results.md').write_text('\n'.join(lines[:4] + head + [''] + lines[4:]), encoding='utf-8')
with open(HERE / 'values.csv', 'w', newline='', encoding='utf-8') as fh:
    csv.writer(fh).writerows([['report', 'field', 'value', 'source sentence']] + value_rows)
print('\n'.join(head))
print(f'\nDetails: {HERE / "results.md"} · values with sources: {HERE / "values.csv"}')
