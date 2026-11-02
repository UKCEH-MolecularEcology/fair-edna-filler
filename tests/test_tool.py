#!/usr/bin/env python3
"""End-to-end checks for the FAIRe Metadata Filler (standard library only).

    python3 tests/test_tool.py                       # test with the bundled template
    python3 tests/test_tool.py path/to/template.xlsx # test a different / newer template

Fails (exit code 1) if a template no longer contains terms the tool fills,
if the filled workbook breaks any FAIRe rule we check, or if index.html is
out of date with server.py.
"""
import re, sys, tempfile, zipfile, posixpath
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import server  # noqa: E402

failures = []


def check(cond, msg):
    print(('  ok    ' if cond else '  FAIL  ') + msg)
    if not cond: failures.append(msg)


def sheet_rows(pkg, ss, paths, name):
    return server.parse_rows(pkg[paths[name]].decode('utf-8'))


def row_values(row, ss):
    return {c: server.cell_text(x, ss) for c, x in row['cells'].items()}


template = Path(sys.argv[1]) if len(sys.argv) > 1 else server.TEMPLATE_PATH
print(f'Template: {template.name}')

# ── 1. index.html must be built from the current server.py ──────────────────
sys.path.insert(0, str(ROOT))
import build_standalone  # noqa: E402
check((ROOT / 'index.html').read_text(encoding='utf-8') == build_standalone.build(),
      'index.html is up to date with server.py (run: python3 build_standalone.py)')

# ── 2. the template still has every term the tool writes ────────────────────
with zipfile.ZipFile(template) as z:
    tpl = {n: z.read(n) for n in z.namelist()}
ss = server.shared_strings(tpl)
paths = server.sheet_paths(tpl)
for sheet in ('projectMetadata', 'sampleMetadata', 'experimentRunMetadata', 'Drop-down values'):
    check(sheet in paths, f'template has sheet {sheet}')
if failures: sys.exit(1)

meta = server.parse_report_regex(server.extract_report_text(ROOT / 'tests' / 'sample_report.txt'))
project_terms = {row_values(r, ss).get('C') for n, r in sheet_rows(tpl, ss, paths, 'projectMetadata').items() if n > 1}
wanted = set(server.project_values(meta, '1')) | server.ASSAY_TERMS
missing = sorted(t for t in wanted if t not in project_terms)
check(not missing, f'projectMetadata contains all {len(wanted)} terms the tool fills' + (f' – missing: {missing}' if missing else ''))
samples = server.build_samples(meta)
for sheet, keys in (('sampleMetadata', {k for rec, _, _ in samples for k in rec}),
                    ('experimentRunMetadata', {'samp_name', 'assay_name', 'lib_id', 'seq_run_id', 'lib_conc_meth', 'phix_perc'})):
    hdr = set(row_values(sheet_rows(tpl, ss, paths, sheet)[3], ss).values())
    gone = sorted(k for k in keys if k not in hdr and k != 'org_note')
    check(not gone, f'{sheet} header contains every column the tool fills' + (f' – missing: {gone}' if gone else ''))

# ── 3. extraction from the sample report ────────────────────────────────────
check(meta['project_title'] == 'Example eDNA survey of two upland lakes for automated testing', 'project title extracted')
check([s['code'] for s in meta['sites']] == ['LALP', 'LBET'], 'two lakes extracted from the site table')
check(len(meta['rounds']) == 2 and meta['rounds'][0]['iso_date'] == '2024-06-03/2024-06-09', 'two sampling rounds with ISO dates')
check(sorted(a['marker'] for a in meta['assays']) == ['12S', 'COI'], 'assays 12S and COI detected')
coi = next(a for a in meta['assays'] if a['marker'] == 'COI')
check((coi.get('primer_name_fwd'), coi.get('primer_name_rev')) == ('fwhF2', 'fwhR2n'), 'COI primer names extracted')
check('doi.org/10.3897/mbmg.1.14625' in (coi.get('primer_ref_fwd') or ''), 'COI primer reference resolved to its DOI')
check(coi.get('amplificationReactionVolume') == '25' and coi.get('pcr_dna_vol') == '2', 'PCR volumes computed (25 µL, 2 µL template)')
check(meta.get('ext_kit', '').startswith('DNeasy PowerWater Kit'), 'extraction kit extracted')

# ── 4. fill the template and check the result ───────────────────────────────
with tempfile.TemporaryDirectory() as tmp:
    out = Path(tmp) / 'filled.xlsx'
    summary = server.fill_workbook(template, meta, 'sample_report.txt', out)
    with zipfile.ZipFile(out) as z:
        pkg = {n: z.read(n) for n in z.namelist()}
        check(z.testzip() is None, 'output is a valid zip/xlsx package')

bad_xml = []
for n, data in pkg.items():
    if n.endswith('.xml') or n.endswith('.rels'):
        try: ET.fromstring(data)
        except ET.ParseError: bad_xml.append(n)
check(not bad_xml, 'every XML part parses' + (f' – broken: {bad_xml}' if bad_xml else ''))

dangling = []
for n in pkg:
    if n.startswith('xl/worksheets/_rels/'):
        for tgt in re.findall(r'Target="([^"]+)"', pkg[n].decode()):
            if tgt.startswith('http'): continue
            p = tgt.lstrip('/') if tgt.startswith('/') else posixpath.normpath(posixpath.join('xl/worksheets', tgt))
            if p not in pkg: dangling.append(p)
check(not dangling, 'no relationships point to missing parts' + (f' – {dangling}' if dangling else ''))

ss = server.shared_strings(pkg)
paths = server.sheet_paths(pkg)
vocab = server.load_vocab(pkg, paths, ss)
MISSING = re.compile(r'^(missing:|not applicable)')


def vocab_ok(term, value):
    if term not in vocab or MISSING.match(value) or value.startswith('other:'): return True
    ftype, opts = vocab[term]
    return value in opts


# expected sizes: 2 lakes × 2 rounds × (5 samples + 1 field blank) + 2 extraction blanks + 2 PCR blanks × 2 assays
check(summary['sampleRows'] == 2 * 2 * 6 + 2 + 4, f'sampleMetadata rows = 30 (got {summary["sampleRows"]})')
check(summary['experimentRows'] == 26 * 2 + 4, f'experimentRunMetadata rows = 56 (got {summary["experimentRows"]})')

proj = sheet_rows(pkg, ss, paths, 'projectMetadata')
empty_m, vocab_bad = [], []
for n, r in proj.items():
    if n < 2: continue
    v = row_values(r, ss)
    term, req = v.get('C', ''), v.get('A', '')
    vals = [x for c, x in v.items() if server.col_num(c) >= 4 and x]
    if req == 'M' and not vals: empty_m.append(term)
    vocab_bad += [f'{term}={x}' for x in vals if not vocab_ok(term, x)]
check(not empty_m, 'every mandatory projectMetadata term has a value' + (f' – empty: {empty_m}' if empty_m else ''))

ids = {}
for sheet in ('sampleMetadata', 'experimentRunMetadata'):
    rows = sheet_rows(pkg, ss, paths, sheet)
    hdr = row_values(rows[3], ss)
    req = row_values(rows[1], ss)
    seen, empty_m2 = [], set()
    for n, r in rows.items():
        if n <= 3: continue
        v = row_values(r, ss)
        for c, term in hdr.items():
            if req.get(c) == 'M' and not v.get(c): empty_m2.add(term)
            if v.get(c): vocab_bad += [f'{sheet}.{term}={v[c]}'] if not vocab_ok(term, v[c]) else []
        key_col = next((c for c, t in hdr.items() if t == ('samp_name' if sheet == 'sampleMetadata' else 'lib_id')), None)
        if key_col: seen.append(v.get(key_col))
        if sheet == 'experimentRunMetadata': ids.setdefault('run_samples', set()).add(v.get('A'))
    ids[sheet] = seen
    check(not empty_m2, f'every mandatory {sheet} column is filled' + (f' – empty: {sorted(empty_m2)}' if empty_m2 else ''))
    check(bool(seen) and len(seen) == len(set(seen)), f'{sheet} IDs are present and unique')
check(ids['run_samples'] <= set(ids['sampleMetadata']), 'every run row links to a sample in sampleMetadata')
check(not vocab_bad, 'all controlled-vocabulary values are valid' + (f' – {sorted(set(vocab_bad))[:5]}' if vocab_bad else ''))

# ── 5. automatic checks ─────────────────────────────────────────────────────
import json, threading
from http.server import BaseHTTPRequestHandler, HTTPServer

text = server.extract_report_text(ROOT / 'tests' / 'sample_report.txt')
meta = server.prepare(text)
coi = next(a for a in meta['assays'] if a['marker'] == 'COI')
check(coi.get('primer_fwd') == 'GGDACWGGWTGAACWGTWTAYCCHCC', 'primer library fills the fwhF2 sequence the report only names')
items = server.review(meta, template)
by_id = {it['id']: it for it in items}
ci = next(i for i, a in enumerate(meta['assays']) if a['marker'] == 'COI')
check(by_id[f'assay:{ci}:pcr_primer_forward']['status'] == 'check', 'a library-filled sequence is marked for checking')
check(by_id['project:institution']['status'] == 'ok', 'a value with a source sentence is accepted automatically')
counts = server.review_summary(items)
check(counts['ok'] > counts['check'] > 0, f'most values accepted automatically ({counts["ok"]} ok, {counts["check"]} to check)')
check(server.check_value('numeric', '25 µL') == ('25', None) and server.check_value('fixed format (PnYnMnWnDTnHnMnS)', '24 hours') == ('PT24H', None)
      and server.check_value('fixed format (yyyy-mm-dd)', '14/07/2025') == ('2025-07-14', None)
      and server.check_value('integer', 'three')[1], 'validation fixes units, durations and dates, and flags non-numbers')

edits = {'project:license': 'CC-BY 4.0', f'assay:{ci}:commercial_mm': 'Q5 master mix (NEB)'}
prof = server.profile_from_review(meta, edits, 'test lab')
again = server.prepare(text, profile=prof)
items2 = {it['id']: it for it in server.review(again, template)}
check(items2['project:license']['value'] == 'CC-BY 4.0' and items2['project:license']['status'] == 'ok'
      and items2[f'assay:{ci}:commercial_mm']['value'] == 'Q5 master mix (NEB)', 'a saved profile re-applies corrections to the next report')

reply = {'project_title': 'Example eDNA survey of two upland lakes for automated testing', 'filter_diameter': 47,
         'project_contact': 'nobody@example.org', 'sequencing_location': 'UKCEH Wallingford',
         'assays': [{'marker': 'COI', 'annealingTemp': 50, 'pcr_dna_vol': 2}],
         '_evidence': {'project_title': 'Example eDNA survey of two upland lakes for automated testing',
                       'filter_diameter': 'vacuum filtered within 24 hours using 47 mm diameter',
                       'project_contact': 'contact Jane at nobody@example.org',
                       'sequencing_location': 'prior to DNA extraction and sequencing at UKCEH Wallingford',
                       'COI.pcr_dna_vol': 'prior to the addition of 2 µL of template DNA'}}


class FakeOllama(BaseHTTPRequestHandler):
    def do_POST(self):
        self.rfile.read(int(self.headers['Content-Length']))
        body = json.dumps({'message': {'content': json.dumps(reply)}}).encode()
        self.send_response(200); self.send_header('Content-Length', str(len(body))); self.end_headers(); self.wfile.write(body)

    def log_message(self, *a): pass


srv = HTTPServer(('127.0.0.1', 0), FakeOllama)
threading.Thread(target=srv.serve_forever, daemon=True).start()
m3 = server.parse_report_regex(text)
server.local_ai_opinion(m3, text, 'fake', f'http://127.0.0.1:{srv.server_address[1]}')
srv.shutdown()
r3 = {it['id']: it for it in server.review(m3, template)}
check(r3['project:sequencing_location']['value'] == 'UKCEH Wallingford' and r3['project:sequencing_location']['status'] == 'check',
      'local AI fills a gap, marked for checking')
check(r3['sample:filter_diameter']['status'] == 'ok' and 'agrees' in ' '.join(r3['sample:filter_diameter']['reasons']),
      'agreement with the local AI is recorded')
check(not r3['project:project_contact']['value'], 'a local-AI value without a real quote from the report is discarded')
try:
    server.parse_with_ollama(text, 'x', 'http://example.org:11434'); check(False, 'refuses to send reports to another computer')
except RuntimeError:
    check(True, 'refuses to send reports to another computer')

# ── 6. title, authors and contact from the first page of a paper or report ──
paper = """Mol Ecol Resour. 2024;24:e13903.  | 1 of 15
https://doi.org/10.1111/1755-0998.13903
RESOURCE ARTICLE
Evaluating the use of lake sedimentary DNA in palaeolimnology:
A comparison with long-term monitoring
Amy C. Thorpe1,2 | Eleanor B. Mackay3 | Daniel S. Read1
1UK Centre for Ecology & Hydrology (UKCEH), Wallingford, UK
Correspondence
Amy C. Thorpe (amytho@ceh.ac.uk)
"""
fm = server.front_matter(paper)
check(fm.get('project_title') == 'Evaluating the use of lake sedimentary DNA in palaeolimnology: A comparison with long-term monitoring',
      'paper title found under journal header lines')
check(fm.get('recorded_by') == 'Amy C. Thorpe | Eleanor B. Mackay | Daniel S. Read', 'author names found (superscripts removed)')
check('amytho@ceh.ac.uk' in (fm.get('project_contact') or '') and 'Centre for Ecology' in (fm.get('institution') or ''),
      'corresponding e-mail and first affiliation found')
report = """Page 1 of 39  [Report title and catalogue code]
GCN sample analysis via single
species assay and metabarcoding
Extraction and analysis of Great Crested Newt eDNA samples
April 2024
Natural England Commissioned Report NECR534

Page 2 of 39 GCN sample analysis via single species assay and metabarcoding
NECR534
Report details
Author(s)
Helen C. Rees, Steven D. Kane
Contractor
RSK ADAS Ltd.
Keywords
eDNA
"""
fm = server.front_matter(report)
check(fm.get('project_title') == 'GCN sample analysis via single species assay and metabarcoding', 'report title found; subtitle left out')
check(fm.get('recorded_by') == 'Helen C. Rees | Steven D. Kane' and fm.get('institution') == 'RSK ADAS Ltd',
      'authors and contractor read from the report details page')

# ── 7. sampling design described in words, and a user sample list ─────────
design = """Methods
Water samples were collected to target autumn spawners (7th November to 5th December 2025) and spring spawners (28th January to 3rd March 2026).
Potential autumn spawning sites covered 1) the west shore (NBASW, N=9), 2) North Point (NTH, N=3 and NTHS, N=1). Each of these sites was sampled five times within the autumn sampling period, with the exception of “NTHS” which was sampled nine times. Sites targeted for spring spawners were Holbeck Point (HOLP, N=4). Each of these sites was sampled five times during the spring spawning period. Seven samples were collected at each of the two additional lakes: Rydal Water (1st December 2025), Goat’s Water (5th February 2026). A single field blank was carried into the field on each day of sampling. The additional lakes (RYDA and GOAT) were sampled from the shore.
""" + "PCR and sequencing details. " * 80
plan = server.sample_plan(design, server.norm(design))
by = {g['code']: g for g in plan}
check(sorted(by) == ['GOAT', 'HOLP', 'NBASW', 'NTH', 'NTHS', 'RYDA'], 'sites read from "(CODE, N=x)" and "N samples at each of the lakes"')
check(by['NBASW']['points'] * by['NBASW']['visits'] == 45 and by['NTHS']['visits'] == 9 and by['HOLP']['date'] == '2026-01-28/2026-03-03'
      and by['GOAT']['date'] == '2026-02-05', 'points × visits, exceptions and campaign dates')
p1 = server.sample_plan('Five replicate water samples were collected at 37 sites across three catchments.', 'Five replicate water samples were collected at 37 sites across three catchments.')
p2 = server.sample_plan('x', 'Sample collection. A total of 450 river biofilm samples were collected over three years from 146 sites.')
check(len(p1) == 37 and p1[0]['points'] == 5 and p2 and p2[0]['points'] == 450,
      'plain sample counts ("5 samples at 37 sites", "a total of 450 samples") give placeholder sample rows')
with tempfile.TemporaryDirectory() as tmp:
    sl = Path(tmp) / 'samples.csv'
    sl.write_text('sample,site,date,latitude,longitude\nA1,Lake A,07/11/2025,54.1,-2.9\nA_FB,Lake A,07/11/2025,,\nEXT_BLANK_1,,,,\n', encoding='utf-8')
    m4 = server.prepare(text, samples=sl)
    rows4 = {rec['samp_name']: (rec, ctl) for rec, _, ctl in server.build_samples(m4)}
check(set(rows4) == {'A1', 'A_FB', 'EXT_BLANK_1'} and rows4['A1'][0]['eventDate'] == '2025-11-07'
      and rows4['A_FB'][0]['neg_cont_type'] == 'field negative' and rows4['EXT_BLANK_1'][0]['neg_cont_type'] == 'extraction negative'
      and rows4['A1'][0]['decimalLatitude'] == '54.1', 'a sample list gives one row per sample, with dates, coordinates and blanks')

# ── 8. extra sources (protocol, dataset record, GitHub, SRA) – offline, with canned responses ──
O = server.FETCH_OVERRIDES
O['https://api.datacite.org/dois/10.5281/zenodo.111?affiliation=true'] = json.dumps({'data': {'attributes': {
    'creators': [{'givenName': 'Ann', 'familyName': 'Lee', 'nameIdentifiers': [{'nameIdentifier': '0000-0001-0000-0001', 'nameIdentifierScheme': 'ORCID'}],
                  'affiliation': [{'name': 'University of Testing', 'affiliationIdentifier': 'https://ror.org/000000000'}]}],
    'rightsList': [{'rights': 'Creative Commons Attribution 4.0 International', 'rightsUri': 'https://creativecommons.org/licenses/by/4.0/'}]}}}).encode()
O['https://zenodo.org/api/records/111/files'] = json.dumps({'entries': []}).encode()
G = 'https://raw.githubusercontent.com/ann/edna/main/'
O['https://api.github.com/repos/ann/edna'] = b'{"default_branch": "main"}'
O['https://api.github.com/repos/ann/edna/git/trees/main?recursive=1'] = json.dumps({'tree': [
    {'path': p, 'type': 'blob', 'size': 10} for p in ('README.md', 'cooc_meta.csv', 'samples.csv', 'reads.csv', 'blanks.csv')]}).encode()
O[G + 'cooc_meta.csv'] = b',sample,type,habitat\n1,A1,W,B\n2,A2,W,B\n3,CAM1_2023-05-01,Cam,F\n4,CAM1_2023-05-02,Cam,F\n5,CAM1_2023-05-03,Cam,F\n'
O[G + 'README.md'] = b"Blanks: NEG (PCR negative), 'EB' (e.g. XEB1) are extraction blanks, 'BK' is a field blank."
O[G + 'samples.csv'] = b'sample,type,site,date,time,volume_filtered,notes\nA1,W,Lake A,07/06/2023,10:15,800,clear\nA2,W,Lake A,07/06/2023,10:30,750,\n'
O[G + 'reads.csv'] = (b'"",A1,A2,ABK1,AEB1,NEG01,taxonomy\nRana_temporaria,10,0,0,0,0,"d__Eukaryota; p__Chordata; c__Amphibia; o__Anura; f__Ranidae; g__Rana; s__Rana_temporaria"\n'
                      b'Salmo_trutta,500,90,3,0,0,"d__Eukaryota; p__Chordata; c__Actinopteri; o__Salmoniformes; f__Salmonidae; g__Salmo; s__Salmo_trutta"\n')
O[G + 'blanks.csv'] = b'blank,sample\nABK1,A1\nABK1,A2\nAEB1,A1\nNEG01,A1\n'
O['https://www.ebi.ac.uk/ena/portal/api/search?result=read_run&query=study_accession=PRJNA999&fields=run_accession,library_name,sample_alias,'
  'sample_accession,instrument_model,library_layout,read_count,fastq_ftp,fastq_md5'] = (
    b'run_accession\tlibrary_name\tsample_alias\tsample_accession\tinstrument_model\tlibrary_layout\tread_count\tfastq_ftp\tfastq_md5\n'
    b'SRR1\tL1\tA1\tS1\tIllumina MiSeq\tPAIRED\t1000\tftp/x/SRR1_1.fastq.gz;ftp/x/SRR1_2.fastq.gz\taaa;bbb\n')
with tempfile.TemporaryDirectory() as tmp:
    proto = Path(tmp) / 'protocol.txt'
    proto.write_text("""DOI https://dx.doi.org/10.17504/protocols.io.test123/v1
Section 7 Library Preparation
22 Step 1: first PCR PCR1
1. Use primers 12S-V5-F (5'-ACTGGGATTAGATACCCC-3') and 12SV5-R (5'-TAGAACAGGCTCCTCTAG-3').
2. Prepare PCR1 reactions in triplicate for each sample:
12.5 µL Q5 Master Mix
7 µL molecular-grade water
1.5 µL of each 10 µM tagged primer
2 µL template DNA
3. Run PCR1 with the following conditions:
98°C for 5 mins
35 cycles of 98°C for 10s, 58°C for 20s and 72°C for 30s
72°C for 7 minutes
Step 3: PCR2 - Indexing
1. Prepare duplicate 50 µL PCR2 reactions:
25 µL Q5 Master Mix
4 µL cleaned template DNA
""", encoding='utf-8')
    m5 = server.prepare(text, sources=[str(proto), '10.5281/zenodo.111', 'https://github.com/ann/edna', 'PRJNA999'])
    wb5 = Path(tmp) / 'multi.xlsx'
    s5 = server.fill_workbook(template, m5, 'x', wb5)
r5 = {it['id']: it for it in server.review(m5, template)}
a12 = next(i for i, a in enumerate(m5['assays']) if a['marker'] == '12S')
check(r5[f'assay:{a12}:annealingTemp']['value'] == '58' and r5[f'assay:{a12}:pcr_cycles']['value'] == '35'
      and r5[f'assay:{a12}:amplificationReactionVolume']['value'] == '24.5' and 'protocols.io' in ' '.join(r5[f'assay:{a12}:annealingTemp']['reasons']),
      'protocol: PCR recipe and cycling read from step lists, with the protocol as source')
check(r5['project:nucl_acid_amp']['value'] if 'project:nucl_acid_amp' in r5 else r5[f'assay:{a12}:nucl_acid_amp']['value'] == 'https://doi.org/10.17504/protocols.io.test123/v1',
      'protocol DOI used for nucl_acid_amp')
check(r5['project:recordedByID']['value'] == 'https://orcid.org/0000-0001-0000-0001' and 'Creative Commons' in r5['project:license']['value'],
      'dataset record: ORCIDs and licence')
names = {rec['samp_name']: rec for rec, _, _ in server.build_samples(m5)}
check(set(names) == {'A1', 'A2', 'ABK1', 'AEB1', 'NEG01'} and names['ABK1']['neg_cont_type'] == 'field negative'
      and names['AEB1']['neg_cont_type'] == 'extraction negative' and names['NEG01']['neg_cont_type'] == 'PCR negative',
      'GitHub: samples from the data table, controls named via the README legend')
check(names['A1']['rel_cont_id'] == 'ABK1 | AEB1 | NEG01' and names['A1']['eventDate'] == '2023-06-07T10:15' and names['A1']['_extra'].get('sample_type') == 'W',
      'blank links, date + time, unmatched columns kept as user-defined columns')
check(names['A1']['_extra'].get('habitat') == 'B' and names['A1'].get('samp_vol_we_dna_ext') == '800',
      'the richest sample table is used, with gaps filled from the other tables')
check(names['A1']['_run'].get('filename') == 'SRR1_1.fastq.gz' and names['A1']['_run'].get('checksum_filename') == 'aaa'
      and names['A1']['_run'].get('output_read_count') == '510', 'SRA run files and read counts linked to their sample')
check(s5['taxa'].get('taxaRaw') == 2, 'taxaRaw filled from the read table')
check(server.find_links('Data on Zenodo (https:// doi. org/ 10. 5281/ zenodo. 15723454) and GitHub (https:// github. com/ ann/ ed na ). BioProject PRJNA1345715.')
      == ['10.5281/zenodo.15723454', 'https://github.com/ann/edna', 'PRJNA1345715'], 'links found in PDF text with broken spacing')
# broken or awkward links must never stop the run
check(server.quote_url('https://raw.githubusercontent.com/a/b/main/Reference database/tax map.txt')
      == 'https://raw.githubusercontent.com/a/b/main/Reference%20database/tax%20map.txt'
      and server.quote_url('https://x.org/a%20b\n') == 'https://x.org/a%20b', 'links with spaces / line breaks are encoded safely')
O['https://api.github.com/repos/ann/broken'] = b'{"default_branch": "main"}'
O['https://api.github.com/repos/ann/broken/git/trees/main?recursive=1'] = json.dumps({'tree': [
    {'path': 'Reference database/map.txt', 'type': 'blob', 'size': 10}, {'path': 'gone.csv', 'type': 'blob', 'size': 10}]}).encode()
O['https://raw.githubusercontent.com/ann/broken/main/Reference database/map.txt'] = b'id\ttaxid\nA\t1\n'
O['https://raw.githubusercontent.com/ann/broken/main/gone.csv'] = OSError('HTTP Error 404: Not Found')
O['https://api.datacite.org/dois/10.5281/zenodo.404?affiliation=true'] = OSError('HTTP Error 404: Not Found')
m6 = server.prepare(text, sources=['https://github.com/ann/broken', '10.5281/zenodo.404', 'not-a-link'])
check(any('could not read gone.csv' in n for n in m6['notes']) and any('zenodo' in n.lower() and '404' in n for n in m6['notes'])
      and any('not-a-link' in n for n in m6['notes']), 'a broken link becomes a note, the rest still runs')
O.clear()

print()
if failures:
    print(f'{len(failures)} check(s) failed.')
    sys.exit(1)
print('All checks passed.')
