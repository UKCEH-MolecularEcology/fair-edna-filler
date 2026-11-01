#!/usr/bin/env python3
"""
FAIRe Metadata Filler.

Reads an eDNA project report (.doc, .docx, .pdf, .txt), extracts methods and
sampling metadata, and writes them into the FAIRe full template workbook.

    python3 server.py            # web app at http://127.0.0.1:8787

Environment variables
    ANTHROPIC_API_KEY   enables Claude-assisted extraction (optional)
    FAIRE_CLAUDE_MODEL  Claude model id (default: claude-haiku-5-5)
    FAIRE_TOOL_HOST     bind address (default 127.0.0.1)
    FAIRE_TOOL_PORT     port (default 8787)

The same functions are used by the Colab notebook:
    extract_report_text(path) -> str
    parse_report_regex(text)  -> metadata dict     (rule-based)
    parse_with_claude(text, key) -> dict ; facts_from_claude(dict) -> metadata dict
    fill_workbook(template, metadata, report_name, output) -> summary dict
"""
from __future__ import annotations

import html, json, os, re, shutil, struct, subprocess, tempfile, time, uuid, zipfile
import xml.etree.ElementTree as ET
from datetime import date, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

APP_DIR = Path(__file__).resolve().parent
STATIC_DIR = APP_DIR / 'static'
OUTPUT_DIR = APP_DIR / 'outputs'
UPLOAD_DIR = APP_DIR / 'uploads'


# ════════════════════════════════════════════════════════════════════════════
#  1. Reading reports
# ════════════════════════════════════════════════════════════════════════════

def read_doc_text(data: bytes) -> str:
    """Extract text from a Word 97-2003 .doc (CFB container + piece table). Pure Python."""
    if data[:8] != b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1':
        raise ValueError('Not a Word 97-2003 .doc file (no OLE header).')
    sec_size  = 1 << struct.unpack_from('<H', data, 0x1E)[0]
    mini_size = 1 << struct.unpack_from('<H', data, 0x20)[0]
    first_dir = struct.unpack_from('<I', data, 0x30)[0]
    cutoff    = struct.unpack_from('<I', data, 0x38)[0]
    first_mf  = struct.unpack_from('<I', data, 0x3C)[0]
    first_dif = struct.unpack_from('<I', data, 0x44)[0]
    n_dif     = struct.unpack_from('<I', data, 0x48)[0]
    sec = lambda n: 512 + n * sec_size if sec_size == 512 else (n + 1) * sec_size
    # DIFAT -> FAT sector list
    difat = [x for x in struct.unpack_from('<109I', data, 0x4C) if x < 0xFFFFFFFA]
    d = first_dif
    for _ in range(n_dif):
        if d >= 0xFFFFFFFA: break
        vals = struct.unpack_from(f'<{sec_size//4}I', data, sec(d))
        difat += [x for x in vals[:-1] if x < 0xFFFFFFFA]; d = vals[-1]
    fat = []
    for s in difat:
        fat += struct.unpack_from(f'<{sec_size//4}I', data, sec(s))
    def chain(start, table):
        out, s, guard = [], start, 0
        while s < 0xFFFFFFFA and guard < len(table) + 1:
            out.append(s); s = table[s]; guard += 1
        return out
    def read_big(start, size):
        b = b''.join(data[sec(s):sec(s) + sec_size] for s in chain(start, fat))
        return b[:size]
    dir_bytes = read_big(first_dir, 1 << 30)
    entries = {}
    for i in range(len(dir_bytes) // 128):
        e = dir_bytes[i*128:(i+1)*128]
        nlen = struct.unpack_from('<H', e, 64)[0]
        if nlen < 2: continue
        name = e[:nlen-2].decode('utf-16-le', 'replace')
        etype = e[66]; start = struct.unpack_from('<I', e, 116)[0]; size = struct.unpack_from('<I', e, 120)[0]
        entries[name] = (etype, start, size)
    root = entries.get('Root Entry')
    minifat = []
    if first_mf < 0xFFFFFFFA:
        for s in chain(first_mf, fat):
            minifat += struct.unpack_from(f'<{sec_size//4}I', data, sec(s))
    ministream = read_big(root[1], root[2]) if root else b''
    def stream(name):
        if name not in entries: return None
        _, start, size = entries[name]
        if size < cutoff:
            return b''.join(ministream[s*mini_size:(s+1)*mini_size] for s in chain(start, minifat))[:size]
        return read_big(start, size)
    wd = stream('WordDocument')
    if wd is None: raise ValueError('No WordDocument stream found.')
    flags = struct.unpack_from('<H', wd, 0x0A)[0]
    table = stream('1Table' if flags & 0x0200 else '0Table')
    # FibRgFcLcb97: fcClx/lcbClx are entry 33
    csw = struct.unpack_from('<H', wd, 32)[0]
    p = 34 + csw * 2
    cslw = struct.unpack_from('<H', wd, p)[0]
    p += 2 + cslw * 4 + 2           # skip rgLw and cbRgFcLcb
    fc_clx, lcb_clx = struct.unpack_from('<II', wd, p + 33 * 8)
    clx = table[fc_clx:fc_clx + lcb_clx]
    i = 0
    while i < len(clx) and clx[i] == 0x01:           # skip Prc entries
        i += 3 + struct.unpack_from('<H', clx, i + 1)[0]
    if i >= len(clx) or clx[i] != 0x02: raise ValueError('Could not find the document text table.')
    lcb = struct.unpack_from('<I', clx, i + 1)[0]; plc = clx[i+5:i+5+lcb]
    n = (lcb - 4) // 12
    cps = struct.unpack_from(f'<{n+1}I', plc, 0)
    parts = []
    for k in range(n):
        fc = struct.unpack_from('<I', plc, (n+1)*4 + k*8 + 2)[0]
        count = cps[k+1] - cps[k]
        if fc & 0x40000000:
            off = (fc & 0x3FFFFFFF) // 2
            parts.append(wd[off:off+count].decode('cp1252', 'replace'))
        else:
            parts.append(wd[fc:fc+2*count].decode('utf-16-le', 'replace'))
    return clean_word_text(''.join(parts))

def clean_word_text(t: str) -> str:
    t = re.sub(r'\x13[^\x13\x14\x15]*\x14', '', t)   # drop field instructions (e.g. FORMTEXT)
    t = re.sub(r'[\x13\x14\x15]', '', t)
    t = t.replace('\r', '\n').replace('\x0b', '\n').replace('\x0c', '\n').replace('\x07', '\t')
    t = re.sub(r'[\x00-\x08\x0e-\x1f]', '', t)
    return t

def extract_report_text(path) -> str:
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix in ('.txt', '.md'):
        return path.read_text(errors='replace')
    if suffix == '.docx':
        return _extract_docx(path)
    if suffix == '.pdf':
        return _extract_pdf(path)
    if suffix == '.doc':
        data = path.read_bytes()
        try:
            return read_doc_text(data)
        except Exception as first_error:
            text = _convert_with_office(path)
            if text: return text
            raise ValueError(f'Could not read this .doc file ({first_error}). Please save it as .docx and try again.')
    raise ValueError(f'Unsupported file type "{suffix}". Use .doc, .docx, .pdf or .txt.')


def _extract_docx(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read('word/document.xml'))
    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    out = []
    for p in root.iter(W + 'p'):
        parts = []
        for el in p.iter():
            if el.tag == W + 't' and el.text: parts.append(el.text)
            elif el.tag == W + 'tab': parts.append('\t')
            elif el.tag in (W + 'br', W + 'cr'): parts.append('\n')
        out.append(''.join(parts))
    return '\n'.join(out)


def _extract_pdf(path: Path) -> str:
    try:
        import pdfplumber  # type: ignore
        with pdfplumber.open(path) as pdf:
            return '\n'.join(page.extract_text() or '' for page in pdf.pages)
    except ImportError:
        pass
    try:
        from pypdf import PdfReader  # type: ignore
        return '\n'.join(page.extract_text() or '' for page in PdfReader(str(path)).pages)
    except ImportError:
        raise ValueError('PDF support needs pdfplumber: pip install pdfplumber')


def _convert_with_office(path: Path):
    """Fallback for unusual .doc files: LibreOffice or macOS textutil."""
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if soffice:
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run([soffice, '--headless', '--convert-to', 'txt:Text', '--outdir', tmp, str(path)],
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False, timeout=120)
            out = Path(tmp) / (path.stem + '.txt')
            if out.exists(): return out.read_text(errors='replace')
    textutil = shutil.which('textutil')
    if textutil:
        r = subprocess.run([textutil, '-convert', 'txt', '-stdout', str(path)], stdout=subprocess.PIPE, check=False)
        if r.returncode == 0 and r.stdout: return r.stdout.decode('utf-8', errors='replace')
    return None



# ════════════════════════════════════════════════════════════════════════════
#  2. Rule-based extraction (generic patterns; nothing hard-coded per report)
# ════════════════════════════════════════════════════════════════════════════

MARKERS = ['16S', '18S', 'rbcL', 'COI', '12S', 'ITS', 'CytB', '28S', '23S', 'trnL', 'matK']
MARKER_RE = r'(?:16S|18S|rbcL|RBCL|COI|CO1|12S|ITS2?|CytB|cytb|28S|23S|trnL|matK)'
TARGET_GENE = {   # FAIRe target_gene controlled vocabulary
    '16S': '16S rRNA (SSU prokaryote)', '18S': '18S rRNA (SSU eukaryote)', '12S': '12S rRNA (SSU mitochondria)',
    'rbcL': 'rbcL', 'COI': 'COI', 'ITS': 'ITS', 'CytB': 'CytB', '28S': '28S rRNA (LSU eukaryote)',
    '23S': '23S rRNA (LSU prokaryote)', 'trnL': 'trnL', 'matK': 'matK',
}
TAXON_WORDS = [  # descriptor in text -> (targetTaxonomicAssay, scope word)
    (r'prokaryot\w*|bacteri\w*', 'Bacteria | Archaea'),
    (r'diatom\w*', 'Bacillariophyta'),
    (r'eukaryot\w*|protist\w*', 'Eukaryota'),
    (r'metazoan\w*|invertebrate\w*|macroinvertebrate\w*', 'Metazoa'),
    (r'fish', 'Actinopterygii'),
    (r'vertebrate\w*', 'Vertebrata'),
    (r'fung\w*', 'Fungi'),
    (r'plant\w*', 'Viridiplantae'),
]
NUMBER_WORDS = {w: i for i, w in enumerate(
    'zero one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen sixteen'
    ' seventeen eighteen nineteen twenty'.split())}
MONTHS = {m: i + 1 for i, m in enumerate(
    'january february march april may june july august september october november december'.split())}
MONTH_RE = r'(?:January|February|March|April|May|June|July|August|September|October|November|December)'
UK_REGIONS = {  # adjective/name -> INSDC region
    'cumbria': 'Cumbria', 'cumbrian': 'Cumbria', 'scotland': 'Scotland', 'scottish': 'Scotland',
    'wales': 'Wales', 'welsh': 'Wales', 'northern ireland': 'Northern Ireland', 'england': 'England',
}
INSTITUTIONS = {  # acronym -> (name, ROR)
    'UKCEH': ('UK Centre for Ecology & Hydrology', 'https://ror.org/00pggkr55'),
}


def norm(s: str) -> str:
    return re.sub(r'\s+', ' ', s or '').strip()


def num_word(w: str):
    w = w.strip().lower()
    return int(w) if w.isdigit() else NUMBER_WORDS.get(w)


def canon_marker(m: str) -> str:
    m = m.strip()
    up = m.upper()
    if up == 'RBCL': return 'rbcL'
    if up in ('COI', 'CO1'): return 'COI'
    if up.startswith('ITS'): return 'ITS'
    if up == 'CYTB': return 'CytB'
    return m


def markers_in(s: str) -> list[str]:
    out = []
    for m in re.findall(r'(?<![A-Za-z0-9])' + MARKER_RE + r'(?![A-Za-z0-9])', s):
        c = canon_marker(m)
        if c not in out: out.append(c)
    return out


def sentences(text: str) -> list[str]:
    t = norm(text)
    parts = re.split(r'(?<=[.!?])\s+(?=[A-Z(])', t)
    return [p for p in parts if p]


def parse_date(s: str, default_year=None):
    """'1st July 2025' / '14th July' / '19/2/26' / '2025-07-01' -> ISO date string or None."""
    s = s.strip()
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', s)
    if m: return s[:10]
    m = re.match(r'(\d{1,2})(?:st|nd|rd|th)?\s+(' + MONTH_RE + r')(?:\s+(\d{4}))?', s, re.I)
    if m:
        y = int(m.group(3)) if m.group(3) else default_year
        if not y: return None
        return f'{y:04d}-{MONTHS[m.group(2).lower()]:02d}-{int(m.group(1)):02d}'
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{2,4})', s)
    if m:
        y = int(m.group(3)); y = y + 2000 if y < 100 else y
        return f'{y:04d}-{int(m.group(2)):02d}-{int(m.group(1)):02d}'
    return None


def after_label(text: str, label: str, max_len=300):
    """First non-empty line/cell after a form label."""
    m = re.search(label + r'[ \t]*(?::)?[ \t]*\n?(.*?)(?:\n|$)', text, re.I)
    if not m: return None
    val = norm(m.group(1))
    if not val:  # value on a following line
        rest = text[m.end():]
        for line in rest.splitlines():
            if norm(line): val = norm(line); break
    return val[:max_len] or None


def find(pattern, text, group=1, flags=re.I):
    m = re.search(pattern, text, flags)
    return norm(m.group(group)) if m else None


def references(text: str) -> list[str]:
    """Reference list as one string per entry (handles entries wrapped over several lines)."""
    m = re.search(r'References to published material(.*)', text, re.S | re.I) or \
        re.search(r'\n\s*References\s*\n(.*)', text, re.S)
    block = m.group(1) if m else ''
    entries, cur = [], ''
    for line in block.splitlines():
        l = norm(line)
        if not l:
            if cur: entries.append(cur); cur = ''
            continue
        new = re.match(r"^[A-Z\u00c0-\u00dd][A-Za-z\u00c0-\u00ff'\-]+,\s+(?:[A-Z]\.|[A-Z][a-z]+)", l) or \
              re.match(r'^[A-Z][\w\-]+ (?:et al\.?|and [A-Z]|\(\d{4}\))', l)
        if new and cur: entries.append(cur); cur = l
        else: cur = (cur + ' ' + l).strip()
    if cur: entries.append(cur)
    return [e for e in entries if re.search(r'\(\d{4}\)|\d{4}\.', e) and len(e) > 30]


def ref_for(citation: str, refs: list[str]):
    """'Vamos et al. 2017' -> full reference line (with DOI) if present."""
    m = re.match(r'([A-Z][A-Za-zÀ-ÿ\-]+).*?(\d{4})', citation)
    if not m: return citation
    author, year = m.group(1), m.group(2)
    for r in refs:
        if r.startswith(author) and year in r[:120]:
            doi = re.search(r'https?://doi\.org/\S+', r)
            return f'{citation.strip()} {doi.group(0).rstrip(".")}' if doi else r
    return citation


# ── main extractor ───────────────────────────────────────────────────────────

def extract(text: str) -> dict:
    text = text.replace('\u00a0', ' ').replace('\ufeff', '')
    text = re.sub('[\ue000-\uf8ff]', '', text)                     # private-use glyphs (bullets)
    text = re.sub(r'(?<=[a-z])-\n(?=[a-z])', '-', text)              # "high-\nfidelity"
    refs = references(text)
    cut = re.search(r'\n\s*(?:\d+\.\s*)?(?:\t)?References(?: to published material)?\s*\n', text)
    text = text[:cut.start()] if cut else text            # body only from here on
    text = text.replace('\u03bc', '\u00b5')                # Greek mu -> micro sign
    flat = norm(text)
    sents = sentences(text)
    meta: dict = {'assays': [], 'sites': [], 'rounds': [], 'runs': [], 'notes': []}

    # ── project ──────────────────────────────────────────────────────────────
    tm = re.search(r'Project title[ \t]*:?[ \t]*\n?((?:[ \t]*\S.*\n?){1,3})', text, re.I)
    title_lines = []
    if tm:
        for line in tm.group(1).splitlines():
            if not norm(line) or re.match(r'\s*\d+\.', line) or re.match(r'\s*(Contractor|organisation)', line, re.I): break
            title_lines.append(norm(line))
    meta['project_title'] = ' '.join(title_lines) or find(r'^\s*(.{20,200})\n', text, flags=re.M)
    org = after_label(text, r'organisation\(s\)') or after_label(text, r'Contractor')
    for acr, (name, ror) in INSTITUTIONS.items():
        if (org and acr in org) or re.search(r'\b' + acr + r'\b', flat):
            meta['institution'], meta['institution_ror'] = name, ror
            break
    else:
        meta['institution'] = org
    meta['start_date'] = parse_date(after_label(text, r'start date') or '')
    meta['end_date'] = parse_date(after_label(text, r'end date') or '')
    if re.search(r'\bDefra\b', flat): meta['funder'] = 'Defra'

    # ── controls ─────────────────────────────────────────────────────────────
    fb = re.search(r'(\w+) field blanks? .{0,80}?(?:per|for each) (?:lake|site|pond|river|location)', flat, re.I)
    meta['field_blanks_per_site_round'] = num_word(fb.group(1)) if fb else (1 if re.search(r'field blank', flat, re.I) else 0)
    eb = re.search(r'(\w+) extraction (?:blanks?|negatives?|controls?) were included', flat, re.I)
    meta['extraction_blanks'] = num_word(eb.group(1)) if eb else 0
    pb = re.search(r'(\w+) PCR (?:blanks?|negatives?) were included for each primer pair', flat, re.I)
    meta['pcr_negatives_per_assay'] = num_word(pb.group(1)) if pb else 0
    meta['neg_controls'] = bool(meta['field_blanks_per_site_round'] or meta['extraction_blanks'] or
                                meta['pcr_negatives_per_assay'] or re.search(r'negative control', flat, re.I))
    meta['pos_controls'] = bool(re.search(r'positive control|mock communit', flat, re.I))

    # ── sterilisation ────────────────────────────────────────────────────────
    st = [s for s in sents if re.search(r'UV[- ]steril|bleach|sterile scalpel|decontaminat|DNA[- ]?away', s, re.I)]
    meta['sterilise_method'] = ' '.join(st[:2]) or None

    # ── sampling design ──────────────────────────────────────────────────────
    sm = re.search(r'(\w+)\s+(\d+(?:\.\d+)?)\s*(L|mL|ml|litre|liter)s?\s+water samples', flat, re.I)
    if sm:
        meta['samples_per_site_round'] = num_word(sm.group(1))
        meta['samp_size'] = sm.group(2)
        meta['samp_size_unit'] = 'L' if sm.group(3).lower() in ('l', 'litre', 'liter') else 'mL'
    sub = re.search(r'(\d+)\s*[x×]\s*(\d+)\s*m[lL]\s+subsamples', flat)
    if sub: meta['subsamples'] = f'{sub.group(1)} x {sub.group(2)} mL subsamples'
    where = find(r'collected from the (shoreline|shore|littoral zone|centre|center|deepest point|surface)', flat)
    parts = []
    if sm: parts.append(f'{sm.group(1).capitalize()} {sm.group(2)} {meta["samp_size_unit"]} water samples per site per sampling round')
    if sub: parts.append(f'each composed of {meta["subsamples"]}')
    if where: parts.append(f'collected from the {where}')
    meta['samp_collect_method'] = ', '.join(parts) or None

    # habitat → ENVO
    if re.search(r'\blakes?\b|\btarns?\b|still waters|lentic', flat, re.I):
        meta['env_broad_scale'] = 'freshwater lake biome [ENVO:01000252]'
        meta['env_local_scale'] = 'lake [ENVO:00000020]'
        meta['env_medium'] = 'fresh water [ENVO:00002011]'
    elif re.search(r'\bponds?\b', flat, re.I):
        meta['env_broad_scale'] = 'aquatic biome [ENVO:00002030]'
        meta['env_local_scale'] = 'pond water [ENVO:00002228]'
        meta['env_medium'] = 'liquid water [ENVO:00002006]'
    meta['habitat_natural'] = not re.search(r'mesocosm|aquari|tank experiment', flat, re.I)

    # country / region
    if re.search(r'\bUK\b|United Kingdom|England|Scotland|Wales|Defra|UKCEH', flat):
        meta['country'] = 'United Kingdom'
    for k, v in UK_REGIONS.items():
        if re.search(r'\b' + k + r'\b', flat, re.I) and k not in ('england',):
            meta['region'] = v; break

    # sites from a table of "<Name>\n<CODE>\t<type>\t..." rows
    for m in re.finditer(r'(?:^|\t)([A-Z][A-Za-z\' ]{2,40}?)\s*\n([A-Z]{3,5})\t([^\t\n]+)\t([^\t\n]+)\t([^\t\n]+)\t([^\t\n]+?)(?=\t|\n)', text, re.M):
        name, code = norm(m.group(1)), m.group(2)
        if name.lower() in ('lake', 'code', 'site', 'name') or any(s['code'] == code for s in meta['sites']): continue
        meta['sites'].append({'name': name, 'code': code,
                              'attributes': ' '.join(norm(m.group(g)) for g in (3, 4, 5, 6))})

    if not meta['sites']:   # tables without tab separators (e.g. from PDF): "<Name>\n<CODE> v1 .. vk <NextName>"
        lines = text.splitlines()
        rows = [(i, m.group(1), m.group(2).split()) for i, l in enumerate(lines)
                if (m := re.match(r'^([A-Z]{3,5})\s+(\S.*)$', l.strip()))]
        # keep the longest run of consecutive table rows
        best, cur = [], []
        for r in rows:
            if cur and r[0] != cur[-1][0] + 1: best, cur = (cur if len(cur) > len(best) else best), []
            cur.append(r)
        rows = cur if len(cur) > len(best) else best
        if len(rows) >= 3:
            k = min(len(t) for _, _, t in rows)          # value columns per row (last row has no trailing name)
            first = []
            for tok in reversed(lines[rows[0][0] - 1].split()):
                if not tok[:1].isupper(): break
                first.insert(0, tok)
            names = [' '.join(first)] + [' '.join(t[k:]) for _, _, t in rows[:-1]]
            if all(names) and len(set(names)) == len(names):
                meta['sites'] = [{'name': nm, 'code': code, 'attributes': ' '.join(t[:k])}
                                 for nm, (_, code, t) in zip(names, rows)]

    # sampling rounds: "weeks commencing 14th July and 13th October 2025"
    wc = re.search(r'weeks? commencing ((?:\d{1,2}(?:st|nd|rd|th)?\s+' + MONTH_RE + r'(?:\s+\d{4})?(?:,\s*|\s+and\s+)?)+)', flat, re.I)
    if wc:
        year = find(r'(\d{4})', wc.group(1)) or find(r'commencing.{0,60}?(\d{4})', flat)
        for i, d in enumerate(re.findall(r'\d{1,2}(?:st|nd|rd|th)?\s+' + MONTH_RE, wc.group(1), re.I)):
            iso = parse_date(d, int(year) if year else None)
            if iso:
                end = (date.fromisoformat(iso) + timedelta(days=6)).isoformat()
                meta['rounds'].append({'code': f'R{i+1}', 'iso_date': f'{iso}/{end}', 'verbatim': f'week commencing {d} {year}',
                                       'month': re.search(MONTH_RE, d, re.I).group(0).lower()})
    # per-round mean volumes: "... October ... compared to July (1207 mL and 720 mL respectively)"
    for s in sents:
        par = re.search(r'\(([^()]*\d\s*m[lL][^()]*)respectively\)', s)
        if par and meta['rounds']:
            vols = re.findall(r'(\d[\d,]*)\s*m[lL]', par.group(1))
            months = [mm.lower() for mm in re.findall(MONTH_RE, s[:par.start()], re.I)]
            for mo, v in zip(months, vols):
                for r in meta['rounds']:
                    if r['month'] == mo: r['avg_vol_ml'] = v.replace(',', '')
            break
    avg = re.search(r'average volume filtered was (\d[\d,]*)\s*m[lL](?:\s*\(range (\d+) to (\d+)\s*m[lL]\))?', flat, re.I)
    if avg: meta['notes'].append(f'Mean volume filtered {avg.group(1)} mL' + (f' (range {avg.group(2)}–{avg.group(3)} mL)' if avg.group(2) else ''))

    # ── filtration / storage ─────────────────────────────────────────────────
    fl = re.search(r'(\d+)\s*mm diameter\s*(\d+(?:\.\d+)?)\s*[µu]m?\s*([a-z ]+?)\s+filters', flat, re.I)
    if fl:
        meta['filter_diameter'], meta['size_frac'], meta['filter_material'] = fl.group(1), fl.group(2), norm(fl.group(3))
    meta['filter_active'] = bool(re.search(r'vacuum|peristaltic|pump', flat, re.I))
    proc = [s for s in sents if re.search(r'filter(ed|ing)? within|filters were used per sample|Filtering was stopped', s, re.I)]
    meta['samp_mat_process'] = ' '.join(proc[:3]) or None
    within = re.search(r'filtered within (\d+)\s*(hours?|h|days?)', flat, re.I)
    if within: meta['samp_store_dur'] = f'PT{within.group(1)}H' if within.group(2).lower().startswith('h') else f'P{within.group(1)}D'
    if re.search(r'transported on ice', flat, re.I): meta['samp_store_method_additional'] = 'Transported on ice; ' + (f'filtered within {within.group(1)} {within.group(2)}' if within else '')
    sl = re.search(r'transported on ice to ([A-Z][\w ]+?) and', flat)
    if sl: meta['samp_store_loc'] = sl.group(1)
    sol = re.search(r'containing\s+(?:\d+\s*m[lL]\s+)?(DNA/RNA Shield|ethanol|Longmire\'?s?(?: buffer)?|ATL buffer|RNAlater)\s*(\([^)]+\))?', flat, re.I)
    if sol: meta['prepped_samp_store_sol'] = norm(sol.group(1) + ' ' + (sol.group(2) or ''))
    tmp = re.search(r'(?:stored|frozen) at (-?\d+)\s*°?\s*C', flat, re.I)
    if tmp: meta['prepped_samp_store_temp'] = tmp.group(1)

    # ── extraction ───────────────────────────────────────────────────────────
    kit = re.search(r'performed using ([A-Z][^()]{3,80}?Kit)\s*\(([^)]+)\)', flat)
    if kit: meta['ext_kit'] = f'{kit.group(1)} ({kit.group(2)})'
    sop = re.search(r'follow those outlined in ([A-Z][\w\-]+ et al\.? \(\d{4}\))', flat)
    if sop:
        meta['nucl_acid_ext'] = ref_for(sop.group(1).replace('(', '').replace(')', ''), refs)
        if re.search(r'(?:metabarcoding|PCR|amplification) conditions follow', flat, re.I):
            meta['nucl_acid_amp'] = meta['nucl_acid_ext']
    lys = [s for s in sents if re.search(r'mechanical lysis|bead[- ]beat|homogeni[sz]er|Proteinase K', s, re.I)]
    meta['ext_lysis'] = ' '.join(lys[:2]) or None
    kitname = meta.get('ext_kit') or ''
    if re.search(r'DNeasy|Power(?:Water|Soil|Max)|Quick-DNA|NucleoSpin|QIAamp|E\.Z\.N\.A|Monarch', kitname, re.I):
        meta['ext_sep'] = 'column-based'
    elif re.search(r'MagMAX|KingFisher|MagAttract|magnetic', kitname + flat, re.I):
        meta['ext_sep'] = 'magnetic beads'
    mod = re.search(r'with the following amendments.*?(?=Thereafter|DNA extraction was verified)', flat, re.I)
    if mod: meta['ext_modifications'] = norm(re.sub(r'^with the following amendments[^.]*\.\s*', '', mod.group(0)))[:600]
    conc = [s for s in sents if re.search(r'Qubit|Nanodrop|NanoDrop|fluorometer', s) and 'librar' not in s.lower()]
    meta['conc_method'] = conc[0] if conc else None
    nm = re.search(r'normalised to (?:a standard template concentration of )?(\d+(?:\.\d+)?\s*ng/[µu]L)', flat, re.I)
    if nm: meta['notes'].append(f'Sample DNA normalised to {nm.group(1)} before PCR')

    # ── assays ───────────────────────────────────────────────────────────────
    REJECT = r'(?:instead of|rather than|in preference to|as opposed to|alternatives? such as|compared with)\s+'
    accepted = norm(re.sub(REJECT + r'(?:the\s+)?[\w/\-]+', ' ', flat, flags=re.I))
    counts = {}
    for mk in markers_in(accepted):
        counts[mk] = len(re.findall(r'(?<![A-Za-z0-9])' + mk + r'(?![A-Za-z0-9])', accepted, re.I))
    seen_markers = [m for m, n in counts.items() if n >= 3]
    cap = re.findall(r'\([a-h]\)\s*(' + MARKER_RE + r')\b', flat)
    order = [canon_marker(m) for m in cap] or MARKERS
    seen_markers.sort(key=lambda m: order.index(m) if m in order else 99)
    taxa_for = {}
    for mk in seen_markers:
        counts = {}
        for pat, taxon in TAXON_WORDS:
            n = len(re.findall(r'(?<![A-Za-z0-9])' + mk + r'\)?\s+(?:' + pat + r')\b', flat, re.I))
            if n: counts[taxon] = counts.get(taxon, 0) + n
        if counts: taxa_for[mk] = max(counts, key=counts.get)
    marker_for_taxon = {v: k for k, v in taxa_for.items()}

    scope_word = {}
    for mk in seen_markers:
        w = re.findall(r'(?<![A-Za-z0-9])' + mk + r'\)?\s+(prokaryotes|bacteria\w*|eukaryotes|diatoms|metazoans|fish|vertebrates|invertebrates|algae|fungi)\b', flat, re.I)
        if w: scope_word[mk] = max(set(x.lower() for x in w), key=[x.lower() for x in w].count)
    assays = {mk: {'marker': mk, 'target_gene': TARGET_GENE.get(mk), 'targetTaxonomicAssay': taxa_for.get(mk),
                   'targetTaxonomicScope': scope_word.get(mk)} for mk in seen_markers}

    # primer names: "<marker> ... primers A and B" / "A/B primers of Author YEAR"; skip rejected alternatives
    for s in sents:
        if 'primer' not in s.lower(): continue
        main = re.split(r'in preference to|as opposed to|instead of|alternatives such as|rather than', s)[0]
        pm = re.search(r'primers?\s+([A-Za-z][\w\-]*\d\w*)\s*(?:and|/)\s*([A-Za-z][\w\-]*\d\w*)', main) or \
             re.search(r'([A-Za-z][\w\-]*\d\w*)\s*/\s*([A-Za-z][\w\-]*\d\w*)\s+primers?', main)
        if not pm: continue
        mk = (markers_in(main) or [None])[0]
        if not mk:
            for pat, taxon in TAXON_WORDS:
                if re.search(r'For\s+(?:' + pat + r')', main, re.I) and taxon in marker_for_taxon:
                    mk = marker_for_taxon[taxon]; break
            if not mk and re.search(r'For invertebrates', main, re.I):
                mk = marker_for_taxon.get('Metazoa')
        if mk in assays:
            a = assays[mk]
            a['primer_name_fwd'], a['primer_name_rev'] = pm.group(1), pm.group(2)
            ref = re.search(r'primers? of ([A-Z][\w\-]+ et al\.? \(?\d{4}\)?)', main)
            if ref:
                full = ref_for(ref.group(1).replace('(', '').replace(')', ''), refs)
                a['primer_ref_fwd'] = a['primer_ref_rev'] = full
    for mk, a in assays.items():
        fwd, rev = a.get('primer_name_fwd'), a.get('primer_name_rev')
        a['assay_name'] = f'{mk}_{fwd}-{rev}' if fwd and rev else mk

    # PCR recipes scoped by "For <markers>" sentences
    scope, recipes = None, {}
    for s in sents:
        fm = re.match(r'For ((?:' + MARKER_RE + r'[\s,&]*(?:and\s+)?)+)(?:assays?)?,', s)
        if fm: scope = markers_in(fm.group(1))
        elif scope and not re.search(r'[µu]L', s): scope = scope if re.search(r'dispensed|template', s) else None
        if scope and re.search(r'[µu]L', s):
            for mk in scope: recipes.setdefault(mk, []).append(s)
    mm_re = r'(Q5[\w\- ]*?(?:polymerase|Master Mix)|Phusion[\w ]*?(?:master mix|polymerase)|Platinum[\w ]+|KAPA[\w ]+|AmpliTaq[\w ]+|GoTaq[\w ]+|MyTaq[\w ]+|Qiagen Multiplex[\w ]+|HotStar[\w ]+)'
    for mk, ss in recipes.items():
        if mk not in assays: continue
        a, joined = assays[mk], ' '.join(ss)
        m = re.search(mm_re, joined, re.I)
        if m:
            mf = re.search(re.escape(m.group(1)) + r'\s*\(([^)]+)\)', joined)
            a['commercial_mm'] = norm(m.group(1)) + (f' ({mf.group(1)})' if mf else '')
        a['custom_mm'] = ss[0][:500]
        dna = re.search(r'(\d+(?:\.\d+)?)\s*[µu]L\s*(?:of\s+)?template DNA', joined)
        if dna: a['pcr_dna_vol'] = dna.group(1)
        recipe = next((x for x in ss if re.search(r'contained|combined with|consisted of|comprised', x, re.I)), ss[0])
        vols = [float(v) for v in re.findall(r'(\d+(?:\.\d+)?)\s*[µu]L', recipe)]
        each = re.search(r'(\d+(?:\.\d+)?)\s*[µu]L\s*(?:of\s*)?(?:(?:\d+\s*[µu]?M|\d+\s*mM)\s*)?(?:of )?each', recipe)
        total = sum(vols) + (float(each.group(1)) if each else 0)   # "each" primer counted twice
        if dna and dna.group(0) not in recipe: total += float(dna.group(1))
        if vols: a['amplificationReactionVolume'] = f'{total:g}'
        pv = re.search(r'(\d+(?:\.\d+)?)\s*[µu]L\s*(?:of\s*)?(?:each\s*)?(\d+)\s*[µu]M\s*(?:of\s*)?(?:each\s*)?forward', recipe) or \
             re.search(r'(\d+(?:\.\d+)?)\s*[µu]L\s*of each\s*(\d+)\s*[µu]M', recipe)
        if pv: a['pcr_primer_vol'], a['pcr_primer_conc'] = pv.group(1), pv.group(2)
        elif re.search(r'\d+\s*mM\s*of each forward', recipe):
            meta['notes'].append(f'{mk}: primer stock concentration is reported in mM (likely µM) – please check')
    # PCR replicates
    if re.search(r'single reactions|single PCR|one PCR replicate|without (?:PCR )?replication', flat, re.I): rep = '1'
    else:
        rm = re.search(r'(\w+) PCR replicates (?:were|per)', flat, re.I) or re.search(r'in (duplicate|triplicate)', flat, re.I)
        rep = {'duplicate': '2', 'triplicate': '3'}.get(rm.group(1).lower(), str(num_word(rm.group(1)) or '')) if rm else None
    an = re.search(r'anneal\w*\s+(?:at\s+)?(\d{2}(?:\.\d)?)\s*°\s*C', flat, re.I)
    cy = re.search(r'(\d{2})\s+cycles\b(?! for)', flat, re.I)
    for a in assays.values():
        if rep: a['pcr_rep'] = rep
        if an: a['annealingTemp'] = an.group(1)
        if cy: a['pcr_cycles'] = cy.group(1)
    if re.search(r'thermocycling (?:parameters|conditions) are (?:described|given|shown) in Table', flat, re.I):
        meta['notes'].append('Thermocycling conditions and primer sequences are in a report table that is not part of the text – please add them manually')

    # ── library prep / sequencing ────────────────────────────────────────────
    if re.search(r'two[- ]step', flat, re.I): meta['barcoding_pcr_appr'] = 'two-step PCR'
    elif re.search(r'one[- ]step|single[- ]step', flat, re.I): meta['barcoding_pcr_appr'] = 'one-step PCR'
    elif re.search(r'ligation', flat, re.I): meta['barcoding_pcr_appr'] = 'ligation-based'
    ad = re.search(r'forward adaptor:\s*5\W{0,2}-\s?([ACGTN]+)\s?-3\W{0,2},\s*reverse adaptor:\s*5\W{0,2}-\s?([ACGTN]+)\s?-3', flat)
    if ad: meta['adapter_forward'], meta['adapter_reverse'] = ad.group(1), ad.group(2)
    ad2 = list(re.finditer(r'forward adaptor:\s*5\W{0,2}-\s?([ACGTN]+)\s?-3\W{0,2},\s*reverse adaptor:\s*5\W{0,2}-\s?([ACGTN]+)\s?-3', flat))
    if len(ad2) > 1: meta['notes'].append(f'Indexing (step II) adaptors: forward {ad2[1].group(1)}, reverse {ad2[1].group(2)}')
    p2 = re.search(r'Per sample (\d+(?:\.\d+)?)\s*[µu]L of step 1 product.*?to a total of (\d+(?:\.\d+)?)\s*[µu]L', flat, re.I)
    if p2:
        meta['pcr2_dna_vol'], meta['pcr2_amplificationReactionVolume'] = p2.group(1), p2.group(2)
        m = re.search(mm_re, p2.group(0), re.I)
        if m: meta['pcr2_commercial_mm'] = norm(m.group(1))
        meta['pcr2_custom_mm'] = norm(p2.group(0))
    lq = re.search(r'libraries were quantified using (?:a |an )?([^.,;]+?)(?: and | to |,|\.)', flat, re.I)
    if lq: meta['lib_conc_meth'] = norm(lq.group(1))
    lib = [s for s in sents if re.search(r'librar(y|ies)', s, re.I) and re.search(r'Qubit|quantif|normali[sz]|gel purif|Bioanalyzer|TapeStation', s)]
    meta['lib_screen'] = ' '.join(lib[:3]) or None
    ph = re.search(r'(\d+(?:\.\d+)?)\s*%\s*(?:addition of\s*)?(?:[\w ]{0,30})PhiX', flat, re.I)
    if ph: meta['phix_perc'] = ph.group(1)
    if re.search(r'\bIllumina\b', flat): meta['platform'] = 'ILLUMINA'
    elif re.search(r'Nanopore|MinION|GridION|PromethION', flat): meta['platform'] = 'OXFORD_NANOPORE'
    elif re.search(r'PacBio', flat): meta['platform'] = 'PACBIO_SMRT'
    ins = re.search(r'Illumina (NextSeq|MiSeq|NovaSeq|HiSeq|iSeq)\b', flat) or re.search(r'\b(NextSeq|MiSeq|NovaSeq|HiSeq|iSeq|MinION|GridION|PromethION)(?:™)?\s*(\d+(?:/\d+)?)?', flat)
    if ins:
        model = re.search(r'(NextSeq|MiSeq|NovaSeq|HiSeq|iSeq)(?:™)?\s*(\d+(?:/\d+)?)', flat)
        meta['instrument'] = ('Illumina ' if meta.get('platform') == 'ILLUMINA' else '') + (f'{model.group(1)} {model.group(2)}' if model else ins.group(1))
    if re.search(r'paired[- ]end|mergePairs|forward and reverse reads', flat, re.I): meta['lib_layout'] = 'paired end'
    loc = re.search(r'to ([A-Z][\w ]+?) \([^)]*\) for DNA extraction and sequencing|to ([A-Z][\w ]+?) for (?:DNA extraction and )?sequencing', flat)
    if loc: meta['sequencing_location'] = norm(loc.group(1) or loc.group(2))
    # kit per assay: "<kit> 600 Cycles for pooled 16S,18S, rbcL and COI, and 200 Cycles for 12S"
    kb = re.search(r'((?:NextSeq|MiSeq|NovaSeq)(?:(?!NextSeq|MiSeq|NovaSeq)[^.])*?)\s*(\d+)\s*Cycles for', flat)
    if kb:
        base = norm(re.sub(r'Chemistry', '', kb.group(1))).replace('™', '')
        for cyc, mlist in re.findall(r'(\d+)\s*Cycles for (?:pooled\s*)?((?:' + MARKER_RE + r'[\s,]*(?:and\s+)?)+)', flat):
            for mk in markers_in(mlist):
                if mk in assays: assays[mk]['seq_kit'] = f'{base} {cyc} cycles'
    # runs: dates in order, run descriptions in order
    rd = re.search(r'(?:Sequenc\w* runs? (?:were|was) completed[^.]*? on|runs? on)\s*([\d/]+(?:\s*(?:and|,)\s*[\d/]+)*)', flat, re.I)
    rdesc = re.findall(r'(\d+)\s*cycle run \(((?:' + MARKER_RE + r'[\s,]*(?:and\s+)?)+)\)', flat, re.I)
    if rd:
        dates = [parse_date(d) for d in re.findall(r'\d{1,2}/\d{1,2}/\d{2,4}', rd.group(1))]
        for i, d in enumerate(dates):
            cyc, mlist = rdesc[i] if i < len(rdesc) else ('', '')
            mks = markers_in(mlist) if mlist else list(assays)
            meta['runs'].append({'seq_run_id': f'run_{d.replace("-", "")}' + (f'_{cyc}cycle' if cyc else ''), 'date': d, 'assays': mks})

    # ── bioinformatics ───────────────────────────────────────────────────────
    cut = re.search(r'cutadapt\s*\((v?[\d.]+)', flat, re.I) or re.search(r'cutadapt\s*(v?[\d.]+)?', flat, re.I)
    if cut: meta['trim_method'] = 'cutadapt' + (f' {cut.group(1)}' if cut.group(1) else '')
    if re.search(r'DADA2', flat):
        rv = find(r'\bin R\s*v?\.?\s*([\d.]+\d)', flat)
        meta['sop_bioinformatics'] = 'DADA2 pipeline (Callahan et al. 2016)' + (f' in R v{rv}' if rv else '')
        meta['error_rate_tool'] = 'DADA2 filterAndTrim' if re.search(r'maxEE', flat) else None
        if re.search(r'mergePairs', flat): meta['merge_tool'] = 'DADA2 mergePairs'
        meta['otu_clust_tool'] = 'DADA2 (amplicon sequence variants, no clustering)'
        meta['otu_clust_cutoff'] = '100'
    ee = re.search(r'maxEE\)?\s*=\s*([\d.]+)(?:\s*,\s*([\d.]+))?', flat)
    if ee:
        meta['error_rate_type'] = 'expected error rate'
        meta['error_rate_cutoff'] = ee.group(1)
    trims = {}
    for m in re.finditer(r'((?:' + MARKER_RE + r'[\s,&]*(?:and\s+)?)+)amplicon reads (?:were|both)[^.]*?trimmed to (\d+)(?: and (\d+))? bases', flat):
        for mk in markers_in(m.group(1)): trims[mk] = (m.group(2), m.group(3) or m.group(2))
    for m in re.finditer(r'reads of ((?:' + MARKER_RE + r'[\s,&]*(?:and\s+)?)+)amplicon reads were both trimmed to (\d+) bases', flat):
        for mk in markers_in(m.group(1)): trims[mk] = (m.group(2), m.group(2))
    maxn = find(r'maxN\)?\s*=\s*(\d+)', flat)
    for mk, a in assays.items():
        bits = []
        if mk in trims: bits.append(f'truncLen forward {trims[mk][0]} bp, reverse {trims[mk][1]} bp')
        if maxn is not None: bits.append(f'maxN = {maxn}')
        if ee: bits.append(f'maxEE = {ee.group(1)} (forward)' + (f', {ee.group(2)} (reverse)' if ee.group(2) else ''))
        if bits: a['trim_param'] = '; '.join(bits)
    if re.search(r'removeBimeraDenovo', flat):
        meta['chimera_check_method'] = 'DADA2 removeBimeraDenovo'
        if re.search(r'removeBimeraDenovo default', flat, re.I): meta['chimera_check_param'] = 'default settings'
    # taxonomy: global method + per-marker overrides and databases
    bio = re.search(r'Bioinformatic analysis(.*?)(?:Statistical analysis|Results)', flat, re.S | re.I)
    btxt = bio.group(1) if bio else flat
    global_method = 'DADA2 assignTaxonomy (naive Bayesian classifier)' if 'assignTaxonomy' in btxt else None
    btxt = re.sub(r'mined from GenBank', 'mined from G-e-n-B-a-n-k', btxt)
    db_re = (r'(Silva\s*v?[\d.]*\d|SILVA\s*v?[\d.]*\d|PR2(?:\s*V?\s*\d[\d.]*\d)?|Midori2?|MIDORI2?|BOLD(?: database)?|GenBank|NCBI(?: database)?|diat\.barcode'
             r'|UNITE[\w .]*?|Greengenes2?|COI classifier|curated list for [^;,(]+?first compiled in [^;]+?\(\d{4}\)(?: and recently updated by [^;]+?\(\d{4}\))?'
             r'|curated list for [\w ]+|reference database from the [\w ]+\([A-Z]+,?\s*https?://[^)\s]+\))')
    bsents = sentences(btxt)
    scope = []
    for s in bsents:
        fm = re.match(r'For ((?:' + MARKER_RE + r'[\s,&]*(?:and\s+)?)+)', s)
        if fm: scope = markers_in(fm.group(1))
        # "Prokaryotes (16S), Silva v138.2 ... Eukaryotes (18S) PR2 V 5.0.1"
        for m in re.finditer(r'\((' + MARKER_RE + r')\),?\s*' + db_re, s):
            mk = canon_marker(m.group(1))
            if mk in assays: assays[mk].setdefault('dbs', []).append(norm(m.group(2)))
        if fm:
            for d in re.findall(db_re, s):
                for mk in scope:
                    if mk in assays: assays[mk].setdefault('dbs', []).append(norm(d))
            if re.search(r'SINTAX', s):
                for mk in scope:
                    if mk in assays: assays[mk]['otu_seq_comp_appr'] = 'SINTAX (Edgar 2016)'
        elif scope and not re.match(r'(?:ASVs|Primer|Sequences|Reference)', s):
            for d in re.findall(db_re, s):
                for mk in scope:
                    if mk in assays: assays[mk].setdefault('dbs', []).append(norm(d))
            if re.search(r'BLAST', s):
                for mk in scope:
                    if mk in assays: assays[mk]['blast_note'] = norm(s)
        else:
            scope = []
    for mk, a in assays.items():
        dbs = []
        for d in a.pop('dbs', []):
            cm = re.match(r'(?i)curated list for (.+?) first compiled in (.+?),?\s*\((\d{4})\)(?: and recently updated by (.+?),?\s*\((\d{4})\))?', d)
            if cm:
                d = f'Curated {cm.group(1)} reference database ({cm.group(2)} {cm.group(3)}' + (f'; {cm.group(4)} {cm.group(5)})' if cm.group(4) else ')')
            im = re.match(r'reference database from the (.+?)\(([A-Z]+),?\s*(https?://[^)\s]+)\)', d)
            if im: d = f'{im.group(2)} reference database ({norm(im.group(1))}; {im.group(3)})'
            d = re.sub(r'(?i)^curated list for (.*)', r'Curated \1 reference database', d)
            d = re.sub(r'^Midori2?$|^MIDORI2?$', 'MIDORI2', d)
            d = re.sub(r'\s+', ' ', d).replace('V 5', 'v5')
            vm = None
            if d.lower() == 'coi classifier':
                vm = re.search(r'COI classifier.*?(v\d[\d.]+)', btxt, re.I)
            d = d + (f' {vm.group(1)}' if vm else '')
            if d not in dbs: dbs.append(d)
        if dbs: a['otu_db'] = ' | '.join(dbs)
        a.setdefault('otu_seq_comp_appr', global_method)
        if a.get('otu_seq_comp_appr'):
            a['tax_assign_cat'] = 'sequence composition'
        if a.pop('blast_note', None):
            a['otu_seq_comp_appr'] += '; BLAST against NCBI for confirmation of selected taxa'
    # rarefaction depths
    rare = re.search(r'rarefaction to (.*?) using', flat, re.I)
    if rare:
        for n, mk in re.findall(r'([\d,]+)\s*(?:sequences\s*)?for\s+(' + MARKER_RE + r')', rare.group(1)):
            mk = canon_marker(mk)
            if mk in assays:
                assays[mk]['otu_final_description'] = f'ASV table rarefied to {n} reads per sample' + (' (rarefy_even_depth, phyloseq)' if 'phyloseq' in flat else '')
    raw_bits = []
    if meta.get('trim_method'): raw_bits.append(f'primers removed with {meta["trim_method"]}')
    if re.search(r'quality filter', flat, re.I): raw_bits.append('quality filtering')
    if re.search(r'denois', flat, re.I): raw_bits.append('denoising')
    if meta.get('merge_tool'): raw_bits.append('merging of paired reads')
    if meta.get('chimera_check_method'): raw_bits.append('chimera removal')
    if raw_bits: meta['otu_raw_description'] = 'ASV table after ' + ', '.join(raw_bits) + (' (DADA2)' if 'DADA2' in flat else '')
    cont = [s for s in sents if re.search(r'blanks?', s, re.I) and re.search(r'contaminat|not reliable|mis-?binning', s, re.I)]
    if cont: meta['screen_contam_method'] = ' '.join(cont[:2])[:700]
    stats = [p for p in ('phyloseq', 'vegan', 'microeco', 'lme4', 'emmeans', 'ggplot2') if re.search(r'\b' + p, flat)]
    if stats: meta['bioinfo_method_additional'] = 'Downstream analysis in R using ' + ', '.join(stats)

    meta['assays'] = list(assays.values())
    # biological replicates
    if meta.get('samples_per_site_round'): meta['biological_rep'] = str(meta['samples_per_site_round'])
    # bibliographic citation for an unpublished report
    if re.search(r'Final Report', flat, re.I) and meta.get('project_title'):
        yr = (meta.get('end_date') or '')[:4]
        who = meta.get('institution') or ''
        meta['bibliographicCitation'] = f'{who} ({yr}). {meta["project_title"]}. Evidence project final report' + (f' to {meta["funder"]}' if meta.get('funder') else '') + '.'
    return meta


# ════════════════════════════════════════════════════════════════════════════
#  3. Writing the FAIRe workbook
# ════════════════════════════════════════════════════════════════════════════

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
MISS = 'missing: not provided'
NA = 'not applicable'
NA_S = 'not applicable: sample group'
NA_C = 'not applicable: control sample'

# projectMetadata terms that may differ between assays (written to assay1..N columns)
ASSAY_TERMS = {
    'assay_name', 'assay_validation', 'targetTaxonomicAssay', 'targetTaxonomicScope', 'target_gene',
    'target_subfragment', 'ampliconSize', 'pcr_primer_forward', 'pcr_primer_reverse',
    'pcr_primer_name_forward', 'pcr_primer_name_reverse', 'pcr_primer_reference_forward',
    'pcr_primer_reference_reverse', 'pcr_primer_vol_forward', 'pcr_primer_vol_reverse',
    'pcr_primer_conc_forward', 'pcr_primer_conc_reverse', 'commercial_mm', 'custom_mm',
    'amplificationReactionVolume', 'pcr_dna_vol', 'pcr_rep', 'nucl_acid_amp', 'pcr_cond',
    'annealingTemp', 'pcr_cycles', 'seq_kit', 'trim_param', 'min_len_cutoff', 'otu_db',
    'tax_assign_cat', 'otu_seq_comp_appr', 'otu_final_description',
}
ASSAY_KEY = {  # FAIRe term -> key in an assay dict
    'pcr_primer_name_forward': 'primer_name_fwd', 'pcr_primer_name_reverse': 'primer_name_rev',
    'pcr_primer_reference_forward': 'primer_ref_fwd', 'pcr_primer_reference_reverse': 'primer_ref_rev',
    'pcr_primer_forward': 'primer_fwd', 'pcr_primer_reverse': 'primer_rev',
    'pcr_primer_vol_forward': 'pcr_primer_vol', 'pcr_primer_vol_reverse': 'pcr_primer_vol',
    'pcr_primer_conc_forward': 'pcr_primer_conc', 'pcr_primer_conc_reverse': 'pcr_primer_conc',
}


# ── low-level xlsx helpers ───────────────────────────────────────────────────

def col_num(col: str) -> int:
    n = 0
    for ch in col: n = n * 26 + ord(ch) - 64
    return n


def col_letter(n: int) -> str:
    s = ''
    while n:
        n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s


_ILLEGAL = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]')


def esc(v) -> str:
    v = _ILLEGAL.sub('', str(v))
    return v.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def cell_xml(ref: str, value, style: str | None = None) -> str:
    s = f' s="{style}"' if style else ''
    v = str(value)
    if re.fullmatch(r'-?\d+(?:\.\d+)?', v) and len(v) < 15:
        return f'<c r="{ref}"{s}><v>{v}</v></c>'
    sp = ' xml:space="preserve"' if v != v.strip() or '\n' in v else ''
    return f'<c r="{ref}"{s} t="inlineStr"><is><t{sp}>{esc(v)}</t></is></c>'


ROW_RE = re.compile(r'<row\b([^>]*?)(?:/>|>(.*?)</row>)', re.S)
CELL_RE = re.compile(r'<c\b[^>]*?\br="([A-Z]+)(\d+)"[^>]*?(?:/>|>.*?</c>)', re.S)


def parse_rows(sheet_xml: str):
    sd = re.search(r'<sheetData\s*/>|<sheetData>(.*?)</sheetData>', sheet_xml, re.S)
    rows = {}
    for m in ROW_RE.finditer(sd.group(1) or '' if sd else ''):
        rn = int(re.search(r'\br="(\d+)"', m.group(1)).group(1))
        cells = {c.group(1): c.group(0) for c in CELL_RE.finditer(m.group(2) or '')}
        rows[rn] = {'attrs': re.sub(r'\s*spans="[^"]*"', '', m.group(1)), 'cells': cells}
    return rows


def build_rows(rows: dict) -> str:
    out = []
    for rn in sorted(rows):
        r = rows[rn]
        cells = ''.join(r['cells'][c] for c in sorted(r['cells'], key=col_num))
        attrs = r.get('attrs') or f' r="{rn}"'
        out.append(f'<row{attrs}>{cells}</row>' if cells else f'<row{attrs}/>')
    return ''.join(out)


def replace_sheet_data(sheet_xml: str, rows: dict) -> str:
    body = build_rows(rows)
    sheet_xml = re.sub(r'<sheetData\s*/>|<sheetData>.*?</sheetData>', lambda _: f'<sheetData>{body}</sheetData>', sheet_xml, count=1, flags=re.S)
    max_r = max(rows) if rows else 1
    max_c = max((col_num(c) for r in rows.values() for c in r['cells']), default=1)
    return re.sub(r'<dimension ref="[^"]*"', f'<dimension ref="A1:{col_letter(max_c)}{max_r}"', sheet_xml, count=1)


VOCAB: dict = {}


def set_cell(rows, rn, col, value, keep_style=True, term=None):
    if term: value = normalise(term, value, VOCAB)
    if value in (None, ''): return
    row = rows.setdefault(rn, {'attrs': f' r="{rn}"', 'cells': {}})
    old = row['cells'].get(col, '')
    style = re.search(r'\bs="(\d+)"', old).group(1) if keep_style and re.search(r'\bs="(\d+)"', old) else None
    row['cells'][col] = cell_xml(f'{col}{rn}', value, style)


def shared_strings(pkg) -> list[str]:
    if 'xl/sharedStrings.xml' not in pkg: return []
    root = ET.fromstring(pkg['xl/sharedStrings.xml'])
    return [''.join(t.text or '' for t in si.iter(f'{{{MAIN}}}t')) for si in root.findall(f'{{{MAIN}}}si')]


def cell_text(cell_xml_str: str, ss: list[str]) -> str:
    if 't="s"' in cell_xml_str:
        v = re.search(r'<v>(\d+)</v>', cell_xml_str)
        return ss[int(v.group(1))] if v else ''
    if 't="inlineStr"' in cell_xml_str:
        return ''.join(re.findall(r'<t[^>]*>([^<]*)</t>', cell_xml_str))
    v = re.search(r'<v>([^<]*)</v>', cell_xml_str)
    return v.group(1) if v else ''


def sheet_paths(pkg) -> dict:
    wb = ET.fromstring(pkg['xl/workbook.xml'])
    rels = ET.fromstring(pkg['xl/_rels/workbook.xml.rels'])
    target = {r.get('Id'): r.get('Target') for r in rels.findall(f'{{{PKG_REL}}}Relationship')}
    out = {}
    for sh in wb.find(f'{{{MAIN}}}sheets'):
        t = target[sh.get(f'{{{REL}}}id')].lstrip('/')
        out[sh.get('name')] = t if t.startswith('xl/') else 'xl/' + t
    return out


# ── controlled vocabularies (read from the template itself) ──────────────────

VOCAB_RULES = {
    'nucl_acid_ext_lysis': [('physical', r'mechanical|bead|homogeni|vortex|sonicat|grind|FastPrep|TissueLyser|physical'),
                            ('enzymatic', r'proteinase|lysozyme|enzym'), ('chemical', r'\bSDS\b|guanidin|CTAB|chemical'),
                            ('thermal', r'\bheat|boil|thermal'), ('osmotic', r'osmotic')],
    'nucl_acid_ext_sep': [('column-based', r'column|spin|silica|column-based'), ('magnetic beads', r'magnetic|bead-based|SPRI'),
                          ('phenol chloroform', r'phenol'), ('precipitation', r'precipitat'), ('centrifugation', r'centrifug')],
    'filter_material': [('cellulose ester', r'mixed cellulose ester|\bMCE\b'), ('glass fiber', r'glass fib|GF/F|GF/C'),
                        ('track etched polycarbonate', r'polycarbonate'), ('polyethersulfone', r'polyethersulfone|\bPES\b'),
                        ('nylon', r'nylon')],
    'prepped_samp_store_sol': [('longmire', r'longmire'), ('ethanol', r'ethanol'), ('sodium acetate', r'sodium acetate'),
                               ('none', r'^none$|^dry$')],
}
VOCAB_RULES['samp_store_sol'] = VOCAB_RULES['prepped_samp_store_sol']


def load_vocab(pkg, paths, ss) -> dict:
    if 'Drop-down values' not in paths: return {}
    rows = parse_rows(pkg[paths['Drop-down values']].decode('utf-8'))
    out = {}
    for rn, r in rows.items():
        if rn == 1: continue
        cells = {c: cell_text(x, ss) for c, x in r['cells'].items()}
        term, ftype = cells.get('A'), cells.get('B')
        opts = [v.strip() for c, v in sorted(cells.items(), key=lambda kv: col_num(kv[0])) if col_num(c) >= 5 and v.strip()]
        if term and opts: out[term] = (ftype, opts)
    return out


def normalise(term, value, vocab):
    """Coerce a value into the template's controlled vocabulary for this term."""
    if value is None or term not in vocab: return value
    v = str(value).strip()
    ftype, opts = vocab[term]
    if v in MISSING_TERMS or v.startswith('other:'): return v
    if ftype == 'Boolean':
        return v if v in ('0', '1') else None
    low = {o.lower(): o for o in opts}
    if v.lower() in low: return low[v.lower()]
    hits = [o for o, pat in VOCAB_RULES.get(term, []) if re.search(pat, v, re.I)]
    if hits:
        return hits[0]
    return f'other: {v}'


MISSING_TERMS = {MISS, NA, NA_S, NA_C, 'missing: not collected', 'missing: restricted access'}


# ── value helpers ────────────────────────────────────────────────────────────

def present(v) -> bool:
    return v not in (None, '', [], {})


def missing_for(req: str):
    return MISS if req in ('M', 'HR', 'R') else None


# ── main ─────────────────────────────────────────────────────────────────────

def project_values(meta: dict, version: str) -> dict:
    """Project-level FAIRe term -> value (None = not found)."""
    title = meta.get('project_title') or ''
    pid = meta.get('project_id') or re.sub(r'_+', '_', re.sub(r'[^a-z0-9]+', '_', ' '.join(
        w for w in title.lower().split() if w not in ('of', 'the', 'and', 'to', 'a', 'an', 'for', 'in')))).strip('_')
    while len(pid) > 40 and '_' in pid: pid = pid.rsplit('_', 1)[0]
    notes = meta.get('notes') or []
    return {
        'recordedBy': meta.get('recorded_by'), 'recordedByID': meta.get('recorded_by_id'),
        'project_contact': meta.get('project_contact'), 'institution': meta.get('institution'),
        'institutionID': meta.get('institution_ror'), 'project_name': title or None, 'project_id': pid or None,
        'study_factor': meta.get('study_factor'), 'assay_type': 'metabarcoding',
        'sterilise_method': meta.get('sterilise_method'),
        'neg_cont_0_1': '1' if meta.get('neg_controls') else '0', 'pos_cont_0_1': '1' if meta.get('pos_controls') else '0',
        'checkls_ver': version, 'mod_date': time.strftime('%Y-%m-%d'),
        'license': meta.get('license'), 'rightsHolder': meta.get('institution'), 'accessRights': meta.get('access_rights'),
        'bibliographicCitation': meta.get('bibliographicCitation'), 'code_repo': meta.get('code_repo'),
        'biological_rep': meta.get('biological_rep'), 'pcr_0_1': '1',
        'pcr_method_additional': '; '.join(n for n in notes if 'normalised' in n) or None,
        'barcoding_pcr_appr': meta.get('barcoding_pcr_appr'),
        'pcr2_amplificationReactionVolume': meta.get('pcr2_amplificationReactionVolume'),
        'pcr2_commercial_mm': meta.get('pcr2_commercial_mm'), 'pcr2_custom_mm': meta.get('pcr2_custom_mm'),
        'pcr2_dna_vol': meta.get('pcr2_dna_vol'), 'pcr2_cond': meta.get('pcr2_cond'),
        'sequencing_location': meta.get('sequencing_location'), 'platform': meta.get('platform'),
        'instrument': meta.get('instrument'), 'lib_layout': meta.get('lib_layout'),
        'adapter_forward': meta.get('adapter_forward'), 'adapter_reverse': meta.get('adapter_reverse'),
        'lib_screen': meta.get('lib_screen'),
        'seq_method_additional': '; '.join([n for n in notes if 'adaptor' in n] +
                                           ([f'{meta["phix_perc"]}% PhiX spike-in'] if meta.get('phix_perc') else [])) or None,
        'sop_bioinformatics': meta.get('sop_bioinformatics'), 'trim_method': meta.get('trim_method'),
        'demux_tool': meta.get('demux_tool'), 'merge_tool': meta.get('merge_tool'),
        'error_rate_tool': meta.get('error_rate_tool'), 'error_rate_type': meta.get('error_rate_type'),
        'error_rate_cutoff': meta.get('error_rate_cutoff'), 'chimera_check_method': meta.get('chimera_check_method'),
        'chimera_check_param': meta.get('chimera_check_param'), 'otu_clust_tool': meta.get('otu_clust_tool'),
        'otu_clust_cutoff': meta.get('otu_clust_cutoff'), 'screen_contam_method': meta.get('screen_contam_method'),
        'otu_raw_description': meta.get('otu_raw_description'),
        'bioinfo_method_additional': meta.get('bioinfo_method_additional'),
    }


def assay_value(a: dict, term: str, meta: dict):
    if term == 'nucl_acid_amp': return a.get('nucl_acid_amp') or meta.get('nucl_acid_amp')
    if term in ('probeReporter', 'probeQuencher', 'probe_seq', 'probe_ref', 'probe_conc'): return NA
    return a.get(ASSAY_KEY.get(term, term))


def fill_project(rows, ss, meta, version):
    assays = meta.get('assays') or []
    pv = project_values(meta, version)
    d1 = rows.get(1, {}).get('cells', {}).get('D', '')
    hdr_style = re.search(r'\bs="(\d+)"', d1).group(1) if re.search(r'\bs="(\d+)"', d1) else None
    if len(assays) > 1:
        for i in range(len(assays)):
            rows[1]['cells'][col_letter(5 + i)] = cell_xml(f'{col_letter(5 + i)}1', f'assay{i + 1}', hdr_style)
    filled = 0
    for rn, row in rows.items():
        if rn < 2: continue
        term = cell_text(row['cells'].get('C', ''), ss)
        req = cell_text(row['cells'].get('A', ''), ss)
        section = cell_text(row['cells'].get('B', ''), ss)
        if not term: continue
        if section == 'Targeted assay detection':
            set_cell(rows, rn, 'D', NA); continue
        if term in ASSAY_TERMS or term.startswith('probe'):
            vals = [assay_value(a, term, meta) for a in assays]
            if term == 'assay_name':
                set_cell(rows, rn, 'D', ' | '.join(str(v) for v in vals) or MISS)
                if len(assays) > 1:
                    for i, v in enumerate(vals): set_cell(rows, rn, col_letter(5 + i), v)
                filled += 1; continue
            if not vals or not any(present(v) for v in vals):
                m = missing_for(req)
                if m: set_cell(rows, rn, 'D', m)
                continue
            vals = [v if present(v) else (missing_for(req) or '') for v in vals]
            if len(set(map(str, vals))) == 1 or len(assays) == 1:
                set_cell(rows, rn, 'D', vals[0], term=term)
            else:
                for i, v in enumerate(vals):
                    if v != '': set_cell(rows, rn, col_letter(5 + i), v, term=term)
            filled += 1
            continue
        v = pv.get(term)
        if present(v):
            set_cell(rows, rn, 'D', v, term=term); filled += 1
        else:
            m = missing_for(req)
            if m: set_cell(rows, rn, 'D', m)
    return filled


def header_map(rows, ss, header_row=3):
    hdr = {cell_text(x, ss): c for c, x in rows.get(header_row, {}).get('cells', {}).items()}
    req = {c: cell_text(x, ss) for c, x in rows.get(1, {}).get('cells', {}).items()}
    return {k: v for k, v in hdr.items() if k}, req


def write_record(rows, rn, rec, hdr, req, control=False):
    """control: False/'field' -> unreported values are 'missing: not provided'; 'lab' -> 'not applicable: control sample'."""
    rows[rn] = {'attrs': f' r="{rn}"', 'cells': {}}
    for term, col in hdr.items():
        if term == 'samp_name' or col == 'A' and term.startswith('#'): pass
        v = rec.get(term)
        if not present(v):
            r = req.get(col, 'O')
            v = (NA_C if control == 'lab' else MISS) if r in ('M', 'HR', 'R') else None
        if present(v): v = normalise(term, v, VOCAB)
        if present(v): rows[rn]['cells'][col] = cell_xml(f'{col}{rn}', v)


def build_samples(meta):
    """Return list of (sample_record, assays_for_sample)."""
    assays = [a['assay_name'] for a in meta.get('assays') or []] or ['assay1']
    sites = meta.get('sites') or [{'name': None, 'code': 'SITE1'}]
    rounds = meta.get('rounds') or [{'code': 'R1'}]
    n = int(meta.get('samples_per_site_round') or 1)
    fb = int(meta.get('field_blanks_per_site_round') or 0)
    country, region = meta.get('country'), meta.get('region')
    out = []
    common = {
        'env_broad_scale': meta.get('env_broad_scale'), 'env_local_scale': meta.get('env_local_scale'),
        'samp_store_dur': meta.get('samp_store_dur'), 'samp_store_loc': meta.get('samp_store_loc'),
        'samp_store_method_additional': meta.get('samp_store_method_additional'),
        'samp_mat_process': meta.get('samp_mat_process'),
        'filter_passive_active_0_1': '1' if meta.get('filter_active') else None,
        'size_frac': meta.get('size_frac'), 'filter_diameter': meta.get('filter_diameter'),
        'filter_material': meta.get('filter_material'),
        'prepped_samp_store_temp': meta.get('prepped_samp_store_temp'),
        'prepped_samp_store_sol': meta.get('prepped_samp_store_sol'),
        'samp_vol_we_dna_ext_unit': 'mL',
        'nucl_acid_ext': meta.get('nucl_acid_ext'), 'nucl_acid_ext_kit': meta.get('ext_kit'),
        'nucl_acid_ext_lysis': meta.get('ext_lysis'), 'nucl_acid_ext_sep': meta.get('ext_sep'),
        'nucl_acid_ext_modify': meta.get('ext_modifications'),
        'concentration_unit': 'ng/µl', 'concentration_method': meta.get('conc_method'),
        'detected_notDetected': NA, 'assay_name': ' | '.join(assays),
    }
    for site in sites:
        loc = site.get('name')
        geo = (f'{country}: {region}, {loc}' if region else f'{country}: {loc}') if country and loc else None
        for rd in rounds:
            group = [f'{site["code"]}_{rd["code"]}_{i:02d}' for i in range(1, n + 1)]
            fb_ids = [f'{site["code"]}_{rd["code"]}_FB' + (f'{j}' if fb > 1 else '') for j in range(1, fb + 1)]
            where = {'decimalLatitude': site.get('latitude'), 'decimalLongitude': site.get('longitude'),
                     'geo_loc_name': geo, 'eventDate': rd.get('iso_date'), 'verbatimEventDate': rd.get('verbatim')}
            prep = f'Mean volume filtered in this sampling round: {rd["avg_vol_ml"]} mL (per-sample volumes not reported)' if rd.get('avg_vol_ml') else None
            for sid in group:
                rec = dict(common, **where, samp_name=sid, samp_category='sample', neg_cont_type=NA_S, pos_cont_type=NA_S,
                           rel_cont_id=' | '.join(fb_ids) or None,
                           biological_rep_relation=' | '.join(group) if n > 1 else None,
                           env_medium=meta.get('env_medium'),
                           habitat_natural_artificial_0_1='0' if meta.get('habitat_natural', True) else '1',
                           samp_collect_method=meta.get('samp_collect_method'),
                           samp_size=meta.get('samp_size'), samp_size_unit=meta.get('samp_size_unit'),
                           prep_method_additional=prep, org_note=site.get('attributes'))
                out.append((rec, assays, False))
            for bid in fb_ids:
                rec = dict(common, **where, samp_name=bid, samp_category='negative control', neg_cont_type='field negative',
                           pos_cont_type=NA_C, rel_cont_id=NA_C, env_broad_scale=NA_C, env_local_scale=NA_C, env_medium=NA_C,
                           habitat_natural_artificial_0_1=NA_C, materialSampleID=NA_C,
                           samp_collect_method='Field blank of purified water collected alongside the samples')
                out.append((rec, assays, 'field'))
    _unused = {k: NA_C for k in ('decimalLatitude', 'decimalLongitude', 'geo_loc_name', 'eventDate', 'env_broad_scale',
                             'env_local_scale', 'env_medium', 'habitat_natural_artificial_0_1', 'samp_collect_method',
                             'samp_size', 'samp_size_unit', 'samp_store_temp', 'samp_store_sol', 'samp_store_dur',
                             'samp_mat_process', 'filter_passive_active_0_1', 'size_frac', 'filter_diameter',
                             'filter_material', 'prepped_samp_store_temp', 'prepped_samp_store_sol', 'samp_vol_we_dna_ext',
                             'samp_vol_we_dna_ext_unit')}
    ext_fields = {k: common[k] for k in ('nucl_acid_ext', 'nucl_acid_ext_kit', 'nucl_acid_ext_lysis', 'nucl_acid_ext_sep',
                                         'nucl_acid_ext_modify', 'concentration_unit', 'concentration_method', 'detected_notDetected', 'assay_name')}
    for i in range(1, int(meta.get('extraction_blanks') or 0) + 1):
        rec = dict(ext_fields, samp_name=f'EXTBLANK_{i:02d}', samp_category='negative control',
                   neg_cont_type='extraction negative', pos_cont_type=NA_C, concentration=MISS)
        out.append((rec, assays, 'lab'))
    for a in meta.get('assays') or []:
        for i in range(1, int(meta.get('pcr_negatives_per_assay') or 0) + 1):
            rec = dict(detected_notDetected=NA, samp_name=f'PCRNEG_{a["marker"]}_{i:02d}', samp_category='negative control',
                       neg_cont_type='PCR negative', pos_cont_type=NA_C, assay_name=a['assay_name'],
                       nucl_acid_ext=NA_C, nucl_acid_ext_kit=NA_C, nucl_acid_ext_lysis=NA_C, nucl_acid_ext_sep=NA_C,
                       nucl_acid_ext_modify=NA_C, concentration=NA_C, concentration_unit=NA_C, concentration_method=NA_C)
            out.append((rec, [a['assay_name']], 'lab'))
    return out


def fill_workbook(template_path, meta: dict, report_filename: str, output_path) -> dict:
    with zipfile.ZipFile(template_path) as z:
        pkg = {i.filename: z.read(i.filename) for i in z.infolist()}
        infos = z.infolist()
    ss = shared_strings(pkg)
    paths = sheet_paths(pkg)
    VOCAB.clear(); VOCAB.update(load_vocab(pkg, paths, ss))
    ver = re.search(r'checklist_v([\d.]+\d)', ' '.join(ss))
    version = ver.group(1) if ver else '1.0'

    # projectMetadata
    p = paths['projectMetadata']; xml = pkg[p].decode('utf-8')
    rows = parse_rows(xml)
    filled = fill_project(rows, ss, meta, version)
    pkg[p] = replace_sheet_data(xml, rows).encode('utf-8')

    # sampleMetadata
    p = paths['sampleMetadata']; xml = pkg[p].decode('utf-8')
    rows = parse_rows(xml)
    hdr, req = header_map(rows, ss)
    rows = {k: v for k, v in rows.items() if k <= 3}
    samples = build_samples(meta)
    for i, (rec, _, control) in enumerate(samples):
        write_record(rows, 4 + i, rec, hdr, req, control)
    pkg[p] = replace_sheet_data(xml, rows).encode('utf-8')

    # experimentRunMetadata
    run_for = {}
    for r in meta.get('runs') or []:
        for mk in r.get('assays') or []: run_for[mk] = r['seq_run_id']
    p = paths['experimentRunMetadata']; xml = pkg[p].decode('utf-8')
    rows = parse_rows(xml)
    hdr, req = header_map(rows, ss)
    rows = {k: v for k, v in rows.items() if k <= 3}
    marker_of = {a['assay_name']: a['marker'] for a in meta.get('assays') or []}
    rn = 4
    for rec, assays, control in samples:
        for an in assays:
            write_record(rows, rn, {
                'samp_name': rec['samp_name'], 'assay_name': an, 'lib_id': f'{rec["samp_name"]}_{an}',
                'seq_run_id': run_for.get(marker_of.get(an)), 'lib_conc_meth': meta.get('lib_conc_meth'),
                'phix_perc': meta.get('phix_perc'),
            }, hdr, req)
            rn += 1
    n_runs = rn - 4
    pkg[p] = replace_sheet_data(xml, rows).encode('utf-8')

    # README notes
    if 'README' in paths:
        p = paths['README']; xml = pkg[p].decode('utf-8')
        rows = parse_rows(xml)
        start = (max(rows) if rows else 0) + 2
        notes = ['Filled by FAIRe Metadata Filler on ' + time.strftime('%Y-%m-%d') + f' from: {report_filename}',
                 f'Extraction method: {meta.get("_method", "rule-based")}',
                 'Sample IDs, library IDs and sequencing-run IDs were generated from the report structure – replace them with your lab IDs.',
                 'Values the report does not state are entered as "missing: not provided"; please review them.'] + \
                [f'Note: {n}' for n in meta.get('notes') or []]
        for i, t in enumerate(notes):
            rows[start + i] = {'attrs': f' r="{start + i}"', 'cells': {'A': cell_xml(f'A{start + i}', t)}}
        pkg[p] = replace_sheet_data(xml, rows).encode('utf-8')

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info, pkg[info.filename])

    return {
        'projectTitle': meta.get('project_title'), 'projectId': project_values(meta, version)['project_id'],
        'institution': meta.get('institution'), 'lakes': len(meta.get('sites') or []),
        'rounds': len(meta.get('rounds') or []), 'assays': [a['assay_name'] for a in meta.get('assays') or []],
        'sampleRows': len(samples), 'experimentRows': n_runs, 'projectTermsFilled': filled,
        'notes': meta.get('notes') or [],
    }


# ════════════════════════════════════════════════════════════════════════════
#  4. Claude-assisted extraction (same schema as the rule-based extractor)
# ════════════════════════════════════════════════════════════════════════════

CLAUDE_MODEL = os.environ.get('FAIRE_CLAUDE_MODEL', 'claude-haiku-5-5')
MAX_REPORT_CHARS = 120_000

CLAUDE_PROMPT = """You extract FAIR eDNA (FAIRe) metadata from an eDNA project report.
Return ONLY a JSON object (no prose, no code fences). Use null for anything the report does not state.
Never guess: do not fill primer sequences, coordinates, database versions or names that are not in the text.

Keys (all optional except assays):
project_title, institution, institution_ror, project_contact, recorded_by, start_date (YYYY-MM-DD), end_date (YYYY-MM-DD),
funder, study_factor, sterilise_method, license, access_rights, bibliographicCitation, code_repo,
neg_controls (true/false), pos_controls (true/false),
field_blanks_per_site_round (int), extraction_blanks (int, total), pcr_negatives_per_assay (int),
samples_per_site_round (int), samp_size (number), samp_size_unit ("mL"|"L"), samp_collect_method,
env_broad_scale, env_local_scale, env_medium   (ENVO "label [ENVO:00000000]" only if you are certain of the ID),
habitat_natural (true for natural habitats), country (INSDC country), region,
samp_store_dur (ISO 8601 duration e.g. PT24H), samp_store_loc, samp_store_method_additional, samp_mat_process,
filter_active (true if pump/vacuum), size_frac (pore size µm, number), filter_diameter (mm, number), filter_material,
prepped_samp_store_temp (°C number), prepped_samp_store_sol,
ext_kit, nucl_acid_ext (reference/DOI for the extraction protocol), ext_lysis, ext_sep, ext_modifications, conc_method,
nucl_acid_amp (reference/DOI for the PCR protocol),
barcoding_pcr_appr ("one-step PCR"|"two-step PCR"|"ligation-based"), pcr2_amplificationReactionVolume, pcr2_commercial_mm,
pcr2_custom_mm, pcr2_dna_vol, pcr2_cond,
platform (e.g. "ILLUMINA", "OXFORD_NANOPORE"), instrument, lib_layout ("paired end"|"single end"), adapter_forward,
adapter_reverse, lib_screen, lib_conc_meth, phix_perc (number), sequencing_location,
sop_bioinformatics, trim_method, demux_tool, merge_tool, error_rate_tool, error_rate_type ("Phred score"|"expected error rate"),
error_rate_cutoff (number), chimera_check_method, chimera_check_param, otu_clust_tool, otu_clust_cutoff (number, 100 for ASVs),
screen_contam_method, otu_raw_description, bioinfo_method_additional, biological_rep (int),
notes: [short strings about inconsistencies or information that must be added by hand],
assays: [{marker ("16S","18S","12S","COI","rbcL","ITS",...), assay_name (short id, no spaces),
          target_gene, targetTaxonomicAssay, targetTaxonomicScope, ampliconSize,
          primer_fwd, primer_rev (5'-3' sequences ONLY if printed in the text), primer_name_fwd, primer_name_rev,
          primer_ref_fwd, primer_ref_rev, pcr_primer_vol (µL), pcr_primer_conc (stock µM),
          commercial_mm, custom_mm, amplificationReactionVolume (µL), pcr_dna_vol (µL), pcr_rep (int),
          pcr_cond, annealingTemp (°C), pcr_cycles (int), seq_kit, trim_param, otu_db, tax_assign_cat
          ("sequence similarity"|"sequence composition"|"phylogeny"|"probabilistic"), otu_seq_comp_appr,
          otu_final_description}],
sites: [{name, code (short id, no spaces), latitude, longitude, attributes}],
rounds: [{code ("R1","R2",...), iso_date (YYYY-MM-DD or YYYY-MM-DD/YYYY-MM-DD), verbatim, avg_vol_ml}],
runs: [{seq_run_id, date, assays: [markers sequenced on this run]}]

REPORT:
"""


def parse_with_claude(text: str, api_key: str) -> dict:
    try:
        import anthropic  # type: ignore
    except ImportError:
        raise RuntimeError('The anthropic package is not installed (pip install anthropic).')
    client = anthropic.Anthropic(api_key=api_key)
    msg = client.messages.create(model=CLAUDE_MODEL, max_tokens=8000,
                                 messages=[{'role': 'user', 'content': claude_prompt(text)}])
    raw = ''.join(b.text for b in msg.content if getattr(b, 'type', '') == 'text')
    return json_from_reply(raw)


def json_from_reply(raw: str) -> dict:
    """Parse the JSON object in a Claude reply (tolerates code fences / stray prose)."""
    raw = re.sub(r'^```(?:json)?\s*|\s*```$', '', raw.strip())
    return json.loads(raw[raw.find('{'): raw.rfind('}') + 1])


def claude_prompt(text: str) -> str:
    return CLAUDE_PROMPT + text[:MAX_REPORT_CHARS]


def facts_from_claude(data: dict) -> dict:
    """Tidy Claude's JSON into the same shape the rule-based extractor returns."""
    meta = {k: v for k, v in (data or {}).items() if v not in (None, '', [])}
    for key in ('assays', 'sites', 'rounds', 'runs', 'notes'):
        meta[key] = [x for x in meta.get(key) or [] if x]
    for a in meta['assays']:
        a['marker'] = canon_marker(str(a.get('marker') or a.get('assay_name') or 'assay'))
        a['assay_name'] = re.sub(r'\s+', '_', str(a.get('assay_name') or a['marker']))
        if not a.get('target_gene') and a['marker'] in TARGET_GENE: a['target_gene'] = TARGET_GENE[a['marker']]
        for k, v in list(a.items()):
            if isinstance(v, (int, float)) and not isinstance(v, bool): a[k] = f'{v:g}'
    for s in meta['sites']:
        s['code'] = re.sub(r'\W+', '_', str(s.get('code') or s.get('name') or 'SITE')).strip('_')
    for i, r in enumerate(meta['rounds']):
        r.setdefault('code', f'R{i + 1}')
    for k, v in list(meta.items()):
        if isinstance(v, (int, float)) and not isinstance(v, bool): meta[k] = f'{v:g}'
    for k in ('field_blanks_per_site_round', 'extraction_blanks', 'pcr_negatives_per_assay', 'samples_per_site_round'):
        try: meta[k] = int(float(meta.get(k) or 0))
        except (TypeError, ValueError): meta[k] = 0
    if meta.get('samples_per_site_round') and not meta.get('biological_rep'):
        meta['biological_rep'] = str(meta['samples_per_site_round'])
    meta['_method'] = f'Claude ({CLAUDE_MODEL})'
    return meta


def parse_report_regex(text: str) -> dict:
    """Rule-based extraction (kept under its old name for the Colab notebook)."""
    meta = extract(text)
    meta['_method'] = 'rule-based'
    return meta


# ════════════════════════════════════════════════════════════════════════════
#  5. Web server
# ════════════════════════════════════════════════════════════════════════════

def safe_name(name):
    return re.sub(r'[^A-Za-z0-9._-]+', '_', name).strip('_') or 'file'


def parse_multipart(content_type, body):
    m = re.search(r'boundary=(?P<b>[^;]+)', content_type)
    if not m: raise ValueError('Missing multipart boundary.')
    boundary = m.group('b').strip().strip('"').encode()
    parts = {}
    for raw in body.split(b'--' + boundary):
        raw = raw.strip(b'\r\n')
        if not raw or raw == b'--': continue
        if raw.endswith(b'--'): raw = raw[:-2].rstrip(b'\r\n')
        head_blob, _, payload = raw.partition(b'\r\n\r\n')
        headers = head_blob.decode('utf-8', errors='replace')
        disp = re.search(r'Content-Disposition: form-data;\s*(.*)', headers, re.I)
        if not disp: continue
        nm = re.search(r'name="([^"]+)"', disp.group(1))
        fn = re.search(r'filename="([^"]*)"', disp.group(1))
        if not nm: continue
        parts[nm.group(1)] = (fn.group(1) if fn else nm.group(1), payload)
    return parts


def process_upload(report_name: str, report_bytes: bytes, template_bytes: bytes, api_key: str = '') -> dict:
    OUTPUT_DIR.mkdir(exist_ok=True)
    job = uuid.uuid4().hex[:10]
    with tempfile.TemporaryDirectory() as tmp:                 # uploads are never kept
        rp = Path(tmp) / safe_name(Path(report_name or 'report.txt').name)
        tp = Path(tmp) / 'template.xlsx'
        rp.write_bytes(report_bytes); tp.write_bytes(template_bytes)
        text = extract_report_text(rp)
        if len(text.strip()) < 200:
            raise ValueError('Very little text could be read from the report – is it a scanned PDF?')
        method = 'rule-based'
        if api_key:
            try:
                meta = facts_from_claude(parse_with_claude(text, api_key)); method = 'claude-ai'
            except Exception as e:
                meta = parse_report_regex(text)
                meta['notes'].append(f'Claude extraction failed ({str(e)[:120]}); rule-based extraction used instead')
        else:
            meta = parse_report_regex(text)
        out_name = f'FAIRe_filled_{safe_name(project_values(meta, "x")["project_id"] or "project")}_{job}.xlsx'
        summary = fill_workbook(tp, meta, Path(report_name).name, OUTPUT_DIR / out_name)
    summary.update(downloadUrl=f'/download/{out_name}', outputName=out_name, parsingMethod=method)
    return summary


class Handler(BaseHTTPRequestHandler):
    def _cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def do_OPTIONS(self):
        self.send_response(204); self._cors(); self.end_headers()

    def do_GET(self):
        path = unquote(self.path.split('?', 1)[0])
        if path == '/':
            self._serve(STATIC_DIR / 'index.html', 'text/html; charset=utf-8')
        elif path.startswith('/static/'):
            tgt = (STATIC_DIR / path[len('/static/'):]).resolve()
            if STATIC_DIR.resolve() not in tgt.parents: return self.send_error(404)
            ct = {'.css': 'text/css', '.js': 'application/javascript', '.png': 'image/png', '.svg': 'image/svg+xml'}.get(tgt.suffix, 'text/html; charset=utf-8')
            self._serve(tgt, ct)
        elif path.startswith('/download/'):
            self._serve(OUTPUT_DIR / safe_name(path[len('/download/'):]),
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', attachment=True)
        else:
            self.send_error(404)

    do_HEAD = do_GET

    def do_POST(self):
        if self.path != '/api/fill': return self.send_error(404)
        try:
            length = int(self.headers.get('Content-Length', '0'))
            form = parse_multipart(self.headers.get('Content-Type', ''), self.rfile.read(length))
            if 'report' not in form or 'template' not in form:
                raise ValueError('Please upload both a report file and the FAIRe template workbook.')
            key_field = form.get('apiKey', ('', b''))[1].decode('utf-8', errors='replace').strip()
            api_key = key_field or os.environ.get('ANTHROPIC_API_KEY', '')
            self._json(process_upload(form['report'][0], form['report'][1], form['template'][1], api_key))
        except Exception as exc:
            self._json({'error': str(exc)}, 500)

    def _serve(self, target, content_type, attachment=False):
        if not Path(target).is_file(): return self.send_error(404)
        data = Path(target).read_bytes()
        self.send_response(200)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(data)))
        self._cors()
        if attachment: self.send_header('Content-Disposition', f'attachment; filename="{Path(target).name}"')
        self.end_headers()
        if self.command != 'HEAD': self.wfile.write(data)

    def _json(self, payload, status=200):
        data = json.dumps(payload).encode()
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(data)))
        self._cors(); self.end_headers(); self.wfile.write(data)

    def log_message(self, fmt, *args):
        print(f'{self.address_string()} {html.escape(fmt % args)}')


def main():
    host = os.environ.get('FAIRE_TOOL_HOST', '127.0.0.1')
    port = int(os.environ.get('FAIRE_TOOL_PORT', '8787'))
    print(f'FAIRe Metadata Filler at http://{host}:{port}')
    print('  Claude extraction:', 'enabled' if os.environ.get('ANTHROPIC_API_KEY') else 'off (paste a key in the page, or set ANTHROPIC_API_KEY)')
    ThreadingHTTPServer((host, port), Handler).serve_forever()


if __name__ == '__main__':
    main()
