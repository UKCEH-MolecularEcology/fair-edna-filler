#!/usr/bin/env python3
"""
FAIRe Metadata Filler – fills the FAIRe eDNA metadata template from project reports.

    python3 server.py                         # open the tool in your browser (served from this computer)
    python3 server.py report.pdf              # fill the template from one report  -> outputs/
    python3 server.py uploads/                # every report in a folder           -> outputs/
    python3 server.py report.doc -t my_template.xlsx -o results/
    python3 server.py report.doc --edits outputs/report_review.csv --save-profile lab.json   # reuse your corrections
    python3 server.py uploads/ --profile lab.json --local-ai    # + second opinion from a local model (Ollama)

Each workbook comes with <name>_review.csv: values that need a look first, then everything accepted automatically,
each with the report sentence it came from. Put corrections in its 'your_value' column and rerun with --edits.
Reports are read on this computer only.

The same functions are used by index.html (in the browser, via Pyodide) and the Colab notebook:
    extract_report_text(path) -> str
    parse_report_regex(text)  -> metadata dict
    apply_library(meta)       -> primers/databases checked against published ones
    review(meta, template)    -> every value with its source sentence and a status: ok / check / missing
    apply_review(meta, edits); profile_from_review(...); apply_profile(meta, profile)
    fill_workbook(template, metadata, report_name, output) -> summary dict
Optional, only for non-confidential reports: parse_with_claude(text, key) -> facts_from_claude(data, text)
"""
from __future__ import annotations

import csv, html, io, json, os, posixpath, re, shutil, struct, subprocess, tempfile, time, zipfile
import xml.etree.ElementTree as ET
from datetime import date, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

APP_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = APP_DIR / 'outputs'
UPLOAD_DIR = APP_DIR / 'uploads'     # reserved for batch processing (multiple reports / a folder)

# FAIRe FULL template this tool is tested against. Updated by .github/workflows/check-template.yml
TOOL_VERSION = '1.1'
TEMPLATE_VERSION = '1.0.2'
TEMPLATE_FILE = f'FAIRe_checklist_v{TEMPLATE_VERSION}_FULLtemplate.xlsx'
TEMPLATE_URL = f'https://github.com/FAIR-eDNA/FAIRe_checklist/raw/main/{TEMPLATE_FILE}'   # official source
TEMPLATE_PATH = APP_DIR / 'template' / TEMPLATE_FILE                                          # bundled copy


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


PDF_NOTES: list = []      # pages that could not be read, reported to the user


def _quiet_pdf_logs():
    import logging
    for name in ('pypdf', 'pypdf._cmap', 'pypdf.generic', 'pdfminer'): logging.getLogger(name).setLevel(logging.ERROR)


def _extract_pdf(path: Path) -> str:
    _quiet_pdf_logs()
    """Text of a PDF, page by page. Pages too complex to read (e.g. huge vector figures) are skipped."""
    PDF_NOTES.clear()
    try:
        from pypdf import PdfReader  # type: ignore
    except ImportError:
        PdfReader = None
    if PdfReader is not None:
        reader = PdfReader(str(path))
        pages, skipped = [], []
        for n, page in enumerate(reader.pages, 1):
            try:
                pages.append(page.extract_text() or '')
            except Exception:
                skipped.append(n); pages.append('')
        if skipped:
            PDF_NOTES.append(f'PDF page(s) {", ".join(map(str, skipped))} could not be read (too complex, probably a figure) and were skipped')
        return '\n'.join(pages)
    try:
        import pdfplumber  # type: ignore
        with pdfplumber.open(path) as pdf:
            out = []
            for page in pdf.pages:
                out.append(page.extract_text() or ''); page.close()
            return '\n'.join(out)
    except ImportError:
        tool = shutil.which('pdftotext')                     # poppler, if installed
        if tool:
            r = subprocess.run([tool, '-layout', str(path), '-'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
            if r.returncode == 0 and r.stdout.strip(): return r.stdout.decode('utf-8', 'replace')
        raise ValueError('reading PDFs needs the pypdf package. Install it once with:  python3 -m pip install --user pypdf  '
                         '(or: pip install -r requirements.txt)')


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



_STOPWORDS = {
    'en': ('English', 'the and of to in is were was with for that from by on this are using as'),
    'fr': ('French', 'le la les des et du en est une dans pour par sur avec que qui au aux été sont'),
    'de': ('German', 'der die das und den des ist mit von für auf im dem zu wurden wurde ein eine nicht'),
    'es': ('Spanish', 'el la los las de y en del que con por para una se fue fueron como al'),
    'it': ('Italian', 'il lo la gli le di e del della che con per sono stato una nel dei delle'),
    'pt': ('Portuguese', 'o a os as de e do da dos das que com para em uma foi foram por no na'),
    'nl': ('Dutch', 'de het een en van in is met voor op dat zijn werd werden door bij uit'),
    'sv': ('Swedish', 'och att det som en är av för med till på den har var inte från'),
    'da': ('Danish', 'og at det som en er af for med til på den har var ikke fra blev'),
    'no': ('Norwegian', 'og at det som en er av for med til på den har var ikke fra ble'),
    'fi': ('Finnish', 'ja on oli että ei se kun ovat olivat myös joka mukaan sekä'),
    'pl': ('Polish', 'i w na z do że się jest nie to od przez dla oraz był były'),
}


def detect_language(text: str):
    """Rough language guess from common words: returns (code, name). Runs locally."""
    sample = text[:60000]
    if len(re.findall(r'[぀-ヿ]', sample)) > 50: return 'ja', 'Japanese'
    if len(re.findall(r'[一-鿿]', sample)) > 200: return 'zh', 'Chinese'
    if len(re.findall(r'[Ѐ-ӿ]', sample)) > 200: return 'ru', 'Russian (Cyrillic)'
    words = re.findall(r'[a-zà-ÿąćęłńóśźż]+', sample.lower())
    if len(words) < 50: return 'en', 'English'
    counts = {}
    for w in words: counts[w] = counts.get(w, 0) + 1
    scores = {code: sum(counts.get(w, 0) for w in sw.split()) for code, (_, sw) in _STOPWORDS.items()}
    best = max(scores, key=scores.get)
    if scores['en'] >= 0.6 * scores[best]: best = 'en'      # mixed text (e.g. English report with foreign references)
    return best, _STOPWORDS[best][0]


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


# ── targeted assays (qPCR / ddPCR / dPCR) and oligonucleotide tables ─────────
TMARKER_RE = r'(?:mt-?)?(?:ND[1-6]L?|nd[1-6]|D-?[Ll]oop|d-?loop|COIII|COII|CytB|cytb|Cyt ?b|COI|CO1|12S|16S|18S|rbcL|ITS2?)'
GENE_PHRASES = [
    (r'cytochrome[- ]b\b', 'CytB'), (r'cytochrome c oxidase (?:subunit )?(?:III|3)\b', 'COIII'),
    (r'cytochrome c oxidase (?:subunit )?(?:II|2)\b', 'COII'), (r'cytochrome c oxidase(?: (?:subunit )?(?:I|1)\b)?', 'COI'),
    (r'NADH[- ]dehydrogenase (?:subunit )?([1-6])', 'ND'), (r'\bcontrol region\b', 'D-loop'),
    (r'\b12S rRNA\b|mt-?RNR1', '12S'), (r'\b16S rRNA\b|mt-?RNR2', '16S')]
for _g in ('ND1', 'ND2', 'ND3', 'ND4', 'ND4L', 'ND5', 'ND6', 'D-loop', 'COII', 'COIII'): TARGET_GENE.setdefault(_g, _g)
TARGETED_WORDS = re.compile(r'\bqPCR\b|\bddPCR\b|\bdPCR\b|digital PCR|droplet digital|real[- ]time (?:PCR|quantitative)|quantitative PCR|'
                            r'TaqMan|\bprobes?\b|\bC[tq] values?\b|species[- ]specific (?:assay|qPCR|PCR)|single[- ]species (?:assay|qPCR)', re.I)
SEQ_WORDS = re.compile(r'metabarcod|amplicon sequencing|high[- ]throughput sequencing|\bIllumina\b|MiSeq|NovaSeq|NextSeq|Nanopore|'
                       r'sequencing librar|library prep', re.I)
DYES = r'(?:6-?FAM|FAM|HEX|VIC|NED|TET|JOE|ROX|Cy5|Cy3|ABY|JUN|Texas Red|ATTO ?\d{3})'
QUENCHERS = r'(?:BHQ-?[123]|MGB-?NFQ|NFQ-?MGB|MGB|NFQ|QSY\d*|ZEN|3IABkFQ|IBFQ|Iowa Black(?: ?FQ)?|TAMRA|BBQ|Dabcyl|DDQ)'
QPCR_INSTRUMENTS = (r'((?:Bio-?Rad )?CFX ?(?:96|384|Opus|Connect|Duet)(?: Touch)?|QuantStudio ?\d+(?: Flex)?|StepOne(?:Plus)?|'
                    r'(?:Applied Biosystems )?7[59]00(?: Fast)?(?: Real-Time PCR System)?|LightCycler ?\d+|QX ?(?:200|ONE|600)|QIAcuity(?: One| Four| Eight| digital PCR system| dPCR system| instrument)|'
                    r'Rotor-Gene(?: Q)?|Mic qPCR|AriaMx|Mx3005P|ViiA ?7|Naica|Absolute Q)(?! Probe| PCR [Kk]it| [Kk]it)')


def canon_gene(tok: str) -> str:
    t = re.sub(r'^mt-?', '', tok.strip(), flags=re.I)
    u = t.upper().replace(' ', '')
    if re.fullmatch(r'ND[1-6]L?', u): return u
    if re.fullmatch(r'D-?LOOP', u): return 'D-loop'
    if u in ('COII', 'COIII'): return u
    if u == 'CYTB': return 'CytB'
    return canon_marker(t)


def genes_in(s: str) -> list[str]:
    """Genes named in a passage (abbreviations and full names), in order of appearance."""
    hits = [(m.start(), canon_gene(m.group(0))) for m in re.finditer(r'(?<![A-Za-z0-9])' + TMARKER_RE + r'(?![A-Za-z0-9])', s)]
    for pat, g in GENE_PHRASES:
        hits += [(m.start(), 'ND' + m.group(1) if g == 'ND' else g) for m in re.finditer(pat, s, re.I)]
    out = []
    for _, g in sorted(hits):
        if g not in out: out.append(g)
    return out


OLIGO_LABEL = re.compile(r'(?:^|[\s(|:;,])(forward|fwd|reverse|rev|probe)\b', re.I)
GENERIC_WORDS = re.compile(r'^(?:forward|reverse|fwd|rev|primers?|probes?|sequences?|oligo(?:nucleotide)?s?|names?|target|assay|'
                           r'and|or|the|of|with|primer|component|5|3)$', re.I)


def oligo_records(text: str) -> list[dict]:
    """Primer and probe sequences with role, name, dyes and gene, from sentences and from table rows.

    Works on the raw text with its line breaks, so table rows such as
    'TCCB.probe  FAM-CATCC…-BHQ1' or 'Forward Primer: CGAAGG…' keep their labels."""
    IUPAC = 'ACGTRYKMSWBDHVNI'
    t = text.replace('−', '-').replace('–', '-').replace('′', "'").replace('’', "'")
    for _ in range(4):
        t = re.sub(r'\b([' + IUPAC + r']{8,})[ \t]*\n[ \t]*([' + IUPAC + r']{2,})\b(?![ \t]+[' + IUPAC + r']{12,})(?![.\-]?[a-z])', r'\1\2', t)
    out, seen = [], set()
    for m in re.finditer(r'(?<![A-Za-z0-9])([' + IUPAC + r']{15,45})(?![A-Za-z0-9])', t):
        seq = m.group(1)
        if sum(c in 'ACGT' for c in seq) < 0.6 * len(seq): continue
        ls = t.rfind('\n', 0, m.start()) + 1
        le = t.find('\n', m.end()); le = len(t) if le < 0 else le
        line_before, line_after = t[ls:m.start()], t[m.end():le]
        prev_lines = t[max(0, t.rfind('\n', 0, max(0, ls - 1)) + 1): ls]
        pre, post = t[max(0, m.start() - 30): m.start()], t[m.end(): m.end() + 30]
        rep_ = re.search(DYES + r"\W{0,4}(?:5'?\W{0,3})?$", pre)
        q = re.search(r"^\W{0,4}(?:3'?\W{0,3})?" + QUENCHERS, post) or re.search(r"^\W{0,4}" + QUENCHERS, post)
        label_zone = line_before[-60:] if line_before.strip() else prev_lines[-60:]
        local = re.split(r"[;,(]|\s{2,}|\t", line_before)[-1]               # the table cell / phrase just before the sequence
        name = None
        for tok in reversed(re.findall(r"[A-Za-z0-9][\w.\-/]*", re.sub(DYES + r"\W*$|5'\W*$", '', local))):
            if GENERIC_WORDS.match(tok) or len(tok) > 25 or re.fullmatch(r'[ACGTRYKMSWBDHVNI]{8,}\W*\d*', tok): continue
            if re.search(r'\d', tok) or len(re.findall(r'[A-Z]', tok)) >= 2: name = tok; break
        role = None
        if rep_ or q or re.search(r'probe', label_zone, re.I) or (name and re.search(r'(?:[._\-]probe|[._\-]?P(?:r|rb)?)$', name, re.I) and not re.search(r'primer', label_zone, re.I)):
            role = 'probe'
        else:
            lab = [x.group(1).lower() for x in OLIGO_LABEL.finditer(label_zone)]
            if lab: role = 'fwd' if lab[-1] in ('forward', 'fwd') else 'rev' if lab[-1] in ('reverse', 'rev') else None
            if role is None and name:
                if re.search(r'(?:F|Fw|Fwd|For|L)$|[._\-]f$', name): role = 'fwd'
                elif re.search(r'(?:R|Rv|Rev)$|[._\-]r$', name): role = 'rev'
        window = t[max(0, m.start() - 600): m.start()]
        tw, sw = list(TARGETED_WORDS.finditer(window)), list(SEQ_WORDS.finditer(window))
        targeted = role == 'probe' or (bool(tw) and (not sw or tw[-1].start() > sw[-1].start()))
        genes = genes_in(window[-400:] + line_before + line_after)
        label = None
        for ln in reversed(prev_lines.strip().splitlines()[-3:] if prev_lines.strip() else []):
            mm = re.match(r'\s*([A-Za-z][\w.\-]*\d[\w.\-]*)\s*(?:\(|$)', ln)
            if mm and not re.search('[' + IUPAC + ']{15,}', ln): label = mm.group(1); break
        key = (seq, role)
        if key in seen: continue
        seen.add(key)
        out.append({'seq': seq, 'role': role, 'name': name, 'reporter': rep_.group(0).strip(" -/5'") if rep_ else None,
                    'quencher': q.group(0).strip(" -/3'") if q else None, 'gene': genes[-1] if genes else None,
                    'targeted': targeted, 'label': label, 'pos': m.start(),
                    'src': norm(t[max(0, m.start() - 160): m.end() + 40])})
    return out


def _same_site(x: str, y: str) -> bool:
    """Two oligos bind the same site (≥70 % identical, either strand)."""
    comp = str.maketrans('ACGTRYKMBDHVN', 'TGCAYRMKVHDBN')
    if not x or not y: return x == y
    for z in (y, y.translate(comp)[::-1]):
        if abs(len(x) - len(z)) <= 2:
            n = min(len(x), len(z))
            if sum(a_ == b_ for a_, b_ in zip(x[:n], z[:n])) >= 0.7 * n: return True
    return False


def oligo_groups(recs: list[dict]) -> list[dict]:
    """Group primer/probe records into assays: forward + reverse (+ probe)."""
    # specificity / alignment tables repeat the same primer sites for many species: keep the first of each family
    fam = [[j for j, o in enumerate(recs) if _same_site(r['seq'], o['seq'])] for r in recs]
    recs = [r for i, r in enumerate(recs) if len(fam[i]) < 3 or fam[i][0] == i or r.get('role')]
    groups = []
    for r in recs:
        g = groups[-1] if groups else None
        role = r['role'] or next((x for x in ('fwd', 'rev', 'probe') if g and not g.get(x) and (x != 'probe' or g.get('targeted') or r['targeted'])), 'fwd')
        if not g: role = r['role'] or 'fwd'
        new = g is None or g.get(role) or (r['label'] and g.get('label') and r['label'] != g['label']) or r['pos'] - g['last'] > 2500
        if new:
            g = {'label': r['label'], 'gene': r['gene'], 'targeted': r['targeted'], 'first': r['pos'], 'src': r['src']}
            groups.append(g)
        g[role] = r; g['last'] = r['pos']
        g['targeted'] = g['targeted'] or r['targeted'] or role == 'probe'
        g['gene'] = g.get('gene') or r['gene']
        g['label'] = g.get('label') or r['label']
    for g in groups:
        names = [g[k]['name'] for k in ('fwd', 'rev') if g.get(k) and g[k].get('name')]
        stem = None
        if len(names) == 2:
            a, b = names
            i = 0
            while i < min(len(a), len(b)) and a[i] == b[i]: i += 1
            stem = a[:i].rstrip('._-') if i >= 3 else None
        g['name'] = g.get('label') or stem or (g['probe']['name'].split('.')[0] if g.get('probe') and g['probe'].get('name') else None)
    groups = [g for g in groups if (g.get('fwd') and g.get('rev')) or (g.get('probe') and (g.get('fwd') or g.get('rev')))]
    # specificity / alignment tables list the same primer sites in many species: keep the first of each family
    similar = _same_site
    kept = []
    for g in groups:
        dup = next((k for k in kept if similar((g.get('fwd') or {}).get('seq'), (k.get('fwd') or {}).get('seq'))
                    and similar((g.get('rev') or {}).get('seq'), (k.get('rev') or {}).get('seq'))), None)
        if dup:                                                       # same assay again (or a specificity table row): merge
            for kk in ('gene', 'name', 'label', 'probe'):
                if g.get(kk) and not dup.get(kk): dup[kk] = g[kk]
            dup['targeted'] = dup['targeted'] or g['targeted']
            continue
        kept.append(g)
    return kept


def pipe_rows(text: str) -> dict:
    """Assay-per-column table rows -> {label (lower): [a, b, c]}.

    'Label: a | b | c', 'Label<TAB>a<TAB>b' and 'Label  a  b  c' (two or more spaces) are read."""
    rows = {}
    for m in re.finditer(r'^[ \t]*([A-Za-z][^:\n|\t]{1,40}?)(?::[ \t]*|\t+| {2,})([^\n]+)$', text, re.M):
        cells = re.split(r'\s*\|\s*|\t+| {2,}', m.group(2).strip())
        if len(cells) >= 2: rows.setdefault(m.group(1).strip().lower(), [c for c in cells if c])
    return rows


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

def strip_line_numbers(text: str) -> str:
    """Remove manuscript line numbers (bioRxiv / journal submissions number every line, at the end or the start)."""
    lines = text.split('\n')
    for where in ('end', 'start'):
        pat = re.compile(r'(\d{1,4})\s*$') if where == 'end' else re.compile(r'^\s*(\d{1,4})(?=\s|$)')
        found = [(i, int(m.group(1))) for i, l in enumerate(lines) if (m := pat.search(l))]
        # line numbers form a long, almost continuous increasing run
        run = sum(1 for (i1, a), (i2, b) in zip(found, found[1:]) if 0 < b - a <= 3)
        if len(found) > 40 and run > 0.6 * len(found):
            expected, out = None, []
            for l in lines:
                m = pat.search(l)
                if m and (expected is None or expected - 2 <= int(m.group(1)) <= expected + 6):
                    expected = int(m.group(1)) + 1
                    l = (l[:m.start()] if where == 'end' else l[m.end():]).rstrip() if where == 'end' else l[m.end():]
                out.append(l)
            return '\n'.join(out)
    return text


# ── front matter: title, authors, contact, institution (reports and papers) ──
FM_LABEL = (r'(?:Authors?(?:\(s\))?|(?:Natural England|Environment Agency.s) Project [Mm]anager|(?:Research )?Contractors?|Keywords|'
            r'Acknowledge?ments|Citation|Dissemination Status|Project Number|Report details|Foreword|Further Information|Copyright)')
JUNK_ANY = re.compile(r'Page \d+ of \d+|\[Report title|\bdoi\b|doi\.org|https?://|www\.|©|\b(?:19|20)\d\d\s*[;,]\s*\d+|'
                      r'\b(?:Received|Revised|Accepted|Published|Keywords|Citation)\b\s*:|Commissioned Report|^\d+ of \d+|\d of \d+\s*$|'
                      r'\bet al\.?\s*$|\((?:19|20)\d\d\)\s*\d+\s*:\s*\d+|wileyonlinelibrary|journal/|Open Access|licen[cs]e|^\W*$', re.I)
JUNK_FULL = re.compile(r'(?:(?:RESEARCH|RESOURCE|ORIGINAL|REVIEW|METHODS?)\s*)?(?:ARTICLE|RESEARCH|PAPER)(?:\s+OPEN ACCESS)?|Article|'
                       r'(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}|'
                       r'Report\s*[–-].*|[A-Z]{2,}[\d/]+\S*|[ivx]+|\d+|(?:Edition|Version)\b.*|Natural England|Environment Agency|'
                       r'Evidence Project Final Report|Project identification', re.I)
ORG_WORDS = (r'University|Universit\w+|Institut\w*|Centre|Center|Agency|Administration|Organi[sz]ation|Survey|Ltd\.?|Limited|'
             r'Museum|Council|Commission|Authority|Laborator\w+|College|Foundation|Cefas|CSIRO|NOAA|Division|Society|Trust|plc|LLP|Inc\.?|'
             r'Consultancy|Consulting')
GENERIC_EMAIL = re.compile(r'^(?:enquiries|info|contact|admin|office|help|support|noreply|strategicevidence)@', re.I)
FUNCTION_WORDS = re.compile(r'\b(?:of|the|in|for|on|to|with|from|a|an|by|using|via|as|at|into|across|through|is|are)\b')


def _junk(s):
    return bool(JUNK_ANY.search(s) or JUNK_FULL.fullmatch(s))


def _labelled_block(text, label, max_lines=8):
    """Lines under a stand-alone label such as 'Authors' or 'Contractor:' (not a table-of-contents line)."""
    for m in re.finditer(r'(?im)^[ \t]*' + label + r'[ \t]*:?[ \t]*$', text[:60000]):
        out = []
        for line in text[m.end():].split('\n')[1:max_lines + 1]:
            s = norm(line)
            if not s:
                if out: break
                continue
            if re.fullmatch(FM_LABEL + r'\s*:?', s): break
            out.append(s)
        if out: return out
    return []


def _org_from(line):
    """'1Department of X, University of Y, City, Country' -> 'University of Y'."""
    line = re.sub(r'^\s*[\d*†,]+\s*', '', line).rstrip(' .')
    if not re.search(r'\d', line) and re.search(ORG_WORDS, line) and len(line) < 120 and not re.match(r'(?:Department|School|Faculty)\b', line):
        return norm(line)
    parts = [p.strip() for p in re.split(r',|\|', line) if p.strip()]
    best = [p for p in parts if re.search(ORG_WORDS, p) and not re.match(r'(?:Department|Dept|School|Faculty|Group)\b', p)]
    org = (best or [p for p in parts if re.search(ORG_WORDS, p)] or [None])[0]
    return norm(org).rstrip(' ,.;') if org and len(org) < 120 else None


def _names(s):
    s = re.sub(r'\([^)]*@[^)]*\)|\S+@\S+', '', s)
    s = re.sub(r'(?<=[A-Za-zÀ-ÿ.])\s?\d+(?:\s?,\s?\d+)*[*†‡§]*|[*†‡§]+', '', s)
    out = []
    for p in re.split(r'\s*(?:,|\||&|\band\b)\s*', s):
        p = norm(re.sub(r'^(?:Dr\.?|Prof\.?|Professor)\s+', '', p.strip(' .;')))
        if re.fullmatch(r"[A-ZÀ-Ý][\w'’‑\-.]+(?:\s+[A-ZÀ-Ý][\w'’‑\-.]*){1,4}", p) and not re.search(ORG_WORDS, p) \
                and not re.fullmatch(r'[A-Z]\.?', p.split()[-1]): out.append(p)
    return out


def _authorish(s):
    return len(_names(s)) >= 2 and not FUNCTION_WORDS.search(re.sub(r'\([^)]*\)', '', s)) or \
           bool(re.match(r'\d+\s*[A-Z]', s) and re.search(ORG_WORDS, s))


def front_matter(text: str) -> dict:
    """Title, authors, contact e-mail and institution from the first pages of a report or paper."""
    head = text[:12000]
    fm = {}
    lines = [re.sub(r'^(?:[A-Z] ){3,}[A-Z]\b\s*', '', norm(l)) for l in head.split('\n')[:100]]   # "R E S E A R C H  A R T I C L E"
    block, start, end = [], None, None
    for i, s in enumerate(lines):
        words = len(' '.join(block).split())
        if not s or _junk(s):
            if block and words >= 4: end = i; break
            block = []; continue
        if _authorish(s):
            if words >= 4: end = i; break
            block = []; continue
        if not block: start = i
        block.append(s)
        if len(' '.join(block)) > 300: block = []
    title = re.sub(r'(\w)- (\w)', r'\1-\2', norm(' '.join(block))) if len(' '.join(block).split()) >= 4 else ''
    if title:
        rest = norm(text[:20000])      # a running page header repeating the title shows where a subtitle starts
        for m in list(re.finditer(re.escape(title[:30]), rest))[1:]:
            tw, rw = title.split(), rest[m.start(): m.start() + len(title) + 80].split()
            k = 0
            while k < min(len(tw), len(rw)) and tw[k].strip('.,:;') == rw[k].strip('.,:;'): k += 1
            if 4 <= k <= len(tw) - 3: title = ' '.join(tw[:k]); break   # only cut off a real subtitle
        fm['project_title'] = title.rstrip(' .:')
    auth = _labelled_block(text, r'Authors?(?:\(s\))?')
    if auth:
        one_per_line = all(not l.rstrip().endswith(',') and len(_names(l)) >= 1 and
                           _names(re.split(r',|&|\band\b', l)[-1]) for l in auth)
        names = _names((', ' if one_per_line else ' ').join(auth))
        if names: fm['recorded_by'], fm['_ev_authors'] = ' | '.join(names), 'Authors: ' + ' '.join(auth)
        em = re.search(r'([\w.+\-]+@[\w\-]+\.[\w.\-]*\w)', ' '.join(auth))
        if em: fm['project_contact'] = em.group(1)
    elif end is not None:
        by = []
        for s in lines[end:end + 6]:
            if not s or _junk(s): continue
            if re.match(r'\d+\s*[A-Z]', s) and re.search(ORG_WORDS, s): break
            if not _authorish(s) and not by: break
            by.append(s)
            if not re.search(r'[,|&]\s*$|\band\s*$|\b[A-Z]\.\s*$', s): break
        names = _names(' '.join(by))
        if names: fm['recorded_by'], fm['_ev_authors'] = ' | '.join(names), ' '.join(by)[:300]
    con = _labelled_block(text, r'(?:Research )?Contractors?')
    if con:
        org = next((o for o in map(_org_from, con) if o), None) or (con[0].split(',')[0] if len(con) <= 4 else None)
        if org: fm['institution'], fm['_ev_institution'] = org.rstrip('.'), 'Contractor: ' + ' '.join(con)[:300]
    if not fm.get('project_contact'):
        for m in re.finditer(r'(?:Correspond\w*|Corresponding author\w*|e-?mail)\s*:?\s*(.{0,160})', head, re.I | re.S):
            em = re.search(r'([\w.+\-]+@[\w\-]+\.[\w.\-]*\w)', m.group(1))
            if em and not GENERIC_EMAIL.match(em.group(1)):
                nm = re.search(r'([A-Z][\w.\- ]{3,40}?)\s*\(\s*' + re.escape(em.group(1)), m.group(1))
                fm['project_contact'] = (norm(nm.group(1)) + f' ({em.group(1)})') if nm else em.group(1)
                fm['_ev_contact'] = norm(m.group(0))[:200]; break
    if not fm.get('institution'):
        aff = re.search(r'(?ms)^\s*1\s*([A-Z].{8,300}?)(?=^\s*2\s*[A-Z]|^\s*(?:Correspond|Abstract|Funding|Keywords|Received)|\Z)', head)
        org = _org_from(norm(aff.group(1))) if aff else None
        if org: fm['institution'], fm['_ev_institution'] = org, norm(aff.group(0))
    return fm


HEADING = r'^\s*(?:\d+(?:\.\d+)*\.?\s*\|?\s*|[IVX]+\.\s*)?'
METHODS_H = re.compile(HEADING + r'(?:MATERIALS? (?:AND|&) METHODS?|Materials? (?:and|&) [Mm]ethods?|METHODS?|Methods?|METHODOLOGY|Methodology|Experimental (?:design|procedures))\s*$', re.M)
END_H = re.compile(HEADING + r'(?:RESULTS?(?: AND DISCUSSION)?|Results?(?: and [Dd]iscussion)?|DISCUSSION|Discussion|CONCLUSIONS?|Conclusions?|REFERENCES|References|ACKNOWLEDGE?MENTS?|Acknowledge?ments?)\s*$', re.M)


def methods_text(text: str):
    """The longest Methods section (heading to the next Results/Discussion/References heading), or None."""
    best = None
    for m in METHODS_H.finditer(text):
        e = END_H.search(text, m.end())
        span = text[m.end(): e.start() if e else len(text)]
        if len(span) > 1500 and (best is None or len(span) > len(best)): best = span
    return best


SEASONS = r'(autumn|spring|summer|winter|wet season|dry season)'


def _iso_range(a, b):
    da, db = parse_date(a), parse_date(b)
    if da and db and da > db: da = None
    if not da and db and re.match(r'\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+$', a.strip()):
        da = parse_date(a + ' ' + db[:4])
    return f'{da}/{db}' if da and db else (db or da)


def sample_plan(text: str, flat: str) -> list:
    """Sampling groups described in words, e.g. 'Troutbeck Fan (TFAN, N=4) ... each site was sampled five times'
    or 'Seven samples were collected at each of the four additional lakes: Rydal Water (1st December 2025), …'.
    Returns [{code, name, points, visits, date, campaign, evidence}]."""
    groups, sents_ = [], sentences(text)
    seasons = {}
    for m in re.finditer(SEASONS + r'[^().]{0,100}?\(\s*(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+(?:\s+\d{4})?)\s+(?:to|-|–|until)\s+'
                         r'(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4})', flat, re.I):
        seasons.setdefault(m.group(1).lower(), _iso_range(m.group(2), m.group(3)))
    lines = [norm(l) for l in text.split('\n')]
    def name_for(code, fallback):
        for i, l in enumerate(lines):
            if l == code and i >= 2:
                prev = [x for x in lines[max(0, i - 3):i] if x]
                cand = [x for x in prev if len(x) > 6 and not re.fullmatch(r'(?:North|South|East|West)(?: \w+)?', x)]
                if cand: return cand[-1]
        return fallback
    pending, campaign = [], None
    for sent in sents_:
        cm = re.search(SEASONS, sent, re.I)
        found = list(re.finditer(r'(?:\d\)\s*)?([A-Z][^,;()]{2,60}?|[a-z][^,;()]{2,60}?)\s*\(\s*([A-Z][A-Z0-9]{1,7}),\s*[Nn]\s*=\s*(\d+)'
                                 r'(?:\s+and\s+([A-Z][A-Z0-9]{1,7}),\s*[Nn]\s*=\s*(\d+))?\s*\)', sent))
        if found:
            campaign = cm.group(1).lower() if cm else campaign
            for m in found:
                for code, n in ((m.group(2), m.group(3)), (m.group(4), m.group(5))):
                    if not code: continue
                    fb = re.sub(r'^(?:the|and|including|\d\))\s+', '', norm(m.group(1)), flags=re.I)
                    g = {'code': code, 'name': name_for(code, fb), 'points': int(n), 'visits': None,
                         'date': seasons.get(campaign), 'campaign': campaign, 'evidence': sent}
                    groups.append(g); pending.append(g)
        vm = re.search(r'sampled\s+(\w+)\s+times', sent, re.I)
        if vm and pending:
            k = num_word(vm.group(1))
            for g in pending: g['visits'] = k
            for ex in re.finditer(r'exception of\s*[“"‘\']?([A-Z][A-Z0-9]{1,7})[”"’\']?\s*(?:,\s*)?which was sampled\s+(\w+)\s+times', sent, re.I):
                for g in pending:
                    if g['code'] == ex.group(1): g['visits'] = num_word(ex.group(2))
            for g in pending: g['evidence'] = g['evidence'] + ' ' + sent
            pending = []
    # "Seven samples were collected at each of the four additional lakes: Rydal Water (1st December 2025), …"
    for m in re.finditer(r'(\w+)\s+samples\s+were\s+(?:collected|taken)\s+(?:at|from)\s+each\s+of\s+the\s+(?:\w+\s+){0,3}?'
                         r'(?:lakes|sites|ponds|rivers|locations|tarns|reservoirs)\s*:?\s*([^.]{10,600})', flat, re.I):
        n = num_word(m.group(1))
        codes = []
        for cl in re.finditer(r'\(((?:[A-Z][A-Z0-9]{2,7}\s*(?:,|and)\s*)+[A-Z][A-Z0-9]{2,7})\)', flat):
            codes += re.findall(r'[A-Z][A-Z0-9]{2,7}', cl.group(1))
        for it in re.finditer(r'([A-Z][\w’\'\- ]{2,40}?)\s*\(\s*(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4})\s*\)', m.group(2)):
            name = norm(it.group(1))
            key = re.sub(r'[^a-z]', '', name.lower())
            code = next((c for c in codes if key.startswith(c.lower()[:4]) and c not in [g['code'] for g in groups]), None) or \
                   re.sub(r'[^A-Z]', '', name.upper())[:4]
            groups.append({'code': code, 'name': name, 'points': n, 'visits': 1, 'date': parse_date(it.group(2)),
                           'campaign': None, 'evidence': norm(m.group(0))[:400]})
    for g in groups:
        g['visits'] = g['visits'] or 1
    if len(groups) >= 2: return groups
    # plain counts: "Five replicate water samples were collected at 37 sites", "A total of 450 samples were collected from 146 sites"
    NUM = r'(\d[\d,]*|' + '|'.join(NUMBER_WORDS) + r')'
    m = re.search(NUM + r'\s+(?:replicate\s+|independent\s+)?(?:\w+\s+){0,2}samples?\s+(?:\([^)]*\)\s+)?were\s+(?:collected|taken|sampled)\s+'
                  r'(?:at|from)\s+(?:each\s+of\s+)?(?:the\s+)?' + NUM + r'\s+(?:\w+\s+)?(sites|locations|ponds|lakes|stations|rivers|points)', flat, re.I)
    if m and not re.search(r'total', flat[max(0, m.start() - 15): m.start()], re.I):
        n, k = num_word(m.group(1).replace(',', '')), num_word(m.group(2).replace(',', ''))
        if n and k and n * k <= 5000:
            w = len(str(k))
            return [{'code': f'SITE{i:0{w}d}', 'name': None, 'points': n, 'visits': 1, 'date': None, 'campaign': None,
                     'evidence': at_sentence(flat, m.start()), 'estimated': True} for i in range(1, k + 1)]
    m = re.search(r'(?:A total of|In total,?)\s+' + NUM + r'\s+(?:\w+\s+){0,3}samples?\s+(?:\([^)]*\)\s+)?were\s+(?:collected|taken|sampled)', flat, re.I) or \
        re.search(NUM + r'\s+(?:\w+\s+){0,2}samples?\s+\((?:including|incl\.)[^)]*\)\s+were\s+(?:collected|taken)', flat, re.I)
    if m:
        t = num_word(m.group(1).replace(',', ''))
        if t and 2 <= t <= 5000:
            return [{'code': 'S', 'name': None, 'points': t, 'visits': 1, 'date': None, 'campaign': None,
                     'evidence': at_sentence(flat, m.start()), 'estimated': True, 'total': True}]
    return []


def at_sentence(flat, pos):
    a = flat.rfind('. ', 0, pos) + 2 if flat.rfind('. ', 0, pos) >= 0 else 0
    b = flat.find('. ', pos)
    return flat[a: b + 1 if b > 0 else pos + 300]


def despace_headings(text: str) -> str:
    """Letter-spaced headings from some PDF readers: 'R E F E R E N C E S' -> 'REFERENCES'."""
    return re.sub(r'(?m)^((?:[A-Z] ){2,}[A-Z]{1,2})(?= |$)', lambda m: m.group(1).replace(' ', ''), text)


def extract(text: str) -> dict:
    text = text.replace('\u00a0', ' ').replace('\ufeff', '').replace('ﬁ', 'fi').replace('ﬂ', 'fl').replace('ﬀ', 'ff').replace('\u2011', '-')
    text = strip_line_numbers(text)
    text = despace_headings(text)
    text = re.sub('[\ue000-\uf8ff]', '', text)                     # private-use glyphs (bullets)
    _words = set(re.findall(r'[a-z]{5,}', text.lower()))
    text = re.sub(r'([A-Za-z]+)-\n([a-z]+)', lambda m_: m_.group(1) + m_.group(2) if (m_.group(1) + m_.group(2)).lower() in _words
                  else m_.group(1) + '-' + m_.group(2), text)          # "tem-\nplate" -> "template", "high-\nfidelity" stays
    text = re.sub(r'(?<=\d)\s?(?:o|º|˚|⁰)C\b|℃', '°C', text)            # "98oC", "60℃" -> °C
    text = re.sub(r'(?<=[a-z]) - ?\n?(?=[a-z]{3,})', '', text)          # "tem -\nperature"
    refs = references(text)
    cuts = [m for m in re.finditer(r'\n[ \t]*(?:\d+\.\s*)?(?:References(?: to published material)?|REFERENCES|Literature [Cc]ited|Bibliography)[ \t]*(?:\n|\t)', text)
            if m.start() > 0.3 * len(text)]                 # skip a table of contents
    cut = cuts[0] if cuts else None
    appendix = ''
    if cut:                                               # appendices after the reference list often hold lab protocols
        ap = re.search(r'\n[ \t]*(?:Appendix|APPENDIX|Annex|ANNEX|Supplementary (?:Methods|Information))\b[^\n]{0,100}\n', text[cut.start():])
        if ap: appendix = text[cut.start() + ap.start():].replace('\u03bc', '\u00b5')
    text = text[:cut.start()] if cut else text            # body only from here on
    text = text.replace('\u03bc', '\u00b5')                # Greek mu -> micro sign
    flat = norm(text)
    sents = sentences(text)
    _spans, _p = [], 0
    for _s in sents:
        _i = flat.find(_s, _p)
        if _i >= 0: _spans.append((_i, _i + len(_s), _s)); _p = _i + len(_s)
    evidence = {}

    def at(pos):
        for a0, a1, ss in _spans:
            if a0 <= pos < a1: return ss
        return flat[max(0, pos - 150): pos + 150]

    def cite(key, src):
        if src is None or key in evidence: return
        evidence[key] = (src if isinstance(src, str) else at(src.start()))[:600]

    meta: dict = {'assays': [], 'sites': [], 'rounds': [], 'runs': [], 'notes': [], '_evidence': evidence}

    # ── project ──────────────────────────────────────────────────────────────
    tm = re.search(r'Project title[ \t]*:?[ \t]*\n?((?:[ \t]*\S.*\n?){1,3})', text, re.I)
    title_lines = []
    if tm:
        for line in tm.group(1).splitlines():
            if not norm(line) or re.match(r'\s*\d+\.', line) or re.match(r'\s*(Contractor|organisation)', line, re.I): break
            title_lines.append(norm(line))
    fm = front_matter(text)
    meta['project_title'] = ' '.join(title_lines) or fm.get('project_title')
    cite('project_title', 'Project title: ' + meta['project_title'] if title_lines else meta['project_title'])
    org = after_label(text, r'organisation\(s\)') if title_lines else None
    meta['institution'] = org or fm.get('institution')
    if meta['institution']: cite('institution', f'Contractor organisation(s): {org}' if org else fm.get('_ev_institution'))
    if not meta['institution']:
        um = re.search(r'\bUKCEH\b|Centre for Ecology (?:&|and) Hydrology', flat[:6000])
        if um: meta['institution'] = 'UK Centre for Ecology & Hydrology'; cite('institution', um)
    for acr, (name, ror) in INSTITUTIONS.items():
        if meta['institution'] and (acr in meta['institution'] or name.lower() in meta['institution'].lower() or
                                    re.search(r'Centre for Ecology (?:&|and) Hydrology', meta['institution'])):
            meta['institution'], meta['institution_ror'] = name, ror; cite('institution_ror', evidence.get('institution'))
    for k in ('recorded_by', 'project_contact'):
        if fm.get(k): meta[k] = fm[k]
    if fm.get('recorded_by'): cite('recorded_by', fm.get('_ev_authors'))
    if fm.get('project_contact'): cite('project_contact', fm.get('_ev_contact') or fm.get('_ev_authors'))
    meta['start_date'] = parse_date(after_label(text, r'start date') or '')
    if meta['start_date']: cite('start_date', 'start date: ' + (after_label(text, r'start date') or ''))
    meta['end_date'] = parse_date(after_label(text, r'end date') or '')
    if re.search(r'\bDefra\b', flat): meta['funder'] = 'Defra'

    # ── controls ─────────────────────────────────────────────────────────────
    fb = re.search(r'(\w+) field blanks? .{0,80}?(?:per|for each) (?:lake|site|pond|river|location)', flat, re.I)
    meta['field_blanks_per_site_round'] = num_word(fb.group(1)) if fb else (1 if re.search(r'field blank', flat, re.I) else 0)
    cite('field_blanks', fb or re.search(r'field blank', flat, re.I))
    eb = re.search(r'(\w+) extraction (?:blanks?|negatives?|controls?) were included', flat, re.I)
    meta['extraction_blanks'] = num_word(eb.group(1)) if eb else 0
    cite('extraction_blanks', eb)
    pb = re.search(r'(\w+) PCR (?:blanks?|negatives?) were included for each primer pair', flat, re.I)
    meta['pcr_negatives_per_assay'] = num_word(pb.group(1)) if pb else 0
    cite('pcr_negatives', pb)
    meta['neg_controls'] = bool(meta['field_blanks_per_site_round'] or meta['extraction_blanks'] or
                                meta['pcr_negatives_per_assay'] or re.search(r'negative control', flat, re.I))
    meta['pos_controls'] = bool(re.search(r'positive control|mock communit', flat, re.I))
    cite('neg_controls', fb or eb or pb or re.search(r'negative control', flat, re.I))
    cite('pos_controls', re.search(r'positive control|mock communit', flat, re.I))

    # ── sterilisation ────────────────────────────────────────────────────────
    st = [s for s in sents if re.search(r'UV[- ]steril|bleach|hypochlorite|Virkon|Clinell|decontaminat|DNA[- ]?away|DNA[- ]?Zap|'
                                        r'wiped with|sterile gloves|autoclav|steril\w+ (?:with|in|by|using)', s, re.I)
          and not re.search(r'scalpel|sliced', s, re.I)] or \
         [s for s in sents if re.search(r'sterile scalpel', s, re.I)]
    meta['sterilise_method'] = ' '.join(st[:2]) or None
    cite('sterilise_method', meta['sterilise_method'])

    # ── sampling design ──────────────────────────────────────────────────────
    sm = re.search(r'(\w+)\s+(\d+(?:\.\d+)?)\s*(L|mL|ml|litre|liter)s?\s+water samples', flat, re.I)
    if sm:
        meta['samples_per_site_round'] = num_word(sm.group(1))
        meta['samp_size'] = sm.group(2)
        meta['samp_size_unit'] = 'L' if sm.group(3).lower() in ('l', 'litre', 'liter') else 'mL'
        for _k in ('samples_per_site_round', 'samp_size', 'samp_size_unit', 'samp_collect_method'): cite(_k, sm)
    sub = re.search(r'(\d+)\s*[x×]\s*(\d+)\s*m[lL]\s+subsamples', flat)
    if sub: meta['subsamples'] = f'{sub.group(1)} x {sub.group(2)} mL subsamples'
    if sub and not sm:
        meta['samp_size'], meta['samp_size_unit'] = f'{int(sub.group(1)) * float(sub.group(2)):g}', 'mL'
        cite('samp_size', sub); cite('samp_size_unit', sub)
    where = find(r'collected from the (shoreline|shore|littoral zone|centre|center|deepest point|surface)', flat)
    parts = []
    if sm: parts.append(f'{sm.group(1).capitalize()} {sm.group(2)} {meta["samp_size_unit"]} water samples per site per sampling round')
    if sub: parts.append(f'each composed of {meta["subsamples"]}')
    if where: parts.append(f'collected from the {where}')
    meta['samp_collect_method'] = ', '.join(parts) or None

    # habitat → ENVO
    HABITATS = [  # (words, broad, local, medium)
        (r'\blakes?\b|\btarns?\b|\blochs?\b|still waters?|lentic|reservoirs?', 'freshwater lake biome [ENVO:01000252]', 'lake [ENVO:00000020]', 'fresh water [ENVO:00002011]'),
        (r'\bponds?\b', 'aquatic biome [ENVO:00002030]', 'pond water [ENVO:00002228]', 'liquid water [ENVO:00002006]'),
        (r'\brivers?\b|\bstreams?\b|riverine|lotic', None, None, 'river water [ENVO:01000599]'),
        (r'\bseawater\b|sea water|\bocean\w*|\bmarine\b|offshore|\bcoastal\b', 'marine biome [ENVO:00000447]', 'marine pelagic zone [ENVO:00000208]', 'sea water [ENVO:00002149]'),
        (r'Continuous Plankton Recorder|plankton (?:net|tow|samples)', 'marine biome [ENVO:00000447]', None, 'planktonic material [ENVO:01000063]'),
    ]
    hab_counts = [(len(re.findall(p, flat, re.I)) * (10 if 'Plankton' in p else 1), (p, b, l, m)) for p, b, l, m in HABITATS]
    hab_counts.sort(key=lambda x: -x[0])
    if hab_counts[0][0] >= 3:
        p, b, l, m = hab_counts[0][1]
        if re.search(r'sediment cores?|sedimentary|\bsedDNA\b', flat, re.I) and len(re.findall(r'sediment', flat, re.I)) > hab_counts[0][0]:
            b, l, m = (b, l, 'marine sediment [ENVO:03000033]') if 'marine' in (b or '') else (b, l, 'lake sediment [ENVO:00000546]')
        hm = re.search(p, flat, re.I)
        for _k, _v in (('env_broad_scale', b), ('env_local_scale', l), ('env_medium', m)):
            if _v: meta[_k] = _v; cite(_k, hm)
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

    if not meta['sites']:   # one cell per line, rows separated by blank lines: "Name\nCODE\nv1\n…\n\n"
        blocks = [b.split('\n') for b in re.split(r'\n[ \t]*\n', text)]
        run = []
        for b in blocks + [[]]:
            b = [x.strip() for x in b if x.strip()]
            ok = len(b) >= 3 and re.fullmatch(r'[A-Z][A-Za-z\' ]{2,40}', b[0]) and re.fullmatch(r'[A-Z]{3,5}', b[1]) and all(len(x) < 60 for x in b)
            if ok and (not run or len(b) == len(run[0])): run.append(b); continue
            if len(run) >= 3: break
            run = [b] if ok else []
        if len(run) >= 3 and len({b[1] for b in run}) == len(run):
            meta['sites'] = [{'name': b[0], 'code': b[1], 'attributes': ' '.join(b[2:])} for b in run]

    # sampling rounds: "weeks commencing 14th July and 13th October 2025"
    wc = re.search(r'weeks? commencing ((?:\d{1,2}(?:st|nd|rd|th)?\s+' + MONTH_RE + r'(?:\s+\d{4})?(?:,\s*|\s+and\s+)?)+)', flat, re.I)
    if wc:
        cite('rounds', wc)
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
        for _k in ('filter_diameter', 'size_frac', 'filter_material'): cite(_k, fl)
    meta['filter_active'] = bool(re.search(r'vacuum|peristaltic|pump', flat, re.I))
    cite('filter_active', re.search(r'vacuum|peristaltic|pump', flat, re.I))
    proc = [s for s in sents if re.search(r'filter(ed|ing)? within|filters were used per sample|Filtering was stopped', s, re.I)]
    meta['samp_mat_process'] = ' '.join(proc[:3]) or None
    cite('samp_mat_process', meta['samp_mat_process'])
    within = re.search(r'filtered within (\d+)\s*(hours?|h|days?)', flat, re.I)
    if within: meta['samp_store_dur'] = f'PT{within.group(1)}H' if within.group(2).lower().startswith('h') else f'P{within.group(1)}D'
    cite('samp_store_dur', within); cite('samp_store_method_additional', re.search(r'transported on ice', flat, re.I))
    if re.search(r'transported on ice', flat, re.I): meta['samp_store_method_additional'] = 'Transported on ice; ' + (f'filtered within {within.group(1)} {within.group(2)}' if within else '')
    sl = re.search(r'transported on ice to ([A-Z][\w ]+?) and', flat)
    if sl: meta['samp_store_loc'] = sl.group(1)
    cite('samp_store_loc', sl)
    sol = re.search(r'containing\s+(?:\d+\s*m[lL]\s+)?(DNA/RNA Shield|ethanol|Longmire\'?s?(?: buffer)?|ATL buffer|RNAlater)\s*(\([^)]+\))?', flat, re.I)
    if sol: meta['prepped_samp_store_sol'] = norm(sol.group(1) + ' ' + (sol.group(2) or ''))
    cite('prepped_samp_store_sol', sol)
    tmp = next((m for m in re.finditer(r'(?:stored|frozen) at (-?\d+)\s*°?\s*C', flat, re.I)
                if not re.search(r'PCR products?|amplicons?|libraries|DNA extracts?\b|extracted DNA|eluted', at(m.start()), re.I)), None)
    if tmp: meta['prepped_samp_store_temp'] = tmp.group(1)
    cite('prepped_samp_store_temp', tmp)

    # ── extraction ───────────────────────────────────────────────────────────
    kit = re.search(r'performed using ([A-Z][^()]{3,80}?Kit)\s*\(([^)]+)\)', flat)
    if kit: meta['ext_kit'] = f'{kit.group(1)} ({kit.group(2)})'
    cite('ext_kit', kit); cite('ext_sep', kit)
    sop = re.search(r'follow those outlined in ([A-Z][\w\-]+ et al\.? \(\d{4}\))', flat)
    if sop:
        meta['nucl_acid_ext'] = ref_for(sop.group(1).replace('(', '').replace(')', ''), refs)
        cite('nucl_acid_ext', sop); cite('nucl_acid_amp', sop)
        if re.search(r'(?:metabarcoding|PCR|amplification) conditions follow', flat, re.I):
            meta['nucl_acid_amp'] = meta['nucl_acid_ext']
    lys = [s for s in sents if re.search(r'mechanical lysis|bead[- ]beat|homogeni[sz]er|Proteinase K', s, re.I)]
    meta['ext_lysis'] = ' '.join(lys[:2]) or None
    cite('ext_lysis', meta['ext_lysis'])
    kitname = meta.get('ext_kit') or ''
    if re.search(r'DNeasy|Power(?:Water|Soil|Max)|Quick-DNA|NucleoSpin|QIAamp|E\.Z\.N\.A|Monarch', kitname, re.I):
        meta['ext_sep'] = 'column-based'
    elif re.search(r'MagMAX|KingFisher|MagAttract|magnetic', kitname + flat, re.I):
        meta['ext_sep'] = 'magnetic beads'
    mod = re.search(r'with the following amendments.*?(?=Thereafter|DNA extraction was verified)', flat, re.I)
    if mod: meta['ext_modifications'] = norm(re.sub(r'^with the following amendments[^.]*\.\s*', '', mod.group(0)))[:600]
    cite('ext_modifications', meta.get('ext_modifications'))
    conc = [s for s in sents if re.search(r'Qubit|Nanodrop|NanoDrop|fluorometer', s) and 'librar' not in s.lower()]
    meta['conc_method'] = conc[0] if conc else None
    cite('conc_method', meta['conc_method'])
    nm = re.search(r'normalised to (?:a standard template concentration of )?(\d+(?:\.\d+)?\s*ng/[µu]L)', flat, re.I)
    if nm: meta['notes'].append(f'Sample DNA normalised to {nm.group(1)} before PCR')

    # ── assays ───────────────────────────────────────────────────────────────
    REJECT = r'(?:instead of|rather than|in preference to|as opposed to|alternatives? such as|compared with)\s+'
    accepted = norm(re.sub(REJECT + r'(?:the\s+)?[\w/\-]+', ' ', flat, flags=re.I))
    counts = {}
    for raw in re.findall(r'(?<![A-Za-z0-9])' + MARKER_RE + r'(?![A-Za-z0-9])', accepted):   # case-sensitive: 'its' is not ITS
        counts[canon_marker(raw)] = counts.get(canon_marker(raw), 0) + 1
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
        _tm = re.search(r'(?<![A-Za-z0-9])' + mk + r'\)?\s+(?:' + '|'.join(p for p, _ in TAXON_WORDS) + r')\b', flat, re.I)
        if _tm: cite(f'{mk}.targetTaxonomicAssay', _tm); cite(f'{mk}.targetTaxonomicScope', _tm)
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
            for _k in ('primer_name_fwd', 'primer_name_rev', 'assay_name'): cite(f'{mk}.{_k}', s)
            ref = re.search(r'primers?(?:\s+[\w\-]+\s*(?:and|/)\s*[\w\-]+)?\s+of ([A-Z][\w\-]+ et al\.?,? \(?\d{4}\)?)', main) or \
                  re.search(r'([\w\-]+)\s*/\s*([\w\-]+)\s+primers? of ([A-Z][\w\-]+ et al\.?,? \(?\d{4}\)?)', main)
            if ref:
                full = ref_for(ref.group(ref.lastindex).replace('(', '').replace(')', '').replace(',', ''), refs)
                a['primer_ref_fwd'] = a['primer_ref_rev'] = full
                cite(f'{mk}.primer_ref_fwd', s); cite(f'{mk}.primer_ref_rev', s)
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
        cite(f'{mk}.custom_mm', ss[0]); cite(f'{mk}.commercial_mm', ' '.join(ss)[:600] if m else None)
        dna = re.search(r'(\d+(?:\.\d+)?)\s*[µu]L\s*(?:of\s+)?template DNA', joined)
        if dna: a['pcr_dna_vol'] = dna.group(1)
        if dna: cite(f'{mk}.pcr_dna_vol', next((x for x in ss if dna.group(0) in x), None))
        recipe = next((x for x in ss if re.search(r'contained|combined with|consisted of|comprised', x, re.I)), ss[0])
        vols = [float(v) for v in re.findall(r'(\d+(?:\.\d+)?)\s*[µu]L', recipe)]
        each = re.search(r'(\d+(?:\.\d+)?)\s*[µu]L\s*(?:of\s*)?(?:(?:\d+\s*[µu]?M|\d+\s*mM)\s*)?(?:of )?each', recipe)
        total = sum(vols) + (float(each.group(1)) if each else 0)   # "each" primer counted twice
        if dna and dna.group(0) not in recipe: total += float(dna.group(1))
        if vols: a['amplificationReactionVolume'] = f'{total:g}'
        if vols: cite(f'{mk}.amplificationReactionVolume', ' '.join(ss)[:600])
        pv = re.search(r'(\d+(?:\.\d+)?)\s*[µu]L\s*(?:of\s*)?(?:each\s*)?(\d+)\s*[µu]M\s*(?:of\s*)?(?:each\s*)?forward', recipe) or \
             re.search(r'(\d+(?:\.\d+)?)\s*[µu]L\s*of each\s*(\d+)\s*[µu]M', recipe)
        if pv: a['pcr_primer_vol'], a['pcr_primer_conc'] = pv.group(1), pv.group(2)
        if pv: cite(f'{mk}.pcr_primer_vol', recipe); cite(f'{mk}.pcr_primer_conc', recipe)
        elif re.search(r'\d+\s*mM\s*of each forward', recipe):
            meta['notes'].append(f'{mk}: primer stock concentration is reported in mM (likely µM) – please check')
    # PCR replicates
    _rp = re.search(r'single reactions|single PCR|one PCR replicate|without (?:PCR )?replication', flat, re.I)
    if _rp: rep = '1'
    else:
        rm = re.search(r'(\w+) PCR replicates (?:were|per)', flat, re.I) or re.search(r'in (duplicate|triplicate)', flat, re.I)
        rep = {'duplicate': '2', 'triplicate': '3'}.get(rm.group(1).lower(), str(num_word(rm.group(1)) or '')) if rm else None
    an = re.search(r'anneal\w*\s+(?:at\s+)?(\d{2}(?:\.\d)?)\s*°\s*C', flat, re.I)
    cy = re.search(r'(\d{2})\s+cycles\b(?! for)', flat, re.I)
    for a in assays.values():
        if rep: a['pcr_rep'] = rep; cite(f"{a['marker']}.pcr_rep", _rp or (rm if not _rp else None))
    if re.search(r'thermocycling (?:parameters|conditions) are (?:described|given|shown) in Table', flat, re.I):
        meta['notes'].append('Thermocycling conditions and primer sequences are in a report table that is not part of the text – please add them manually')

    # ── library prep / sequencing ────────────────────────────────────────────
    if re.search(r'two[- ]step', flat, re.I): meta['barcoding_pcr_appr'] = 'two-step PCR'; cite('barcoding_pcr_appr', re.search(r'two[- ]step', flat, re.I))
    elif re.search(r'one[- ]step|single[- ]step', flat, re.I): meta['barcoding_pcr_appr'] = 'one-step PCR'
    elif re.search(r'ligation', flat, re.I): meta['barcoding_pcr_appr'] = 'ligation-based'
    ad = re.search(r'forward adaptor:\s*5\W{0,2}-\s?([ACGTN]+)\s?-3\W{0,2},\s*reverse adaptor:\s*5\W{0,2}-\s?([ACGTN]+)\s?-3', flat)
    if ad: meta['adapter_forward'], meta['adapter_reverse'] = ad.group(1), ad.group(2)
    cite('adapter_forward', ad); cite('adapter_reverse', ad)
    ad2 = list(re.finditer(r'forward adaptor:\s*5\W{0,2}-\s?([ACGTN]+)\s?-3\W{0,2},\s*reverse adaptor:\s*5\W{0,2}-\s?([ACGTN]+)\s?-3', flat))
    if len(ad2) > 1: meta['notes'].append(f'Indexing (step II) adaptors: forward {ad2[1].group(1)}, reverse {ad2[1].group(2)}'); cite('seq_method_additional', ad2[1])
    p2 = re.search(r'Per sample (\d+(?:\.\d+)?)\s*[µu]L of step 1 product.*?to a total of (\d+(?:\.\d+)?)\s*[µu]L', flat, re.I)
    if p2:
        meta['pcr2_dna_vol'], meta['pcr2_amplificationReactionVolume'] = p2.group(1), p2.group(2)
        m = re.search(mm_re, p2.group(0), re.I)
        if m: meta['pcr2_commercial_mm'] = norm(m.group(1))
        meta['pcr2_custom_mm'] = norm(p2.group(0))
        for _k in ('pcr2_dna_vol', 'pcr2_amplificationReactionVolume', 'pcr2_commercial_mm', 'pcr2_custom_mm'): cite(_k, p2)
    lq = re.search(r'libraries were quantified using (?:a |an )?([^.,;]+?)(?: and | to |,|\.)', flat, re.I)
    if lq: meta['lib_conc_meth'] = norm(lq.group(1))
    cite('lib_conc_meth', lq)
    lib = [s for s in sents if re.search(r'librar(y|ies)', s, re.I) and re.search(r'Qubit|quantif|normali[sz]|gel purif|Bioanalyzer|TapeStation', s)]
    meta['lib_screen'] = ' '.join(lib[:3]) or None
    cite('lib_screen', meta['lib_screen'])
    ph = re.search(r'(\d+(?:\.\d+)?)\s*%\s*(?:addition of\s*)?(?:[\w ]{0,30})PhiX', flat, re.I)
    if ph: meta['phix_perc'] = ph.group(1)
    cite('phix_perc', ph)
    if re.search(r'\bIllumina\b', flat): meta['platform'] = 'ILLUMINA'; cite('platform', re.search(r'\bIllumina\b', flat))
    elif re.search(r'Nanopore|MinION|GridION|PromethION', flat): meta['platform'] = 'OXFORD_NANOPORE'
    elif re.search(r'PacBio', flat): meta['platform'] = 'PACBIO_SMRT'
    ins = re.search(r'Illumina (NextSeq|MiSeq|NovaSeq|HiSeq|iSeq)\b', flat) or re.search(r'\b(NextSeq|MiSeq|NovaSeq|HiSeq|iSeq|MinION|GridION|PromethION)(?:™)?\s*(\d+(?:/\d+)?)?', flat)
    if ins:
        model = re.search(r'(NextSeq|NovaSeq|HiSeq|iSeq)(?:™)?\s*(500|550|1000/2000|1000|2000|2500|3000|4000|6000|100)\b', flat)
        meta['instrument'] = ('Illumina ' if meta.get('platform') == 'ILLUMINA' else '') + (f'{model.group(1)} {model.group(2)}' if model else ins.group(1))
        cite('instrument', model or ins)
    _pe = re.search(r'paired[- ]end|mergePairs|forward and reverse reads', flat, re.I)
    if _pe: meta['lib_layout'] = 'paired end'; cite('lib_layout', _pe)
    loc = re.search(r'to ([A-Z][\w ]+?) \([^)]*\) for DNA extraction and sequencing|to ([A-Z][\w ]+?) for (?:DNA extraction and )?sequencing', flat)
    if loc: meta['sequencing_location'] = norm(loc.group(1) or loc.group(2))
    cite('sequencing_location', loc)
    # kit per assay: "<kit> 600 Cycles for pooled 16S,18S, rbcL and COI, and 200 Cycles for 12S"
    kb = re.search(r'((?:NextSeq|MiSeq|NovaSeq)(?:(?!NextSeq|MiSeq|NovaSeq)[^.])*?)\s*(\d+)\s*Cycles for', flat)
    if kb:
        base = norm(re.sub(r'Chemistry', '', kb.group(1))).replace('™', '')
        for cyc, mlist in re.findall(r'(\d+)\s*Cycles for (?:pooled\s*)?((?:' + MARKER_RE + r'[\s,]*(?:and\s+)?)+)', flat):
            for mk in markers_in(mlist):
                if mk in assays: assays[mk]['seq_kit'] = f'{base} {cyc} cycles'; cite(f'{mk}.seq_kit', kb)
    # runs: dates in order, run descriptions in order
    rd = re.search(r'(?:Sequenc\w* runs? (?:were|was) completed[^.]*? on|runs? on)\s*([\d/]+(?:\s*(?:and|,)\s*[\d/]+)*)', flat, re.I)
    rdesc = re.findall(r'(\d+)\s*cycle run \(((?:' + MARKER_RE + r'[\s,]*(?:and\s+)?)+)\)', flat, re.I)
    if rd:
        dates = [parse_date(d) for d in re.findall(r'\d{1,2}/\d{1,2}/\d{2,4}', rd.group(1))]
        for i, d in enumerate(dates):
            cyc, mlist = rdesc[i] if i < len(rdesc) else ('', '')
            mks = markers_in(mlist) if mlist else list(assays)
            meta['runs'].append({'seq_run_id': f'run_{d.replace("-", "")}' + (f'_{cyc}cycle' if cyc else ''), 'date': d, 'assays': mks})
        cite('runs', rd)

    # ── bioinformatics ───────────────────────────────────────────────────────
    cut = re.search(r'cutadapt\s*\((v?[\d.]+)', flat, re.I) or re.search(r'cutadapt\s*(v?[\d.]+)?', flat, re.I)
    if cut: meta['trim_method'] = 'cutadapt' + (f' {cut.group(1)}' if cut.group(1) else '')
    cite('trim_method', cut)
    if re.search(r'DADA2', flat):
        rv = find(r'\bin R\s*v?\.?\s*([\d.]+\d)', flat)
        meta['sop_bioinformatics'] = 'DADA2 pipeline (Callahan et al. 2016)' + (f' in R v{rv}' if rv else '')
        cite('sop_bioinformatics', re.search(r'DADA2', flat)); cite('otu_clust_tool', re.search(r'DADA2', flat))
        cite('error_rate_tool', re.search(r'maxEE', flat)); cite('merge_tool', re.search(r'mergePairs', flat)); cite('otu_clust_cutoff', re.search(r'DADA2', flat))
        meta['error_rate_tool'] = 'DADA2 filterAndTrim' if re.search(r'maxEE', flat) else None
        if re.search(r'mergePairs', flat): meta['merge_tool'] = 'DADA2 mergePairs'
        meta['otu_clust_tool'] = 'DADA2 (amplicon sequence variants, no clustering)'
        meta['otu_clust_cutoff'] = '100'
    ee = re.search(r'maxEE\)?\s*=\s*(\d+(?:\.\d+)?)(?:\s*,\s*(\d+(?:\.\d+)?))?', flat)
    if ee:
        meta['error_rate_type'] = 'expected error rate'
        meta['error_rate_cutoff'] = ee.group(1)
        cite('error_rate_type', ee); cite('error_rate_cutoff', ee)
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
        if bits: cite(f'{mk}.trim_param', ' '.join(x for x in [at(m2.start()) for m2 in re.finditer(r'trimmed to|maxEE', flat)][:2]))
    if re.search(r'removeBimeraDenovo', flat):
        meta['chimera_check_method'] = 'DADA2 removeBimeraDenovo'
        cite('chimera_check_method', re.search(r'removeBimeraDenovo', flat)); cite('chimera_check_param', re.search(r'removeBimeraDenovo', flat))
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
            if mk in assays: assays[mk].setdefault('dbs', []).append(norm(m.group(2))); cite(f'{mk}.otu_db', s)
        if fm:
            for d in re.findall(db_re, s):
                for mk in scope:
                    if mk in assays: assays[mk].setdefault('dbs', []).append(norm(d)); cite(f'{mk}.otu_db', s)
            if re.search(r'SINTAX', s):
                for mk in scope: cite(f'{mk}.otu_seq_comp_appr', s)
                for mk in scope:
                    if mk in assays: assays[mk]['otu_seq_comp_appr'] = 'SINTAX (Edgar 2016)'
        elif scope and not re.match(r'(?:ASVs|Primer|Sequences|Reference)', s):
            for d in re.findall(db_re, s):
                for mk in scope:
                    if mk in assays: assays[mk].setdefault('dbs', []).append(norm(d)); cite(f'{mk}.otu_db', s)
            if re.search(r'BLAST', s):
                for mk in scope:
                    if mk in assays: assays[mk]['blast_note'] = norm(s)
        else:
            scope = []
    if len(assays) == 1:
        mk1 = list(assays)[0]
        for s in bsents:
            if re.search(r'database|reference', s, re.I) and not assays[mk1].get('dbs'):
                found = re.findall(db_re, s) + re.findall(r'\b(Midori2?|MIDORI2?)\b', s)
                if len(found) > 1:
                    assays[mk1]['dbs'] = list(dict.fromkeys(found)); cite(f'{mk1}.otu_db', s)
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
        cite(f"{mk}.otu_seq_comp_appr", re.search(r'assignTaxonomy', flat)); cite(f"{mk}.tax_assign_cat", evidence.get(f"{mk}.otu_seq_comp_appr"))
        if a.get('otu_seq_comp_appr'):
            a['tax_assign_cat'] = 'sequence composition'
        if a.pop('blast_note', None):
            a['otu_seq_comp_appr'] = ((a.get('otu_seq_comp_appr') + '; ') if a.get('otu_seq_comp_appr') else '') + 'BLAST against NCBI for confirmation of selected taxa'
    # rarefaction depths
    rare = re.search(r'rarefaction to (.*?) using', flat, re.I)
    if rare:
        for n, mk in re.findall(r'([\d,]+)\s*(?:sequences\s*)?for\s+(' + MARKER_RE + r')', rare.group(1)):
            mk = canon_marker(mk)
            if mk in assays:
                assays[mk]['otu_final_description'] = f'ASV table rarefied to {n} reads per sample' + (' (rarefy_even_depth, phyloseq)' if 'phyloseq' in flat else '')
                cite(f'{mk}.otu_final_description', rare)
    raw_bits = []
    if meta.get('trim_method'): raw_bits.append(f'primers removed with {meta["trim_method"]}')
    if re.search(r'quality filter', flat, re.I): raw_bits.append('quality filtering')
    if re.search(r'denois', flat, re.I): raw_bits.append('denoising')
    if meta.get('merge_tool'): raw_bits.append('merging of paired reads')
    if meta.get('chimera_check_method'): raw_bits.append('chimera removal')
    if raw_bits: meta['otu_raw_description'] = 'ASV table after ' + ', '.join(raw_bits) + (' (DADA2)' if 'DADA2' in flat else '')
    cite('otu_raw_description', re.search(r'DADA2', flat))
    cont = [s for s in sents if re.search(r'blanks?', s, re.I) and re.search(r'contaminat|not reliable|mis-?binning', s, re.I)]
    if cont: meta['screen_contam_method'] = ' '.join(cont[:2])[:700]
    cite('screen_contam_method', meta.get('screen_contam_method'))
    stats = [p for p in ('phyloseq', 'vegan', 'microeco', 'lme4', 'emmeans', 'ggplot2') if re.search(r'\b' + p, flat)]
    if stats: meta['bioinfo_method_additional'] = 'Downstream analysis in R using ' + ', '.join(stats)
    cite('bioinfo_method_additional', re.search(r'\b' + stats[0], flat) if stats else None)

    # ── general patterns for other report / paper styles ─────────────────────
    # Runs after the specific rules above and only fills what is still empty
    # (except PCR cycling, which is assigned per assay here).
    mtext = methods_text(text)
    fl = norm(mtext or text).replace('\u2212', '-').replace('μ', 'µ')
    fsents = [x.replace('\u2212', '-').replace('μ', 'µ') for x in (sentences(mtext) if mtext else sents)]
    one_assay = list(assays)[0] if len(assays) == 1 else None

    meta['_methods'] = fl if mtext else None
    conflicts = meta.setdefault('_conflicts', {})

    def differs(a, b):
        tok = lambda v: set(re.findall(r'[a-z0-9]+(?:\.[0-9]+)*', re.sub(r'\b(?:v|version)\.?\s*(?=\d)', '', str(v).lower()))) - \
            {'the', 'a', 'of', 'and', 'kit', 'chemistry', 'reagent', 'database'}
        x, y = tok(a), tok(b)
        return bool(x and y) and not (x <= y or y <= x)

    def put(key, value, src):
        if value in (None, ''): return
        if not meta.get(key): meta[key] = value; cite(key, src)
        elif isinstance(meta[key], str) and differs(meta[key], value):
            conflicts.setdefault(key, []).append(str(value))

    def put_a(mk, key, value, src, override=False):
        a = assays.get(mk)
        if a is None or value in (None, ''): return
        if override or not a.get(key):
            a[key] = value; evidence.pop(f'{mk}.{key}', None); cite(f'{mk}.{key}', src)
        elif isinstance(a[key], str) and differs(a[key], value):
            conflicts.setdefault(f'{mk}.{key}', []).append(str(value))

    def segments(sentence):
        """'16S … 25 cycles … 18S … 35 cycles' -> [('16S', '16S … 25 cycles …'), ('18S', '18S … 35 cycles')]."""
        hits = [(m.start(), canon_marker(m.group(0))) for m in re.finditer(r'(?<![A-Za-z0-9])' + MARKER_RE + r'(?![A-Za-z0-9])', sentence)]
        hits = [(p, mk) for p, mk in hits if mk in assays]
        if len({mk for _, mk in hits}) < 2: return [(None, sentence)]
        out = []
        for k, (p, mk) in enumerate(hits):
            end = hits[k + 1][0] if k + 1 < len(hits) else len(sentence)
            if out and out[-1][0] == mk: out[-1] = (mk, out[-1][1] + sentence[p:end])
            else: out.append((mk, sentence[p:end]))
        if hits[0][0] > 0: out[0] = (out[0][0], sentence[:hits[0][0]] + out[0][1])
        return out

    def nearest_marker(src, pos):
        found = [canon_marker(m.group(0)) for m in re.finditer(r'(?<![A-Za-z0-9])' + MARKER_RE + r'(?![A-Za-z0-9])', src[max(0, pos - 250):pos])]
        found = [mk for mk in found if mk in assays]
        return found[-1] if found else one_assay

    def targets(sentence, carry):
        """Assays a methods sentence is about: markers it names, else the carried scope, else all assays."""
        mks = [m for m in markers_in(sentence) if m in assays]
        return mks or carry or list(assays)

    # ── targeted assays: primer/probe sets read from sentences and tables ────────
    body_plus = text + ('\n' + appendix if appendix else '')
    t_groups = oligo_groups([r for r in oligo_records(body_plus) if r['targeted']])
    prow = pipe_rows(body_plus)
    gene_cols = next((v for k_, v in prow.items() if re.search(r'target(?:ed)? gene|\bgene\b|\bmarker\b|\blocus\b', k_)), None)
    name_cols = next((v for k_, v in prow.items() if re.search(r'assay(?: name)?$|assay id', k_)), None)
    methods_flat = norm(mtext or text)
    targeted_study = len(TARGETED_WORDS.findall(methods_flat)) >= 3
    if targeted_study and not t_groups and not SEQ_WORDS.search(methods_flat):
        gs = genes_in(' '.join(x for x in fsents if TARGETED_WORDS.search(x)))
        t_groups = [{'gene': gs[0] if gs else None, 'name': None, 'targeted': True, 'src': next((x for x in fsents if TARGETED_WORDS.search(x)), None)}]

    def new_key(base):
        base = re.sub(r'[^\w.\-]+', '_', base).strip('_') or 'assay'
        k_, n_ = base, 2
        while k_ in assays: k_, n_ = f'{base}_{n_}', n_ + 1
        return k_

    t_keys = []
    for gi, g in enumerate(t_groups):
        gene = g.get('gene')
        if gene_cols and len(gene_cols) == len(t_groups): gene = canon_gene(gene_cols[gi]) or gene
        nm_ = g.get('name') or (name_cols[gi] if name_cols and len(name_cols) == len(t_groups) else None)
        k_ = new_key(f'{gene}_{nm_ or "targeted"}' if gene else (nm_ or 'targeted'))
        a = {'marker': k_, 'gene': gene, 'assay_type': 'targeted', 'assay_name': nm_ or k_,
             'target_gene': TARGET_GENE.get(gene) if gene in TARGET_GENE else (f'other: {gene}' if gene else None)}
        assays[k_] = a; t_keys.append(k_)
        if g.get('src'): cite(f'{k_}.assay_type', g['src']); cite(f'{k_}.target_gene', g['src'] if gene else None)
        for role in ('fwd', 'rev'):
            r = g.get(role)
            if not r: continue
            a[f'primer_{role}'] = r['seq']; cite(f'{k_}.primer_{role}', r['src'])
            if r.get('name') and r['name'] != nm_: a[f'primer_name_{role}'] = r['name']; cite(f'{k_}.primer_name_{role}', r['src'])
        pr = g.get('probe')
        if pr:
            a['probe_seq'] = pr['seq']; cite(f'{k_}.probe_seq', pr['src'])
            if pr.get('reporter'): a['probeReporter'] = pr['reporter']; cite(f'{k_}.probeReporter', pr['src'])
            if pr.get('quencher'): a['probeQuencher'] = pr['quencher']; cite(f'{k_}.probeQuencher', pr['src'])
        a['_names'] = sorted({x for x in [nm_] + [g[r]['name'] for r in ('fwd', 'rev', 'probe') if g.get(r) and g[r].get('name')] if x}, key=len)
    if t_keys:
        lod_rows = [(k_, v) for k_, v in prow.items() if re.match(r'(?:lod|limit of detection|loq|limit of quantification)\b', k_) and len(v) == len(t_keys)]
        unit_hint = re.search(r'cop\w*\s*(?:#\s*)?/\s*(rxn|reaction|well|[µu]L)|cop\w* per (reaction|[µu]L)', body_plus, re.I)
        for lab, vals in lod_rows:
            kind = 'lod' if re.match(r'lod|limit of detection', lab) else 'loq'
            for k_, v in zip(t_keys, vals):
                num = re.match(r'\s*(\d+(?:\.\d+)?)', v)
                if num:
                    assays[k_][f'pcr_assay_{kind}'] = num.group(1); cite(f'{k_}.pcr_assay_{kind}', f'{lab}: ' + ' | '.join(vals))
                    if unit_hint: assays[k_][f'pcr_assay_{kind}_unit'] = 'copies/' + ('reaction' if re.match(r'rxn|reaction|well', unit_hint.group(1) or unit_hint.group(2) or '', re.I) else 'µL')
        dig = len(re.findall(r'\bdd?PCR\b|digital PCR|droplet digital', methods_flat, re.I)); rt = len(re.findall(r'\bqPCR\b|real[- ]time|quantitative PCR', methods_flat, re.I))
        for k_ in t_keys: assays[k_]['amp_vis_method'] = 'dPCR' if dig > rt else 'qPCR'
    mb_keys = [k_ for k_ in assays if k_ not in t_keys]
    one_assay = mb_keys[0] if len(mb_keys) == 1 else None

    def t_named(s_):
        """Targeted assays a passage names (assay or oligo names, else its gene)."""
        hit = [k_ for k_ in t_keys if any(re.search(r'(?<![\w.])' + re.escape(n_) + r'(?![\w])', s_) for n_ in assays[k_].get('_names', []))]
        if not hit and TARGETED_WORDS.search(s_):         # gene named in a qPCR/dPCR passage
            gs = genes_in(s_)
            hit = [k_ for k_ in t_keys if assays[k_].get('gene') and assays[k_]['gene'] in gs]
        return hit

    def pcr_details(k_, s_, specific, done):
        """Reaction details for one targeted assay from a sentence or a protocol block (recipe lists included)."""
        s_ = norm(s_.replace('−', '-').replace('μ', 'µ'))
        if not re.search(r'µ[lL]|cycles?|°\s*C|[Mm]aster ?[Mm]ix|[Ss]upermix|[Pp]olymerase|PCR|replicate|probe', s_, re.I): return
        ov = specific and k_ not in done
        cyc = re.search(r'(\d{2})\s*(?:PCR\s*|amplification\s*)?cycles\b', s_) or re.search(r'\b(?:number of )?cycles\W{1,3}(\d{2})\b', s_, re.I)
        if cyc and 15 <= int(cyc.group(1)) <= 60: put_a(k_, 'pcr_cycles', cyc.group(1), s_, override=ov)
        ann = re.search(r'anneal\w*(?:[^.;°]{0,40}?)(\d{2}(?:\.\d)?)\s*°\s*C', s_, re.I) or \
              re.search(r'(\d{2}(?:\.\d)?)\s*°\s*C(?:(?!°)[^.;,]){0,25}\banneal', s_, re.I) or \
              re.search(r'cycles? of\s+(?:\d+\s*(?:s|sec\w*|min\w*)\s+(?:at\s+)?9\d\s*°\s*C|9\d\s*°\s*C\s+for\s+\d+\s*(?:s|sec\w*|min\w*))\s*(?:,|and)\s*'
                        r'(?:\d+\s*(?:s|sec\w*|min\w*)\s+(?:at\s+)?(\d{2}(?:\.\d)?)\s*°\s*C|(\d{2}(?:\.\d)?)\s*°\s*C)', s_, re.I)
        if ann:
            v = next(x for x in ann.groups() if x)
            if 45 <= float(v) <= 72: put_a(k_, 'annealingTemp', v, s_, override=ov)
        tot = re.search(r'(?:total|final|reaction)\s+volumes?\s+(?:of\s+)?(\d+(?:\.\d+)?)\s*µ[lL]', s_, re.I) or \
              re.search(r'(\d+(?:\.\d+)?)[ -]?µ[lL]\s+(?:q?PCR\s+|dPCR\s+|ddPCR\s+)?reactions?\b', s_, re.I)
        if tot and float(tot.group(1)) <= 100: put_a(k_, 'amplificationReactionVolume', tot.group(1), s_)
        two = re.search(r'(\d+(?:\.\d+)?)\s*µ[lL]\s+(?:of\s+)?(?:\w+\s+)?(?:template|DNA|eDNA|extract)\b[^.;]{0,40}?\band\s+(\d+(?:\.\d+)?)\s*µ[lL]\s+(?:of\s+)?(?:\w+\s+)?(?:master ?mix|mastermix|reaction mix|PCR mix|mix)\b', s_, re.I) or \
              re.search(r'(\d+(?:\.\d+)?)\s*µ[lL]\s+(?:of\s+)?(?:\w+\s+)?(?:master ?mix|mastermix|reaction mix|PCR mix)\b[^.;]{0,40}?\band\s+(\d+(?:\.\d+)?)\s*µ[lL]\s+(?:of\s+)?(?:\w+\s+)?(?:template|DNA|eDNA|extract)\b', s_, re.I)
        if two and not tot: put_a(k_, 'amplificationReactionVolume', f'{float(two.group(1)) + float(two.group(2)):g}', s_)
        items = re.findall(r'(?:^|[\s;,:])([a-z])[.)]\s*(\d+(?:\.\d+)?)\s*µ[lL]\b([^;]{0,90})', s_)
        letters = [ord(x[0]) for x in items]
        items = [(v, rest) for _, v, rest in items] if letters and letters == list(range(letters[0], letters[0] + len(letters))) else []
        if len(items) >= 3 and not tot:            # recipe list: a. 3 µL template, b. 1 µL of each primer/probe (…), c. 12.5 µL mix, d. water
            total = 0.0
            for v, rest in items:
                each = re.search(r'\beach\b', rest)
                n = 1
                if each:
                    inside = re.search(r'\(([^)]*)', rest)
                    n = len(re.findall(r';', inside.group(1))) + 1 if inside and ';' in inside.group(1) else (3 if re.search(r'primers?/probe|primer and probe', rest) else 2)
                total += float(v) * n
            if total <= 100: put_a(k_, 'amplificationReactionVolume', f'{total:g}', s_)
        dna = re.search(r'(\d+(?:\.\d+)?)\s*µ[lL]\s+(?:of\s+)?(?:the\s+)?(?:\w+\s+){0,2}?(?:template|DNA|eDNA|extract)\b', s_, re.I)
        if dna and 'elut' not in s_.lower()[max(0, dna.start() - 40): dna.end()] and float(dna.group(1)) <= 20: put_a(k_, 'pcr_dna_vol', dna.group(1), s_)
        mm = re.search(MM, s_)
        if mm: put_a(k_, 'commercial_mm', norm(mm.group(1)), s_)
        rp = re.search(r'\b(?:in|run in|performed in|carried out in|run as|tested in)\s+(duplicate|triplicate|quadruplicate)', s_, re.I) or \
             re.search(r'(?<![/\d.])\b(two|three|four|five|six|eight|ten|twelve|\d+)\s+(?:independent\s+|technical\s+|replicate\s+|qPCR\s+|PCR\s+|dPCR\s+|ddPCR\s+)*(?:replicates?|reactions? per sample)\b', s_, re.I)
        if rp:
            r_ = rp.group(1).lower()
            n = {'duplicate': 2, 'triplicate': 3, 'quadruplicate': 4}.get(r_) or num_word(r_)
            if n and n < 50: put_a(k_, 'pcr_rep', str(n), s_)
        for v, unit, nm_ in re.findall(r'(\d+(?:\.\d+)?)\s*(µmol/L|µM|nM)\s+(?:of\s+)?([A-Za-z][\w.\-]*)', s_):
            names = assays[k_].get('_names', [])
            conc = f'{float(v) / 1000:g}' if unit == 'nM' else v
            if re.search(r'probe', nm_, re.I) or (assays[k_].get('probe_seq') and nm_.lower().endswith(('p', 'pr'))):
                if not names or any(n_ in nm_ for n_ in names): put_a(k_, 'probe_conc', conc, s_)
            elif not names or any(n_ in nm_ for n_ in names):
                put_a(k_, 'pcr_primer_conc', conc, s_)
        cfg = re.search(r'(?:conditions?|program(?:me)?|profile|cycling)[^:]{0,40}:\s*((?:[^.]*?\d+\s*°C){3,}[^.]*)', s_, re.I) or \
              re.search(r'((?:initial|activation)[^.]{0,80}?\d+\s*°\s*C[^.]{0,200}?\d{2}\s*cycles[^.]{0,200})', s_, re.I)
        if cfg: put_a(k_, 'pcr_cond', norm(cfg.group(1))[:300], s_)
        th = re.search(QPCR_INSTRUMENTS, s_)
        if th: put('thermocycler', th.group(1), s_)
        lod = re.search(r'\b(?:LOD|limit of detection)\b[^.;]{0,40}?(?:was|of|=|:)\s*(\d+(?:\.\d+)?)\s*(copies|copy)\s*(?:/|per)\s*(µ[lL]|reaction|rxn|well|L)\b', s_, re.I)
        if lod: put_a(k_, 'pcr_assay_lod', lod.group(1), s_); put_a(k_, 'pcr_assay_lod_unit', f'copies/{lod.group(3)}', s_)
        loq = re.search(r'\b(?:LOQ|limit of quantification)\b[^.;]{0,40}?(?:was|of|=|:)\s*(\d+(?:\.\d+)?)\s*(copies|copy)\s*(?:/|per)\s*(µ[lL]|reaction|rxn|well|L)\b', s_, re.I)
        if loq: put_a(k_, 'pcr_assay_loq', loq.group(1), s_); put_a(k_, 'pcr_assay_loq_unit', f'copies/{loq.group(3)}', s_)
        if re.search(r'gBlock|synthetic (?:DNA|oligo|gene|standard)|plasmid standard|standard curve', s_, re.I):
            st = re.search(r'gBlock|synthetic \w+|plasmid', s_, re.I)
            if st: put_a(k_, 'std_type', {'gblock': 'gBlock'}.get(st.group(0).lower(), st.group(0).lower()), s_)
        done.add(k_)

    # PCR, per assay
    MM = (r'((?:Q5|Phusion|Platinum|KAPA|AmpliTaq|GoTaq|MyTaq|HotStar|Phire|TaqMan|Qiagen Multiplex|QIAGEN Multiplex|Multiplex PCR'
          r'|Environmental Master|Invitrogen|Takara|Accuprime|AccuPrime|Bioline|NEBNext|OneTaq|DreamTaq|Maxima|PrimeSTAR|Ex Taq'
          r'|QIAcuity|ddPCR|TaqPath|PrimeTime|Luna|SsoAdvanced|SsoFast|PowerUp|PowerTrack|SensiFAST|PerfeCTa|GoTaq Probe|Fast Advanced)'
          r'[\w®™\-.,/ ]{0,45}?(?:Master ?[Mm]ix|ReadyMix|Mastermix|[Mm]ix|Supermix|PCR Kit|[Kk]it|[Pp]olymerase|Gold|Taq)(?:\s*\d\.\d)?)')
    carry = None
    pcr_segments = []
    for sent in fsents:
        if not re.search(r'µ[lL]|cycles?|°\s*C|[Mm]aster ?[Mm]ix|[Pp]olymerase|PCR', sent): carry = None; continue
        for seg_mk, seg in segments(sent):
            pcr_segments.append((seg_mk, seg, sent))
    for s in fsents:
        amp = re.search(r'(\d{2,4})\s*(?:-\s*)?(?:bp|base[- ]pairs?)\s+(?:long\s+)?(?:region|fragment|amplicon|product|section)', s, re.I)
        named = [m for m in markers_in(s) if m in assays] or ([one_assay] if one_assay else [])
        if amp and named and re.search(r'amplif|primer|target', s, re.I):
            for mk in named: put_a(mk, 'ampliconSize', amp.group(1), s)
    t_done, t_carry, last_sent, t_weak = set(), (None, 0), None, {}
    for seg_mk, s, sent in pcr_segments:
        named = [seg_mk] if seg_mk else [m for m in markers_in(s) if m in mb_keys]
        tn = t_named(s) if t_keys else []
        if sent is not last_sent: t_carry = (t_carry[0], t_carry[1] - 1) if t_carry[1] > 0 else (None, 0)
        last_sent = sent
        mb_turn = named or SEQ_WORDS.search(s) or re.search(r'first[- ](?:round|step) PCR|PCR\s*1\b|librar', s, re.I)
        if t_keys and (tn or (TARGETED_WORDS.search(s) and not SEQ_WORDS.search(s) and not named) or (t_carry[0] and not mb_turn)):
            explicit = bool(tn) or bool(TARGETED_WORDS.search(s))
            ks = tn or (t_carry[0] if t_carry[0] and not TARGETED_WORDS.search(s) else t_keys)
            for k_ in ks:
                before_ = dict(assays[k_])
                if explicit:                                  # an explicit qPCR/dPCR sentence beats values carried over from a neighbour
                    weak = t_weak.get(k_, set())
                    for f_ in weak: assays[k_].pop(f_, None)
                pcr_details(k_, s, bool(tn), t_done)
                if explicit:
                    for f_ in list(t_weak.get(k_, set())):
                        if f_ not in assays[k_]: assays[k_][f_] = before_[f_]
                        else: t_weak[k_].discard(f_)
                else:
                    t_weak.setdefault(k_, set()).update(f_ for f_ in assays[k_] if f_ not in before_)
            t_carry = (ks, 2); carry = None; continue
        if mb_turn: t_carry = (None, 0)
        if re.search(r'\bqPCR\b|real[- ]time|\bprobe\b|\bC[tq]\b|ddPCR|\bdPCR\b|single[- ]species|TaqMan', s) and not named: carry = None; continue
        if re.search(r'second[- ](?:step|round|stage)|step[- ](?:2|II)\b|index(?:ing)?\s+PCR|\bPCR\s*2\b|PCR2\b|second PCR|library PCR', s, re.I):
            continue                                   # indexing PCR: not the assay's own amplification
        mks = named or carry or list(mb_keys)
        carry = named or carry
        cyc = re.search(r'(\d{2})\s*cycles?\s*(?:of|consisting of|comprising|each of|:)?\s*[^.;]{0,30}?(\d{2,3}(?:\.\d)?)\s*°\s*C[^,;]{0,30}[,;]\s*(?:followed by\s*)?(?:and\s*)?(\d{2}(?:\.\d)?)\s*°\s*C', s)
        ann = re.search(r'anneal\w*(?:[^.;°]{0,40}?)(\d{2}(?:\.\d)?)\s*°\s*C', s, re.I) or \
              re.search(r'(\d{2}(?:\.\d)?)\s*°\s*C(?:(?!°)[^.;,]){0,25}\banneal', s, re.I)   # "annealing at 60 °C" before "60 °C (annealing)"
        cyc_only = re.search(r'(\d{2})\s*(?:PCR\s*)?cycles\b', s)
        pcr_ctx = re.search(r'PCR|reaction|amplif|[Mm]aster ?[Mm]ix|polymerase|primer|template', s)
        disp = re.search(r'(\d+(?:\.\d+)?)\s*µ[lL]\s+(?:was|were|of (?:the )?(?:master ?mix|mix|solution) was)\s+(?:dispensed|aliquoted|added)[^.]{0,90}?'
                         r'(?:addition of|followed by|with)\s+(\d+(?:\.\d+)?)\s*µ[lL]\s+(?:of\s+)?(?:\w+\s+)?(?:template|DNA)', s, re.I)
        vol = pcr_ctx and (re.search(r'(?:total|final|reaction)\s+volumes?\s+(?:of\s+)?(\d+(?:\.\d+)?)\s*µ[lL]', s, re.I)
               or re.search(r'(\d+(?:\.\d+)?)[ -]?µ[lL]\s+(?:PCR\s+)?reactions?', s, re.I)
               or re.search(r'\((\d+(?:\.\d+)?)\s*µ[lL]\s*(?:per|/)\s*(?:sample|reaction)\)', s, re.I)
               or re.search(r'reactions?\s+(?:of|with a volume of)\s+(\d+(?:\.\d+)?)\s*µ[lL]', s, re.I))
        dna = re.search(r'(\d+(?:\.\d+)?)\s*µ[lL]\s+(?:of\s+)?(?:the\s+)?(?:\w+\s+)?(?:template|DNA|eDNA|extract)\b', s, re.I)
        mm = re.search(MM, s)
        rep = re.search(r'\b(?:in|run in|performed in|carried out in)\s+(duplicate|triplicate|quadruplicate)', s, re.I) or \
              re.search(r'\b(two|three|four|five|six|eight|ten|twelve|\d+)\s+(?:independent\s+|technical\s+|replicate\s+)*(?:PCR\s+)?replicates?\b', s, re.I)
        for mk in mks:
            specific = bool(named)
            if cyc:
                put_a(mk, 'pcr_cycles', cyc.group(1), s, override=specific); put_a(mk, 'annealingTemp', cyc.group(3), s, override=specific)
            elif cyc_only: put_a(mk, 'pcr_cycles', cyc_only.group(1), s, override=specific)
            if ann: put_a(mk, 'annealingTemp', ann.group(1), s, override=specific or bool(cyc))
            if vol and float(vol.group(1)) <= 200: put_a(mk, 'amplificationReactionVolume', vol.group(1), s)
            if disp: put_a(mk, 'amplificationReactionVolume', f'{float(disp.group(1)) + float(disp.group(2)):g}', s)
            pv = re.search(r'(\d+(?:\.\d+)?)\s*µ[lL]\s+of\s+(?:each\s+)?(?:the\s+)?(\d+(?:\.\d+)?)\s*[µu]M\s+(?:of\s+)?(?:each\s+)?(?:forward and reverse\s+)?primers?', s, re.I)
            if pv: put_a(mk, 'pcr_primer_vol', pv.group(1), s); put_a(mk, 'pcr_primer_conc', pv.group(2), s)
            amp = re.search(r'(\d{2,4})\s*(?:-\s*)?(?:bp|base[- ]pairs?)\s+(?:long\s+)?(?:region|fragment|amplicon|product|section)', s, re.I)
            if amp: put_a(mk, 'ampliconSize', amp.group(1), s)
            cond = re.search(r'(?:conditions?|program(?:me)?|profile|cycling)[^:]{0,40}:\s*((?:[^.]*?\d+\s*°C){3,}[^.]*)', s, re.I)
            if cond: put_a(mk, 'pcr_cond', norm(cond.group(1)), s)
            if dna and pcr_ctx and 'elut' not in s.lower() and float(dna.group(1)) <= 20: put_a(mk, 'pcr_dna_vol', dna.group(1), s)
            if mm: put_a(mk, 'commercial_mm', norm(mm.group(1)), s)
            if rep:
                r = rep.group(1).lower()
                n = {'duplicate': 2, 'triplicate': 3, 'quadruplicate': 4}.get(r) or num_word(r)
                if n and n < 20: put_a(mk, 'pcr_rep', str(n), s)

    if t_keys:
        all_names = {n_ for k_ in t_keys for n_ in assays[k_].get('_names', [])}
        proto = (mtext or text) + ('\n' + appendix if appendix else '')       # methods and appendices only
        proto = re.sub(r'(?m)^.*\bPage \d+ of \d+\b.*\n?', '', proto)              # running page headers
        heads = [m_.start() + 1 for m_ in re.finditer(r'\n[ \t]*(?:\d+(?:\.\d+)*\.?\s+)?[A-Z][^\n.:;,]{2,60}[ \t]*\n', proto)]
        cuts_ = [0] + heads + [len(proto)]
        for b0, b1 in zip(cuts_, cuts_[1:]):
            blk = proto[b0:b1]
            if not TARGETED_WORDS.search(blk) and not t_named(blk): continue
            if SEQ_WORDS.search(blk): continue                 # mixed methods text: handled sentence by sentence above
            hit = t_named(blk)
            foreign = re.findall(r'(?:µmol/L|µM|nM)\s+(?:of\s+)?([A-Za-z][\w.\-]*\d*[\w.\-]*)', blk) + re.findall(r'\b(\w+\.probe)\b', blk)
            foreign = [f_ for f_ in foreign if not any(n_ in f_ or f_ in n_ for n_ in all_names) and not re.match(r'(?:each|the|of|primers?|probes?|forward|reverse)$', f_, re.I)]
            if not hit and foreign: continue
            for k_ in hit or (t_keys if len(t_keys) == 1 or re.search(r'\beach (?:assay|target)|all (?:three|four|two|\d) assays|the (?:three|four|two) assays', blk, re.I) else []):
                pcr_details(k_, blk, False, t_done)

    # primers: DNA sequences in the text (tables included), adapters/probes removed
    NEXTERA_F, NEXTERA_R = 'TCGTCGGCAGCGTCAGATGTGTATAAGAGACAG', 'GTCTCGTGGGCTCGGAGATGTGTATAAGAGACAG'
    IUPAC = 'ACGTRYKMSWBDHVNI'
    wrapped = text.replace('\u2212', '-')
    for _ in range(4):   # only sequences wrapped over a line break: "…AGACAGTA\nGAACAGGC…"
        wrapped = re.sub(r'\b([' + IUPAC + r']{8,})[ \t]*\n[ \t]*([' + IUPAC + r']{2,})\b(?![ \t]+[' + IUPAC + r']{12,})(?![.\-]?[a-z])', r'\1\2', wrapped)
    joined = norm(wrapped).replace('μ', 'µ')
    seen_seq, order, cur_pair = set(), {}, None
    for m in re.finditer(r'(?<![A-Za-z])([' + IUPAC + r']{15,})(?![A-Za-z])', joined):
        seq = m.group(1)
        if sum(c in 'ACGT' for c in seq) < 0.6 * len(seq) or seq in seen_seq: continue
        seen_seq.add(seq)
        before = joined[max(0, m.start() - 120): m.start()]
        ctx = list(re.finditer(r'(probe|\bFAM\b|BHQ|\bMGB\b|quencher|TaqMan|qPCR|ddPCR|single[- ]species)|(metabarcod|amplicon sequencing|Illumina|first[- ]round PCR|librar)',
                               joined[max(0, m.start() - 400): m.start()], re.I))
        if ctx and ctx[-1].group(1): continue
        if re.search(r'\bFAM\b|BHQ|\bMGB\b|quencher', joined[m.end(): m.end() + 30]): continue
        if 'AATGATACGGCGACCACC' in seq or 'CAAGCAGAAGACGGCATACGA' in seq: continue          # index adapters
        overhang = seq.startswith(NEXTERA_F) or seq.startswith(NEXTERA_R)    # adapter + primer: the primer is kept below
        if NEXTERA_F.startswith(seq) or NEXTERA_R.startswith(seq) or \
           (not overhang and re.search(r'adaptor|adapter|index|overhang|tag(?:ged)? sequence|barcode', joined[max(0, m.start() - 45): m.start()], re.I)): continue
        direction = None
        for ov, d in ((NEXTERA_F, 'fwd'), (NEXTERA_R, 'rev')):
            if seq.startswith(ov): seq, direction = seq[len(ov):], d
        if len(seq) < 14 or seq in (NEXTERA_F, NEXTERA_R): continue
        near = before[-70:]
        if direction is None:
            dw = list(re.finditer(r'(forward|fwd|\b[\w\-]*\d[\w\-]*F\b|\bF\s*[:(])|(reverse|\brev\b|\b[\w\-]*\d[\w\-]*R\b|\bR\s*[:(])', near, re.I))
            if dw: direction = 'fwd' if dw[-1].group(1) else 'rev'
        nm = re.findall(r'\b([A-Za-z0-9][\w\-]*?\d[\w\-]*[FRfr](?:wd|ev)?)\b', near)
        lab = re.search(r'\(\s*([A-Za-z][\w.\-]{1,20})\s*[;:,]\s*5\W{0,3}-?\s*$', near)
        if lab: nm = [lab.group(1)]
        mk = nearest_marker(joined, m.start())
        if not mk: continue
        if cur_pair and m.start() - cur_pair[1] < 60 and not assays[cur_pair[0]].get('primer_rev'): mk = cur_pair[0]   # second primer of the pair
        pair = re.search(r'\b((?=[\w\-]*\d)(?=[\w\-]*[A-Za-z])[\w\-]{3,})\s*\(\s*$', near)            # "Tele01 (SEQ, SEQ, …)"
        base = assays[mk]
        if pair and base.get('primer_fwd') and base.get('primer_rev') and pair.group(1) != base.get('_pair_label'):
            k2 = new_key(f"{base.get('_base', mk)}_{pair.group(1)}")       # a further primer pair for the same marker -> its own assay
            assays[k2] = {kk: vv for kk, vv in base.items() if kk in ('target_gene', 'targetTaxonomicAssay', 'targetTaxonomicScope', 'target_subfragment')}
            assays[k2].update(marker=k2, _base=base.get('_base', mk), assay_name=f"{base.get('_base', mk)}_{pair.group(1)}")
            mb_keys.append(k2); mk = k2
        if pair and not assays[mk].get('_pair_label'): assays[mk]['_pair_label'] = pair.group(1)
        cur_pair = (mk, m.end())
        k = order.setdefault(mk, 0)
        if direction is None: direction = 'fwd' if k % 2 == 0 else 'rev'
        order[mk] = k + 1
        src = (before[-150:] + seq)[-300:]
        put_a(mk, 'primer_fwd' if direction == 'fwd' else 'primer_rev', seq, src)
        if nm: put_a(mk, 'primer_name_fwd' if direction == 'fwd' else 'primer_name_rev', nm[-1], src)

    # sequencing kit / instrument / indexing approach
    kit = re.search(r'((?:MiSeq|NextSeq|NovaSeq|HiSeq|iSeq) (?:Reagent|Nano|Micro|High Output|Mid Output|SBS)[^.;)]{0,40}?v\d(?:\.\d)?[^.;)]{0,25}?\d{3}[ -]?cycles?)', fl, re.I) or \
          re.search(r'((?:MiSeq|NextSeq|NovaSeq|HiSeq|iSeq)[\w ]{0,15}?(?:Reagent\s+(?:Kit\s+)?|Nano\s+)?v\d(?:\.\d)?[^.;)]{0,25}?\d{3}[ -]?cycles?)', fl, re.I) or \
          re.search(r'((?:MiSeq|NextSeq|NovaSeq|HiSeq|iSeq)[\w ]{0,20}?(?:Reagent Kit|Nano Kit|Micro Kit|flow ?cell)[^.;)]{0,20})', fl, re.I)
    if kit:
        for mk in mb_keys: put_a(mk, 'seq_kit', norm(kit.group(1)), kit)
    if not meta.get('instrument'):
        im = re.search(r'\b(MiSeq|iSeq 100|NextSeq (?:500|550|1000|2000|1000/2000)|NovaSeq (?:6000|X(?: Plus)?)|HiSeq (?:2000|2500|4000|X)|MinION|GridION|PromethION|Ion Torrent PGM|Ion S5)\b', fl)
        if im: put('instrument', ('Illumina ' if re.match(r'MiSeq|iSeq|NextSeq|NovaSeq|HiSeq', im.group(1)) else '') + im.group(1), im)
    if not meta.get('platform'):
        if re.search(r'\bIllumina\b|MiSeq|NovaSeq|NextSeq|HiSeq', fl): put('platform', 'ILLUMINA', re.search(r'\bIllumina\b|MiSeq|NovaSeq|NextSeq|HiSeq', fl))
        elif re.search(r'Nanopore|MinION|GridION', fl): put('platform', 'OXFORD_NANOPORE', re.search(r'Nanopore|MinION|GridION', fl))
        elif re.search(r'Ion Torrent|Ion S5|PGM', fl): put('platform', 'ION_TORRENT', re.search(r'Ion Torrent|Ion S5|PGM', fl))
    two = re.search(r'second(?: round of| round| step)? PCR|indexing PCR|index PCR|two rounds of PCR|PCR2\b|second-step PCR|two-step|Fluidigm|Nextera XT|overhang adapter', fl, re.I)
    if two: put('barcoding_pcr_appr', 'two-step PCR', two)
    elif re.search(r'fusion primers|one-step|single-step|tagged primers', fl, re.I):
        put('barcoding_pcr_appr', 'one-step PCR', re.search(r'fusion primers|one-step|single-step|tagged primers', fl, re.I))
    if re.search(r'paired[- ]end|\d\s*[x×]\s*\d{2,3}\s*bp|forward and reverse reads|R1 and R2', fl, re.I):
        put('lib_layout', 'paired end', re.search(r'paired[- ]end|\d\s*[x×]\s*\d{2,3}\s*bp|forward and reverse reads|R1 and R2', fl, re.I))
    phx = re.search(r'(\d+(?:\.\d+)?)\s*%\s*(?:of\s+)?PhiX|PhiX[^.;]{0,30}?(\d+(?:\.\d+)?)\s*%', fl, re.I)
    if phx: put('phix_perc', phx.group(1) or phx.group(2), phx)

    # bioinformatics tools (with versions where given)
    def tools(names, context):
        found = []
        for s in fsents:
            if not re.search(context, s, re.I): continue
            for t in names:
                for m in re.finditer(r'\b(' + t + r')\b\s*(?:\(\s*)?(?:v(?:ersion)?\.?\s*)?(\d+(?:\.\d+)+)?', s, re.I):
                    label = m.group(1) + (f' {m.group(2)}' if m.group(2) else '')
                    if not any(label.split()[0].lower() == f.split()[0].lower() for f, _ in found): found.append((label, s))
        return found
    tr = tools(['cutadapt', 'Trimmomatic', 'fastp', 'Trim Galore', 'AdapterRemoval', 'BBDuk', 'Sickle', 'PRINSEQ', 'ngsfilter', 'Skewer'],
               r'trim|primer|adapter|quality|clip|remov')
    if len(tr) > 1: meta['trim_method'] = '; '.join(t for t, _ in tr); cite('trim_method', tr[0][1])
    elif tr: put('trim_method', tr[0][0], tr[0][1])
    mg = tools(['FLASH2?', 'PEAR', 'VSEARCH', 'USEARCH', 'PANDAseq', 'fastp', 'BBMerge', 'illuminapairedend', 'DADA2'],
               r'merg|join|pair|overlap|assembl')
    if mg: put('merge_tool', mg[0][0] + (' mergePairs' if mg[0][0].upper().startswith('DADA2') else ''), mg[0][1])
    cl = tools(['VSEARCH', 'USEARCH', 'UPARSE', 'UNOISE3?', 'SWARM', 'CD-HIT', 'Deblur', 'DADA2', 'QIIME ?2?', 'OBITools'],
               r'cluster|OTU|ASV|denois|zOTU|amplicon sequence variant')
    if cl: put('otu_clust_tool', cl[0][0], cl[0][1])
    cc = re.search(r'(\d{2,3}(?:\.\d+)?)\s*%\s*(?:sequence\s*)?(?:similarity|identity)[^.;]{0,40}(?:OTU|cluster)|(?:OTU|cluster)[^.;]{0,60}?(\d{2,3}(?:\.\d+)?)\s*%', fl, re.I)
    if cc and not meta.get('otu_clust_cutoff'): put('otu_clust_cutoff', cc.group(1) or cc.group(2), cc)
    tax = None
    for pat, cat in ((r'\bBLAST\w*\b|blastn|megablast|usearch_global|global alignment', 'sequence similarity'),
                     (r'na[iï]ve Bayes|RDP classifier|feature-classifier|IDTAXA|assignTaxonomy|SINTAX|k-mer', 'sequence composition'),
                     (r'phylogenetic placement|\bEPA\b|pplacer|SEPP', 'phylogeny')):
        m = re.search(pat, fl, re.I)
        if m:
            tax = (cat, m); break
    idc = re.search(r'(\d{2,3}(?:\.\d+)?)\s*%\s*(?:sequence\s*|query\s*)?(?:identity|similarity|match)', fl, re.I)
    qcv = re.search(r'(\d{2,3})\s*%\s*(?:query\s*)?coverage', fl, re.I)
    mrd = re.search(r'(?:fewer|less) than (\d+) (?:reads|copies|sequences)[^.]{0,80}?(?:removed|excluded|discarded)|(?:removed|excluded|discarded)[^.]{0,60}?(?:fewer|less) than (\d+) (?:reads|copies|sequences)', fl, re.I)
    mrd = mrd or re.search(r'(?:with|had|having)\s*(?:<|fewer than|less than)\s*(\d[\d,]*)\s*(?:total\s+)?reads()', fl, re.I)
    if mrd: put('min_reads_cutoff', (mrd.group(1) or mrd.group(2)).replace(',', ''), mrd); put('min_reads_cutoff_unit', 'reads', mrd)
    if idc: put('tax_class_id_cutoff', idc.group(1), idc)
    if qcv: put('tax_class_query_cutoff', qcv.group(1), qcv)
    DB = (r'(SILVA|Silva|PR2|MIDORI2?|Midori2?|UNITE|Greengenes2?|RDP|BOLD|GenBank|NCBI(?: nt| nr| GenBank)?|MitoFish|diat\.barcode|CRABS|EMBL)'
          r'(?:[^.;,()]{0,25}?\(?\s*(?:v|version|release|r)\.?\s*(\d+(?:\.\d+)*))?')
    last_named, since = [], 99
    for s in sorted(fsents, key=lambda x: 'respectively' not in x):
        named = [mm for mm in markers_in(s) if mm in assays]
        if named: last_named, since = named, 0
        else: since += 1
        if not re.search(r'taxonom|classif|BLAST|reference database|database|assign', s, re.I): continue
        dbs = [(m.group(1) + (f' v{m.group(2)}' if m.group(2) else '')) for m in re.finditer(DB, s)]
        dbs = list(dict.fromkeys(dbs))
        if not dbs: continue
        if 'respectively' in s and len(named) == len(dbs):
            for mk, d in zip(named, dbs): put_a(mk, 'otu_db', d, s)
            continue
        scope = named or (last_named if since <= 4 else []) or ([one_assay] if one_assay else [])
        for mk in scope: put_a(mk, 'otu_db', ' | '.join(dbs), s)
    for mk in mb_keys:
        if tax: put_a(mk, 'tax_assign_cat', tax[0], tax[1]); put_a(mk, 'otu_seq_comp_appr', norm(tax[1].group(0)), tax[1])

    # sampling, filtration, storage, extraction
    fsz = (re.search(r'(\d+(?:\.\d+)?)\s*[-‐]?\s*µm(?:\s*pore(?:\s*size)?)?[^.;]{0,50}?(?:filter|Sterivex|membrane|capsule)', fl, re.I)
           or re.search(r'(?:filter|Sterivex|membrane|capsule)[^.;]{0,50}?(\d+(?:\.\d+)?)\s*µm', fl, re.I)
           or re.search(r'pore size of (\d+(?:\.\d+)?)\s*µm', fl, re.I))
    if fsz: put('size_frac', fsz.group(1), fsz)
    fdm = re.search(r'(\d{2,3})\s*[-‐]?\s*mm(?:\s*diameter)?[^.;]{0,40}?(?:filter|membrane)|(?:filter|membrane)[^.;]{0,30}?(\d{2,3})\s*mm', fl, re.I)
    if fdm: put('filter_diameter', fdm.group(1) or fdm.group(2), fdm)
    for pat, mat in ((r'Sterivex[- ]?GP|polyethersulfone|\bPES\b', 'polyethersulfone'), (r'Sterivex[- ]?HV|PVDF|Durapore', 'other: PVDF'),
                     (r'cellulose nitrate|nitrocellulose', 'other: cellulose nitrate'), (r'mixed cellulose ester|\bMCE\b', 'cellulose ester'),
                     (r'glass[- ]fib(?:re|er)|GF/F|GF/C', 'glass fiber'), (r'polycarbonate', 'track etched polycarbonate'), (r'\bnylon\b', 'nylon'),
                     (r'Sterivex', 'polyethersulfone')):
        m = re.search(pat, fl, re.I)
        if m and re.search(r'filter|Sterivex|membrane', at(m.start()), re.I):
            put('filter_material', mat, m); break
    fnm = re.search(r'\b(Sterivex[\w\-]*(?: [\w\-]+ Filter Unit)?|Whatman[^.,;()]{0,25}|Millipore[^.,;()]{0,25}filter|NatureMetrics[^.,;()]{0,20}|Smith-Root[^.,;()]{0,20})', fl)
    if fnm: put('filter_name', norm(fnm.group(1)), fnm)
    act = re.search(r'peristaltic|vacuum|pump|syringe', fl, re.I)
    if act and 'filter_active' not in evidence: meta['filter_active'] = True; cite('filter_active', act)
    ssz = (re.search(r'(?:filter(?:ed|ing)|collected|sampled|took|taken)\s+(?:~|approximately\s+|up to\s+|a total of\s+)?(\d+(?:\.\d+)?)\s*(L|l|litres?|liters?|mL|ml)\b', fl)
           or re.search(r'(\d+(?:\.\d+)?)\s*(L|l|litres?|liters?|mL|ml)\s+(?:of\s+)?(?:\w+\s+){0,2}(?:water|seawater|sample)s?\b', fl))
    if ssz and not meta.get('samp_size'):
        unit = 'mL' if ssz.group(2).lower() in ('ml',) else 'L'
        put('samp_size', ssz.group(1), ssz); put('samp_size_unit', unit, ssz)
    for s in fsents:
        st = re.search(r'(?:stored|kept|frozen|preserved|transported)\s+(?:at|in)\s+(-?\d+)\s*°\s*C', s, re.I)
        if st and re.search(r'PCR products?|amplicons?|libraries|DNA extracts?\b|extracted DNA|eluted', s, re.I): continue
        if st:
            put('prepped_samp_store_temp' if re.search(r'filter|extract|tube|cartridge', s, re.I) else 'samp_store_temp', st.group(1), s); break
    sol = re.search(r"(?:in|with|containing|added)\s+(?:\d+\s*m[lL]\s+(?:of\s+)?)?(?:\d+\s*%\s*)?(ethanol|Longmire'?s?(?: (?:buffer|solution))?|ATL buffer|RNAlater|DNA/RNA Shield|lysis buffer|CTAB)", fl, re.I)
    if sol: put('prepped_samp_store_sol', sol.group(1), sol)
    for s in fsents:
        if not re.search(r'extract', s, re.I): continue
        km = re.search(r"((?:[A-Z][\w®™'\-/&.]*\s){0,6}?(?:[Bb]lood (?:and|&) [Tt]issue |PowerSoil |PowerWater |Soil |Plant )?(?:Mini |Maxi |Pro |Plus )?[Kk]it)(\s*\([^)]{2,60}\))?", s)
        if km and len(km.group(1)) > 8 and not re.search(r'Assay Kit|Qubit|Quant|PicoGreen|library|clean-?up|purification|Index', km.group(0), re.I):
            put('ext_kit', norm(km.group(1) + (km.group(2) or '')), s); break
    if not meta.get('ext_sep'):
        kitname = meta.get('ext_kit') or ''
        if re.search(r'DNeasy|Power(?:Water|Soil|Max)|Quick-DNA|NucleoSpin|QIAamp|Blood (?:and|&) Tissue|Monarch|spin column', kitname + ' ' + fl, re.I):
            meta['ext_sep'] = 'column-based'; cite('ext_sep', meta.get('ext_kit') or re.search(r'spin column|DNeasy|NucleoSpin', fl, re.I))
        elif re.search(r'MagBead|MagMAX|KingFisher|MagAttract|magnetic bead', kitname + ' ' + fl, re.I):
            meta['ext_sep'] = 'magnetic beads'; cite('ext_sep', re.search(r'MagBead|MagMAX|KingFisher|MagAttract|magnetic bead', fl, re.I))
        elif re.search(r'phenol[- ]chloroform', fl, re.I):
            meta['ext_sep'] = 'phenol chloroform'; cite('ext_sep', re.search(r'phenol[- ]chloroform', fl, re.I))

    # controls
    nc = re.search(r'\bblanks?\b|negative controls?|no[- ]template controls?|\bNTCs?\b', fl, re.I)
    if nc and not meta.get('neg_controls'): meta['neg_controls'] = True; cite('neg_controls', nc)
    pc = re.search(r'positive controls?|mock communit|Microbial Community Standard', fl, re.I)
    if pc and not meta.get('pos_controls'): meta['pos_controls'] = True; cite('pos_controls', pc)

    seq_words = re.search(r'\bIllumina\b|MiSeq|NovaSeq|NextSeq|HiSeq|Nanopore|MinION|Ion Torrent|metabarcod|high[- ]throughput sequencing|amplicon sequencing', fl, re.I)
    q = re.search(r'\bqPCR\b|ddPCR|droplet digital|TaqMan|real[- ]time PCR|quantitative PCR', fl, re.I)
    if q and not seq_words and not t_keys:
        meta['assay_type'] = 'targeted'; cite('assay_type', q)
        meta['notes'].append('This looks like a targeted (qPCR/ddPCR) study. The tool is built for metabarcoding; '
                             'targeted-assay fields (probe, standard curve, LOD/LOQ) need to be filled by hand.')
    # further primer pairs on the same marker share the base assay's protocol unless stated otherwise
    for k_, a in assays.items():
        b = assays.get(a.get('_base')) if a.get('_base') and a.get('_base') != k_ else None
        if b:
            for kk in ('commercial_mm', 'amplificationReactionVolume', 'pcr_dna_vol', 'pcr_rep', 'annealingTemp', 'pcr_cycles', 'pcr_cond',
                       'seq_kit', 'otu_db', 'tax_assign_cat', 'nucl_acid_amp'):
                if b.get(kk) and not a.get(kk):
                    a[kk] = b[kk]
                    if f"{b['marker']}.{kk}" in evidence: evidence.setdefault(f'{k_}.{kk}', evidence[f"{b['marker']}.{kk}"])
    types = {a.get('assay_type') or 'metabarcoding' for a in assays.values()}
    if assays and types == {'targeted'}:
        meta['assay_type'] = 'targeted'; cite('assay_type', evidence.get(f'{t_keys[0]}.assay_type') if t_keys else None)
    elif len(types) > 1:
        meta['assay_type'] = None                        # mixed: written per assay
        for a in assays.values(): a.setdefault('assay_type', 'metabarcoding')
    if t_keys:
        meta['notes'] = [n_ for n_ in meta['notes'] if not n_.startswith('This looks like a targeted')]
        meta['notes'].append('Targeted (qPCR/dPCR) assay(s) found: ' + ', '.join(assays[k_]['assay_name'] for k_ in t_keys) +
                             '. Check the probe, LOD/LOQ and standard-curve fields; values not stated in the text are marked missing.')
    for a in assays.values():
        for kk in ('_names', '_pair_label', '_base'): a.pop(kk, None)
    meta['assays'] = list(assays.values())
    # biological replicates
    if meta.get('samples_per_site_round'): meta['biological_rep'] = str(meta['samples_per_site_round'])
    cite('biological_rep', evidence.get('samples_per_site_round'))
    # bibliographic citation for an unpublished report
    if re.search(r'Final Report', flat, re.I) and meta.get('project_title'):
        yr = (meta.get('end_date') or '')[:4]
        who = meta.get('institution') or ''
        meta['bibliographicCitation'] = f'{who} ({yr}). {meta["project_title"]}. Evidence project final report' + (f' to {meta["funder"]}' if meta.get('funder') else '') + '.'
    # ── more general project-level facts ──
    doi_head = re.search(r'(?:doi\.org/|DOI:?\s*)(10\.\d{4,5}/[^\s,;]+)', text[:3000])
    if doi_head and not meta.get('bibliographicCitation'):
        meta['bibliographicCitation'] = 'https://doi.org/' + doi_head.group(1).rstrip('.'); cite('bibliographicCitation', doi_head)
    links = find_links(text)
    if links:
        meta['associated_resource'] = ' | '.join(('https://doi.org/' + l) if l.startswith('10.') else
                                                 (f'https://www.ncbi.nlm.nih.gov/bioproject/{l}' if l.startswith('PRJ') else l)
                                                 for l in links if 'github.com' not in l)
        cite('associated_resource', next((s_ for s_ in sents if re.search(r'Zenodo|protocols\.io|BioProject|Dryad|figshare', s_)), None))
        gh = [l for l in links if 'github.com' in l and any(re.search(r'\bcode\b|\bscripts?\b|data availab|\brepositor|reproduc|\bworkflow|\bpipeline', s_, re.I) and not re.search(r'database is available|database.{0,20}download', s_, re.I)
                                                             and l.split('/')[-1][:12].lower() in re.sub(r'\s', '', s_).lower() for s_ in sents)]
        if gh and not meta.get('code_repo'): meta['code_repo'] = ' | '.join(gh); cite('code_repo', next((s_ for s_ in sents if 'github' in s_.lower()), None))
    for s_ in sorted(fsents, key=lambda x: not re.search(r'primer', x, re.I)):
        tv = re.search(r'\b(vertebrate|fish|amphibian|mammal|bird|eukaryot|prokaryot|bacteri|fung|metazoa|invertebrate|diatom|plant)(?:s|es|a|ic|al|ia|i|e)?\b', s_, re.I)
        if tv and re.search(r'primer|target|amplif|assay|coverage', s_, re.I):
            tax = {'vertebrate': 'Vertebrata', 'fish': 'Actinopterygii', 'amphibian': 'Amphibia', 'mammal': 'Mammalia', 'bird': 'Aves',
                   'eukaryot': 'Eukaryota', 'prokaryot': 'Bacteria | Archaea', 'bacteri': 'Bacteria', 'fung': 'Fungi', 'metazoa': 'Metazoa',
                   'invertebrate': 'Metazoa', 'diatom': 'Bacillariophyta', 'plant': 'Viridiplantae'}[tv.group(1).lower()]
            for mk, a in assays.items():
                if not a.get('targetTaxonomicAssay') and (len(assays) == 1 or mk in s_):
                    a['targetTaxonomicAssay'] = a['targetTaxonomicScope'] = tax
                    cite(f'{mk}.targetTaxonomicAssay', s_); cite(f'{mk}.targetTaxonomicScope', s_)
            break
    for mk, a in assays.items():
        sub = re.search(r'\b' + re.escape(mk) + r'[\s-]*(V\d{1,2}(?:-V?\d)?)\b', fl) or (re.search(r'\b(V\d)\s+region', fl) if len(assays) == 1 else None)
        if sub and not a.get('target_subfragment'): a['target_subfragment'] = sub.group(1); cite(f'{mk}.target_subfragment', sub)
    rp = re.search(r'\b(Duplicate|Triplicate|Quadruplicate)\s+PCRs?\b', fl, re.I)
    if rp:
        for mk in mb_keys: put_a(mk, 'pcr_rep', {'duplicate': '2', 'triplicate': '3', 'quadruplicate': '4'}[rp.group(1).lower()], rp)
    nt = next((s_ for s_ in fsents if re.search(r'remov|exclud|discard', s_, re.I) and re.search(r'human|domestic|non-?target|livestock|contamina\w+ species', s_, re.I)), None)
    if nt and not meta.get('screen_nontarget_method'): meta['screen_nontarget_method'] = nt[:400]; cite('screen_nontarget_method', nt)
    cr = re.search(r'©\s*(?:\d{4}\s+)?(The Author\(s\)|The Authors|[A-Z][^.]{2,80}?)(?:\.|\s+(?:Published|This))', text[:6000])
    if cr: meta['rights_holder'] = norm(cr.group(1)); cite('rights_holder', cr.group(0))
    if not any(a.get('otu_db') for a in assays.values()):
        db = re.search(r'((?:(?-i:[A-Z])[\w-]*\s+){0,4}(?:vertebrate|fish|eukaryote|metazoan|invertebrate|plant)?\s*(?:12S|16S|18S|COI|ITS|rbcL|CytB)\s+(?:reference\s+)?(?:database|library))(\s*\([^)]{0,80}\))?', fl, re.I)
        if db:
            for mk in mb_keys: put_a(mk, 'otu_db', norm(db.group(0)), db)
    sk = re.search(r'(MiSeq|NovaSeq|NextSeq|iSeq|HiSeq)[^.]{0,60}?(\d)\s*[x×]\s*(\d{3})\s*(?:bp)?\s*\(?\s*(v\d)\s*(?:chemistry|reagent)', fl, re.I) or \
         re.search(r'(MiSeq|NovaSeq|NextSeq|iSeq|HiSeq)[^.]{0,60}?\b(v\d)\b\s*(\d)\s*[x×]?\s*(\d{3})\s*bp', fl, re.I)
    if sk:
        g = sk.groups()
        ver, n, ln = (g[3], g[1], g[2]) if g[1].isdigit() else (g[1], g[2], g[3])
        for mk in mb_keys: put_a(mk, 'seq_kit', f'{g[0]} Reagent Kit {ver.lower()} ({n} × {ln} bp)', sk)
    plan = sample_plan(mtext or text, norm(mtext) if mtext else flat) or (sample_plan(text, flat) if mtext else [])
    meta['_primer_text'] = fl
    if plan:
        meta['sample_plan'] = plan
        cite('sample_plan', plan[0]['evidence'])
        n = sum(g['points'] * g['visits'] for g in plan)
        if plan[0].get('estimated'):
            parts = (f'{n} samples' if plan[0].get('total') else f'{len(plan)} sites × {plan[0]["points"]} samples') + \
                    ' – sample and site names are placeholders'
        else:
            parts = ', '.join('%s %d×%d' % (g['code'], g['points'], g['visits']) for g in plan)
        meta['notes'].append(f'Sampling design read from the text: {n} samples ({parts}). '
                             'Exact dates per visit and coordinates are not in the report – add them with a sample list.')
        eb_all = re.findall(r'(\w+) extraction (?:blanks?|negatives?|controls?) were (?:included|processed|extracted)', flat, re.I)
        if eb_all: meta['extraction_blanks'] = sum(num_word(x) or 0 for x in eb_all)
        pb = re.search(r'(\w+) PCR (?:blanks?|negatives?|negative controls?) were included', flat, re.I)
        if pb and not meta.get('pcr_negatives_per_assay'): meta['pcr_negatives_per_assay'] = num_word(pb.group(1)) or 0
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
    'probeReporter', 'probeQuencher', 'probe_seq', 'probe_ref', 'probe_conc',
}
TARGETED_TERMS = {   # 'Targeted assay detection' section + probes: only for targeted (qPCR/dPCR) assays
    'probeReporter', 'probeQuencher', 'probe_seq', 'probe_ref', 'probe_conc', 'amp_vis_method', 'detection_criteria',
    'lod_method', 'pcr_assay_lod_techreps', 'pcr_assay_lod', 'pcr_assay_lod_unit', 'pcr_assay_lod_UL', 'pcr_assay_lod_LL',
    'loq_method', 'pcr_assay_loq_techreps', 'pcr_assay_loq', 'pcr_assay_loq_unit', 'pcr_assay_loq_UL', 'pcr_assay_loq_LL',
    'std_type', 'std_source', 'std_seq', 'thresholdQuantificationCycle', 'baselineValue', 'automaticThresholdQuantificationCycle',
    'automaticBaselineValue', 'targeted_detection_method_additional',
}
ASSAY_TERMS |= TARGETED_TERMS
METABARCODING_TERMS = {'seq_kit', 'trim_param', 'min_len_cutoff', 'otu_db', 'tax_assign_cat', 'otu_seq_comp_appr', 'otu_final_description'}
SEQUENCING_SECTIONS = ('Library preparation sequencing', 'Bioinformatics')
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
        return html.unescape(''.join(re.findall(r'<t[^>]*>([^<]*)</t>', cell_xml_str)))
    v = re.search(r'<v>([^<]*)</v>', cell_xml_str)
    return html.unescape(v.group(1)) if v else ''


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
VOCAB_RULES.update({
    'probeQuencher': [('Black Hole Quencher (BHQ)', r'BHQ|Black Hole'), ('Minor Groove Binder (MGB)', r'MGB|minor groove'),
                      ('Zero-End Quencher (ZEN)', r'\bZEN\b'), ('TAMRA', r'TAMRA'), ('lowa Black', r'Iowa Black|IBFQ|IABkFQ')],
    'std_type': [('synthetic double-stranded DNA', r'synthetic|gBlock|gene fragment|oligo'), ('plasmid', r'plasmid'),
                 ('gDNA', r'genomic|gDNA|tissue DNA'), ('amplicon', r'amplicon|PCR product')],
    'amp_vis_method': [('dPCR', r'\bdd?PCR\b|digital'), ('qPCR', r'qPCR|real[- ]time|quantitative'), ('gel electrophoresis', r'\bgel\b')],
    'pcr_assay_lod_unit': [('copies/reaction', r'cop\w*\s*(?:#\s*)?/\s*(?:reaction|rxn|well)|cop\w* per (?:reaction|rxn|well)'), ('copies/µL', r'cop\w*\s*/\s*[µu]L|cop\w* per [µu]L')],
})
VOCAB_RULES['pcr_assay_loq_unit'] = VOCAB_RULES['pcr_assay_lod_unit']


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
        'study_factor': meta.get('study_factor'),
        'assay_type': meta.get('assay_type') or ' | '.join(sorted({a.get('assay_type') or 'metabarcoding' for a in meta.get('assays') or []})) or 'metabarcoding',
        'sterilise_method': meta.get('sterilise_method'),
        'neg_cont_0_1': '1' if meta.get('neg_controls') else '0', 'pos_cont_0_1': '1' if meta.get('pos_controls') else '0',
        'checkls_ver': version, 'mod_date': time.strftime('%Y-%m-%d'),
        'license': meta.get('license'), 'rightsHolder': meta.get('rights_holder') or meta.get('institution'), 'accessRights': meta.get('access_rights'),
        'bibliographicCitation': meta.get('bibliographicCitation') if re.search(r'10\.\d{4,}/|https?://', str(meta.get('bibliographicCitation') or '')) else None, 'code_repo': meta.get('code_repo'),
        'biological_rep': meta.get('biological_rep'), 'pcr_0_1': '1',
        'pcr_method_additional': meta.get('pcr_method_additional') or '; '.join(n for n in notes if 'normalised' in n) or None,
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
        'tax_class_id_cutoff': meta.get('tax_class_id_cutoff'), 'tax_class_query_cutoff': meta.get('tax_class_query_cutoff'),
        'min_reads_cutoff': meta.get('min_reads_cutoff'), 'min_reads_cutoff_unit': meta.get('min_reads_cutoff_unit'),
        'otu_clust_cutoff': meta.get('otu_clust_cutoff'), 'screen_contam_method': meta.get('screen_contam_method'),
        'otu_raw_description': meta.get('otu_raw_description'),
        'bioinfo_method_additional': meta.get('bioinfo_method_additional'),
        **{t: meta.get(t) for t in EXTRA_PROJECT_TERMS},
    }


EXTRA_PROJECT_TERMS = ('associated_resource', 'checksum_method', 'thermocycler', 'pcr2_annealingTemp', 'pcr2_cycles',
                       'pcr2_thermocycler', 'merge_min_overlap', 'min_len_tool', 'tax_class_collapse', 'tax_class_other',
                       'screen_nontarget_method', 'screen_geograph_method', 'screen_other', 'demux_max_mismatch', 'min_reads_tool',
                       'parent_project_id', 'informationWithheld', 'dataGeneralizations')


def assay_kind(a: dict, meta: dict) -> str:
    return a.get('assay_type') or meta.get('assay_type') or 'metabarcoding'


def assay_value(a: dict, term: str, meta: dict):
    kind = assay_kind(a, meta)
    if term in TARGETED_TERMS and kind != 'targeted': return NA
    if term in METABARCODING_TERMS and kind == 'targeted': return NA
    if term == 'nucl_acid_amp': return a.get('nucl_acid_amp') or meta.get('nucl_acid_amp')
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
        kinds = {assay_kind(a, meta) for a in assays} or {meta.get('assay_type') or 'metabarcoding'}
        if (section == 'Targeted assay detection' or term in TARGETED_TERMS) and 'targeted' not in kinds:
            set_cell(rows, rn, 'D', NA); continue
        if term == 'assay_type' and len(kinds) > 1 and len(assays) > 1 and 'assay_type' not in (meta.get('_overrides') or {}).get('project', {}):
            for i, a in enumerate(assays): set_cell(rows, rn, col_letter(5 + i), assay_kind(a, meta), term=term)
            filled += 1; continue
        if term in ASSAY_TERMS or term.startswith('probe'):
            aov = (meta.get('_overrides') or {}).get('assay', {})
            vals = [aov.get(str(i), {}).get(term, assay_value(a, term, meta)) for i, a in enumerate(assays)]
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
        if term in (meta.get('_overrides') or {}).get('project', {}):
            v = meta['_overrides']['project'][term]
        if present(v):
            set_cell(rows, rn, 'D', v, term=term); filled += 1
        elif kinds == {'targeted'} and section in SEQUENCING_SECTIONS:
            set_cell(rows, rn, 'D', NA)                  # no sequencing in a qPCR/dPCR-only project
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
        'filter_material': meta.get('filter_material'), 'filter_name': meta.get('filter_name'),
        'samp_store_temp': meta.get('samp_store_temp'), 'samp_store_sol': meta.get('samp_store_sol'),
        'prepped_samp_store_temp': meta.get('prepped_samp_store_temp'),
        'prepped_samp_store_sol': meta.get('prepped_samp_store_sol'),
        'samp_vol_we_dna_ext_unit': 'mL',
        'nucl_acid_ext': meta.get('nucl_acid_ext'), 'nucl_acid_ext_kit': meta.get('ext_kit'),
        'nucl_acid_ext_lysis': meta.get('ext_lysis'), 'nucl_acid_ext_sep': meta.get('ext_sep'),
        'nucl_acid_ext_modify': meta.get('ext_modifications'),
        'concentration_unit': 'ng/µl', 'concentration_method': meta.get('conc_method'),
        'detected_notDetected': NA, 'assay_name': ' | '.join(assays),
    }
    plan = meta.get('sample_plan') or []
    if plan:
        fb_ids_by_campaign = {}
        if meta.get('field_blanks_per_site_round'):
            for g in plan:
                key = g.get('campaign') or g['code']
                n_fb = max([x['visits'] for x in plan if (x.get('campaign') or x['code']) == key]) if g.get('campaign') else 1
                fb_ids_by_campaign.setdefault(key, [f'{key.upper().replace(" ", "")}_V{v}_FB' if g.get('campaign') else f'{g["code"]}_FB'
                                                    for v in range(1, n_fb + 1)])
        for g in plan:
            loc = g.get('name')
            geo = (f'{country}: {region}, {loc}' if region else f'{country}: {loc}') if country and loc else None
            key = g.get('campaign') or g['code']
            for v in range(1, g['visits'] + 1):
                sep, w = ('_', 2) if g['code'][-1].isdigit() else ('', 3 if g['points'] < 1000 else 4)
                ids = [f'{g["code"]}{sep}{p:0{w}d}' + (f'_V{v}' if g['visits'] > 1 else '') for p in range(1, g['points'] + 1)] if g['points'] > 1 \
                    else [f'{g["code"]}_V{v}' if g['visits'] > 1 else g['code']]
                fbs = fb_ids_by_campaign.get(key, [])
                fb = fbs[v - 1] if g.get('campaign') and len(fbs) >= v else (fbs[0] if fbs else None)
                verb = ' '.join(x for x in [f'{g["campaign"]} survey' if g.get('campaign') else '',
                                            f'visit {v} of {g["visits"]}' if g['visits'] > 1 else ''] if x) or None
                for sid in ids:
                    rec = dict(common, samp_name=sid, samp_category='sample', neg_cont_type=NA_S, pos_cont_type=NA_S,
                               rel_cont_id=fb, geo_loc_name=geo, eventDate=g.get('date'), verbatimEventDate=verb,
                               env_medium=meta.get('env_medium'), habitat_natural_artificial_0_1='0' if meta.get('habitat_natural', True) else '1',
                               samp_collect_method=meta.get('samp_collect_method'), samp_size=meta.get('samp_size'),
                               samp_size_unit=meta.get('samp_size_unit'), org_note=f'{g["code"]}: {loc}' if loc else None)
                    out.append((rec, assays, False))
        for key, fbs in fb_ids_by_campaign.items():
            dates = [g.get('date') for g in plan if (g.get('campaign') or g['code']) == key]
            for bid in fbs:
                rec = dict(common, samp_name=bid, samp_category='negative control', neg_cont_type='field negative',
                           pos_cont_type=NA_C, rel_cont_id=NA_C, env_broad_scale=NA_C, env_local_scale=NA_C, env_medium=NA_C,
                           habitat_natural_artificial_0_1=NA_C, materialSampleID=NA_C, eventDate=dates[0] if dates else None,
                           samp_collect_method='Field blank of purified water collected alongside the samples')
                out.append((rec, assays, 'field'))
        sites = []
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
    if meta.get('sample_list'):
        out = samples_from_list(meta, common, assays, out)
    sov = (meta.get('_overrides') or {}).get('sample', {})
    if sov:
        for rec, _, control in out:
            for term, v in sov.items():
                if not control or (rec.get(term) not in (None, NA_C, NA_S) and present(rec.get(term))):
                    rec[term] = v
    return out


def repair_package(pkg: dict) -> int:
    """Drop relationships that point to parts missing from the file (the official FAIRe
    template references empty drawing parts that are not in the zip). Returns the count."""
    import posixpath
    removed = 0
    for rels_name in [n for n in pkg if n.startswith('xl/worksheets/_rels/') and n.endswith('.rels')]:
        rels = pkg[rels_name].decode('utf-8')
        base = 'xl/worksheets'
        dead = []
        for m in re.finditer(r'<Relationship\b[^>]*/>', rels):
            el = m.group(0)
            if 'TargetMode="External"' in el: continue
            rid = re.search(r'\bId="([^"]+)"', el).group(1)
            tgt = re.search(r'\bTarget="([^"]+)"', el).group(1)
            path = tgt.lstrip('/') if tgt.startswith('/') else posixpath.normpath(posixpath.join(base, tgt))
            if path not in pkg: dead.append((el, rid))
        if not dead: continue
        sheet_name = 'xl/worksheets/' + posixpath.basename(rels_name)[:-len('.rels')]
        sheet = pkg[sheet_name].decode('utf-8') if sheet_name in pkg else None
        for el, rid in dead:
            rels = rels.replace(el, '')
            if sheet is not None:
                sheet = re.sub(r'<(?:\w+:)?(?:drawing|legacyDrawing|legacyDrawingHF)\b[^>]*\br:id="' + re.escape(rid) + r'"[^>]*/>', '', sheet)
            removed += 1
        pkg[rels_name] = rels.encode('utf-8')
        if sheet is not None: pkg[sheet_name] = sheet.encode('utf-8')
    if '[Content_Types].xml' in pkg:
        ct = pkg['[Content_Types].xml'].decode('utf-8')
        ct2 = re.sub(r'<Override\b[^>]*PartName="/([^"]+)"[^>]*/>', lambda m: m.group(0) if m.group(1) in pkg else '', ct)
        pkg['[Content_Types].xml'] = ct2.encode('utf-8')
    return removed


def fill_workbook(template_path, meta: dict, report_filename: str, output_path) -> dict:
    with zipfile.ZipFile(template_path) as z:
        pkg = {i.filename: z.read(i.filename) for i in z.infolist()}
        infos = z.infolist()
    repair_package(pkg)
    ss = shared_strings(pkg)
    paths = sheet_paths(pkg)
    VOCAB.clear(); VOCAB.update(load_vocab(pkg, paths, ss))
    info = {'vocab': VOCAB, 'types': template_info(template_path).get('types', {})}
    issues = {'fixed': {}, 'problems': {}}
    ver = re.search(r'checklist_v([\d.]+\d)', ' '.join(ss))
    version = ver.group(1) if ver else '1.0'

    # projectMetadata
    p = paths['projectMetadata']; xml = pkg[p].decode('utf-8')
    rows = parse_rows(xml)
    filled = fill_project(rows, ss, meta, version)
    validate_rows('projectMetadata', rows, ss, info, issues)
    pkg[p] = replace_sheet_data(xml, rows).encode('utf-8')

    # sampleMetadata
    p = paths['sampleMetadata']; xml = pkg[p].decode('utf-8')
    rows = parse_rows(xml)
    hdr, req = header_map(rows, ss)
    rows = {k: v for k, v in rows.items() if k <= 3}
    samples = build_samples(meta)
    extra_cols = []                                   # user-defined columns (FAIRe allows extra terms at the end)
    for rec, _, _ in samples:
        for k in (rec.get('_extra') or {}):
            if k not in hdr and k not in extra_cols: extra_cols.append(k)
    if extra_cols:
        last = max(col_num(c) for c in rows[3]['cells'])
        for j, k in enumerate(extra_cols, start=1):
            col = col_letter(last + j)
            rows[1]['cells'][col] = cell_xml(f'{col}1', 'O'); rows[2]['cells'][col] = cell_xml(f'{col}2', 'User defined')
            rows[3]['cells'][col] = cell_xml(f'{col}3', k); hdr[k] = col; req[col] = 'O'
    for i, (rec, _, control) in enumerate(samples):
        write_record(rows, 4 + i, dict(rec, **(rec.get('_extra') or {})), hdr, req, control)
    validate_rows('sampleMetadata', rows, ss, info, issues)
    sample_rows = rows
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
    no_run = {a['assay_name'] for a in meta.get('assays') or [] if assay_kind(a, meta) == 'targeted'}   # qPCR/dPCR: no sequencing run
    rn = 4
    for rec, assays, control in samples:
        for an in assays:
            if an in no_run: continue
            run = {'samp_name': rec['samp_name'], 'assay_name': an, 'lib_id': f'{rec["samp_name"]}_{an}',
                   'seq_run_id': run_for.get(marker_of.get(an)), 'lib_conc_meth': meta.get('lib_conc_meth'),
                   'phix_perc': meta.get('phix_perc')}
            run.update({k: v for k, v in (rec.get('_run') or {}).items() if present(v)})
            if (rec.get('_run') or {}).get('lib_id') and len(assays) > 1: run['lib_id'] = f'{rec["_run"]["lib_id"]}_{an}'
            write_record(rows, rn, run, hdr, req)
            rn += 1
    n_runs = rn - 4
    validate_rows('experimentRunMetadata', rows, ss, info, issues)
    check_links(sample_rows, rows, ss, issues)
    pkg[p] = replace_sheet_data(xml, rows).encode('utf-8')

    # taxaRaw / taxaFinal (only when taxon tables were given)
    taxa_counts = {}
    for sheet_name, key in (('taxaRaw', 'taxa_raw'), ('taxaFinal', 'taxa_final')):
        recs = meta.get(key) or []
        if not recs or sheet_name not in paths: continue
        p = paths[sheet_name]; xml = pkg[p].decode('utf-8')
        rows = parse_rows(xml)
        hdr, req = header_map(rows, ss)
        rows = {k: v for k, v in rows.items() if k <= 3}
        for i, rec in enumerate(recs):
            write_record(rows, 4 + i, rec, hdr, req)
        validate_rows(sheet_name, rows, ss, info, issues)
        pkg[p] = replace_sheet_data(xml, rows).encode('utf-8')
        taxa_counts[sheet_name] = len(recs)

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
        'notes': meta.get('notes') or [], 'validation': summarise_issues(issues), 'taxa': taxa_counts,
    }


# ── review step: values + source sentences, user edits as overrides ──────────

SECTIONS = [
    ('Project', ['recordedBy', 'recordedByID', 'project_contact', 'institution', 'institutionID', 'project_name', 'project_id',
                 'parent_project_id', 'associated_resource', 'informationWithheld', 'dataGeneralizations',
                 'study_factor', 'assay_type', 'license', 'rightsHolder', 'accessRights', 'bibliographicCitation', 'code_repo']),
    ('Controls & contamination', ['sterilise_method', 'neg_cont_0_1', 'pos_cont_0_1', 'screen_contam_method']),
    ('Sampling', ['biological_rep', 'samp_collect_method', 'samp_size', 'samp_size_unit', 'env_broad_scale', 'env_local_scale',
                  'env_medium', 'habitat_natural_artificial_0_1']),
    ('Filtration & storage', ['samp_mat_process', 'filter_passive_active_0_1', 'size_frac', 'filter_diameter', 'filter_material',
                              'filter_name', 'samp_store_temp', 'samp_store_sol',
                              'samp_store_dur', 'samp_store_loc', 'samp_store_method_additional', 'prepped_samp_store_temp',
                              'prepped_samp_store_sol']),
    ('DNA extraction', ['nucl_acid_ext', 'nucl_acid_ext_kit', 'nucl_acid_ext_lysis', 'nucl_acid_ext_sep', 'nucl_acid_ext_modify',
                        'concentration_method']),
    ('PCR', ['pcr_0_1', 'assay_name', 'target_gene', 'targetTaxonomicAssay', 'targetTaxonomicScope', 'pcr_primer_name_forward',
             'pcr_primer_name_reverse', 'pcr_primer_forward', 'pcr_primer_reverse', 'pcr_primer_reference_forward',
             'pcr_primer_reference_reverse', 'pcr_primer_vol_forward', 'pcr_primer_vol_reverse', 'pcr_primer_conc_forward',
             'pcr_primer_conc_reverse', 'ampliconSize', 'commercial_mm', 'custom_mm', 'amplificationReactionVolume', 'pcr_dna_vol',
             'pcr_rep', 'nucl_acid_amp', 'pcr_cond', 'annealingTemp', 'pcr_cycles', 'thermocycler', 'pcr_method_additional',
             'probe_seq', 'probeReporter', 'probeQuencher', 'probe_ref', 'probe_conc']),
    ('Targeted detection (qPCR/dPCR)', ['amp_vis_method', 'detection_criteria', 'lod_method', 'pcr_assay_lod_techreps', 'pcr_assay_lod',
                                        'pcr_assay_lod_unit', 'pcr_assay_lod_UL', 'pcr_assay_lod_LL', 'loq_method', 'pcr_assay_loq_techreps',
                                        'pcr_assay_loq', 'pcr_assay_loq_unit', 'pcr_assay_loq_UL', 'pcr_assay_loq_LL', 'std_type', 'std_source',
                                        'std_seq', 'thresholdQuantificationCycle', 'baselineValue', 'automaticThresholdQuantificationCycle',
                                        'automaticBaselineValue', 'targeted_detection_method_additional']),
    ('Library preparation & sequencing', ['barcoding_pcr_appr', 'pcr2_annealingTemp', 'pcr2_cycles', 'pcr2_thermocycler', 'checksum_method', 'pcr2_commercial_mm', 'pcr2_custom_mm', 'pcr2_amplificationReactionVolume',
                                          'pcr2_dna_vol', 'pcr2_cond', 'platform', 'instrument', 'seq_kit', 'lib_layout',
                                          'adapter_forward', 'adapter_reverse', 'lib_screen', 'sequencing_location',
                                          'seq_method_additional']),
    ('Bioinformatics', ['sop_bioinformatics', 'trim_method', 'trim_param', 'demux_tool', 'demux_max_mismatch', 'merge_tool',
                        'merge_min_overlap', 'min_len_cutoff', 'min_len_tool', 'min_reads_tool', 'tax_class_collapse', 'tax_class_other',
                        'screen_nontarget_method', 'screen_geograph_method', 'screen_other',
                        'error_rate_tool', 'error_rate_type', 'error_rate_cutoff', 'chimera_check_method', 'chimera_check_param',
                        'otu_clust_tool', 'otu_clust_cutoff', 'min_reads_cutoff', 'min_reads_cutoff_unit', 'otu_db', 'tax_assign_cat',
                        'otu_seq_comp_appr', 'tax_class_id_cutoff', 'tax_class_query_cutoff',
                        'otu_raw_description', 'otu_final_description', 'bioinfo_method_additional']),
]
SAMPLE_TERMS = ['samp_collect_method', 'samp_size', 'samp_size_unit', 'env_broad_scale', 'env_local_scale', 'env_medium',
                'habitat_natural_artificial_0_1', 'samp_mat_process', 'filter_passive_active_0_1', 'size_frac', 'filter_diameter',
                'filter_material', 'filter_name', 'samp_store_temp', 'samp_store_sol',
                'samp_store_dur', 'samp_store_loc', 'samp_store_method_additional', 'prepped_samp_store_temp',
                'prepped_samp_store_sol', 'nucl_acid_ext', 'nucl_acid_ext_kit', 'nucl_acid_ext_lysis', 'nucl_acid_ext_sep',
                'nucl_acid_ext_modify', 'concentration_method']
GENERATED = {'project_id', 'assay_type', 'pcr_0_1', 'neg_cont_0_1', 'pos_cont_0_1', 'checkls_ver', 'mod_date', 'rightsHolder',
             'assay_name', 'target_gene', 'habitat_natural_artificial_0_1', 'filter_passive_active_0_1',
             'otu_raw_description', 'tax_assign_cat'}
# FAIRe term -> key under which the extractor recorded its source sentence
EVIDENCE_KEY = {
    'recordedBy': 'recorded_by', 'project_contact': 'project_contact', 'rightsHolder': 'rights_holder',
    'institution': 'institution', 'institutionID': 'institution', 'project_name': 'project_title',
    'bibliographicCitation': 'project_title',
    'sterilise_method': 'sterilise_method', 'neg_cont_0_1': 'neg_controls', 'pos_cont_0_1': 'pos_controls',
    'screen_contam_method': 'screen_contam_method', 'biological_rep': 'biological_rep',
    'samp_collect_method': 'samp_collect_method', 'samp_size': 'samp_size', 'samp_size_unit': 'samp_size_unit',
    'env_broad_scale': 'env_broad_scale', 'env_local_scale': 'env_local_scale', 'env_medium': 'env_medium',
    'samp_mat_process': 'samp_mat_process', 'filter_passive_active_0_1': 'filter_active', 'size_frac': 'size_frac',
    'filter_diameter': 'filter_diameter', 'filter_material': 'filter_material', 'samp_store_dur': 'samp_store_dur',
    'samp_store_loc': 'samp_store_loc', 'samp_store_method_additional': 'samp_store_method_additional',
    'prepped_samp_store_temp': 'prepped_samp_store_temp', 'prepped_samp_store_sol': 'prepped_samp_store_sol',
    'nucl_acid_ext': 'nucl_acid_ext', 'nucl_acid_ext_kit': 'ext_kit', 'nucl_acid_ext_lysis': 'ext_lysis',
    'nucl_acid_ext_sep': 'ext_sep', 'nucl_acid_ext_modify': 'ext_modifications', 'concentration_method': 'conc_method',
    'barcoding_pcr_appr': 'barcoding_pcr_appr', 'pcr2_commercial_mm': 'pcr2_commercial_mm', 'pcr2_custom_mm': 'pcr2_custom_mm',
    'pcr2_amplificationReactionVolume': 'pcr2_amplificationReactionVolume', 'pcr2_dna_vol': 'pcr2_dna_vol',
    'platform': 'platform', 'instrument': 'instrument', 'lib_layout': 'lib_layout', 'adapter_forward': 'adapter_forward',
    'adapter_reverse': 'adapter_reverse', 'lib_screen': 'lib_screen', 'sequencing_location': 'sequencing_location',
    'sop_bioinformatics': 'sop_bioinformatics', 'trim_method': 'trim_method', 'merge_tool': 'merge_tool',
    'error_rate_tool': 'error_rate_tool', 'error_rate_type': 'error_rate_type', 'error_rate_cutoff': 'error_rate_cutoff',
    'chimera_check_method': 'chimera_check_method', 'chimera_check_param': 'chimera_check_param',
    'otu_clust_tool': 'otu_clust_tool', 'otu_clust_cutoff': 'otu_clust_cutoff', 'otu_raw_description': 'otu_raw_description',
    'bioinfo_method_additional': 'bioinfo_method_additional', 'nucl_acid_amp': 'nucl_acid_amp',
    'seq_method_additional': 'phix_perc', 'filter_name': 'filter_name', 'samp_store_temp': 'samp_store_temp',
    'samp_store_sol': 'samp_store_sol', 'tax_class_id_cutoff': 'tax_class_id_cutoff', 'tax_class_query_cutoff': 'tax_class_query_cutoff',
    'min_reads_cutoff': 'min_reads_cutoff', 'min_reads_cutoff_unit': 'min_reads_cutoff_unit',
}
ASSAY_EVIDENCE_KEY = {'pcr_primer_name_forward': 'primer_name_fwd', 'pcr_primer_name_reverse': 'primer_name_rev',
                      'pcr_primer_reference_forward': 'primer_ref_fwd', 'pcr_primer_reference_reverse': 'primer_ref_rev',
                      'pcr_primer_vol_forward': 'pcr_primer_vol', 'pcr_primer_vol_reverse': 'pcr_primer_vol',
                      'pcr_primer_conc_forward': 'pcr_primer_conc', 'pcr_primer_conc_reverse': 'pcr_primer_conc',
                      'target_gene': 'primer_name_fwd', 'assay_name': 'assay_name'}


def _evidence_for(meta, term, value, assay=None):
    if value in (None, '') or str(value).startswith(('missing:', 'not applicable')): return None
    ev = meta.get('_evidence') or {}
    if assay is not None:
        key = ASSAY_EVIDENCE_KEY.get(term) or ASSAY_KEY.get(term, term)
        return ev.get(f"{assay['marker']}.{key}") or (ev.get(EVIDENCE_KEY[term]) if term in EVIDENCE_KEY else None)
    return ev.get(EVIDENCE_KEY.get(term, term))


def _section_of(term):
    for name, terms in SECTIONS:
        if term in terms: return name
    return 'Other'


def review_items(meta: dict, text: str = '') -> list:
    """Everything the tool would write, with the sentence each value came from, for the user to check."""
    ov = meta.get('_overrides') or {}
    items = []
    pv = project_values(meta, '1')
    for term, value in pv.items():
        if term in ('checkls_ver', 'mod_date'): continue
        if term in ov.get('project', {}): value = ov['project'][term]
        if term == 'pcr_method_additional': continue
        items.append({'id': f'project:{term}', 'section': _section_of(term), 'term': term, 'scope': 'project',
                      'value': '' if value is None else str(value), 'generated': term in GENERATED,
                      'evidence': _evidence_for(meta, term, value)})
    for i, a in enumerate(meta.get('assays') or []):
        for term in sorted(ASSAY_TERMS, key=lambda t: [x for _, ts in SECTIONS for x in ts].index(t) if any(t in ts for _, ts in SECTIONS) else 999):
            value = ov.get('assay', {}).get(str(i), {}).get(term, assay_value(a, term, meta))
            if value == NA and (term in TARGETED_TERMS or term in METABARCODING_TERMS): continue
            items.append({'id': f'assay:{i}:{term}', 'section': _section_of(term), 'term': term, 'scope': 'assay',
                          'assay': a.get('assay_name'), 'marker': a.get('marker'), 'value': '' if value is None else str(value),
                          'generated': term in GENERATED,
                          'evidence': _evidence_for(meta, term, value, a)})
    samples = build_samples(meta)
    first = next((rec for rec, _, c in samples if not c), {})
    for term in SAMPLE_TERMS:
        value = ov.get('sample', {}).get(term, first.get(term))
        items.append({'id': f'sample:{term}', 'section': _section_of(term), 'term': term, 'scope': 'samples',
                      'value': '' if value is None else str(value), 'generated': term in GENERATED,
                      'evidence': _evidence_for(meta, term, value)})
    order = [name for name, _ in SECTIONS] + ['Other']
    items.sort(key=lambda x: order.index(x['section']))
    return items


def apply_review(meta: dict, edits: dict) -> dict:
    """edits: {item id: new value}. An empty value means 'not stated' (written as missing wording)."""
    ov = meta.setdefault('_overrides', {'project': {}, 'assay': {}, 'sample': {}})
    for key, value in (edits or {}).items():
        value = None if value is None or str(value).strip() == '' else str(value).strip()
        parts = key.split(':')
        if parts[0] == 'project': ov.setdefault('project', {})[parts[1]] = value
        elif parts[0] == 'assay': ov.setdefault('assay', {}).setdefault(parts[1], {})[parts[2]] = value
        elif parts[0] == 'sample': ov.setdefault('sample', {})[parts[1]] = value
    return meta



# ── optional sample list (CSV / Excel) with one row per sample ───────────────

SAMPLE_ALIASES = {
    'samp_name': r'samp_name|sample(?:[ _]?(?:id|name|code|no\.?|number))?|id|name',
    'eventDate': r'event_?date|date|sampling[ _]date|collection[ _]date|date[ _]collected',
    'decimalLatitude': r'decimal_?latitude|lat(?:itude)?(?:[ _]?\(?dd\)?)?',
    'decimalLongitude': r'decimal_?longitude|lon(?:g|gitude)?(?:[ _]?\(?dd\)?)?',
    'site': r'site(?:[ _]?(?:name|code))?|location|locality|station|lake|river|pond',
    'samp_category': r'samp_category|category|sample[ _]type|type',
    'samp_vol_we_dna_ext': r'samp_vol_we_dna_ext|volume[ _]filtered(?:[ _]?\(?ml\)?)?|filtered[ _]volume(?:[ _]?\(?ml\)?)?|vol(?:ume)?(?:[ _]?\(?ml\)?)?',
    'samp_size': r'samp_size|sample[ _]volume(?:[ _]?\(?ml\)?)?|volume[ _]collected(?:[ _]?\(?ml\)?)?',
    'depth': r'depth(?:[ _]?\(?m\)?)?',
    'notes': r'notes?|comments?|remarks?',
    'samp_collect_method': r'collection[ _]method|sampling[ _]method|collect(?:ion)?[ _]device',
    'samp_mat_process': r'filtration[ _]method|filter(?:ing)?[ _]method|processing(?:[ _]method)?',
    'samp_store_method_additional': r'storage(?:[ _]method)?|preservation',
    'temp': r'water[ _]?temp(?:erature)?(?:[ _]?\(?°?c\)?)?|temp(?:erature)?(?:[ _]?\(?°?c\)?)?',
    'seq_run_id': r'library|library[ _]name|lib|sub[ _-]?library|seq(?:uencing)?[ _]run(?:[ _]id)?|run(?:[ _]id)?',
    'habitat': r'habitat(?:[ _]type)?|substrate',
}
RUN_TERMS = {'pcr_plate_id', 'lib_id', 'seq_run_id', 'lib_conc', 'lib_conc_unit', 'mid_forward', 'mid_reverse', 'filename',
             'filename2', 'checksum_filename', 'checksum_filename2', 'associatedSequences', 'input_read_count', 'output_read_count',
             'output_otu_num', 'otu_num_tax_assigned'}


def read_table(path) -> list:
    """Rows of a .csv / .tsv / .xlsx file as dicts (header = first row that names the sample column)."""
    import csv, io
    path = Path(path)
    if path.suffix.lower() in ('.xlsx', '.xlsm'):
        with zipfile.ZipFile(path) as z:
            pkg = {n: z.read(n) for n in z.namelist()}
        ss, paths = shared_strings(pkg), sheet_paths(pkg)
        sheet = paths.get('sampleMetadata') or next(iter(paths.values()))
        rows = parse_rows(pkg[sheet].decode('utf-8'))
        table = [[cell_text(r['cells'].get(col_letter(c), ''), ss) for c in range(1, max((col_num(k) for k in r['cells']), default=0) + 1)]
                 for _, r in sorted(rows.items())]
    else:
        raw = path.read_bytes().decode('utf-8-sig', errors='replace')
        table = list(csv.reader(io.StringIO(raw), dialect=csv.Sniffer().sniff(raw[:4000], delimiters=',;\t')))
    for i, row in enumerate(table[:10]):
        if any(re.fullmatch(SAMPLE_ALIASES['samp_name'], str(c).strip(), re.I) for c in row):
            hdr = [str(c).strip() for c in row]
            return [dict(zip(hdr, r)) for r in table[i + 1:] if any(str(x).strip() for x in r)]
    raise ValueError('The sample list needs a column called "sample" or "samp_name" (one row per sample).')


CONTROL_CODES: dict = {}      # code -> kind, read from a data README ("EB = extraction blank")


def control_codes_from(text: str) -> dict:
    codes = {}
    kinds = [(r'PCR\s*neg\w*|no[- ]template|NTC', 'PCR negative'), (r'PCR\s*pos\w*|positive', 'positive control'),
             (r'extraction\s*(?:blank|neg\w*)', 'extraction negative'), (r'filter\s*blank', 'process negative'),
             (r'field\s*(?:blank|neg\w*)', 'field negative'), (r'lysis\s*blank', 'extraction negative')]
    for m in re.finditer(r"['‘’\"]?\b([A-Z]{2,5})\b['‘’\"]?\s*(?:\((?:e\.g\.\s*)?([^)]{0,40})\)|\(e\.g\.[^)]*\))?\s*(?:are|is|=|denotes|:)?\s*(?:an?\s+)?([A-Za-z -]{3,40})?", text):
        code, inner, after = m.group(1), m.group(2) or '', m.group(3) or ''
        if code in ('PCR', 'DNA', 'RNA', 'UK', 'USA'): continue
        for pat, kind in kinds:
            if re.match(pat, inner.strip(), re.I) or re.match(pat, after.strip(), re.I):
                codes.setdefault(code, kind); break
    return codes


def _blank_kind(name, cat):
    up = str(name).upper()
    for code, kind in sorted(CONTROL_CODES.items(), key=lambda kv: -len(kv[0])):
        if up.startswith(code) or re.search(code + r'\d*[A-Z]?$', up):
            return kind
    t = re.sub(r'[_\-.]', ' ', f'{name} {cat}'.lower())
    if re.search(r'positive|mock|^pos\d|\bpos\b', t): return 'positive control'
    if not re.search(r'blank|negative|\bneg\d*\b|^neg|\bntc\b|control|\bfb\d*\b|\beb\d*\b|\bnc\d*\b|bk\d*\b', t): return None
    if re.search(r'\bpcr\b|pcrneg|\bntc\b|no template|^neg', t): return 'PCR negative'
    if re.search(r'extract|\bext\b|eb\d*\b', t): return 'extraction negative'
    return 'field negative'


def _note(meta, text):
    if text not in meta.setdefault('notes', []): meta['notes'].append(text)


def samples_from_list(meta, common, assays, generated):
    """One sampleMetadata row per row of the user's sample list, with the report's methods filled in."""
    terms_known = set(common) | {'samp_name', 'samp_category', 'neg_cont_type', 'pos_cont_type', 'rel_cont_id', 'geo_loc_name',
                                 'eventDate', 'verbatimEventDate', 'decimalLatitude', 'decimalLongitude', 'env_medium',
                                 'samp_collect_method', 'samp_size', 'samp_size_unit', 'minimumDepthInMeters', 'maximumDepthInMeters'}
    hdr_terms = set(meta.get('_sample_headers') or []) | terms_known
    country, region = meta.get('country'), meta.get('region')
    out, ignored = [], set()
    for row in meta['sample_list']:
        r, unmatched = {}, {}
        for k, v in row.items():
            v = str(v or '').strip(); k0 = str(k or '').strip()
            if not k0 or not v: continue
            term = next((t for t in list(hdr_terms) + list(RUN_TERMS) if t.lower() == k0.lower()), None) or \
                   next((t for t, pat in SAMPLE_ALIASES.items() if re.fullmatch(pat, k0, re.I)), None)
            if term: r[term] = v
            else: ignored.add(k0); unmatched[re.sub(r'\W+', '_', k0).strip('_')] = v
        if not r.get('samp_name'): continue
        cat = r.pop('samp_category', '')
        kind = _blank_kind(r['samp_name'], f'{cat} {r.get("notes", "")}')
        site = r.pop('site', None)
        extra = []
        if cat and not re.search(r'sample|control|blank|neg|pos|mock', cat, re.I): extra.append(f'type: {cat}'); cat = ''
        if r.get('habitat'): extra.append(f'habitat: {r.pop("habitat")}')
        if site and (len(site) <= 5 or re.fullmatch(r'(north|south|east|west|upper|lower|centre|center)\w*', site, re.I)):
            extra.append(f'location: {site}'); site = None
        rec = dict(common)
        if kind == 'positive control':
            rec.update(samp_category='positive control', pos_cont_type=meta.get('pos_cont_desc') or 'positive control (see projectMetadata)',
                       neg_cont_type=NA_C, rel_cont_id=NA_C, env_broad_scale=NA_C, env_local_scale=NA_C, env_medium=NA_C,
                       habitat_natural_artificial_0_1=NA_C)
        elif kind:
            rec.update(samp_category='negative control', neg_cont_type=kind, pos_cont_type=NA_C, rel_cont_id=NA_C,
                       env_broad_scale=NA_C, env_local_scale=NA_C, env_medium=NA_C, habitat_natural_artificial_0_1=NA_C)
        elif re.search(r'positive|mock', cat, re.I):
            rec.update(samp_category='positive control', pos_cont_type=cat, neg_cont_type=NA_C)
        else:
            rec.update(samp_category='sample', neg_cont_type=NA_S, pos_cont_type=NA_S, env_medium=meta.get('env_medium'),
                       habitat_natural_artificial_0_1='0' if meta.get('habitat_natural', True) else '1',
                       samp_collect_method=meta.get('samp_collect_method'), samp_size=meta.get('samp_size'),
                       samp_size_unit=meta.get('samp_size_unit'))
        if site:
            rec['geo_loc_name'] = (f'{country}: {region}, {site}' if region else f'{country}: {site}') if country else site
        if r.get('depth'): rec['minimumDepthInMeters'] = rec['maximumDepthInMeters'] = r.pop('depth')
        rec['_extra'] = dict(unmatched)
        for e in extra:
            k, v = e.split(': ', 1); rec['_extra'][{'type': 'sample_type'}.get(k, k)] = v
        if r.get('notes'): rec['_extra']['notes'] = r.pop('notes')
        links = (meta.get('_blank_links') or {}).get(r['samp_name'])
        if links and not kind: rec['rel_cont_id'] = ' | '.join(dict.fromkeys(links))
        mx = meta.get('_read_matrix')
        if mx and r['samp_name'] in mx['samples']:
            counts = [t['counts'].get(r['samp_name'], 0) for t in mx['taxa']]
            r.setdefault('output_read_count', f'{sum(counts):g}')
            r.setdefault('output_otu_num', str(sum(1 for c in counts if c)))
        lib = r.get('seq_run_id')
        sra = (meta.get('sra_by_sample') or {}).get(r['samp_name']) or (meta.get('sra_by_library') or {}).get(lib or '')
        if sra:
            files = [f.split('/')[-1] for f in (sra.get('fastq_ftp') or '').split(';') if f]
            md5 = [m for m in (sra.get('fastq_md5') or '').split(';') if m]
            r.setdefault('seq_run_id', sra.get('library_name') or sra.get('run_accession'))
            r.setdefault('associatedSequences', sra.get('run_accession'))
            if files: r.setdefault('filename', files[0])
            if len(files) > 1: r.setdefault('filename2', files[1])
            if md5: r.setdefault('checksum_filename', md5[0])
            if len(md5) > 1: r.setdefault('checksum_filename2', md5[1])
            if (meta.get('sra_by_sample') or {}).get(r['samp_name']) and sra.get('read_count'): r.setdefault('input_read_count', sra['read_count'])
        if re.fullmatch(r'\d{5}(?:\.\d+)?', r.get('eventDate', '')):        # Excel date serial number
            r['eventDate'] = (date(1899, 12, 30) + timedelta(days=int(float(r['eventDate'])))).isoformat()
        if r.get('eventDate') and not re.match(r'\d{4}-\d{2}', r['eventDate']):
            d = parse_date(r['eventDate'])                      # day first (UK style) for 07/11/2025
            if d: r['eventDate'] = d; meta['_dayfirst'] = True
        elif r.get('eventDate'): r['eventDate'] = r['eventDate'][:10] if re.match(r'\d{4}-\d{2}-\d{2} 00:00', r['eventDate']) else r['eventDate']
        if r.get('samp_vol_we_dna_ext'): rec['samp_vol_we_dna_ext_unit'] = 'mL'
        run = {k: r.pop(k) for k in list(r) if k in RUN_TERMS}
        rec.update(r); rec['_run'] = run
        an = [a.strip() for a in str(rec.get('assay_name') or '').split('|') if a.strip()]
        out.append((rec, an or assays, 'field' if kind == 'field negative' else 'lab' if kind and kind != 'positive control' else False))
    if ignored:
        _note(meta, 'Sample list columns with no FAIRe term were added as user-defined columns at the end of sampleMetadata: ' + ', '.join(sorted(ignored)[:15]))
    if meta.pop('_dayfirst', None):
        _note(meta, 'Dates in the sample list were read day-first (e.g. 07/11/2025 = 7 November 2025).')
    _note(meta, f'{len(out)} sample rows taken from your sample list (instead of the {len(generated)} estimated from the report).')
    return out


# ════════════════════════════════════════════════════════════════════════════
#  4. Automatic checks: primer library, validation, confidence, profiles, second opinions
# ════════════════════════════════════════════════════════════════════════════

# ── primer & reference-database library ──────────────────────────────────────
# Published primers commonly used for eDNA metabarcoding. Used to (a) name and cite a
# primer whose sequence is printed in the report, (b) fill the sequence of a primer that
# is only named (flagged for checking), (c) flag sequences that differ from the published one.
# To add one: append (names, marker, direction, sequence 5'-3', reference).
PRIMER_LIBRARY = [
    (['MiFish-U-F', 'MiFish-UF', 'MiFishU-F'], '12S', 'fwd', 'GTCGGTAAAACTCGTGCCAGC', 'https://doi.org/10.1098/rsos.150088'),
    (['MiFish-U-R', 'MiFish-UR', 'MiFishU-R'], '12S', 'rev', 'CATAGTGGGGTATCTAATCCCAGTTTG', 'https://doi.org/10.1098/rsos.150088'),
    (['MiFish-E-F', 'MiFish-EF'], '12S', 'fwd', 'GTTGGTAAATCTCGTGCCAGC', 'https://doi.org/10.1098/rsos.150088'),
    (['MiFish-E-R', 'MiFish-ER'], '12S', 'rev', 'CATAGTGGGGTATCTAATCCTAGTTTG', 'https://doi.org/10.1098/rsos.150088'),
    (['12S-V5-F', '12SV5F', '12S-V5F', '12SV5-F'], '12S', 'fwd', 'ACTGGGATTAGATACCCC', 'https://doi.org/10.1093/nar/gkr732'),
    (['12S-V5-R', '12SV5R', '12S-V5R', '12SV5-R'], '12S', 'rev', 'TAGAACAGGCTCCTCTAG', 'https://doi.org/10.1093/nar/gkr732'),
    (['teleo_F', 'teleo-F', 'Tele02-F', 'Tele02F'], '12S', 'fwd', 'ACACCGCCCGTCACTCT', 'https://doi.org/10.1111/mec.13428'),
    (['teleo_R', 'teleo-R', 'Tele02-R', 'Tele02R'], '12S', 'rev', 'CTTCCGGTACACTTACCATG', 'https://doi.org/10.1111/mec.13428'),
    (['mlCOIintF'], 'COI', 'fwd', 'GGWACWGGWTGAACWGTWTAYCCYCC', 'https://doi.org/10.1186/1742-9994-10-34'),
    (['mlCOIintF-XT'], 'COI', 'fwd', 'GGWACWRGWTGRACWITITAYCCYCC', 'https://doi.org/10.7717/peerj.4705'),
    (['jgHCO2198', 'jgHCO'], 'COI', 'rev', 'TAIACYTCIGGRTGICCRAARAAYCA', 'https://doi.org/10.1111/1755-0998.12138'),
    (['LCO1490'], 'COI', 'fwd', 'GGTCAACAAATCATAAAGATATTGG', 'https://pubmed.ncbi.nlm.nih.gov/7881515/'),
    (['HCO2198'], 'COI', 'rev', 'TAAACTTCAGGGTGACCAAAAAATCA', 'https://pubmed.ncbi.nlm.nih.gov/7881515/'),
    (['fwhF2'], 'COI', 'fwd', 'GGDACWGGWTGAACWGTWTAYCCHCC', 'https://doi.org/10.3897/mbmg.1.14625'),
    (['fwhR2n'], 'COI', 'rev', 'GTRATWGCHCCDGCTARWACWGG', 'https://doi.org/10.3897/mbmg.1.14625'),
    (['BF3'], 'COI', 'fwd', 'CCHGAYATRGCHTTYCCHCG', 'https://doi.org/10.3389/fenvs.2017.00011'),
    (['BR2'], 'COI', 'rev', 'TCDGGRTGNCCRAARAAYCA', 'https://doi.org/10.3389/fenvs.2017.00011'),
    (['515F-Y', '515FY', '515F', '515f', '515F (Parada)'], '16S', 'fwd', 'GTGYCAGCMGCCGCGGTAA', 'https://doi.org/10.1111/1462-2920.13023'),
    (['926R', '926r'], '16S', 'rev', 'CCGYCAATTYMTTTRAGTTT', 'https://doi.org/10.1111/1462-2920.13023'),
    (['806RB', '806R (Apprill)'], '16S', 'rev', 'GGACTACNVGGGTWTCTAAT', 'https://doi.org/10.3354/ame01753'),
    (['515F (Caporaso)'], '16S', 'fwd', 'GTGCCAGCMGCCGCGGTAA', 'https://doi.org/10.1073/pnas.1000080107'),
    (['806R', '806R (Caporaso)'], '16S', 'rev', 'GGACTACHVGGGTWTCTAAT', 'https://doi.org/10.1073/pnas.1000080107'),
    (['341F', 'S-D-Bact-0341-b-S-17'], '16S', 'fwd', 'CCTACGGGNGGCWGCAG', 'https://doi.org/10.1093/nar/gks808'),
    (['785R', '805R', 'S-D-Bact-0785-a-A-21'], '16S', 'rev', 'GACTACHVGGGTATCTAATCC', 'https://doi.org/10.1093/nar/gks808'),
    (['1391f', '1391F', '1391_F'], '18S', 'fwd', 'GTACACACCGCCCGTC', 'https://doi.org/10.1371/journal.pone.0006372'),
    (['EukBr', 'EukB', 'EukBR'], '18S', 'rev', 'TGATCCTTCTGCAGGTTCACCTAC', 'https://doi.org/10.1371/journal.pone.0006372'),
    (['TAReuk454FWD1', 'TAReukFWD1'], '18S', 'fwd', 'CCAGCASCYGCGGTAATTCC', 'https://doi.org/10.1111/j.1365-294X.2009.04480.x'),
    (['TAReukREV3'], '18S', 'rev', 'ACTTTCGTTCTTGATYRA', 'https://doi.org/10.1111/j.1365-294X.2009.04480.x'),
    (['Diat_rbcL_708F_1', 'Diat_rbcL_708F'], 'rbcL', 'fwd', 'AGGTGAAGTAAAAGGTTCWTACTTAAA', 'https://doi.org/10.1016/j.ecolind.2017.06.024'),
    (['R3_1', 'Diat_rbcL_R3_1'], 'rbcL', 'rev', 'CCTTCTAATTTACCWACWACTG', 'https://doi.org/10.1016/j.ecolind.2017.06.024'),
    (['ITS1F', 'ITS1-F'], 'ITS', 'fwd', 'CTTGGTCATTTAGAGGAAGTAA', 'https://doi.org/10.1111/j.1365-294X.1993.tb00005.x'),
    (['ITS2'], 'ITS', 'rev', 'GCTGCGTTCTTCATCGATGC', None),
    (['ITS3'], 'ITS', 'fwd', 'GCATCGATGAAGAACGCAGC', None),
    (['ITS4'], 'ITS', 'rev', 'TCCTCCGCTTATTGATATGC', None),
    (['fITS7'], 'ITS', 'fwd', 'GTGARTCATCGAATCTTTG', 'https://doi.org/10.1111/j.1574-6941.2012.01437.x'),
    (['gITS7'], 'ITS', 'fwd', 'GTGARTCATCGARTCTTTG', 'https://doi.org/10.1111/j.1574-6941.2012.01437.x'),
    (['trnL-g', 'g', 'trnL g'], 'trnL', 'fwd', 'GGGCAATCCTGAGCCAA', 'https://doi.org/10.1093/nar/gkl938'),
    (['trnL-h', 'h', 'trnL h'], 'trnL', 'rev', 'CCATTGAGTCTCTGCACCTATC', 'https://doi.org/10.1093/nar/gkl938'),
]
# Primer pairs often named only as a set in reports ("12SrRNA V5 primers (Riaz et al., 2011)")
PRIMER_PAIRS = [
    (r'12S\s*-?\s*(?:rRNA\s*)?V5|12SV5|Riaz et al', '12S', '12S-V5-F', '12S-V5-R'),
    (r'MiFish[- ]?U\b|MiFish\s+primers', '12S', 'MiFish-U-F', 'MiFish-U-R'),
    (r'MiFish[- ]?E\b', '12S', 'MiFish-E-F', 'MiFish-E-R'),
    (r'Tele0?2\b|teleo\b', '12S', 'teleo_F', 'teleo_R'),
    (r'Leray\b|mlCOIintF', 'COI', 'mlCOIintF', 'jgHCO2198'),
    (r'\bfwh\b|fwhF2', 'COI', 'fwhF2', 'fwhR2n'),
    (r'BF3\s*/\s*BR2', 'COI', 'BF3', 'BR2'),
    (r'Folmer', 'COI', 'LCO1490', 'HCO2198'),
    (r'515F[- ]?Y?\s*/\s*926R|Parada et al', '16S', '515F-Y', '926R'),
    (r'515F\s*/\s*806R', '16S', '515F (Caporaso)', '806R'),
    (r'341F\s*/\s*785R|Klindworth', '16S', '341F', '785R'),
    (r'1391f?\s*/\s*EukBr|Amaral-Zettler', '18S', '1391f', 'EukBr'),
    (r'TAReuk|Stoeck et al', '18S', 'TAReuk454FWD1', 'TAReukREV3'),
    (r'Diat_rbcL|Vasselon', 'rbcL', 'Diat_rbcL_708F_1', 'R3_1'),
    (r'trnL\s*\(?g\s*/\s*h|trnL g-h', 'trnL', 'trnL-g', 'trnL-h'),
    (r'ITS1F\s*/\s*ITS2', 'ITS', 'ITS1F', 'ITS2'),
    (r'fITS7\s*/\s*ITS4', 'ITS', 'fITS7', 'ITS4'),
]
REFERENCE_DBS = [  # (canonical name, pattern)
    ('SILVA', r'\bSILVA\b'), ('PR2', r'\bPR2\b|Protist Ribosomal Reference'), ('MIDORI2', r'MIDORI'),
    ('BOLD', r'\bBOLD\b|Barcode of Life Data'), ('NCBI GenBank', r'\bNCBI\b|GenBank|\bnt database'),
    ('UNITE', r'\bUNITE\b'), ('Diat.barcode', r'diat\.?barcode|R-Syst'), ('MitoFish', r'MitoFish'),
    ('RDP', r'\bRDP\b'), ('Greengenes', r'Greengenes'), ('GTDB', r'\bGTDB\b'), ('CRUX', r'\bCRUX\b'),
    ('MARES', r'\bMARES\b'), ('PhytoREF', r'PhytoREF'), ('EMBL', r'\bEMBL\b'), ('custom reference database', r'custom|in-house|curated local'),
]
IUPAC_SET = set('ACGTRYKMSWBDHVNI')


def _seq_key(s):
    return re.sub(r'[^ACGTRYKMSWBDHVN]', '', str(s or '').upper().replace('I', 'N'))


def _name_key(s):
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def primer_by_seq(seq):
    k = _seq_key(seq)
    if len(k) < 14: return None
    for e in PRIMER_LIBRARY:
        lk = _seq_key(e[3])
        if k == lk or (len(k) >= 15 and len(lk) >= 15 and (k.endswith(lk) or lk.endswith(k))): return e
    return None


def primer_by_name(name):
    k = _name_key(name)
    if len(k) < 3: return None          # single letters ('g', 'h') only match by sequence
    for e in PRIMER_LIBRARY:
        if any(_name_key(n) == k for n in e[0]): return e
    return None


def _flag(meta, item_id, level, reason, origin=None, evidence=None):
    a = meta.setdefault('_auto', {}).setdefault(item_id, {'level': 'ok', 'reasons': []})
    if level == 'check': a['level'] = 'check'
    if reason and reason not in a['reasons']: a['reasons'].append(reason)
    if origin: a['origin'] = origin
    if evidence: a['evidence'] = evidence


def apply_library(meta: dict) -> dict:
    """Cross-check primers against the library; fill names, references and (flagged) sequences."""
    text = meta.get('_primer_text') or ''
    for i, a in enumerate(meta.get('assays') or []):
        mk = a.get('marker')
        if text and not any(a.get(k) for k in ('primer_fwd', 'primer_rev', 'primer_name_fwd', 'primer_name_rev')):
            for pat, pmk, fname, rname in PRIMER_PAIRS:
                hit = re.search(pat, text, re.I)
                if pmk == mk and hit:
                    s0 = max(text.rfind('.', 0, hit.start()) + 1, 0); s1 = text.find('.', hit.end())
                    src = text[s0: s1 if s1 > 0 else hit.end() + 200].strip()
                    a['primer_name_fwd'], a['primer_name_rev'] = fname, rname
                    for nid in (f'assay:{i}:pcr_primer_name_forward', f'assay:{i}:pcr_primer_name_reverse'):
                        _flag(meta, nid, 'check', 'the report names the primer set, not each primer – names taken from the primer library',
                              origin='primer library', evidence=src)
                    break
        for d, word in (('fwd', 'forward'), ('rev', 'reverse')):
            sid, nid, rid = (f'assay:{i}:pcr_primer_{word}', f'assay:{i}:pcr_primer_name_{word}',
                             f'assay:{i}:pcr_primer_reference_{word}')
            seq, name, ref = a.get(f'primer_{d}'), a.get(f'primer_name_{d}'), a.get(f'primer_ref_{d}')
            e = primer_by_seq(seq) if seq else None
            if e:
                label = f'{e[0][0]} ({e[4] or "published primer"})'
                _flag(meta, sid, 'ok', f'matches the published primer {label}')
                if e[1] != mk: _flag(meta, sid, 'check', f'the primer library lists {e[0][0]} as a {e[1]} primer, not {mk}')
                if e[2] != d: _flag(meta, sid, 'check', f'{e[0][0]} is published as a {"forward" if e[2] == "fwd" else "reverse"} primer – '
                                                         'check the direction')
                if not name and e[2] == d:
                    a[f'primer_name_{d}'] = e[0][0]
                    _flag(meta, nid, 'ok', 'named from the primer library (sequence matches exactly)', origin='primer library')
                elif _name_key(name) in {_name_key(n) for n in e[0]}:
                    _flag(meta, nid, 'ok', f'name and sequence agree with the published {e[0][0]}')
                if not ref and e[4]:
                    a[f'primer_ref_{d}'] = e[4]
                    _flag(meta, rid, 'ok', 'reference taken from the primer library', origin='primer library')
                continue
            e = primer_by_name(name) if name else None
            if not e: continue
            if seq:
                _flag(meta, sid, 'check', f'differs from the published {e[0][0]} sequence ({e[3]}) – check for a typo or a modified primer')
            else:
                a[f'primer_{d}'] = e[3]
                _flag(meta, sid, 'check', f'the report names {name} but does not print the sequence; filled from the primer library',
                      origin='primer library')
                if not ref and e[4]:
                    a[f'primer_ref_{d}'] = e[4]
                    _flag(meta, rid, 'ok', 'reference taken from the primer library', origin='primer library')
            if e[1] != mk: _flag(meta, nid, 'check', f'{e[0][0]} is a {e[1]} primer in the library, but this assay is {mk}')
        db = a.get('otu_db')
        if db:
            known = [n for n, p in REFERENCE_DBS if re.search(p, str(db), re.I)]
            if known: _flag(meta, f'assay:{i}:otu_db', 'ok', 'recognised reference database: ' + ', '.join(known))
    return meta


# ── template-driven field types and validation ───────────────────────────────

_INFO_CACHE: dict = {}


def template_info(path=None) -> dict:
    """{'vocab': {term: (type, options)}, 'types': {term: field type}, 'req': {term: M/HR/R/O}} from a FAIRe template."""
    path = path or (TEMPLATE_PATH if 'TEMPLATE_PATH' in globals() else None)
    try:
        key = (str(path), Path(path).stat().st_mtime)
    except Exception:
        return {'vocab': {}, 'types': {}, 'req': {}}
    if key in _INFO_CACHE: return _INFO_CACHE[key]
    with zipfile.ZipFile(path) as z:
        pkg = {n: z.read(n) for n in z.namelist()}
    ss = shared_strings(pkg); paths = sheet_paths(pkg)
    types, req = {}, {}
    for name in ('projectMetadata', 'sampleMetadata', 'experimentRunMetadata'):
        sp = paths.get(name)
        if not sp: continue
        rows = parse_rows(pkg[sp].decode('utf-8'))
        if name == 'projectMetadata':
            for rn, r in rows.items():
                t = cell_text(r['cells'].get('C', ''), ss)
                if t and rn > 1: req.setdefault(t, cell_text(r['cells'].get('A', ''), ss))
        else:
            hdr = {c: cell_text(x, ss) for c, x in rows.get(3, {}).get('cells', {}).items()}
            for c, x in rows.get(1, {}).get('cells', {}).items():
                if hdr.get(c): req.setdefault(hdr[c], cell_text(x, ss))
        rels = pkg.get(posixpath.join(posixpath.dirname(sp), '_rels', posixpath.basename(sp) + '.rels'), b'').decode('utf-8')
        m = re.search(r'Target="([^"]*comments[^"]*)"', rels)
        cp = posixpath.normpath(posixpath.join(posixpath.dirname(sp), m.group(1))) if m else None
        if not cp or cp not in pkg: continue
        for ref, body in re.findall(r'<comment ref="([A-Z]+\d+)"[^>]*>(.*?)</comment>', pkg[cp].decode('utf-8'), re.S):
            ft = re.search(r'Field type\s*:\s*([^\n<]+)', html.unescape(re.sub(r'<[^>]+>', '', body)))
            col, rn = re.match(r'([A-Z]+)(\d+)', ref).groups()
            term = cell_text(rows.get(int(rn), {}).get('cells', {}).get(col, ''), ss)
            if ft and term: types.setdefault(term, ft.group(1).strip())
    info = {'vocab': load_vocab(pkg, paths, ss), 'types': types, 'req': req}
    _INFO_CACHE.clear(); _INFO_CACHE[key] = info
    return info


ISO_DUR = re.compile(r'P(?!$)(\d+Y)?(\d+M)?(\d+W)?(\d+D)?(T(?=\d)(\d+H)?(\d+M)?(\d+(\.\d+)?S)?)?')
ISO_DATE = r'\d{4}(?:-\d{2}(?:-\d{2}(?:T\d{2}(?::\d{2}(?::\d{2})?)?(?:Z|[+-]\d{2}:?\d{2})?)?)?)?'
UNIT_SUFFIX = r'\s*(?:µ[lL]|u[lL]|[mn][lL]|L|°\s*C|ºC|C|mm|cm|µm|um|nm|%|x|×|bp|cycles?|reads|ng(?:/µ?[lL])?|µM|uM|nM|rpm|g|s|min)?\s*'


def _num(v):
    m = re.fullmatch(r'\s*~?\s*(-?\d+(?:[.,]\d+)?)' + UNIT_SUFFIX, str(v))
    return m.group(1).replace(',', '.') if m else None


def check_value(ftype, value, options=None):
    """-> (value to write, problem or None). Fixes only unambiguous formatting."""
    v = str(value).strip()
    if not v or v in MISSING_TERMS or v.startswith(('missing:', 'not applicable')): return v, None
    ft = (ftype or '').lower()
    if ft in ('numeric', 'integer'):
        if ' | ' in v:  # one value per assay/list
            parts = [check_value(ftype, p)[0] for p in v.split(' | ')]
            return ' | '.join(parts), None if all(_num(p) for p in parts) else f'should be a number'
        n = _num(v)
        if n is None: return v, 'should be a number'
        if ft == 'integer':
            if float(n) != int(float(n)): return v, 'should be a whole number'
            n = str(int(float(n)))
        return n, None
    if ft == 'boolean':
        b = {'1': '1', '0': '0', 'yes': '1', 'true': '1', 'no': '0', 'false': '0'}.get(v.lower())
        return (b, None) if b else (v, 'should be 0 or 1')
    if 'yyyy-mm-dd' in ft:
        if re.fullmatch(ISO_DATE + r'(?:/' + ISO_DATE + r')?', v): return v, None
        m = re.fullmatch(r'(\d{1,2})[/.](\d{1,2})[/.](\d{4})', v)
        if m and int(m.group(1)) > 12:   # day first, unambiguous
            return f'{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}', None
        return v, 'should be a date like 2025-07-14'
    if 'pnynmnwndtnhnmns' in ft:
        if ISO_DUR.fullmatch(v): return v, None
        m = re.fullmatch(r'(\d+)\s*(h|hrs?|hours?|d|days?|min|minutes?|w|weeks?|months?|y|years?)', v, re.I)
        if m:
            n, u = m.group(1), m.group(2).lower()
            code = 'PT{}M' if u.startswith('min') else 'P{}M' if u.startswith('mo') else \
                   {'h': 'PT{}H', 'd': 'P{}D', 'w': 'P{}W', 'y': 'P{}Y'}[u[0]]
            return code.format(n), None
        return v, 'should be an ISO 8601 duration like PT24H'
    if 'termlabel [termid]' in ft:
        return (v, None) if re.fullmatch(r'.+ \[[A-Za-z_]+:[\w.]+\](?:\s*\|\s*.+ \[[A-Za-z_]+:[\w.]+\])*', v) else \
               (v, 'should be an ontology term like "lake [ENVO:00000020]"')
    if ft == 'doi':
        return (v, None) if re.search(r'10\.\d{4,}/|https?://', v) else (v, 'should be a DOI or link')
    if ft.startswith('controlled vocabulary') or options:
        opts = options or [o.strip() for o in re.sub(r'^controlled vocabulary\s*\(|\)$', '', ftype or '').split('|') if o.strip()]
        if v.startswith('other:') or not opts: return v, None
        low = {o.lower(): o for o in opts}
        if v in opts: return v, None
        if v.lower() in low: return low[v.lower()], None
        return v, 'not one of the FAIRe options'
    return v, None


def validate_rows(sheet, rows, ss, info, issues):
    """Check (and safely fix) every value in a filled sheet. Adds to issues {'fixed': {...}, 'problems': {...}}."""
    types, vocab = info.get('types', {}), info.get('vocab', {})
    if sheet == 'projectMetadata':
        cells = []
        for rn, r in rows.items():
            if rn < 2: continue
            term = cell_text(r['cells'].get('C', ''), ss); req = cell_text(r['cells'].get('A', ''), ss)
            vals = [(c, x) for c, x in r['cells'].items() if col_num(c) >= 4]
            if term and req == 'M' and not any(cell_text(x, ss) for _, x in vals):
                issues['problems'].setdefault((sheet, term, 'mandatory but empty'), []).append(f'C{rn}')
            cells += [(term, c, rn) for c, _ in vals]
    else:
        hdr = {c: cell_text(x, ss) for c, x in rows.get(3, {}).get('cells', {}).items()}
        cells = [(hdr.get(c), c, rn) for rn, r in rows.items() if rn > 3 for c in r['cells']]
    for term, col, rn in cells:
        if not term or term in ('checkls_ver',): continue
        x = rows[rn]['cells'][col]
        v = cell_text(x, ss)
        if not v: continue
        fixed, problem = check_value(types.get(term), v, (vocab.get(term) or (None, None))[1])
        if fixed != v and not problem:
            style = re.search(r'\bs="(\d+)"', x)
            rows[rn]['cells'][col] = cell_xml(f'{col}{rn}', fixed, style.group(1) if style else None)
            issues['fixed'].setdefault((sheet, term, f'"{v}" → "{fixed}"'), []).append(f'{col}{rn}')
        elif problem:
            issues['problems'].setdefault((sheet, term, problem), []).append(f'{col}{rn}')


def check_links(sample_rows, run_rows, ss, issues):
    def col_of(rows, name):
        return next((c for c, x in rows.get(3, {}).get('cells', {}).items() if cell_text(x, ss) == name), None)
    sc, rc, lc = col_of(sample_rows, 'samp_name'), col_of(run_rows, 'samp_name'), col_of(run_rows, 'lib_id')
    names = [cell_text(r['cells'].get(sc, ''), ss) for rn, r in sample_rows.items() if rn > 3] if sc else []
    if len(names) != len(set(names)):
        issues['problems'][('sampleMetadata', 'samp_name', 'duplicate sample names')] = sorted({n for n in names if names.count(n) > 1})[:5]
    if rc:
        orphans = sorted({cell_text(r['cells'].get(rc, ''), ss) for rn, r in run_rows.items() if rn > 3} - set(names) - {''})
        if orphans: issues['problems'][('experimentRunMetadata', 'samp_name', 'not found in sampleMetadata')] = orphans[:5]
    if lc:
        libs = [cell_text(r['cells'].get(lc, ''), ss) for rn, r in run_rows.items() if rn > 3]
        if len(libs) != len(set(libs)):
            issues['problems'][('experimentRunMetadata', 'lib_id', 'duplicate library IDs')] = sorted({n for n in libs if libs.count(n) > 1})[:5]


def summarise_issues(issues) -> dict:
    out = {}
    for kind in ('fixed', 'problems'):
        out[kind] = [{'sheet': s, 'term': t, 'what': w, 'cells': len(cells), 'where': ', '.join(map(str, cells[:4])) +
                      ('…' if len(cells) > 4 else '')} for (s, t, w), cells in issues[kind].items()]
    return out


# ── confidence: which values need a human look ───────────────────────────────

RANGES = {'annealingTemp': (35, 75), 'pcr_cycles': (10, 60), 'amplificationReactionVolume': (2, 200), 'pcr_dna_vol': (0.1, 20),
          'pcr_rep': (1, 24), 'size_frac': (0.05, 50), 'filter_diameter': (5, 300), 'otu_clust_cutoff': (80, 100),
          'biological_rep': (1, 100), 'phix_perc': (0, 50), 'pcr_primer_vol_forward': (0.05, 20), 'pcr_primer_vol_reverse': (0.05, 20),
          'pcr_primer_conc_forward': (0.01, 100), 'pcr_primer_conc_reverse': (0.01, 100), 'tax_class_id_cutoff': (50, 100),
          'tax_class_query_cutoff': (20, 100), 'prepped_samp_store_temp': (-200, 40), 'samp_store_temp': (-200, 40),
          'ampliconSize': (30, 2000), 'min_len_cutoff': (20, 1000), 'pcr2_amplificationReactionVolume': (2, 200),
          'pcr2_dna_vol': (0.1, 20), 'samp_size': (0.001, 1e6)}
METHOD_TERMS = set(RANGES) | {'pcr_primer_forward', 'pcr_primer_reverse', 'pcr_primer_name_forward', 'pcr_primer_name_reverse',
                                 'commercial_mm', 'custom_mm', 'pcr_cond', 'seq_kit', 'instrument', 'otu_db', 'trim_method',
                                 'merge_tool', 'otu_clust_tool', 'nucl_acid_ext_kit', 'filter_name', 'filter_material'}
TERM_KEY = dict(EVIDENCE_KEY, **{t: ASSAY_KEY.get(t, t) for t in ASSAY_TERMS})


def _flatten(s):
    return re.sub(r'\s+', ' ', str(s or '').replace('−', '-').replace('μ', 'µ')).strip()


def _tokens(v):
    v = re.sub(r'\b(?:v|ver|version)\.?\s*(?=\d)', '', str(v).lower().replace('µ', 'u'))
    return set(re.findall(r'[a-z0-9]+(?:\.[0-9]+)*', v)) - {'the', 'a', 'an', 'of', 'and', 'kit', 'chemistry', 'reagent', 'database'}


def same_value(a, b) -> bool:
    if str(a).strip().lower() == str(b).strip().lower(): return True
    try: return abs(float(_num(a) or 'x') - float(_num(b) or 'x')) < 1e-6
    except ValueError: pass
    x, y = _tokens(a), _tokens(b)
    return bool(x and y) and (x <= y or y <= x)


def assess(meta: dict, items: list, info: dict | None = None) -> list:
    """Give every review item a status: 'ok' (accepted automatically), 'check' (needs a look) or 'missing'."""
    info = info or {}
    types, vocab, req = info.get('types', {}), info.get('vocab', {}), info.get('req', {})
    auto, conflicts = meta.get('_auto') or {}, meta.get('_conflicts') or {}
    methods = _flatten(meta.get('_methods')) if meta.get('_methods') else None
    for it in items:
        v, term = it['value'], it['term']
        a = auto.get(it['id']) or {}
        it['required'] = req.get(term, '')
        it['origin'] = a.get('origin') or ('tool' if it['generated'] else 'report')
        if not it.get('evidence') and a.get('evidence'): it['evidence'] = a['evidence']
        if not v or v.startswith(('missing:', 'not applicable')):
            it['status'] = 'missing'
            it['reasons'] = a.get('reasons') or ([] if v.startswith('not applicable') else ['not found in the report'])
            if v.startswith('not applicable'): it['status'] = 'ok'; it['reasons'] = ['not applicable']
            continue
        check = list(a.get('reasons', [])) if a.get('level') == 'check' else []
        good = [] if check else list(a.get('reasons', []))
        key = TERM_KEY.get(term, term)
        alt = conflicts.get(f'{it.get("marker")}.{key}' if it['scope'] == 'assay' else key)
        if alt and not a.get('origin') == 'profile':
            check.append('the report also gives: ' + '; '.join(dict.fromkeys(alt))[:160])
        if it['generated']:
            good.append('set by the tool')
        elif a.get('origin'):               # from the user, a profile, the primer library or another source
            pass
        elif not it.get('evidence'):
            if a.get('origin') != 'user': check.append('no source sentence – the value was inferred')
        elif methods and term in METHOD_TERMS and _flatten(it['evidence'])[:70] not in methods:
            check.append('the source sentence is outside the Methods section')
        else:
            good.append('found in the Methods' if methods else 'source sentence found')
        if term in vocab:
            written = str(normalise(term, v, vocab))
            if written.startswith('other:') and not v.startswith('other:'):
                good.append(f'not in the FAIRe list – written as "{written[:60]}"')
            elif written != v:
                good.append(f'written as the FAIRe option "{written}"')
        elif term not in ('checkls_ver',):
            fixed, problem = check_value(types.get(term), v)
            if problem: check.append(problem)
            elif fixed != v: good.append(f'will be written as "{fixed}"')
        n = _num(v)
        if term in RANGES and n is not None and not (RANGES[term][0] <= float(n) <= RANGES[term][1]):
            check.append(f'unusual value for {term} (expected {RANGES[term][0]:g}–{RANGES[term][1]:g})')
        if term in ('pcr_primer_forward', 'pcr_primer_reverse') and a.get('origin') != 'primer library':
            s = re.sub(r'\s', '', v.upper())
            if set(s) - IUPAC_SET or not 14 <= len(s) <= 45: check.append('does not look like a single primer sequence')
        it['status'] = 'check' if check else 'ok'
        it['reasons'] = check or good
    return items


def review_summary(items) -> dict:
    c = {'ok': 0, 'check': 0, 'missing': 0}
    for it in items:
        if it['generated'] and it.get('status') == 'ok': continue
        c[it.get('status', 'ok')] = c.get(it.get('status', 'ok'), 0) + 1
    c['missing_required'] = sum(1 for it in items if it.get('status') == 'missing' and it.get('required') in ('M', 'HR'))
    return c


# ── profiles: learn from corrections ─────────────────────────────────────────

def _portable_id(item_id, meta):
    p = item_id.split(':')
    if p[0] == 'assay':
        a = (meta.get('assays') or [])[int(p[1])]
        return f'assay:{a.get("marker")}:{p[2]}'
    return item_id


def _local_id(pid, meta):
    p = pid.split(':')
    if p[0] != 'assay': return pid
    for i, a in enumerate(meta.get('assays') or []):
        if a.get('marker') == p[1]: return f'assay:{i}:{p[2]}'
    return None


def profile_from_review(meta: dict, edits: dict, name: str = 'my profile', profile: dict | None = None) -> dict:
    """Turn review edits into a reusable profile: {id: {'from': value the tool found, 'to': your value}}."""
    profile = dict(profile or {'faire_profile': 1, 'name': name, 'entries': {}})
    profile['name'] = name or profile.get('name') or 'my profile'
    profile['updated'] = time.strftime('%Y-%m-%d')
    entries = dict(profile.get('entries') or {})
    found = {it['id']: it['value'] for it in review_items(dict(meta, _overrides={}))}
    for item_id, value in (edits or {}).items():
        if item_id not in found: continue
        entries[_portable_id(item_id, meta)] = {'from': found[item_id], 'to': str(value or '').strip()}
    profile['entries'] = entries
    return profile


def apply_profile(meta: dict, profile: dict) -> dict:
    """Apply a saved profile: a correction is used when the tool finds the same value as last time (or nothing)."""
    if not profile: return meta
    name = profile.get('name') or 'profile'
    current = {it['id']: it['value'] for it in review_items(meta)}
    edits = {}
    for pid, e in (profile.get('entries') or {}).items():
        lid = _local_id(pid, meta)
        if not lid or lid not in current: continue
        cur, frm, to = current[lid], e.get('from') or '', e.get('to') or ''
        empty = not cur or cur.startswith(('missing:', 'not applicable'))
        if same_value(cur, to) or (cur == to):
            _flag(meta, lid, 'ok', f'matches your profile "{name}"', origin='profile')
        elif empty or cur == frm:
            edits[lid] = to
            _flag(meta, lid, 'ok', f'from your profile "{name}"', origin='profile')
        else:
            _flag(meta, lid, 'check', f'your profile "{name}" says "{to[:80]}"')
    apply_review(meta, edits)
    meta.setdefault('notes', [])
    return meta


# ── second opinion from another reader (local AI or Claude) ──────────────────

def second_opinion(meta: dict, other: dict, label: str) -> dict:
    """Compare with another reader's (grounded) result: agreement -> accepted, disagreement -> check, gaps -> filled + check."""
    mine = {it['id']: it for it in review_items(meta)}
    m_of = {a.get('marker'): i for i, a in enumerate(meta.get('assays') or [])}
    edits, agree, differ, filled = {}, 0, 0, 0
    for it in review_items(other):
        if it['generated']: continue
        lid = it['id']
        if it['scope'] == 'assay':
            mk = (other.get('assays') or [])[int(lid.split(':')[1])].get('marker')
            if mk not in m_of: continue
            lid = f'assay:{m_of[mk]}:{it["term"]}'
        if lid not in mine: continue
        theirs, ours = it['value'], mine[lid]['value']
        if not theirs or theirs.startswith(('missing:', 'not applicable')): continue
        if not ours or ours.startswith('missing:'):
            edits[lid] = theirs; filled += 1
            _flag(meta, lid, 'check', f'not found by the built-in reader; found by {label} – please confirm', origin=label,
                  evidence=it.get('evidence'))
        elif same_value(ours, theirs):
            agree += 1; _flag(meta, lid, 'ok', f'{label} agrees')
        else:
            differ += 1; _flag(meta, lid, 'check', f'{label} read "{theirs[:80]}"')
    apply_review(meta, edits)
    for k in ('sites', 'rounds'):
        if not meta.get(k) and other.get(k): meta[k] = other[k]
    meta.setdefault('notes', []).append(f'{label}: agreed on {agree} value(s), disagreed on {differ}, filled {filled} gap(s) (marked for checking).')
    meta['_method'] = f'{meta.get("_method", "rule-based")} + {label}'
    return meta


def review(meta: dict, template=None) -> list:
    """review_items + a status for each: what can be accepted automatically and what needs a look."""
    return assess(meta, review_items(meta), template_info(template))


# ════════════════════════════════════════════════════════════════════════════
#  5. AI readers (optional): local model via Ollama, or Claude for non-confidential reports (same schema as the rule-based extractor)
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
runs: [{seq_run_id, date, assays: [markers sequenced on this run]}],
_evidence: {key: "exact quote from the report supporting that value"}  — one entry for EVERY top-level value you give,
           copied word for word (max ~30 words); for assay fields use the key "<marker>.<field>", e.g. "COI.primer_name_fwd".
Values without an exact supporting quote will be discarded.

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


LOCAL_MODEL = os.environ.get('FAIRE_LOCAL_MODEL', 'qwen2.5:7b')
OLLAMA_HOST = os.environ.get('OLLAMA_HOST', 'http://127.0.0.1:11434')
LOCAL_MAX_CHARS = 40_000


def parse_with_ollama(text: str, model: str = None, host: str = None, timeout: int = 900) -> dict:
    """Ask a model running locally in Ollama (nothing leaves this computer). Reads the Methods section when there is one."""
    import urllib.request
    host = (host or OLLAMA_HOST).rstrip('/')
    if not host.startswith('http'): host = 'http://' + host
    if not re.match(r'https?://(127\.0\.0\.1|localhost|\[::1\])(:\d+)?$', host):
        raise RuntimeError(f'Refusing to send the report to {host}: the local AI reader only talks to this computer.')
    body = methods_text(strip_line_numbers(text)) or text
    payload = {'model': model or LOCAL_MODEL, 'stream': False, 'format': 'json',
               'options': {'temperature': 0, 'num_ctx': 16384},
               'messages': [{'role': 'user', 'content': CLAUDE_PROMPT + body[:LOCAL_MAX_CHARS]}]}
    req = urllib.request.Request(host + '/api/chat', data=json.dumps(payload).encode('utf-8'),
                                 headers={'Content-Type': 'application/json'})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            reply = json.loads(r.read().decode('utf-8'))
    except OSError as e:
        raise RuntimeError(f'Could not reach Ollama at {host} ({e}). Is it running? Install from https://ollama.com, '
                           f'then: ollama pull {model or LOCAL_MODEL}')
    return json_from_reply(reply.get('message', {}).get('content', ''))


def local_ai_opinion(meta: dict, text: str, model: str = None, host: str = None) -> dict:
    """Second opinion from a local model: grounded like Claude, then compared value by value."""
    other = facts_from_claude(parse_with_ollama(text, model, host), text)
    return second_opinion(meta, other, f'local AI ({model or LOCAL_MODEL})')


def json_from_reply(raw: str) -> dict:
    """Parse the JSON object in a Claude reply (tolerates code fences / stray prose)."""
    raw = re.sub(r'^```(?:json)?\s*|\s*```$', '', raw.strip())
    return json.loads(raw[raw.find('{'): raw.rfind('}') + 1])


def claude_prompt(text: str) -> str:
    return CLAUDE_PROMPT + text[:MAX_REPORT_CHARS]


def _quote_found(quote: str, flat_text: str) -> bool:
    q = re.sub(r'\s+', ' ', str(quote or '')).strip().lower()
    if len(q) < 8: return False
    if q in flat_text: return True
    words = q.split()                                   # tolerate small differences (hyphenation, ellipses)
    chunks = [' '.join(words[i:i + 6]) for i in range(0, max(1, len(words) - 5), 3)]
    return sum(c in flat_text for c in chunks) >= max(1, len(chunks) // 2)


GROUNDING_EXEMPT = {'notes', 'assays', 'sites', 'rounds', 'runs', '_evidence', 'language'}


def facts_from_claude(data: dict, text: str = None) -> dict:
    """Tidy Claude's JSON into the same shape the rule-based extractor returns.
    If the report text is given, every value must be backed by a quote that really appears in it; others are dropped."""
    data = dict(data or {})
    evidence = {k: v for k, v in (data.pop('_evidence', None) or {}).items() if v}
    dropped = []
    if text is not None:
        flat_text = re.sub(r'\s+', ' ', text).lower()
        for k in list(data):
            if k in GROUNDING_EXEMPT or data[k] in (None, '', []): continue
            if not _quote_found(evidence.get(k), flat_text):
                dropped.append(k); data.pop(k)
        for a in data.get('assays') or []:
            mk = canon_marker(str(a.get('marker') or a.get('assay_name') or ''))
            for k in list(a):
                if k in ('marker', 'assay_name') or a[k] in (None, ''): continue
                if not _quote_found(evidence.get(f'{mk}.{k}'), flat_text):
                    dropped.append(f'{mk}.{k}'); a.pop(k)
        evidence = {k: v for k, v in evidence.items() if _quote_found(v, flat_text)}
    meta = {k: v for k, v in data.items() if v not in (None, '', [])}
    meta['_evidence'] = evidence
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
    if dropped:
        meta['notes'].append(f'{len(dropped)} value(s) from Claude were discarded because their quoted source could not be found in the report: '
                             + ', '.join(dropped[:12]) + ('…' if len(dropped) > 12 else ''))
    meta['_method'] = f'Claude ({CLAUDE_MODEL})'
    return meta


def parse_report_regex(text: str) -> dict:
    """Rule-based extraction (kept under its old name for the Colab notebook)."""
    meta = extract(text)
    meta['_method'] = 'rule-based'
    code, name = detect_language(text)
    meta['language'] = name
    meta['notes'].extend(PDF_NOTES)
    if code != 'en':
        meta['notes'].insert(0, f'This report appears to be in {name}. The built-in patterns read English methods text, so most fields '
                                f'will be empty – please fill them in during the review step (or use an English version of the report).')
    return meta


# ════════════════════════════════════════════════════════════════════════════
#  6. Extra sources: protocols, dataset records, code repositories, data files, sequence archives
# ════════════════════════════════════════════════════════════════════════════

# ── fetching (only when the user gives a public link or accession) ───────────

FETCH_OVERRIDES: dict = {}          # url -> bytes, used by tests (no network)
FETCH_LOG: list = []
MAX_DOWNLOAD = 60 * 1024 * 1024


def fetch(url: str, max_bytes: int = MAX_DOWNLOAD) -> bytes:
    if url in FETCH_OVERRIDES:
        FETCH_LOG.append(url)
        if isinstance(FETCH_OVERRIDES[url], Exception): raise RuntimeError(f'could not download {url} ({FETCH_OVERRIDES[url]})')
        return FETCH_OVERRIDES[url]
    url_q = quote_url(url)
    if __import__('sys').platform == 'emscripten':          # in the browser page: synchronous request through the browser
        return _browser_fetch(url, url_q, max_bytes)
    import urllib.request
    data, last = None, None
    for ua in ('FAIRe-Metadata-Filler (+https://github.com/UKCEH-MolecularEcology/fair-edna-filler)',
               'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36'):
        try:
            req = urllib.request.Request(url_q, headers={'User-Agent': ua, 'Accept': '*/*'})
            with urllib.request.urlopen(req, timeout=60) as r:
                size = int(r.headers.get('Content-Length') or 0)
                if size and size > max_bytes: raise RuntimeError(f'{url} is {size / 1e6:.0f} MB – skipped (too large)')
                data = r.read(max_bytes + 1)
            break
        except RuntimeError:
            raise
        except Exception as e:          # HTTP errors, refused connections, odd URLs
            last = e
    if data is None:
        raise RuntimeError(f'could not download {url} ({last})')
    if len(data) > max_bytes: raise RuntimeError(f'{url} is larger than {max_bytes // 1_000_000} MB – skipped')
    FETCH_LOG.append(url)
    return data


def _browser_fetch(url, url_q, max_bytes):
    try:
        import js  # type: ignore
        xhr = js.XMLHttpRequest.new()
        xhr.open('GET', url_q, False)
        xhr.overrideMimeType('text/plain; charset=x-user-defined')     # raw bytes
        xhr.send(None)
    except Exception as e:
        raise RuntimeError(f'could not download {url} (the site does not allow downloads from a web page – '
                           f'use the Python or Colab version, or download the file and add it yourself)')
    if xhr.status != 200:
        raise RuntimeError(f'could not download {url} (HTTP {xhr.status})')
    text = str(xhr.responseText)
    if len(text) > max_bytes: raise RuntimeError(f'{url} is too large – skipped')
    FETCH_LOG.append(url)
    return bytes(ord(c) & 0xFF for c in text)


def quote_url(url: str) -> str:
    """Make a link safe to request: spaces and other odd characters in paths/queries are percent-encoded, line breaks removed."""
    import urllib.parse
    url = re.sub(r'[\x00-\x1f\x7f]', '', str(url).strip())
    parts = urllib.parse.urlsplit(url)
    return urllib.parse.urlunsplit(parts._replace(
        path=urllib.parse.quote(urllib.parse.unquote(parts.path), safe="/:@!$&'()*+,;=-._~"),
        query=urllib.parse.quote(urllib.parse.unquote(parts.query), safe='=&?/:,;+'),
        fragment=''))


def fetch_json(url):
    return json.loads(fetch(url).decode('utf-8'))


# ── recognising what the user (or the report) points to ─────────────────────

def _squash(s):
    return re.sub(r'\s+', '', s)


def find_links(text: str) -> list:
    """Public identifiers mentioned in a report: protocols.io / Zenodo / Dryad / Figshare DOIs, GitHub repos, BioProjects."""
    t = text.replace(' ', ' ')
    found = []
    # URLs and DOIs in PDFs are often broken by spaces: "https:// doi. org/ 10. 5281/ zenodo. 15723454"
    for m in re.finditer(r'(?:https?:\s*/\s*/\s*)?(?:dx\.\s*)?(?:doi\.\s*org\s*/\s*)?10\.\s*\d{4,5}\s*/[^\s;,)]*(?:\s+[^\s;,)]+){0,10}', t):
        cand = _squash(m.group(0))
        d = re.search(r'10\.\d{4,5}/[\w.\-/]+', cand)
        if d:
            doi = d.group(0).rstrip('.')
            doi = re.match(r'(10\.17504/protocols\.io\.[a-z0-9]+(?:/v\d+)?|10\.5281/zenodo\.\d+|10\.5061/dryad\.[\w]+|10\.6084/m9\.figshare\.\d+(?:\.v\d+)?|.*)', doi).group(1)
            if re.match(r'10\.(17504|5281|5061|6084)/', doi): found.append(doi)
    for m in re.finditer(r'github\.\s*com\s*/\s*([^\s)]+(?:\s+[^\s)]+){0,16})', t):
        parts = re.match(r'([A-Za-z0-9\-_.]+)/([A-Za-z0-9\-_.]+)', _squash(m.group(1)))
        if parts: found.append(f'https://github.com/{parts.group(1)}/{parts.group(2).rstrip(".")}')
    found += re.findall(r'\b(PRJ(?:NA|EB|DB)\d+|[SED]RP\d{5,})\b', t)
    out = []
    for f in found:
        if f not in out: out.append(f)
    return out


def classify(src: str) -> str:
    s = str(src).strip()
    p = Path(s)
    if p.exists():
        if re.match(r'(extra_)?readme', p.name, re.I): return 'file:code'
        if p.suffix.lower() in ('.csv', '.tsv', '.txt', '.xlsx') and not re.search(r'protocol', p.name, re.I): return 'file:data'
        if p.suffix.lower() in ('.r', '.py', '.sh', '.smk', '.nf', '.yaml', '.yml'): return 'file:code'
        if p.suffix.lower() == '.zip': return 'file:zip'
        return 'file:protocol'
    if re.search(r'protocols\.io', s, re.I): return 'protocol'
    if re.search(r'github\.com/', s, re.I): return 'github'
    if re.search(r'zenodo', s, re.I): return 'zenodo'
    if re.fullmatch(r'PRJ(?:NA|EB|DB)\d+|[SED]RP\d+', s): return 'sra'
    if re.search(r'10\.\d{4,5}/', s): return 'doi'
    return 'unknown'


# ── protocols ────────────────────────────────────────────────────────────────

def protocol_to_prose(text: str) -> str:
    """Step-list protocols (protocols.io exports) -> sentences the extractor can read."""
    lines = [l.strip() for l in text.replace('ʼ', "'").split('\n')]
    out, cur = [], ''
    for l in lines:
        if not l: continue
        if re.match(r'protocols\.io \||^\d+/\d+$|^Create & collaborate|^Create free account', l): continue
        new = re.match(r'(\d+(\.\d+)*\s+[A-Z]|\d+\.\s|[-•▪]\s|Section \d+|Step \d+)', l) or \
              re.match(r'\d+(\.\d+)?\s*(µ|u)L\b|\d+\s*°C|\d+\s*cycles', l)
        if new and cur:
            out.append(cur.rstrip(' .:') + '.'); cur = l
        else:
            cur = (cur + ' ' + l).strip()
    if cur: out.append(cur.rstrip(' .:') + '.')
    return '\n'.join(out)


def _pcr_block(text, start_pat, end_pat):
    m = re.search(start_pat, text, re.I)
    if not m: return None
    e = re.search(end_pat, text[m.end():], re.I)
    return text[m.start(): m.end() + (e.start() if e else 4000)]


def _cycling(block):
    temps = re.findall(r'(\d{2,3})\s*°C\s*(?:for\s*)?\(?\s*(\d+(?:\.\d+)?)\s*(s|sec\w*|min\w*|m)\b', block)
    cyc = re.search(r'(\d{1,2})\s*cycles', block)
    if not temps or not cyc: return None
    unit = lambda u: 'min' if u.startswith('m') else 's'
    pre = [t for t in re.findall(r'(\d{2,3})\s*°C[^0-9]{0,12}?(\d+(?:\.\d+)?)\s*(s|sec\w*|min\w*)', block[:cyc.start()])]
    inside = re.findall(r'(\d{2,3})\s*°C[^0-9]{0,12}?(\d+(?:\.\d+)?)\s*(s|sec\w*|min\w*)', block[cyc.end(): cyc.end() + 160])
    after = re.findall(r'(\d{2,3})\s*°C[^0-9]{0,12}?(\d+(?:\.\d+)?)\s*(s|sec\w*|min\w*)', block[cyc.end() + 160:])
    parts = []
    if pre: parts.append(f'initial denaturation:{pre[0][0]}_{pre[0][1]}{unit(pre[0][2])}')
    names = ['denaturation', 'annealing', 'elongation'] if len(inside) >= 3 else ['denaturation', 'annealing/elongation']
    for n, (tp, d, u) in zip(names, inside):
        parts.append(f'{n}:{tp}_{d}{unit(u)}')
    if after: parts.append(f'final elongation:{after[0][0]}_{after[0][1]}{unit(after[0][2])}')
    parts.append(f'{cyc.group(1)} cycles')
    ann = inside[1][0] if len(inside) >= 2 else None
    return '; '.join(parts), ann, cyc.group(1)


def _recipe(block):
    comps = re.findall(r'(\d+(?:\.\d+)?)\s*(?:µ|u)L\s+(?:of\s+)?([^\n.;]{3,80})', block)
    if not comps: return {}
    total, mm, custom, out = 0.0, None, [], {}
    stated = re.search(r'(\d+(?:\.\d+)?)\s*(?:µ|u)L\s+(?:PCR\s*\d?\s+)?reactions?', block, re.I)
    for v, what in comps:
        if re.match(r'(PCR\s*\d?\s+)?reactions?\b', what.strip(), re.I): continue
        v = float(v); each = re.search(r'\beach\b', what, re.I)
        total += v * (2 if each and re.search(r'primer', what, re.I) else 1)
        if re.search(r'master ?mix|ready ?mix|\bmix\b|polymerase', what, re.I) and not mm: mm = norm(what)
        elif re.search(r'template|DNA', what, re.I): out['dna'] = f'{v:g}'
        elif re.search(r'primer', what, re.I):
            out['primer_vol'] = f'{v:g}'
            c = re.search(r'(\d+(?:\.\d+)?)\s*(?:µ|u)M', what)
            if c: out['primer_conc'] = c.group(1)
        else: custom.append(f'{v:g} µL {norm(what)}')
    out.update(total=stated.group(1) if stated else f'{total:g}', mm=mm, custom='; '.join(custom) or None)
    return out


def read_protocol_text(text: str, label: str, doi: str | None) -> dict:
    """Rule-based reading of a protocol, plus list-style PCR recipes and cycling."""
    prose = protocol_to_prose(text)
    meta = parse_report_regex(prose)
    meta['_primer_text'] = norm(prose)
    if not meta.get('assays'):                       # protocols rarely repeat the marker name: use printed primer sequences
        seqs = re.findall(r'[ACGTRYKMSWBDHVNI]{15,}', prose)
        for e in PRIMER_LIBRARY:
            if any(_seq_key(x) == _seq_key(e[3]) for x in seqs) and not any(a['marker'] == e[1] for a in meta['assays']):
                meta['assays'].append({'marker': e[1], 'assay_name': e[1], 'target_gene': TARGET_GENE.get(e[1], e[1])})
        for a in meta['assays']:
            for x in seqs:
                hit = primer_by_seq(x)
                if hit and hit[1] == a['marker']:
                    a.setdefault(f'primer_{hit[2]}', x); a.setdefault(f'primer_name_{hit[2]}', hit[0][0])
                    meta['_evidence'][f'{a["marker"]}.primer_{hit[2]}'] = norm(prose[max(0, prose.find(x) - 120): prose.find(x) + 40])
    apply_library(meta)
    flat = norm(prose)
    blocks = {'pcr1': _pcr_block(prose, r'first PCR|PCR\s*1\b|Step 1\b.*PCR', r'PCR\s*2\b|second PCR|indexing|Step [2-9]'),
              'pcr2': _pcr_block(prose, r'PCR\s*2\b|second PCR|indexing PCR', r'Step [4-9]|Library Quantification|Sequencing|Section \d')}
    ev = meta.setdefault('_evidence', {})
    for key, block in blocks.items():
        if not block: continue
        rc, cy = _recipe(block), _cycling(block)
        if key == 'pcr1':
            for a in meta['assays']:
                if rc.get('total'): a['amplificationReactionVolume'] = rc['total']
                if rc.get('dna'): a['pcr_dna_vol'] = rc['dna']
                if rc.get('mm'): a['commercial_mm'] = rc['mm']
                if rc.get('custom'): a['custom_mm'] = rc['custom']
                if rc.get('primer_vol'): a['pcr_primer_vol'] = rc['primer_vol']
                if rc.get('primer_conc'): a['pcr_primer_conc'] = rc['primer_conc']
                if cy: a['pcr_cond'], a['annealingTemp'], a['pcr_cycles'] = cy
                rep = re.search(r'in (duplicate|triplicate|quadruplicate)', block, re.I)
                if rep: a['pcr_rep'] = {'duplicate': '2', 'triplicate': '3', 'quadruplicate': '4'}[rep.group(1).lower()]
                for k in ('amplificationReactionVolume', 'pcr_dna_vol', 'commercial_mm', 'custom_mm', 'pcr_primer_vol', 'pcr_primer_conc',
                          'pcr_cond', 'annealingTemp', 'pcr_cycles', 'pcr_rep'):
                    ev[f'{a["marker"]}.{k}'] = norm(block)[:400]
        else:
            if rc.get('total'): meta['pcr2_amplificationReactionVolume'] = rc['total']
            if rc.get('dna'): meta['pcr2_dna_vol'] = rc['dna']
            if rc.get('mm'): meta['pcr2_commercial_mm'] = rc['mm']
            if rc.get('custom'): meta['pcr2_custom_mm'] = rc['custom']
            if cy: meta['pcr2_cond'], meta['pcr2_annealingTemp'], meta['pcr2_cycles'] = cy
            for k in ('pcr2_amplificationReactionVolume', 'pcr2_dna_vol', 'pcr2_commercial_mm', 'pcr2_custom_mm', 'pcr2_cond',
                      'pcr2_annealingTemp', 'pcr2_cycles'):
                ev[k] = norm(block)[:400]
    rules = [  # (key, pattern, group or callable)
        ('thermocycler', r'((?:Applied Biosystems|Bio-?Rad|Eppendorf|Thermo Fisher|Analytik Jena)[\w ®™-]{0,40}?(?:Thermal ?Cycler|Thermocycler|PCR System|C1000[\w ]*|T100[\w ]*))', 1),
        ('trim_param', r'(\d+\s*bp sliding window[^.]{0,80})', 1),
        ('min_len_cutoff', r'(?:retain|keep)\w* reads (?:between|of)\s*(\d+)\s*[–-]', 1),
        ('merge_min_overlap', r'minimum\s*(\d+)\s*bp\s*overlap', 1),
        ('chimera_check_method', r'(--uchime3?_denovo|uchime3?_denovo|removeBimeraDenovo|--uchime_ref)', 1),
        ('demux_tool', r'(demultiplex[^.]{0,200})', 1),
        ('tax_class_collapse', r'((?:Majority )?Lowest Common Ancestor.{0,400}?agree[^.]{0,60})', 1),
        ('tax_class_other', r'([^.]{0,80}cannot be assigned below [^.]{0,80})', 1),
        ('error_rate_cutoff', r'Phred score of Q(\d+)', 1),
    ]
    for key, pat, g in rules:          # protocol wording is specific: these replace the general reader's guesses
        m = re.search(pat, flat, re.I | re.S)
        if m: meta[key] = norm(m.group(g))[:400]; ev[key] = norm(m.group(0))[:400]
    cc = _pcr_block(prose, r'Contamination control|thresholds?\b|limit of detection', r'Section \d|Protocol references|References')
    if cc: meta['screen_contam_method'] = norm(re.sub(r'^\d+\s+', '', cc))[:500]; ev['screen_contam_method'] = norm(cc)[:400]
    if meta.get('error_rate_cutoff') and meta.get('trim_method'):
        meta.setdefault('error_rate_tool', meta['trim_method']); ev.setdefault('error_rate_tool', ev.get('error_rate_cutoff'))
    if meta.get('min_len_cutoff') and meta.get('trim_method'):
        meta.setdefault('min_len_tool', meta['trim_method']); ev.setdefault('min_len_tool', ev.get('min_len_cutoff'))
    pm = [norm(x) for x in prose.split('\n') if re.search(r'MID tag|multiplex identifier|heterogeneity spacer|inhibit|dilut\w* 1:\d+|nested', x, re.I)]
    if pm: meta['pcr_method_additional'] = ' '.join(pm[:4])[:600]; ev['pcr_method_additional'] = meta['pcr_method_additional']
    if meta.get('thermocycler') and not meta.get('pcr2_thermocycler') and meta.get('pcr2_cond'):
        meta['pcr2_thermocycler'] = meta['thermocycler']; ev['pcr2_thermocycler'] = ev.get('thermocycler')
    for k in ASSAY_TERMS:                # assay-level terms read from the protocol apply to every assay
        if meta.get(k):
            for a in meta['assays']:
                if not a.get(ASSAY_KEY.get(k, k)):
                    a[ASSAY_KEY.get(k, k)] = meta[k]; ev[f'{a["marker"]}.{ASSAY_KEY.get(k, k)}'] = ev.get(k)
    lq = _pcr_block(prose, r'Library Quantification|Library pooling|Quantif\w+ (?:and|&) Pooling', r'Section \d|Step [5-9]|Sequencing')
    if lq:
        bits = [norm(x) for x in re.split(r'\n', lq) if re.search(r'Qubit|qPCR|Bioanalyzer|TapeStation|\bnM\b|Quant', x)]
        if bits: meta['lib_screen'] = '; '.join(re.sub(r'^\d+\.\s*', '', b).rstrip('.') for b in bits); ev['lib_screen'] = norm(lq)[:400]
    st = [norm(x) for x in prose.split('\n') if re.search(r'wip\w*|soak\w*|rins\w*', x, re.I) and re.search(r'bleach|ethanol|lipsol|Virkon|DNA[- ]?away|UV', x, re.I)]
    if st: meta['sterilise_method'] = ' '.join(st[:3])[:600]; ev['sterilise_method'] = meta['sterilise_method']
    if meta.get('error_rate_cutoff'): meta['error_rate_type'] = 'Phred score'; ev['error_rate_type'] = ev['error_rate_cutoff']
    cl = re.search(r'--cluster_unoise|unoise3?\b|--cluster_size|--cluster_fast|swarm', flat, re.I)
    if cl:
        meta['otu_clust_tool'] = ('VSEARCH ' if re.search(r'vsearch', flat, re.I) else '') + cl.group(0)
        ev['otu_clust_tool'] = norm(flat[max(0, cl.start() - 100): cl.end() + 60])
    mr = re.search(r'fewer than\s+(\w+)\s+reads', flat, re.I)
    if mr and num_word(mr.group(1)):
        meta['min_reads_cutoff'], meta['min_reads_cutoff_unit'] = str(num_word(mr.group(1))), 'reads'; ev['min_reads_cutoff'] = norm(mr.group(0))
    if meta.get('min_reads_cutoff') and re.search(r'vsearch', flat, re.I):
        meta.setdefault('min_reads_tool', 'VSEARCH'); ev.setdefault('min_reads_tool', ev.get('min_reads_cutoff'))
    if doi:
        url = 'https://doi.org/' + doi if not doi.startswith('http') else doi
        for k in ('nucl_acid_amp', 'nucl_acid_ext', 'sop_bioinformatics'):
            meta[k] = url if k != 'sop_bioinformatics' or not meta.get(k) else f'{meta[k]}; {url}'
            ev[k] = f'{label}: {url}'
    meta['_label'] = label
    return meta


def protocol_from_url(url: str) -> dict:
    doi = re.search(r'10\.17504/protocols\.io\.[a-z0-9]+(?:/v\d+)?', url)
    if doi:
        slug = doi.group(0)
    view = url if 'protocols.io/view/' in url else None
    if not view and doi:   # resolve the DOI to its protocols.io page (DOI handle API)
        if ('https://doi.org/' + slug) in FETCH_OVERRIDES:
            view = FETCH_OVERRIDES['https://doi.org/' + slug].decode()
        else:
            try:
                h = fetch_json('https://doi.org/api/handles/' + slug)
                view = next((v['data']['value'] for v in h.get('values', []) if v.get('type') == 'URL'), None)
            except Exception:
                view = None
    if not view: raise RuntimeError(f'could not resolve {url} to a protocols.io page')
    pdf = re.sub(r'(\?.*)?$', '', view).rstrip('/')
    pdf = pdf if pdf.endswith('.pdf') else pdf + '.pdf'
    try:
        data = fetch(pdf)
    except RuntimeError as e:
        raise RuntimeError(f'protocols.io did not let the tool download the protocol ({e}). Download it yourself from {view} '
                           '(⋯ menu → Download PDF) and add it with --source protocol.pdf')
    text = _pdf_bytes_text(data) if data[:4] == b'%PDF' else data.decode('utf-8', 'replace')
    d = re.search(r'10\.17504/protocols\.io\.[a-z0-9]+(?:/v\d+)?', text)
    return read_protocol_text(text, 'protocols.io', d.group(0) if d else (doi.group(0) if doi else None))


def _pdf_bytes_text(data: bytes) -> str:
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        p = Path(tmp) / 'doc.pdf'; p.write_bytes(data)
        return extract_report_text(p)


# ── dataset records (DataCite: Zenodo, Dryad, Figshare, …) ──────────────────

def read_datacite(doi: str) -> dict:
    doi = re.search(r'10\.\d{4,5}/[^\s?#]+', doi).group(0)
    d = fetch_json(f'https://api.datacite.org/dois/{doi}?affiliation=true')['data']['attributes']
    creators = d.get('creators') or []
    names, orcids, affs = [], [], {}
    for c in creators:
        nm = ' '.join(x for x in (c.get('givenName'), c.get('familyName')) if x) or c.get('name')
        names.append(nm)
        oid = next((i['nameIdentifier'] for i in c.get('nameIdentifiers') or [] if 'ORCID' in (i.get('nameIdentifierScheme') or '')), None)
        if oid: orcids.append(oid if oid.startswith('http') else f'https://orcid.org/{oid}')
        for a in c.get('affiliation') or []:
            nm_a = a.get('name') if isinstance(a, dict) else a
            ror = a.get('affiliationIdentifier') if isinstance(a, dict) else None
            if nm_a: affs.setdefault(nm_a, [0, ror])[0] += 1
    out = {'_label': f'dataset record {doi}', 'doi': doi}
    if names: out['recorded_by'] = ' | '.join(names)
    if orcids and len(orcids) == len(names): out['recorded_by_id'] = ' | '.join(orcids)
    elif orcids: out['recorded_by_id_partial'] = ' | '.join(orcids)
    if affs:
        top = max(affs.items(), key=lambda kv: kv[1][0])
        out['institution'] = top[0]
        if top[1][1] and 'ror.org' in str(top[1][1]): out['institution_ror'] = top[1][1]
    rights = d.get('rightsList') or []
    if rights:
        r = rights[0]
        lic = r.get('rights') or r.get('rightsIdentifier')
        out['license'] = lic + (f" ({r['rightsUri']})" if r.get('rightsUri') else '')
        out['access_rights'] = 'open access' if re.search(r'creative commons|cc-?by|cc0|open', lic or '', re.I) else None
    out['title'] = (d.get('titles') or [{}])[0].get('title')
    out['related'] = [r.get('relatedIdentifier') for r in d.get('relatedIdentifiers') or []]
    out['funder'] = ' | '.join(f.get('funderName') for f in d.get('fundingReferences') or [] if f.get('funderName')) or None
    out['url'] = d.get('url')
    out['_evidence'] = {k: f'DataCite record {doi}: {k}' for k in out if not k.startswith('_')}
    return out


# ── code repositories and dataset files ──────────────────────────────────────

DATA_EXT = ('.csv', '.tsv', '.txt')
CODE_EXT = ('.r', '.py', '.sh', '.smk', '.nf', '.yaml', '.yml', '.config', '.cfg', '.ini', '.json')


def read_github(url: str) -> dict:
    m = re.search(r'github\.com/([^/\s]+)/([^/\s#?]+)', url)
    owner, repo = m.group(1), m.group(2).removesuffix('.git')
    info = fetch_json(f'https://api.github.com/repos/{owner}/{repo}')
    branch = info.get('default_branch') or 'main'
    tree = fetch_json(f'https://api.github.com/repos/{owner}/{repo}/git/trees/{branch}?recursive=1').get('tree') or []
    files, skipped = {}, []
    for t in tree:
        p = t.get('path', '')
        if t.get('type') != 'blob' or (t.get('size') or 0) > 5_000_000: continue
        if p.lower().endswith(DATA_EXT + CODE_EXT) or re.match(r'readme', Path(p).name, re.I):
            try: files[p] = fetch(f'https://raw.githubusercontent.com/{owner}/{repo}/{branch}/{p}')
            except Exception as e: skipped.append(f'{p} ({e})')
    return {'_label': f'GitHub {owner}/{repo}', 'code_repo': f'https://github.com/{owner}/{repo}', 'files': files,
            'license': (info.get('license') or {}).get('name'), 'skipped': skipped}


def read_zenodo(ref: str) -> dict:
    rid = re.search(r'zenodo\.(\d+)|records?/(\d+)', ref)
    rid = rid.group(1) or rid.group(2)
    out = {'_label': f'Zenodo {rid}', 'files': {}}
    try: out.update({k: v for k, v in read_datacite(f'10.5281/zenodo.{rid}').items() if k not in ('_label',)})
    except Exception as e: out['_error'] = str(e)
    entries = None
    for attempt in (rid, 'latest'):
        try:
            if attempt == 'latest':     # a concept DOI (all versions): ask Zenodo for its latest version
                hits = fetch_json(f'https://zenodo.org/api/records?q=conceptrecid:{rid}&sort=mostrecent&size=1').get('hits', {}).get('hits', [])
                if not hits: break
                rid = str(hits[0].get('id')); out['_label'] = f'Zenodo {rid}'
            entries = fetch_json(f'https://zenodo.org/api/records/{rid}/files').get('entries') or []
            out.pop('_error', None); break
        except Exception as e:
            out['_error'] = str(e)
    entries = entries or []
    out['file_list'] = [(e.get('key'), e.get('size'), e.get('checksum')) for e in entries]
    for e in entries:
        key, size = e.get('key') or '', e.get('size') or 0
        url = (e.get('links') or {}).get('content') or f'https://zenodo.org/api/records/{rid}/files/{key}/content'
        if size > 20_000_000: continue
        if key.lower().endswith('.zip'):
            try: out['files'].update(unzip_bytes(fetch(url), prefix=key))
            except RuntimeError: pass
        elif key.lower().endswith(DATA_EXT + CODE_EXT):
            try: out['files'][key] = fetch(url)
            except RuntimeError: pass
    return out


def unzip_bytes(data: bytes, prefix='') -> dict:
    files = {}
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for n in z.namelist():
            if n.endswith('/') or '__MACOSX' in n: continue
            if n.lower().endswith(DATA_EXT + CODE_EXT) or re.match(r'readme', Path(n).name, re.I):
                if z.getinfo(n).file_size <= 10_000_000: files[f'{prefix}/{n}'] = z.read(n)
    return files


def read_sra(acc: str) -> dict:
    fields = 'run_accession,library_name,sample_alias,sample_accession,instrument_model,library_layout,read_count,fastq_ftp,fastq_md5'
    tsv = fetch(f'https://www.ebi.ac.uk/ena/portal/api/search?result=read_run&query=study_accession={acc}&fields={fields}').decode('utf-8')
    rows = list(csv.DictReader(io.StringIO(tsv), delimiter='\t'))
    return {'_label': f'ENA/SRA {acc}', 'accession': acc, 'runs': rows}


# ── understanding data tables ────────────────────────────────────────────────

def parse_table(name: str, data: bytes):
    raw = data.decode('utf-8-sig', errors='replace')
    if not raw.strip(): return None
    try: dialect = csv.Sniffer().sniff(raw[:5000], delimiters=',;\t')
    except csv.Error: dialect = csv.excel_tab if name.lower().endswith('.tsv') else csv.excel
    rows = [r for r in csv.reader(io.StringIO(raw), dialect) if any(c.strip() for c in r)]
    return rows if len(rows) >= 2 and len(rows[0]) >= 2 else None


def table_kind(rows) -> str:
    hdr = [h.strip().lower() for h in rows[0]]
    body = rows[1:]
    if len(hdr) == 2 and re.search(r'blank|control|neg', hdr[0]) and re.search(r'sample', hdr[1]): return 'blank_links'
    numeric = lambda v: re.fullmatch(r'-?\d+(\.\d+)?', v.strip() or 'x') is not None
    num_cols = sum(1 for j in range(1, len(hdr)) if all(numeric(r[j]) for r in body[:50] if j < len(r) and r[j].strip()))
    if num_cols >= 5 and num_cols >= 0.6 * (len(hdr) - 1):
        vals = {r[j].strip() for r in body[:50] for j in range(1, len(r)) if j < len(r) and numeric(r[j])}
        return 'presence' if vals <= {'0', '1'} else 'reads'
    if any(re.fullmatch(SAMPLE_ALIASES['samp_name'], h, re.I) for h in hdr) and \
       sum(1 for h in hdr if re.search(r'date|time|lat|lon|site|habitat|type|volume|depth|location', h)) >= 2: return 'samples'
    return 'other'


def matrix_from(rows):
    """{'samples': [...], 'taxa': [{'label', 'taxonomy', 'counts': {sample: n}}]} – taxa as rows or as columns."""
    hdr = [h.strip() for h in rows[0]]
    tax_col = next((j for j, h in enumerate(hdr) if re.fullmatch(r'taxonomy|lineage|taxon_?path|classification', h, re.I)), None)
    seq_col = next((j for j, h in enumerate(hdr) if re.fullmatch(r'sequence|seq|dna_sequence|asv_sequence', h, re.I)), None)
    cols = [j for j in range(1, len(hdr)) if j not in (tax_col, seq_col)]
    taxa = []
    for r in rows[1:]:
        if not r or not r[0].strip(): continue
        counts = {}
        for j in cols:
            if j < len(r) and re.fullmatch(r'\d+(\.\d+)?', r[j].strip() or 'x'): counts[hdr[j]] = float(r[j])
        taxa.append({'label': r[0].strip(), 'taxonomy': r[tax_col].strip() if tax_col is not None and tax_col < len(r) else '',
                     'sequence': r[seq_col].strip() if seq_col is not None and seq_col < len(r) else '', 'counts': counts})
    return {'samples': [hdr[j] for j in cols], 'taxa': taxa}


RANKS = [('kingdom', 'k'), ('phylum', 'p'), ('class', 'c'), ('order', 'o'), ('family', 'f'), ('genus', 'g'), ('species', 's')]


def taxonomy_fields(label: str, taxonomy: str) -> dict:
    """'d__Eukaryota; p__Chordata; …; s__Rana_temporaria' (or 'Animalia;Chordata;…') -> FAIRe taxa columns."""
    out = {}
    parts = [p.strip() for p in re.split(r'[;|]', taxonomy or '') if p.strip()]
    if parts and re.match(r'[a-z]__', parts[0]):
        for p in parts:
            code, name = p.split('__', 1)
            name = name.replace('_', ' ').strip()
            rank = {'d': None, 'k': 'kingdom', 'p': 'phylum', 'c': 'class', 'o': 'order', 'f': 'family', 'g': 'genus', 's': 'species'}.get(code)
            if rank and name and not re.match(r'(unassigned|unknown|NA)$', name, re.I): out[rank] = name
    elif len(parts) >= 3:
        for (rank, _), name in zip(RANKS, parts):
            if name and not re.match(r'(unassigned|unknown|NA)$', name, re.I): out[rank] = name.replace('_', ' ')
    if 'kingdom' not in out and out.get('phylum') in ('Chordata', 'Arthropoda', 'Mollusca', 'Annelida', 'Cnidaria'): out['kingdom'] = 'Animalia'
    sp = out.pop('species', None) or (label.replace('_', ' ') if re.fullmatch(r'[A-Z][a-z]+[_ ][a-z]+', label) else None)
    rec = {k: out.get(k) for k, _ in RANKS[:-1]}
    if sp:
        rec['genus'] = rec.get('genus') or sp.split()[0]
        rec['specificEpithet'] = sp.split()[1] if len(sp.split()) > 1 else None
        rec['scientificName'], rec['taxonRank'] = sp, 'species'
    else:
        deepest = next(((k, out[k]) for k, _ in reversed(RANKS[:-1]) if out.get(k)), None)
        if deepest: rec['scientificName'], rec['taxonRank'] = deepest[1], deepest[0]
        else: rec['scientificName'], rec['taxonRank'] = label.replace('_', ' '), None
    rec['verbatimIdentification'] = taxonomy or label
    return {k: v for k, v in rec.items() if v}


# ── putting the sources together ─────────────────────────────────────────────

PREFER_SOURCE = {  # which source wins when sources disagree
    'protocols.io': METHOD_TERMS | {'sterilise_method', 'pcr_cond', 'custom_mm', 'pcr_rep', 'trim_param', 'demux_tool', 'merge_tool',
                                    'merge_min_overlap', 'min_len_cutoff', 'chimera_check_method', 'otu_clust_tool', 'min_reads_cutoff',
                                    'screen_contam_method', 'tax_class_collapse', 'otu_seq_comp_appr', 'nucl_acid_ext_sep',
                                    'nucl_acid_ext_lysis', 'pcr2_commercial_mm', 'pcr2_custom_mm', 'pcr2_cond', 'pcr2_amplificationReactionVolume',
                                    'pcr2_dna_vol', 'thermocycler', 'lib_screen', 'seq_kit'},
    'dataset': {'recordedByID', 'license', 'institution', 'institutionID', 'accessRights', 'rightsHolder'},
}


def merge_source(meta: dict, other: dict, label: str, prefer: set) -> dict:
    """Fill gaps from another source and settle disagreements (preferred terms take the other source's value)."""
    mine = {it['id']: it for it in review_items(meta)}
    if not meta.get('assays') and other.get('assays'):
        meta['assays'] = [dict(a) for a in other['assays']]
        for k, v in (other.get('_evidence') or {}).items():
            if re.match(r'[A-Za-z0-9]+\.', k): meta.setdefault('_evidence', {}).setdefault(k, v)
        mine = {it['id']: it for it in review_items(meta)}
    m_of = {a.get('marker'): i for i, a in enumerate(meta.get('assays') or [])}
    edits, filled, overridden, agreed = {}, 0, 0, 0
    for it in review_items(other):
        if it['generated'] or not it['value'] or it['value'].startswith(('missing:', 'not applicable')): continue
        lid = it['id']
        if it['scope'] == 'assay':
            mk = (other.get('assays') or [])[int(lid.split(':')[1])].get('marker')
            if mk not in m_of: continue
            lid = f'assay:{m_of[mk]}:{it["term"]}'
        if lid not in mine: continue
        ours, theirs = mine[lid]['value'], it['value']
        ev = f'{label}: ' + (it.get('evidence') or '')[:300]
        if not ours or ours.startswith('missing:'):
            edits[lid] = theirs; filled += 1
            _flag(meta, lid, 'ok', f'from {label}', origin=label, evidence=ev)
        elif same_value(ours, theirs):
            agreed += 1; _flag(meta, lid, 'ok', f'{label} agrees')
        elif it['term'] in prefer:
            edits[lid] = theirs; overridden += 1
            _flag(meta, lid, 'check', f'taken from {label}; the report says "{ours[:80]}"', origin=label, evidence=ev)
        else:
            _flag(meta, lid, 'check', f'{label} says "{theirs[:80]}"')
    apply_review(meta, edits)
    meta.setdefault('notes', []).append(f'{label}: filled {filled} value(s), replaced {overridden}, agreed on {agreed}.')
    meta.setdefault('_sources', []).append(label)
    return meta


def apply_dataset_record(meta: dict, rec: dict):
    """DataCite metadata (Zenodo/Dryad/Figshare): authors with ORCIDs, licence, custodian institution."""
    label = rec.get('_label') or 'dataset record'
    edits = {}
    pairs = [('recordedBy', rec.get('recorded_by'), False), ('recordedByID', rec.get('recorded_by_id'), True),
             ('license', rec.get('license'), True), ('accessRights', rec.get('access_rights'), True),
             ('institution', rec.get('institution'), True), ('institutionID', rec.get('institution_ror'), True)]
    cur = {it['id']: it['value'] for it in review_items(meta)}
    for term, val, prefer in pairs:
        if not val: continue
        lid = f'project:{term}'
        ours = cur.get(lid, '')
        if not ours or ours.startswith('missing:'):
            edits[lid] = val; _flag(meta, lid, 'ok', f'from {label}', origin=label, evidence=f'{label}: {term}')
        elif same_value(ours, val): _flag(meta, lid, 'ok', f'{label} agrees')
        elif prefer:
            edits[lid] = val; _flag(meta, lid, 'check', f'taken from {label}; the report says "{ours[:80]}"', origin=label, evidence=f'{label}: {term}')
        else: _flag(meta, lid, 'check', f'{label} says "{val[:80]}"')
    if rec.get('institution') and not rec.get('institution_ror') and rec['institution'] != cur.get('project:institution'):
        edits['project:institutionID'] = ''           # the old ROR belonged to the replaced institution
    doi = rec.get('doi')
    if doi:
        assoc = (meta.get('associated_resource') or '')
        if doi not in assoc: meta['associated_resource'] = ' | '.join(x for x in (assoc, f'https://doi.org/{doi}') if x)
    apply_review(meta, edits)
    meta.setdefault('_sources', []).append(label)


def apply_files(meta: dict, files: dict, label: str):
    """Data tables and code from a repository / archive / local files."""
    tables = {}
    for name, data in sorted(files.items(), key=lambda kv: not re.match(r'readme', Path(kv[0]).name, re.I)):
        if re.match(r'readme', Path(name).name, re.I): CONTROL_CODES.update(control_codes_from(data.decode('utf-8', 'replace')))
        if name.lower().endswith(DATA_EXT):
            rows = parse_table(name, data)
            if rows: tables[name] = (table_kind(rows), rows)
    reads = [(n, matrix_from(r)) for n, (k, r) in tables.items() if k == 'reads']
    pres = [(n, matrix_from(r)) for n, (k, r) in tables.items() if k == 'presence']
    samples = [(n, r) for n, (k, r) in tables.items() if k == 'samples']
    links = [(n, r) for n, (k, r) in tables.items() if k == 'blank_links']
    notes = []
    if reads:
        name, mx = max(reads, key=lambda x: len(x[1]['taxa']))
        meta['_read_matrix'] = mx
        meta['taxa_raw'] = [dict(seq_id=t['label'], dna_sequence=t['sequence'] or None, **taxonomy_fields(t['label'], t['taxonomy']))
                            for t in mx['taxa']]
        meta['otu_raw_file'] = name
        if not meta.get('otu_raw_description') or 'quality filtering' in meta.get('otu_raw_description', ''):
            meta['otu_raw_description'] = f'Taxon × sample read table before filtering ({Path(name).name})'
            meta.setdefault('_evidence', {})['otu_raw_description'] = f'{label}: {name}'
        notes.append(f'taxaRaw: {len(mx["taxa"])} taxa from {name}' + ('' if any(t['sequence'] for t in mx['taxa']) else ' (no sequences in the file)'))
    if samples:
        in_matrix = set(meta['_read_matrix']['samples']) if meta.get('_read_matrix') else None
        def table_recs(rows):
            hdr = [h.strip() for h in rows[0]]
            sid_ = next(h for h in hdr if re.fullmatch(SAMPLE_ALIASES['samp_name'], h, re.I))
            return sid_, [{k: v for k, v in zip(hdr, r) if k} for r in rows[1:]]
        def score(item):                 # samples matching the read table, then how much each row tells us
            sid_, recs_ = table_recs(item[1])
            hit = [r for r in recs_ if not in_matrix or r.get(sid_, '').strip() in in_matrix]
            filled = sum(1 for r in hit for k, v in r.items() if k != sid_ and str(v).strip())
            return (len(hit) > 0, filled)
        samples.sort(key=score, reverse=True)
        name = samples[0][0]
        sid, recs = table_recs(samples[0][1])
        by_id = {r.get(sid, '').strip(): r for r in recs}
        for other_name, other_rows in samples[1:]:          # fill gaps from the other sample tables
            osid, orecs = table_recs(other_rows)
            for r in orecs:
                key = r.get(osid, '').strip()
                if key in by_id:
                    for k, v in r.items():
                        if k != osid and str(v).strip() and not str(by_id[key].get(k, '')).strip(): by_id[key][k] = v
                elif in_matrix and key in in_matrix:
                    nr = {sid: key}; nr.update({k: v for k, v in r.items() if k != osid}); recs.append(nr); by_id[key] = nr
        if in_matrix:
            keep = [r for r in recs if r.get(sid, '').strip() in in_matrix]
            extra = sorted(in_matrix - {r.get(sid, '').strip() for r in recs})
            recs = keep + [{sid: s} for s in extra]           # controls and replicates only in the read table
            notes.append(f'{len(keep)} samples from {name} matched the read table; {len(extra)} more (mostly controls) taken from the read table')
        for r in recs:
            if r.get('date') and r.get('time') and re.match(r'\d', r['time']):
                d = parse_date(r['date']) or r['date']
                r['date'] = f'{d}T{r["time"][:5]}'; r.pop('time')
        meta['sample_list'] = recs
        meta['sample_list_source'] = name
    elif meta.get('_read_matrix'):
        meta['sample_list'] = [{'sample': s} for s in meta['_read_matrix']['samples']]
    if links:
        name, rows = links[0]
        rel = {}
        for b, s in rows[1:]:
            if b.strip() != s.strip(): rel.setdefault(s.strip(), []).append(b.strip())
        meta['_blank_links'] = rel
        notes.append(f'control–sample links from {name}')
    if pres:
        name, mx = max(pres, key=lambda x: len(x[1]['taxa']))
        sample_set = set(meta.get('_read_matrix', {}).get('samples', []))
        cols = [s for s in mx['samples'] if not sample_set or s in sample_set]
        if len(cols) < 3 and len(mx['taxa']) and sample_set:      # samples as rows, taxa as columns
            rows_ = tables[name][1]
            taxa_cols = rows_[0][1:]
            kept = {t for r in rows_[1:] if r and r[0].strip() in sample_set for t, v in zip(taxa_cols, r[1:]) if v.strip() == '1'}
        else:
            kept = {t['label'] for t in mx['taxa'] if any(t['counts'].get(c) for c in cols)}
        raw = {t['seq_id']: t for t in meta.get('taxa_raw') or []}
        meta['taxa_final'] = [raw.get(k) or dict(seq_id=k, **taxonomy_fields(k, '')) for k in sorted(kept)]
        for a in meta.get('assays') or []:
            a.setdefault('otu_final_description', f'Presence–absence table after filtering ({Path(name).name})')
        notes.append(f'taxaFinal: {len(kept)} taxa from {name}')
    for n, d in files.items():
        if re.match(r'readme', Path(n).name, re.I): CONTROL_CODES.update(control_codes_from(d.decode('utf-8', 'replace')))
    if CONTROL_CODES: notes.append('control naming: ' + ', '.join(f'{k} = {v}' for k, v in CONTROL_CODES.items()))
    code = {n: d.decode('utf-8', 'replace') for n, d in files.items() if n.lower().endswith(CODE_EXT) or re.match(r'readme', Path(n).name, re.I)}
    if code:
        txt = '\n'.join(code.values())
        hints = []
        for pat, what in ((r'filterAndTrim\([^)]*maxEE\s*=\s*c?\(?([\d., ]+)', 'maxEE'), (r'truncLen\s*=\s*c?\(?([\d, ]+)', 'truncLen'),
                          (r'minBoot\s*=\s*(\d+)', 'minBoot'), (r'--cluster_unoise|--minsize\s+\d+|--uchime3?_denovo', None)):
            m = re.search(pat, txt)
            if m: hints.append(f'{what} = {m.group(1).strip()}' if what else m.group(0))
        if hints: meta.setdefault('bioinfo_method_additional', None); notes.append('settings seen in the code: ' + '; '.join(hints))
    for n in notes: meta.setdefault('notes', []).append(f'{label}: {n}')
    meta.setdefault('_sources', []).append(label)


def apply_sra(meta: dict, rec: dict):
    runs = rec.get('runs') or []
    if not runs: return
    label = rec['_label']
    names = {str(s.get('samp_name') or s.get('sample') or '') for s in meta.get('sample_list') or []}
    per_sample = {r['sample_alias']: r for r in runs if r.get('sample_alias') in names} or \
                 {r['library_name']: r for r in runs if r.get('library_name') in names}
    meta['sra_runs'] = runs
    meta['sra_by_sample'] = per_sample
    meta['sra_by_library'] = {r.get('library_name'): r for r in runs if r.get('library_name')}
    inst = runs[0].get('instrument_model')
    edits = {}
    cur = {it['id']: it['value'] for it in review_items(meta)}
    for term, val in (('instrument', inst), ('platform', 'ILLUMINA' if re.search(r'illumina|miseq|novaseq|nextseq|hiseq', inst or '', re.I) else None),
                      ('lib_layout', {'PAIRED': 'paired end', 'SINGLE': 'single end'}.get((runs[0].get('library_layout') or '').upper())),
                      ('checksum_method', 'MD5' if runs[0].get('fastq_md5') else None)):
        if val and (not cur.get(f'project:{term}') or cur[f'project:{term}'].startswith('missing')):
            edits[f'project:{term}'] = val; _flag(meta, f'project:{term}', 'ok', f'from {label}', origin=label, evidence=f'{label}: {term}')
        elif val: _flag(meta, f'project:{term}', 'ok' if same_value(val, cur.get(f'project:{term}', '')) else 'check', f'{label}: {val}')
    apply_review(meta, edits)
    acc = rec.get('accession')
    if acc and acc not in (meta.get('associated_resource') or ''):
        meta['associated_resource'] = ' | '.join(x for x in (meta.get('associated_resource'), f'https://www.ncbi.nlm.nih.gov/bioproject/{acc}') if x)
    if not per_sample:
        meta.setdefault('notes', []).append(
            f'{label}: {len(runs)} sequencing runs are pooled libraries ({", ".join(sorted(r.get("library_name") or r["run_accession"] for r in runs)[:6])}…), '
            'not one run per sample. Add a "library" (or seq_run_id) column to the sample list to link samples to their run files.')
    meta.setdefault('_sources', []).append(label)


def add_sources(meta: dict, sources: list, report_text: str = '', follow_links: bool = False) -> dict:
    """Read every extra source and merge it into the report's metadata. Problems become notes, never errors."""
    todo = list(sources or [])
    code_only = set()
    if follow_links and report_text:
        found = find_links(report_text)
        avail = ' '.join(_squash(x) for x in sentences(report_text) if re.search(r'data availab|available (?:via|at|on|from|in)|deposited|archived|accession', x, re.I)).lower()
        for f in found:
            if 'github.com' in f and f.split('github.com/')[-1].lower()[:20] not in avail: code_only.add(f)   # e.g. a pipeline's repo
        todo += [f for f in found if f not in todo]
        if found: meta.setdefault('notes', []).append('Links found in the report: ' + ', '.join(found))
    seen_github = set()
    local_files = {}
    order = {'sra': 2}
    todo = sorted(todo, key=lambda x: order.get(classify(x), 0))       # sequence archives last: they link to samples
    for i_src, src in enumerate(todo):
        kind = classify(src)
        if kind == 'sra' and local_files:                               # local data files first, all together
            apply_files(meta, local_files, 'your files'); local_files = {}
        try:
            if kind in ('protocol', 'file:protocol'):
                if kind == 'protocol': proto = protocol_from_url(src)
                else:
                    ptxt = extract_report_text(src)
                    d = re.search(r'10\.17504/protocols\.io\.[a-z0-9]+(?:/v\d+)?', ptxt)
                    proto = read_protocol_text(ptxt, 'protocols.io' if d else Path(src).name, d.group(0) if d else None)
                merge_source(meta, proto, proto.get('_label') or 'protocol', PREFER_SOURCE['protocols.io'])
                for k in ('pcr2_amplificationReactionVolume', 'pcr2_dna_vol', 'pcr2_commercial_mm', 'pcr2_custom_mm', 'pcr2_cond',
                          'thermocycler', 'pcr2_annealingTemp', 'pcr2_cycles', 'pcr2_thermocycler', 'nucl_acid_ext', 'nucl_acid_amp'):
                    if proto.get(k) and not meta.get(k):
                        meta[k] = proto[k]; meta.setdefault('_evidence', {})[k] = proto['_evidence'].get(k)
            elif kind == 'zenodo':
                z = read_zenodo(src)
                if z.get('_error'): meta.setdefault('notes', []).append(f'{z["_label"]}: {z["_error"]}')
                meta.setdefault('notes', []).append(f'{z["_label"]}: {len(z.get("file_list") or [])} files listed, '
                                                    f'{len(z.get("files") or {})} data/code files read' +
                                                    (f'; authors, ORCIDs and licence from the record' if z.get('recorded_by') else ''))
                apply_dataset_record(meta, z)
                if z.get('files'): apply_files(meta, z['files'], z['_label'])
                big = [k for k, s, _ in z.get('file_list', []) if (s or 0) > 20_000_000]
                if big: meta.setdefault('notes', []).append(f'{z["_label"]}: large files not downloaded: ' + ', '.join(big))
            elif kind == 'doi':
                apply_dataset_record(meta, read_datacite(src))
            elif kind == 'github':
                if src in seen_github: continue
                seen_github.add(src)
                if src in code_only:
                    if src not in (meta.get('code_repo') or ''): meta['code_repo'] = ' | '.join(x for x in (meta.get('code_repo'), src) if x)
                    continue
                g = read_github(src)
                if g.get('skipped'): meta.setdefault('notes', []).append(f'{g["_label"]}: could not read ' + '; '.join(g['skipped'][:3]))
                if g.get('code_repo') and g['code_repo'] not in (meta.get('code_repo') or ''):
                    meta['code_repo'] = ' | '.join(x for x in (meta.get('code_repo'), g['code_repo']) if x)
                apply_files(meta, g['files'], g['_label'])
            elif kind == 'sra':
                apply_sra(meta, read_sra(src))
            elif kind == 'file:data':
                p = Path(src); nm = re.sub(r'^extra_', '', p.name)
                if p.suffix.lower() == '.xlsx':
                    rows = read_table(p); hdr = list(rows[0]) if rows else []
                    data = '\n'.join(','.join(f'"{str(r.get(h, ""))}"' for h in hdr) for r in rows)
                    local_files[re.sub(r'\.xlsx$', '.csv', nm)] = (','.join(hdr) + '\n' + data).encode()
                else: local_files[nm] = p.read_bytes()
            elif kind == 'file:code':
                local_files[re.sub(r'^extra_', '', Path(src).name)] = Path(src).read_bytes()
            elif kind == 'file:zip':
                local_files.update(unzip_bytes(Path(src).read_bytes(), Path(src).stem))
            else:
                meta.setdefault('notes', []).append(f'Not sure what "{src}" is – skipped.')
        except Exception as e:
            meta.setdefault('notes', []).append(f'Could not use {src}: {e}')
    if local_files:
        try: apply_files(meta, local_files, 'your files')
        except Exception as e: meta.setdefault('notes', []).append(f'Could not use your files: {e}')
    return meta


# ════════════════════════════════════════════════════════════════════════════
#  7. Google Colab: a point-and-click review screen (used by the notebook)
# ════════════════════════════════════════════════════════════════════════════

_COLAB: dict = {}
_CSS = """<style>
.fx,.fx-card,.fx-card *{font-family:system-ui,-apple-system,Segoe UI,Roboto,sans-serif}
.fx{font-size:14px;color:inherit;line-height:1.5}
.fx h3{margin:.6em 0 .3em;font-size:17px}
.fx .sum{background:#F5F9F8;color:#1d2b2a;border:1px solid #D5EAE6;border-radius:10px;padding:.7em 1em;margin:.4em 0}
.fx .sum small{color:#3d4b4a}
.fx .b{display:inline-block;border-radius:99px;padding:.1em .7em;font-weight:600;font-size:13px;margin:0 .3em .3em 0}
.fx .ok{background:#DCE8CF;color:#2F4519}.fx .ck{background:#FFE9A8;color:#5C4100}.fx .ms{background:#E4E9E9;color:#3d4b4a}
.fx-card{background:#ffffff !important;color:#1d2b2a !important;border:1px solid #D5EAE6;border-left:5px solid #4679E1;
         border-radius:10px;padding:10px 14px !important;margin:8px 2px !important}
.fx-card.check{border-left-color:#E0A800}.fx-card.missing{border-left-color:#9AA5A4}.fx-card.okc{border-left-color:#90A968}
.fx-card .term{font-weight:700;font-size:15px;color:#1A3B8B}
.fx-card .tag{display:inline-block;margin-left:6px;background:#D1DEF8;color:#1A3B8B;border-radius:99px;padding:0 8px;font-size:12px;font-weight:600}
.fx-card .why{color:#7A4E00;font-size:13px;margin-top:4px}
.fx-card .why.good{color:#3A5020}
.fx-card .src{color:#3d4b4a;font-size:13px;margin-top:4px;background:#F5F9F8;border-radius:6px;padding:6px 8px}
.fx-card .src b{color:#1d2b2a}
.fx-card textarea{background:#ffffff !important;color:#111111 !important;border:1.5px solid #9DB8B3 !important;border-radius:6px !important;
                  font-size:14px !important;padding:6px 8px !important}
.fx-card textarea:focus{border-color:#4679E1 !important}
</style>"""


def _esc(t):
    return html.escape(str(t or ''))


def colab_read(use_my_own_template=False, use_a_profile=False, report_is_not_confidential=False, use_a_sample_list=False,
               extra_sources='', read_links_in_the_report=False):
    """Notebook step: upload a report, read it, and show the review screen with a Fill button."""
    from google.colab import files as colab_files
    from IPython.display import display, HTML
    import urllib.request
    display(HTML(_CSS + '<div class="fx"><h3>① Upload your report</h3>.doc, .docx, .pdf or .txt – it stays in this Colab session</div>'))
    up = colab_files.upload()
    if not up: return display(HTML('<div class="fx">No file chosen – run this cell again.</div>'))
    report = Path(list(up)[0])
    template = APP_DIR / 'template' / TEMPLATE_FILE
    if use_my_own_template:
        display(HTML('<div class="fx">Upload your FAIRe <b>template</b> (.xlsx):</div>'))
        template = Path(list(colab_files.upload())[0])
    elif not template.exists():
        template = Path(TEMPLATE_FILE)
        urllib.request.urlretrieve('https://github.com/UKCEH-MolecularEcology/fair-edna-filler/raw/main/template/' + TEMPLATE_FILE, template)
    profile = None
    if use_a_profile:
        display(HTML('<div class="fx">Upload your <b>profile</b> (.json):</div>'))
        profile = json.loads(Path(list(colab_files.upload())[0]).read_text(encoding='utf-8'))
    samples = None
    if use_a_sample_list:
        display(HTML('<div class="fx">Upload your <b>sample list</b> (.csv or .xlsx, one row per sample):</div>'))
        samples = Path(list(colab_files.upload())[0])
    key = ''
    if report_is_not_confidential:
        try:
            from google.colab import userdata
            key = userdata.get('ANTHROPIC_API_KEY') or ''
        except Exception:
            key = ''
        if not key:
            from getpass import getpass
            key = getpass('Claude API key: ').strip()
    text = extract_report_text(report)
    srcs = [x for x in re.split(r'[\s,;]+', extra_sources or '') if x]
    if srcs or read_links_in_the_report:
        display(HTML('<div class="fx">Reading the extra sources (public links only)…</div>'))
    meta = prepare(text, profile=profile, local_ai=None, samples=samples, sources=srcs, follow_links=read_links_in_the_report)
    if key:
        try: second_opinion(meta, facts_from_claude(parse_with_claude(text, key), text), 'Claude')
        except Exception as e: meta['notes'].append(f'Claude could not be used ({e}).')
    items = review(meta, template)
    _COLAB.update(meta=meta, items=items, template=template, report=report, profile=profile)
    _review_screen(items, meta)


def _review_screen(items, meta):
    import ipywidgets as w
    from IPython.display import display, HTML
    c = review_summary(items)
    notes = ''.join(f'<li>{_esc(n)}</li>' for n in meta.get('notes') or [])
    display(HTML(_CSS + f'''<div class="fx"><h3>② Check the values</h3><div class="sum">
      <span class="b ok">{c["ok"]} accepted automatically</span><span class="b ck">{c["check"]} need a look</span>
      <span class="b ms">{c["missing"]} not found in the report</span><br>
      <small>Read with: {_esc(meta.get("_method"))}. Change anything that is wrong – your changes are used when you click
      <b>Fill workbook</b>. Leave a box empty if the report doesn't say it.</small></div>
      {f"<b>Please note</b><ul>{notes}</ul>" if notes else ""}</div>'''))
    boxes = {}

    def rows(sel):
        out = []
        for it in sel:
            val = '' if it['value'].startswith(('missing:', 'not applicable')) else it['value']
            box = w.Textarea(value=val, placeholder='type a value if you know it', layout=w.Layout(width='100%', height='38px' if len(val) < 90 else '70px'))
            boxes[it['id']] = (box, val)
            tags = (f'<span class="tag">{_esc(it.get("assay"))}</span>' if it.get('assay') else '') + \
                   ('<span class="tag">all samples</span>' if it['scope'] == 'samples' else '') + \
                   ('<span class="tag">required</span>' if it.get('required') in ('M', 'HR') else '')
            reasons = ' · '.join(it.get('reasons') or [])
            why = f'<div class="why{"" if it["status"] == "check" else " good"}">{"⚠ " if it["status"] == "check" else ""}{_esc(reasons)}</div>' if reasons else ''
            src = f'<div class="src"><b>From the report:</b> “{_esc(it["evidence"][:400])}”</div>' if it.get('evidence') else ''
            card = w.VBox([w.HTML(f'<div><span class="term">{_esc(it["term"])}</span>{tags}</div>'), box, w.HTML(why + src)])
            card.add_class('fx-card'); card.add_class({'check': 'check', 'missing': 'missing'}.get(it['status'], 'okc'))
            out.append(card)
        return w.VBox(out or [w.HTML('<div class="fx">Nothing here.</div>')])

    shown = [it for it in items if not (it['generated'] and it['status'] == 'ok')]
    tabs = w.Tab(children=[rows([i for i in shown if i['status'] == 'check']),
                           rows([i for i in shown if i['status'] == 'missing' and i.get('required') in ('M', 'HR')]),
                           rows([i for i in shown if i['status'] == 'ok'])])
    for k, t in enumerate([f'Needs a look ({c["check"]})', f'Not found – required ({c["missing_required"]})', f'Accepted ({c["ok"]})']):
        tabs.set_title(k, t)
    save = w.Checkbox(value=False, description='Also save my corrections as a profile', indent=False)
    name = w.Text(value=(_COLAB.get('profile') or {}).get('name', 'my FAIRe profile'), description='Profile name:',
                  style={'description_width': 'initial'})
    btn = w.Button(description='Fill workbook and download', button_style='success', icon='download', layout=w.Layout(width='260px'))
    out = w.Output()

    def go(_):
        btn.disabled = True
        with out:
            out.clear_output()
            edits = {k: b.value.strip() for k, (b, orig) in boxes.items() if b.value.strip() != orig}
            colab_fill(edits, save.value, name.value)
        btn.disabled = False

    btn.on_click(go)
    display(tabs, HTML(_CSS + '<div class="fx"><h3>③ Fill the workbook</h3></div>'), w.HBox([save, name]), btn, out)


def colab_fill(edits=None, save_profile=False, profile_name='my FAIRe profile'):
    """Notebook step: apply corrections, write the workbook (and profile) and download them."""
    from google.colab import files as colab_files
    from IPython.display import display, HTML
    meta, template, report = _COLAB['meta'], _COLAB['template'], _COLAB['report']
    edits = edits or {}
    apply_review(meta, edits)
    tmp = Path('FAIRe_filled.xlsx')
    s = fill_workbook(template, meta, report.name, tmp)
    out = Path(f"FAIRe_filled_{s['projectId'] or 'project'}.xlsx"); tmp.replace(out)
    v = s['validation']
    checks = ''.join(f'<li>⚠️ {_esc(x["sheet"])} · {_esc(x["term"])}: {_esc(x["what"])} ({_esc(x["where"])})</li>' for x in v['problems']) + \
             ''.join(f'<li>🔧 fixed {_esc(x["sheet"])} · {_esc(x["term"])}: {_esc(x["what"])}</li>' for x in v['fixed'])
    display(HTML(_CSS + f'''<div class="fx"><div class="sum">✅ <b>{_esc(out.name)}</b> – {s["lakes"]} sites · {len(s["assays"])} assays ·
      {s["sampleRows"]} sample rows · {s["experimentRows"]} run rows · {len(edits)} of your corrections</div>
      <ul>{checks or "<li>✅ Passes the FAIRe format, vocabulary and ID checks</li>"}</ul>
      Next: open it in Excel, fill anything still marked <i>missing: not provided</i>, and validate at
      <a href="https://shiny.csiro.au/FAIRe-fier/" target="_blank">FAIRe-fier</a>.</div>'''))
    colab_files.download(str(out))
    if save_profile and edits:
        prof = profile_from_review(meta, edits, profile_name, _COLAB.get('profile'))
        pf = Path(re.sub(r'[^A-Za-z0-9_-]+', '_', profile_name or 'profile') + '.faire-profile.json')
        pf.write_text(json.dumps(prof, indent=1, ensure_ascii=False), encoding='utf-8')
        display(HTML(f'<div class="fx">💾 Profile saved: <b>{_esc(pf.name)}</b> ({len(prof["entries"])} corrections). '
                     'Tick “use a profile” next time to apply them.</div>'))
        colab_files.download(str(pf))


# ════════════════════════════════════════════════════════════════════════════
#  8. Running the tool: local web page or command line
# ════════════════════════════════════════════════════════════════════════════

REPORT_TYPES = ('.doc', '.docx', '.pdf', '.txt')


def prepare(text: str, profile: dict | None = None, local_ai: str | None = None, samples=None, sources=None, follow_links=False) -> dict:
    """Built-in reader + primer library + your profile (+ optional local AI second opinion, + optional sample list)."""
    CONTROL_CODES.clear()
    meta = parse_report_regex(text)
    if samples: meta['sample_list'] = read_table(samples) if isinstance(samples, (str, Path)) else list(samples)
    apply_library(meta)
    if sources or follow_links:
        user_list = meta.get('sample_list')
        add_sources(meta, sources or [], text, follow_links)
        if user_list: meta['sample_list'] = user_list   # your own sample list wins over files found online
    if profile: apply_profile(meta, profile)
    if local_ai:
        try: local_ai_opinion(meta, text, None if local_ai is True else local_ai)
        except Exception as e: meta['notes'].append(f'Local AI reader not used: {e}')
    return meta


REVIEW_COLUMNS = ['id', 'status', 'section', 'field', 'assay', 'value', 'your_value', 'why', 'source sentence in the report']


def write_review_csv(items, path):
    import csv
    order = {'check': 0, 'missing': 1, 'ok': 2}
    with open(path, 'w', newline='', encoding='utf-8') as fh:
        w = csv.writer(fh)
        w.writerow(REVIEW_COLUMNS)
        for it in sorted(items, key=lambda x: (order.get(x['status'], 3), 0 if x.get('required') in ('M', 'HR') else 1)):
            if it['generated'] and it['status'] == 'ok': continue
            w.writerow([it['id'], {'ok': 'accepted', 'check': 'CHECK', 'missing': 'not found'}[it['status']], it['section'],
                        it['term'] + (' (mandatory)' if it.get('required') == 'M' else ''),
                        it.get('assay') or ('all samples' if it['scope'] == 'samples' else ''), it['value'], '',
                        '; '.join(it.get('reasons') or []), it.get('evidence') or ''])


def read_review_csv(path) -> dict:
    """Your corrections from a review CSV: the 'your_value' column ('-' empties a value)."""
    import csv
    edits = {}
    with open(path, newline='', encoding='utf-8-sig') as fh:
        for row in csv.DictReader(fh):
            v = (row.get('your_value') or '').strip()
            if v and row.get('id'): edits[row['id']] = '' if v == '-' else v
    return edits


def process_report(report_path, template_path=None, out_dir=OUTPUT_DIR, name=None, profile=None, local_ai=None, edits=None, samples=None,
                   sources=None, follow_links=False) -> dict:
    """Read one report, fill the template, and write <name>_FAIRe.xlsx plus <name>_review.csv into out_dir.
    The review CSV lists what needs checking first; put corrections in its 'your_value' column and run again with --edits."""
    report_path, out_dir = Path(report_path), Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    text = extract_report_text(report_path)
    meta = prepare(text, profile=profile, local_ai=local_ai, samples=samples, sources=sources, follow_links=follow_links)
    if edits:
        apply_review(meta, edits)
        for k in edits: meta.setdefault('_auto', {})[k] = {'level': 'ok', 'reasons': ['your correction'], 'origin': 'user'}
    template = template_path or TEMPLATE_PATH
    items = review(meta, template)
    stem = name or re.sub(r'[^A-Za-z0-9._-]+', '_', report_path.stem).strip('_') or 'report'
    out = out_dir / f'{stem}_FAIRe.xlsx'
    summary = fill_workbook(template, meta, report_path.name, out)
    write_review_csv(items, out_dir / f'{stem}_review.csv')
    summary.update(output=str(out), review=review_summary(items), review_csv=str(out_dir / f'{stem}_review.csv'), meta=meta)
    return summary


class Handler(BaseHTTPRequestHandler):
    """Serves the same browser page as the online version, plus the bundled template."""

    def do_GET(self):
        path = unquote(self.path.split('?', 1)[0])
        if path in ('/', '/index.html'):
            return self._serve(APP_DIR / 'index.html', 'text/html; charset=utf-8')
        if path.startswith('/template/'):
            target = (APP_DIR / 'template' / Path(path).name)
            if target.suffix == '.xlsx' and target.is_file():
                return self._serve(target, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            if target.suffix == '.csv' and target.is_file():
                return self._serve(target, 'text/csv; charset=utf-8')
        self.send_error(404)

    do_HEAD = do_GET

    def _serve(self, target, content_type):
        data = Path(target).read_bytes()
        self.send_response(200)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        if self.command != 'HEAD': self.wfile.write(data)

    def log_message(self, fmt, *args):
        pass


def main(argv=None):
    import argparse, webbrowser
    p = argparse.ArgumentParser(
        description='Fill the FAIRe eDNA metadata template from project reports.',
        epilog='With no reports given, opens the tool in your web browser. '
               'Reports are processed on this computer only.')
    p.add_argument('reports', nargs='*', help='report files (.doc .docx .pdf .txt) or folders containing them')
    p.add_argument('-t', '--template', help='FAIRe template to fill (default: the bundled FULL template)')
    p.add_argument('-o', '--out', default=str(OUTPUT_DIR), help='folder for the filled workbooks (default: outputs/)')
    p.add_argument('--profile', help='apply a saved profile of corrections (JSON)')
    p.add_argument('--source', action='append', default=[], metavar='LINK_OR_FILE',
                   help='extra source: protocols.io link/DOI, Zenodo/Dryad/Figshare DOI, GitHub repo, BioProject (PRJNA…), or a local '
                        'protocol PDF / data table / code file. Repeat for several.')
    p.add_argument('--follow-links', action='store_true',
                   help='also read the protocol, dataset, code and sequence links printed in the report (downloads public files)')
    p.add_argument('--samples', help='your sample list (.csv or .xlsx, one row per sample: sample, date, site, latitude, longitude, …)')
    p.add_argument('--edits', help='a review CSV with your corrections in the your_value column (one report only)')
    p.add_argument('--save-profile', metavar='FILE', help='save the corrections from --edits into this profile (JSON)')
    p.add_argument('--profile-name', default=None, help='name stored in a new profile')
    p.add_argument('--local-ai', nargs='?', const=True, metavar='MODEL',
                   help=f'second opinion from a model running in Ollama on this computer (default model: {LOCAL_MODEL})')
    p.add_argument('--port', type=int, default=int(os.environ.get('FAIRE_TOOL_PORT', '8787')))
    args = p.parse_args(argv)

    if not args.reports:
        host = os.environ.get('FAIRE_TOOL_HOST', '127.0.0.1')
        url = f'http://{host}:{args.port}/'
        print(f'FAIRe Metadata Filler: {url}   (Ctrl+C to stop)')
        try: webbrowser.open(url)
        except Exception: pass
        ThreadingHTTPServer((host, args.port), Handler).serve_forever()
        return

    files = []
    for r in args.reports:
        r = Path(r)
        if not r.exists():
            print(f'✗ {r}: file not found – check the name (put it in quotes if it has spaces).'); continue
        files += sorted(f for f in r.iterdir() if f.suffix.lower() in REPORT_TYPES) if r.is_dir() else [r]
    if not files: raise SystemExit('No reports found.')
    if args.edits and len(files) > 1: raise SystemExit('--edits works with one report at a time.')
    profile = json.loads(Path(args.profile).read_text(encoding='utf-8')) if args.profile else None
    edits = read_review_csv(args.edits) if args.edits else None
    stems = [f.stem for f in files]
    for f in files:
        name = re.sub(r'[^A-Za-z0-9._-]+', '_', f.stem).strip('_') or 'report'
        if stems.count(f.stem) > 1: name += '_' + f.suffix.lstrip('.').lower()     # report.doc + report.pdf
        try:
            s = process_report(f, args.template, args.out, name, profile, args.local_ai, edits, args.samples, args.source, args.follow_links)
            r = s['review']
            print(f'✓ {f.name}: {len(s["assays"])} assays, {s["sampleRows"]} sample rows' +
                  ''.join(f', {n} {k} rows' for k, n in (s.get('taxa') or {}).items()) + f' -> {s["output"]}')
            print(f'    {r["ok"]} values accepted automatically · {r["check"]} to check · {r["missing"]} not found '
                  f'({r["missing_required"]} mandatory/highly recommended) -> {s["review_csv"]}')
            for v in s['validation']['fixed']: print(f'    fixed: {v["sheet"]}.{v["term"]} {v["what"]} ({v["cells"]} cells)')
            for v in s['validation']['problems']: print(f'    problem: {v["sheet"]}.{v["term"]} {v["what"]} ({v["where"]})')
            for n in s['notes']: print(f'    note: {n}')
            if args.save_profile and edits:
                old = json.loads(Path(args.save_profile).read_text(encoding='utf-8')) if Path(args.save_profile).exists() else None
                prof = profile_from_review(s['meta'], edits, args.profile_name or (old or {}).get('name') or Path(args.save_profile).stem, old)
                Path(args.save_profile).write_text(json.dumps(prof, indent=1, ensure_ascii=False), encoding='utf-8')
                print(f'    profile saved: {args.save_profile} ({len(prof["entries"])} corrections)')
        except Exception as e:
            print(f'✗ {f.name}: {e}')


if __name__ == '__main__':
    main()
