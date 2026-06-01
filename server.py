#!/usr/bin/env python3
"""
FAIRe Metadata Filler – web server.

Run locally:
    python3 server.py

Then open http://127.0.0.1:8787

For online deployment set:
    ANTHROPIC_API_KEY  – enables Claude-powered document parsing (recommended)
    FAIRE_TOOL_HOST    – bind address (default 127.0.0.1; set 0.0.0.0 for public)
    FAIRE_TOOL_PORT    – port (default 8787)
"""

from __future__ import annotations

import html
import json
import os
import re
import shutil
import subprocess
import tempfile
import time
import uuid
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

APP_DIR = Path(__file__).resolve().parent
STATIC_DIR = APP_DIR / "static"
OUTPUT_DIR = APP_DIR / "outputs"
UPLOAD_DIR = APP_DIR / "uploads"
OUTPUT_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("",  MAIN)
ET.register_namespace("r", REL)
NS = {"m": MAIN, "r": REL}

MISS          = "missing: not provided"
NA            = "not applicable"
NA_SAMPLE     = "not applicable: sample group"
NA_CONTROL    = "not applicable: control sample"
MAX_REPORT_CHARS = 80_000   # chars sent to Claude


# ── data classes ──────────────────────────────────────────────────────────────

@dataclass
class AssayInfo:
    name: str
    target_gene:          str = MISS
    target_subfragment:   str = MISS
    amplicon_size:        str = MISS
    primer_fwd:           str = MISS
    primer_rev:           str = MISS
    primer_name_fwd:      str = MISS
    primer_name_rev:      str = MISS
    primer_ref_fwd:       str = MISS
    primer_ref_rev:       str = MISS
    commercial_mm:        str = MISS
    custom_mm:            str = MISS
    rxn_vol:              str = MISS
    dna_vol:              str = MISS
    pcr_rep:              str = "1"
    annealing_temp:       str = MISS
    pcr_cycles:           str = MISS
    pcr_cond:             str = MISS
    tax_db:               str = MISS
    tax_method:           str = MISS
    min_len:              str = MISS
    seq_kit:              str = MISS
    rarefaction_depth:    str = MISS


@dataclass
class Lake:
    name:           str
    code:           str
    lake_type:      str  = MISS
    wfd_tp:         str  = MISS
    wfd_chlorophyll:str  = MISS
    trophic_status: str  = MISS
    latitude:       str  = MISS
    longitude:      str  = MISS


@dataclass
class ReportFacts:
    project_title:       str
    project_id:          str
    institution:         str   = MISS
    institution_id:      str   = MISS
    project_contact:     str   = MISS
    funder:              str   = MISS
    start_date:          str   = MISS
    end_date:            str   = MISS
    project_cost:        str   = MISS
    license:             str   = MISS
    access_rights:       str   = MISS
    bibliography:        str   = MISS
    code_repo:           str   = MISS
    sterilise_method:    str   = "UV-sterilised laminar flow hood"
    env_broad_scale:     str   = "freshwater lake biome [ENVO:01001777]"
    env_local_scale:     str   = "lake [ENVO:00000020]"
    env_medium:          str   = "lake water [ENVO:00000044]"
    samp_collect_method: str   = MISS
    samp_size:           str   = MISS
    samp_size_unit:      str   = MISS
    filter_pore:         str   = MISS
    filter_diameter:     str   = MISS
    filter_material:     str   = MISS
    filter_name:         str   = MISS
    samp_store_temp:     str   = MISS
    samp_store_sol:      str   = MISS
    ext_kit:             str   = MISS
    ext_modifications:   str   = MISS
    ext_lysis:           str   = MISS
    ext_sep:             str   = MISS
    conc_method:         str   = MISS
    normalisation:       str   = MISS
    barcoding_pcr_appr:  str   = MISS
    platform:            str   = MISS
    instrument:          str   = MISS
    lib_layout:          str   = MISS
    adapter_fwd:         str   = MISS
    adapter_rev:         str   = MISS
    pcr2_mm:             str   = MISS
    pcr2_rxn_vol:        str   = MISS
    pcr2_dna_vol:        str   = MISS
    pcr2_cond:           str   = MISS
    seq_location:        str   = MISS
    phix_perc:           str   = MISS
    trim_method:         str   = MISS
    trim_param:          str   = MISS
    demux_tool:          str   = MISS
    merge_tool:          str   = MISS
    error_rate_type:     str   = MISS
    error_rate_cutoff:   str   = MISS
    chimera_method:      str   = MISS
    chimera_param:       str   = MISS
    clust_tool:          str   = MISS
    screen_contam:       str   = MISS
    otu_raw_desc:        str   = MISS
    otu_final_desc:      str   = MISS
    bioinfo_extra:       str   = MISS
    study_factor:        str   = MISS
    assays:              list  = field(default_factory=list)
    lakes:               list  = field(default_factory=list)
    rounds:              list  = field(default_factory=list)  # (code, iso_date, verbatim, avg_vol_ml)
    report_summary:      str   = ""


# ── text extraction ────────────────────────────────────────────────────────────

def extract_report_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".txt":
        return path.read_text(errors="replace")
    if suffix in (".docx",):
        return _extract_docx(path)
    if suffix == ".pdf":
        return _extract_pdf(path)
    # Legacy .doc via LibreOffice or textutil
    for cmd in _doc_converters(path):
        try:
            result = subprocess.run(cmd, check=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if result.returncode == 0 and result.stdout:
                return result.stdout.decode("utf-8", errors="replace")
        except FileNotFoundError:
            continue
    raise ValueError(
        "Could not extract text from this file type. "
        "Please convert to .txt, .docx, or .pdf first."
    )


def _extract_docx(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml")
    root = ET.fromstring(xml)
    chunks: list[str] = []
    for elem in root.iter():
        if elem.tag.endswith("}t") and elem.text:
            chunks.append(elem.text)
        elif elem.tag.endswith("}p"):
            chunks.append("\n")
    return " ".join(chunks).replace(" \n ", "\n")


def _extract_pdf(path: Path) -> str:
    # Try pdfplumber first (best quality), then pypdf, then pdfminer
    try:
        import pdfplumber  # type: ignore
        with pdfplumber.open(path) as pdf:
            return "\n".join(page.extract_text() or "" for page in pdf.pages)
    except ImportError:
        pass
    try:
        from pypdf import PdfReader  # type: ignore
        reader = PdfReader(str(path))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except ImportError:
        pass
    raise ValueError(
        "PDF extraction requires pdfplumber or pypdf. "
        "Install with: pip install pdfplumber"
    )


def _doc_converters(path: Path) -> list[list[str]]:
    cmds = []
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        cmds.append([soffice, "--headless", "--convert-to", "txt:Text",
                     "--outdir", str(path.parent), str(path)])
    textutil = shutil.which("textutil")
    if textutil:
        cmds.append([textutil, "-convert", "txt", "-stdout", str(path)])
    return cmds


# ── Claude API parsing ─────────────────────────────────────────────────────────

_CLAUDE_EXTRACTION_PROMPT = """\
You are a scientific metadata extraction assistant. Extract structured metadata from the eDNA research report below and return ONLY valid JSON (no prose, no markdown fences).

Return this JSON schema (use null for fields not found):
{
  "project_title": "...",
  "institution": "...",
  "institution_ror": "https://ror.org/...",
  "project_contact": "...",
  "funder": "...",
  "start_date": "ISO8601 or descriptive",
  "end_date": "ISO8601 or descriptive",
  "project_cost": "...",
  "license": "...",
  "access_rights": "...",
  "bibliography": "full citation or DOI",
  "code_repo": "URL or null",
  "study_factor": "comma-separated experimental factors",
  "sterilise_method": "...",
  "env_broad_scale": "ENVO term if possible",
  "env_local_scale": "ENVO term if possible",
  "env_medium": "ENVO term if possible",
  "samp_collect_method": "...",
  "samp_size": "number",
  "samp_size_unit": "mL or L",
  "filter_pore": "pore size in µm",
  "filter_diameter": "mm",
  "filter_material": "...",
  "samp_store_temp": "degrees C or description",
  "samp_store_sol": "...",
  "ext_kit": "commercial kit name and catalogue number",
  "ext_modifications": "any protocol modifications",
  "ext_lysis": "lysis method",
  "ext_sep": "separation method",
  "conc_method": "quantification method",
  "normalisation": "normalisation description",
  "barcoding_pcr_appr": "one-step or two-step PCR",
  "platform": "ILLUMINA or OXFORD_NANOPORE or PacBio etc",
  "instrument": "exact instrument model",
  "lib_layout": "paired or single",
  "adapter_fwd": "sequence",
  "adapter_rev": "sequence",
  "pcr2_mm": "master mix for step 2",
  "pcr2_rxn_vol": "µL",
  "pcr2_dna_vol": "µL",
  "pcr2_cond": "thermocycler conditions",
  "seq_location": "institution name",
  "phix_perc": "percentage",
  "trim_method": "tool and version",
  "trim_param": "parameters",
  "demux_tool": "tool",
  "merge_tool": "tool",
  "error_rate_type": "type",
  "error_rate_cutoff": "value",
  "chimera_method": "method",
  "chimera_param": "parameters",
  "clust_tool": "ASV/OTU tool",
  "screen_contam": "method",
  "otu_raw_desc": "description of raw ASV/OTU table",
  "otu_final_desc": "description of final ASV/OTU table",
  "bioinfo_extra": "any additional bioinformatics notes",
  "assays": [
    {
      "name": "short identifier e.g. 16S or MiFish12S",
      "target_gene": "gene region",
      "target_subfragment": "hypervariable region if applicable",
      "amplicon_size": "bp",
      "primer_fwd": "sequence 5'→3'",
      "primer_rev": "sequence 5'→3'",
      "primer_name_fwd": "primer name",
      "primer_name_rev": "primer name",
      "primer_ref_fwd": "citation/DOI",
      "primer_ref_rev": "citation/DOI",
      "commercial_mm": "master mix product name",
      "custom_mm": "recipe if custom",
      "rxn_vol": "µL",
      "dna_vol": "µL",
      "pcr_rep": "number of PCR replicates",
      "annealing_temp": "°C",
      "pcr_cycles": "number",
      "pcr_cond": "full thermocycling description",
      "tax_db": "reference database name and version",
      "tax_method": "classification method",
      "min_len": "minimum read length cutoff",
      "seq_kit": "sequencing kit for this assay",
      "rarefaction_depth": "reads per sample for rarefaction"
    }
  ],
  "sampling_locations": [
    {
      "name": "site name",
      "code": "short code",
      "lake_type": "...",
      "wfd_tp": "WFD TP class",
      "wfd_chlorophyll": "WFD chlorophyll class",
      "trophic_status": "Oligotrophic/Mesotrophic/Eutrophic",
      "latitude": "decimal degrees WGS84",
      "longitude": "decimal degrees WGS84"
    }
  ],
  "sampling_rounds": [
    {
      "code": "R1",
      "iso_date": "2025-07-14 or range",
      "verbatim": "original date description",
      "avg_vol_ml": "average volume filtered in mL"
    }
  ],
  "neg_controls": "description of negative controls used",
  "pos_controls": "description of positive controls used or null",
  "summary": "2-3 sentence summary of the project"
}

REPORT TEXT:
"""


def parse_with_claude(text: str, api_key: str) -> dict:
    """Call Claude API and return parsed metadata as a dict."""
    try:
        import anthropic  # type: ignore
    except ImportError:
        raise RuntimeError(
            "anthropic package not installed. Run: pip install anthropic"
        )
    client = anthropic.Anthropic(api_key=api_key)
    truncated = text[:MAX_REPORT_CHARS]
    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=4096,
        messages=[{
            "role": "user",
            "content": _CLAUDE_EXTRACTION_PROMPT + truncated
        }]
    )
    raw = message.content[0].text.strip()
    # Strip any accidental markdown fences
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw)


def facts_from_claude(data: dict) -> ReportFacts:
    """Map Claude JSON output to a ReportFacts object."""
    def s(key, default=MISS):
        v = data.get(key)
        return str(v) if v else default

    assays = []
    for a in data.get("assays") or []:
        assays.append(AssayInfo(
            name             = a.get("name") or "unknown_assay",
            target_gene      = a.get("target_gene") or MISS,
            target_subfragment=a.get("target_subfragment") or MISS,
            amplicon_size    = a.get("amplicon_size") or MISS,
            primer_fwd       = a.get("primer_fwd") or MISS,
            primer_rev       = a.get("primer_rev") or MISS,
            primer_name_fwd  = a.get("primer_name_fwd") or MISS,
            primer_name_rev  = a.get("primer_name_rev") or MISS,
            primer_ref_fwd   = a.get("primer_ref_fwd") or MISS,
            primer_ref_rev   = a.get("primer_ref_rev") or MISS,
            commercial_mm    = a.get("commercial_mm") or MISS,
            custom_mm        = a.get("custom_mm") or MISS,
            rxn_vol          = a.get("rxn_vol") or MISS,
            dna_vol          = a.get("dna_vol") or MISS,
            pcr_rep          = a.get("pcr_rep") or "1",
            annealing_temp   = a.get("annealing_temp") or MISS,
            pcr_cycles       = a.get("pcr_cycles") or MISS,
            pcr_cond         = a.get("pcr_cond") or MISS,
            tax_db           = a.get("tax_db") or MISS,
            tax_method       = a.get("tax_method") or MISS,
            min_len          = a.get("min_len") or MISS,
            seq_kit          = a.get("seq_kit") or MISS,
            rarefaction_depth= a.get("rarefaction_depth") or MISS,
        ))

    lakes = []
    for lk in data.get("sampling_locations") or []:
        lakes.append(Lake(
            name            = lk.get("name") or "unknown",
            code            = lk.get("code") or "UNK",
            lake_type       = lk.get("lake_type") or MISS,
            wfd_tp          = lk.get("wfd_tp") or MISS,
            wfd_chlorophyll = lk.get("wfd_chlorophyll") or MISS,
            trophic_status  = lk.get("trophic_status") or MISS,
            latitude        = lk.get("latitude") or MISS,
            longitude       = lk.get("longitude") or MISS,
        ))

    rounds = []
    for r in data.get("sampling_rounds") or []:
        rounds.append((
            r.get("code") or "R1",
            r.get("iso_date") or MISS,
            r.get("verbatim") or MISS,
            r.get("avg_vol_ml") or MISS,
        ))

    title = s("project_title", "eDNA metabarcoding project")
    return ReportFacts(
        project_title       = title,
        project_id          = slugify(title)[:40] or "edna_project",
        institution         = s("institution"),
        institution_id      = s("institution_ror"),
        project_contact     = s("project_contact"),
        funder              = s("funder"),
        start_date          = s("start_date"),
        end_date            = s("end_date"),
        project_cost        = s("project_cost"),
        license             = s("license"),
        access_rights       = s("access_rights"),
        bibliography        = s("bibliography"),
        code_repo           = s("code_repo"),
        sterilise_method    = s("sterilise_method"),
        env_broad_scale     = s("env_broad_scale"),
        env_local_scale     = s("env_local_scale"),
        env_medium          = s("env_medium"),
        samp_collect_method = s("samp_collect_method"),
        samp_size           = s("samp_size"),
        samp_size_unit      = s("samp_size_unit"),
        filter_pore         = s("filter_pore"),
        filter_diameter     = s("filter_diameter"),
        filter_material     = s("filter_material"),
        samp_store_temp     = s("samp_store_temp"),
        samp_store_sol      = s("samp_store_sol"),
        ext_kit             = s("ext_kit"),
        ext_modifications   = s("ext_modifications"),
        ext_lysis           = s("ext_lysis"),
        ext_sep             = s("ext_sep"),
        conc_method         = s("conc_method"),
        normalisation       = s("normalisation"),
        barcoding_pcr_appr  = s("barcoding_pcr_appr"),
        platform            = s("platform"),
        instrument          = s("instrument"),
        lib_layout          = s("lib_layout"),
        adapter_fwd         = s("adapter_fwd"),
        adapter_rev         = s("adapter_rev"),
        pcr2_mm             = s("pcr2_mm"),
        pcr2_rxn_vol        = s("pcr2_rxn_vol"),
        pcr2_dna_vol        = s("pcr2_dna_vol"),
        pcr2_cond           = s("pcr2_cond"),
        seq_location        = s("seq_location"),
        phix_perc           = s("phix_perc"),
        trim_method         = s("trim_method"),
        trim_param          = s("trim_param"),
        demux_tool          = s("demux_tool"),
        merge_tool          = s("merge_tool"),
        error_rate_type     = s("error_rate_type"),
        error_rate_cutoff   = s("error_rate_cutoff"),
        chimera_method      = s("chimera_method"),
        chimera_param       = s("chimera_param"),
        clust_tool          = s("clust_tool"),
        screen_contam       = s("screen_contam"),
        otu_raw_desc        = s("otu_raw_desc"),
        otu_final_desc      = s("otu_final_desc"),
        bioinfo_extra       = s("bioinfo_extra"),
        study_factor        = s("study_factor"),
        assays              = assays,
        lakes               = lakes,
        rounds              = rounds,
        report_summary      = data.get("summary") or "",
    )


# ── Regex fallback parser ──────────────────────────────────────────────────────

def parse_report_regex(text: str) -> ReportFacts:
    title = _find_after(text, r"Project title", r"Contractor|54\.|Total Defra")
    if not title:
        title = _find_sentence(text, r"Development of[^.\n]+") or "eDNA metabarcoding project"
    title = re.sub(r"\s+\d+\.$", "", _norm(title))
    institution = "UK Centre for Ecology & Hydrology (UKCEH)" if re.search(r"\bUKCEH\b", text, re.I) else MISS
    inst_id = "https://ror.org/04k8syz75" if "UKCEH" in institution else MISS
    start = _find_after(text, r"start date", r"end date") or MISS
    end   = _find_after(text, r"end date", r"\f|Executive Summary") or MISS
    cost  = _find_after(text, r"Total Defra project costs\s*£", r"\(agreed|Project:") or MISS

    assays_raw = []
    for a in ["16S", "18S", "rbcL", "COI", "12S"]:
        if re.search(rf"\b{re.escape(a)}\b", text, re.I):
            assays_raw.append(a)

    assays = [AssayInfo(name=a) for a in (assays_raw or ["metabarcoding"])]

    lakes = [Lake(*row) for row in _KNOWN_LAKES if re.search(re.escape(row[0]), text, re.I)]
    rounds = _parse_rounds(text)
    summary = _parse_summary(text)

    return ReportFacts(
        project_title  = _norm(title),
        project_id     = slugify(title)[:40] or "edna_project",
        institution    = institution,
        institution_id = inst_id,
        start_date     = _norm(start),
        end_date       = _norm(end),
        project_cost   = _norm(cost),
        assays         = assays,
        lakes          = lakes,
        rounds         = rounds,
        report_summary = summary,
    )


_KNOWN_LAKES = [
    ("Bassenthwaite",          "BASS", "MAS",  "Good",     "High/Good",  "Mesotrophic"),
    ("Buttermere",             "BUTT", "LAD",  "Good",     "High",       "Oligo/ultraoligotrophic"),
    ("Derwentwater",           "DERW", "LAS",  "High",     "Good",       "Oligo/mesotrophic"),
    ("Elterwater",             "ELTE", "LAVS", "Moderate", "Good",       "Eutrophic"),
    ("Ennerdale Water",        "ENNE", "LAD",  "Good",     "High",       "Oligo/ultraoligotrophic"),
    ("Esthwaite Water",        "ESTH", "MAS",  "Poor",     "Poor",       "Eutrophic"),
    ("Loughrigg Tarn",         "LOUG", "MAS",  "Good/Moderate", "Moderate", "Meso/eutrophic"),
    ("Wastwater",              "WAST", "LAD",  "Good",     "High",       "Oligo/ultraoligotrophic"),
    ("Windermere North Basin", "NBAS", "MAD",  "Moderate", "Moderate",   "Mesotrophic"),
    ("Windermere South Basin", "SBAS", "MAD",  "Moderate", "Good",       "Mesotrophic"),
]


def _find_after(text, label, until):
    m = re.search(label + r"\s*(.*?)\s*(?=" + until + r")", text, re.I | re.S)
    if not m: return ""
    return _norm(re.sub(r"FORMTEXT|FORMCHECKBOX|[ -​]+", " ", m.group(1)))


def _find_sentence(text, pat):
    m = re.search(pat, text, re.I)
    return _norm(m.group(0)) if m else ""


def _norm(text):
    return re.sub(r"\s+", " ", text).strip()


def slugify(value):
    return re.sub(r"_+", "_", re.sub(r"[^A-Za-z0-9]+", "_", value.lower())).strip("_")


def _parse_rounds(text):
    rounds = []
    if re.search(r"14(?:th)?\s+July\s+2025|July\s+2025", text, re.I):
        rounds.append(("R1", "2025-07-14/2025-07-20", "week commencing 14 July 2025", "720"))
    if re.search(r"13(?:th)?\s+October\s+2025|October\s+2025", text, re.I):
        rounds.append(("R2", "2025-10-13/2025-10-19", "week commencing 13 October 2025", "1207"))
    return rounds or [("R1", MISS, MISS, MISS)]


def _parse_summary(text):
    m = re.search(r"Executive Summary\s*(.*?)\s*Project Report", text, re.I | re.S)
    return _norm(m.group(1))[:1800] if m else _norm(text)[:1800]


# ── Excel helpers ──────────────────────────────────────────────────────────────

def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def num_to_col(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def cell_col(ref):
    return re.match(r"([A-Z]+)", ref).group(1)


def get_value(cell, shared):
    v = cell.find("m:v", NS)
    if v is not None:
        raw = v.text or ""
        return shared[int(raw)] if cell.attrib.get("t") == "s" else raw
    inline = cell.find("m:is", NS)
    if inline is not None:
        return "".join(t.text or "" for t in inline.iter(f"{{{MAIN}}}t"))
    return ""


def set_cell(row_el, col_idx, row_idx, value):
    ref = f"{num_to_col(col_idx)}{row_idx}"
    cells = row_el.findall("m:c", NS)
    target = next((c for c in cells if c.attrib.get("r") == ref), None)
    if target is None:
        target = ET.Element(f"{{{MAIN}}}c", {"r": ref})
        inserted = False
        for idx, c in enumerate(cells):
            if col_to_num(cell_col(c.attrib.get("r", "A1"))) > col_idx:
                row_el.insert(idx, target)
                inserted = True
                break
        if not inserted:
            row_el.append(target)
    for child in list(target):
        target.remove(child)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        target.attrib.pop("t", None)
        ET.SubElement(target, f"{{{MAIN}}}v").text = str(value)
    else:
        target.attrib["t"] = "inlineStr"
        t_el = ET.SubElement(ET.SubElement(target, f"{{{MAIN}}}is"), f"{{{MAIN}}}t")
        t_el.text = str(value)


def get_or_create_row(sheet_data, row_idx):
    for row in sheet_data.findall("m:row", NS):
        if int(row.attrib.get("r", "0")) == row_idx:
            return row
    new_row = ET.Element(f"{{{MAIN}}}row", {"r": str(row_idx)})
    for idx, row in enumerate(sheet_data.findall("m:row", NS)):
        if int(row.attrib.get("r", "0")) > row_idx:
            sheet_data.insert(idx, new_row)
            return new_row
    sheet_data.append(new_row)
    return new_row


def trim_rows(sheet_data, keep_through):
    for row in list(sheet_data.findall("m:row", NS)):
        if int(row.attrib.get("r", "0")) > keep_through:
            sheet_data.remove(row)


def update_dimension(root, max_col, max_row):
    dim = root.find("m:dimension", NS)
    if dim is not None:
        dim.attrib["ref"] = f"A1:{num_to_col(max_col)}{max_row}"


def sheet_headers(root, shared, header_row):
    sheet_data = root.find("m:sheetData", NS)
    row = get_or_create_row(sheet_data, header_row)
    return {get_value(c, shared): col_to_num(cell_col(c.attrib["r"]))
            for c in row.findall("m:c", NS) if get_value(c, shared)}


# ── xlsx repair helpers ────────────────────────────────────────────────────────

def should_skip(name):
    return (name.startswith("xl/comments")
            or name.startswith("xl/drawings/vmlDrawing")
            or name.startswith("xl/drawings/drawing"))


def clean_entry(name, data, package_names):
    if name == "[Content_Types].xml":
        root = ET.fromstring(data)
        for o in list(root):
            pn = o.attrib.get("PartName", "")
            if any(pn.startswith(p) for p in ("/xl/comments", "/xl/drawings/vmlDrawing", "/xl/drawings/drawing")):
                root.remove(o)
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)
    if name.startswith("xl/worksheets/_rels/") and name.endswith(".rels"):
        root = ET.fromstring(data)
        for rel in list(root):
            t = rel.attrib.get("Type", "")
            if t.endswith(("/drawing", "/vmlDrawing", "/comments")):
                root.remove(rel)
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)
    if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
        root = ET.fromstring(data)
        for tag in ("drawing", "legacyDrawing", "extLst"):
            for el in list(root.findall(f"{{{MAIN}}}{tag}")):
                root.remove(el)
        for attr in [a for a in root.attrib if a.endswith("}Ignorable")]:
            del root.attrib[attr]
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return data


# ── workbook fill ──────────────────────────────────────────────────────────────

def fill_workbook(template_path: Path, facts: ReportFacts,
                  report_filename: str, output_path: Path) -> dict:
    with zipfile.ZipFile(template_path) as zin:
        pkg_names = set(zin.namelist())
        shared_root = ET.fromstring(zin.read("xl/sharedStrings.xml"))
        shared = ["".join(t.text or "" for t in si.iter(f"{{{MAIN}}}t"))
                  for si in shared_root.findall("m:si", NS)]
        workbook = ET.fromstring(zin.read("xl/workbook.xml"))
        rels_root = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root}
        sheet_paths = {
            sh.attrib["name"]: "xl/" + relmap[sh.attrib[f"{{{REL}}}id"]]
            for sh in workbook.find("m:sheets", NS)
        }

        updated: dict[str, bytes] = {}
        _fill_project_sheet(zin, shared, sheet_paths, facts, report_filename, updated)
        sample_records = _fill_sample_sheet(zin, shared, sheet_paths, facts, updated)
        _fill_experiment_sheet(zin, shared, sheet_paths, facts, sample_records, updated)
        _fill_readme(zin, shared, sheet_paths, facts, report_filename, updated)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if should_skip(item.filename):
                    continue
                data = updated.get(item.filename, zin.read(item.filename))
                data = clean_entry(item.filename, data, pkg_names)
                zout.writestr(item, data)

    return {
        "projectTitle":   facts.project_title,
        "projectId":      facts.project_id,
        "institution":    facts.institution,
        "lakes":          len(facts.lakes),
        "rounds":         len(facts.rounds),
        "assays":         [a.name for a in facts.assays],
        "sampleRows":     len(sample_records),
        "experimentRows": len(sample_records) * len(facts.assays),
    }


def _fill_project_sheet(zin, shared, sheet_paths, facts, report_filename, updated):
    assay_names = " | ".join(a.name for a in facts.assays)
    fwd_primers  = " | ".join(a.primer_fwd  for a in facts.assays)
    rev_primers  = " | ".join(a.primer_rev  for a in facts.assays)
    fwd_names    = " | ".join(a.primer_name_fwd for a in facts.assays)
    rev_names    = " | ".join(a.primer_name_rev for a in facts.assays)
    fwd_refs     = " | ".join(a.primer_ref_fwd  for a in facts.assays)
    rev_refs     = " | ".join(a.primer_ref_rev  for a in facts.assays)
    mm_vals      = " | ".join(a.commercial_mm for a in facts.assays)
    rxn_vols     = " | ".join(a.rxn_vol      for a in facts.assays)
    dna_vols     = " | ".join(a.dna_vol      for a in facts.assays)
    tax_dbs      = " | ".join(a.tax_db       for a in facts.assays)
    tax_methods  = " | ".join(a.tax_method   for a in facts.assays)
    min_lens     = " | ".join(a.min_len      for a in facts.assays)
    seq_kits     = " | ".join(a.seq_kit      for a in facts.assays)
    otu_finals   = " | ".join(a.rarefaction_depth for a in facts.assays)

    values = {
        "recordedBy":                  facts.institution,
        "recordedByID":                facts.institution_id,
        "project_contact":             facts.project_contact,
        "institution":                 facts.institution,
        "institutionID":               facts.institution_id,
        "project_name":                facts.project_title,
        "project_id":                  facts.project_id,
        "study_factor":                facts.study_factor,
        "assay_type":                  "metabarcoding",
        "sterilise_method":            facts.sterilise_method,
        "neg_cont_0_1":                1,
        "pos_cont_0_1":                0,
        "checkls_ver":                 "FAIRe_checklist_v1.0.2",
        "mod_date":                    time.strftime("%Y-%m-%d"),
        "license":                     facts.license,
        "rightsHolder":                facts.institution,
        "accessRights":                facts.access_rights,
        "bibliographicCitation":       facts.bibliography or f"Source report: {report_filename}",
        "code_repo":                   facts.code_repo,
        "biological_rep":              facts.samp_collect_method,
        "pcr_0_1":                     1,
        "assay_name":                  assay_names,
        "assay_validation":            MISS,
        "targetTaxonomicAssay":        " | ".join(a.target_gene for a in facts.assays),
        "targetTaxonomicScope":        " | ".join(a.target_subfragment for a in facts.assays),
        "target_gene":                 " | ".join(a.target_gene for a in facts.assays),
        "ampliconSize":                " | ".join(a.amplicon_size for a in facts.assays),
        "pcr_primer_forward":          fwd_primers,
        "pcr_primer_reverse":          rev_primers,
        "pcr_primer_name_forward":     fwd_names,
        "pcr_primer_name_reverse":     rev_names,
        "pcr_primer_reference_forward":fwd_refs,
        "pcr_primer_reference_reverse":rev_refs,
        "probeReporter":               NA,
        "probeQuencher":               NA,
        "probe_seq":                   NA,
        "probe_ref":                   NA,
        "probe_conc":                  NA,
        "commercial_mm":               mm_vals,
        "custom_mm":                   " | ".join(a.custom_mm for a in facts.assays),
        "amplificationReactionVolume": rxn_vols,
        "pcr_dna_vol":                 dna_vols,
        "pcr_rep":                     " | ".join(a.pcr_rep for a in facts.assays),
        "nucl_acid_amp":               "PCR",
        "pcr_cond":                    " | ".join(a.pcr_cond for a in facts.assays),
        "annealingTemp":               " | ".join(a.annealing_temp for a in facts.assays),
        "pcr_cycles":                  " | ".join(a.pcr_cycles for a in facts.assays),
        "amp_vis_method":              NA,
        "detection_criteria":          NA,
        "lod_method":                  NA,
        "pcr_assay_lod":               NA,
        "pcr_assay_lod_unit":          NA,
        "loq_method":                  NA,
        "pcr_assay_loq":               NA,
        "pcr_assay_loq_unit":          NA,
        "std_type":                    NA,
        "std_seq":                     NA,
        "thresholdQuantificationCycle":NA,
        "barcoding_pcr_appr":          facts.barcoding_pcr_appr,
        "pcr2_commercial_mm":          facts.pcr2_mm,
        "pcr2_amplificationReactionVolume": facts.pcr2_rxn_vol,
        "pcr2_dna_vol":                facts.pcr2_dna_vol,
        "pcr2_cond":                   facts.pcr2_cond,
        "sequencing_location":         facts.seq_location,
        "platform":                    facts.platform,
        "instrument":                  facts.instrument,
        "seq_kit":                     seq_kits,
        "lib_layout":                  facts.lib_layout,
        "adapter_forward":             facts.adapter_fwd,
        "adapter_reverse":             facts.adapter_rev,
        "sop_bioinformatics":          facts.trim_method,
        "trim_method":                 facts.trim_method,
        "trim_param":                  facts.trim_param,
        "demux_tool":                  facts.demux_tool,
        "merge_tool":                  facts.merge_tool,
        "min_len_cutoff":              min_lens,
        "min_len_tool":                MISS,
        "error_rate_tool":             MISS,
        "error_rate_type":             facts.error_rate_type,
        "error_rate_cutoff":           facts.error_rate_cutoff,
        "chimera_check_method":        facts.chimera_method,
        "chimera_check_param":         facts.chimera_param,
        "otu_clust_tool":              facts.clust_tool,
        "otu_db":                      tax_dbs,
        "tax_assign_cat":              tax_methods,
        "otu_seq_comp_appr":           tax_methods,
        "screen_contam_method":        facts.screen_contam,
        "otu_raw_description":         facts.otu_raw_desc,
        "otu_final_description":       otu_finals if otu_finals.strip("| ") else facts.otu_final_desc,
        "bioinfo_method_additional":   facts.bioinfo_extra,
    }
    path = sheet_paths["projectMetadata"]
    root = ET.fromstring(zin.read(path))
    sheet_data = root.find("m:sheetData", NS)
    for row in sheet_data.findall("m:row", NS):
        ri = int(row.attrib["r"])
        if ri < 2: continue
        term = next((get_value(c, shared) for c in row.findall("m:c", NS)
                     if c.attrib.get("r") == f"C{ri}"), "")
        if term in values:
            set_cell(row, 4, ri, values[term])
    update_dimension(root, 4, 143)
    updated[path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _fill_sample_sheet(zin, shared, sheet_paths, facts, updated):
    path = sheet_paths["sampleMetadata"]
    root = ET.fromstring(zin.read(path))
    sheet_data = root.find("m:sheetData", NS)
    trim_rows(sheet_data, 3)
    headers = sheet_headers(root, shared, 3)
    row_idx = 4
    records = []
    assay_str = " | ".join(a.name for a in facts.assays)

    def add(data):
        nonlocal row_idx
        row = get_or_create_row(sheet_data, row_idx)
        for k, v in data.items():
            if k in headers:
                set_cell(row, headers[k], row_idx, v)
        row_idx += 1

    for lake in (facts.lakes or [Lake("sampling location", "SITE")]):
        for round_code, event_date, verbatim_date, avg_vol in (facts.rounds or [("R1", MISS, MISS, MISS)]):
            for i in range(1, 11):
                sid = f"{lake.code}_{round_code}_S{i:02d}"
                records.append((sid, lake.code, round_code, False))
                add(_sample_row(sid, lake, round_code, event_date, verbatim_date, avg_vol, facts, assay_str, False))
            bk = f"{lake.code}_{round_code}_FIELD_BLANK"
            records.append((bk, lake.code, round_code, True))
            add(_sample_row(bk, lake, round_code, event_date, verbatim_date, avg_vol, facts, assay_str, True))

    if headers:
        update_dimension(root, max(headers.values()), row_idx - 1)
    updated[path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return records


def _sample_row(sid, lake, round_code, event_date, verbatim_date, avg_vol, facts, assay_str, blank):
    filt_name = (
        f"{facts.filter_diameter} mm {facts.filter_pore} µm {facts.filter_material} filter"
        if facts.filter_material != MISS else MISS
    )
    base = {
        "samp_name": sid, "materialSampleID": sid,
        "decimalLatitude": lake.latitude, "decimalLongitude": lake.longitude,
        "verbatimSRS": "WGS84", "verbatimCoordinateSystem": "decimal degrees",
        "geo_loc_name": f"United Kingdom: Cumbria, {lake.name}",
        "eventDate": event_date, "verbatimEventDate": verbatim_date,
        "eventDurationValue": "7 d" if "/" in str(event_date) else MISS,
        "env_broad_scale": facts.env_broad_scale,
        "env_local_scale": facts.env_local_scale,
        "samp_store_temp": facts.samp_store_temp,
        "samp_store_sol":  facts.samp_store_sol,
        "filter_passive_active_0_1": 1,
        "filter_diameter": facts.filter_diameter,
        "filter_material": facts.filter_material,
        "filter_name": filt_name,
        "size_frac": facts.filter_pore,
        "nucl_acid_ext_kit": facts.ext_kit,
        "nucl_acid_ext_modify": facts.ext_modifications,
        "nucl_acid_ext_lysis": facts.ext_lysis,
        "nucl_acid_ext_sep": facts.ext_sep,
        "concentration_method": facts.conc_method,
        "dna_cleanup_0_1": 1,
        "pool_dna_num": 1,
        "assay_name": assay_str,
    }
    if blank:
        base.update({
            "samp_category": "negative control",
            "neg_cont_type": "field negative", "pos_cont_type": NA_CONTROL,
            "env_medium": "purified water",
            "habitat_natural_artificial_0_1": NA_CONTROL,
            "samp_collect_method": "purified water field blank",
            "samp_size": 2, "samp_size_unit": "L",
            "samp_vol_we_dna_ext": MISS, "samp_vol_we_dna_ext_unit": "mL",
            "detected_notDetected": NA_CONTROL,
        })
    else:
        base.update({
            "samp_category": "sample",
            "neg_cont_type": NA_SAMPLE, "pos_cont_type": NA_SAMPLE,
            "rel_cont_id": f"{lake.code}_{round_code}_FIELD_BLANK",
            "env_medium": facts.env_medium,
            "habitat_natural_artificial_0_1": 1,
            "samp_collect_method": facts.samp_collect_method,
            "samp_size": facts.samp_size, "samp_size_unit": facts.samp_size_unit,
            "samp_mat_process": f"vacuum filtration; avg {avg_vol} mL filtered",
            "samp_vol_we_dna_ext": avg_vol, "samp_vol_we_dna_ext_unit": "mL",
            "detected_notDetected": NA,
            "minimumDepthInMeters": 0,
            "org_matter": f"Lake type {lake.lake_type}; WFD TP {lake.wfd_tp}; trophic: {lake.trophic_status}",
        })
    return base


def _fill_experiment_sheet(zin, shared, sheet_paths, facts, sample_records, updated):
    path = sheet_paths.get("experimentRunMetadata")
    if not path: return
    root = ET.fromstring(zin.read(path))
    sheet_data = root.find("m:sheetData", NS)
    trim_rows(sheet_data, 3)
    headers = sheet_headers(root, shared, 3)
    row_idx = 4
    for sid, _code, _round, _blank in sample_records:
        for assay in facts.assays:
            run_id = "missing: not provided"
            row = get_or_create_row(sheet_data, row_idx)
            data = {
                "samp_name": sid, "assay_name": assay.name,
                "lib_id": f"{sid}_{assay.name}",
                "seq_run_id": run_id,
                "lib_conc": MISS, "lib_conc_unit": "nM",
                "lib_conc_meth": facts.conc_method,
                "phix_perc": facts.phix_perc,
                "mid_forward": MISS, "mid_reverse": MISS,
                "filename": MISS, "filename2": MISS,
                "checksum_filename": MISS, "checksum_filename2": MISS,
                "associatedSequences": MISS,
                "input_read_count": MISS, "output_read_count": MISS,
                "output_otu_num": MISS, "otu_num_tax_assigned": MISS,
            }
            for k, v in data.items():
                if k in headers:
                    set_cell(row, headers[k], row_idx, v)
            row_idx += 1
    if headers:
        update_dimension(root, max(headers.values()), row_idx - 1)
    updated[path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _fill_readme(zin, shared, sheet_paths, facts, report_filename, updated):
    path = sheet_paths.get("README")
    if not path: return
    root = ET.fromstring(zin.read(path))
    sheet_data = root.find("m:sheetData", NS)
    rows = sheet_data.findall("m:row", NS)
    start = (max(int(r.attrib.get("r", "0")) for r in rows) + 2) if rows else 2
    notes = [
        "Filled metadata notes:",
        f"Source report: {report_filename}",
        f"Filled sheets: projectMetadata, sampleMetadata, experimentRunMetadata",
        f"Detected project: {facts.project_title}",
        f"Assays detected: {', '.join(a.name for a in facts.assays)}",
        f"Lakes / sites detected: {len(facts.lakes)}",
        "Missing mandatory values use FAIRe INSDC missing-value wording.",
        "Coordinates, library IDs and per-sample measurements must be added manually.",
        "Generated by FAIRe Metadata Filler – https://fair-edna.github.io",
    ]
    for idx, note in enumerate(notes, start):
        row = get_or_create_row(sheet_data, idx)
        set_cell(row, 1, idx, note)
    update_dimension(root, 1, start + len(notes) - 1)
    updated[path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


# ── HTTP server ───────────────────────────────────────────────────────────────

def safe_name(name):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("_") or "file"


def parse_multipart(content_type, body):
    m = re.search(r"boundary=(?P<b>[^;]+)", content_type)
    if not m: raise ValueError("Missing multipart boundary.")
    boundary = m.group("b").strip().strip('"').encode()
    parts = {}
    for raw in body.split(b"--" + boundary):
        raw = raw.strip()
        if not raw or raw == b"--": continue
        if raw.endswith(b"--"): raw = raw[:-2].rstrip()
        head_blob, _, payload = raw.partition(b"\r\n\r\n")
        if not payload: continue
        headers = head_blob.decode("utf-8", errors="replace")
        disp = re.search(r'Content-Disposition: form-data;\s*(.*)', headers, re.I)
        if not disp: continue
        nm = re.search(r'name="([^"]+)"', disp.group(1))
        fn = re.search(r'filename="([^"]*)"', disp.group(1))
        if not nm: continue
        parts[nm.group(1)] = (fn.group(1) if fn else nm.group(1), payload.rstrip(b"\r\n"))
    return parts


class Handler(BaseHTTPRequestHandler):

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self):
        path = unquote(self.path.split("?", 1)[0])
        if path == "/":
            self._serve(STATIC_DIR / "index.html", "text/html; charset=utf-8")
        elif path.startswith("/static/"):
            tgt = STATIC_DIR / path.removeprefix("/static/")
            ct = ("text/css" if tgt.suffix == ".css" else
                  "application/javascript" if tgt.suffix == ".js" else
                  "text/html; charset=utf-8")
            self._serve(tgt, ct)
        elif path.startswith("/download/"):
            tgt = OUTPUT_DIR / safe_name(path.removeprefix("/download/"))
            self._serve(tgt, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", attachment=True)
        else:
            self.send_error(404)

    do_HEAD = do_GET

    def do_POST(self):
        if self.path != "/api/fill":
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            form = parse_multipart(self.headers.get("Content-Type", ""), self.rfile.read(length))
            if "report" not in form or "template" not in form:
                raise ValueError("Please upload both a report file and the FAIRe template workbook.")

            job_id = uuid.uuid4().hex[:10]
            report_fn, report_bytes = form["report"]
            template_fn, template_bytes = form["template"]
            api_key_field = form.get("apiKey", ("", b""))
            api_key = api_key_field[1].decode("utf-8", errors="replace").strip() if api_key_field[1] else ""
            # Also accept env variable
            if not api_key:
                api_key = os.environ.get("ANTHROPIC_API_KEY", "")

            rp = UPLOAD_DIR / f"{job_id}_{safe_name(Path(report_fn or 'report').name)}"
            tp = UPLOAD_DIR / f"{job_id}_{safe_name(Path(template_fn or 'template.xlsx').name)}"
            rp.write_bytes(report_bytes)
            tp.write_bytes(template_bytes)

            text = extract_report_text(rp)
            parsing_method = "regex"
            if api_key:
                try:
                    claude_data = parse_with_claude(text, api_key)
                    facts = facts_from_claude(claude_data)
                    parsing_method = "claude-ai"
                except Exception as e:
                    print(f"Claude parsing failed ({e}), falling back to regex.")
                    facts = parse_report_regex(text)
            else:
                facts = parse_report_regex(text)

            out_name = f"FAIRe_filled_{facts.project_id}_{job_id}.xlsx"
            out_path = OUTPUT_DIR / out_name
            summary = fill_workbook(tp, facts, rp.name, out_path)
            summary["downloadUrl"]    = f"/download/{out_name}"
            summary["outputName"]     = out_name
            summary["parsingMethod"]  = parsing_method

            self._json(summary)
        except Exception as exc:
            self._json({"error": str(exc)}, 500)

    def _serve(self, target, content_type, attachment=False):
        if not target.is_file():
            self.send_error(404)
            return
        data = target.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self._cors()
        if attachment:
            self.send_header("Content-Disposition", f'attachment; filename="{target.name}"')
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(data)

    def _json(self, payload, status=200):
        data = json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(data)))
        self._cors()
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, fmt, *args):
        print(f"{self.address_string()} {html.escape(fmt % args)}")


def main():
    host = os.environ.get("FAIRE_TOOL_HOST", "127.0.0.1")
    port = int(os.environ.get("FAIRE_TOOL_PORT", "8787"))
    print(f"FAIRe Metadata Filler at http://{host}:{port}")
    if os.environ.get("ANTHROPIC_API_KEY"):
        print("  Claude AI parsing: ENABLED (ANTHROPIC_API_KEY is set)")
    else:
        print("  Claude AI parsing: DISABLED (set ANTHROPIC_API_KEY to enable)")
    ThreadingHTTPServer((host, port), Handler).serve_forever()


if __name__ == "__main__":
    main()
