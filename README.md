# FAIRe Metadata Filler

**Automatically fill the [FAIRe eDNA metadata template](https://fair-edna.github.io/download.html) from a project report.**

**[▶ Open the tool in your browser](https://ukceh-molecularecology.github.io/fair-edna-filler/)** · [![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/UKCEH-MolecularEcology/fair-edna-filler/blob/main/FAIRe_Metadata_Filler.ipynb)

**Your report stays private:** it is read on your own computer (in your browser or in Python) and is never uploaded anywhere.

Upload a Word document, PDF, or plain-text report and the tool extracts project information, sampling details, PCR/sequencing methods, and bioinformatics parameters — then writes them into the correct cells of the FAIRe Excel workbook. Fields that cannot be found are filled with approved [FAIRe missing-value wording](https://fair-edna.github.io/guidelines.html#missing-values) so the workbook is always valid for the [FAIRe-fier validator](https://shiny.csiro.au/FAIRe-fier/).

---

## What it fills

| Sheet | What gets populated |
|---|---|
| `projectMetadata` | Project, institution, controls, PCR (per assay), library prep, sequencing, bioinformatics, reference databases |
| `sampleMetadata` | One row per sample, plus field, extraction and PCR negative controls, with location, dates, filtration, storage and extraction details |
| `experimentRunMetadata` | One row per sample × assay, with a `lib_id` that links back to `sampleMetadata` and a sequencing-run ID |

The tool does **not** modify `taxaRaw`, `taxaFinal`, `ampData`, `stdData`, or `eLowQuantData` — those require your actual sequence data.

### How values are placed

- **Multi-assay projects** follow the layout of the official FAIRe examples: project-wide values go in `project_level` (column D), and anything that differs between assays (primers, PCR recipe, databases, sequencing kit…) goes in one column per assay (`assay1`, `assay2`, …).
- **Controlled vocabularies** are read from the template's own `Drop-down values` sheet; values outside the list are written as `other: …`, as FAIRe requires.
- **Nothing is guessed.** If the report doesn't state something (e.g. primer sequences in a table that is an image, exact coordinates), the cell gets the FAIRe wording `missing: not provided` (or `not applicable: control sample` for lab controls).
- **IDs are generated** from the report structure (e.g. `BASS_R1_01`, `BASS_R1_FB`, `EXTBLANK_01`, `PCRNEG_16S_01`, `run_20260219_600cycle`). Replace them with your lab's IDs before publishing.
- **You check every value before it is written.** The tool shows a review table: each value, the sentence from the report it came from, and an edit box. Correct anything that's wrong, empty a box if the report doesn't really say it, or fill in something it missed.
- After filling, the page lists **"Check these by hand"** notes (for example inconsistencies spotted in the report).
- The rest of the template — dropdowns, term notes, formatting, other sheets — is left exactly as it was.

---

## Inputs & outputs

| Input | Description |
|---|---|
| **Report file** | `.docx`, `.pdf`, `.doc`, or `.txt` |
| **FAIRe template** *(optional)* | The official FAIRe FULL template is built in ([`template/FAIRe_checklist_v1.0.2_FULLtemplate.xlsx`](template/FAIRe_checklist_v1.0.2_FULLtemplate.xlsx)). Only choose a template to use a different one, e.g. made with FAIRe-ator. |
| **Claude API key** *(optional, off by default)* | Only for reports that are **not confidential** — see [Privacy](#privacy-and-how-values-are-found) |

**Output:** A filled `.xlsx` workbook ready to upload to the FAIRe-fier validator. The Python version also writes `<name>_review.csv`, listing every value with the report sentence it came from.

---

## How to use it

Pick whichever suits you — all three give the same result. Works on **macOS, Windows and Linux**.

### 1. In your browser — no installation (recommended)

**Online:** open **https://ukceh-molecularecology.github.io/fair-edna-filler/**

**Offline copy:** green **Code** button → *Download ZIP* → extract → double-click **`index.html`**.

Then choose your report, click **Read report**, check the values in the review table, click **Fill workbook with these values**, and download the result — the official FAIRe template is built in. (When opening the downloaded `index.html` directly, also choose the template file from the `template` folder: browsers don't let a local page read neighbouring files.) Everything runs inside your browser — your files are never uploaded anywhere. The first visit downloads the tool's Python engine (about 10 MB), so you need an internet connection once. Works in current Chrome, Edge, Firefox and Safari.

### 2. Google Colab

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/UKCEH-MolecularEcology/fair-edna-filler/blob/main/FAIRe_Metadata_Filler.ipynb)

Open the notebook, run the cells in order: upload your report, check the review table (put corrections in the `edits` cell), then fill and download. The official FAIRe FULL template is fetched automatically (tick a box to upload your own, e.g. one made with FAIRe-ator). Needs a Google account.

### 3. Python (for coders)

```bash
git clone https://github.com/UKCEH-MolecularEcology/fair-edna-filler.git
cd fair-edna-filler
pip install -r requirements.txt        # pypdf (PDF support), anthropic (optional)

python3 server.py                      # open the tool in your browser, served from this computer
python3 server.py report.pdf           # fill from one report      -> outputs/report_FAIRe.xlsx + report_review.csv
python3 server.py uploads/             # every report in a folder  -> outputs/
python3 server.py report.doc -t my_template.xlsx -o results/
```

The `uploads/` folder is a convenient place to drop several reports for batch filling; its contents are never committed to git.

Or use it from your own scripts:

```python
import server
meta = server.parse_report_regex(server.extract_report_text("report.doc"))
for item in server.review_items(meta):  # each value with the sentence it came from
    print(item["id"], "=", item["value"], "|", item["evidence"])
server.apply_review(meta, {"project.project_name": "My corrected title"})   # optional corrections
summary = server.fill_workbook(server.TEMPLATE_PATH, meta, "report.doc", "filled.xlsx")   # bundled template
print(summary["notes"])                # things to check by hand
```

| Environment variable | Default | Description |
|---|---|---|
| `FAIRE_CLAUDE_MODEL` | `claude-haiku-5-5` | Claude model used for extraction |
| `FAIRE_TOOL_HOST` | `127.0.0.1` | Bind address (`0.0.0.0` to serve on a network) |
| `FAIRE_TOOL_PORT` | `8787` | Port |

---

## Privacy and how values are found

**By default nothing leaves your computer.** A built-in reader looks for the usual ways methods are written in eDNA reports and papers: primer sequences and names, PCR recipes, filters, kits, sequencers, bioinformatics tools and reference databases. It works out which assay each value belongs to. Where a report has a *Methods* section, it reads that section only, so numbers from the results or introduction are not picked up. References are ignored. Every value keeps the sentence it came from, and you see it in the review table.

**Claude (optional, off by default).** For reports that are *not confidential*, you can tick *"This report is not confidential"* and enter a Claude API key. The report text is then sent to Anthropic's API. Claude must quote the sentence behind each value, and any value whose quote is not in the report is dropped. Never use this for sensitive or unpublished reports.

**Other languages.** The built-in reader is written for English. Reports in other languages are detected and flagged; primer sequences, numbers, kit and tool names are still found, but expect more gaps to fill by hand.

**Targeted (qPCR / ddPCR) studies.** The tool is built for metabarcoding. If a report looks like a targeted study, `assay_type` is set to `targeted` and you are told which fields (probe, standard curve, LOD/LOQ) to fill by hand.

### How well does it work?

`bench/run_benchmark.py` runs the tool on a folder of reports and compares the result with the published FAIRe example workbooks for the same studies. On 14 public reports and papers, three of which have a published FAIRe workbook, the values it filled were mostly right: 61 right, 13 different from the published workbook. Most of those 13 are naming differences, e.g. "Qiagen DNeasy Blood & Tissue kit" vs "Qiagen Blood & Tissue Kit". The other 46 checked values were left for you to fill in, typically things a report doesn't state in words, or values in tables that are images. That is why the review step exists.

```bash
python3 bench/run_benchmark.py path/to/reports/   # writes bench/results.md and bench/values.csv
```

The reports themselves are never copied into the repository.

---

## Supported file types

| Format | Support |
|---|---|
| `.doc` (Word 97–2003) | ✅ Always — read directly, no Word or LibreOffice needed |
| `.docx` | ✅ Always |
| `.txt` | ✅ Always |
| `.pdf` | ✅ Text-based PDFs (the Python version needs `pypdf`). Scanned PDFs need OCR first. |

---

## Project structure

```
fair-edna-filler/
├── index.html                   # Browser version (runs server.py in the browser via Pyodide)
├── FAIRe_Metadata_Filler.ipynb  # Google Colab notebook
├── server.py                    # All the logic: reading reports, extraction, review, filling, command line
├── static/ukceh_logo.png        # Logo embedded into index.html
├── bench/                       # Accuracy benchmark against published FAIRe examples
├── template/                    # Bundled official FAIRe FULL template (kept current by CI)
├── tests/                       # End-to-end checks + a fictional sample report
├── scripts/                     # check_template_update.py (used by CI)
├── .github/workflows/           # Tests on every push; template check twice a year
├── uploads/                     # Working folder for batch processing (contents not tracked)
├── build_standalone.py          # Rebuilds index.html from server.py
├── standalone/shell.html        # Page layout used by build_standalone.py
└── requirements.txt
```

### Keeping the FAIRe template current (maintainers)

On **1 January and 1 July** the *Check for a new FAIRe template* workflow reads the [FAIRe download page](https://fair-edna.github.io/download.html). If it links to a different FULL template version, the workflow downloads it, runs `tests/test_tool.py` against it, and opens a **pull request** that switches the tool to the new version, with the test results in the description. Nothing changes until you merge it. To check right away: **Actions → Check for a new FAIRe template → Run workflow**.

One-time setup: **Settings → Actions → General → Workflow permissions** → tick *Allow GitHub Actions to create and approve pull requests*.

Run the checks yourself with `python3 tests/test_tool.py` (or `python3 tests/test_tool.py some_template.xlsx` to try another template).

### Publishing the browser version (maintainers)

In the GitHub repo: **Settings → Pages → Build and deployment → Deploy from a branch → `main` / `(root)`**. The page is then served at the address above and updates automatically on every push.

---

## Citation

If you use this tool, please also cite the FAIRe eDNA checklist and guidelines — see the project's [How to cite](https://fair-edna.github.io/citation.html) page.

---

## Contributing

Issues and pull requests welcome. Everything lives in `server.py`: the rule-based extractor (`extract`), the workbook writer (`fill_workbook`) and the Claude prompt (`CLAUDE_PROMPT`). After changing it, run `python3 build_standalone.py` so `index.html` uses the same code.

## Licence

MIT — © UK Centre for Ecology & Hydrology
