# FAIRe Metadata Filler

**Automatically fill the [FAIRe eDNA metadata template](https://fair-edna.github.io/download.html) from a project report.**

**[▶ Open the tool in your browser](https://ukceh-molecularecology.github.io/fair-edna-filler/)** · [![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/UKCEH-MolecularEcology/fair-edna-filler/blob/main/FAIRe_Metadata_Filler.ipynb)

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
- After filling, the page lists **"Check these by hand"** notes (for example inconsistencies spotted in the report).
- The rest of the template — dropdowns, term notes, formatting, other sheets — is left exactly as it was.

---

## Inputs & outputs

| Input | Description |
|---|---|
| **Report file** | `.docx`, `.pdf`, `.doc`, or `.txt` |
| **FAIRe template** *(optional)* | The official FAIRe FULL template is built in ([`template/FAIRe_checklist_v1.0.2_FULLtemplate.xlsx`](template/FAIRe_checklist_v1.0.2_FULLtemplate.xlsx)). Only choose a template to use a different one, e.g. made with FAIRe-ator. |
| **Claude API key** *(optional)* | Lets Claude read unusual report layouts — [get a key](https://console.anthropic.com) (paid, usage-based) |

**Output:** A filled `.xlsx` workbook ready to review in Excel and upload to the FAIRe-fier validator.

---

## How to use it

Pick whichever suits you — all three give the same result. Works on **macOS, Windows and Linux**.

### 1. In your browser — no installation (recommended)

**Online:** open **https://ukceh-molecularecology.github.io/fair-edna-filler/**

**Offline copy:** green **Code** button → *Download ZIP* → extract → double-click **`index.html`**.

Then choose your report, click **Fill metadata workbook**, and download the result — the official FAIRe template is built in. (When opening the downloaded `index.html` directly, also choose the template file from the `template` folder: browsers don't let a local page read neighbouring files.) Everything runs inside your browser — your files are never uploaded anywhere. The first visit downloads the tool's Python engine (about 10 MB), so you need an internet connection once. Works in current Chrome, Edge, Firefox and Safari.

### 2. Google Colab

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/UKCEH-MolecularEcology/fair-edna-filler/blob/main/FAIRe_Metadata_Filler.ipynb)

Open the notebook, choose *Runtime → Run all*, upload your report, and download the result. The official FAIRe FULL template is fetched automatically (tick a box to upload your own, e.g. one made with FAIRe-ator). Needs a Google account.

### 3. Python (for coders)

```bash
git clone https://github.com/UKCEH-MolecularEcology/fair-edna-filler.git
cd fair-edna-filler
pip install -r requirements.txt        # pdfplumber (PDF support), anthropic (optional Claude extraction)
python3 server.py                      # local web app at http://127.0.0.1:8787
```

Or use it from your own scripts:

```python
import server
meta = server.parse_report_regex(server.extract_report_text("report.doc"))
summary = server.fill_workbook(server.TEMPLATE_PATH, meta, "report.doc", "filled.xlsx")   # bundled template
print(summary["notes"])                # things to check by hand
```

| Environment variable | Default | Description |
|---|---|---|
| `ANTHROPIC_API_KEY` | *(none)* | Claude API key for AI-assisted extraction |
| `FAIRE_CLAUDE_MODEL` | `claude-haiku-5-5` | Claude model used for extraction |
| `FAIRE_TOOL_HOST` | `127.0.0.1` | Bind address (`0.0.0.0` to serve on a network) |
| `FAIRE_TOOL_PORT` | `8787` | Port |

---

## AI vs rule-based parsing

| Mode | When | Quality |
|---|---|---|
| **Claude AI** | API key provided | Understands free-text, extracts any format |
| **Rule-based** | No API key | Generic patterns for typical eDNA methods text (works offline) |

---

## Supported file types

| Format | Support |
|---|---|
| `.doc` (Word 97–2003) | ✅ Always — read directly, no Word or LibreOffice needed |
| `.docx` | ✅ Always |
| `.txt` | ✅ Always |
| `.pdf` | ✅ Text-based PDFs (the Python version needs `pdfplumber`). Scanned PDFs need OCR first. |

---

## Project structure

```
fair-edna-filler/
├── index.html                   # Browser version (runs server.py in the browser via Pyodide)
├── FAIRe_Metadata_Filler.ipynb  # Google Colab notebook
├── server.py                    # All the logic: reading reports, extraction, filling, local web app
├── static/                      # Page used by `python3 server.py`
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
