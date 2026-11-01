# FAIRe Metadata Filler

**Automatically fill the [FAIRe eDNA metadata template](https://fair-edna.github.io/download.html) from a project report.**

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/UKCEH-MolecularEcology/fair-edna-filler/blob/main/FAIRe_Metadata_Filler.ipynb)

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
| **FAIRe template** | FAIRe full template `.xlsx` — [download here](https://fair-edna.github.io/download.html) |
| **Claude API key** *(optional)* | Lets Claude read unusual report layouts — [get a key](https://console.anthropic.com) (paid, usage-based) |

**Output:** A filled `.xlsx` workbook ready to review in Excel and upload to the FAIRe-fier validator.

---

## Option A — Just open a web page (no installation)

Download the ZIP (green **Code** button → *Download ZIP*), extract it, and double-click **`index.html`**. It opens in your browser and does everything locally — your files are never uploaded. The first time, it downloads its Python engine (about 10 MB), so you need an internet connection once.

---

## Option A2 — Google Colab

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/UKCEH-MolecularEcology/fair-edna-filler/blob/main/FAIRe_Metadata_Filler.ipynb)

Run entirely in the cloud — no Python or terminal needed. Just open the notebook, run all cells, upload your files, and download the result.

---

## Option B — Desktop app (no terminal, runs in background)

> **One-time requirement:** Python 3.9+ must be installed. [Download from python.org](https://www.python.org/downloads/) — tick **"Add Python to PATH"** on Windows.

### macOS

1. Download the ZIP from the green **Code** button above → Extract.
2. Double-click **`FAIRe Metadata Filler.app`**.
   - First launch: macOS may show a security warning → right-click → Open → Open.
   - Dependencies install automatically (internet needed the first time, ~30 seconds).
3. The tool icon appears in your **menu bar** (top-right). Your browser opens automatically.
4. To reopen: click the menu-bar icon → *Open FAIRe Filler*.
5. To stop: click the icon → *Quit*.

### Windows

1. Download the ZIP → Extract.
2. Double-click **`start_windows_silent.vbs`**.
   - A tray icon appears in the **system tray** (bottom-right corner).
   - Your browser opens automatically.
3. To reopen: double-click the tray icon or double-click the `.vbs` again.
4. To stop: right-click the tray icon → *Quit*.

> **First-time tip:** Pin the launcher to your taskbar / Dock for instant access.

---

## Option C — Command line

```bash
git clone https://github.com/UKCEH-MolecularEcology/fair-edna-filler.git
cd fair-edna-filler
pip install -r requirements.txt          # pystray, pillow, pdfplumber, anthropic
ANTHROPIC_API_KEY=sk-ant-... python3 tray_app.py   # tray icon + browser
```

Or run the web server directly:
```bash
python3 server.py                        # opens http://127.0.0.1:8787
```

### Environment variables

| Variable | Default | Description |
|---|---|---|
| `ANTHROPIC_API_KEY` | *(none)* | Claude API key for AI-powered extraction |
| `FAIRE_CLAUDE_MODEL` | `claude-haiku-5-5` | Claude model used for extraction |
| `FAIRE_TOOL_HOST` | `127.0.0.1` | Bind address (`0.0.0.0` for public) |
| `FAIRE_TOOL_PORT` | `8787` | Port |

---

## Option D — Online / hosted

### Docker

```bash
docker build -t faire-filler .
docker run -p 8787:8787 -e ANTHROPIC_API_KEY=sk-ant-... faire-filler
```

### Railway / Render / Fly.io

Add a `Procfile`:
```
web: python3 server.py
```
Set `ANTHROPIC_API_KEY`, `FAIRE_TOOL_HOST=0.0.0.0`, `FAIRE_TOOL_PORT=$PORT` in the platform dashboard.

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
| `.pdf` | ✅ Text-based PDFs (needs `pdfplumber` for the desktop/server versions; installed by the launchers). Scanned PDFs need OCR first. |

---

## Project structure

```
fair-edna-filler/
├── server.py                    # All the logic: reading reports, extraction, filling, web server
├── index.html                   # No-install version (runs server.py in the browser via Pyodide)
├── build_standalone.py          # Rebuilds index.html from server.py — run after editing server.py
├── standalone/shell.html        # Page layout used by build_standalone.py
├── tray_app.py                  # System-tray launcher (no terminal needed)
├── FAIRe_Metadata_Filler.ipynb  # Google Colab notebook
├── requirements.txt
├── Dockerfile
├── FAIRe Metadata Filler.app/   # macOS app bundle (double-click)
├── start_windows_silent.vbs     # Windows silent launcher (double-click)
├── start_mac.command            # macOS terminal fallback
├── start_windows.bat            # Windows terminal fallback
├── start_linux.sh               # Linux launcher
└── static/
    ├── index.html               # Browser UI (UKCEH branded)
    ├── app.js
    └── styles.css
```

---

## Citation

If you use this tool, please also cite the FAIRe eDNA checklist and guidelines — see the project's [How to cite](https://fair-edna.github.io/citation.html) page.

---

## Contributing

Issues and pull requests welcome. Everything lives in `server.py`: the rule-based extractor (`extract`), the workbook writer (`fill_workbook`) and the Claude prompt (`CLAUDE_PROMPT`). After changing it, run `python3 build_standalone.py` so `index.html` uses the same code.

## Licence

MIT — © UK Centre for Ecology & Hydrology
