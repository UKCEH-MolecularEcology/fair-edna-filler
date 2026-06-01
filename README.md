# FAIRe Metadata Filler

**Automatically fill the [FAIRe eDNA metadata template](https://fair-edna.github.io/download.html) from a project report.**

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/UKCEH-MolecularEcology/fair-edna-filler/blob/main/FAIRe_Metadata_Filler.ipynb)

Upload a Word document, PDF, or plain-text report and the tool extracts project information, sampling details, PCR/sequencing methods, and bioinformatics parameters — then writes them into the correct cells of the FAIRe Excel workbook. Fields that cannot be found are filled with approved [FAIRe missing-value wording](https://fair-edna.github.io/guidelines.html#missing-values) so the workbook is always valid for the [FAIRe-fier validator](https://shiny.csiro.au/FAIRe-fier/).

---

## What it fills

| Sheet | What gets populated |
|---|---|
| `projectMetadata` | Project name/ID, institution, funder, dates, PCR conditions, primer sequences, sequencing platform/kit, bioinformatics pipeline, reference databases |
| `sampleMetadata` | One row per water sample and per field blank, with location, date, filtration, extraction, and assay details |
| `experimentRunMetadata` | One row per sample × assay combination, with library and sequencing run identifiers |

The tool does **not** modify `taxaRaw`, `taxaFinal`, `ampData`, `stdData`, or `eLowQuantData` — those require your actual sequence data.

---

## Inputs & outputs

| Input | Description |
|---|---|
| **Report file** | `.docx`, `.pdf`, `.doc`, or `.txt` |
| **FAIRe template** | FAIRe full template `.xlsx` — [download here](https://fair-edna.github.io/download.html) |
| **Claude API key** *(optional)* | Enables AI-powered extraction — [get a free key](https://console.anthropic.com) |

**Output:** A filled `.xlsx` workbook ready to review in Excel and upload to the FAIRe-fier validator.

---

## Option A — No installation: Google Colab

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
| **Rule-based** | No API key | Regex patterns optimised for Defra/UKCEH reports |

---

## Supported file types

| Format | Support |
|---|---|
| `.docx` | ✅ Always |
| `.txt` | ✅ Always |
| `.pdf` | ✅ Requires `pdfplumber` (installed by launchers) |
| `.doc` | ✅ Requires LibreOffice (`soffice`) or macOS `textutil` |

---

## Project structure

```
fair-edna-filler/
├── server.py                    # Web server, extraction logic, workbook filling
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

If you use this tool, please also cite the FAIRe eDNA guidelines:

> Turon X et al. (2025). FAIR eDNA (FAIRe) metadata checklist and guidelines. https://fair-edna.github.io

---

## Contributing

Issues and pull requests welcome. The Claude extraction prompt is in `server.py` (`_CLAUDE_EXTRACTION_PROMPT`).

## Licence

MIT — © UK Centre for Ecology & Hydrology
