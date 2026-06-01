# FAIRe Metadata Filler

**Automatically fill the [FAIRe eDNA metadata template](https://fair-edna.github.io/download.html) from a project report.**

Upload a Word document, PDF, or plain-text report and the tool extracts project information, sampling details, PCR/sequencing methods, and bioinformatics parameters — then writes them into the correct cells of the FAIRe Excel workbook. Fields that cannot be found are filled with approved [FAIRe missing-value wording](https://fair-edna.github.io/guidelines.html#missing-values) so the workbook always passes the [FAIRe-fier validator](https://shiny.csiro.au/FAIRe-fier/).

---

## What it fills

| Sheet | What gets populated |
|---|---|
| `projectMetadata` | Project name/ID, institution, funder, dates, PCR conditions, primer sequences, sequencing platform/kit, bioinformatics pipeline, reference databases |
| `sampleMetadata` | One row per water sample and per field blank, with location, date, filtration, extraction, and assay details |
| `experimentRunMetadata` | One row per sample × assay combination, with library and sequencing run identifiers |

The tool does **not** modify `taxaRaw`, `taxaFinal`, `ampData`, `stdData`, or `eLowQuantData` — those require your actual sequence data.

---

## Inputs

| Input | Description |
|---|---|
| **Report file** | `.docx`, `.pdf`, `.doc`, or `.txt` — your project's final report or methods document |
| **FAIRe template** | The FAIRe full template workbook (`.xlsx`) — download from [fair-edna.github.io/download.html](https://fair-edna.github.io/download.html) |
| **Claude API key** *(optional)* | Enables AI-powered extraction. Without it, rule-based parsing is used (works well for standard Defra/UKCEH report formats). [Get a free key →](https://console.anthropic.com) |

## Output

A filled `.xlsx` workbook, ready to open in Excel or upload to the [FAIRe-fier validator](https://shiny.csiro.au/FAIRe-fier/). You will still need to:
- Add exact sample coordinates (the tool fills approximate lake centroids where known)
- Add sample-specific DNA concentrations and sequencing file names
- Review and correct any values marked `missing: not provided`

---

## ▶ Usage — No coding required

> **You need Python installed.** If you are not sure, follow the steps below — it only takes a couple of minutes.

### Step 1 — Check / install Python

- **Windows:** Open the Start menu, search for *Command Prompt*, type `python --version` and press Enter. If you see `Python 3.x.x` you are ready. Otherwise download Python from [python.org/downloads](https://www.python.org/downloads/) — tick **"Add Python to PATH"** during installation.
- **Mac:** Open *Terminal* (search for it in Spotlight), type `python3 --version`. If not found, download from [python.org/downloads](https://www.python.org/downloads/).

### Step 2 — Download this tool

Click the green **Code** button at the top of this page → **Download ZIP**. Extract the ZIP anywhere on your computer (e.g. your Desktop).

### Step 3 — Start the tool

| System | What to do |
|---|---|
| **Windows** | Open the extracted folder and double-click **`start_windows.bat`** |
| **Mac** | Open the extracted folder. Right-click **`start_mac.command`** → Open → click *Open* in the security dialog |
| **Linux** | Open a terminal in the folder and run `bash start_linux.sh` |

The launcher will install any missing packages (internet required the first time), start a local web server, and open your browser automatically.

### Step 4 — Use the web interface

Your browser will open at **http://127.0.0.1:8787**

1. Click *Choose a report file* and select your `.docx` / `.pdf` / `.txt` report.
2. Click *Choose the FAIRe full template* and select the template `.xlsx` you downloaded from the FAIRe website.
3. *(Optional)* Paste your Claude API key into the *Claude API key* box for better extraction.
4. Click **Fill metadata workbook**.
5. When the status bar reaches 100% click **Download filled workbook**.

### Step 5 — Review and complete

Open the downloaded workbook in Excel. Review all cells and fill in anything marked `missing: not provided`. Pay particular attention to:
- **Exact coordinates** (`decimalLatitude` / `decimalLongitude` in `sampleMetadata`)
- **Sample-specific volumes and concentrations** (`concentration`, `samp_vol_we_dna_ext`)
- **Sequencing file names** (`filename`, `filename2` in `experimentRunMetadata`)
- **License and access rights** (`license`, `accessRights` in `projectMetadata`)

> **To stop the tool:** close the terminal / command window that opened when you double-clicked the launcher.

---

## ▶ Usage — Command line

```bash
# Clone or download
git clone https://github.com/UKCEH-MolecularEcology/fair-edna-filler.git
cd fair-edna-filler

# Install dependencies
pip install -r requirements.txt

# Run (Claude AI parsing enabled via env variable)
ANTHROPIC_API_KEY=sk-ant-... python3 server.py
```

Then open **http://127.0.0.1:8787**.

### Environment variables

| Variable | Default | Description |
|---|---|---|
| `ANTHROPIC_API_KEY` | *(none)* | Claude API key. Enables AI-powered extraction |
| `FAIRE_TOOL_HOST` | `127.0.0.1` | Bind address. Set to `0.0.0.0` for public access |
| `FAIRE_TOOL_PORT` | `8787` | Port number |

---

## ▶ Hosting online

### Any Linux VPS

```bash
pip install -r requirements.txt
FAIRE_TOOL_HOST=0.0.0.0 FAIRE_TOOL_PORT=8787 \
ANTHROPIC_API_KEY=sk-ant-... \
python3 server.py
```

Add nginx or Caddy as a reverse proxy for HTTPS.

### Railway / Render / Fly.io

Create a `Procfile` in the repo root:
```
web: python3 server.py
```
Set `ANTHROPIC_API_KEY`, `FAIRE_TOOL_HOST=0.0.0.0`, and `FAIRE_TOOL_PORT=$PORT` in the platform's environment settings.

### Docker

```bash
docker build -t faire-filler .
docker run -p 8787:8787 -e ANTHROPIC_API_KEY=sk-ant-... faire-filler
```

---

## AI vs rule-based parsing

| Mode | When used | Extraction quality |
|---|---|---|
| **Claude AI** | API key provided | Extracts any report format, understands context, returns structured JSON |
| **Rule-based** | No API key | Regex patterns tuned for Defra/UKCEH report formats — reliable for standard layouts |

Both modes fill the same sheets; the AI mode usually populates more fields correctly and handles unusual formats (e.g. grant applications, lab notebooks, short SOPs).

---

## Supported file types

| Format | Support | Notes |
|---|---|---|
| `.docx` | ✅ Always | Word 2007+ |
| `.txt` | ✅ Always | Plain text |
| `.pdf` | ✅ With `pdfplumber` | `pip install pdfplumber` |
| `.doc` | ✅ With LibreOffice or macOS | Requires `soffice` or `textutil` |

---

## Project structure

```
fair-edna-filler/
├── server.py           # Web server + extraction + workbook filling
├── requirements.txt    # Python dependencies
├── start_mac.command   # macOS double-click launcher
├── start_windows.bat   # Windows double-click launcher
├── start_linux.sh      # Linux launcher
├── Dockerfile          # Container deployment
├── static/
│   ├── index.html      # Browser UI
│   ├── app.js          # Frontend logic
│   └── styles.css      # Styling
├── uploads/            # Temporary upload storage (auto-cleared)
└── outputs/            # Generated workbooks
```

---

## Citation

If you use this tool, please also cite the FAIRe eDNA guidelines:

> Turon X et al. (2025). FAIR eDNA (FAIRe) metadata checklist and guidelines. https://fair-edna.github.io

---

## Contributing

Issues and pull requests welcome. The extraction prompt sent to Claude is in `server.py` (`_CLAUDE_EXTRACTION_PROMPT`). Improving it for specific report types is the easiest way to increase coverage.

## Licence

MIT
