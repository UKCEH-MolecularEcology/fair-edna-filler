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
- **Sample rows.** A report rarely lists every sample, so by default the rows are estimated from how sampling is described. For example, "Troutbeck Fan (TFAN, N=4) … each site was sampled five times" gives TFAN001_V1 … TFAN004_V5, plus field, extraction and PCR blanks. For the real list, with dates, coordinates, volumes and run IDs per sample, add an optional **sample list** (`.csv`/`.xlsx`, one row per sample — see [`template/sample_list_template.csv`](template/sample_list_template.csv)). Column names can be plain ("sample", "date", "site", "latitude", "volume filtered (ml)") or any FAIRe sampleMetadata / experimentRunMetadata term. Blanks are recognised from their names or type.
- **IDs are generated** from the report structure (e.g. `BASS_R1_01`, `BASS_R1_FB`, `EXTBLANK_01`, `PCRNEG_16S_01`, `run_20260219_600cycle`). Replace them with your lab's IDs before publishing.
- **You only check what needs checking.** Values that come with a clear source sentence and pass the FAIRe checks are accepted automatically. The review table opens on the few flagged ones, each with the reason, the sentence it came from and an edit box. See [Automatic review](#automatic-review).
- **Before download the workbook is validated** against the template's own field types: numbers, dates, ISO durations, ontology terms, controlled vocabularies, mandatory fields and sample/library ID links. Safe fixes are made automatically, e.g. `25 µL` → `25`, `24 hours` → `PT24H`, `14/07/2025` → `2025-07-14`.
- After filling, the page lists **"Check these by hand"** notes (for example inconsistencies spotted in the report).
- The rest of the template — dropdowns, term notes, formatting, other sheets — is left exactly as it was.

---

## Inputs & outputs

| Input | Description |
|---|---|
| **Report file** | `.docx`, `.pdf`, `.doc`, or `.txt` |
| **Links / extra files** *(optional)* | Protocol, dataset archive, GitHub, BioProject, or your own tables and scripts – read automatically from the report or added by you. See [More sources](#more-sources-than-the-report) |
| **Sample list** *(optional)* | `.csv` or `.xlsx`, one row per sample — gives one sampleMetadata row (and one run row per assay) for each sample |
| **FAIRe template** *(optional)* | The official FAIRe FULL template is built in ([`template/FAIRe_checklist_v1.0.2_FULLtemplate.xlsx`](template/FAIRe_checklist_v1.0.2_FULLtemplate.xlsx)). Only choose a template to use a different one, e.g. made with FAIRe-ator. |
| **Claude API key** *(optional, off by default)* | Only for reports that are **not confidential** — see [Privacy](#privacy-and-how-values-are-found) |

**Output:** A filled `.xlsx` workbook ready to upload to the FAIRe-fier validator. The Python version also writes `<name>_review.csv`, listing every value with the report sentence it came from.

---

## How to use it

Pick whichever suits you — all three give the same result. Works on **macOS, Windows and Linux**.

### 1. In your browser — no installation (recommended)

**Online:** open **https://ukceh-molecularecology.github.io/fair-edna-filler/**

**Offline copy:** green **Code** button → *Download ZIP* → extract → double-click **`index.html`**.

Then choose your report (and, optionally, a saved profile), click **Read report**, look at the values under *Needs a look*, click **Fill workbook with these values**, and download the result — the official FAIRe template is built in. (When opening the downloaded `index.html` directly, also choose the template file from the `template` folder: browsers don't let a local page read neighbouring files.) Everything runs inside your browser — your files are never uploaded anywhere. The first visit downloads the tool's Python engine (about 10 MB), so you need an internet connection once. Works in current Chrome, Edge, Firefox and Safari.

### 2. Google Colab

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/UKCEH-MolecularEcology/fair-edna-filler/blob/main/FAIRe_Metadata_Filler.ipynb)

Open the notebook and click ▶ on its two cells: **① Set up**, then **② Fill from my report**. Upload your report, check the flagged values in the review screen (no code to edit), and click **Fill workbook and download**. The official FAIRe FULL template is fetched automatically; tick a box to use your own template or a saved profile. To add the study's protocol, dataset, GitHub or BioProject, tick **read_links_in_the_report** or paste links into **extra_sources** (see [More sources](#more-sources-than-the-report)). Needs a Google account.

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

To combine the report with its protocol, dataset archive, GitHub repository, BioProject or your own files, add `--follow-links` and/or `--source …`. See [More sources than the report](#more-sources-than-the-report).

The review CSV lists what needs a look first. To correct values without the browser, type your value into its `your_value` column (`-` empties a value), then run:

```bash
python3 server.py report.pdf --edits outputs/report_review.csv --save-profile lab.json   # apply corrections, remember them
python3 server.py uploads/ --profile lab.json                                            # next reports: corrections re-applied
python3 server.py report.docx --samples my_samples.xlsx                                  # one row per sample from your list
python3 server.py report.pdf --local-ai                                                  # + second opinion from a local model
```

The `uploads/` folder is a convenient place to drop several reports for batch filling; its contents are never committed to git.

Or use it from your own scripts:

```python
import server
meta = server.prepare(server.extract_report_text("report.doc"))   # built-in reader + primer library
for item in server.review(meta):        # status: ok / check / missing, with reasons and source sentence
    if item["status"] == "check":
        print(item["id"], "=", item["value"], "|", item["reasons"], "|", item["evidence"])
server.apply_review(meta, {"project:project_name": "My corrected title"})   # optional corrections
summary = server.fill_workbook(server.TEMPLATE_PATH, meta, "report.doc", "filled.xlsx")   # bundled template
print(summary["notes"])                # things to check by hand
```

| Environment variable | Default | Description |
|---|---|---|
| `FAIRE_CLAUDE_MODEL` | `claude-haiku-5-5` | Claude model (non-confidential reports only) |
| `FAIRE_LOCAL_MODEL` | `qwen2.5:7b` | Ollama model used by `--local-ai` |
| `OLLAMA_HOST` | `http://127.0.0.1:11434` | Ollama address; only this computer is allowed |
| `FAIRE_TOOL_HOST` | `127.0.0.1` | Bind address (`0.0.0.0` to serve on a network) |
| `FAIRE_TOOL_PORT` | `8787` | Port |

---

## More sources than the report

A report or paper rarely contains everything FAIRe asks for. The details are usually spread over:

- a **protocol** (e.g. on protocols.io),
- a **dataset archive** (Zenodo, Dryad, Figshare),
- a **code/data repository** (GitHub),
- a **sequence archive** (SRA/ENA BioProject),
- and your own **files** (sample sheets, read tables, scripts).

The tool can read all of these and combine them with the report. There are two ways to tell it where to look, and you can use both together.

### Option A – let it follow the links printed in the report

Most papers have a *Data availability* section with links. Switch on **follow links** and the tool finds them and reads each one.

- **Colab:** in step ②, tick **read_links_in_the_report**.
- **Python:**
  ```bash
  python3 server.py paper.pdf --follow-links
  ```

It recognises these links, even when the PDF has broken them up with spaces (e.g. `https:// doi. org/ 10. 5281/ zenodo. 123`):

| In the report | Read as |
|---|---|
| `10.17504/protocols.io.…` | protocol |
| `10.5281/zenodo.…`, `10.5061/dryad.…`, `10.6084/m9.figshare.…` | dataset archive |
| `github.com/owner/repo` | code/data repository |
| `PRJNA…`, `PRJEB…`, `PRJDB…`, `SRP…`, `ERP…` | sequence archive |

A GitHub link that is only mentioned in passing, outside the data-availability text, is recorded as `code_repo` but not downloaded. An example is the bioinformatics pipeline's own repository.

### Option B – add sources yourself

Use this for links that are not in the report, or for files on your computer. You can add as many as you like.

- **Colab:** in step ②, paste them into **extra_sources**, separated by spaces.
- **Python:** add one `--source` per link or file:
  ```bash
  python3 server.py report.docx \
      --source 10.17504/protocols.io.eq2lyqm8mvx9/v1 \
      --source 10.5281/zenodo.15723454 \
      --source https://github.com/clarecowgill/Cowgill-et-al.-2025-Multi-tool-monitoring \
      --source PRJNA1345715 \
      --source my_protocol.pdf --source lab_samples.xlsx --source reads_table.csv --source pipeline.R
  ```

### What each source can add

| Source | Accepted as | What it fills |
|---|---|---|
| **Protocol** | protocols.io link or DOI, or a protocol file (`.pdf` or `.docx`; a `.txt` file needs "protocol" in its name) | PCR recipe and cycling (also for the indexing PCR), thermocycler, sterilisation, extraction, library quantification, bioinformatics settings: trimming, merging, length filter, denoising, chimera check, taxonomy and contamination thresholds. The protocol DOI goes into `nucl_acid_ext`, `nucl_acid_amp` and `sop_bioinformatics`. |
| **Dataset archive** | Zenodo, Dryad or Figshare DOI or link | Authors with ORCIDs, licence, the data-holding institution and its ROR ID, plus the data and code files inside the record (files over 20 MB are skipped) |
| **Repository** | GitHub link | `code_repo`; sample tables; blank–sample links (`rel_cont_id`); control names explained in the README; taxon × sample read tables → `taxaRaw`, read counts per sample; presence/absence tables → `taxaFinal` |
| **Sequence archive** | BioProject / study accession | Instrument, read layout, file names, MD5 checksums and run accessions for each sample. If the runs are pooled libraries, add a `library` column to your sample list to link each sample to its run. |
| **Your files** | `.csv`, `.tsv`, `.xlsx`, `.zip`, `.R`, `.py`, … | The same as above, without going online. The kind of table (sample sheet, read table, blank links, presence/absence) is recognised from its contents. |

### What happens when sources disagree

Every value shows where it came from in the review table and in the review CSV, e.g. *"from protocols.io"* or *"taken from protocols.io; the report says …"*. When two sources differ, the more specific one wins and the value is flagged for you to check:

- **the protocol** wins for methods;
- **the dataset archive** wins for ORCIDs, licence and institution;
- **data files and your sample list** win for samples (your own `--samples` list always takes priority);
- **the report** is used for everything else.

Sample-sheet columns with no FAIRe term (e.g. `habitat`, `water_flow`) are not thrown away. They are added as user-defined columns at the end of `sampleMetadata`.

### If something cannot be read

A broken or unreachable link never stops the run. It is listed in the notes, printed on screen and saved in the review CSV. Common cases:

| Note says | What to do |
|---|---|
| *protocols.io did not let the tool download the protocol* | Open the protocol page, choose **⋯ → Download PDF**, and add it with `--source protocol.pdf` |
| *large files not downloaded* | Normal for raw sequences or media; they are not needed for the metadata |
| *runs are pooled libraries* | Add a `library` column to your sample list (`--samples`) |
| *could not read …* (a single file) | The other files are still used; add that file yourself with `--source` if it matters |

**Privacy:** nothing is downloaded unless you add a link or switch on *follow links*. Only public records are fetched, and your report never leaves your computer. Online sources work in **Python and Colab**. The browser page reads the report and your sample list only.

**Example:** on Cowgill et al. (2026), the paper alone gave 25 correct project terms out of 86 checked. With follow links switched on (protocols.io, Zenodo, GitHub and the BioProject), that rose to 74. The run also produced 224 sample rows linked to their blanks, 93 `taxaRaw` and 45 `taxaFinal` taxa.

## Automatic review

The tool decides for each value whether it can be **accepted automatically** or **needs a look**:

| Accepted automatically when… | Flagged for a look when… |
|---|---|
| it comes with a source sentence from the report (from the *Methods* section, for methods values) | there's no source sentence, or the sentence is outside the Methods section |
| it passes the template's format check (number, date, duration, ontology term, vocabulary) | the format is wrong, or the number is unusual (e.g. annealing at 95 °C) |
| a primer sequence matches a published primer in the built-in library | a primer differs from the published sequence, or is listed for another marker or direction |
| the report gives one value | the report gives different values in different places ("the report also gives: …") |
| your saved profile or a second reader agrees | your profile or a second reader disagrees, or filled a gap |

**Primer and database library.** Common eDNA primers are built in: MiFish, 12S-V5, Tele02, Leray/Geller COI, Folmer, fwh, BF3/BR2, 515F/806R/926R, 341F/785R, 1391f/EukBr, TAReuk, diatom rbcL, ITS and trnL g/h. Each is stored with its sequence and reference. If a report prints a sequence, the tool adds the primer's name and reference. If it only names the primer, the sequence is filled from the library and flagged. Add more primers to `PRIMER_LIBRARY` in `server.py`.

**Profiles: learning from your corrections.** After you correct values, click **Save my corrections** in the browser; in Colab, tick *save_profile*; on the command line, use `--save-profile`. You get a small `.json` file. Load it with the next report and the same corrections are applied automatically. A correction is applied when the tool finds the same value it found last time, or finds nothing; anything else is flagged. Use profiles for things like the lab's contact, licence, kit naming or a primer that is always named in a table image. Share the file with colleagues to reuse the same corrections.

**Local AI second opinion (optional, Python only).** With [Ollama](https://ollama.com) installed (`ollama pull qwen2.5:7b`), `python3 server.py report.pdf --local-ai` asks a model running on your own computer to read the Methods. It must quote the report for every value, like Claude. Where it agrees, values are accepted. Where it disagrees, they are flagged. Gaps it fills are flagged for confirmation. The report never leaves your computer: the tool refuses to send it to any address other than this machine.

## Privacy and how values are found

**By default nothing leaves your computer.** A built-in reader looks for the usual ways methods are written in eDNA reports and papers: primer sequences and names, PCR recipes, filters, kits, sequencers, bioinformatics tools and reference databases. It works out which assay each value belongs to. Where a report has a *Methods* section, it reads that section only, so numbers from the results or introduction are not picked up. References are ignored. Every value keeps the sentence it came from, and you see it in the review table.

**Claude (optional, off by default).** For reports that are *not confidential*, you can tick *"This report is not confidential"* and enter a Claude API key. The report text is then sent to Anthropic's API, and Claude gives a second opinion in the same way as the local AI. Claude must quote the sentence behind each value, and any value whose quote is not in the report is dropped. Never use this for sensitive or unpublished reports.

**Other languages.** The built-in reader is written for English. Reports in other languages are detected and flagged; primer sequences, numbers, kit and tool names are still found, but expect more gaps to fill by hand.

**Targeted (qPCR / ddPCR) studies.** The tool is built for metabarcoding. If a report looks like a targeted study, `assay_type` is set to `targeted` and you are told which fields (probe, standard curve, LOD/LOQ) to fill by hand.

### How well does it work?

`bench/run_benchmark.py` runs the tool on a folder of reports and compares the result with the published FAIRe example workbooks for the same studies. On 14 public reports and papers, three of which have a published workbook, 120 values could be checked:

- 61 were right and 13 differed from the published workbook. Most of the 13 are naming or rounding differences, e.g. "Qiagen DNeasy Blood & Tissue kit" vs "Qiagen Blood & Tissue Kit", or 12.5 vs 13 µL. About four are real errors, such as a storage temperature.
- 46 were not found: things a report doesn't state in words, or values in tables that are images.
- Of the values **accepted automatically**, 42 were right and 11 differed (the naming differences above, plus about three real errors). 21 values were **flagged for a look**; 19 of those were right.

So expect to look at roughly 10–20 flagged values per report, plus the mandatory fields the report doesn't state.

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
