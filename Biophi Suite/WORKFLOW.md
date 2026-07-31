# Biophi Suite — Behavior & Setup Notes

## Directory layout

```
Biophi Suite/
├── App/
│   ├── biophi_app.py          Core analysis engine (leaf measurement, Excel output)
│   ├── variety_overview.py    Workbook extraction + Roboflow leaf-cutout rendering
│   ├── brochure_maker.py      Airtable fetch, asset download, brochure orchestration
│   ├── build_brochure_pptx.py Python/pptx brochure builder — Linux/macOS/Windows ✓
│   ├── build_brochure.ps1     Legacy Windows COM brochure builder (reference only)
│   └── main.py                Windows desktop GUI (tkinter + winreg — Windows only)
├── Assets/
│   ├── airtable_cache.json    Offline Airtable record cache (auto-populated)
│   └── Airtable HD Cache/     Downloaded HD images for brochure hero backgrounds
├── template.pptx              Branded slide master (not included — add separately)
├── run_biophi_suite.py        Server-compatible (headless) entry point
├── run_biophi_suite.bat       Windows desktop launcher (starts main.py GUI)
└── WORKFLOW.md                This file
```

---

## Environment variables

| Variable | Required | Description |
|---|---|---|
| `ROBOFLOW_API_KEY` | **Yes** | Roboflow cloud API key |
| `AIRTABLE_PAT` | **Yes** | Airtable Personal Access Token |
| `BROCHURE_AIRTABLE_BASE` | Optional | Airtable base ID (e.g. `appXXXXXXXX`). Falls back to the value compiled into `brochure_maker.py`. |
| `BROCHURE_AIRTABLE_TABLE` | Optional | Airtable table name or ID. Same fallback. |

> **Security rule — never commit API keys.**
> Set them via:
> - Shell: `export ROBOFLOW_API_KEY="…"` (Linux/macOS)
> - Shell: `setx ROBOFLOW_API_KEY "…"` (Windows — persisted to user environment)
> - Hosting platform: add as *secrets* or *environment variables* in your server/CI settings
> - `.env` file loaded by `python-dotenv` (add `.env` to `.gitignore`)

---

## Python dependencies

```
numpy>=1.24
openpyxl>=3.1
Pillow>=10.0
python-pptx>=0.6.21      # required for build_brochure_pptx.py (Linux-compatible deck builder)
requests                 # used by Roboflow HTTP calls
```

Install with:

```bash
pip install -r requirements.txt
```

---

## Pipeline overview

### Full pipeline (server / headless)

```bash
python run_biophi_suite.py \
    --analysis  "/path/to/Biophi Analysis Output.xlsx" \
    --template  template.pptx \
    --output-dir output/ \
    --trial-name "Trial 42"
```

Produces:
- `output/Brochure Output/<timestamp>/Brochure Maker Data.xlsx` — Airtable data workbook
- `output/Brochure Output/<timestamp>/Variety Brochure.pptx` — branded deck

### Windows desktop GUI

Double-click `run_biophi_suite.bat` or run `python App/main.py`.

---

## PowerPoint output layer

### Current Windows approach (legacy)
`brochure_maker.py` calls `build_brochure.ps1` via a PowerShell subprocess.
That script uses Windows COM automation (`New-Object -ComObject PowerPoint.Application`)
and requires Microsoft PowerPoint to be installed.

### Linux / server approach (new)
`brochure_maker.py` now tries to import `build_brochure_pptx` first.
If found, `build_brochure_pptx.build_deck_from_manifest()` is called instead of the
PowerShell script.  This uses `python-pptx` and has no OS-level dependencies.

The Python builder replicates the PS1 slide layout:
- Slide duplication with independent chart parts per variety
- Named shape discovery (`TextBox 7`, `Picture 36`, `Chart 13`, etc.)
- Text substitution in identity, description, harvest, and environmental boxes
- Image replacement for hero and leaf cutout photos
- Scale ruler drawn with lines and textboxes
- Yield chart data replacement

**Known difference from COM version:**
- Chart formatting (number formats, data-label styles) is reset to python-pptx
  defaults when `chart.replace_data()` is called.  Visual appearance may differ
  from the COM-built version.  Adjust chart styles in the template if needed.

---

## Airtable cache

`brochure_maker.py` writes `Assets/airtable_cache.json` after a successful
live fetch.  Subsequent runs use this cache automatically when:
- `AIRTABLE_PAT` is absent, **or**
- The live fetch returns HTTP 401 / 403

To force a fresh fetch, delete `Assets/airtable_cache.json` or set
`AIRTABLE_PAT` and run the pipeline.

---

## Offline / air-gapped operation

1. Run the pipeline at least once with live credentials to populate:
   - `Assets/airtable_cache.json`
   - `Assets/Airtable HD Cache/` (HD images)
2. On subsequent offline runs the cache is used automatically.

---

## Adding the slide template

`template.pptx` is not included in this repository to keep the package size
small.  Copy the branded template file into this directory before running the
pipeline.  The template must contain at least one slide with the named shapes
listed in `App/build_brochure_pptx.py` (`TextBox 7`, `Picture 36`, etc.).
