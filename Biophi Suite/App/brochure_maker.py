from __future__ import annotations

import json, os, re, shutil, urllib.parse, urllib.request, sys
from urllib.error import HTTPError
from pathlib import Path
from datetime import datetime
from io import BytesIO

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from PIL import Image

BASE_ID = os.environ.get("BROCHURE_AIRTABLE_BASE", "apphOKNN9qjm6IghM")
TABLE_ID = os.environ.get("BROCHURE_AIRTABLE_TABLE", "tblVJa7JUazw7SnCO")

FIELDS = {
    "id": "VarietyID", "name": "Name (from E#)", "use": "UseinBrochure",
    "description": "M&SDescription", "hr": "Resistance HR (from E#)",
    "ir": "Resistance IR (from E#)", "inventory": "Available Inventory # (from Variety)",
    "variety": "Variety", "marketing": "Photo (marketing)", "hd": "HD Image",
    "hd2": "HD Image 2", "wh": "WH Image", "wh2": "WH Image 2", "background": "BackgroundPhoto",
    "segment": "Broad Segment",
}

def _attachments(value):
    return [{"url": x.get("url"), "filename": x.get("filename", "")} for x in (value or []) if x.get("url")]

def fetch_records():
    cache_path = Path(__file__).resolve().parents[1] / "Assets" / "airtable_cache.json"
    token = os.environ.get("AIRTABLE_PAT") or os.environ.get("AIRTABLE_API_KEY")
    if not token:
        if cache_path.is_file():
            return json.loads(cache_path.read_text(encoding="utf-8")).get("records", [])
        raise RuntimeError("Set AIRTABLE_PAT (a Personal Access Token with read access) before using Brochure Maker.")
    rows, offset = [], None
    while True:
        params = {"pageSize": "100"}
        if offset: params["offset"] = offset
        url = f"https://api.airtable.com/v0/{BASE_ID}/{TABLE_ID}?" + urllib.parse.urlencode(params)
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
        try:
            with urllib.request.urlopen(req, timeout=45) as response:
                payload = json.load(response)
        except HTTPError as error:
            if error.code in (401, 403) and cache_path.is_file():
                return json.loads(cache_path.read_text(encoding="utf-8")).get("records", [])
            if error.code == 401:
                raise RuntimeError("Airtable rejected the saved token. Create a Personal Access Token with data.records:read access and access to the John Gray base, then use Set Airtable Token.") from error
            if error.code == 403:
                raise RuntimeError("The Airtable token is valid but does not have access to the John Gray / Brochure Maker table.") from error
            raise
        for record in payload.get("records", []):
            f = record.get("fields", {})
            row = {k: f.get(v) for k, v in FIELDS.items()}
            row["segment"] = (
                f.get("Broad Segment")
                or f.get("Broad Segment (from Variety)")
                or f.get("Segment")
            )
            for key in ("marketing", "hd", "hd2", "wh", "wh2", "background"):
                row[key] = _attachments(row[key])
            row["record_id"] = record.get("id")
            rows.append(row)
        offset = payload.get("offset")
        if not offset: break
    return rows

def download_assets(rows, folder: Path, keys=None):
    folder.mkdir(parents=True, exist_ok=True)
    asset_keys = tuple(keys or ("marketing", "hd", "hd2", "wh", "wh2", "background"))
    for row in rows:
        slug = re.sub(r"[^A-Za-z0-9_-]+", "_", str(row.get("variety") or row.get("name") or row.get("id") or "variety"))
        row.setdefault("local_assets", {})
        for key in asset_keys:
            items = row.get(key) or []
            if not items: continue
            asset = items[0]
            suffix = Path(asset.get("filename", "image.jpg")).suffix or ".jpg"
            if row["local_assets"].get(key) and Path(row["local_assets"][key]).is_file():
                continue
            target = folder / f"{slug}_{key}{suffix}"
            if target.is_file():
                row["local_assets"][key] = str(target.resolve())
                continue
            try:
                urllib.request.urlretrieve(asset["url"], target)
                row["local_assets"][key] = str(target.resolve())
            except Exception:
                pass

def write_data_workbook(rows, output: Path):
    wb = Workbook(); ws = wb.active; ws.title = "Brochure Data"
    headers = ["VarietyID", "Variety", "Broad Segment", "Description", "Resistance HR", "Resistance IR", "Inventory", "Strengths", "Weaknesses", "Final DLI", "Final Temperature", "Final RH", "Annualized Yield by DTM (kg/m2/yr)", "Light Use Efficiency by DTM (g/mol)", "Hero Image", "Marketing Image", "HD Image", "HD Image 2", "WH Image", "WH Image 2", "Background Photo"]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="64D273"); c.alignment = Alignment(wrap_text=True)
    for r in rows:
        assets = r.get("local_assets", {})
        trial = r.get("trial_data") or {}; env = trial.get("environment") or {}
        curve = "; ".join(f"DTM {p.get('dtm')}: {p.get('value'):.2f}" for p in trial.get("yield", []) if isinstance(p.get("value"), (int, float)))
        efficiency = "; ".join(f"DTM {p.get('dtm')}: {p.get('gmol'):.2f}" for p in trial.get("yield", []) if isinstance(p.get("gmol"), (int, float)))
        ws.append([r.get("id"), r.get("variety") or r.get("name"), r.get("segment"), r.get("description"), r.get("hr"), r.get("ir"), r.get("inventory"), trial.get("strengths"), trial.get("weaknesses"), env.get("dli"), env.get("temperature"), env.get("rh"), curve, efficiency, r.get("hero_path"), assets.get("marketing"), assets.get("hd"), assets.get("hd2"), assets.get("wh"), assets.get("wh2"), assets.get("background")])
    for column in range(1, len(headers) + 1):
        ws.column_dimensions[ws.cell(1, column).column_letter].width = 24
    ws.freeze_panes = "A2"; output.parent.mkdir(parents=True, exist_ok=True); wb.save(output)

def _name_key(value) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def _marketing_match(marketing_rows, trial_name):
    trial_key = _name_key(trial_name)
    for row in marketing_rows:
        if trial_key and trial_key in {
            _name_key(row.get("variety")),
            _name_key(row.get("name")),
            _name_key(row.get("id")),
        }:
            return row
    return None


def _prepare_hero(row, fallback_row, cache_dir: Path):
    source_row = row if (row.get("local_assets") or {}).get("hd") else fallback_row
    source = (source_row.get("local_assets") or {}).get("hd") if source_row else None
    if not source or not Path(source).is_file():
        row["hero_path"] = None
        return
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", str(row.get("variety") or row.get("name") or row.get("id") or "variety"))
    destination = cache_dir / f"{safe_name}_airtable_hd_hero_v3.png"
    if destination.is_file():
        row["hero_path"] = str(destination.resolve())
        return
    from variety_overview import cutout_largest_leaf, roboflow_key
    key = roboflow_key()
    if not key:
        raise RuntimeError("Roboflow API key is required to prepare transparent Airtable HD images.")
    cutout = cutout_largest_leaf(Path(source).read_bytes(), destination, key)
    if not cutout:
        row["hero_path"] = None
        return
    with Image.open(destination) as image:
        rgba = image.convert("RGBA")
        rgba.putalpha(rgba.getchannel("A").point(lambda value: round(value * 0.35)))
        rgba.save(destination)
    row["hero_path"] = str(destination.resolve())


def build_brochure(analysis_workbook: Path | None, template: Path, output_folder: Path, trial_name: str = ""):
    run_folder = output_folder / "Brochure Output" / datetime.now().strftime("%Y-%m-%d_%H%M%S")
    marketing_rows = fetch_records()
    rows = []
    overview = None
    if analysis_workbook and Path(analysis_workbook).is_file():
        overview_dir = Path(__file__).resolve().parent
        if str(overview_dir) not in sys.path: sys.path.insert(0, str(overview_dir))
        from variety_overview import extract_workbook
        photo_cache = output_folder / "Variety Overview Photos"
        overview = extract_workbook(Path(analysis_workbook), photo_cache)
        trial_name = trial_name or str(overview.get("trial") or "")
        for trial in overview.get("varieties", []):
            marketing = _marketing_match(marketing_rows, trial.get("name"))
            row = dict(marketing or {})
            row.setdefault("id", trial.get("name"))
            row.setdefault("name", trial.get("name"))
            row.setdefault("variety", trial.get("name"))
            trial["environment"] = overview.get("environment") or {}
            row["trial_data"] = trial
            rows.append(row)
    else:
        rows = [dict(row) for row in marketing_rows if row.get("use") is True]
    if not rows:
        raise RuntimeError("No varieties were found in the selected analysis workbook.")
    download_rows = list(rows)
    cristabel = _marketing_match(marketing_rows, "Cristabel")
    if cristabel and cristabel not in download_rows:
        download_rows.append(cristabel)
    persistent_assets = Path(__file__).resolve().parents[1] / "Assets" / "Airtable HD Cache"
    download_assets(download_rows, persistent_assets, keys=("hd",))
    hero_cache = output_folder / "Variety Overview Photos" / "Airtable HD Heroes"
    hero_cache.mkdir(parents=True, exist_ok=True)
    for row in rows:
        _prepare_hero(row, cristabel, hero_cache)
    data_path = run_folder / "Brochure Maker Data.xlsx"; write_data_workbook(rows, data_path)
    manifest = {"trial": trial_name, "analysis": str(analysis_workbook or ""), "rows": rows}
    manifest_path = run_folder / "brochure_manifest.json"; manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    # The PowerPoint template is intentionally kept as the source of truth.  The
    # builder duplicates its first slide and replaces matching text/image slots.
    # On Linux/macOS/Windows we use the python-pptx builder; the legacy
    # build_brochure.ps1 (Windows COM automation) is kept alongside as a reference.
    output_pptx = run_folder / "Variety Brochure.pptx"
    try:
        from build_brochure_pptx import build_deck_from_manifest
        build_deck_from_manifest(str(manifest_path), str(template), str(output_pptx))
    except ImportError:
        # Fallback: Windows COM automation via PowerShell
        import subprocess
        script = Path(__file__).with_name("build_brochure.ps1")
        env = os.environ.copy()
        env["BROCHURE_MANIFEST"] = str(manifest_path)
        env["BROCHURE_TEMPLATE"] = str(template)
        env["BROCHURE_OUTPUT"] = str(output_pptx)
        env["BROCHURE_ANALYSIS"] = str(analysis_workbook or "")
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script)],
            env=env, check=True,
        )
    manifest_path.unlink(missing_ok=True)
    return data_path, output_pptx
