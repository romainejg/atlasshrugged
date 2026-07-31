"""
Airtable client for Biophi Suite.

Fetches records from an Airtable base and maintains a local JSON cache so that
the suite can run offline or without hammering the API on every launch.

Required environment variables:
    AIRTABLE_PAT              -- Airtable personal access token
    BROCHURE_AIRTABLE_BASE    -- Base ID (e.g. "appXXXXXXXXXXXXXX"); defaults to
                                 the value in Assets/airtable_cache/config.json
                                 if the env-var is absent.
    BROCHURE_AIRTABLE_TABLE   -- Table name or ID; same fallback logic.
"""

import json
import os
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests

_CACHE_DIR = Path(__file__).parent.parent / "Assets" / "airtable_cache"
_CACHE_FILE = _CACHE_DIR / "records.json"
_META_FILE = _CACHE_DIR / "meta.json"
_AIRTABLE_API = "https://api.airtable.com/v0"
_CACHE_TTL_SECONDS = 3600  # re-fetch after 1 hour


def _get_config() -> Dict[str, str]:
    """Load base/table IDs from env-vars or fall back to cached config."""
    config_file = _CACHE_DIR / "config.json"
    disk_config: Dict[str, str] = {}
    if config_file.exists():
        with config_file.open() as fh:
            disk_config = json.load(fh)

    base = os.environ.get("BROCHURE_AIRTABLE_BASE") or disk_config.get("base_id", "")
    table = os.environ.get("BROCHURE_AIRTABLE_TABLE") or disk_config.get("table_name", "")
    pat = os.environ.get("AIRTABLE_PAT", "")

    if not base:
        raise EnvironmentError(
            "Airtable base ID not set. "
            "Provide the BROCHURE_AIRTABLE_BASE environment variable or "
            "set 'base_id' in Assets/airtable_cache/config.json."
        )
    if not table:
        raise EnvironmentError(
            "Airtable table name not set. "
            "Provide the BROCHURE_AIRTABLE_TABLE environment variable or "
            "set 'table_name' in Assets/airtable_cache/config.json."
        )
    if not pat:
        raise EnvironmentError(
            "AIRTABLE_PAT environment variable is not set. "
            "Never commit API keys -- use environment variables or server secrets."
        )

    return {"base": base, "table": table, "pat": pat}


def _cache_is_fresh() -> bool:
    """Return True if the on-disk cache exists and is younger than TTL."""
    if not _META_FILE.exists():
        return False
    meta = json.loads(_META_FILE.read_text())
    return (time.time() - meta.get("fetched_at", 0)) < _CACHE_TTL_SECONDS


def fetch_records(force_refresh: bool = False) -> List[Dict[str, Any]]:
    """
    Return all Airtable records for the configured base/table.

    Uses a local JSON cache to avoid unnecessary API calls.  Pass
    ``force_refresh=True`` to bypass the cache.
    """
    _CACHE_DIR.mkdir(parents=True, exist_ok=True)

    if not force_refresh and _cache_is_fresh() and _CACHE_FILE.exists():
        with _CACHE_FILE.open() as fh:
            return json.load(fh)

    config = _get_config()
    headers = {"Authorization": "Bearer " + config["pat"]}
    url = _AIRTABLE_API + "/" + config["base"] + "/" + config["table"]

    records: List[Dict[str, Any]] = []
    params: Dict[str, Any] = {}
    while True:
        response = requests.get(url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        data = response.json()
        records.extend(data.get("records", []))
        offset = data.get("offset")
        if not offset:
            break
        params["offset"] = offset

    _CACHE_FILE.write_text(json.dumps(records, indent=2))
    _META_FILE.write_text(json.dumps({"fetched_at": time.time()}, indent=2))
    return records


def get_record_by_id(record_id: str, force_refresh: bool = False) -> Optional[Dict[str, Any]]:
    """Return a single record dict matching *record_id*, or None."""
    for record in fetch_records(force_refresh=force_refresh):
        if record.get("id") == record_id:
            return record
    return None


def invalidate_cache() -> None:
    """Delete the local cache so the next call fetches fresh data."""
    for path in (_CACHE_FILE, _META_FILE):
        if path.exists():
            path.unlink()
