"""
Image analysis pipeline for Biophi Suite.

Orchestrates fetching Airtable metadata, running Roboflow inference, and
returning structured analysis results ready for deck generation.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from airtable_client import fetch_records
from roboflow_client import get_cached_predictions, run_inference

logger = logging.getLogger(__name__)

_IMAGE_CACHE_DIR = Path(__file__).parent.parent / "Assets" / "image_cache"


def analyse_record(
    record: Dict[str, Any],
    model_id: str,
    image_field: str = "Image",
    confidence: float = 0.4,
) -> Dict[str, Any]:
    """
    Run inference on the image(s) attached to an Airtable *record*.

    Args:
        record:       An Airtable record dict (as returned by airtable_client).
        model_id:     Roboflow model identifier (``project/version``).
        image_field:  Name of the Airtable attachment field containing images.
        confidence:   Minimum confidence threshold forwarded to Roboflow.

    Returns:
        A dict with keys ``record_id``, ``fields``, and ``predictions`` (list).
    """
    record_id: str = record.get("id", "")
    fields: Dict[str, Any] = record.get("fields", {})
    attachments: List[Dict[str, Any]] = fields.get(image_field, [])

    predictions: List[Dict[str, Any]] = []
    for attachment in attachments:
        filename: str = attachment.get("filename", "")
        cached = get_cached_predictions(Path(filename).stem)
        if cached is not None:
            predictions.append({"filename": filename, "result": cached})
            continue

        url: Optional[str] = attachment.get("url")
        if not url:
            logger.warning("Attachment %s has no URL; skipping inference.", filename)
            continue

        local_path = _download_image(url, filename)
        result = run_inference(str(local_path), model_id, confidence=confidence)
        predictions.append({"filename": filename, "result": result})

    return {"record_id": record_id, "fields": fields, "predictions": predictions}


def analyse_all(
    model_id: str,
    image_field: str = "Image",
    confidence: float = 0.4,
    force_airtable_refresh: bool = False,
) -> List[Dict[str, Any]]:
    """
    Fetch every Airtable record and run inference on each attached image.

    Returns a list of result dicts (one per record) in the same format as
    :func:`analyse_record`.
    """
    records = fetch_records(force_refresh=force_airtable_refresh)
    results: List[Dict[str, Any]] = []
    for record in records:
        try:
            results.append(
                analyse_record(record, model_id, image_field=image_field, confidence=confidence)
            )
        except Exception:
            logger.exception("Failed to analyse record %s", record.get("id"))
    return results


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _download_image(url: str, filename: str) -> Path:
    """Download *url* into the image cache and return the local path."""
    import requests  # local import to keep top-level imports light

    _IMAGE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    dest = _IMAGE_CACHE_DIR / filename
    if dest.exists():
        return dest

    logger.info("Downloading %s …", filename)
    response = requests.get(url, timeout=60)
    response.raise_for_status()
    dest.write_bytes(response.content)
    return dest
