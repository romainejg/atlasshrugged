"""
Roboflow inference client for Biophi Suite.

Downloads prediction results from a Roboflow-hosted model and caches the
returned bounding-box / classification data alongside the HD image cache.

Required environment variable:
    ROBOFLOW_API_KEY  -- Roboflow API key; never commit this value to source.
"""

import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests

_IMAGE_CACHE_DIR = Path(__file__).parent.parent / "Assets" / "image_cache"
_ROBOFLOW_INFER_URL = "https://detect.roboflow.com"


def _get_api_key() -> str:
    key = os.environ.get("ROBOFLOW_API_KEY", "")
    if not key:
        raise EnvironmentError(
            "ROBOFLOW_API_KEY environment variable is not set. "
            "Never commit API keys -- use environment variables or server secrets."
        )
    return key


def run_inference(
    image_path: str,
    model_id: str,
    confidence: float = 0.4,
    overlap: float = 0.3,
) -> Dict[str, Any]:
    """
    Submit *image_path* to the Roboflow hosted model *model_id* and return
    the raw prediction dict.

    Results are cached to ``Assets/image_cache/<stem>_predictions.json`` so
    repeated calls on the same image do not consume API quota.

    Args:
        image_path: Path to the local image file.
        model_id:   Roboflow model identifier in the form ``project/version``.
        confidence: Minimum confidence threshold (0–1).
        overlap:    Maximum bounding-box overlap threshold (0–1).
    """
    _IMAGE_CACHE_DIR.mkdir(parents=True, exist_ok=True)

    p = Path(image_path)
    cache_file = _IMAGE_CACHE_DIR / (p.stem + "_predictions.json")

    if cache_file.exists():
        with cache_file.open() as fh:
            return json.load(fh)

    api_key = _get_api_key()
    url = _ROBOFLOW_INFER_URL + "/" + model_id

    with open(image_path, "rb") as image_fh:
        response = requests.post(
            url,
            params={
                "api_key": api_key,
                "confidence": int(confidence * 100),
                "overlap": int(overlap * 100),
            },
            files={"file": image_fh},
            timeout=60,
        )
    response.raise_for_status()
    result = response.json()

    cache_file.write_text(json.dumps(result, indent=2))
    return result


def get_cached_predictions(image_stem: str) -> Optional[Dict[str, Any]]:
    """Return cached predictions for *image_stem* (filename without extension)."""
    cache_file = _IMAGE_CACHE_DIR / (image_stem + "_predictions.json")
    if cache_file.exists():
        with cache_file.open() as fh:
            return json.load(fh)
    return None


def list_cached_predictions() -> List[str]:
    """Return the stems of all images that have cached predictions."""
    if not _IMAGE_CACHE_DIR.exists():
        return []
    return [
        f.stem.replace("_predictions", "")
        for f in _IMAGE_CACHE_DIR.glob("*_predictions.json")
    ]
