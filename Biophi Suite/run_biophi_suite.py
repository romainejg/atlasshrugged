#!/usr/bin/env python3
"""
Biophi Suite — server-compatible (non-GUI) entry point.

Runs the full brochure pipeline from the command line or a web server:
  1. Fetch (or load cached) Airtable variety records.
  2. Download and cache Airtable HD images.
  3. Run Roboflow inference on analysis photos (if an analysis workbook is given).
  4. Build the branded PowerPoint brochure using python-pptx (no Office required).

Required environment variables:
    AIRTABLE_PAT       — Airtable personal access token
    ROBOFLOW_API_KEY   — Roboflow API key

Optional environment variables:
    BROCHURE_AIRTABLE_BASE   — Airtable base ID (falls back to cached config)
    BROCHURE_AIRTABLE_TABLE  — Airtable table name (falls back to cached config)

Never store API keys in source files.  Use environment variables, .env files
loaded by a tool such as python-dotenv, or your hosting platform's secret store.
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
from pathlib import Path

# Add App/ to the module search path so sibling modules are importable.
_APP_DIR = Path(__file__).parent / "App"
sys.path.insert(0, str(_APP_DIR))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
_log = logging.getLogger("biophi")


def _check_env() -> None:
    required = ["AIRTABLE_PAT", "ROBOFLOW_API_KEY"]
    missing = [v for v in required if not os.environ.get(v, "").strip()]
    if missing:
        _log.error(
            "Missing required environment variables: %s\n"
            "Set them via your shell, a .env file, or your hosting platform's secret store.\n"
            "Never commit API keys to source code.",
            ", ".join(missing),
        )
        sys.exit(1)


def _parse_args(argv=None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Run the Biophi brochure pipeline (server-compatible, no GUI)."
    )
    p.add_argument(
        "--analysis",
        metavar="WORKBOOK.xlsx",
        help="Optional Biophi Analysis output workbook to include trial data.",
    )
    p.add_argument(
        "--template",
        metavar="template.pptx",
        default=str(Path(__file__).parent / "template.pptx"),
        help="Branded slide template (default: ./template.pptx).",
    )
    p.add_argument(
        "--output-dir",
        metavar="DIR",
        default=str(Path(__file__).parent / "output"),
        help="Directory for generated files (default: ./output).",
    )
    p.add_argument(
        "--trial-name",
        metavar="NAME",
        default="",
        help="Trial name to embed in the brochure.",
    )
    p.add_argument(
        "--force-airtable-refresh",
        action="store_true",
        help="Bypass cached Airtable data and fetch fresh records.",
    )
    return p.parse_args(argv)


def main(argv=None) -> int:
    _check_env()
    args = _parse_args(argv)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    template = Path(args.template)
    if not template.is_file():
        _log.error("Template not found: %s", template)
        return 1

    analysis_workbook = Path(args.analysis) if args.analysis else None
    if analysis_workbook and not analysis_workbook.is_file():
        _log.error("Analysis workbook not found: %s", analysis_workbook)
        return 1

    from brochure_maker import build_brochure

    _log.info("Starting brochure pipeline…")
    try:
        data_file, deck = build_brochure(
            analysis_workbook,
            template,
            output_dir,
            args.trial_name,
        )
        _log.info("Airtable data saved: %s", data_file)
        _log.info("Brochure deck saved: %s", deck)
    except Exception as exc:
        _log.exception("Pipeline failed: %s", exc)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
