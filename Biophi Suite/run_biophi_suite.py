#!/usr/bin/env python3
"""
Biophi Suite — main entry point.

Run this script to execute the full pipeline:
  1. Fetch (or load cached) Airtable records.
  2. Run Roboflow inference on attached images.
  3. Generate a branded PowerPoint brochure.

Environment variables (required):
    ROBOFLOW_API_KEY       -- Roboflow API key
    AIRTABLE_PAT           -- Airtable personal access token
    ROBOFLOW_MODEL_ID      -- Roboflow model in the form "project/version"

Environment variables (optional):
    BROCHURE_AIRTABLE_BASE  -- Airtable base ID
    BROCHURE_AIRTABLE_TABLE -- Airtable table name
    BIOPHI_OUTPUT_DIR       -- Override output directory (default: ./output)

Never store API keys in this file or any other source file.
"""

import logging
import os
import sys
from pathlib import Path

# Ensure App/ is on the path so sibling modules resolve correctly.
_APP_DIR = Path(__file__).parent / "App"
sys.path.insert(0, str(_APP_DIR))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-8s %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("biophi")


def _check_env() -> None:
    missing = [v for v in ("ROBOFLOW_API_KEY", "AIRTABLE_PAT", "ROBOFLOW_MODEL_ID") if not os.environ.get(v)]
    if missing:
        logger.error(
            "Missing required environment variables: %s\n"
            "Set them in your shell, a .env file (loaded externally), or as server secrets.\n"
            "Never commit API keys to source code.",
            ", ".join(missing),
        )
        sys.exit(1)


def main() -> int:
    _check_env()

    output_dir = Path(os.environ.get("BIOPHI_OUTPUT_DIR", Path(__file__).parent / "output"))
    output_dir.mkdir(parents=True, exist_ok=True)

    results_path = output_dir / "results.json"
    deck_path = output_dir / "brochure.pptx"

    # Step 1 + 2: Fetch records and run inference.
    logger.info("Running analysis pipeline …")
    from analysis import analyse_all
    import json

    model_id = os.environ["ROBOFLOW_MODEL_ID"]
    results = analyse_all(model_id=model_id)
    results_path.write_text(json.dumps(results, indent=2))
    logger.info("Analysis complete. %d records processed.", len(results))

    # Step 3: Build the deck.
    logger.info("Building brochure deck …")
    from build_deck import build_deck
    out = build_deck(results, output_path=str(deck_path))
    logger.info("Done. Brochure written to: %s", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
