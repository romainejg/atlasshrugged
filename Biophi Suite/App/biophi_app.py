from __future__ import annotations

import math
import os
import re
import traceback
import base64
import json
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from io import BytesIO
from collections import deque
from copy import copy
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import numpy as np
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image, ImageDraw, ImageFilter


GRID_SQUARE_CM = 1.27  # 1/2 inch, supplied by the user
IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}


def clean_name(value: Any) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\xa0", " ").strip()
    text = re.sub(r"\s*\(\s*C116\??\s*\)\s*", "", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def key_name(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", clean_name(value).lower())


def number(value: Any) -> float | None:
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)) and math.isfinite(float(value)):
        return float(value)
    return None


def find_sheets_by_day(workbook) -> list[tuple[int, Any]]:
    candidates = []
    for sheet in workbook.worksheets:
        match = re.search(r"(?<!\d)(\d{1,3})\s*days?(?:\s*harvest)?\b", sheet.title, re.I)
        if match:
            candidates.append((int(match.group(1)), sheet))
    return sorted(candidates, key=lambda item: item[0])


def trial_info(workbook) -> dict[str, Any]:
    result: dict[str, Any] = {}
    sheet = next((s for s in workbook.worksheets if "trial info" in s.title.lower()), None)
    if sheet is None:
        return result
    for row in sheet.iter_rows(min_col=1, max_col=min(sheet.max_column, 8)):
        label = clean_name(row[0].value)
        if label:
            result[key_name(label)] = next(
                (cell.value for cell in row[1:] if cell.value not in (None, "", "\xa0")), None
            )
    return result


def evaluation_rows(path: Path) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    formula_book = load_workbook(path, data_only=False)
    values_book = load_workbook(path, data_only=True)
    formula_sheets = find_sheets_by_day(formula_book)
    value_sheets = dict(find_sheets_by_day(values_book))
    if not formula_sheets:
        raise ValueError("No harvest worksheet with a day number was found (for example, '12 Days' or '23 Day Harvest').")
    info = trial_info(formula_book)
    start_date = info.get("pondingdate") or info.get("sowingdate")
    rows = []
    for sheet_day, sheet in formula_sheets:
        value_sheet = value_sheets.get(sheet_day)
        if value_sheet is None:
            continue
        header_row = next(
            (
                row_index
                for row_index in range(1, min(sheet.max_row, 12) + 1)
                if any(key_name(cell.value) in {"variety", "varietyname"} for cell in sheet[row_index])
            ),
            None,
        )
        if header_row is None:
            continue
        headers = {key_name(cell.value): cell.column for cell in sheet[header_row] if clean_name(cell.value)}

        def col(*names: str) -> int | None:
            return next((headers[key_name(name)] for name in names if key_name(name) in headers), None)

        mapping = {
            "variety": col("Variety Name", "Variety"),
            "plot_size": col("Plot Size (m2) (1 raft)", "Plot Size (m2)"),
            "total_plot_size": col("Total Plot Size (m2) (3 rafts)"),
            "cycle": col("DTM from Ponding", "Cycle length"),
            "raft1": col("Plot Weight (g) (Raft 1)"),
            "raft2": col("Plot Weight (g) (Raft 2)"),
            "raft3": col("Plot Weight (g) (Raft 3)"),
            "extra": col("Extra Weight (g)"),
            "total": col("Total Weight (g)"),
            "germ": col("Germ Rating (1-9)"),
            "uniformity": col("Variety Uniformity Rating (1-9)"),
            "tipburn": col("Tipburn Rating (1-9)"),
            "yellowing": col("Leaf Base Yellowing Rating (1-9)"),
            "root": col("Root Quality Rating (1-9)"),
            "bolting": col("Plant Bolting Rating (1-9)"),
            "glassiness": col("Glassiness (1-9)"),
            "crispiness": col("Crispiness Rating (1-9)"),
            "overall": col("Variety Overall Value Rating (1-9)"),
            "taste": col("Taste Notes"),
            "notes": col("Strong and Weak Points (including any disease notes)"),
        }
        if not mapping["variety"]:
            continue
        for row_index in range(header_row + 1, sheet.max_row + 1):
            name = clean_name(sheet.cell(row_index, mapping["variety"]).value)
            if not name:
                continue
            cycle_value = number(value_sheet.cell(row_index, mapping["cycle"]).value) if mapping["cycle"] else None
            harvest_day = int(cycle_value) if cycle_value is not None else sheet_day
            record = {
                "variety": name,
                "key": key_name(name),
                "harvest_day": harvest_day,
                "harvest_date": (
                    start_date + timedelta(days=harvest_day)
                    if isinstance(start_date, datetime)
                    else None
                ),
            }
            for field, column in mapping.items():
                if field == "variety" or not column:
                    continue
                value = value_sheet.cell(row_index, column).value
                if value is None:
                    value = sheet.cell(row_index, column).value
                    if isinstance(value, str) and value.startswith("="):
                        value = None
                record[field] = value
            if number(record.get("total_plot_size")) is None:
                plot = number(record.get("plot_size"))
                record["total_plot_size"] = plot * 3 if plot is not None else None
            if number(record.get("total")) is None:
                weights = [number(record.get(f)) for f in ("raft1", "raft2", "raft3", "extra")]
                record["total"] = sum(weights) if all(v is not None for v in weights) else None
            rows.append(record)

    days = sorted({record["harvest_day"] for record in rows})
    positive_totals = [
        value
        for record in rows
        if (value := number(record.get("total"))) is not None and value > 0
    ]
    # Some evaluation templates store weights as kilograms even though the
    # column caption says grams, while others (notably brassica) store grams.
    # Detect the workbook convention once so every downstream output uses
    # grams for weight fields and kg for yield calculations.
    weight_to_kg = 0.001 if positive_totals and float(np.median(positive_totals)) > 50 else 1.0
    for record in rows:
        record["_weight_to_kg"] = weight_to_kg
    info["weightunitdetected"] = "g" if weight_to_kg == 0.001 else "kg"
    info["harvestdays"] = days
    info["lastharvestdate"] = max(
        (record["harvest_date"] for record in rows if record["harvest_date"] is not None),
        default=None,
    )
    return rows, info


def shelf_life_days(path: Path) -> dict[str, int]:
    workbook = load_workbook(path, data_only=True)
    observations: dict[str, list[tuple[int, float | None]]] = {}
    for sheet in workbook.worksheets:
        day_match = re.search(r"(\d+)", sheet.title)
        if not day_match:
            continue
        day = int(day_match.group(1))
        for row in range(2, sheet.max_row + 1):
            name = clean_name(sheet.cell(row, 1).value)
            if not name:
                continue
            observations.setdefault(key_name(name), []).append(
                (day, number(sheet.cell(row, 3).value))
            )

    result: dict[str, int] = {}
    for variety, values in observations.items():
        acceptable = [day for day, score in values if score is not None and score >= 6]
        if acceptable:
            result[variety] = max(acceptable)
    return result


def grid_spacing_pixels(rgb: np.ndarray) -> float | None:
    height, width, _ = rgb.shape
    gray = rgb.mean(axis=2)
    region = gray[int(height * 0.12) : int(height * 0.62), int(width * 0.06) : int(width * 0.94)]
    dark = region < 120
    projection = dark.mean(axis=0)
    threshold = max(0.01, float(np.percentile(projection, 94)))
    candidate = np.where(projection >= threshold)[0]
    if candidate.size < 5:
        return None
    groups = []
    start = previous = int(candidate[0])
    for value in candidate[1:]:
        value = int(value)
        if value > previous + 2:
            groups.append((start + previous) / 2)
            start = value
        previous = value
    groups.append((start + previous) / 2)
    gaps = np.diff(groups)
    valid = gaps[(gaps > width * 0.018) & (gaps < width * 0.08)]
    return float(np.median(valid)) if valid.size >= 4 else None


def components(
    mask: np.ndarray, min_area: int | None = None
) -> list[dict[str, int]]:
    height, width = mask.shape
    seen = np.zeros(mask.shape, dtype=bool)
    found = []
    if min_area is None:
        min_area = max(30, int(mask.size * 0.00045))
    for y, x in zip(*np.where(mask & ~seen)):
        if seen[y, x]:
            continue
        queue = deque([(int(y), int(x))])
        seen[y, x] = True
        area = 0
        top = bottom = int(y)
        left = right = int(x)
        while queue:
            cy, cx = queue.popleft()
            area += 1
            top, bottom = min(top, cy), max(bottom, cy)
            left, right = min(left, cx), max(right, cx)
            for ny, nx in ((cy - 1, cx), (cy + 1, cx), (cy, cx - 1), (cy, cx + 1)):
                if 0 <= ny < height and 0 <= nx < width and mask[ny, nx] and not seen[ny, nx]:
                    seen[ny, nx] = True
                    queue.append((ny, nx))
        if area >= min_area:
            found.append({
                "area": area,
                "left": left,
                "top": top,
                "right": right,
                "bottom": bottom,
                "height": bottom - top + 1,
                "width": right - left + 1,
            })
    return found


def stem_anchored_leaf_boxes(
    mask: np.ndarray, spacing: float, reduction: int
) -> tuple[list[dict[str, Any]], list[dict[str, int]], list[str]]:
    height, width = mask.shape
    pixels_per_cm = spacing / reduction / GRID_SQUARE_CM
    foreground_y, _ = np.where(mask)
    if not len(foreground_y):
        return [], [], ["No distinct stem bases detected"]

    # Leaves are placed side-by-side and do not overlap at the bottom.  A
    # shallow band immediately above the common baseline therefore contains
    # one disconnected component per physical leaf, even when the blades touch
    # above it.
    foreground_groups = [
        group for group in components(mask) if group["bottom"] > height * 0.55
    ]
    baseline = (
        int(round(np.median([group["bottom"] for group in foreground_groups])))
        if foreground_groups
        else int(np.percentile(foreground_y, 99.5))
    )
    band_depth = max(4, round(1.2 * pixels_per_cm))
    band_top = max(int(height * 0.65), baseline - band_depth)
    base_components: list[dict[str, int]] = []
    band = mask[band_top : baseline + 1, :]
    for component in components(
        band, min_area=max(3, round(pixels_per_cm * 0.5))
    ):
        component = dict(component)
        component["top"] += band_top
        component["bottom"] += band_top
        component["height"] = component["bottom"] - component["top"] + 1
        component_width_cm = component["width"] / pixels_per_cm
        component_height_cm = component["height"] / pixels_per_cm
        if component_width_cm > 4.5 or component_height_cm < 0.35:
            continue
        base_components.append(component)

    # Some specimens are set slightly above the common baseline.  Recover
    # those from a deeper band, but only when the component does not contain an
    # already-detected shallow base.  Existing shallow bases therefore remain
    # the authoritative split for close leaves.
    deep_top = max(int(height * 0.60), baseline - round(3.0 * pixels_per_cm))
    deep_band = mask[deep_top : baseline + 1, :]
    for component in components(
        deep_band, min_area=max(4, round(pixels_per_cm * 0.7))
    ):
        component = dict(component)
        component["top"] += deep_top
        component["bottom"] += deep_top
        component["height"] = component["bottom"] - component["top"] + 1
        component_width_cm = component["width"] / pixels_per_cm
        component_height_cm = component["height"] / pixels_per_cm
        if component_width_cm > 4.5 or component_height_cm < 0.7:
            continue
        if any(
            component["left"] - 1
            <= (base["left"] + base["right"]) / 2
            <= component["right"] + 1
            for base in base_components
        ):
            continue
        base_components.append(component)

    # If a raised stem has already touched a neighbour's blade within the
    # deeper band, it is no longer a separate component.  Recover only strong
    # downward tips located well away from every existing base.  The distance
    # guard prevents serrated blade edges from becoming extra leaves.
    envelope = np.full(width, -1, dtype=int)
    support = np.zeros(width, dtype=int)
    support_height = max(4, round(1.2 * pixels_per_cm))
    for x in range(width):
        ys = np.where(mask[deep_top : baseline + 1, x])[0]
        if not len(ys):
            continue
        tip = deep_top + int(ys.max())
        envelope[x] = tip
        support_top = max(deep_top, tip - support_height + 1)
        support[x] = int(
            mask[support_top : tip + 1, max(0, x - 1) : min(width, x + 2)].sum()
        )
    peak_radius = max(3, round(1.0 * pixels_per_cm))
    existing_centers = [
        (base["left"] + base["right"]) / 2 for base in base_components
    ]
    candidate_columns = []
    for x, tip in enumerate(envelope):
        if tip < 0 or support[x] < support_height:
            continue
        if existing_centers and min(abs(x - center) for center in existing_centers) < 2.3 * pixels_per_cm:
            continue
        left, right = max(0, x - peak_radius), min(width, x + peak_radius + 1)
        if tip >= int(envelope[left:right].max()) - 1:
            candidate_columns.append(x)
    candidate_runs: list[list[int]] = []
    for x in candidate_columns:
        if not candidate_runs or x > candidate_runs[-1][-1] + 1:
            candidate_runs.append([x])
        else:
            candidate_runs[-1].append(x)
    for run in candidate_runs:
        x = max(run, key=lambda column: (support[column], envelope[column]))
        if existing_centers and min(abs(x - center) for center in existing_centers) < 2.3 * pixels_per_cm:
            continue
        tip = int(envelope[x])
        seed_top = max(deep_top, tip - support_height + 1)
        seed_left, seed_right = max(0, x - 1), min(width - 1, x + 1)
        seed = mask[seed_top : tip + 1, seed_left : seed_right + 1]
        seed_y, seed_x = np.where(seed)
        if not len(seed_y):
            continue
        left = seed_left + int(seed_x.min())
        right = seed_left + int(seed_x.max())
        top = seed_top + int(seed_y.min())
        bottom = seed_top + int(seed_y.max())
        base_components.append({
            "area": int(len(seed_y)),
            "left": left,
            "top": top,
            "right": right,
            "bottom": bottom,
            "height": bottom - top + 1,
            "width": right - left + 1,
        })
        existing_centers.append((left + right) / 2)
    base_components.sort(key=lambda item: (item["left"] + item["right"]) / 2)
    # Suppress duplicate anchors created by different depth passes while
    # preserving genuinely close neighbouring stems.
    separated_bases: list[dict[str, int]] = []
    minimum_separation = 1.5 * pixels_per_cm
    for base in base_components:
        center = (base["left"] + base["right"]) / 2
        if separated_bases:
            prior_center = (
                separated_bases[-1]["left"] + separated_bases[-1]["right"]
            ) / 2
            if center - prior_center < minimum_separation:
                continue
        separated_bases.append(base)
    base_components = separated_bases

    issues = []
    if not base_components:
        return [], [], ["No distinct stem bases detected"]
    if len(base_components) > 20:
        issues.append("Implausibly many stem bases")
    for base in base_components:
        base_width_cm = base["width"] * reduction / spacing * GRID_SQUARE_CM
        if base_width_cm > 5.0:
            issues.append("A stem base is too wide to confirm as one leaf")
            break

    base_centers = np.array([
        (base["left"] + base["right"]) / 2 for base in base_components
    ])
    # Grow one region upward from every confirmed stem.  The temporary closing
    # bridges grid/highlight gaps, while the final dimensions are still taken
    # only from original foreground pixels.  Multi-source growth separates
    # touching blades by their path back to the correct stem instead of by a
    # rigid vertical cut.
    growth_image = Image.fromarray((mask * 255).astype(np.uint8))
    growth_image = growth_image.filter(ImageFilter.MaxFilter(5)).filter(
        ImageFilter.MinFilter(5)
    )
    growth_mask = np.asarray(growth_image) > 0
    labels = np.full(mask.shape, -1, dtype=np.int16)
    distances = np.full(mask.shape, np.iinfo(np.int32).max, dtype=np.int32)
    queue = deque()
    for label, base in enumerate(base_components):
        top, bottom = max(0, base["top"] - 1), min(height - 1, base["bottom"] + 1)
        left, right = max(0, base["left"] - 1), min(width - 1, base["right"] + 1)
        seed_y, seed_x = np.where(growth_mask[top : bottom + 1, left : right + 1])
        for local_y, local_x in zip(seed_y, seed_x):
            y, x = top + int(local_y), left + int(local_x)
            labels[y, x] = label
            distances[y, x] = 0
            queue.append((y, x))
    while queue:
        y, x = queue.popleft()
        next_distance = int(distances[y, x]) + 1
        for ny, nx in ((y - 1, x), (y + 1, x), (y, x - 1), (y, x + 1)):
            if not (
                0 <= ny < height
                and 0 <= nx < width
                and growth_mask[ny, nx]
            ):
                continue
            if next_distance < distances[ny, nx]:
                distances[ny, nx] = next_distance
                labels[ny, nx] = labels[y, x]
                queue.append((ny, nx))

    lane_edges = [0]
    lane_edges.extend(
        int(round((base_centers[index] + base_centers[index + 1]) / 2))
        for index in range(len(base_centers) - 1)
    )
    lane_edges.append(width)
    boxes = []
    for label, base in enumerate(base_components):
        ys, xs = np.where((labels == label) & mask)
        candidates: list[tuple[int, int, int, int]] = []
        if len(ys):
            candidates.append((
                int(xs.min()),
                int(ys.min()),
                int(xs.max()),
                int(ys.max()),
            ))

        # Conservative fallback for a strongly leaning leaf whose highlighted
        # petiole still has a gap after closing.  It cannot create another leaf
        # because it reuses the same confirmed base and midpoint lane.
        lane_left, lane_right = lane_edges[label], lane_edges[label + 1]
        lane_components = components(mask[:, lane_left:lane_right], min_area=3)
        if lane_components:
            base_center = float(base_centers[label] - lane_left)
            selected = min(
                lane_components,
                key=lambda item: (
                    0
                    if item["left"] - 1 <= base_center <= item["right"] + 1
                    else min(
                        abs(base_center - item["left"]),
                        abs(base_center - item["right"]),
                    ),
                    abs(item["bottom"] - base["bottom"]),
                    -item["area"],
                ),
            )
            connected = dict(selected)
            remaining = [
                item for item in lane_components if item is not selected
            ]
            initial_length = max(
                [
                    (bounds[3] - bounds[1] + 1) / pixels_per_cm
                    for bounds in candidates
                ]
                + [connected["height"] / pixels_per_cm]
            )
            changed = True
            while changed:
                changed = False
                for item in list(remaining):
                    horizontal_gap = max(
                        0,
                        item["left"] - connected["right"] - 1,
                        connected["left"] - item["right"] - 1,
                    )
                    vertical_gap = max(
                        0,
                        item["top"] - connected["bottom"] - 1,
                        connected["top"] - item["bottom"] - 1,
                    )
                    center_gap = abs(
                        (item["left"] + item["right"]) / 2
                        - (connected["left"] + connected["right"]) / 2
                    )
                    short_highlight_gap = (
                        horizontal_gap <= 1 and vertical_gap <= 3
                    )
                    aligned_petiole_gap = (
                        2.0 <= initial_length < 3.0
                        and
                        center_gap <= 1.0 * pixels_per_cm
                        and vertical_gap <= 2.5 * pixels_per_cm
                    )
                    if short_highlight_gap or aligned_petiole_gap:
                        connected["left"] = min(connected["left"], item["left"])
                        connected["right"] = max(connected["right"], item["right"])
                        connected["top"] = min(connected["top"], item["top"])
                        connected["bottom"] = max(connected["bottom"], item["bottom"])
                        remaining.remove(item)
                        changed = True
            candidates.append((
                lane_left + connected["left"],
                connected["top"],
                lane_left + connected["right"],
                connected["bottom"],
            ))
        if not candidates:
            continue
        # Prefer the candidate with the complete vertical extent; if tied, use
        # the narrower ownership to avoid borrowing a neighbour's blade.
        left, top, right, bottom = max(
            candidates,
            key=lambda bounds: (
                bounds[3] - bounds[1],
                -(bounds[2] - bounds[0]),
            ),
        )
        leaf_height = bottom - top + 1
        leaf_width = right - left + 1
        length_cm = leaf_height * reduction / spacing * GRID_SQUARE_CM
        width_cm = leaf_width * reduction / spacing * GRID_SQUARE_CM
        if length_cm < 2.0:
            continue
        boxes.append({
            "left": left,
            "top": top,
            "right": right,
            "bottom": bottom,
            "length": round(length_cm, 1),
            "width": round(width_cm, 1),
            "base": base,
        })

    crop_line = int(height * 0.40)
    if any(box["top"] <= crop_line + 1 for box in boxes):
        issues.append("Foreground touches the analysis boundary")
    if any(box["length"] > 35 or box["width"] > 25 for box in boxes):
        issues.append("Implausible leaf dimensions")
    return boxes, base_components, issues


def roboflow_leaf_boxes(
    path: Path, scale: float, spacing: float, reduction: int
) -> tuple[list[dict[str, Any]] | None, list[str]]:
    """Return one box per Roboflow leaf instance, or None when not configured.

    The API key is intentionally read only from the machine environment; it is
    never written to this application, the workbook, or the photo audit.
    """
    api_key = os.environ.get("ROBOFLOW_API_KEY", "").strip()
    if not api_key:
        return None, []
    try:
        workspace = os.environ.get("ROBOFLOW_WORKSPACE", "").strip()
        if not workspace:
            auth_request = Request(
                f"https://api.roboflow.com/?{urlencode({'api_key': api_key})}",
                headers={"Accept": "application/json"},
            )
            with urlopen(auth_request, timeout=30) as auth_result:
                workspace = str(json.loads(auth_result.read().decode("utf-8")).get("workspace", "")).strip()
        if not workspace:
            return None, ["Roboflow authenticated but did not return a workspace ID"]
        workflow_id = os.environ.get("ROBOFLOW_WORKFLOW_ID", "leafy").strip()
        endpoint = f"https://serverless.roboflow.com/{workspace}/workflows/{workflow_id}"
        image_body = base64.b64encode(path.read_bytes()).decode("ascii")
        payload = json.dumps({
            "api_key": api_key,
            "inputs": {"image": {"type": "base64", "value": image_body}},
            "use_cache": False,
        }).encode("utf-8")
        request = Request(
            endpoint,
            data=payload,
            headers={"Content-Type": "application/json", "Accept": "application/json"},
            method="POST",
        )
        with urlopen(request, timeout=90) as result:
            response = json.loads(result.read().decode("utf-8"))
    except Exception as error:
        return None, [f"Roboflow inference failed: {str(error)[:120]}"]

    def find_predictions(value: Any) -> list[dict[str, Any]] | None:
        if isinstance(value, dict):
            predictions = value.get("predictions")
            if isinstance(predictions, list):
                return predictions
            for nested in value.values():
                found = find_predictions(nested)
                if found is not None:
                    return found
        elif isinstance(value, list):
            for nested in value:
                found = find_predictions(nested)
                if found is not None:
                    return found
        return None

    predictions = find_predictions(response)
    if predictions is None:
        return None, ["Leafy workflow response contained no instance predictions"]
    pixels_per_cm = spacing / reduction / GRID_SQUARE_CM
    boxes: list[dict[str, Any]] = []
    for prediction in predictions:
        points = prediction.get("points") or prediction.get("polygon") or []
        if points:
            xs = [float(point["x"] if isinstance(point, dict) else point[0]) for point in points]
            ys = [float(point["y"] if isinstance(point, dict) else point[1]) for point in points]
            left, right, top, bottom = min(xs), max(xs), min(ys), max(ys)
        else:
            x, y = float(prediction.get("x", 0)), float(prediction.get("y", 0))
            width, height = float(prediction.get("width", 0)), float(prediction.get("height", 0))
            left, right, top, bottom = x - width / 2, x + width / 2, y - height / 2, y + height / 2
        # Predictions use the original image coordinate system; measurements
        # below use the resized working image and then its reduced mask scale.
        left, right = left * scale / reduction, right * scale / reduction
        top, bottom = top * scale / reduction, bottom * scale / reduction
        leaf_height, leaf_width = bottom - top, right - left
        length_cm, width_cm = leaf_height / pixels_per_cm, leaf_width / pixels_per_cm
        if length_cm < 2.0 or leaf_width <= 0:
            continue
        boxes.append({
            "left": int(round(left)), "top": int(round(top)),
            "right": int(round(right)), "bottom": int(round(bottom)),
            "length": round(length_cm, 1), "width": round(width_cm, 1),
            "base": {"left": int(round((left + right) / 2 - 1)), "right": int(round((left + right) / 2 + 1)), "top": int(round(bottom - 2)), "bottom": int(round(bottom)), "width": 3},
        })
    if not boxes:
        return [], ["Roboflow detected no measurable leaf instances"]
    boxes.sort(key=lambda item: (item["left"] + item["right"]) / 2)
    return boxes, []


def measure_leaf(path: Path) -> dict[str, Any] | None:
    with Image.open(path) as source:
        source = source.convert("RGB")
        scale = min(1.0, 800 / max(source.size))
        if scale < 1:
            source = source.resize(
                (round(source.width * scale), round(source.height * scale)), Image.Resampling.LANCZOS
            )
        rgb = np.asarray(source, dtype=np.float32)

    spacing = grid_spacing_pixels(rgb)
    if spacing is None or spacing <= 0:
        return None

    reduction = 4
    boxes, issues = roboflow_leaf_boxes(path, scale, spacing, reduction)
    cloud_issues = list(issues)
    measurement_method = "Roboflow Workflow Leafy" if boxes is not None else "Local fallback"
    bases: list[dict[str, int]] = []
    if boxes is None:
        issues = []

    if boxes is None:
        red, green, blue = rgb[:, :, 0], rgb[:, :, 1], rgb[:, :, 2]
        maximum = np.maximum(np.maximum(red, green), blue)
        minimum = np.minimum(np.minimum(red, green), blue)
        saturation = (maximum - minimum) / np.maximum(maximum, 1)
        # Color-agnostic fallback for photos that cannot use Roboflow.
        chroma = maximum - minimum
        blue_marker = (blue > red + 15) & (blue >= green * 0.95)
        vivid_leaf = (saturation > 0.18) & (chroma > 22)
        green_dark = (green > red + 4) & (green > blue + 4)
        red_dark = (red > green + 5) & (red >= blue - 5)
        dark_leaf = ((maximum < 150) & (saturation > 0.06) & (chroma > 7) & (green_dark | red_dark))
        leaf = (maximum > 35) & vivid_leaf & ~blue_marker
        leaf[: int(leaf.shape[0] * 0.40), :] = False
        leaf[:, : int(leaf.shape[1] * 0.09)] = False
        leaf[:, int(leaf.shape[1] * 0.91) :] = False
        if int(leaf.sum()) < max(200, int(leaf.size * 0.003)):
            leaf = (maximum > 35) & (vivid_leaf | dark_leaf) & ~blue_marker
            leaf[: int(leaf.shape[0] * 0.40), :] = False
            leaf[:, : int(leaf.shape[1] * 0.09)] = False
            leaf[:, int(leaf.shape[1] * 0.91) :] = False
        mask_image = Image.fromarray((leaf * 255).astype(np.uint8))
        mask_image = mask_image.filter(ImageFilter.MaxFilter(3)).filter(ImageFilter.MinFilter(3))
        reduced = mask_image.resize(
            (max(1, mask_image.width // reduction), max(1, mask_image.height // reduction)),
            Image.Resampling.NEAREST,
        )
        boxes, bases, issues = stem_anchored_leaf_boxes(np.asarray(reduced) > 0, spacing, reduction)
    largest_index = max(range(len(boxes)), key=lambda index: boxes[index]["length"]) if boxes else None
    largest = boxes[largest_index] if largest_index is not None else None
    if largest and not (2.0 <= largest["length"] <= 100 and 0.5 <= largest["width"] <= 100):
        issues.append("Largest leaf dimensions are implausible")
    original_stream = BytesIO()
    source.save(original_stream, format="JPEG", quality=78, optimize=True)
    original_stream.seek(0)
    overlay_image = source.copy()
    draw = ImageDraw.Draw(overlay_image)
    line_width = max(4, round(max(overlay_image.size) / 250))
    for index, leaf_box in enumerate(boxes):
        resized_box = tuple(leaf_box[key] * reduction for key in ("left", "top", "right", "bottom"))
        color = (255, 30, 30) if index == largest_index else (20, 90, 255)
        draw.rectangle(resized_box, outline=color, width=line_width)
        draw.text(
            (resized_box[0] + 3, max(0, resized_box[1] - 16)),
            f"{index + 1}: {leaf_box['length']:.1f} cm",
            fill=color,
            stroke_width=2,
            stroke_fill=(255, 255, 255),
        )
        base = leaf_box["base"]
        base_box = tuple(base[key] * reduction for key in ("left", "top", "right", "bottom"))
        draw.ellipse(base_box, outline=(255, 215, 0), width=max(2, line_width - 1))
    overlay_stream = BytesIO()
    overlay_image.save(overlay_stream, format="JPEG", quality=78, optimize=True)
    overlay_stream.seek(0)
    if not boxes:
        issues.append("No leaf could be assigned confidently to a stem base")
        draw.rectangle((10, 10, 500, 48), fill=(255, 255, 255), outline=(255, 140, 0), width=3)
        draw.text((18, 20), "REVIEW REQUIRED — automated measurement withheld", fill=(210, 90, 0))
    review_required = bool(issues)
    return {
        "length": None if review_required or largest is None else largest["length"],
        "width": None if review_required or largest is None else largest["width"],
        "leaf_count": len(boxes),
        "leaves": boxes,
        "status": "Review required" if review_required else "OK",
        "method": measurement_method,
        "cloud_issues": cloud_issues,
        "issues": issues,
        "path": path,
        "original_stream": original_stream,
        "overlay_stream": overlay_stream,
    }


def photo_measurements(folder: Path, harvest_days: list[int]) -> dict[int, dict[str, dict[str, Any]]]:
    results: dict[int, dict[str, dict[str, Any]]] = {day: {} for day in harvest_days}
    accepted_days = set(harvest_days)
    for path in folder.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in IMAGE_SUFFIXES:
            continue
        if re.search(r"off\s*type", path.stem, re.I):
            continue
        relative = path.relative_to(folder)
        harvest_day = None
        # Prefer directory labels ("12", "Day 12", "12 DTM", etc.).
        for part in reversed(relative.parts[:-1]):
            numbers = [int(value) for value in re.findall(r"(?<!\d)(\d{1,3})(?!\d)", part)]
            harvest_day = next((value for value in numbers if value in accepted_days), None)
            if harvest_day is not None:
                break
        # Also support flat photo folders when the filename contains a DTM.
        if harvest_day is None:
            numbers = [int(value) for value in re.findall(r"(?<!\d)(\d{1,3})(?!\d)", path.stem)]
            harvest_day = next((value for value in numbers if value in accepted_days), None)
        if harvest_day is None:
            continue
        base_stem = re.sub(r"[-_]+\d+$", "", path.stem)
        key = key_name(base_stem)
        measured = measure_leaf(path)
        if measured is not None:
            day_results = results[harvest_day]
            previous = day_results.get(key)
            measured_score = (1 if measured.get("status") == "OK" else 0, measured.get("length") or 0)
            previous_score = (1 if previous and previous.get("status") == "OK" else 0, previous.get("length") or 0 if previous else 0)
            if previous is None or measured_score > previous_score:
                day_results[key] = measured
    return results


def photo_key_for(variety_key: str, available: dict[str, Any]) -> str | None:
    candidates = [variety_key]
    if variety_key == "crismari":
        candidates.append("c116")
    for candidate in candidates:
        if candidate in available:
            return candidate
        prefix = next((key for key in available if key.startswith(candidate)), None)
        if prefix:
            return prefix
    return None


def copy_cell_style(source, target) -> None:
    if source.has_style:
        target._style = copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    target.alignment = copy(source.alignment)


def weight_kg(record: dict[str, Any], field: str = "total") -> float | None:
    value = number(record.get(field))
    if value is None:
        return None
    factor = number(record.get("_weight_to_kg"))
    if factor is None:
        factor = 0.001 if abs(value) > 50 else 1.0
    return value * factor


def weight_grams(record: dict[str, Any], field: str = "total") -> float | None:
    value = weight_kg(record, field)
    return round(value * 1000, 2) if value is not None else None


def annual_yield(record: dict[str, Any]) -> float | None:
    total = weight_kg(record)
    area = number(record.get("total_plot_size"))
    cycle = number(record.get("cycle"))
    return round(total / area * 365 / cycle, 2) if total is not None and area and cycle else None


def comparative_points(records: list[dict[str, Any]]) -> dict[str, tuple[str | None, str | None]]:
    trait_labels = {
        "overall": "overall value",
        "uniformity": "uniformity",
        "germ": "germination",
        "root": "root quality",
        "crispiness": "crispiness",
        "tipburn": "tipburn resistance",
        "yellowing": "yellowing resistance",
        "glassiness": "glassiness resistance",
        "bolting": "bolting resistance",
    }
    yields = [annual_yield(record) for record in records]
    valid_yields = [value for value in yields if value is not None]
    yield_median = float(np.median(valid_yields)) if valid_yields else None
    medians = {}
    for field in trait_labels:
        values = [number(record.get(field)) for record in records]
        valid = [value for value in values if value is not None]
        medians[field] = float(np.median(valid)) if valid and max(valid) != min(valid) else None

    result = {}
    for record in records:
        pros, cons = [], []
        value = annual_yield(record)
        if value is not None and yield_median:
            difference = (value - yield_median) / yield_median
            if difference >= 0.10:
                pros.append("higher yield")
            elif difference <= -0.10:
                cons.append("lower yield")
        for field, label in trait_labels.items():
            score = number(record.get(field))
            median = medians[field]
            if score is None or median is None:
                continue
            if score >= median + 1:
                pros.append(f"strong {label}")
            elif score <= median - 1:
                cons.append(f"weak {label}")
        result[record["key"]] = (
            "; ".join(pros) if pros else None,
            "; ".join(cons) if cons else None,
        )
    return result


def add_tdm_sheet(sheet, records: list[dict[str, Any]], harvest_days: list[int], photos, shelf) -> None:
    tipburn_percent = {9: 0, 7: 25, 5: 50, 3: 75, 1: 100}
    varieties = []
    for record in records:
        if record["key"] not in [item["key"] for item in varieties]:
            varieties.append(record)
    trait_headers = [
        "Leaf 3D None to extreme",
        "Leaf Colour Green Light to dark",
        "Leaf Colouring Red Percentage class",
        "Leaf Length Harvested Centimeter (cm)",
        "Leaf margin Yellowing Extreme to none",
        "Leaf Thickness Thin to thick",
        "Leaf Tipburn Plot Affection class",
        "Leaf Twisting Extreme to none",
        "Leaf Width Centimeter (cm)",
        "Plant Bolting Index",
        "Plant Shelflife Days (days)",
        "Plot Size 1M2",
        "Plot Weight Gram (g)",
        "Root Quality Bad to good",
        "Seed Germination Bad to good",
        "Variety Maturity Days (days)",
        "Variety Remark Text",
        "Variety Strong points Text",
        "Variety Uniformity Bad to good",
        "Variety Value Overall Bad to good",
        "Variety Weak points Text",
    ]
    sheet.delete_rows(1, sheet.max_row)
    if sheet.max_column:
        sheet.delete_cols(1, sheet.max_column)
    sheet.sheet_view.showGridLines = False
    sheet["A1"] = "Harvest"
    sheet["A2"] = "Date"
    sheet.column_dimensions["A"].width = 13
    header_fill = PatternFill("solid", fgColor="1F4E78")
    sub_fill = PatternFill("solid", fgColor="D9EAF7")
    sheet["A1"].fill = header_fill
    sheet["A1"].font = Font(color="FFFFFF", bold=True)
    sheet["A2"].fill = sub_fill
    sheet["A2"].font = Font(bold=True)

    block_width = len(trait_headers)
    for variety_index, variety in enumerate(varieties):
        start = 2 + variety_index * block_width
        end = start + block_width - 1
        sheet.merge_cells(start_row=1, start_column=start, end_row=1, end_column=end)
        cell = sheet.cell(1, start, variety["variety"])
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
        for offset, header in enumerate(trait_headers):
            target = sheet.cell(2, start + offset, header)
            target.fill = sub_fill
            target.font = Font(bold=True)
            target.alignment = Alignment(wrap_text=True, vertical="top")
            sheet.column_dimensions[target.column_letter].width = 15

    by_day_key = {(record["harvest_day"], record["key"]): record for record in records}
    for row_index, day in enumerate(harvest_days, start=3):
        day_records = [record for record in records if record["harvest_day"] == day]
        harvest_date = next(
            (record["harvest_date"] for record in day_records if record["harvest_date"] is not None),
            None,
        )
        sheet.cell(row_index, 1, harvest_date)
        sheet.cell(row_index, 1).number_format = "yyyy-mm-dd"
        for variety_index, variety in enumerate(varieties):
            record = by_day_key.get((day, variety["key"]))
            if record is None:
                continue
            day_photos = photos.get(day, {})
            photo_key = photo_key_for(record["key"], day_photos)
            measurement = day_photos.get(photo_key) if photo_key else None
            pros, cons = comparative_points(day_records).get(record["key"], (None, None))
            values = [
                None, None, None,
                measurement.get("length") if measurement else None,
                record.get("yellowing"),
                None,
                tipburn_percent.get(int(record["tipburn"])) if number(record.get("tipburn")) is not None else None,
                None,
                measurement.get("width") if measurement else None,
                record.get("bolting"),
                shelf.get(record["key"]),
                record.get("plot_size"),
                weight_grams(record),
                record.get("root"),
                record.get("germ"),
                record.get("cycle"),
                record.get("taste"),
                pros,
                record.get("uniformity"),
                record.get("overall"),
                cons,
            ]
            start = 2 + variety_index * block_width
            for offset, value in enumerate(values):
                sheet.cell(row_index, start + offset, value)
    sheet.freeze_panes = "B3"
    sheet.row_dimensions[1].height = 24
    sheet.row_dimensions[2].height = 80


def add_photo_audit_sheet(workbook, records, photos) -> int:
    if "Output4 - Photo Audit" in workbook.sheetnames:
        del workbook["Output4 - Photo Audit"]
    sheet = workbook.create_sheet("Output4 - Photo Audit")
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A2"
    sheet.column_dimensions["A"].width = 3
    for column in range(2, 9):
        sheet.column_dimensions[sheet.cell(1, column).column_letter].width = 12
    for column in range(10, 17):
        sheet.column_dimensions[sheet.cell(1, column).column_letter].width = 12
    image_count = 0
    row = 1
    for record in records:
        available = photos.get(record["harvest_day"], {})
        photo_key = photo_key_for(record["key"], available)
        measurement = available.get(photo_key) if photo_key else None
        if not measurement:
            continue
        sheet.merge_cells(start_row=row, start_column=2, end_row=row, end_column=8)
        sheet.merge_cells(start_row=row, start_column=10, end_row=row, end_column=16)
        sheet.cell(row, 2, f"{record['variety']} — DTM {record['harvest_day']} — Original")
        if measurement.get("status") == "OK":
            audit_title = (
                f"Measurement overlay — Largest {measurement['length']:.1f} × "
                f"{measurement['width']:.1f} cm — Leaf count {measurement['leaf_count']} "
                f"(gold = stem base, red = largest, blue = other leaves ≥2 cm)"
            )
        else:
            reason = "; ".join(measurement.get("issues", []))
            audit_title = (
                f"REVIEW REQUIRED — Measurement withheld — Suggested leaf count "
                f"{measurement['leaf_count']} — {reason}"
            )
        sheet.cell(row, 10, audit_title)
        for cell in (sheet.cell(row, 2), sheet.cell(row, 10)):
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center")
        original_stream = measurement["original_stream"]
        overlay_stream = measurement["overlay_stream"]
        original_stream.seek(0)
        overlay_stream.seek(0)
        original = ExcelImage(original_stream)
        overlay = ExcelImage(overlay_stream)
        for image in (original, overlay):
            ratio = min(480 / image.width, 340 / image.height)
            image.width *= ratio
            image.height *= ratio
        sheet.add_image(original, f"B{row + 1}")
        sheet.add_image(overlay, f"J{row + 1}")
        for detail_row in range(row + 1, row + 19):
            sheet.row_dimensions[detail_row].height = 15
        row += 20
        image_count += 1
    return image_count


def save_rendered_photo_audits(photos, folder: Path) -> int:
    """Persist the exact Roboflow/audit renders so later brochure runs can reuse them."""
    folder.mkdir(parents=True, exist_ok=True)
    saved = 0
    for day, items in photos.items():
        day_folder = folder / f"DTM {day}"
        day_folder.mkdir(parents=True, exist_ok=True)
        for key, measurement in items.items():
            for kind in ("original_stream", "overlay_stream"):
                stream = measurement.get(kind)
                if stream is None:
                    continue
                suffix = "original" if kind.startswith("original") else "roboflow-overlay"
                target = day_folder / f"{key}_{suffix}.png"
                if target.exists():
                    continue
                stream.seek(0)
                target.write_bytes(stream.read())
                saved += 1
    return saved


def add_environment_summary_sheet(workbook, environment_path: Path, records, info) -> int:
    if "Output5 - Environment" in workbook.sheetnames:
        del workbook["Output5 - Environment"]
    sheet = workbook.create_sheet("Output5 - Environment")
    sheet.sheet_view.showGridLines = False
    source_book = load_workbook(environment_path, data_only=True, read_only=True)
    source = next(
        (item for item in source_book.worksheets if item.title.lower() == "summary"),
        source_book.worksheets[0],
    )
    headers = [clean_name(cell.value) for cell in source[1]]
    date_column = next(
        (index for index, header in enumerate(headers, start=1) if key_name(header) == "date"),
        None,
    )
    rows_by_date = []
    if date_column:
        for row in source.iter_rows(min_row=2, values_only=True):
            date_value = row[date_column - 1]
            if isinstance(date_value, datetime):
                rows_by_date.append((date_value, row))
    source_book.close()

    numeric_columns = []
    for column, header in enumerate(headers, start=1):
        if column == date_column or not header:
            continue
        if any(number(row[column - 1]) is not None for _, row in rows_by_date):
            numeric_columns.append((column, header))

    output_headers = ["Cycle Length (DTM)", "Start Date", "Harvest Date", "Days Included"]
    output_headers.extend(header for _, header in numeric_columns)
    for column, header in enumerate(output_headers, start=1):
        cell = sheet.cell(1, column, header)
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        sheet.column_dimensions[cell.column_letter].width = 18
    sheet.row_dimensions[1].height = 55

    start_date = info.get("pondingdate") or info.get("sowingdate")
    harvest_days = sorted({record["harvest_day"] for record in records})
    for output_row, day in enumerate(harvest_days, start=2):
        harvest_date = next(
            (
                record["harvest_date"]
                for record in records
                if record["harvest_day"] == day and record["harvest_date"] is not None
            ),
            None,
        )
        included = [
            row
            for date_value, row in rows_by_date
            if (not isinstance(start_date, datetime) or date_value >= start_date)
            and (not isinstance(harvest_date, datetime) or date_value <= harvest_date)
        ]
        base_values = [day, start_date, harvest_date, len(included) if included else None]
        means = []
        for source_column, _ in numeric_columns:
            values = [number(row[source_column - 1]) for row in included]
            valid = [value for value in values if value is not None]
            means.append(round(sum(valid) / len(valid), 2) if valid else None)
        for column, value in enumerate(base_values + means, start=1):
            sheet.cell(output_row, column, value)
            if column in (2, 3):
                sheet.cell(output_row, column).number_format = "yyyy-mm-dd"
            elif column > 4:
                sheet.cell(output_row, column).number_format = "0.00"
    sheet.freeze_panes = "A2"
    return len(harvest_days)


def environment_dli_by_cycle(workbook) -> dict[int, float]:
    if "Output5 - Environment" not in workbook.sheetnames:
        return {}
    sheet = workbook["Output5 - Environment"]
    headers = {key_name(cell.value): cell.column for cell in sheet[1] if cell.value}
    dli_column = next(
        (column for key, column in headers.items() if "daylightintegral" in key),
        None,
    )
    if not dli_column:
        return {}
    result = {}
    for row in range(2, sheet.max_row + 1):
        cycle = number(sheet.cell(row, 1).value)
        dli = number(sheet.cell(row, dli_column).value)
        if cycle is not None and dli is not None:
            result[int(cycle)] = dli
    return result


def light_use_efficiency(record, average_dli) -> float | None:
    total_weight_kg = weight_kg(record)
    total_plot_size = number(record.get("total_plot_size"))
    cycle = number(record.get("cycle"))
    average_dli = number(average_dli)
    if not total_weight_kg or not total_plot_size or not cycle or not average_dli:
        return None
    harvested_grams_per_m2 = total_weight_kg * 1000 / total_plot_size
    cumulative_light_mol_per_m2 = average_dli * cycle
    return round(harvested_grams_per_m2 / cumulative_light_mol_per_m2, 2)


def populate_workbook(
    template: Path,
    output_path: Path,
    evaluation_path: Path,
    environment_path: Path,
    shelf_path: Path | None,
    photos_folder: Path,
    trial_name_override: str | None = None,
) -> dict[str, int]:
    rows, info = evaluation_rows(evaluation_path)
    shelf = shelf_life_days(shelf_path) if shelf_path and shelf_path.is_file() else {}
    harvest_days = info.get("harvestdays", [])
    photos = photo_measurements(photos_folder, harvest_days)
    rendered_photo_folder = output_path.parent / "Variety Overview Photos"
    save_rendered_photo_audits(photos, rendered_photo_folder)

    workbook = load_workbook(template)
    out1 = workbook["Output1 - Excel"]
    out2 = workbook["Output2 - TDM"]
    out3 = workbook["Output3 - Airtable"]
    trial = clean_name(trial_name_override) or info.get("trialname")
    farm = info.get("farm")
    environment_cycles = add_environment_summary_sheet(workbook, environment_path, rows, info)
    dli_by_cycle = environment_dli_by_cycle(workbook)
    out1.cell(1, 27).value = "Light Use Efficiency (g/mol)"
    copy_cell_style(out1.cell(1, 26), out1.cell(1, 27))
    out1.column_dimensions[out1.cell(1, 27).column_letter].width = out1.column_dimensions[out1.cell(1, 26).column_letter].width
    rank_map = {}
    for day in harvest_days:
        day_rows = [record for record in rows if record["harvest_day"] == day]
        ordered = sorted(
            (record for record in day_rows if annual_yield(record) is not None),
            key=annual_yield,
            reverse=True,
        )
        rank_map.update({(day, record["key"]): index + 1 for index, record in enumerate(ordered)})

    for index, record in enumerate(rows, start=2):
        day_photos = photos.get(record["harvest_day"], {})
        size_key = photo_key_for(record["key"], day_photos)
        measurement = day_photos.get(size_key) if size_key else None
        values = [
            trial,
            record["variety"],
            record.get("plot_size"),
            record.get("total_plot_size"),
            record.get("cycle"),
            weight_grams(record, "raft1"),
            weight_grams(record, "raft2"),
            weight_grams(record, "raft3"),
            weight_grams(record, "extra"),
            weight_grams(record),
            record.get("germ"),
            record.get("uniformity"),
            record.get("tipburn"),
            record.get("yellowing"),
            record.get("root"),
            record.get("bolting"),
            record.get("glassiness"),
            record.get("crispiness"),
            record.get("overall"),
            rank_map.get((record["harvest_day"], record["key"])),
            record.get("taste"),
            record.get("notes"),
            annual_yield(record),
            measurement.get("length") if measurement else None,
            None,
            shelf.get(record["key"]),
            light_use_efficiency(record, dli_by_cycle.get(int(record["cycle"]))),
        ]
        for column, value in enumerate(values, start=1):
            out1.cell(index, column).value = value
            copy_cell_style(out1.cell(1, column), out1.cell(index, column))
        out1.cell(index, 23).number_format = "0.00"
        out1.cell(index, 27).number_format = "0.00"

    add_tdm_sheet(out2, rows, harvest_days, photos, shelf)

    latest_day = max(harvest_days) if harvest_days else None
    latest_rows = [record for record in rows if record["harvest_day"] == latest_day]
    points = comparative_points(latest_rows)
    for index, record in enumerate(latest_rows, start=3):
        pros, cons = points.get(record["key"], (None, None))
        file_key = photo_key_for(record["key"], photos.get(latest_day, {}))
        values = [
            farm,
            None,
            record["variety"],
            info.get("lastharvestdate"),
            pros,
            cons,
            record.get("notes"),
            annual_yield(record),
            None,
            "Yes" if file_key else None,
        ]
        for column, value in enumerate(values, start=1):
            out3.cell(index, column).value = value
            copy_cell_style(out3.cell(2, column), out3.cell(index, column))
        out3.cell(index, 4).number_format = "yyyy-mm-dd"
        out3.cell(index, 8).number_format = "0.00"

    photo_count = add_photo_audit_sheet(workbook, rows, photos)
    out1.freeze_panes = "A2"
    out2.freeze_panes = "A3"
    out3.freeze_panes = "A3"
    for sheet in (out1, out2, out3):
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell.alignment = copy(cell.alignment)
                    cell.alignment = Alignment(
                        horizontal=cell.alignment.horizontal,
                        vertical="top",
                        text_rotation=cell.alignment.text_rotation,
                        wrap_text=True,
                        shrink_to_fit=cell.alignment.shrink_to_fit,
                        indent=cell.alignment.indent,
                    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.save(output_path)
    return {
        "varieties": len(latest_rows),
        "harvests": len(harvest_days),
        "records": len(rows),
        "leaf_sizes": sum(
            1
            for day in photos.values()
            for measurement in day.values()
            if measurement.get("status") == "OK"
        ),
        "photo_reviews": sum(
            1
            for day in photos.values()
            for measurement in day.values()
            if measurement.get("status") != "OK"
        ),
        "shelf_lives": len(shelf),
        "photo_audits": photo_count,
        "environment_cycles": environment_cycles,
    }


@dataclass
class Inputs:
    evaluation: Path
    environment: Path
    shelf_life: Path | None
    photos: Path
    output_folder: Path
    trial_name: str = ""


def run_analysis(inputs: Inputs, template: Path | None = None) -> tuple[Path, dict[str, int]]:
    for path in (inputs.evaluation, inputs.environment):
        if not path.is_file():
            raise FileNotFoundError(f"Input workbook not found: {path}")
    if inputs.shelf_life is not None and not inputs.shelf_life.is_file():
        raise FileNotFoundError(f"Shelf-life workbook not found: {inputs.shelf_life}")
    if not inputs.photos.is_dir():
        raise FileNotFoundError(f"Photo folder not found: {inputs.photos}")

    # Environment data is validated even though the supplied output formats have no
    # environment fields. Unsupported values are intentionally not invented.
    environment_book = load_workbook(inputs.environment, read_only=True, data_only=True)
    if not environment_book.sheetnames:
        raise ValueError("The environment workbook contains no worksheets.")
    environment_book.close()

    app_folder = Path(__file__).resolve().parent
    template_path = template or app_folder / "Example Output.xlsx"
    if not template_path.is_file():
        raise FileNotFoundError(
            "Example Output.xlsx is missing from the application folder. "
            "Copy the provided example output workbook beside biophi_app.py."
        )
    output_name = f"Biophi Analysis {datetime.now():%Y-%m-%d_%H%M%S}.xlsx"
    output_path = inputs.output_folder / output_name
    return output_path, populate_workbook(
        template_path,
        output_path,
        inputs.evaluation,
        inputs.environment,
        inputs.shelf_life,
        inputs.photos,
        inputs.trial_name,
    )


def write_error_log(output_folder: Path, error: BaseException) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    log = output_folder / "Biophi Analysis Error.txt"
    log.write_text(
        f"{type(error).__name__}: {error}\n\n{traceback.format_exc()}",
        encoding="utf-8",
    )
    return log
