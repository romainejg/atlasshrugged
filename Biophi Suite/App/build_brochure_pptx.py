"""
Server-compatible PowerPoint brochure builder for Biophi Suite.

This module replaces ``build_brochure.ps1``, which required Windows COM
automation (Microsoft PowerPoint installed).  It uses ``python-pptx`` and
runs on Linux, macOS, and Windows without any Office installation.

Install dependency:  pip install python-pptx

Entry points
------------
``build_deck_from_manifest(manifest_path, template_path, output_path)``
    Called by ``brochure_maker.build_brochure()``.  Reads the JSON manifest
    written by that function and renders the deck.

``main()``
    CLI entry point for standalone use.
"""

from __future__ import annotations

import copy
import json
import math
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Optional

# ---------------------------------------------------------------------------
# Named shape identifiers used in the branded template
# (these match the shape names used in build_brochure.ps1)
# ---------------------------------------------------------------------------
_SHAPE_IDENTITY = "TextBox 7"
_SHAPE_DESCRIPTION = "TextBox 17"
_SHAPE_HARVEST_1 = "TextBox 9"
_SHAPE_HARVEST_2 = "TextBox 12"
_SHAPE_HARVEST_3 = "TextBox 15"
_SHAPE_DETAILS = "TextBox 29"
_SHAPE_DENSITY = "TextBox 34"
_SHAPE_HERO = "Picture 36"
_SHAPE_LEAF_1 = "Picture 8"
_SHAPE_LEAF_2 = "Picture 11"
_SHAPE_LEAF_3 = "Picture 14"
_SHAPE_CHART = "Chart 13"

# EMU (English Metric Units, used by python-pptx): 914 400 EMU = 1 inch
_EMU_PER_INCH = 914_400
_EMU_PER_CM = 360_000  # 914 400 / 2.54
# The template uses PowerPoint "points" for the leaf baseline value.
# 1 point = 12 700 EMU.
_EMU_PER_PT = 12_700
_LEAF_BASELINE_PT = 396.3  # same value as in build_brochure.ps1


def build_deck_from_manifest(
    manifest_path: str,
    template_path: str,
    output_path: str,
) -> str:
    """
    Build a branded brochure deck from a ``brochure_manifest.json`` file.

    Args:
        manifest_path:  Path to the JSON manifest produced by
                        ``brochure_maker.build_brochure()``.
        template_path:  Path to the branded ``.pptx`` template.
        output_path:    Destination path for the generated ``.pptx``.

    Returns:
        The resolved output path as a string.
    """
    from pptx import Presentation

    manifest = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    rows: list[dict[str, Any]] = manifest.get("rows", [])
    if not rows:
        raise ValueError("No varieties found in the brochure manifest.")

    prs = Presentation(template_path)
    _build(prs, rows, output_path)
    return output_path


def _build(prs: Any, rows: list[dict[str, Any]], output_path: str) -> None:
    """Core deck-building logic (separated for testability)."""
    # --- Calculate deck-wide physical leaf scale (same logic as PS1) ---
    max_leaf_h = 0.0
    max_leaf_w = 0.0
    for row in rows:
        trial = row.get("trial_data") or {}
        for leaf in trial.get("leaves") or []:
            h = _to_float(leaf.get("height_cm"))
            if h is None:
                continue
            max_leaf_h = max(max_leaf_h, h)
            w = _to_float(leaf.get("width_cm"))
            if w is not None:
                max_leaf_w = max(max_leaf_w, w)
            elif leaf.get("pixel_width") and leaf.get("pixel_height"):
                max_leaf_w = max(max_leaf_w, h * leaf["pixel_width"] / leaf["pixel_height"])

    # EMU per cm, capped so the tallest/widest leaf fits in the available space
    leaf_emu_per_cm = float(_EMU_PER_CM)
    if max_leaf_h > 0:
        leaf_emu_per_cm = min(leaf_emu_per_cm, _EMU_PER_PT * 180.0 / max_leaf_h)
    if max_leaf_w > 0:
        leaf_emu_per_cm = min(leaf_emu_per_cm, _EMU_PER_PT * 100.0 / max_leaf_w)

    leaf_baseline_emu = int(_LEAF_BASELINE_PT * _EMU_PER_PT)
    ruler_max_cm = max(5, int(math.ceil(max_leaf_h / 5.0) * 5)) if max_leaf_h > 0 else 5

    # --- Duplicate the template slide for each variety ---
    # Variety 1 reuses the original template slide; 2…N get fresh copies.
    slides = [prs.slides[0]]
    for _ in range(1, len(rows)):
        slides.append(_duplicate_slide(prs, 0))

    # --- Fill each slide ---
    for slide, row in zip(slides, rows):
        _fill_slide(slide, row, leaf_emu_per_cm, leaf_baseline_emu, ruler_max_cm)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)


# ---------------------------------------------------------------------------
# Slide duplication
# ---------------------------------------------------------------------------

def _rId_num(rId: str) -> int:
    """Return the numeric suffix of an rId string for stable sort ordering."""
    return int(rId[3:]) if rId.startswith("rId") and rId[3:].isdigit() else 0


def _duplicate_slide(prs: Any, source_index: int = 0) -> Any:
    """
    Append an independent copy of ``prs.slides[source_index]`` and return it.

    Chart parts (embedded workbooks) are deep-copied so each slide's chart
    can be updated independently.
    """
    from pptx.parts.slide import SlidePart
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    src_slide = prs.slides[source_index]
    src_part: SlidePart = src_slide.part

    # --- Next available slide partname (same logic as PresentationPart._next_slide_partname) ---
    prs_part = prs.part
    partname = prs_part._next_slide_partname

    # --- Create new SlidePart with a deep copy of the source XML element ---
    new_elem = copy.deepcopy(src_part._element)
    new_part = SlidePart(partname, src_part.content_type, src_part._package, new_elem)

    # --- Re-wire relationships in rId order so embedded XML references remain valid ---
    sorted_rels = sorted(src_part.rels.items(), key=lambda kv: _rId_num(kv[0]))
    for _rId, rel in sorted_rels:
        if rel.is_external:
            new_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        elif RT.CHART in rel.reltype:
            # Deep-copy chart (and its embedded xlsx) so each slide's chart is independent
            cloned_chart = _clone_chart_part(rel.target_part, len(prs.slides))
            new_part.relate_to(cloned_chart, rel.reltype)
        else:
            # Layout, images, and all other parts are shared (read-only in this context)
            new_part.relate_to(rel.target_part, rel.reltype)

    # --- Register the new slide with the presentation ---
    rId = prs_part.relate_to(new_part, RT.SLIDE)
    prs.slides._sldIdLst.add_sldId(rId)

    return new_part.slide


def _clone_chart_part(chart_part: Any, slide_idx: int) -> Any:
    """Return a deep, independent copy of *chart_part* with a unique partname."""
    from pptx.parts.chart import ChartPart
    from pptx.opc.constants import CONTENT_TYPE as CT

    package = chart_part._package
    new_partname = package.next_partname(ChartPart.partname_template)
    new_elem = copy.deepcopy(chart_part._element)
    new_chart = ChartPart(new_partname, CT.DML_CHART, package, new_elem)

    # Copy nested relationships (style, colours, embedded xlsx workbook)
    sorted_rels = sorted(chart_part.rels.items(), key=lambda kv: _rId_num(kv[0]))
    for _rId, rel in sorted_rels:
        if rel.is_external:
            new_chart.relate_to(rel.target_ref, rel.reltype, is_external=True)
        elif "package" in rel.reltype.lower() or "oleObject" in rel.reltype:
            # The embedded xlsx workbook — copy so its data can be changed independently
            wb_clone = _clone_binary_part(rel.target_part, slide_idx)
            new_chart.relate_to(wb_clone, rel.reltype)
        else:
            # Style/colour parts are shared (read-only)
            new_chart.relate_to(rel.target_part, rel.reltype)

    return new_chart


def _clone_binary_part(part: Any, slide_idx: int) -> Any:
    """Return an independent copy of a binary part (e.g. embedded xlsx) with a unique name."""
    from pptx.opc.package import Part
    from pptx.opc.packuri import PackURI

    old_name = str(part.partname)
    dot_pos = old_name.rfind(".")
    base, ext = old_name[:dot_pos], old_name[dot_pos:]
    new_name = PackURI(f"{base}_s{slide_idx}{ext}")

    blob = part.blob
    return Part(new_name, part.content_type, part._package, blob)


# ---------------------------------------------------------------------------
# Slide filling
# ---------------------------------------------------------------------------

def _fill_slide(
    slide: Any,
    row: dict[str, Any],
    leaf_emu_per_cm: float,
    leaf_baseline_emu: int,
    ruler_max_cm: int,
) -> None:
    """Apply all variety-specific content to *slide*."""
    trial = row.get("trial_data") or {}
    assets = row.get("local_assets") or {}

    name = row.get("variety") or row.get("name") or row.get("id") or ""
    segment = row.get("segment") or ""
    hr = row.get("hr") or ""
    ir = row.get("ir") or ""
    description = row.get("description") or ""
    hero_path = row.get("hero_path") or ""
    env_data = trial.get("environment") or {}
    strengths = trial.get("strengths") or ""
    weaknesses = trial.get("weaknesses") or ""

    points = sorted(trial.get("yield") or [], key=lambda p: p.get("dtm", 0))[:3]
    leaves = sorted(trial.get("leaves") or [], key=lambda l: l.get("dtm", 0))[:3]

    def leaf_for(point: Optional[dict]) -> Optional[dict]:
        if not point:
            return None
        return next((lf for lf in leaves if lf.get("dtm") == point.get("dtm")), None)

    point1, point2, point3 = (points + [None, None, None])[:3]
    leaf1, leaf2, leaf3 = leaf_for(point1), leaf_for(point2), leaf_for(point3)

    shapes_by_name = {s.name: s for s in slide.shapes}

    # --- Identity box: name, segment, HR, IR ---
    id_shape = shapes_by_name.get(_SHAPE_IDENTITY)
    if id_shape and id_shape.has_text_frame:
        _replace_text(id_shape, "C127", name)
        _replace_text(id_shape, "Segment: Crispy", f"Segment: {segment}" if segment else "Segment: ")
        _replace_text(id_shape, "Nr:0", str(hr))
        _replace_text(id_shape, "LMV:1", str(ir))

    # --- Description ---
    desc_shape = shapes_by_name.get(_SHAPE_DESCRIPTION)
    if desc_shape and desc_shape.has_text_frame and description:
        _set_full_text(desc_shape, description)

    # --- Harvest labels ---
    _set_harvest_label(shapes_by_name.get(_SHAPE_HARVEST_1), "19", "(10cm)", "10 g/mol", point1, leaf1)
    _set_harvest_label(shapes_by_name.get(_SHAPE_HARVEST_2), "21", "(13 cm)", "10 g/mol", point2, leaf2)
    _set_harvest_label(shapes_by_name.get(_SHAPE_HARVEST_3), "21", "(14 cm)", "10/mol", point3, leaf3)

    # --- Environmental details ---
    detail_shape = shapes_by_name.get(_SHAPE_DETAILS)
    if detail_shape and detail_shape.has_text_frame:
        _replace_text(detail_shape, "23.16", _fmt_num(env_data.get("dli"), "0.2f"))
        _replace_text(detail_shape, "20.65", _fmt_num(env_data.get("temperature"), "0.2f"))
        rh_val = env_data.get("rh")
        rh_text = (f"RH: {_fmt_num(rh_val, '0.2f')}%") if rh_val is not None else "RH:"
        _replace_text(detail_shape, "RH:", rh_text)
        _replace_text(detail_shape, "May 25, 2026", _fmt_date(env_data.get("harvest_date")))
        _replace_text(detail_shape, "higher yield; strong overall value", strengths)
        _replace_text(detail_shape, "weak uniformity", weaknesses)

    # --- Clear density placeholder ---
    density_shape = shapes_by_name.get(_SHAPE_DENSITY)
    if density_shape and density_shape.has_text_frame:
        _replace_text(density_shape, "560 plants/m2", "")

    # --- Hero image ---
    _replace_picture(slide, _SHAPE_HERO, hero_path)

    # --- Scaled leaf images ---
    _place_scaled_leaf(slide, _SHAPE_LEAF_1, leaf1, leaf_emu_per_cm, leaf_baseline_emu)
    _place_scaled_leaf(slide, _SHAPE_LEAF_2, leaf2, leaf_emu_per_cm, leaf_baseline_emu)
    _place_scaled_leaf(slide, _SHAPE_LEAF_3, leaf3, leaf_emu_per_cm, leaf_baseline_emu)

    # --- Ruler ---
    ruler_x_emu = int(58 * _EMU_PER_PT)
    _add_ruler(slide, ruler_x_emu, leaf_baseline_emu, leaf_emu_per_cm, ruler_max_cm)

    # --- Yield chart ---
    chart_shape = shapes_by_name.get(_SHAPE_CHART)
    if chart_shape and hasattr(chart_shape, "chart") and points:
        _update_chart(chart_shape, points)


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def _replace_text(shape: Any, old_text: str, new_text: str) -> None:
    """
    Replace *old_text* with *new_text* in all text runs of *shape*.

    Works across runs within each paragraph by rebuilding runs when the
    placeholder text is split across run boundaries.
    """
    if not old_text or not shape.has_text_frame:
        return
    replacement = new_text or ""
    for para in shape.text_frame.paragraphs:
        # Fast path: placeholder text is entirely within a single run
        for run in para.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, replacement)
                return
        # Slow path: placeholder spans multiple runs — consolidate and replace
        full = "".join(r.text for r in para.runs)
        if old_text in full:
            replaced = full.replace(old_text, replacement)
            if para.runs:
                para.runs[0].text = replaced
                for run in para.runs[1:]:
                    run.text = ""
            return


def _set_full_text(shape: Any, text: str) -> None:
    """Replace the entire text of *shape*'s first paragraph with *text*."""
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    if tf.paragraphs and tf.paragraphs[0].runs:
        tf.paragraphs[0].runs[0].text = text
        for run in tf.paragraphs[0].runs[1:]:
            run.text = ""
    elif tf.paragraphs:
        tf.paragraphs[0].text = text


def _set_harvest_label(
    shape: Optional[Any],
    sample_days: str,
    sample_size: str,
    sample_efficiency: str,
    point: Optional[dict],
    leaf: Optional[dict],
) -> None:
    if shape is None or not shape.has_text_frame:
        return
    days_text = (f"{point['dtm']} Days") if point and point.get("dtm") is not None else ""
    h = (leaf or {}).get("height_cm") or (point or {}).get("height_cm")
    size_text = (f"({_fmt_num(h, '0.#')} cm)") if h is not None else ""
    gmol = (point or {}).get("gmol")
    eff_text = (f"{_fmt_num(gmol, '0.2f')} g/mol") if gmol is not None else ""

    _replace_text(shape, sample_days + " Days", days_text)
    _replace_text(shape, sample_size, size_text)
    _replace_text(shape, sample_efficiency, eff_text)


def _fmt_num(value: Any, fmt: str = ".2f") -> str:
    """Format a numeric *value* or return '' if None/invalid.

    *fmt* follows Python format-spec mini-language (e.g. ``'.2f'`` or ``'.0f'``).
    The special value ``'0.#'`` means up to one decimal place with the trailing
    zero suppressed (matches PowerPoint's ``'0.#'`` number format).
    """
    try:
        f = float(value)
        if fmt in ("0.#", ".#"):
            s = f"{f:.1f}"
            return s.rstrip("0").rstrip(".")
        # Accept PS1-style "0.2f" → normalise to ".2f" for Python's format()
        py_fmt = fmt.lstrip("0") or ".2f"
        if not py_fmt.startswith("."):
            py_fmt = "." + py_fmt
        return format(f, py_fmt)
    except (TypeError, ValueError):
        return ""


def _fmt_date(value: Any) -> str:
    """Parse *value* as a date string and return a human-readable form."""
    if not value:
        return ""
    if isinstance(value, str):
        try:
            return datetime.strptime(value[:10], "%Y-%m-%d").strftime("%B %-d, %Y")
        except (ValueError, AttributeError):
            return str(value)
    return str(value)


# ---------------------------------------------------------------------------
# Image helpers
# ---------------------------------------------------------------------------

def _replace_picture(slide: Any, shape_name: str, image_path: str) -> None:
    """
    Find the placeholder shape named *shape_name*, record its geometry,
    remove it, and insert the image file at *image_path* in its place.
    If *image_path* is empty or missing, the placeholder is simply removed.
    """
    target = _find_shape(slide, shape_name)
    if target is None:
        return

    left, top, width, height = target.left, target.top, target.width, target.height
    rotation = getattr(target, "rotation", 0)
    z_idx = _shape_z_index(slide, target)

    # Remove the placeholder
    sp_tree = target._element.getparent()
    sp_tree.remove(target._element)

    if not image_path or not Path(image_path).is_file():
        return

    pic = slide.shapes.add_picture(image_path, left, top)
    scale = min(width / pic.width, height / pic.height)
    pic.width = int(pic.width * scale)
    pic.height = int(pic.height * scale)
    pic.left = left + (width - pic.width) // 2
    pic.top = top + (height - pic.height) // 2
    if rotation:
        pic.rotation = rotation

    # Restore z-order
    if z_idx is not None:
        _set_z_index(slide, pic, z_idx)


def _place_scaled_leaf(
    slide: Any,
    shape_name: str,
    leaf: Optional[dict],
    emu_per_cm: float,
    baseline_emu: int,
) -> None:
    """
    Replace the placeholder named *shape_name* with a leaf image scaled to
    its physical height, bottom-aligned to *baseline_emu*.
    """
    target = _find_shape(slide, shape_name)
    if target is None:
        return

    center_emu = target.left + target.width // 2
    z_idx = _shape_z_index(slide, target)

    sp_tree = target._element.getparent()
    sp_tree.remove(target._element)

    if (
        not leaf
        or not leaf.get("path")
        or not Path(leaf["path"]).is_file()
        or leaf.get("height_cm") is None
    ):
        return

    leaf_h_emu = int(float(leaf["height_cm"]) * emu_per_cm)
    pic = slide.shapes.add_picture(leaf["path"], 0, 0)
    pic.height = leaf_h_emu
    pic.left = center_emu - pic.width // 2
    pic.top = baseline_emu - leaf_h_emu

    if z_idx is not None:
        _set_z_index(slide, pic, z_idx)


def _add_ruler(
    slide: Any,
    x_emu: int,
    baseline_emu: int,
    emu_per_cm: float,
    max_cm: int,
) -> None:
    """
    Add a scale ruler (vertical line + tick marks + cm labels) to *slide*,
    matching the ruler drawn by build_brochure.ps1.
    """
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

    black = RGBColor(0, 0, 0)
    top_emu = baseline_emu - int(max_cm * emu_per_cm)

    # Vertical line
    vline = slide.shapes.add_connector(
        1,  # MSO_CONNECTOR.STRAIGHT
        Emu(x_emu), Emu(baseline_emu),
        Emu(x_emu), Emu(top_emu),
    )
    vline.line.color.rgb = black
    vline.line.width = Pt(1)

    for cm in range(max_cm + 1):
        y_emu = baseline_emu - int(cm * emu_per_cm)
        major = cm % 5 == 0
        tick_len_emu = int(6 * _EMU_PER_PT) if major else int(3 * _EMU_PER_PT)
        weight = Pt(1.0) if major else Pt(0.5)

        tick = slide.shapes.add_connector(
            1,
            Emu(x_emu), Emu(y_emu),
            Emu(x_emu + tick_len_emu), Emu(y_emu),
        )
        tick.line.color.rgb = black
        tick.line.width = weight

        if major:
            lbl_w = int(20 * _EMU_PER_PT)
            lbl_h = int(12 * _EMU_PER_PT)
            lbl_x = x_emu - int(24 * _EMU_PER_PT)
            lbl_y = y_emu - int(6 * _EMU_PER_PT)
            txbox = slide.shapes.add_textbox(Emu(lbl_x), Emu(lbl_y), Emu(lbl_w), Emu(lbl_h))
            tf = txbox.text_frame
            tf.word_wrap = False
            tf.text = str(cm)
            para = tf.paragraphs[0]
            para.alignment = PP_ALIGN.RIGHT
            if para.runs:
                run = para.runs[0]
                run.font.name = "Rubik"
                run.font.size = Pt(7)

    # "cm" unit label above the ruler
    unit_w = int(25 * _EMU_PER_PT)
    unit_h = int(12 * _EMU_PER_PT)
    unit_x = x_emu - int(7 * _EMU_PER_PT)
    unit_y = top_emu - int(14 * _EMU_PER_PT)
    unit_box = slide.shapes.add_textbox(Emu(unit_x), Emu(unit_y), Emu(unit_w), Emu(unit_h))
    unit_box.text_frame.text = "cm"
    if unit_box.text_frame.paragraphs[0].runs:
        r = unit_box.text_frame.paragraphs[0].runs[0]
        r.font.name = "Rubik"
        r.font.size = Pt(7)


# ---------------------------------------------------------------------------
# Chart helper
# ---------------------------------------------------------------------------

def _update_chart(chart_shape: Any, points: list[dict]) -> None:
    """Update the yield chart with harvest-point data from *points*."""
    try:
        from pptx.chart.data import ChartData

        chart_data = ChartData()
        chart_data.categories = [str(p.get("dtm", "")) for p in points]
        values = [p.get("value") if isinstance(p.get("value"), (int, float)) else None for p in points]
        chart_data.add_series("Yield (kg/m²/yr)", values)
        chart_shape.chart.replace_data(chart_data)
    except Exception:
        pass  # Chart update is best-effort; don't fail the whole build


# ---------------------------------------------------------------------------
# Shape utility helpers
# ---------------------------------------------------------------------------

def _find_shape(slide: Any, name: str) -> Optional[Any]:
    """Return the first shape named *name* in *slide*, or None."""
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


def _shape_z_index(slide: Any, shape: Any) -> Optional[int]:
    """Return the z-order index (0 = bottom) of *shape* in *slide*."""
    for i, s in enumerate(slide.shapes):
        if s._element is shape._element:
            return i
    return None


def _set_z_index(slide: Any, shape: Any, target_idx: int) -> None:
    """Move *shape* to z-order position *target_idx* (0 = bottom)."""
    sp_tree = slide.shapes._spTree
    elems = list(sp_tree)
    try:
        current_idx = elems.index(shape._element)
    except ValueError:
        return
    # The first 2 elements of spTree are metadata (spPr, grpSpPr); shapes start at 2
    desired = max(2, target_idx + 2)
    if current_idx == desired:
        return
    sp_tree.remove(shape._element)
    elems = list(sp_tree)
    if desired >= len(elems):
        sp_tree.append(shape._element)
    else:
        sp_tree.insert(desired, shape._element)


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main() -> int:
    import argparse, sys

    parser = argparse.ArgumentParser(
        description="Build a Biophi brochure .pptx from a manifest JSON file."
    )
    parser.add_argument("manifest", help="Path to brochure_manifest.json")
    parser.add_argument("template", help="Path to template.pptx")
    parser.add_argument("output", help="Output .pptx path")
    args = parser.parse_args()

    out = build_deck_from_manifest(args.manifest, args.template, args.output)
    print(f"Deck written to: {out}")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())
