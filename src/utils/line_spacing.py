"""Utilities for line-spacing semantics across rules."""

from __future__ import annotations

from lxml import etree
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def normalize_line_spacing(line_spacing_type: str, line_spacing_value) -> tuple[str, float] | None:
    """Normalize config line-spacing to ('exact'|'multiple', value)."""
    kind = str(line_spacing_type or "").strip().lower()
    try:
        value = float(line_spacing_value)
    except (TypeError, ValueError):
        value = 0.0

    if kind == "exact":
        if value <= 0:
            value = 20.0
        return ("exact", value)

    if kind == "single":
        return ("multiple", 1.0)
    if kind == "one_half":
        return ("multiple", 1.5)
    if kind == "double":
        return ("multiple", 2.0)
    if kind == "multiple":
        if value <= 0:
            value = 1.0
        return ("multiple", value)

    return None


def apply_line_spacing(paragraph_format, line_spacing_type: str, line_spacing_value) -> None:
    """Apply normalized line-spacing to a python-docx paragraph format object."""
    resolved = normalize_line_spacing(line_spacing_type, line_spacing_value)
    if resolved is None:
        return
    kind, value = resolved
    if kind == "exact":
        paragraph_format.line_spacing = Pt(value)
        paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        return
    paragraph_format.line_spacing = value
    paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE


def _ensure_spacing(container_element):
    ppr = container_element.find(_w("pPr"))
    if ppr is None:
        ppr = etree.SubElement(container_element, _w("pPr"))
    spacing = ppr.find(_w("spacing"))
    if spacing is None:
        spacing = etree.SubElement(ppr, _w("spacing"))
    return ppr, spacing


def _sync_line_spacing_attrs(spacing, line_spacing_type: str, line_spacing_value) -> bool:
    resolved = normalize_line_spacing(line_spacing_type, line_spacing_value)
    if resolved is None:
        return False
    kind, value = resolved
    if kind == "exact":
        spacing.set(_w("line"), str(int(round(value * 20))))
        spacing.set(_w("lineRule"), "exact")
        return True
    spacing.set(_w("line"), str(int(round(value * 240))))
    spacing.set(_w("lineRule"), "auto")
    return True


def sync_line_spacing_ooxml(
    container_element,
    *,
    line_spacing_type: str = "single",
    line_spacing_value=1.0,
) -> None:
    """Synchronize only OOXML line-spacing attrs while preserving before/after spacing."""
    _, spacing = _ensure_spacing(container_element)
    _sync_line_spacing_attrs(spacing, line_spacing_type, line_spacing_value)


def apply_safe_picture_line_spacing(paragraph) -> None:
    """Force image paragraphs to Word-safe single/auto line spacing to avoid clipping."""
    apply_line_spacing(paragraph.paragraph_format, "single", 1.0)
    sync_line_spacing_ooxml(
        paragraph._element,
        line_spacing_type="single",
        line_spacing_value=1.0,
    )


def sync_spacing_ooxml(
    container_element,
    *,
    space_before_pt=0.0,
    space_after_pt=0.0,
    line_spacing_type: str = "exact",
    line_spacing_value=20.0,
) -> None:
    """Synchronize OOXML spacing attrs and clear legacy auto-spacing leftovers."""
    ppr, spacing = _ensure_spacing(container_element)

    try:
        before_pt = max(0.0, float(space_before_pt or 0.0))
    except (TypeError, ValueError):
        before_pt = 0.0
    try:
        after_pt = max(0.0, float(space_after_pt or 0.0))
    except (TypeError, ValueError):
        after_pt = 0.0

    spacing.set(_w("before"), str(int(round(before_pt * 20))))
    spacing.set(_w("after"), str(int(round(after_pt * 20))))
    for attr_name in ("beforeAutospacing", "afterAutospacing", "beforeLines", "afterLines"):
        spacing.attrib.pop(_w(attr_name), None)

    _sync_line_spacing_attrs(spacing, line_spacing_type, line_spacing_value)

    contextual = ppr.find(_w("contextualSpacing"))
    if contextual is not None:
        ppr.remove(contextual)
