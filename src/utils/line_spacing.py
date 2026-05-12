"""Utilities for line-spacing semantics across rules."""

from __future__ import annotations

from lxml import etree
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_CM_TO_PT = 72.0 / 2.54
_MM_TO_PT = 72.0 / 25.4
_IN_TO_PT = 72.0
_SPACING_PHYSICAL_UNITS = {"pt", "in", "cm", "mm"}


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def normalize_paragraph_spacing_unit(unit) -> str:
    """Normalize paragraph before/after spacing units."""
    raw = str(unit or "").strip().lower()
    aliases = {
        "point": "pt",
        "points": "pt",
        "磅": "pt",
        "inch": "in",
        "inches": "in",
        "英寸": "in",
        "centimeter": "cm",
        "centimeters": "cm",
        "厘米": "cm",
        "millimeter": "mm",
        "millimeters": "mm",
        "毫米": "mm",
        "line": "line",
        "lines": "line",
        "行": "line",
        "auto": "auto",
        "automatic": "auto",
        "自动": "auto",
    }
    normalized = aliases.get(raw, raw)
    if normalized in {"pt", "in", "cm", "mm", "line", "auto"}:
        return normalized
    return "pt"


def normalize_paragraph_spacing_value(value, *, default: float = 0.0) -> float:
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        numeric = float(default)
    return max(0.0, numeric)


def paragraph_spacing_value_to_pt(value, unit: str = "pt") -> float:
    """Convert physical paragraph-spacing units to points.

    Line and auto spacing are semantic OpenXML modes and must not be converted
    through this helper.
    """
    normalized = normalize_paragraph_spacing_unit(unit)
    numeric = normalize_paragraph_spacing_value(value)
    if normalized == "in":
        return numeric * _IN_TO_PT
    if normalized == "cm":
        return numeric * _CM_TO_PT
    if normalized == "mm":
        return numeric * _MM_TO_PT
    return numeric


def resolve_paragraph_spacing_side(value, unit: str = "pt", legacy_pt=0.0) -> tuple[float, str]:
    normalized_unit = normalize_paragraph_spacing_unit(unit)
    if value is None:
        if normalized_unit in {"line", "auto"}:
            resolved_value = 0.0
        else:
            resolved_value = normalize_paragraph_spacing_value(legacy_pt)
    else:
        resolved_value = normalize_paragraph_spacing_value(value)
    return resolved_value, normalized_unit


def style_spacing_side(style_config, side: str) -> tuple[float, str]:
    prefix = "space_before" if side == "before" else "space_after"
    legacy_attr = f"{prefix}_pt"
    value = getattr(style_config, f"{prefix}_value", None)
    unit = getattr(style_config, f"{prefix}_unit", "pt")
    return resolve_paragraph_spacing_side(
        value,
        unit,
        getattr(style_config, legacy_attr, 0.0),
    )


def sync_style_config_spacing_fields(style_config) -> None:
    """Keep paragraph-spacing value/unit fields compatible with legacy *_pt."""
    for side in ("before", "after"):
        prefix = f"space_{side}"
        value, unit = style_spacing_side(style_config, side)
        setattr(style_config, f"{prefix}_value", value)
        setattr(style_config, f"{prefix}_unit", unit)
        if unit in _SPACING_PHYSICAL_UNITS:
            setattr(style_config, f"{prefix}_pt", round(paragraph_spacing_value_to_pt(value, unit), 4))
        else:
            setattr(style_config, f"{prefix}_pt", 0.0)


def _clear_spacing_side_attrs(spacing, side: str) -> None:
    names = (side, f"{side}Lines", f"{side}Autospacing")
    for attr_name in names:
        spacing.attrib.pop(_w(attr_name), None)


def _sync_spacing_side(spacing, side: str, value, unit: str = "pt", legacy_pt=0.0) -> None:
    resolved_value, resolved_unit = resolve_paragraph_spacing_side(value, unit, legacy_pt)
    _clear_spacing_side_attrs(spacing, side)
    if resolved_unit == "auto":
        spacing.set(_w(f"{side}Autospacing"), "1")
        return
    if resolved_unit == "line":
        spacing.set(_w(f"{side}Lines"), str(int(round(resolved_value * 100))))
        return
    pt_value = paragraph_spacing_value_to_pt(resolved_value, resolved_unit)
    spacing.set(_w(side), str(int(round(pt_value * 20))))


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


def apply_paragraph_spacing(paragraph_format, *, before_value=0.0, before_unit: str = "pt",
                            after_value=0.0, after_unit: str = "pt",
                            before_legacy_pt=0.0, after_legacy_pt=0.0) -> None:
    """Apply python-docx paragraph spacing where representable.

    python-docx can only express physical before/after spacing. Line/auto
    spacing is written by the OOXML synchronizer.
    """
    before_resolved, before_unit = resolve_paragraph_spacing_side(before_value, before_unit, before_legacy_pt)
    after_resolved, after_unit = resolve_paragraph_spacing_side(after_value, after_unit, after_legacy_pt)
    if before_unit in _SPACING_PHYSICAL_UNITS:
        paragraph_format.space_before = Pt(paragraph_spacing_value_to_pt(before_resolved, before_unit))
    if after_unit in _SPACING_PHYSICAL_UNITS:
        paragraph_format.space_after = Pt(paragraph_spacing_value_to_pt(after_resolved, after_unit))


def apply_style_paragraph_spacing(paragraph_format, style_config) -> None:
    before_value, before_unit = style_spacing_side(style_config, "before")
    after_value, after_unit = style_spacing_side(style_config, "after")
    apply_paragraph_spacing(
        paragraph_format,
        before_value=before_value,
        before_unit=before_unit,
        after_value=after_value,
        after_unit=after_unit,
        before_legacy_pt=getattr(style_config, "space_before_pt", 0.0),
        after_legacy_pt=getattr(style_config, "space_after_pt", 0.0),
    )


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
    space_before_value=None,
    space_before_unit: str = "pt",
    space_after_value=None,
    space_after_unit: str = "pt",
    line_spacing_type: str = "exact",
    line_spacing_value=20.0,
) -> None:
    """Synchronize OOXML spacing attrs and line-spacing attrs."""
    ppr, spacing = _ensure_spacing(container_element)

    _sync_spacing_side(
        spacing,
        "before",
        space_before_value,
        space_before_unit,
        space_before_pt,
    )
    _sync_spacing_side(
        spacing,
        "after",
        space_after_value,
        space_after_unit,
        space_after_pt,
    )

    _sync_line_spacing_attrs(spacing, line_spacing_type, line_spacing_value)

    contextual = ppr.find(_w("contextualSpacing"))
    if contextual is not None:
        ppr.remove(contextual)


def sync_style_spacing_ooxml(
    container_element,
    style_config,
    *,
    line_spacing_type: str | None = None,
    line_spacing_value=None,
) -> None:
    before_value, before_unit = style_spacing_side(style_config, "before")
    after_value, after_unit = style_spacing_side(style_config, "after")
    sync_spacing_ooxml(
        container_element,
        space_before_pt=getattr(style_config, "space_before_pt", 0.0),
        space_after_pt=getattr(style_config, "space_after_pt", 0.0),
        space_before_value=before_value,
        space_before_unit=before_unit,
        space_after_value=after_value,
        space_after_unit=after_unit,
        line_spacing_type=line_spacing_type or getattr(style_config, "line_spacing_type", "exact"),
        line_spacing_value=(
            line_spacing_value
            if line_spacing_value is not None
            else getattr(style_config, "line_spacing_pt", 20.0)
        ),
    )
