"""Utilities for applying unit-aware paragraph indents."""

from __future__ import annotations

from lxml import etree
from docx.shared import Pt

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_CM_TO_PT = 72.0 / 2.54
_SPECIAL_INDENT_NONE = "none"
_SPECIAL_INDENT_FIRST_LINE = "first_line"
_SPECIAL_INDENT_HANGING = "hanging"


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def normalize_indent_unit(unit) -> str:
    raw = str(unit or "").strip().lower()
    if raw in {"pt", "point", "points"}:
        return "pt"
    if raw in {"cm", "centimeter", "centimeters"}:
        return "cm"
    return "chars"


def normalize_special_indent_mode(mode) -> str:
    raw = str(mode or "").strip().lower()
    if raw in {"first", "firstline", "first_line", "first-line"}:
        return _SPECIAL_INDENT_FIRST_LINE
    if raw in {"hanging", "hang"}:
        return _SPECIAL_INDENT_HANGING
    return _SPECIAL_INDENT_NONE


def normalize_indent_value(value) -> float:
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        numeric = 0.0
    return max(0.0, numeric)


def normalize_indent_size_pt(size_pt) -> float:
    try:
        numeric = float(size_pt)
    except (TypeError, ValueError):
        numeric = 12.0
    return numeric if numeric > 0 else 12.0


def config_indent_value_to_pt(value, size_pt, unit: str = "chars") -> float:
    """Convert stored config indent value to its pt equivalent."""
    normalized_value = normalize_indent_value(value)
    normalized_unit = normalize_indent_unit(unit)
    if normalized_unit == "pt":
        return normalized_value
    if normalized_unit == "cm":
        return normalized_value * _CM_TO_PT
    return normalized_value * normalize_indent_size_pt(size_pt)


def resolve_pt_indent_value(value_pt, unit: str, size_pt: float) -> float:
    if normalize_indent_unit(unit) == "cm":
        return value_pt / _CM_TO_PT if _CM_TO_PT > 0 else 0.0
    if normalize_indent_unit(unit) == "pt":
        return value_pt
    normalized_size = normalize_indent_size_pt(size_pt)
    return value_pt / normalized_size if normalized_size > 0 else value_pt


def resolve_style_config_special_indent(style_config) -> dict[str, float | str]:
    """Resolve special-indent mode/value/unit, with backward compatibility."""
    raw_mode = normalize_special_indent_mode(getattr(style_config, "special_indent_mode", "none"))
    raw_value = normalize_indent_value(getattr(style_config, "special_indent_value", 0.0))
    raw_unit = normalize_indent_unit(getattr(style_config, "special_indent_unit", "chars"))

    legacy_first_value = normalize_indent_value(getattr(style_config, "first_line_indent_chars", 0.0))
    legacy_first_unit = normalize_indent_unit(getattr(style_config, "first_line_indent_unit", "chars"))
    legacy_hanging_value = normalize_indent_value(getattr(style_config, "hanging_indent_chars", 0.0))
    legacy_hanging_unit = normalize_indent_unit(getattr(style_config, "hanging_indent_unit", "chars"))

    # Legacy fallback: in older configs first-line and hanging were stored as
    # two independent fields. When both are present, prefer hanging because it
    # is the stricter/less ambiguous effective layout. We intentionally let
    # legacy fields win whenever they are non-zero, because a lot of existing
    # code still mutates them directly after StyleConfig construction.
    if legacy_hanging_value > 0:
        return {"mode": _SPECIAL_INDENT_HANGING, "value": legacy_hanging_value, "unit": legacy_hanging_unit}
    if legacy_first_value > 0:
        return {"mode": _SPECIAL_INDENT_FIRST_LINE, "value": legacy_first_value, "unit": legacy_first_unit}
    if raw_mode != _SPECIAL_INDENT_NONE and raw_value > 0:
        return {"mode": raw_mode, "value": raw_value, "unit": raw_unit}
    return {"mode": _SPECIAL_INDENT_NONE, "value": 0.0, "unit": raw_unit}


def sync_style_config_indent_fields(style_config) -> dict[str, float | str]:
    """Keep new special-indent fields and legacy first/hanging fields consistent."""
    size_pt = normalize_indent_size_pt(getattr(style_config, "size_pt", 12.0))
    left_value = normalize_indent_value(getattr(style_config, "left_indent_chars", 0.0))
    left_unit = normalize_indent_unit(getattr(style_config, "left_indent_unit", "chars"))
    right_value = normalize_indent_value(getattr(style_config, "right_indent_chars", 0.0))
    right_unit = normalize_indent_unit(getattr(style_config, "right_indent_unit", "chars"))
    setattr(style_config, "left_indent_chars", left_value)
    setattr(style_config, "left_indent_unit", left_unit)
    setattr(style_config, "right_indent_chars", right_value)
    setattr(style_config, "right_indent_unit", right_unit)

    special = resolve_style_config_special_indent(style_config)
    mode = normalize_special_indent_mode(special["mode"])
    value = normalize_indent_value(special["value"])
    unit = normalize_indent_unit(special["unit"])

    setattr(style_config, "special_indent_mode", mode)
    setattr(style_config, "special_indent_value", value)
    setattr(style_config, "special_indent_unit", unit)

    if mode == _SPECIAL_INDENT_FIRST_LINE and value > 0:
        setattr(style_config, "first_line_indent_chars", value)
        setattr(style_config, "first_line_indent_unit", unit)
        setattr(style_config, "hanging_indent_chars", 0.0)
        setattr(style_config, "hanging_indent_unit", "chars")
    elif mode == _SPECIAL_INDENT_HANGING and value > 0:
        setattr(style_config, "first_line_indent_chars", 0.0)
        setattr(style_config, "first_line_indent_unit", "chars")
        setattr(style_config, "hanging_indent_chars", value)
        setattr(style_config, "hanging_indent_unit", unit)
    else:
        setattr(style_config, "first_line_indent_chars", 0.0)
        setattr(style_config, "first_line_indent_unit", "chars")
        setattr(style_config, "hanging_indent_chars", 0.0)
        setattr(style_config, "hanging_indent_unit", "chars")

    first_value = normalize_indent_value(getattr(style_config, "first_line_indent_chars", 0.0))
    first_unit = normalize_indent_unit(getattr(style_config, "first_line_indent_unit", "chars"))
    hanging_value = normalize_indent_value(getattr(style_config, "hanging_indent_chars", 0.0))
    hanging_unit = normalize_indent_unit(getattr(style_config, "hanging_indent_unit", "chars"))
    return {
        "size_pt": size_pt,
        "left_value": left_value,
        "left_unit": left_unit,
        "right_value": right_value,
        "right_unit": right_unit,
        "special_mode": mode,
        "special_value": value,
        "special_unit": unit,
        "first_value": first_value,
        "first_unit": first_unit,
        "hanging_value": hanging_value,
        "hanging_unit": hanging_unit,
    }


def style_config_indent_kwargs(style_config) -> dict[str, float | str]:
    synced = sync_style_config_indent_fields(style_config)
    return {
        "left_value": synced["left_value"],
        "left_unit": synced["left_unit"],
        "right_value": synced["right_value"],
        "right_unit": synced["right_unit"],
        "first_line_value": synced["first_value"],
        "first_line_unit": synced["first_unit"],
        "hanging_value": synced["hanging_value"],
        "hanging_unit": synced["hanging_unit"],
        "size_pt": synced["size_pt"],
    }


def resolve_style_config_indents(style_config) -> dict[str, float | str]:
    """Resolve configured indents, including special-indent compatibility rules."""
    size_pt = normalize_indent_size_pt(getattr(style_config, "size_pt", 12.0))

    left_value = normalize_indent_value(getattr(style_config, "left_indent_chars", 0.0))
    left_unit = normalize_indent_unit(getattr(style_config, "left_indent_unit", "chars"))
    left_pt = config_indent_value_to_pt(left_value, size_pt, left_unit)
    right_value = normalize_indent_value(getattr(style_config, "right_indent_chars", 0.0))
    right_unit = normalize_indent_unit(getattr(style_config, "right_indent_unit", "chars"))
    right_pt = config_indent_value_to_pt(right_value, size_pt, right_unit)

    special = resolve_style_config_special_indent(style_config)
    special_mode = normalize_special_indent_mode(special["mode"])
    special_value = normalize_indent_value(special["value"])
    special_unit = normalize_indent_unit(special["unit"])
    special_pt = config_indent_value_to_pt(special_value, size_pt, special_unit)

    first_value = 0.0
    first_unit = "chars"
    if special_mode == _SPECIAL_INDENT_FIRST_LINE and special_value > 0:
        first_value = special_value
        first_unit = special_unit
    first_pt = config_indent_value_to_pt(first_value, size_pt, first_unit)

    hanging_value = 0.0
    hanging_unit = "chars"
    if special_mode == _SPECIAL_INDENT_HANGING and special_value > 0:
        hanging_value = special_value
        hanging_unit = special_unit
    hanging_pt = config_indent_value_to_pt(hanging_value, size_pt, hanging_unit)

    effective_left_value = left_value
    effective_left_unit = left_unit
    effective_left_pt = left_pt
    effective_first_pt = first_pt

    # Word's hanging-indent model is "left indent + hanging amount".  When users
    # only configure hanging indent, auto-promote the left indent so the first
    # line stays at the page margin instead of being outdented into the margin.
    if hanging_pt > 0:
        effective_first_pt = 0.0
        if effective_left_pt < hanging_pt:
            effective_left_value = hanging_value
            effective_left_unit = hanging_unit
            effective_left_pt = hanging_pt

    return {
        "size_pt": size_pt,
        "left_value": left_value,
        "left_unit": left_unit,
        "left_pt": left_pt,
        "right_value": right_value,
        "right_unit": right_unit,
        "right_pt": right_pt,
        "special_mode": special_mode,
        "special_value": special_value,
        "special_unit": special_unit,
        "special_pt": special_pt,
        "first_value": first_value,
        "first_unit": first_unit,
        "first_pt": first_pt,
        "hanging_value": hanging_value,
        "hanging_unit": hanging_unit,
        "hanging_pt": hanging_pt,
        "effective_left_value": effective_left_value,
        "effective_left_unit": normalize_indent_unit(effective_left_unit),
        "effective_left_pt": effective_left_pt,
        "effective_first_pt": effective_first_pt,
    }


def _ensure_ind(container_element):
    ppr = container_element.find(_w("pPr"))
    if ppr is None:
        ppr = etree.SubElement(container_element, _w("pPr"))
    ind = ppr.find(_w("ind"))
    if ind is None:
        ind = etree.SubElement(ppr, _w("ind"))
    return ind


def _sync_single_indent(ind, *, twips_attr: str, chars_attr: str, value, unit: str, size_pt) -> None:
    normalized_unit = normalize_indent_unit(unit)
    normalized_value = normalize_indent_value(value)
    twips_key = _w(twips_attr)
    chars_key = _w(chars_attr)

    # Word may continue inheriting char-based indentation from a parent style
    # when only the twips attribute is set to zero on the child style/paragraph.
    # When callers want an explicit zero indent, write both representations as 0
    # so inherited firstLineChars/leftChars are fully overridden.
    if normalized_value <= 0:
        ind.set(twips_key, "0")
        ind.set(chars_key, "0")
        return

    if normalized_unit == "chars" and normalized_value > 0:
        ind.set(chars_key, str(int(round(normalized_value * 100))))
        ind.attrib.pop(twips_key, None)
        return

    pt_value = config_indent_value_to_pt(normalized_value, size_pt, normalized_unit)
    ind.set(twips_key, str(int(round(pt_value * 20))))
    ind.attrib.pop(chars_key, None)


def sync_indent_ooxml(
    container_element,
    *,
    left_value=0.0,
    left_unit: str = "chars",
    first_line_value=0.0,
    first_line_unit: str = "chars",
    hanging_value=0.0,
    hanging_unit: str = "chars",
    right_value=0.0,
    right_unit: str = "chars",
    size_pt=12.0,
    force_zero: bool = False,
) -> None:
    ind = _ensure_ind(container_element)

    if force_zero:
        left_value = 0.0
        first_line_value = 0.0
        hanging_value = 0.0
        right_value = 0.0
        left_unit = "pt"
        first_line_unit = "pt"
        hanging_unit = "pt"
        right_unit = "pt"

    normalized_left_unit = normalize_indent_unit(left_unit)
    normalized_first_unit = normalize_indent_unit(first_line_unit)
    normalized_hanging_unit = normalize_indent_unit(hanging_unit)
    normalized_right_unit = normalize_indent_unit(right_unit)
    normalized_left_value = normalize_indent_value(left_value)
    normalized_first_value = normalize_indent_value(first_line_value)
    normalized_hanging_value = normalize_indent_value(hanging_value)
    normalized_right_value = normalize_indent_value(right_value)

    effective_left_value = normalized_left_value
    effective_left_unit = normalized_left_unit
    if normalized_hanging_value > 0:
        left_pt = config_indent_value_to_pt(normalized_left_value, size_pt, normalized_left_unit)
        hanging_pt = config_indent_value_to_pt(normalized_hanging_value, size_pt, normalized_hanging_unit)
        if left_pt < hanging_pt:
            effective_left_value = normalized_hanging_value
            effective_left_unit = normalized_hanging_unit

    _sync_single_indent(
        ind,
        twips_attr="left",
        chars_attr="leftChars",
        value=effective_left_value,
        unit=effective_left_unit,
        size_pt=size_pt,
    )

    if normalized_hanging_value > 0:
        _sync_single_indent(
            ind,
            twips_attr="hanging",
            chars_attr="hangingChars",
            value=normalized_hanging_value,
            unit=normalized_hanging_unit,
            size_pt=size_pt,
        )
        ind.attrib.pop(_w("firstLine"), None)
        ind.attrib.pop(_w("firstLineChars"), None)
    else:
        _sync_single_indent(
            ind,
            twips_attr="firstLine",
            chars_attr="firstLineChars",
            value=normalized_first_value,
            unit=normalized_first_unit,
            size_pt=size_pt,
        )
        ind.attrib.pop(_w("hanging"), None)
        ind.attrib.pop(_w("hangingChars"), None)

    _sync_single_indent(
        ind,
        twips_attr="right",
        chars_attr="rightChars",
        value=normalized_right_value,
        unit=normalized_right_unit,
        size_pt=size_pt,
    )


def apply_style_config_indents(paragraph_format, container_element, style_config) -> tuple[float, float, float, float]:
    resolved = resolve_style_config_indents(style_config)

    left_pt = float(resolved["effective_left_pt"])
    first_pt = float(resolved["effective_first_pt"])
    hanging_pt = float(resolved["hanging_pt"])
    right_pt = float(resolved["right_pt"])

    paragraph_format.left_indent = Pt(left_pt if left_pt > 0 else 0)
    if hanging_pt > 0:
        paragraph_format.first_line_indent = Pt(-hanging_pt)
    else:
        paragraph_format.first_line_indent = Pt(first_pt if first_pt > 0 else 0)
    paragraph_format.right_indent = Pt(right_pt if right_pt > 0 else 0)

    sync_indent_ooxml(
        container_element,
        left_value=resolved["effective_left_value"],
        left_unit=str(resolved["effective_left_unit"]),
        first_line_value=resolved["first_value"],
        first_line_unit=str(resolved["first_unit"]),
        hanging_value=resolved["hanging_value"],
        hanging_unit=str(resolved["hanging_unit"]),
        right_value=resolved["right_value"],
        right_unit=str(resolved["right_unit"]),
        size_pt=resolved["size_pt"],
    )
    return left_pt, first_pt, hanging_pt, right_pt
