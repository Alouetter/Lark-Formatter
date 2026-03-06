"""Utilities for line-spacing semantics across rules."""

from __future__ import annotations

from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt


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

