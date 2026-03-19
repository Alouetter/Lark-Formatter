"""行距类型选项与显示辅助。"""

from __future__ import annotations

LINE_SPACING_OPTIONS: tuple[tuple[str, str], ...] = (
    ("exact", "固定值"),
    ("single", "单倍"),
    ("one_half", "1.5 倍"),
    ("double", "双倍"),
    ("multiple", "多倍"),
)

LINE_SPACING_VALUE_TO_LABEL = {
    value: label for value, label in LINE_SPACING_OPTIONS
}

_LINE_SPACING_ALIASES = {
    "exact": "exact",
    "固定值": "exact",
    "single": "single",
    "单倍": "single",
    "one_half": "one_half",
    "1.5倍": "one_half",
    "1.5 倍": "one_half",
    "double": "double",
    "双倍": "double",
    "multiple": "multiple",
    "多倍": "multiple",
}

_FIXED_LINE_SPACING_VALUES = {
    "single": 1.0,
    "one_half": 1.5,
    "double": 2.0,
}

_DEFAULT_LINE_SPACING_VALUES = {
    "exact": 20.0,
    "single": 1.0,
    "one_half": 1.5,
    "double": 2.0,
    "multiple": 1.0,
}


def normalize_line_spacing_type(value: str) -> str:
    raw = str(value or "").strip()
    if not raw:
        return "exact"
    return _LINE_SPACING_ALIASES.get(raw.lower(), _LINE_SPACING_ALIASES.get(raw, "exact"))


def line_spacing_display_label(value: str) -> str:
    return LINE_SPACING_VALUE_TO_LABEL.get(normalize_line_spacing_type(value), "固定值")


def line_spacing_unit_label(value: str) -> str:
    kind = normalize_line_spacing_type(value)
    return "磅" if kind == "exact" else "倍"


def line_spacing_is_editable(value: str) -> bool:
    kind = normalize_line_spacing_type(value)
    return kind in {"exact", "multiple"}


def resolve_line_spacing_value(kind: str, raw_value) -> float:
    normalized_kind = normalize_line_spacing_type(kind)
    if normalized_kind in _FIXED_LINE_SPACING_VALUES:
        return _FIXED_LINE_SPACING_VALUES[normalized_kind]

    try:
        value = float(raw_value)
    except (TypeError, ValueError):
        value = _DEFAULT_LINE_SPACING_VALUES[normalized_kind]

    if value <= 0:
        value = _DEFAULT_LINE_SPACING_VALUES[normalized_kind]
    return value
