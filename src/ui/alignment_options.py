"""对齐选项映射。"""

from __future__ import annotations

ALIGNMENT_OPTIONS: tuple[tuple[str, str], ...] = (
    ("left", "左对齐"),
    ("right", "右对齐"),
    ("center", "居中对齐"),
    ("justify", "两端对齐"),
)

_ALIGNMENT_ALIASES = {
    "left": "left",
    "左对齐": "left",
    "居左": "left",
    "right": "right",
    "右对齐": "right",
    "居右": "right",
    "center": "center",
    "居中": "center",
    "居中对齐": "center",
    "垂直居中": "center",
    "justify": "justify",
    "两端对齐": "justify",
}

ALIGNMENT_VALUE_TO_LABEL = {
    value: label for value, label in ALIGNMENT_OPTIONS
}


def normalize_alignment_value(value: str) -> str:
    raw = str(value or "").strip()
    if not raw:
        return "justify"
    return _ALIGNMENT_ALIASES.get(raw.lower(), _ALIGNMENT_ALIASES.get(raw, raw.lower()))


def alignment_display_label(value: str) -> str:
    normalized = normalize_alignment_value(value)
    return ALIGNMENT_VALUE_TO_LABEL.get(normalized, normalized)
