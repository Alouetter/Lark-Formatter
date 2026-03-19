"""字号预设与解析工具。"""

from __future__ import annotations

import re

WORD_NAMED_FONT_SIZES: tuple[tuple[str, float], ...] = (
    ("初号", 42.0),
    ("小初", 36.0),
    ("一号", 26.0),
    ("小一", 24.0),
    ("二号", 22.0),
    ("小二", 18.0),
    ("三号", 16.0),
    ("小三", 15.0),
    ("四号", 14.0),
    ("小四", 12.0),
    ("五号", 10.5),
    ("小五", 9.0),
    ("六号", 7.5),
    ("小六", 6.5),
    ("七号", 5.5),
    ("八号", 5.0),
)

NUMERIC_FONT_SIZE_OPTIONS: tuple[float, ...] = (
    5.0,
    5.5,
    6.5,
    7.5,
    8.0,
    9.0,
    10.0,
    10.5,
    11.0,
    12.0,
    14.0,
    16.0,
    18.0,
    20.0,
    22.0,
    26.0,
    28.0,
    36.0,
    48.0,
    56.0,
    72.0,
)

WORD_NAMED_FONT_SIZE_TO_PT = {
    name: pt for name, pt in WORD_NAMED_FONT_SIZES
}

_PT_TO_NAMED_SIZE = {
    pt: name for name, pt in WORD_NAMED_FONT_SIZES
}

_FONT_SIZE_PT_DISPLAY_RE = re.compile(
    r"(?i)(?P<number>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\s*(?:pt|pts)\b"
)

_FULLWIDTH_ASCII_TRANSLATION = str.maketrans({
    "０": "0",
    "１": "1",
    "２": "2",
    "３": "3",
    "４": "4",
    "５": "5",
    "６": "6",
    "７": "7",
    "８": "8",
    "９": "9",
    "．": ".",
    "。": ".",
    "－": "-",
    "＋": "+",
    "（": "(",
    "）": ")",
})


def normalize_font_size_token(text: str) -> str:
    """Normalize a font-size token for tolerant parsing."""
    raw = str(text or "").strip().translate(_FULLWIDTH_ASCII_TRANSLATION)
    return "".join(raw.split())


def is_named_font_size_input(text: str) -> bool:
    """Return whether the token represents a Chinese named font size."""
    token = normalize_font_size_token(text)
    if not token:
        return False
    return token.split("(", 1)[0] in WORD_NAMED_FONT_SIZE_TO_PT


def format_font_size_pt(value: float) -> str:
    """Format pt values without unnecessary trailing zeros."""
    pt = float(value)
    if pt.is_integer():
        return str(int(pt))
    return f"{pt:.2f}".rstrip("0").rstrip(".")


def display_font_size_with_name(pt_value: float) -> str:
    """Return Chinese named size (e.g. '小四') for known pt values, else numeric."""
    name = _PT_TO_NAMED_SIZE.get(round(pt_value, 2))
    if name:
        return name
    return format_font_size_pt(pt_value)


def normalize_font_size_display_text(value) -> str:
    """Normalize visible font-size unit text to Chinese while keeping raw content."""
    raw = str(value or "").strip()
    if not raw:
        return ""
    return _FONT_SIZE_PT_DISPLAY_RE.sub(
        lambda match: f"{match.group('number')}磅",
        raw,
    )


def font_size_display_text(value) -> str:
    """Keep numeric values numeric; normalize visible unit text when needed."""
    if value is None:
        return ""
    try:
        pt = float(value)
    except (TypeError, ValueError):
        return normalize_font_size_display_text(value)

    return format_font_size_pt(pt)


def parse_font_size_input(text: str) -> float:
    """Parse numeric pt values or Chinese named font sizes."""
    token = normalize_font_size_token(text)
    if not token:
        raise ValueError("empty font size")

    if token in WORD_NAMED_FONT_SIZE_TO_PT:
        return WORD_NAMED_FONT_SIZE_TO_PT[token]

    name_prefix = token.split("(", 1)[0]
    if name_prefix in WORD_NAMED_FONT_SIZE_TO_PT:
        return WORD_NAMED_FONT_SIZE_TO_PT[name_prefix]

    numeric_token = token.lower()
    for suffix in ("pt", "pts", "磅"):
        if numeric_token.endswith(suffix):
            numeric_token = numeric_token[:-len(suffix)]
            break

    return float(numeric_token)
