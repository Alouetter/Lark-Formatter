"""字体搜索辅助逻辑。"""

from __future__ import annotations

_FULLWIDTH_ASCII_TRANSLATION = str.maketrans({
    "Ａ": "A",
    "Ｂ": "B",
    "Ｃ": "C",
    "Ｄ": "D",
    "Ｅ": "E",
    "Ｆ": "F",
    "Ｇ": "G",
    "Ｈ": "H",
    "Ｉ": "I",
    "Ｊ": "J",
    "Ｋ": "K",
    "Ｌ": "L",
    "Ｍ": "M",
    "Ｎ": "N",
    "Ｏ": "O",
    "Ｐ": "P",
    "Ｑ": "Q",
    "Ｒ": "R",
    "Ｓ": "S",
    "Ｔ": "T",
    "Ｕ": "U",
    "Ｖ": "V",
    "Ｗ": "W",
    "Ｘ": "X",
    "Ｙ": "Y",
    "Ｚ": "Z",
    "ａ": "a",
    "ｂ": "b",
    "ｃ": "c",
    "ｄ": "d",
    "ｅ": "e",
    "ｆ": "f",
    "ｇ": "g",
    "ｈ": "h",
    "ｉ": "i",
    "ｊ": "j",
    "ｋ": "k",
    "ｌ": "l",
    "ｍ": "m",
    "ｎ": "n",
    "ｏ": "o",
    "ｐ": "p",
    "ｑ": "q",
    "ｒ": "r",
    "ｓ": "s",
    "ｔ": "t",
    "ｕ": "u",
    "ｖ": "v",
    "ｗ": "w",
    "ｘ": "x",
    "ｙ": "y",
    "ｚ": "z",
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
})


def normalize_font_search_text(text: str) -> str:
    """Normalize font names for fuzzy matching."""
    raw = str(text or "").strip().translate(_FULLWIDTH_ASCII_TRANSLATION).casefold()
    return "".join(ch for ch in raw if not ch.isspace() and ch not in "-_.,()[]{}")


def font_matches_query(font_name: str, query: str) -> bool:
    """Case-insensitive fuzzy matching for font family search."""
    normalized_query = normalize_font_search_text(query)
    if not normalized_query:
        return True

    normalized_font = normalize_font_search_text(font_name)
    if not normalized_font:
        return False

    if normalized_query in normalized_font:
        return True

    return _is_subsequence(normalized_query, normalized_font)


def _is_subsequence(query: str, target: str) -> bool:
    if not query:
        return True
    start = 0
    for ch in query:
        found_at = target.find(ch, start)
        if found_at < 0:
            return False
        start = found_at + 1
    return True
