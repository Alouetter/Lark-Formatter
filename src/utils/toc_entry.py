"""Helpers for identifying TOC entry-like paragraph text."""

from __future__ import annotations

import re

_ROMAN_PAGE_CHARS = "\u2160\u2161\u2162\u2163\u2164\u2165\u2166\u2167\u2168\u2169\u216a\u216b"
_CN_NUM_CHARS = "\u3007\u96f6\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341"

_RE_TOC_TAB_PAGE_SUFFIX = re.compile(
    rf"(?:\t\s*(?:\d+|[IVXLCDMivxlcdm]+|[{_ROMAN_PAGE_CHARS}])){{1,3}}\s*$"
)
_RE_TOC_DOT_LEADER_SUFFIX = re.compile(
    rf"[.\uff0e\u3002\u2026\xb7]{{2,}}\s*(?:\d+|[IVXLCDMivxlcdm]+|[{_ROMAN_PAGE_CHARS}])\s*$"
)
_RE_TOC_LOOSE_PAGE_SUFFIX = re.compile(
    rf"(?:\t|[.\uff0e\u3002\u2026\xb7]{{2,}}|\s{{2,}})(?:\d+|[IVXLCDMivxlcdm]+|[{_ROMAN_PAGE_CHARS}])\s*$"
)
_RE_TOC_SINGLE_SPACE_PAGE_SUFFIX = re.compile(
    rf"\s+(?:\d{{1,3}}|[IVXLCDMivxlcdm]+|[{_ROMAN_PAGE_CHARS}])\s*$"
)
_RE_PAGE_SUFFIX = re.compile(
    rf"(?:\d+|[IVXLCDMivxlcdm]+|[{_ROMAN_PAGE_CHARS}])\s*$"
)
_RE_REFERENCE_ENTRY = re.compile(r"^\s*(\[\d{1,4}\]|[\uff08(]\d{1,4}[\uff09)]|\d{1,4}\.)\s+\S")
_RE_REFERENCE_ENTRY_PREFIX = re.compile(
    r"^\s*(?:\[\d{1,4}\]|[\uff08(]\d{1,4}[\uff09)]|\d{1,4}\.)\s*"
)
_RE_REFERENCE_TYPE_MARKER = re.compile(
    r"\[(?:J|M|D|C|R|P|S|N|Z|A|CP|EB/OL|DB/OL|OL)\]",
    re.IGNORECASE,
)
_RE_REFERENCE_YEAR = re.compile(r"(?<!\d)(?:19|20)\d{2}(?!\d)")
_RE_REFERENCE_VOLUME_ISSUE = re.compile(
    r"(?:,|\uff0c)\s*\d+\s*\(\s*\d+\s*\)\s*[:\uff1a]"
)
_RE_REFERENCE_ARTICLE_ID = re.compile(
    r"[:\uff1a]\s*(?:e?\d{4,}|\d+\s*-\s*\d+)\.?\s*$",
    re.IGNORECASE,
)
_RE_REFERENCE_PUB_HINT = re.compile(r"\b(?:doi|vol\.?|no\.?|pp\.?)\b", re.IGNORECASE)
_RE_DATE_PLACEHOLDER_LINE = re.compile(
    rf"^(?:\d{{2,4}}|[{_CN_NUM_CHARS}]{{2,6}})\u5e74(?:\d{{1,2}})?\u6708(?:(?:\d{{1,2}})?\u65e5)?$"
)
_RE_TOC_LEVEL_STYLE = re.compile(r"^(toc|\u76ee\u5f55)\s*\d+$", re.IGNORECASE)
_RE_NARRATIVE_END = re.compile(r"[\u3002\uff01\uff1f!?\uff1b;:]$")
_RE_NUMBERED_HEADING_PREFIX = re.compile(
    r"^(?:"
    r"\u7b2c[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u96f6\d]+[\u7ae0\u8282\u7bc7]"
    r"|[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u96f6]+\u3001"
    r"|\d+(?:\.\d+){0,5}"
    r")"
)


def _norm_no_space(text: str) -> str:
    return re.sub(r"\s+", "", text or "").strip()


def looks_like_toc_entry_line(text: str) -> bool:
    """Return True when a line looks like a TOC entry with page-number suffix."""
    raw = (text or "").strip()
    if not raw:
        return False
    if _RE_TOC_TAB_PAGE_SUFFIX.search(raw):
        return True
    if _RE_TOC_DOT_LEADER_SUFFIX.search(raw):
        return True
    return False


def looks_like_numbered_toc_entry_with_page_suffix(text: str) -> bool:
    """Detect plain-text TOC entries like '\u7b2c\u4e00\u7ae0 \u7eea\u8bba 1' (single-space suffix)."""
    raw = (text or "").strip()
    if not raw:
        return False
    if not _RE_NUMBERED_HEADING_PREFIX.match(raw):
        return False
    if _RE_TOC_LOOSE_PAGE_SUFFIX.search(raw):
        return True
    if _RE_TOC_SINGLE_SPACE_PAGE_SUFFIX.search(raw):
        return True
    return False


def looks_like_bibliographic_reference_line(text: str) -> bool:
    """Return True when a line strongly resembles a bibliography/reference entry."""
    raw = (text or "").strip()
    if len(raw) < 20:
        return False

    has_prefix = bool(_RE_REFERENCE_ENTRY_PREFIX.match(raw))
    has_type_marker = bool(_RE_REFERENCE_TYPE_MARKER.search(raw))
    has_year = bool(_RE_REFERENCE_YEAR.search(raw))
    has_pub_tail = bool(
        _RE_REFERENCE_VOLUME_ISSUE.search(raw)
        or _RE_REFERENCE_ARTICLE_ID.search(raw)
        or _RE_REFERENCE_PUB_HINT.search(raw)
    )
    punctuation_count = sum(
        raw.count(ch)
        for ch in (",", "\uff0c", ".", "\uff0e", ";", "\uff1b", ":", "\uff1a")
    )
    authorish = punctuation_count >= 3

    if has_type_marker and has_year and (has_pub_tail or authorish or has_prefix):
        return True
    if has_prefix and has_year and has_pub_tail and authorish:
        return True
    return False


def looks_like_reference_entry_line(text: str) -> bool:
    raw = (text or "").strip()
    if not raw:
        return False
    if _RE_REFERENCE_ENTRY.match(raw):
        return True
    if not _RE_REFERENCE_ENTRY_PREFIX.match(raw):
        return False
    return looks_like_bibliographic_reference_line(raw)


def looks_like_date_placeholder_line(text: str) -> bool:
    """Detect date placeholders such as '20 \u5e74 \u6708 \u65e5' and concrete Y/M/D lines."""
    norm = _norm_no_space(text)
    if not norm:
        return False
    return bool(_RE_DATE_PLACEHOLDER_LINE.match(norm))


def is_toc_level_style_name(style_name: str) -> bool:
    s = (style_name or "").strip()
    if not s:
        return False
    return bool(_RE_TOC_LEVEL_STYLE.match(s))


def toc_whitelist_score(
    text: str,
    *,
    style_name: str = "",
    style_id: str = "",
    has_pageref: bool = False,
    numbered_heading_like: bool = False,
) -> int:
    """Score TOC-like whitelist signals. Higher means more likely TOC entry."""
    raw = (text or "").strip()
    if not raw:
        return 0

    score = 0
    if is_toc_level_style_name(style_name) or is_toc_level_style_name(style_id):
        score += 4
    if has_pageref:
        score += 4
    if looks_like_toc_entry_line(raw):
        score += 3
    if _RE_TOC_LOOSE_PAGE_SUFFIX.search(raw):
        score += 2
    if numbered_heading_like and looks_like_numbered_toc_entry_with_page_suffix(raw):
        score += 2
    if numbered_heading_like and _RE_PAGE_SUFFIX.search(raw):
        score += 1
    return score


def toc_blacklist_score(text: str) -> int:
    """Score non-TOC signals. Higher means less likely TOC entry."""
    raw = (text or "").strip()
    if not raw:
        return 0

    score = 0
    if looks_like_reference_entry_line(raw):
        score += 4
    if looks_like_date_placeholder_line(raw):
        score += 3
    if len(raw) >= 100:
        score += 2
    if _RE_NARRATIVE_END.search(raw):
        score += 2
    comma_count = raw.count("\uff0c") + raw.count(",") + raw.count("\u3001")
    if comma_count >= 2 and len(raw) >= 24:
        score += 1
    if (
        "\t" not in raw
        and not re.search(r"[.\uff0e\u3002\u2026\xb7]{2,}", raw)
        and len(raw) >= 42
    ):
        score += 1
    return score
