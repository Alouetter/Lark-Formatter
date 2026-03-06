"""Helpers for identifying TOC entry-like paragraph text."""

from __future__ import annotations

import re

_RE_TOC_TAB_PAGE_SUFFIX = re.compile(
    r"(?:\t\s*(?:\d+|[IVXLCDMivxlcdm]+|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]+)){1,3}\s*$"
)
_RE_TOC_DOT_LEADER_SUFFIX = re.compile(
    r"[.．。…·]{2,}\s*(?:\d+|[IVXLCDMivxlcdm]+|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]+)\s*$"
)
_RE_TOC_LOOSE_PAGE_SUFFIX = re.compile(
    r"(?:\t|[.．。…·]{2,}|\s{2,})(?:\d+|[IVXLCDMivxlcdm]+|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]+)\s*$"
)
_RE_TOC_SINGLE_SPACE_PAGE_SUFFIX = re.compile(
    r"\s+(?:\d{1,3}|[IVXLCDMivxlcdm]+|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]+)\s*$"
)
_RE_PAGE_SUFFIX = re.compile(
    r"(?:\d+|[IVXLCDMivxlcdm]+|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]+)\s*$"
)
_RE_REFERENCE_ENTRY = re.compile(r"^\s*(\[\d{1,4}\]|[（(]\d{1,4}[）)]|\d{1,4}\.)\s+\S")
_RE_DATE_PLACEHOLDER_LINE = re.compile(
    r"^(?:\d{2,4}|[〇零一二三四五六七八九十]{2,6})年(?:\d{1,2})?月(?:(?:\d{1,2})?日)?$"
)
_RE_TOC_LEVEL_STYLE = re.compile(r"^(toc|目录)\s*\d+$", re.IGNORECASE)
_RE_NARRATIVE_END = re.compile(r"[。！？!?；;:]$")
_RE_NUMBERED_HEADING_PREFIX = re.compile(
    r"^(?:"
    r"第[一二三四五六七八九十百零\d]+[章节篇]"
    r"|[一二三四五六七八九十百零]+、"
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
    """Detect plain-text TOC entries like '第一章 绪论 1' (single-space suffix)."""
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


def looks_like_reference_entry_line(text: str) -> bool:
    return bool(_RE_REFERENCE_ENTRY.match((text or "").strip()))


def looks_like_date_placeholder_line(text: str) -> bool:
    """Detect date placeholders such as '20 年 月 日' and concrete Y/M/D lines."""
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
    comma_count = raw.count("，") + raw.count(",") + raw.count("、")
    if comma_count >= 2 and len(raw) >= 24:
        score += 1
    if (
        "\t" not in raw
        and not re.search(r"[.．。…·]{2,}", raw)
        and len(raw) >= 42
    ):
        score += 1
    return score
