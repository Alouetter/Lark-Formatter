from __future__ import annotations

from docx.text.paragraph import Paragraph


_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w(tag: str) -> str:
    return f"{{{_W_NS}}}{tag}"


def _paragraph_ppr(para: Paragraph):
    return para._element.find(_w("pPr"))


def get_paragraph_outline_level(para: Paragraph) -> int | None:
    ppr = _paragraph_ppr(para)
    if ppr is None:
        return None
    outline = ppr.find(_w("outlineLvl"))
    if outline is None:
        return None
    raw = (outline.get(_w("val")) or "").strip()
    if not raw:
        return None
    try:
        return int(raw)
    except (TypeError, ValueError):
        return None


def remove_paragraph_outline_level(para: Paragraph) -> bool:
    """Remove outlineLvl from paragraph pPr. Returns True if removed."""
    ppr = _paragraph_ppr(para)
    if ppr is None:
        return False
    outline = ppr.find(_w("outlineLvl"))
    if outline is None:
        return False
    ppr.remove(outline)
    return True


def get_paragraph_numpr(para: Paragraph) -> tuple[str, int] | None:
    ppr = _paragraph_ppr(para)
    if ppr is None:
        return None
    numpr = ppr.find(_w("numPr"))
    if numpr is None:
        return None

    numid = numpr.find(_w("numId"))
    raw_num_id = (numid.get(_w("val")) or "").strip() if numid is not None else ""
    if not raw_num_id:
        return None

    ilvl = numpr.find(_w("ilvl"))
    raw_ilvl = (ilvl.get(_w("val")) or "").strip() if ilvl is not None else ""
    try:
        level = int(raw_ilvl) if raw_ilvl else 0
    except (TypeError, ValueError):
        level = 0
    return raw_num_id, level


def has_explicit_outline_heading(para: Paragraph) -> bool:
    return get_paragraph_outline_level(para) is not None


def get_paragraph_heading_signal(para: Paragraph) -> dict[str, object]:
    outline_level = get_paragraph_outline_level(para)
    numpr = get_paragraph_numpr(para)
    return {
        "outline_level": outline_level,
        "numpr": numpr,
        "has_outline": outline_level is not None,
        "has_numpr": numpr is not None,
        "has_explicit_outline_heading": outline_level is not None,
    }
