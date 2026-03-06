"""Clone style/page setup settings from a reference DOCX into SceneConfig."""

from __future__ import annotations

from collections import Counter, defaultdict
import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn

from src.scene.schema import HeadingLevelConfig, SceneConfig, StyleConfig
from src.utils.toc_entry import (
    looks_like_numbered_toc_entry_with_page_suffix,
    looks_like_toc_entry_line,
)


STYLE_CANDIDATES: dict[str, list[str]] = {
    "normal": ["Normal", "正文"],
    "heading1": ["Heading 1", "标题 1", "章标题", "一级标题"],
    "heading2": ["Heading 2", "标题 2", "二级标题"],
    "heading3": ["Heading 3", "标题 3", "三级标题"],
    "heading4": ["Heading 4", "标题 4", "四级标题", "4", "附录章标题", "附录一级条标题"],
    "heading5": ["Heading 5", "标题 5", "五级标题", "5", "附录二级条标题"],
    "heading6": ["Heading 6", "标题 6", "六级标题", "6", "附录三级条标题"],
    "heading7": ["Heading 7", "标题 7", "七级标题", "7", "附录四级条标题"],
    "heading8": ["Heading 8", "标题 8", "八级标题", "8", "附录五级条标题"],
    "abstract_title_cn": ["中文摘要标题", "摘要标题", "Abstract Title"],
    "abstract_title_en": ["英文摘要标题", "Abstract Title"],
    "abstract_body": ["中文摘要正文", "摘要正文", "Normal"],
    "abstract_body_en": ["英文摘要正文", "Normal"],
    "toc_title": ["TOC Heading", "目录标题"],
    "toc_chapter": ["TOC 1", "目录 1"],
    "toc_level1": ["TOC 2", "目录 2"],
    "toc_level2": ["TOC 3", "目录 3"],
    "references_body": ["参考文献", "Bibliography", "Normal"],
    "acknowledgment_body": ["致谢正文", "Normal"],
    "appendix_body": ["附录正文", "Normal"],
    "resume_body": ["个人简历正文", "Normal"],
    "symbol_table_body": ["符号注释表", "Normal"],
    "code_block": ["Code", "代码块", "Normal"],
    "figure_caption": ["Caption", "图题注"],
    "table_caption": ["Caption", "表题注"],
    "header_cn": ["Header", "页眉"],
    "header_en": ["Header", "页眉"],
    "page_number": ["Footer", "页脚"],
}

ALIGNMENT_REVERSE = {
    WD_ALIGN_PARAGRAPH.LEFT: "left",
    WD_ALIGN_PARAGRAPH.CENTER: "center",
    WD_ALIGN_PARAGRAPH.RIGHT: "right",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
}

_RE_TABLE_CAPTION = re.compile(r"^\s*表\s*\d+(?:[.\-—–:：]\d+)*\s*\S")
_RE_FIGURE_CAPTION = re.compile(r"^\s*(?:图|Fig\.?)\s*\d+(?:[.\-—–:：]\d+)*\s*\S", re.IGNORECASE)
_PARAGRAPH_SAMPLE_RESET_INDENT_KEYS = {
    "abstract_title_cn",
    "abstract_title_en",
    "toc_title",
    "table_caption",
    "figure_caption",
}
_MIN_STYLE_SLOTS_FOR_CLONE = ("header_cn", "header_en", "page_number")
_HEADING_LEVEL_KEYS = [f"heading{i}" for i in range(1, 9)]
_TOC_TITLE_NORMS = {"目录", "目次", "contents", "tableofcontents"}
_RE_TOC_STYLE_LEVEL = re.compile(r"^(?:toc|目录)\s*(\d+)$", re.IGNORECASE)
_RE_TOC_LEVEL2_PREFIX = re.compile(
    r"^(?:\d+\.\d+\.\d+(?:\.\d+)*|[一二三四五六七八九十百零两]+、|[（(][一二三四五六七八九十百零两]+[）)])"
)
_RE_TOC_LEVEL1_PREFIX = re.compile(
    r"^(?:\d+\.\d+(?:\.\d+)*|第[一二三四五六七八九十百零两0-9]+节)"
)

_ARABIC_HEADING_RE = re.compile(
    r"^\s*(?P<num>\d+(?:[.\uFF0E\u3002\uFF61\uFE52]\d+)*)"
    r"(?P<sep>[ \t\u3000]+)"
    r"(?P<title>\S.+?)\s*$"
)
_CHAPTER_HEADING_RE = re.compile(
    r"^\s*(?P<num>第[一二三四五六七八九十百千万零〇两0-9]+章)"
    r"(?P<sep>[ \t\u3000]*)"
    r"(?P<title>\S.+?)\s*$"
)
_SECTION_HEADING_RE = re.compile(
    r"^\s*(?P<num>第[一二三四五六七八九十百千万零〇两0-9]+节)"
    r"(?P<sep>[ \t\u3000]*)"
    r"(?P<title>\S.+?)\s*$"
)
_CN_ORDINAL_HEADING_RE = re.compile(
    r"^\s*(?P<num>[一二三四五六七八九十百千万零〇两]+、)"
    r"(?P<sep>[ \t\u3000]*)"
    r"(?P<title>\S.+?)\s*$"
)
_CN_ORDINAL_PAREN_RE = re.compile(
    r"^\s*(?P<num>[（(][一二三四五六七八九十百千万零〇两]+[）)])"
    r"(?P<sep>[ \t\u3000]*)"
    r"(?P<title>\S.+?)\s*$"
)
_TOC_LEADER_TAIL_RE = re.compile(r"[\.．。…·]{3,}\s*\d+\s*$")


def _iter_style_chain(style):
    cur = style
    while cur is not None:
        yield cur
        cur = cur.base_style


def _first_non_none(style, getter):
    for st in _iter_style_chain(style):
        try:
            value = getter(st)
        except Exception:
            value = None
        if value is not None:
            return value
    return None


def _rfonts_value(style, attr: str) -> str | None:
    key = qn(attr)
    for st in _iter_style_chain(style):
        rpr = st.element.find(qn("w:rPr"))
        if rpr is None:
            continue
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            continue
        val = rfonts.get(key)
        if val:
            return val
    return None


def _size_pt(style) -> float | None:
    size_len = _first_non_none(style, lambda st: st.font.size)
    if size_len is not None:
        try:
            return float(size_len.pt)
        except Exception:
            pass

    for st in _iter_style_chain(style):
        rpr = st.element.find(qn("w:rPr"))
        if rpr is None:
            continue
        sz = rpr.find(qn("w:sz"))
        if sz is None:
            continue
        raw = sz.get(qn("w:val"))
        if not raw:
            continue
        try:
            return float(raw) / 2.0
        except Exception:
            continue
    return None


def _length_pt_or_none(value) -> float | None:
    if value is None:
        return None
    if hasattr(value, "pt"):
        try:
            return float(value.pt)
        except Exception:
            return None
    return None


def _build_style_config(src_style, base: StyleConfig) -> StyleConfig:
    sc = StyleConfig(**vars(base))

    font_cn = _rfonts_value(src_style, "w:eastAsia")
    if font_cn:
        sc.font_cn = font_cn

    font_en = (
        _rfonts_value(src_style, "w:ascii")
        or _rfonts_value(src_style, "w:hAnsi")
        or _first_non_none(src_style, lambda st: st.font.name)
    )
    if font_en:
        sc.font_en = font_en
        if not font_cn:
            sc.font_cn = sc.font_cn or font_en

    size_pt = _size_pt(src_style)
    if size_pt and size_pt > 0:
        sc.size_pt = round(size_pt, 2)

    bold = _first_non_none(src_style, lambda st: st.font.bold)
    if bold is not None:
        sc.bold = bool(bold)
    italic = _first_non_none(src_style, lambda st: st.font.italic)
    if italic is not None:
        sc.italic = bool(italic)

    align = _first_non_none(src_style, lambda st: st.paragraph_format.alignment)
    if align in ALIGNMENT_REVERSE:
        sc.alignment = ALIGNMENT_REVERSE[align]

    before_pt = _length_pt_or_none(_first_non_none(src_style, lambda st: st.paragraph_format.space_before))
    if before_pt is not None:
        sc.space_before_pt = max(0.0, round(before_pt, 2))
    after_pt = _length_pt_or_none(_first_non_none(src_style, lambda st: st.paragraph_format.space_after))
    if after_pt is not None:
        sc.space_after_pt = max(0.0, round(after_pt, 2))

    first_indent_pt = _length_pt_or_none(_first_non_none(src_style, lambda st: st.paragraph_format.first_line_indent))
    if first_indent_pt is not None and sc.size_pt > 0:
        sc.first_line_indent_chars = round(first_indent_pt / sc.size_pt, 2)

    left_indent_pt = _length_pt_or_none(_first_non_none(src_style, lambda st: st.paragraph_format.left_indent))
    if left_indent_pt is not None and sc.size_pt > 0:
        sc.left_indent_chars = round(left_indent_pt / sc.size_pt, 2)

    line_rule = _first_non_none(src_style, lambda st: st.paragraph_format.line_spacing_rule)
    line_spacing = _first_non_none(src_style, lambda st: st.paragraph_format.line_spacing)
    line_spacing_pt = _length_pt_or_none(line_spacing)
    if line_rule in {
        WD_LINE_SPACING.SINGLE,
        WD_LINE_SPACING.ONE_POINT_FIVE,
        WD_LINE_SPACING.DOUBLE,
        WD_LINE_SPACING.MULTIPLE,
    }:
        sc.line_spacing_type = "multiple"
        if isinstance(line_spacing, (int, float)):
            sc.line_spacing_pt = round(max(0.1, float(line_spacing)), 2)
        elif line_rule == WD_LINE_SPACING.SINGLE:
            sc.line_spacing_pt = 1.0
        elif line_rule == WD_LINE_SPACING.ONE_POINT_FIVE:
            sc.line_spacing_pt = 1.5
        elif line_rule == WD_LINE_SPACING.DOUBLE:
            sc.line_spacing_pt = 2.0
    elif line_spacing_pt is not None and line_spacing_pt > 0:
        sc.line_spacing_type = "exact"
        sc.line_spacing_pt = round(line_spacing_pt, 2)
    elif isinstance(line_spacing, (int, float)) and float(line_spacing) > 0:
        sc.line_spacing_type = "multiple"
        sc.line_spacing_pt = round(max(0.1, float(line_spacing)), 2)

    return sc


def _norm_no_space(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def _run_rfonts_value(run, attr: str) -> str | None:
    rpr = run._element.find(qn("w:rPr"))
    if rpr is None:
        return None
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        return None
    return rfonts.get(qn(attr))


def _build_style_config_from_paragraph(para, base: StyleConfig) -> StyleConfig:
    if para.style is not None:
        sc = _build_style_config(para.style, base)
    else:
        sc = StyleConfig(**vars(base))

    pf = para.paragraph_format
    if pf is not None:
        if pf.alignment in ALIGNMENT_REVERSE:
            sc.alignment = ALIGNMENT_REVERSE[pf.alignment]

        before_pt = _length_pt_or_none(pf.space_before)
        if before_pt is not None:
            sc.space_before_pt = max(0.0, round(before_pt, 2))
        after_pt = _length_pt_or_none(pf.space_after)
        if after_pt is not None:
            sc.space_after_pt = max(0.0, round(after_pt, 2))

        first_indent_pt = _length_pt_or_none(pf.first_line_indent)
        if first_indent_pt is not None and sc.size_pt > 0:
            sc.first_line_indent_chars = round(first_indent_pt / sc.size_pt, 2)
        left_indent_pt = _length_pt_or_none(pf.left_indent)
        if left_indent_pt is not None and sc.size_pt > 0:
            sc.left_indent_chars = round(left_indent_pt / sc.size_pt, 2)

        line_rule = pf.line_spacing_rule
        line_spacing = pf.line_spacing
        line_spacing_pt = _length_pt_or_none(line_spacing)
        if line_rule in {
            WD_LINE_SPACING.SINGLE,
            WD_LINE_SPACING.ONE_POINT_FIVE,
            WD_LINE_SPACING.DOUBLE,
            WD_LINE_SPACING.MULTIPLE,
        }:
            sc.line_spacing_type = "multiple"
            if isinstance(line_spacing, (int, float)):
                sc.line_spacing_pt = round(max(0.1, float(line_spacing)), 2)
            elif line_rule == WD_LINE_SPACING.SINGLE:
                sc.line_spacing_pt = 1.0
            elif line_rule == WD_LINE_SPACING.ONE_POINT_FIVE:
                sc.line_spacing_pt = 1.5
            elif line_rule == WD_LINE_SPACING.DOUBLE:
                sc.line_spacing_pt = 2.0
        elif line_spacing_pt is not None and line_spacing_pt > 0:
            sc.line_spacing_type = "exact"
            sc.line_spacing_pt = round(line_spacing_pt, 2)

    runs = [run for run in para.runs if (run.text or "").strip()]
    if runs:
        run0 = runs[0]
        font_cn = _run_rfonts_value(run0, "w:eastAsia")
        font_en = (
            _run_rfonts_value(run0, "w:ascii")
            or _run_rfonts_value(run0, "w:hAnsi")
            or run0.font.name
        )
        if font_cn:
            sc.font_cn = font_cn
        if font_en:
            sc.font_en = font_en
            if not font_cn:
                sc.font_cn = sc.font_cn or font_en

        size_len = run0.font.size
        if size_len is not None:
            try:
                size_pt = float(size_len.pt)
            except Exception:
                size_pt = None
            if size_pt and size_pt > 0:
                sc.size_pt = round(size_pt, 2)

        bold = run0.bold if run0.bold is not None else run0.font.bold
        if bold is not None:
            sc.bold = bool(bold)
        italic = run0.italic if run0.italic is not None else run0.font.italic
        if italic is not None:
            sc.italic = bool(italic)

    return sc


def _collect_paragraph_samples(doc: Document) -> dict[str, object]:
    samples: dict[str, object] = {}
    table_caption_para = None
    figure_caption_para = None
    toc_title_idx = None

    for idx, para in enumerate(doc.paragraphs):
        text = (para.text or "").strip()
        if not text:
            continue

        normalized = _norm_no_space(text)
        normalized_lower = normalized.lower()

        if "abstract_title_cn" not in samples and normalized == "摘要":
            samples["abstract_title_cn"] = para
        if "abstract_title_en" not in samples and normalized_lower == "abstract":
            samples["abstract_title_en"] = para
        if "toc_title" not in samples and normalized_lower in _TOC_TITLE_NORMS:
            samples["toc_title"] = para
            toc_title_idx = idx

        if table_caption_para is None and _RE_TABLE_CAPTION.match(text):
            table_caption_para = para
        if figure_caption_para is None and _RE_FIGURE_CAPTION.match(text):
            figure_caption_para = para

    if table_caption_para is not None:
        samples["table_caption"] = table_caption_para
    if figure_caption_para is not None:
        samples["figure_caption"] = figure_caption_para
    elif table_caption_para is not None:
        # Many规范文档 only include table-caption examples. Reuse for figure caption.
        samples["figure_caption"] = table_caption_para

    if toc_title_idx is not None:
        found_entry = False
        blank_streak = 0
        for para in doc.paragraphs[toc_title_idx + 1:]:
            text = (para.text or "").strip()
            style_name = ((para.style.name if para.style else "") or "").strip()
            style_id = (
                (getattr(para.style, "style_id", "") if para.style else "") or ""
            ).strip()
            if not text:
                if found_entry:
                    blank_streak += 1
                    if blank_streak >= 3:
                        break
                continue

            blank_streak = 0
            level_style_key = _infer_toc_style_key_from_para(text, style_name, style_id)
            if level_style_key:
                if level_style_key not in samples:
                    samples[level_style_key] = para
                found_entry = True
                continue

            if found_entry:
                break

    return samples


def _ensure_clone_style_slots(config: SceneConfig) -> list[str]:
    added: list[str] = []
    for key in (
        "abstract_title_cn",
        "abstract_title_en",
        "abstract_body",
        "abstract_body_en",
        "toc_title",
        "toc_chapter",
        "toc_level1",
        "toc_level2",
        *_MIN_STYLE_SLOTS_FOR_CLONE,
    ):
        if key in config.styles:
            continue
        config.styles[key] = StyleConfig()
        added.append(key)
    return added


def _infer_toc_style_key_from_para(text: str, style_name: str, style_id: str) -> str | None:
    style_levels = []
    for s in (style_name, style_id):
        m = _RE_TOC_STYLE_LEVEL.match((s or "").strip())
        if not m:
            continue
        try:
            style_levels.append(int(m.group(1)))
        except Exception:
            continue
    if style_levels:
        level = min(style_levels)
        if level <= 1:
            return "toc_chapter"
        if level == 2:
            return "toc_level1"
        return "toc_level2"

    if not (
        looks_like_toc_entry_line(text)
        or looks_like_numbered_toc_entry_with_page_suffix(text)
    ):
        return None

    raw = (text or "").strip()
    if _RE_TOC_LEVEL2_PREFIX.match(raw):
        return "toc_level2"
    if _RE_TOC_LEVEL1_PREFIX.match(raw):
        return "toc_level1"
    return "toc_chapter"


def _normalize_separator(raw: str) -> str:
    if not raw:
        return ""
    if "\u3000" in raw:
        return "\u3000"
    if "\t" in raw:
        return "\t"
    if " " in raw:
        return " "
    return raw


def _looks_like_toc_tail(title: str) -> bool:
    tail = (title or "").strip()
    if not tail:
        return False
    return bool(_TOC_LEADER_TAIL_RE.search(tail))


def _infer_heading_pattern(text: str):
    text = (text or "").strip()
    if not text or "\t" in text:
        return None

    m = _CHAPTER_HEADING_RE.match(text)
    if m:
        title = m.group("title")
        if _looks_like_toc_tail(title):
            return None
        return ("heading1", "chinese_chapter", "第{cn}章", _normalize_separator(m.group("sep")), title)

    m = _SECTION_HEADING_RE.match(text)
    if m:
        title = m.group("title")
        if _looks_like_toc_tail(title):
            return None
        return ("heading2", "chinese_section", "第{cn}节", _normalize_separator(m.group("sep")), title)

    m = _CN_ORDINAL_HEADING_RE.match(text)
    if m:
        title = m.group("title")
        if _looks_like_toc_tail(title):
            return None
        return ("heading3", "chinese_ordinal", "{cn}、", _normalize_separator(m.group("sep")), title)

    m = _CN_ORDINAL_PAREN_RE.match(text)
    if m:
        num = m.group("num")
        template = "({cn})" if num.startswith("(") else "（{cn}）"
        title = m.group("title")
        if _looks_like_toc_tail(title):
            return None
        return ("heading4", "chinese_ordinal_paren", template, _normalize_separator(m.group("sep")), title)

    m = _ARABIC_HEADING_RE.match(text)
    if not m:
        return None

    number_text = m.group("num")
    parts = [p for p in re.split(r"[.\uFF0E\u3002\uFF61\uFE52]", number_text) if p]
    if not parts:
        return None
    if len(parts[0]) >= 4:
        try:
            if int(parts[0]) >= 1900:
                return None
        except ValueError:
            pass

    level_idx = min(len(parts), 8)
    level_key = f"heading{level_idx}"
    format_name = "arabic" if level_idx == 1 else "arabic_dotted"
    template = "{n}" if level_idx == 1 else "{parent}.{n}"
    title = m.group("title")
    if _looks_like_toc_tail(title):
        return None
    if title.startswith(("年", "月", "日")):
        return None
    return (level_key, format_name, template, _normalize_separator(m.group("sep")), title)


def _guess_heading_indent_chars(para, style_cfg: StyleConfig | None) -> float | None:
    left_pt = _length_pt_or_none(para.paragraph_format.left_indent)
    if left_pt is None or left_pt <= 0:
        return None
    size_pt = 0.0
    if style_cfg is not None:
        try:
            size_pt = float(getattr(style_cfg, "size_pt", 0) or 0)
        except Exception:
            size_pt = 0.0
    if size_pt <= 0:
        size_pt = 12.0
    return round(max(0.0, left_pt) / size_pt, 2)


def _build_default_heading_numbering_levels(config: SceneConfig) -> dict[str, HeadingLevelConfig]:
    defaults: dict[str, HeadingLevelConfig] = {}
    existing = getattr(config.heading_numbering, "levels", {}) or {}
    for idx in range(1, 5):
        key = f"heading{idx}"
        base = existing.get(key, HeadingLevelConfig())
        lc = HeadingLevelConfig(**vars(base))
        lc.format = "arabic" if idx == 1 else "arabic_dotted"
        lc.template = "{n}" if idx == 1 else "{parent}.{n}"
        lc.separator = lc.separator or "\u3000"
        lc.custom_separator = None
        defaults[key] = lc
    return defaults


def _clone_heading_numbering(config: SceneConfig, doc: Document) -> dict:
    samples_by_level: dict[str, list[dict[str, object]]] = defaultdict(list)
    sample_count = 0

    for para in doc.paragraphs:
        style_name = ((para.style.name if para.style is not None else "") or "").strip().lower()
        if "toc" in style_name or "目录" in style_name:
            continue

        text = (para.text or "").strip()
        inferred = _infer_heading_pattern(text)
        if inferred is None:
            continue

        level_key, fmt, template, separator, _title = inferred
        if level_key not in _HEADING_LEVEL_KEYS:
            continue

        sample = {
            "format": fmt,
            "template": template,
            "separator": separator,
        }
        align = ALIGNMENT_REVERSE.get(para.paragraph_format.alignment)
        if align:
            sample["alignment"] = align

        style_cfg = config.styles.get(level_key) or config.styles.get("normal")
        indent_chars = _guess_heading_indent_chars(para, style_cfg)
        if indent_chars is not None:
            sample["left_indent_chars"] = indent_chars

        samples_by_level[level_key].append(sample)
        sample_count += 1

    inferred_levels: dict[str, HeadingLevelConfig] = {}
    existing_levels = getattr(config.heading_numbering, "levels", {}) or {}

    for level_key in _HEADING_LEVEL_KEYS:
        level_samples = samples_by_level.get(level_key)
        if not level_samples:
            continue

        base = existing_levels.get(level_key, HeadingLevelConfig())
        lc = HeadingLevelConfig(**vars(base))

        fmt_template_counter = Counter(
            (str(s["format"]), str(s["template"])) for s in level_samples
        )
        (fmt, template), _ = fmt_template_counter.most_common(1)[0]
        lc.format = fmt
        lc.template = template

        separator_counter = Counter(str(s["separator"]) for s in level_samples)
        lc.separator = separator_counter.most_common(1)[0][0]
        lc.custom_separator = None

        align_counter = Counter(
            str(s["alignment"]) for s in level_samples if s.get("alignment")
        )
        if align_counter:
            align_value, align_count = align_counter.most_common(1)[0]
            if align_count >= 2:
                lc.alignment = align_value

        indent_values = [
            float(s["left_indent_chars"])
            for s in level_samples
            if s.get("left_indent_chars") is not None
        ]
        if len(indent_values) >= 2:
            lc.left_indent_chars = round(sum(indent_values) / len(indent_values), 2)

        inferred_levels[level_key] = lc

    use_inferred = bool(inferred_levels) and (
        sample_count >= 3 or (sample_count >= 2 and len(inferred_levels) >= 2)
    )

    levels_updated: dict[str, HeadingLevelConfig] | None = None
    inferred_used = False
    fallback_used = False
    before_levels = getattr(config.heading_numbering, "levels", {}) or {}

    if use_inferred:
        levels_updated = inferred_levels
        inferred_used = True
    elif before_levels:
        levels_updated = None
    else:
        levels_updated = _build_default_heading_numbering_levels(config)
        fallback_used = True

    if levels_updated is not None:
        config.heading_numbering.levels = dict(levels_updated)
        scheme_id = str(getattr(config.heading_numbering, "scheme", "") or "").strip()
        if scheme_id:
            schemes = getattr(config.heading_numbering, "schemes", {}) or {}
            schemes[scheme_id] = dict(levels_updated)
            config.heading_numbering.schemes = schemes

    return {
        "numbering_updated": levels_updated is not None,
        "numbering_levels_updated": list(levels_updated.keys()) if levels_updated else [],
        "numbering_inferred": inferred_used,
        "numbering_fallback_used": fallback_used,
        "numbering_samples": sample_count,
    }


def _find_style(doc: Document, style_key: str):
    candidates = [style_key]
    candidates.extend(STYLE_CANDIDATES.get(style_key, []))

    styles = list(doc.styles)
    lowered = {(st.name or "").strip().lower(): st for st in styles if st.name}
    for name in candidates:
        key = (name or "").strip().lower()
        if not key:
            continue
        st = lowered.get(key)
        if st is not None:
            return st
    return None


def _clone_page_setup(config: SceneConfig, doc: Document) -> bool:
    if not doc.sections:
        return False
    sec = doc.sections[0]
    ps = config.page_setup

    changed = False

    def _set_cm(attr: str, value):
        nonlocal changed
        old = getattr(ps.margin, attr)
        new = round(float(value.cm), 2)
        if abs(old - new) > 1e-6:
            setattr(ps.margin, attr, new)
            changed = True

    _set_cm("top_cm", sec.top_margin)
    _set_cm("bottom_cm", sec.bottom_margin)
    _set_cm("left_cm", sec.left_margin)
    _set_cm("right_cm", sec.right_margin)

    header_new = round(float(sec.header_distance.cm), 2)
    footer_new = round(float(sec.footer_distance.cm), 2)
    if abs(ps.header_distance_cm - header_new) > 1e-6:
        ps.header_distance_cm = header_new
        changed = True
    if abs(ps.footer_distance_cm - footer_new) > 1e-6:
        ps.footer_distance_cm = footer_new
        changed = True

    w_cm = round(float(sec.page_width.cm), 2)
    h_cm = round(float(sec.page_height.cm), 2)
    detected_paper_size = None
    paper_specs = {
        "A4": (21.0, 29.7),
        "Letter": (21.59, 27.94),
        "A3": (29.7, 42.0),
    }
    for paper_name, (pw, ph) in paper_specs.items():
        if (
            abs(w_cm - pw) <= 0.3 and abs(h_cm - ph) <= 0.3
        ) or (
            abs(w_cm - ph) <= 0.3 and abs(h_cm - pw) <= 0.3
        ):
            detected_paper_size = paper_name
            break

    if detected_paper_size and ps.paper_size != detected_paper_size:
        ps.paper_size = detected_paper_size
        changed = True

    return changed


def clone_scene_style_from_docx(config: SceneConfig, template_docx_path: str | Path) -> dict:
    """Clone style/page settings from template docx into scene config.

    Returns:
        {
            "styles_updated": [...style_keys...],
            "styles_missing": [...style_keys...],
            "page_setup_updated": bool
        }
    """
    from src.docx_io.sanitize import sanitize_docx

    docx_path = str(template_docx_path)
    try:
        sanitize_docx(docx_path)
    except Exception:
        pass  # best-effort sanitize

    try:
        doc = Document(docx_path)
    except Exception:
        # Fallback for templates with parser-blocking oversized XML attributes.
        sanitize_docx(docx_path, aggressive=True)
        try:
            doc = Document(docx_path)
        except Exception as exc:
            raise RuntimeError(
                "无法读取模板 DOCX：检测到异常 XML 内容（可能存在超长属性值）。"
            ) from exc
    style_slots_added = _ensure_clone_style_slots(config)
    updated: list[str] = []
    missing: list[str] = []

    for key, sc in config.styles.items():
        src_style = _find_style(doc, key)
        if src_style is None:
            missing.append(key)
            continue
        config.styles[key] = _build_style_config(src_style, sc)
        updated.append(key)

    para_samples = _collect_paragraph_samples(doc)
    for key, para in para_samples.items():
        base = config.styles.get(key)
        if base is None:
            continue
        sc = _build_style_config_from_paragraph(para, base)
        if key in _PARAGRAPH_SAMPLE_RESET_INDENT_KEYS:
            sc.first_line_indent_chars = 0.0
            sc.left_indent_chars = 0.0
        config.styles[key] = sc
        if key not in updated:
            updated.append(key)
        if key in missing:
            missing.remove(key)

    page_setup_updated = _clone_page_setup(config, doc)
    numbering_summary = _clone_heading_numbering(config, doc)

    return {
        "styles_updated": updated,
        "styles_missing": missing,
        "styles_added": style_slots_added,
        "page_setup_updated": page_setup_updated,
        **numbering_summary,
    }
