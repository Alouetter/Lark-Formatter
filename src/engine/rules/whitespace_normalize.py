"""Whitespace normalization and context-aware full/half-width conversion."""

import re

from docx import Document
from docx.oxml import OxmlElement

from src.engine.change_tracker import ChangeTracker
from src.engine.rules.base import BaseRule
from src.scene.schema import SceneConfig

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_XML_NS = "http://www.w3.org/XML/1998/namespace"

_SPACE_VARIANTS_RE = re.compile(r"[\u00A0\u1680\u2000-\u200A\u202F\u205F\u3000]")
_ZERO_WIDTH_RE = re.compile(r"[\u200B\u200C\u200D\u2060\uFEFF]")
_MULTI_SPACE_RE = re.compile(r" {2,}")

_URL_EMAIL_PATH_RE = re.compile(
    r"(https?://\S+|www\.\S+|\b[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}\b|[A-Za-z]:\\[^\s]+|\\\\[^\s]+)"
)
_REF_NUM_RE = re.compile(
    r"(\[\s*\d+(?:\s*[-,]\s*\d+)*\s*\]|\(\s*\d+(?:\s*[-,]\s*\d+)*\s*\)|（\s*\d+(?:\s*[-,]\s*\d+)*\s*）)"
)

_HALF_TO_FULL_PUNC = {
    ",": "，",
    ".": "。",
    ":": "：",
    ";": "；",
    "!": "！",
    "?": "？",
}
_FULL_TO_HALF_PUNC = {v: k for k, v in _HALF_TO_FULL_PUNC.items()}


def _w(tag: str) -> str:
    return f"{{{_W_NS}}}{tag}"


def _sync_xml_space_attr(t_el) -> None:
    """Keep xml:space preserve only when text starts/ends with regular spaces."""
    text = t_el.text or ""
    key = f"{{{_XML_NS}}}space"
    if text.startswith(" ") or text.endswith(" "):
        t_el.set(key, "preserve")
    else:
        t_el.attrib.pop(key, None)


def _replace_tab_elements_with_spaces(p_el) -> int:
    """Convert Word native <w:tab/> nodes into literal spaces for later cleanup."""
    changed = 0
    for tab_el in list(p_el.findall(f".//{_w('tab')}")):
        parent = tab_el.getparent()
        if parent is None:
            continue
        t_el = OxmlElement("w:t")
        t_el.text = " "
        _sync_xml_space_attr(t_el)
        parent.replace(tab_el, t_el)
        changed += 1
    return changed


def _normalize_inline_text(
        text: str,
        *,
        normalize_space_variants: bool,
        convert_tabs: bool,
        remove_zero_width: bool,
        collapse_multiple_spaces: bool) -> str:
    s = text or ""
    if not s:
        return s
    if remove_zero_width:
        s = _ZERO_WIDTH_RE.sub("", s)
    if normalize_space_variants:
        s = _SPACE_VARIANTS_RE.sub(" ", s)
    if convert_tabs:
        s = s.replace("\t", " ")
    if collapse_multiple_spaces:
        s = _MULTI_SPACE_RE.sub(" ", s)
    return s


def _is_skipped_style(para) -> bool:
    style = getattr(para, "style", None)
    style_name = (getattr(style, "name", "") or "")
    low = style_name.lower()
    if any(k in low for k in ("heading", "toc", "caption", "code")):
        return True
    if any(k in style_name for k in ("\u6807\u9898", "\u76EE\u5F55", "\u9898\u6CE8", "\u4EE3\u7801")):
        return True
    return False


def _has_complex_content(p_el) -> bool:
    """Skip paragraphs with equation/object/field content."""
    if p_el.findall(f".//{{{_M_NS}}}oMath") or p_el.findall(f".//{{{_M_NS}}}oMathPara"):
        return True
    if p_el.findall(f".//{_w('object')}") or p_el.findall(f".//{_w('drawing')}"):
        return True
    if p_el.findall(f".//{_w('pict')}"):
        return True
    for instr in p_el.findall(f".//{_w('instrText')}"):
        txt = (instr.text or "").upper()
        if "TOC" in txt:
            return True
    return False


def _trim_paragraph_edges(t_nodes: list) -> bool:
    changed = False
    first_idx = None
    last_idx = None

    for i, t_el in enumerate(t_nodes):
        if t_el.text:
            first_idx = i
            break
    for i in range(len(t_nodes) - 1, -1, -1):
        if t_nodes[i].text:
            last_idx = i
            break

    if first_idx is not None:
        old = t_nodes[first_idx].text or ""
        new = old.lstrip(" ")
        if new != old:
            t_nodes[first_idx].text = new
            _sync_xml_space_attr(t_nodes[first_idx])
            changed = True

    if last_idx is not None:
        old = t_nodes[last_idx].text or ""
        new = old.rstrip(" ")
        if new != old:
            t_nodes[last_idx].text = new
            _sync_xml_space_attr(t_nodes[last_idx])
            changed = True

    return changed


def _is_cjk(ch: str) -> bool:
    if not ch:
        return False
    cp = ord(ch)
    return (
        0x3400 <= cp <= 0x4DBF
        or 0x4E00 <= cp <= 0x9FFF
        or 0x3040 <= cp <= 0x30FF
        or 0xAC00 <= cp <= 0xD7AF
    )


def _is_latin_or_digit(ch: str) -> bool:
    if not ch:
        return False
    return ("A" <= ch <= "Z") or ("a" <= ch <= "z") or ch.isdigit()


def _script_kind(ch: str) -> str | None:
    if _is_cjk(ch):
        return "zh"
    if _is_latin_or_digit(ch):
        return "en"
    return None


def _count_lang_chars(text: str) -> tuple[int, int]:
    zh = 0
    en = 0
    for ch in text:
        if _is_cjk(ch):
            zh += 1
        elif _is_latin_or_digit(ch):
            en += 1
    return zh, en


def _merge_spans(spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    if not spans:
        return []
    spans.sort()
    merged = [spans[0]]
    for s, e in spans[1:]:
        ms, me = merged[-1]
        if s <= me:
            merged[-1] = (ms, max(me, e))
        else:
            merged.append((s, e))
    return merged


def _build_protected_spans(text: str, *, protect_reference_numbering: bool) -> list[tuple[int, int]]:
    spans: list[tuple[int, int]] = []
    for m in _URL_EMAIL_PATH_RE.finditer(text):
        spans.append((m.start(), m.end()))
    if protect_reference_numbering:
        for m in _REF_NUM_RE.finditer(text):
            spans.append((m.start(), m.end()))
    return _merge_spans(spans)


def _in_spans(idx: int, spans: list[tuple[int, int]]) -> bool:
    for s, e in spans:
        if s <= idx < e:
            return True
        if idx < s:
            return False
    return False


def _fullwidth_alnum_to_half(ch: str) -> str:
    cp = ord(ch)
    if 0xFF10 <= cp <= 0xFF19:
        return chr(cp - 0xFEE0)
    if 0xFF21 <= cp <= 0xFF3A:
        return chr(cp - 0xFEE0)
    if 0xFF41 <= cp <= 0xFF5A:
        return chr(cp - 0xFEE0)
    return ch


def _nearest_script(chars: list[str], idx: int, step: int, max_hops: int = 12) -> str | None:
    hops = 0
    j = idx + step
    while 0 <= j < len(chars) and hops < max_hops:
        ch = chars[j]
        if ch.isspace():
            j += step
            hops += 1
            continue
        kind = _script_kind(ch)
        if kind:
            return kind
        j += step
        hops += 1
    return None


def _resolve_context_language(
        chars: list[str],
        idx: int,
        para_counts: tuple[int, int],
        min_confidence: int) -> str | None:
    zh_score = 0
    en_score = 0
    para_zh, para_en = para_counts
    if para_zh >= para_en + 3:
        zh_score += 1
    elif para_en >= para_zh + 3:
        en_score += 1

    left = _nearest_script(chars, idx, -1)
    right = _nearest_script(chars, idx, 1)
    if left == "zh":
        zh_score += 1
    elif left == "en":
        en_score += 1
    if right == "zh":
        zh_score += 1
    elif right == "en":
        en_score += 1

    if zh_score == en_score:
        return None
    if abs(zh_score - en_score) < min_confidence:
        return None
    return "zh" if zh_score > en_score else "en"


def _language_for_inner_text(text: str) -> str | None:
    zh, en = _count_lang_chars(text)
    if zh >= en + 1:
        return "zh"
    if en >= zh + 1:
        return "en"
    return None


def _is_numeric_pair(chars: list[str], idx: int) -> bool:
    return (
        0 < idx < len(chars) - 1
        and chars[idx - 1].isdigit()
        and chars[idx + 1].isdigit()
    )


def _convert_brackets_by_inner_language(
        chars: list[str],
        spans: list[tuple[int, int]]) -> None:
    stack: list[int] = []
    for i, ch in enumerate(chars):
        if ch in {"(", "（"}:
            if _in_spans(i, spans):
                continue
            stack.append(i)
            continue
        if ch not in {")", "）"}:
            continue
        if _in_spans(i, spans):
            continue
        if not stack:
            continue
        open_idx = stack.pop()
        inner = "".join(chars[open_idx + 1:i])
        lang = _language_for_inner_text(inner)
        if lang == "zh":
            chars[open_idx] = "（"
            chars[i] = "）"
        elif lang == "en":
            chars[open_idx] = "("
            chars[i] = ")"


def _convert_punctuation_by_context(
        chars: list[str],
        spans: list[tuple[int, int]],
        para_counts: tuple[int, int],
        min_confidence: int) -> None:
    for i, ch in enumerate(chars):
        if _in_spans(i, spans):
            continue

        if ch in {".", ",", ":"} and _is_numeric_pair(chars, i):
            continue

        if ch in _HALF_TO_FULL_PUNC:
            lang = _resolve_context_language(chars, i, para_counts, min_confidence)
            if lang == "zh":
                chars[i] = _HALF_TO_FULL_PUNC[ch]
            continue

        if ch in _FULL_TO_HALF_PUNC:
            lang = _resolve_context_language(chars, i, para_counts, min_confidence)
            if lang == "en":
                chars[i] = _FULL_TO_HALF_PUNC[ch]


def _convert_quotes_by_context(
        chars: list[str],
        spans: list[tuple[int, int]],
        para_counts: tuple[int, int],
        min_confidence: int) -> None:
    double_open = True
    single_open = True
    for i, ch in enumerate(chars):
        if _in_spans(i, spans):
            continue
        if ch not in {'"', "'", "“", "”", "‘", "’"}:
            continue

        if ch == "'" and 0 < i < len(chars) - 1:
            if chars[i - 1].isalnum() and chars[i + 1].isalnum():
                continue

        lang = _resolve_context_language(chars, i, para_counts, min_confidence)
        if lang == "en":
            if ch in {"“", "”", '"'}:
                chars[i] = '"'
            elif ch in {"‘", "’", "'"}:
                chars[i] = "'"
            continue
        if lang != "zh":
            continue

        if ch in {"“", "”", '"'}:
            chars[i] = "“" if double_open else "”"
            double_open = not double_open
        elif ch in {"‘", "’", "'"}:
            chars[i] = "‘" if single_open else "’"
            single_open = not single_open


def _smart_full_half_convert(
        text: str,
        *,
        para_counts: tuple[int, int],
        min_confidence: int,
        punctuation_by_context: bool,
        bracket_by_inner_language: bool,
        fullwidth_alnum_to_halfwidth: bool,
        quote_by_context: bool,
        protect_reference_numbering: bool) -> str:
    if not text:
        return text

    chars = list(text)
    spans = _build_protected_spans(
        text,
        protect_reference_numbering=protect_reference_numbering,
    )

    if fullwidth_alnum_to_halfwidth:
        for i, ch in enumerate(chars):
            if _in_spans(i, spans):
                continue
            chars[i] = _fullwidth_alnum_to_half(ch)

    if bracket_by_inner_language:
        _convert_brackets_by_inner_language(chars, spans)

    if punctuation_by_context:
        _convert_punctuation_by_context(chars, spans, para_counts, min_confidence)

    if quote_by_context:
        _convert_quotes_by_context(chars, spans, para_counts, min_confidence)

    return "".join(chars)


class WhitespaceNormalizeRule(BaseRule):
    name = "whitespace_normalize"
    description = "空白与全半角规范（实验）"

    def apply(self, doc: Document, config: SceneConfig,
              tracker: ChangeTracker, context: dict) -> None:
        ws_cfg = getattr(config, "whitespace_normalize", None)
        enabled = bool(getattr(ws_cfg, "enabled", True))
        normalize_space_variants = bool(
            getattr(ws_cfg, "normalize_space_variants", True)
        )
        convert_tabs = bool(getattr(ws_cfg, "convert_tabs", True))
        remove_zero_width = bool(getattr(ws_cfg, "remove_zero_width", True))
        collapse_multiple_spaces = bool(
            getattr(ws_cfg, "collapse_multiple_spaces", True)
        )
        trim_paragraph_edges = bool(
            getattr(ws_cfg, "trim_paragraph_edges", True)
        )
        smart_full_half_convert = bool(
            getattr(ws_cfg, "smart_full_half_convert", True)
        )
        punctuation_by_context = bool(
            getattr(ws_cfg, "punctuation_by_context", True)
        )
        bracket_by_inner_language = bool(
            getattr(ws_cfg, "bracket_by_inner_language", True)
        )
        fullwidth_alnum_to_halfwidth = bool(
            getattr(ws_cfg, "fullwidth_alnum_to_halfwidth", True)
        )
        quote_by_context = bool(
            getattr(ws_cfg, "quote_by_context", False)
        )
        protect_reference_numbering = bool(
            getattr(ws_cfg, "protect_reference_numbering", True)
        )
        try:
            context_min_confidence = int(
                getattr(ws_cfg, "context_min_confidence", 2)
            )
        except (TypeError, ValueError):
            context_min_confidence = 2
        context_min_confidence = max(1, min(context_min_confidence, 4))

        if not enabled:
            return
        if not any((
            normalize_space_variants,
            convert_tabs,
            remove_zero_width,
            collapse_multiple_spaces,
            trim_paragraph_edges,
            smart_full_half_convert,
        )):
            return

        target_indices = context.get("target_paragraph_indices")
        if target_indices is not None:
            target_indices = set(target_indices)

        changed_para_count = 0
        changed_text_nodes = 0

        for para_idx, para in enumerate(doc.paragraphs):
            if target_indices and para_idx not in target_indices:
                continue
            if _is_skipped_style(para):
                continue

            p_el = para._p
            if _has_complex_content(p_el):
                continue

            para_changed = False
            if convert_tabs:
                converted_tabs = _replace_tab_elements_with_spaces(p_el)
                if converted_tabs:
                    para_changed = True
                    changed_text_nodes += converted_tabs

            t_nodes = list(p_el.findall(f".//{_w('t')}"))
            if not t_nodes:
                continue

            # Phase 1: baseline whitespace normalization.
            for t_el in t_nodes:
                old = t_el.text or ""
                new = _normalize_inline_text(
                    old,
                    normalize_space_variants=normalize_space_variants,
                    convert_tabs=convert_tabs,
                    remove_zero_width=remove_zero_width,
                    collapse_multiple_spaces=collapse_multiple_spaces,
                )
                if new != old:
                    t_el.text = new
                    _sync_xml_space_attr(t_el)
                    para_changed = True
                    changed_text_nodes += 1

            # Phase 2: context-aware full/half-width conversion.
            if smart_full_half_convert and (
                    punctuation_by_context
                    or bracket_by_inner_language
                    or fullwidth_alnum_to_halfwidth
                    or quote_by_context):
                para_text = "".join((t.text or "") for t in t_nodes)
                para_counts = _count_lang_chars(para_text)
                para_new = _smart_full_half_convert(
                    para_text,
                    para_counts=para_counts,
                    min_confidence=context_min_confidence,
                    punctuation_by_context=punctuation_by_context,
                    bracket_by_inner_language=bracket_by_inner_language,
                    fullwidth_alnum_to_halfwidth=fullwidth_alnum_to_halfwidth,
                    quote_by_context=quote_by_context,
                    protect_reference_numbering=protect_reference_numbering,
                )
                if para_new != para_text:
                    node_lengths = [len(t_el.text or "") for t_el in t_nodes]
                    if sum(node_lengths) == len(para_new):
                        pos = 0
                        for t_el, seg_len in zip(t_nodes, node_lengths):
                            old = t_el.text or ""
                            new = para_new[pos:pos + seg_len]
                            pos += seg_len
                            if new != old:
                                t_el.text = new
                                _sync_xml_space_attr(t_el)
                                para_changed = True
                                changed_text_nodes += 1
                    else:
                        # Defensive fallback for future non-length-preserving rules.
                        for t_el in t_nodes:
                            old = t_el.text or ""
                            new = _smart_full_half_convert(
                                old,
                                para_counts=para_counts,
                                min_confidence=context_min_confidence,
                                punctuation_by_context=punctuation_by_context,
                                bracket_by_inner_language=bracket_by_inner_language,
                                fullwidth_alnum_to_halfwidth=fullwidth_alnum_to_halfwidth,
                                quote_by_context=quote_by_context,
                                protect_reference_numbering=protect_reference_numbering,
                            )
                            if new != old:
                                t_el.text = new
                                _sync_xml_space_attr(t_el)
                                para_changed = True
                                changed_text_nodes += 1

            # Phase 3: trim paragraph edges.
            if trim_paragraph_edges and _trim_paragraph_edges(t_nodes):
                para_changed = True

            if not para_changed:
                continue

            changed_para_count += 1
            tracker.record(
                rule_name=self.name,
                target=f"段落 #{para_idx}",
                section="body",
                change_type="text",
                before="空白/全半角未统一",
                after="已统一空白/全半角",
                paragraph_index=para_idx,
            )

        if changed_para_count:
            tracker.record(
                rule_name=self.name,
                target=f"{changed_para_count} 个段落",
                section="global",
                change_type="normalize",
                before="空白字符混用",
                after=f"已规范化（修改 {changed_text_nodes} 个文本节点）",
                paragraph_index=-1,
            )
