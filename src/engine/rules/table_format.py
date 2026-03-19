"""表格宽度自适应与字体格式规则"""

import re
from lxml import etree
from docx import Document
from docx.shared import Pt, Cm, Emu
from src.engine.rules.base import BaseRule
from src.engine.change_tracker import ChangeTracker
from src.formula_core.normalize import (
    looks_like_bibliographic_reference_text,
    looks_like_caption_text,
    looks_like_formula_text,
)
from src.scene.schema import SceneConfig
from src.utils.ooxml import apply_explicit_rfonts

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# A4 纸宽（twips）
_A4_WIDTH_TWIPS = 11906  # 21cm ≈ 11906 twips

# 每列最小宽度（twips），约 1.5cm
_MIN_COL_WIDTH = 850
# 公式表格在“右侧仅占位且无编号”时的最小压缩宽度（twips）
_MIN_EQ_PLACEHOLDER_COL_WIDTH = 220

# 单元格左右内边距合计（twips），Word 默认约 108*2
_CELL_PADDING = 216

# 字符宽度估算（twips）：10.5pt 字号下
_CJK_CHAR_W = 210    # 中文字符 ≈ 字号宽
_ASCII_CHAR_W = 115   # 英文字符 ≈ 字号 × 0.55


def _w(tag: str) -> str:
    return f"{{{_W_NS}}}{tag}"


_RE_EQ_NUM = re.compile(
    r'^\s*(?:[\.．…·•]{2,}\s*)?[\(（]\s*\d+(?:[\-\.]\d+)?\s*[\)）]\s*$'
)
_RE_EQ_NUM_SUFFIX = re.compile(
    r'^(?P<prefix>.*?)(?:[\.．…·•]{2,}\s*)?[\(（]\s*\d+(?:[\-\.]\d+)?\s*[\)）]\s*$'
)
_RE_EQ_NUM_TAIL = re.compile(r'[\(（]\s*(\d+(?:[\-\.]\d+)?)\s*[\)）]\s*$')
_RE_EQ_NUM_LOOSE = re.compile(
    r'^\s*(?:[\.．…·•]{2,}\s*)?(?:[\(（]\s*)?(\d+(?:[\-\.]\d+)?)\s*(?:[\)）])?\s*$'
)
_RE_EQ_ARROW = re.compile(r'(→|←|↔|⇌|->|<-|<->)')
_RE_HEADER_TRAILING_UNIT = re.compile(
    r'^(?P<head>.+?)\s*\((?P<unit>[^()]{1,80})\)\s*$'
)
_RE_TRAILING_PAREN_CHUNK = re.compile(
    r'^(?P<head>.+?)(?P<paren>[（(][^()（）]{1,80}[）)])\s*$'
)
_RE_CJK_TEXT = re.compile(r"[\u4e00-\u9fff]")
_RE_PROSE_PUNCT = re.compile(r"[，。；：！？、]")
_RE_COMPACT_FORMULA_PREFIX = re.compile(r"^[A-Za-zΑ-Ωα-ω0-9\s()+\-*/^_\\{}\[\].,]+$")
_RE_COMPACT_PLUS_TIMES_EXPR = re.compile(
    r"[A-Za-zΑ-Ωα-ω0-9\)\]}]\s*[+*]\s*[A-Za-zΑ-Ωα-ω0-9\(\[{\\]"
)
_RE_COMPACT_SLASH_EXPR = re.compile(
    r"[A-Za-zΑ-Ωα-ω0-9\)\]}]\s*/\s*[A-Za-zΑ-Ωα-ω0-9\(\[{\\]"
)
_RE_OMML_FORMULA_ANCHOR = re.compile(r"(=|≈|≠|≤|≥|∑|∫|∏|√|→|←|↔|⇌|±|∂)")
_MAX_EQ_TABLE_ROWS = 20


def _get_cell_text(tc) -> str:
    """提取单元格纯文本，保留 tab/换行等文本边界。"""
    parts = []
    for p in tc.findall(_w("p")):
        for r in p.findall(_w("r")):
            for child in r:
                local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                if local == "t" and child.text:
                    parts.append(child.text)
                elif local == "tab":
                    parts.append("\t")
                elif local in {"br", "cr"}:
                    parts.append("\n")
    return "".join(parts)


def _cell_has_math_content(tc) -> bool:
    """判断单元格是否包含 Word 数学对象（OMML）。"""
    if tc.findall(f".//{{{_M_NS}}}oMath"):
        return True
    if tc.findall(f".//{{{_M_NS}}}oMathPara"):
        return True
    return False


def _extract_cell_omml_linear_text(tc) -> str:
    return "".join(
        tc.xpath(".//*[namespace-uri()='%s' and local-name()='t']/text()" % _M_NS)
    ).strip()


def _cell_omml_is_prose_like(tc) -> bool:
    linear = _extract_cell_omml_linear_text(tc)
    if not linear:
        return False
    if looks_like_caption_text(linear):
        return True
    if looks_like_bibliographic_reference_text(linear):
        return True
    if len(linear) > 240:
        return True
    if _RE_OMML_FORMULA_ANCHOR.search(linear):
        return False

    chinese_count = len(_RE_CJK_TEXT.findall(linear))
    prose_punct_count = len(_RE_PROSE_PUNCT.findall(linear))
    matched, conf, _ = looks_like_formula_text(linear)
    if matched and float(conf) >= 0.78 and chinese_count <= max(2, int(len(linear) * 0.12)):
        return False

    if chinese_count >= max(6, int(len(linear) * 0.16)) and prose_punct_count >= 1:
        return True
    if chinese_count >= max(12, int(len(linear) * 0.28)):
        return True
    if any(
        marker in linear
        for marker in ("体积比为", "混合液中", "搅拌", "透析", "分别得到", "分别表示", "其中，")
    ):
        return True
    return False


def _cell_has_formula_math_anchor(tc) -> bool:
    if not _cell_has_math_content(tc):
        return False
    return not _cell_omml_is_prose_like(tc)


def _cell_has_equation_object(tc) -> bool:
    """判断单元格是否包含疑似公式 OLE 对象（MathType/Equation）。"""
    for obj in tc.findall(f".//{{{_W_NS}}}object"):
        try:
            raw = etree.tostring(obj, encoding="unicode").lower()
        except Exception:
            raw = ""
        if "equation" in raw or "mathtype" in raw:
            return True
    return False


def _cell_has_non_text_content(tc) -> bool:
    """判断单元格是否含公式/对象等非文本内容。"""
    if _cell_has_math_content(tc):
        return True
    if tc.findall(f".//{{{_W_NS}}}object"):
        return True
    if tc.findall(f".//{{{_W_NS}}}drawing"):
        return True
    if tc.findall(f".//{{{_W_NS}}}pict"):
        return True
    return False


def _table_has_drawing_or_object(tbl_el) -> bool:
    """Skip non-regular tables that embed pictures/objects (e.g. icon badges)."""
    for tc in tbl_el.iter(_w("tc")):
        if tc.findall(f".//{{{_W_NS}}}object"):
            return True
        if tc.findall(f".//{{{_W_NS}}}drawing"):
            return True
        if tc.findall(f".//{{{_W_NS}}}pict"):
            return True
    return False


def _cell_is_effectively_empty(tc) -> bool:
    """文本为空且不含公式/对象时，认为可写入序号。"""
    text = (_get_cell_text(tc) or "").strip()
    return not text and not _cell_has_non_text_content(tc)


def _iter_cell_paragraphs(tc):
    """递归遍历单元格内段落，兼容 sdt/ins 等包装层。"""
    return tc.iter(_w("p"))


def _looks_like_formula_prefix(prefix: str) -> bool:
    """判断编号前缀是否像公式表达式（覆盖“表达式+(2.8)”同格场景）。"""
    value = (prefix or "").strip()
    if not value:
        return False

    if looks_like_caption_text(value):
        return False

    chinese_count = len(_RE_CJK_TEXT.findall(value))
    if chinese_count:
        if _RE_PROSE_PUNCT.search(value):
            return False
        if chinese_count > max(2, int(len(value) * 0.18)):
            return False

    if re.search(r'(=|→|←|↔|->|<-|<->)', value):
        return True

    matched, conf, _ = looks_like_formula_text(value)

    if value.startswith("\\") and matched:
        return float(conf) >= 0.60

    if any(ch in value for ch in ("^", "_", "{", "}")) and matched:
        return float(conf) >= 0.60

    if not _RE_COMPACT_FORMULA_PREFIX.fullmatch(value):
        return False

    if _RE_COMPACT_PLUS_TIMES_EXPR.search(value):
        return True

    # Bare slash is too noisy for chemistry labels/solvent ratios such as CCl4/H2O.
    # Only accept slash-driven same-cell renumbering when grouping makes it look
    # like a compact fraction/expression rather than a sample or solvent name.
    if "/" in value and re.search(r"[(){}\[\]]", value):
        return bool(_RE_COMPACT_SLASH_EXPR.search(value))

    return False


def _looks_like_equation_text(text: str) -> bool:
    """判断文本是否像公式/反应式表达。"""
    s = (text or "").strip()
    if not s:
        return False
    if _RE_EQ_ARROW.search(s):
        return True
    if re.search(r'[A-Za-z0-9\)\]]\s*=\s*[A-Za-z0-9\(\[]', s):
        return True
    # 兜底：存在“空格+空格”连接的写法，且包含字母
    if re.search(r'\s\+\s', s) and re.search(r'[A-Za-zΑ-Ωα-ω]', s):
        return True
    return False


def _looks_like_formula_cell_text(text: str) -> bool:
    """Conservative text-only formula detector for equation-table heuristics."""
    s = (text or "").strip()
    if not s:
        return False
    if looks_like_caption_text(s):
        return False
    if _looks_like_equation_text(s):
        return True

    matched, conf, _ = looks_like_formula_text(s)
    if not matched:
        return False

    if s.startswith("\\"):
        return float(conf) >= 0.60

    if float(conf) >= 0.72:
        return True

    return False


def _analyze_text_only_equation_row(cell_texts: list[str]) -> dict[str, object]:
    last_idx = len(cell_texts) - 1
    equation_like_indices: list[int] = []
    other_text_indices: list[int] = []

    for idx, txt in enumerate(cell_texts):
        value = (txt or "").strip()
        if not value:
            continue
        if _extract_equation_number_id(value):
            continue
        if _looks_like_formula_cell_text(value):
            equation_like_indices.append(idx)
        else:
            other_text_indices.append(idx)

    only_trailing_other_text = bool(other_text_indices) and all(
        idx == last_idx for idx in other_text_indices
    )
    trailing_slot_available = False
    if last_idx >= 0:
        last_text = (cell_texts[last_idx] or "").strip()
        trailing_slot_available = (not last_text) or bool(_extract_equation_number_id(last_text))

    return {
        "equation_like_indices": equation_like_indices,
        "other_text_indices": other_text_indices,
        "only_trailing_other_text": only_trailing_other_text,
        "trailing_slot_available": trailing_slot_available,
        "is_formula_anchor_row": (
            equation_like_indices == [0]
            and (
                not other_text_indices
                or only_trailing_other_text
            )
        ),
    }


def _is_equation_table(tbl_el) -> bool:
    """Detect equation-table with conservative text-only heuristics."""
    rows = tbl_el.findall(_w("tr"))
    if not rows or len(rows) > _MAX_EQ_TABLE_ROWS:
        return False

    formula_anchor_rows = 0
    trailing_slot_rows = 0
    for tr in rows:
        row_cells = tr.findall(_w("tc"))
        if not row_cells:
            continue
        row_texts = [(_get_cell_text(tc) or "").strip() for tc in row_cells]
        has_math_anchor = any(
            _cell_has_formula_math_anchor(tc) or _cell_has_equation_object(tc)
            for tc in row_cells
        )
        has_pure_number_cell = any(
            txt and _RE_EQ_NUM.match(txt)
            for txt in row_texts
        )
        right_non_empty = next((txt for txt in reversed(row_texts) if txt), "")
        if right_non_empty:
            m = _RE_EQ_NUM_SUFFIX.match(right_non_empty)
            if m and _looks_like_formula_prefix(m.group("prefix")):
                return True
        profile = _analyze_text_only_equation_row(row_texts)
        row_has_formula_anchor = has_math_anchor or bool(
            profile.get("is_formula_anchor_row")
        )
        if row_has_formula_anchor and (
            profile.get("trailing_slot_available") or has_pure_number_cell
        ):
            return True
        if row_has_formula_anchor:
            formula_anchor_rows += 1
        if profile.get("trailing_slot_available") or has_pure_number_cell:
            trailing_slot_rows += 1

    if formula_anchor_rows >= 2 and trailing_slot_rows >= 1:
        return True

    return False


def _build_chapter_ranges(headings, body_end: int):
    """从标题列表构建章号范围表 [(start, end, chapter_num), ...]。"""
    chapters = []
    num = 0
    for h in headings:
        if getattr(h, "level", "") == "heading1":
            num += 1
            chapters.append((h.para_index, num))
    ranges = []
    for i, (start, n) in enumerate(chapters):
        end = chapters[i + 1][0] - 1 if i + 1 < len(chapters) else body_end
        ranges.append((start, end, n))
    return ranges


def _get_chapter_num(para_index: int, chapter_ranges) -> int:
    """查询段落所属章号，章前返回 0。"""
    for start, end, num in chapter_ranges:
        if start <= para_index <= end:
            return num
    return 0


def _extract_equation_number_id(text: str) -> str | None:
    """提取文本末尾的公式编号主体（如 3.2 / 2-8 / 3）。"""
    m = _RE_EQ_NUM_TAIL.search((text or "").strip())
    if not m:
        return None
    return m.group(1)


def _extract_equation_number_display(text: str) -> str | None:
    """Extract the tail number display and preserve original bracket shape."""
    m = _RE_EQ_NUM_TAIL.search((text or "").strip())
    if not m:
        return None
    # Keep full/half-width bracket style for width estimation; strip inner spaces.
    return re.sub(r"\s+", "", m.group(0))


def _parse_chapter_seq(number_id: str) -> tuple[int, int] | None:
    """将编号主体解析为 (chapter, seq)，如 3.2 / 3-2。"""
    parts = re.split(r'[\-\.]', number_id)
    if len(parts) != 2:
        return None
    if not parts[0].isdigit() or not parts[1].isdigit():
        return None
    return int(parts[0]), int(parts[1])


def _find_existing_equation_number_in_row(tr):
    """查找单行内已有公式编号，返回 (number_id, pure_number_cell_or_none, matched_cell_or_none)。"""
    cells = tr.findall(_w("tc"))
    for tc in reversed(cells):
        txt = (_get_cell_text(tc) or "").strip()
        if not txt:
            continue
        number_id = _extract_equation_number_id(txt)
        if not number_id:
            continue
        pure_cell = tc if (_RE_EQ_NUM.match(txt) and _cell_safe_for_plain_text_rewrite(tc)) else None
        return number_id, pure_cell, tc
    return None, None, None


def _find_existing_equation_number(tbl_el):
    """查找表内已有公式编号，返回 (number_id, pure_number_cell_or_none)。"""
    for tr in tbl_el.findall(_w("tr")):
        number_id, pure_cell, _ = _find_existing_equation_number_in_row(tr)
        if number_id:
            return number_id, pure_cell
    return None, None


def _is_equation_row(tr) -> bool:
    """判断当前行是否应按“公式行”参与编号。"""
    cells = tr.findall(_w("tc"))
    if not cells:
        return False

    # 含公式对象时直接视为公式行
    if any(_cell_has_formula_math_anchor(tc) or _cell_has_equation_object(tc) for tc in cells):
        return True

    cell_texts = [(_get_cell_text(tc) or "").strip() for tc in cells]
    non_empty = [txt for txt in cell_texts if txt]
    if not non_empty:
        return False

    # 行内已含编号时，视为公式行
    for txt in non_empty:
        if _extract_equation_number_id(txt):
            return True

    profile = _analyze_text_only_equation_row(cell_texts)
    if (
        profile.get("is_formula_anchor_row")
        and profile.get("trailing_slot_available")
    ):
        return True
    return False


def _insert_equation_number_cell_in_row(tr, number_text: str):
    """将编号写入当前行最右空单元格，返回写入的单元格或 None。"""
    cells = tr.findall(_w("tc"))
    if not cells:
        return None

    target = next((tc for tc in reversed(cells) if _cell_is_effectively_empty(tc)), None)
    if target is None:
        return None

    _set_cell_plain_text(target, number_text)
    return target


def _set_cell_plain_text(tc, text: str) -> None:
    """将单元格内容改为纯文本（首个 w:t 赋值，其余清空）。"""
    text_nodes = tc.findall(f".//{_w('t')}")
    if text_nodes:
        text_nodes[0].text = text
        for t in text_nodes[1:]:
            t.text = ""
        for r in tc.iter(_w("r")):
            for child in list(r):
                local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                if local in {"tab", "br", "cr"}:
                    r.remove(child)
        return
    paras = tc.findall(_w("p"))
    p = paras[0] if paras else etree.SubElement(tc, _w("p"))
    r = etree.SubElement(p, _w("r"))
    t = etree.SubElement(r, _w("t"))
    t.text = text


def _strip_equation_number_tail_from_cell(tc) -> bool:
    """Remove a same-cell trailing equation number while preserving non-text content."""
    text = _get_cell_text(tc)
    if not text:
        return False
    match = _RE_EQ_NUM_TAIL.search(text)
    if not match:
        return False
    prefix = text[:match.start()].rstrip()
    text_nodes = tc.findall(f".//{_w('t')}")
    if not text_nodes:
        return False
    text_nodes[0].text = prefix
    for t in text_nodes[1:]:
        t.text = ""
    for r in tc.iter(_w("r")):
        for child in list(r):
            local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if local in {"tab", "br", "cr"}:
                r.remove(child)
    return True


def _cell_has_rich_run_typography(tc) -> bool:
    """Detect rich run typography that should not be flattened.

    When a cell contains superscript/subscript (or mixed styled runs),
    rewriting with plain text would lose typography.
    """
    runs = list(tc.iter(_w("r")))
    signatures = []
    for r in runs:
        txt = "".join((t.text or "") for t in r.findall(_w("t")))
        if not txt:
            continue
        rPr = r.find(_w("rPr"))
        if rPr is None:
            signatures.append((False, False, "", False, False))
            continue

        def _onoff(local: str) -> bool:
            node = rPr.find(_w(local))
            if node is None:
                return False
            val = (node.get(_w("val")) or "").strip().lower()
            return val not in {"0", "false", "off", "none"}

        va = rPr.find(_w("vertAlign"))
        if va is not None:
            val = (va.get(_w("val")) or "").strip()
            if val in {"subscript", "superscript"}:
                return True

        pos = rPr.find(_w("position"))
        pos_val = (pos.get(_w("val")) or "").strip() if pos is not None else ""
        if pos_val not in {"", "0"}:
            return True

        underline = rPr.find(_w("u"))
        underline_val = ""
        if underline is not None:
            underline_val = (underline.get(_w("val")) or "").strip().lower() or "single"
            if underline_val == "none":
                underline_val = ""

        signatures.append(
            (
                _onoff("b") or _onoff("bCs"),
                _onoff("i") or _onoff("iCs"),
                underline_val,
                _onoff("strike") or _onoff("dstrike"),
                _onoff("caps") or _onoff("smallCaps"),
            )
        )
    return len(set(signatures)) > 1


def _cell_safe_for_plain_text_rewrite(tc) -> bool:
    """Whether it is safe to rewrite a cell by flattening to plain text."""
    if _cell_has_non_text_content(tc):
        return False
    if _cell_has_rich_run_typography(tc):
        return False
    return True


def _insert_equation_number_cell(tbl_el, number_text: str):
    """将编号写入首行最右空单元格，返回写入的单元格或 None。"""
    rows = tbl_el.findall(_w("tr"))
    if not rows:
        return None
    return _insert_equation_number_cell_in_row(rows[0], number_text)


def _rewrite_existing_equation_number_in_row(
        tr, number_text: str, pure_cell=None, number_cell=None):
    """重写单行已有编号为给定值，返回实际改写的单元格；失败返回 None。"""
    if pure_cell is not None and _cell_safe_for_plain_text_rewrite(pure_cell):
        _set_cell_plain_text(pure_cell, number_text)
        return pure_cell

    candidates = []
    if number_cell is not None:
        candidates.append(number_cell)
    for tc in reversed(tr.findall(_w("tc"))):
        if tc is number_cell:
            continue
        candidates.append(tc)

    for tc in candidates:
        if not _cell_safe_for_plain_text_rewrite(tc):
            # Keep equation objects / rich typography intact.
            continue
        txt = (_get_cell_text(tc) or "")
        if not txt:
            continue
        if _RE_EQ_NUM.match(txt.strip()):
            _set_cell_plain_text(tc, number_text)
            return tc
        if _RE_EQ_NUM_TAIL.search(txt.strip()):
            # 仅替换末尾编号，保留前缀表达式文本
            new_txt = _RE_EQ_NUM_TAIL.sub(number_text, txt.strip())
            _set_cell_plain_text(tc, new_txt)
            return tc
    return None


def _rewrite_existing_equation_number(tbl_el, number_text: str, pure_cell=None) -> bool:
    """重写表内已有编号为给定值，优先纯编号单元格；失败返回 False。"""
    for tr in tbl_el.findall(_w("tr")):
        rewritten = _rewrite_existing_equation_number_in_row(
            tr, number_text, pure_cell=pure_cell
        )
        if rewritten is not None:
            return True
    return False


def _cm_to_twips(cm: float) -> int:
    """厘米 → twips (1cm ≈ 567 twips)"""
    return int(cm * 567)


def _available_width(config: SceneConfig) -> int:
    """根据页面设置计算可用内容宽度（twips）"""
    ps = config.page_setup
    page_w = _A4_WIDTH_TWIPS
    left = _cm_to_twips(ps.margin.left_cm)
    right = _cm_to_twips(ps.margin.right_cm)
    gutter = _cm_to_twips(ps.gutter_cm)
    return page_w - left - right - gutter


def _count_cols_from_rows(tbl_el) -> int:
    """从表格行中推断列数（取第一行的单元格数，考虑 gridSpan）"""
    for tr in tbl_el.findall(_w("tr")):
        n = 0
        for tc in tr.findall(_w("tc")):
            span = 1
            tcPr = tc.find(_w("tcPr"))
            if tcPr is not None:
                gs = tcPr.find(_w("gridSpan"))
                if gs is not None:
                    span = int(gs.get(_w("val"), "1"))
            n += span
        if n > 0:
            return n
    return 0


def _get_grid_cols(tbl_el):
    """读取 tblGrid 中各列宽度，返回 [int, ...]（twips）"""
    grid = tbl_el.find(_w("tblGrid"))
    if grid is None:
        return []
    return [
        int(col.get(_w("w"), "0"))
        for col in grid.findall(_w("gridCol"))
    ]


def _set_grid_cols(tbl_el, widths: list[int]):
    """设置 tblGrid 各列宽度"""
    grid = tbl_el.find(_w("tblGrid"))
    if grid is None:
        grid = etree.SubElement(tbl_el, _w("tblGrid"))
    # 清除旧列
    for old in grid.findall(_w("gridCol")):
        grid.remove(old)
    for w in widths:
        col = etree.SubElement(grid, _w("gridCol"))
        col.set(_w("w"), str(w))


def top_level_table_anchor_positions(doc: Document) -> list[int]:
    """Map each top-level table to a nearby paragraph index.

    The default python-docx table order follows the top-level ``w:tbl`` order in
    ``document.xml``. For most tables, the nearest preceding paragraph is the
    most stable anchor. For leading tables that appear before the first
    paragraph (common on thesis covers), prefer the first following paragraph so
    they inherit the cover/front-matter scope instead of being forced onto
    paragraph ``0``.

    Returns ``-1`` only when the document has no paragraph to anchor to.
    """
    positions: list[int] = []
    pending_leading_tables = 0
    para_idx = -1

    for child in doc.element.body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "p":
            para_idx += 1
            if pending_leading_tables > 0:
                positions.extend([para_idx] * pending_leading_tables)
                pending_leading_tables = 0
        elif tag == "tbl":
            if para_idx >= 0:
                positions.append(para_idx)
            else:
                pending_leading_tables += 1

    if pending_leading_tables > 0:
        positions.extend([-1] * pending_leading_tables)

    return positions


def _analyze_col_weights(tbl_el, num_grid_cols: int) -> list[float]:
    """分析每个网格列的内容权重（基于最大文本长度）。

    对跨列单元格，将文本权重均分到所跨的各列。
    中文字符计为 2 个权重单位，英文/数字计为 1。
    """
    weights = [0.0] * num_grid_cols

    for tr in tbl_el.findall(_w("tr")):
        grid_idx = 0
        trPr = tr.find(_w("trPr"))
        if trPr is not None:
            gb = trPr.find(_w("gridBefore"))
            if gb is not None:
                grid_idx = int(gb.get(_w("val"), "0"))

        for tc in tr.findall(_w("tc")):
            # 读取 gridSpan
            span = 1
            tcPr = tc.find(_w("tcPr"))
            if tcPr is not None:
                gs = tcPr.find(_w("gridSpan"))
                if gs is not None:
                    span = int(gs.get(_w("val"), "1"))

            # 计算文本权重
            text = ""
            for p in tc.findall(_w("p")):
                for r in p.findall(_w("r")):
                    t = r.find(_w("t"))
                    if t is not None and t.text:
                        text += t.text
            tw = 0.0
            for ch in text:
                tw += 2.0 if ord(ch) > 0x7F else 1.0

            # 均分到所跨列
            per_col = tw / span if span > 0 else 0
            for k in range(span):
                ci = grid_idx + k
                if ci < num_grid_cols:
                    weights[ci] = max(weights[ci], per_col)

            grid_idx += span

    return weights


def _text_width_twips(text: str) -> int:
    """估算文本单行显示所需宽度（twips），基于 10.5pt 字号。"""
    w = 0
    for ch in text:
        w += _CJK_CHAR_W if ord(ch) > 0x7F else _ASCII_CHAR_W
    return w + _CELL_PADDING


def _first_row_min_widths(tbl_el, num_grid_cols: int) -> list[int]:
    """计算第一行各网格列的最小宽度（保证表头文字不换行）。

    跨列单元格：文本宽度均分到所跨列（合并后自然够宽）。
    """
    mins = [0] * num_grid_cols
    rows = tbl_el.findall(_w("tr"))
    if not rows:
        return mins

    tr = rows[0]
    grid_idx = 0
    trPr = tr.find(_w("trPr"))
    if trPr is not None:
        gb = trPr.find(_w("gridBefore"))
        if gb is not None:
            grid_idx = int(gb.get(_w("val"), "0"))

    for tc in tr.findall(_w("tc")):
        span = 1
        tcPr = tc.find(_w("tcPr"))
        if tcPr is not None:
            gs = tcPr.find(_w("gridSpan"))
            if gs is not None:
                span = int(gs.get(_w("val"), "1"))

        lines = _extract_cell_lines(tc)
        non_empty = [ln for ln in lines if ln.strip()]
        if not non_empty:
            need = 0
        elif len(non_empty) >= 2:
            # 首行若已是双行/多行（显式换行或多段），按“最长行”估算，
            # 避免把整段文本宽度都压到同一列导致该列过宽。
            need = max(_text_width_twips(ln) for ln in non_empty)
        else:
            need = _text_width_twips(non_empty[0])
        per_col = need // span if span > 0 else 0
        for k in range(span):
            ci = grid_idx + k
            if ci < num_grid_cols:
                mins[ci] = max(mins[ci], per_col)

        grid_idx += span

    return mins


def _non_first_row_min_widths(tbl_el, num_grid_cols: int) -> list[int]:
    """估算非首行（正文行）各网格列最小宽度，用于二次调宽。"""
    mins = [0] * num_grid_cols
    rows = tbl_el.findall(_w("tr"))
    if len(rows) <= 1:
        return mins

    for tr in rows[1:]:
        grid_idx = 0
        trPr = tr.find(_w("trPr"))
        if trPr is not None:
            gb = trPr.find(_w("gridBefore"))
            if gb is not None:
                grid_idx = int(gb.get(_w("val"), "0"))

        for tc in tr.findall(_w("tc")):
            span = 1
            tcPr = tc.find(_w("tcPr"))
            if tcPr is not None:
                gs = tcPr.find(_w("gridSpan"))
                if gs is not None:
                    span = int(gs.get(_w("val"), "1"))

            lines = _extract_cell_lines(tc)
            non_empty = [ln for ln in lines if ln.strip()]
            need = max((_text_width_twips(ln) for ln in non_empty), default=0)
            per_col = need // span if span > 0 else 0
            for k in range(span):
                ci = grid_idx + k
                if ci < num_grid_cols:
                    mins[ci] = max(mins[ci], per_col)

            grid_idx += span

    return mins


def _distribute_widths(weights: list[float], total: int,
                       min_widths: list[int] | None = None) -> list[int]:
    """按权重分配列宽，保证最小宽度和总宽精确匹配。

    算法：先给每列分配 floor 宽度，再将剩余空间按权重分配。
    这样 floor 约束永远不会被后续归一化破坏。
    """
    n = len(weights)
    if n == 0:
        return []

    for i in range(n):
        if weights[i] < 1.0:
            weights[i] = 1.0

    # 合并最小宽度约束
    floor = [_MIN_COL_WIDTH] * n
    if min_widths:
        for i in range(n):
            floor[i] = max(floor[i], min_widths[i])

    floor_sum = sum(floor)

    # floor 总和超过 total 时按比例缩放（空间不足，无法保障）
    if floor_sum >= total:
        scale = total / floor_sum
        result = [int(f * scale) for f in floor]
        diff = total - sum(result)
        for i in range(abs(diff)):
            result[i % n] += 1 if diff > 0 else -1
        return result

    # 空间充足：先分配 floor，再按权重分配剩余空间
    remaining = total - floor_sum
    total_weight = sum(weights)
    result = [floor[i] + remaining * weights[i] / total_weight
              for i in range(n)]

    # 转整数并修正舍入误差
    result = [int(r) for r in result]
    diff = total - sum(result)
    for i in range(abs(diff)):
        result[i % n] += 1 if diff > 0 else -1

    return result


def _extract_cell_lines(tc) -> list[str]:
    """提取单元格文本行（按段落与 <w:br/> 断行）。"""
    lines = []
    for p in _iter_cell_paragraphs(tc):
        curr = []
        for node in p.iter():
            tag = node.tag.split("}")[-1] if "}" in node.tag else node.tag
            if tag == "t" and node.text:
                parts = node.text.replace("\r", "\n").split("\n")
                for i, part in enumerate(parts):
                    if i:
                        lines.append("".join(curr))
                        curr = []
                    curr.append(part)
            elif tag == "br":
                lines.append("".join(curr))
                curr = []
        lines.append("".join(curr))
    return lines if lines else [(_get_cell_text(tc) or "")]


def _header_text_with_unit_break(text: str) -> str | None:
    """Try converting 'Title (unit)' to two lines:
    'Title\\n(unit)'. Returns None if not suitable.
    """
    s = (text or "").strip()
    if not s:
        return None
    m = _RE_HEADER_TRAILING_UNIT.match(s)
    if not m:
        return None
    head = m.group("head").strip()
    unit = m.group("unit").strip()
    if not head or not unit:
        return None
    # Restrict to latin-heavy scientific headers to avoid touching common CN headers.
    if not re.search(r"[A-Za-z]", head):
        return None
    if len(head) < 8:
        return None
    return f"{head}\n({unit})"


def _normalize_first_row_unit_breaks(tbl_el, widths: list[int]) -> int:
    """When first-row english header is too long, enforce line break before unit '(...)'."""
    if not widths:
        return 0
    rows = tbl_el.findall(_w("tr"))
    if not rows:
        return 0

    tr = rows[0]
    grid_idx = 0
    trPr = tr.find(_w("trPr"))
    if trPr is not None:
        gb = trPr.find(_w("gridBefore"))
        if gb is not None:
            grid_idx = int(gb.get(_w("val"), "0"))

    changed = 0
    for tc in tr.findall(_w("tc")):
        span = 1
        tcPr = tc.find(_w("tcPr"))
        if tcPr is not None:
            gs = tcPr.find(_w("gridSpan"))
            if gs is not None:
                span = int(gs.get(_w("val"), "1"))

        start = grid_idx
        end = min(len(widths), start + span)
        cell_w = sum(widths[start:end]) if start < len(widths) and start < end else 0

        lines = [ln for ln in _extract_cell_lines(tc) if ln.strip()]
        if len(lines) >= 2:
            grid_idx += span
            continue

        text = (lines[0] if lines else (_get_cell_text(tc) or "")).strip()
        rewritten = _header_text_with_unit_break(text)
        if rewritten is None:
            grid_idx += span
            continue

        # Keep superscript/subscript and mixed run styling intact.
        if _cell_has_rich_run_typography(tc):
            grid_idx += span
            continue

        # Only enforce break when the one-line text clearly exceeds assigned width.
        if cell_w > 0 and _text_width_twips(text) <= int(cell_w * 1.03):
            grid_idx += span
            continue

        _set_cell_plain_text(tc, rewritten)
        changed += 1
        grid_idx += span

    return changed


def _body_text_with_paren_break(text: str) -> str | None:
    """Try converting 'name(... )' to two lines:
    'name\\n(...)'. Returns None if not suitable.
    """
    s = (text or "").strip()
    if not s:
        return None
    m = _RE_TRAILING_PAREN_CHUNK.match(s)
    if not m:
        return None
    head = m.group("head").strip()
    paren = m.group("paren").strip()
    if not head or not paren:
        return None
    if len(head) < 6:
        return None
    return f"{head}\n{paren}"


def _normalize_body_paren_breaks(tbl_el, widths: list[int]) -> int:
    """Normalize non-header long body cells with trailing parenthesis to two lines.

    This helps avoid one column over-expanding and squeezing other columns
    (e.g., manufacturer column being wrapped unexpectedly).
    """
    if not widths:
        return 0
    rows = tbl_el.findall(_w("tr"))
    if len(rows) <= 1:
        return 0

    changed = 0
    for tr in rows[1:]:
        grid_idx = 0
        trPr = tr.find(_w("trPr"))
        if trPr is not None:
            gb = trPr.find(_w("gridBefore"))
            if gb is not None:
                grid_idx = int(gb.get(_w("val"), "0"))

        for tc in tr.findall(_w("tc")):
            span = 1
            tcPr = tc.find(_w("tcPr"))
            if tcPr is not None:
                gs = tcPr.find(_w("gridSpan"))
                if gs is not None:
                    span = int(gs.get(_w("val"), "1"))

            start = grid_idx
            end = min(len(widths), start + span)
            cell_w = sum(widths[start:end]) if start < len(widths) and start < end else 0

            lines = [ln for ln in _extract_cell_lines(tc) if ln.strip()]
            if len(lines) >= 2:
                grid_idx += span
                continue
            text = (lines[0] if lines else (_get_cell_text(tc) or "")).strip()
            rewritten = _body_text_with_paren_break(text)
            if rewritten is None:
                grid_idx += span
                continue

            # Keep superscript/subscript and mixed run styling intact.
            if _cell_has_rich_run_typography(tc):
                grid_idx += span
                continue

            # Proactive: if close to edge (>= 90% width), normalize break.
            if cell_w > 0 and _text_width_twips(text) < int(cell_w * 0.90):
                grid_idx += span
                continue

            _set_cell_plain_text(tc, rewritten)
            changed += 1
            grid_idx += span

    return changed


def _row_spans(tr) -> list[int]:
    spans = []
    for tc in tr.findall(_w("tc")):
        span = 1
        tcPr = tc.find(_w("tcPr"))
        if tcPr is not None:
            gs = tcPr.find(_w("gridSpan"))
            if gs is not None:
                span = int(gs.get(_w("val"), "1"))
        spans.append(span)
    return spans


def _set_tc_grid_span(tc, span: int) -> None:
    tcPr = tc.find(_w("tcPr"))
    if tcPr is None:
        tcPr = etree.SubElement(tc, _w("tcPr"))
        tc.insert(0, tcPr)
    gs = tcPr.find(_w("gridSpan"))
    if span <= 1:
        if gs is not None:
            tcPr.remove(gs)
        return
    if gs is None:
        gs = etree.SubElement(tcPr, _w("gridSpan"))
    gs.set(_w("val"), str(span))


def _normalize_first_row_span_pattern(tbl_el) -> bool:
    """修正“首行跨列模式与后续主模式错位”的情况，避免首行与正文对不齐。"""
    rows = tbl_el.findall(_w("tr"))
    if len(rows) < 2:
        return False

    first = rows[0]
    first_spans = _row_spans(first)
    if not first_spans:
        return False
    first_total = sum(first_spans)
    first_len = len(first_spans)

    freq = {}
    for tr in rows[1:]:
        spans = _row_spans(tr)
        if len(spans) != first_len:
            continue
        if sum(spans) != first_total:
            continue
        key = tuple(spans)
        freq[key] = freq.get(key, 0) + 1
    if not freq:
        return False
    ref = list(max(freq.items(), key=lambda kv: kv[1])[0])
    if ref == first_spans:
        return False

    diff = [i for i, (a, b) in enumerate(zip(first_spans, ref)) if a != b]
    # 保守修复：仅处理相邻两列的 1/N 互换错位（本次用户样例即此类）
    if len(diff) != 2 or diff[1] != diff[0] + 1:
        return False
    i, j = diff
    if not (first_spans[i] == ref[j] and first_spans[j] == ref[i]):
        return False
    if 1 not in (first_spans[i], first_spans[j]):
        return False

    cells = first.findall(_w("tc"))
    _set_tc_grid_span(cells[i], ref[i])
    _set_tc_grid_span(cells[j], ref[j])
    return True


def _first_row_cell_specs(tbl_el):
    """返回首行单元格规格：[{'start','span','need','multiline'}, ...]。"""
    rows = tbl_el.findall(_w("tr"))
    if not rows:
        return []
    tr = rows[0]
    specs = []
    grid_idx = 0
    trPr = tr.find(_w("trPr"))
    if trPr is not None:
        gb = trPr.find(_w("gridBefore"))
        if gb is not None:
            grid_idx = int(gb.get(_w("val"), "0"))

    for tc in tr.findall(_w("tc")):
        span = 1
        tcPr = tc.find(_w("tcPr"))
        if tcPr is not None:
            gs = tcPr.find(_w("gridSpan"))
            if gs is not None:
                span = int(gs.get(_w("val"), "1"))

        lines = _extract_cell_lines(tc)
        non_empty = [ln for ln in lines if ln.strip()]
        multiline = len(non_empty) >= 2
        if not non_empty:
            need = 0
        elif multiline:
            need = max(_text_width_twips(ln) for ln in non_empty)
        else:
            need = _text_width_twips(non_empty[0])

        specs.append({
            "start": grid_idx,
            "span": span,
            "need": need,
            "multiline": multiline,
            "line_count": len(non_empty),
        })
        grid_idx += span
    return specs


def _second_pass_rebalance_first_row(
        tbl_el,
        widths: list[int],
        weights: list[float],
        body_min_widths: list[int] | None = None) -> list[int]:
    """二次列宽调整：
    1) 首行多行/长标题单元格可适当收窄；
    2) 优先给正文行仍不足宽的列补宽（如 Light 列）。
    """
    if not widths:
        return widths

    specs = _first_row_cell_specs(tbl_el)
    if not specs:
        return widths

    n = len(widths)
    donor_cap = [0] * n
    recv_need = [0] * n
    multiline_mask = [False] * n

    if body_min_widths:
        for i in range(min(n, len(body_min_widths))):
            deficit = body_min_widths[i] - widths[i]
            if deficit > 40:
                recv_need[i] = deficit
    total_recv = sum(recv_need)

    # 收集“可释放列宽”和“需要补宽列宽”
    for spec in specs:
        start = spec["start"]
        span = spec["span"]
        end = min(n, start + span)
        if start >= n or start >= end:
            continue
        actual = sum(widths[start:end])
        need = spec["need"]

        if spec["multiline"]:
            for ci in range(start, end):
                multiline_mask[ci] = True
            keep = max(need, _MIN_COL_WIDTH * span)
            surplus = actual - keep
            if surplus > 40:
                # 将可释放宽度按当前列宽比例摊到所跨列
                block_sum = sum(widths[start:end]) or 1
                allocated = 0
                for ci in range(start, end):
                    cap = int(surplus * widths[ci] / block_sum)
                    room_base = max(0, widths[ci] - _MIN_COL_WIDTH)
                    if recv_need[ci] > 0:
                        room_base = 0
                    cap = min(cap, room_base)
                    donor_cap[ci] += cap
                    allocated += cap
                rem = surplus - allocated
                for ci in range(start, end):
                    if rem <= 0:
                        break
                    room = max(0, widths[ci] - _MIN_COL_WIDTH - donor_cap[ci])
                    if recv_need[ci] > 0:
                        room = 0
                    if room <= 0:
                        continue
                    d = min(room, rem)
                    donor_cap[ci] += d
                    rem -= d
        elif total_recv > 0:
            # 单行长标题：允许按“两行目标宽度”收窄，为正文长内容列腾空间。
            keep_2line = _CELL_PADDING + int(max(0, need - _CELL_PADDING) / 2) + 60
            keep = max(_MIN_COL_WIDTH * span, keep_2line)
            surplus = actual - keep
            if surplus > 80:
                block_sum = sum(widths[start:end]) or 1
                allocated = 0
                for ci in range(start, end):
                    cap = int(surplus * widths[ci] / block_sum)
                    room_base = max(0, widths[ci] - _MIN_COL_WIDTH)
                    if recv_need[ci] > 0:
                        room_base = 0
                    cap = min(cap, room_base)
                    donor_cap[ci] += cap
                    allocated += cap
                rem = surplus - allocated
                for ci in range(start, end):
                    if rem <= 0:
                        break
                    room = max(0, widths[ci] - _MIN_COL_WIDTH - donor_cap[ci])
                    if recv_need[ci] > 0:
                        room = 0
                    if room <= 0:
                        continue
                    d = min(room, rem)
                    donor_cap[ci] += d
                    rem -= d
        else:
            deficit = need - actual
            if deficit > 40:
                per = deficit // span
                rem = deficit - per * span
                for k, ci in enumerate(range(start, end)):
                    recv_need[ci] += per + (1 if k < rem else 0)

    total_donor = sum(donor_cap)
    if total_donor <= 0:
        return widths

    total_recv = sum(recv_need)
    if total_recv <= 0:
        # 若无明显缺口，把释放宽度分给非多行列（按权重）
        recv_weights = [0.0 if multiline_mask[i] else max(1.0, weights[i]) for i in range(n)]
        sw = sum(recv_weights)
        if sw <= 0:
            return widths
        recv_need = [int(total_donor * w / sw) for w in recv_weights]
        rem = total_donor - sum(recv_need)
        order = sorted(range(n), key=lambda i: recv_weights[i], reverse=True)
        for i in order:
            if rem <= 0:
                break
            if recv_weights[i] <= 0:
                continue
            recv_need[i] += 1
            rem -= 1
        total_recv = sum(recv_need)

    move = min(total_donor, total_recv)
    if move <= 0:
        return widths

    # 贪心转移：从可释放最多列转到最缺列，保持总宽不变
    donor_rem = donor_cap[:]
    recv_rem = recv_need[:]
    while move > 0:
        recv_candidates = [i for i in range(n) if recv_rem[i] > 0]
        donor_candidates = [i for i in range(n) if donor_rem[i] > 0 and recv_rem[i] <= 0]
        if not recv_candidates or not donor_candidates:
            break
        r_idx = max(recv_candidates, key=lambda i: recv_rem[i])
        donor_candidates = [i for i in donor_candidates if i != r_idx]
        if not donor_candidates:
            break
        d_idx = max(donor_candidates, key=lambda i: donor_rem[i])
        step = min(move, donor_rem[d_idx], recv_rem[r_idx], 20)
        if step <= 0:
            break
        if widths[d_idx] - step < _MIN_COL_WIDTH:
            donor_rem[d_idx] = max(0, widths[d_idx] - _MIN_COL_WIDTH)
            continue
        widths[d_idx] -= step
        widths[r_idx] += step
        donor_rem[d_idx] -= step
        recv_rem[r_idx] -= step
        move -= step

    return widths


def _set_table_width(tbl_el, total_w: int):
    """设置 tblPr/tblW 为精确宽度（dxa）"""
    tblPr = tbl_el.find(_w("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(tbl_el, _w("tblPr"))
        tbl_el.insert(0, tblPr)
    tblW = tblPr.find(_w("tblW"))
    if tblW is None:
        tblW = etree.SubElement(tblPr, _w("tblW"))
    tblW.set(_w("w"), str(total_w))
    tblW.set(_w("type"), "dxa")


def _set_table_alignment(tbl_el, align: str):
    """设置表格水平对齐。align: left/center/right"""
    tblPr = tbl_el.find(_w("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(tbl_el, _w("tblPr"))
        tbl_el.insert(0, tblPr)
    jc = tblPr.find(_w("jc"))
    if jc is None:
        jc = etree.SubElement(tblPr, _w("jc"))
    jc.set(_w("val"), align)


def _normalize_horizontal_alignment(value: str | None, *, default: str = "center") -> str:
    raw = str(value or "").strip().lower()
    return raw if raw in {"left", "center", "right", "justify"} else default


def _target_normal_table_width(
        avail_w: int,
        num_cols: int,
        first_min_ws: list[int],
        body_min_ws: list[int]) -> int:
    """估算常规表格“适宜压缩”目标宽度。"""
    if avail_w <= 0 or num_cols <= 0:
        return avail_w

    mins = []
    for i in range(num_cols):
        need = _MIN_COL_WIDTH
        if i < len(first_min_ws):
            need = max(need, first_min_ws[i])
        if i < len(body_min_ws):
            need = max(need, body_min_ws[i])
        mins.append(need)
    content_total = sum(mins)

    # 内容本身已接近满宽时，不压缩。
    if content_total >= int(avail_w * 0.9):
        return avail_w

    # 轻量缓冲，避免过度贴边；并限制压缩下限。
    buffer_w = min(120 * num_cols, 600)
    target = content_total + buffer_w
    target = max(target, int(avail_w * 0.48))
    target = min(target, int(avail_w * 0.92))
    target = max(target, content_total)
    target = min(target, avail_w)
    return target


def _build_smart_width_palette(raw_widths: list[int], avail_w: int, max_levels: int = 4) -> list[int]:
    """Build a small set of representative table widths from raw compact targets."""
    if not raw_widths or avail_w <= 0:
        return []

    vals = [
        max(_MIN_COL_WIDTH, min(avail_w, int(w)))
        for w in raw_widths
        if int(w) > 0
    ]
    if not vals:
        return []
    vals.sort()

    levels = max(1, min(max_levels, len(vals)))
    if levels == 1:
        return [vals[0]]

    # Evenly sample order statistics: shortest, longest, and middle anchors.
    anchors = []
    for i in range(levels):
        idx = int(round(i * (len(vals) - 1) / (levels - 1)))
        anchors.append(vals[idx])

    # Merge overly close neighboring anchors to reduce width categories.
    tol = max(120, int(avail_w * 0.03))
    palette = [anchors[0]]
    for w in anchors[1:-1]:
        if abs(w - palette[-1]) > tol:
            palette.append(w)
    if abs(anchors[-1] - palette[-1]) > tol:
        palette.append(anchors[-1])
    else:
        palette[-1] = anchors[-1]
    return palette


def _nearest_palette_width(raw_w: int, palette: list[int]) -> int:
    """Map raw width to representative width with slight upward preference.

    Smart mode should trend wider than compact mode, but must not force
    most tables to the maximum width bucket.
    """
    if not palette:
        return raw_w

    vals = sorted(int(w) for w in palette if int(w) > 0)
    if not vals:
        return raw_w
    if raw_w <= vals[0]:
        return vals[0]

    # Penalize downward mapping so we still prefer wider buckets when distances
    # are close, while avoiding large jumps to the widest bucket.
    shrink_penalty = 1.25
    for i in range(1, len(vals)):
        hi = vals[i]
        if raw_w > hi:
            continue
        lo = vals[i - 1]
        if hi <= lo:
            return hi
        up_gap = hi - raw_w
        down_gap = raw_w - lo
        return hi if up_gap <= down_gap * shrink_penalty else lo

    # Raw width is above the current palette max (rare): use the largest one.
    return vals[-1]


def _build_smart_width_plan(
        raw_targets: list[tuple[int, int]],
        avail_w: int,
        max_levels: int = 4) -> dict[int, int]:
    """Plan per-table target widths with a limited number of representative widths."""
    if not raw_targets:
        return {}
    palette = _build_smart_width_palette(
        [w for _, w in raw_targets], avail_w, max_levels=max_levels
    )
    return {idx: _nearest_palette_width(raw_w, palette) for idx, raw_w in raw_targets}


def _update_cell_widths(tbl_el, new_widths: list[int]):
    """更新每行中各单元格的 tcW，跨列单元格取所跨列宽之和。"""
    for tr in tbl_el.findall(_w("tr")):
        # 处理 gridBefore（行首空网格列偏移）
        grid_idx = 0
        trPr = tr.find(_w("trPr"))
        if trPr is not None:
            gb = trPr.find(_w("gridBefore"))
            if gb is not None:
                grid_idx = int(gb.get(_w("val"), "0"))

        for tc in tr.findall(_w("tc")):
            span = 1
            tcPr = tc.find(_w("tcPr"))
            if tcPr is not None:
                gs = tcPr.find(_w("gridSpan"))
                if gs is not None:
                    span = int(gs.get(_w("val"), "1"))

            # 计算该单元格应有的宽度
            cell_w = 0
            for k in range(span):
                ci = grid_idx + k
                if ci < len(new_widths):
                    cell_w += new_widths[ci]

            # 设置 tcW
            if tcPr is None:
                tcPr = etree.SubElement(tc, _w("tcPr"))
                tc.insert(0, tcPr)
            tcW = tcPr.find(_w("tcW"))
            if tcW is None:
                tcW = etree.SubElement(tcPr, _w("tcW"))
            tcW.set(_w("w"), str(cell_w))
            tcW.set(_w("type"), "dxa")

            grid_idx += span


def _clear_row_width_exceptions(tbl_el):
    """清理行级表格宽度例外，避免跨页后首行/后续行列宽不一致。"""
    for tr in tbl_el.findall(_w("tr")):
        trPr = tr.find(_w("trPr"))
        if trPr is None:
            continue
        for tblPrEx in trPr.findall(_w("tblPrEx")):
            for tag in ("tblW", "tblInd", "tblCellSpacing"):
                ex = tblPrEx.find(_w(tag))
                if ex is not None:
                    tblPrEx.remove(ex)


def _format_cell_fonts(tbl_el, font_cn, font_en, size_pt):
    """设置表格内所有 run 的字体（5号宋体/TNR）。"""
    size_half_pt = str(int(size_pt * 2))
    for tc in tbl_el.iter(_w("tc")):
        for r in tc.iter(_w("r")):
            rPr = r.find(_w("rPr"))
            if rPr is None:
                rPr = etree.SubElement(r, _w("rPr"))
                r.insert(0, rPr)
            rf = rPr.find(_w("rFonts"))
            if rf is None:
                rf = etree.SubElement(rPr, _w("rFonts"))
            apply_explicit_rfonts(
                rPr,
                font_cn=font_cn,
                font_en=font_en,
            )
            rf.set(_w("ascii"), font_en)
            rf.set(_w("hAnsi"), font_en)
            rf.set(_w("eastAsia"), font_cn)
            for attr in ("asciiTheme", "hAnsiTheme",
                         "eastAsiaTheme", "cstheme"):
                rf.attrib.pop(_w(attr), None)
            sz = rPr.find(_w("sz"))
            if sz is None:
                sz = etree.SubElement(rPr, _w("sz"))
            sz.set(_w("val"), size_half_pt)
            szCs = rPr.find(_w("szCs"))
            if szCs is None:
                szCs = etree.SubElement(rPr, _w("szCs"))
            szCs.set(_w("val"), size_half_pt)


def _center_all_cells(tbl_el):
    """设置所有单元格水平居中 + 垂直居中。"""
    for tc in tbl_el.iter(_w("tc")):
        # 垂直居中：tcPr/vAlign
        tcPr = tc.find(_w("tcPr"))
        if tcPr is None:
            tcPr = etree.SubElement(tc, _w("tcPr"))
            tc.insert(0, tcPr)
        vAlign = tcPr.find(_w("vAlign"))
        if vAlign is None:
            vAlign = etree.SubElement(tcPr, _w("vAlign"))
        vAlign.set(_w("val"), "center")

        # 水平居中：每个段落的 jc
        for p in _iter_cell_paragraphs(tc):
            pPr = p.find(_w("pPr"))
            if pPr is None:
                pPr = etree.SubElement(p, _w("pPr"))
                p.insert(0, pPr)
            jc = pPr.find(_w("jc"))
            if jc is None:
                jc = etree.SubElement(pPr, _w("jc"))
            jc.set(_w("val"), "center")
            # Clear paragraph indents so "center" is visually centered.
            ind = pPr.find(_w("ind"))
            if ind is None:
                ind = etree.SubElement(pPr, _w("ind"))
            for attr in (
                "left", "leftChars", "right", "rightChars",
                "firstLine", "firstLineChars", "hanging", "hangingChars"
            ):
                ind.set(_w(attr), "0")


def _set_table_borders(tbl_el, size_eighth_pt: int = 4):
    """设置表格全框线（外边框 + 内部网格线）。

    size_eighth_pt: 线宽，单位 1/8 磅。4 = 0.5磅，8 = 1磅。
    """
    tblPr = tbl_el.find(_w("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(tbl_el, _w("tblPr"))
        tbl_el.insert(0, tblPr)

    # 移除旧边框定义
    old = tblPr.find(_w("tblBorders"))
    if old is not None:
        tblPr.remove(old)

    borders = etree.SubElement(tblPr, _w("tblBorders"))
    sz = str(size_eighth_pt)
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = etree.SubElement(borders, _w(side))
        b.set(_w("val"), "single")
        b.set(_w("sz"), sz)
        b.set(_w("space"), "0")
        b.set(_w("color"), "000000")


def _clear_cell_border_overrides(tbl_el) -> None:
    """Remove all cell-level border overrides so table-level borders are authoritative."""
    for tc in tbl_el.iter(_w("tc")):
        tcPr = tc.find(_w("tcPr"))
        if tcPr is None:
            continue
        tc_borders = tcPr.find(_w("tcBorders"))
        if tc_borders is not None:
            tcPr.remove(tc_borders)


def _set_tc_side_border(tc, side: str, val: str, sz: int, color: str = "000000") -> None:
    """Set one border side for a table cell."""
    tcPr = tc.find(_w("tcPr"))
    if tcPr is None:
        tcPr = etree.SubElement(tc, _w("tcPr"))
        tc.insert(0, tcPr)
    tc_borders = tcPr.find(_w("tcBorders"))
    if tc_borders is None:
        tc_borders = etree.SubElement(tcPr, _w("tcBorders"))
    border = tc_borders.find(_w(side))
    if border is None:
        border = etree.SubElement(tc_borders, _w(side))
    border.set(_w("val"), val)
    border.set(_w("sz"), str(max(0, int(sz))))
    border.set(_w("space"), "0")
    border.set(_w("color"), color)


def _set_table_border_side(tbl_borders, side: str, val: str, sz: int, color: str = "000000") -> None:
    border = tbl_borders.find(_w(side))
    if border is None:
        border = etree.SubElement(tbl_borders, _w(side))
    border.set(_w("val"), val)
    border.set(_w("sz"), str(max(0, int(sz))))
    border.set(_w("space"), "0")
    border.set(_w("color"), color)


def _set_three_line_borders(tbl_el, header_sz_eighth_pt: int = 8, bottom_sz_eighth_pt: int = 4) -> None:
    """Apply strict three-line style:
    - First row: top + bottom (thick)
    - Last row: bottom (thin)
    - No other borders
    """
    tblPr = tbl_el.find(_w("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(tbl_el, _w("tblPr"))
        tbl_el.insert(0, tblPr)
    # Remove table style inheritance (e.g., Table Grid) to avoid style borders reappearing.
    old_style = tblPr.find(_w("tblStyle"))
    if old_style is not None:
        tblPr.remove(old_style)

    old = tblPr.find(_w("tblBorders"))
    if old is not None:
        tblPr.remove(old)
    tbl_borders = etree.SubElement(tblPr, _w("tblBorders"))

    # Strict mode: no table-level visible line, all visible lines are row/cell-level.
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        _set_table_border_side(tbl_borders, side, "none", 0)

    # Force explicit none for every cell border first (do not rely on style fallback).
    for tc in tbl_el.iter(_w("tc")):
        tcPr = tc.find(_w("tcPr"))
        if tcPr is None:
            tcPr = etree.SubElement(tc, _w("tcPr"))
            tc.insert(0, tcPr)
        _set_tc_borders_none(tcPr)
    rows = tbl_el.findall(_w("tr"))
    if not rows:
        return

    # First row: top + bottom thick.
    for tc in rows[0].findall(_w("tc")):
        _set_tc_side_border(tc, "top", "single", header_sz_eighth_pt)
        _set_tc_side_border(tc, "bottom", "single", header_sz_eighth_pt)

    # Last row: bottom thin.
    last_cells = rows[-1].findall(_w("tc"))
    for tc in last_cells:
        _set_tc_side_border(tc, "bottom", "single", bottom_sz_eighth_pt)


def _format_cell_paragraphs(tbl_el, line_spacing_mode: str = "single"):
    """设置表格内所有段落行距与缩进。"""
    mode = (line_spacing_mode or "single").strip().lower()
    line_map = {
        "single": "240",
        "one_half": "360",
        "double": "480",
    }
    line_val = line_map.get(mode, "240")
    for tc in tbl_el.iter(_w("tc")):
        for p in _iter_cell_paragraphs(tc):
            pPr = p.find(_w("pPr"))
            if pPr is None:
                pPr = etree.SubElement(p, _w("pPr"))
                p.insert(0, pPr)
            # 清理表格内段落首行/悬挂缩进，避免出现“部分单元格文本缩进”异常
            ind = pPr.find(_w("ind"))
            if ind is None:
                ind = etree.SubElement(pPr, _w("ind"))
            for attr in ("firstLine", "firstLineChars", "hanging", "hangingChars"):
                ind.set(_w(attr), "0")
            spacing = pPr.find(_w("spacing"))
            if spacing is None:
                spacing = etree.SubElement(pPr, _w("spacing"))
            spacing.set(_w("before"), "0")
            spacing.set(_w("after"), "0")
            spacing.set(_w("line"), line_val)
            spacing.set(_w("lineRule"), "auto")


def _normalize_equation_number_text(text: str) -> str | None:
    """将公式编号规范为 (x.y) / (x-y) 形式；保留原有分隔符。"""
    m = _RE_EQ_NUM_LOOSE.match(text or "")
    if not m:
        return None
    number_id = m.group(1)
    return f"({number_id})"


def _format_equation_number_cell(
        tc,
        *,
        alignment: str = "right",
        font_name: str = "Times New Roman",
        font_size_pt: float = 10.5):
    """格式化公式表格右侧编号单元格：括号、对齐、字体/字号。"""
    normalized = _normalize_equation_number_text(_get_cell_text(tc))
    if normalized is None:
        return

    target_alignment = _normalize_horizontal_alignment(alignment, default="right")
    target_font_name = str(font_name or "").strip() or "Times New Roman"
    try:
        target_font_size_pt = float(font_size_pt)
    except (TypeError, ValueError):
        target_font_size_pt = 10.5
    if target_font_size_pt <= 0:
        target_font_size_pt = 10.5
    target_size_half_pt = str(max(1, int(round(target_font_size_pt * 2))))

    # 统一文本为规范编号（避免出现无括号或全角括号的情况）
    text_nodes = tc.findall(f".//{_w('t')}")
    if text_nodes:
        text_nodes[0].text = normalized
        for t in text_nodes[1:]:
            t.text = ""
    else:
        paras = tc.findall(_w("p"))
        if paras:
            p = paras[0]
        else:
            p = etree.SubElement(tc, _w("p"))
        r = etree.SubElement(p, _w("r"))
        t = etree.SubElement(r, _w("t"))
        t.text = normalized

    # 段落右对齐
    for p in _iter_cell_paragraphs(tc):
        pPr = p.find(_w("pPr"))
        if pPr is None:
            pPr = etree.SubElement(p, _w("pPr"))
            p.insert(0, pPr)

        # 清理缩进，避免模板/样式继承导致编号列有效宽度变窄而被换行。
        ind = pPr.find(_w("ind"))
        if ind is None:
            ind = etree.SubElement(pPr, _w("ind"))
        for attr in (
            "left",
            "leftChars",
            "right",
            "rightChars",
            "firstLine",
            "firstLineChars",
            "hanging",
            "hangingChars",
        ):
            ind.set(_w(attr), "0")

        jc = pPr.find(_w("jc"))
        if jc is None:
            jc = etree.SubElement(pPr, _w("jc"))
        jc.set(_w("val"), target_alignment)

    # 编号字体/字号
    for r in tc.iter(_w("r")):
        rPr = r.find(_w("rPr"))
        if rPr is None:
            rPr = etree.SubElement(r, _w("rPr"))
            r.insert(0, rPr)
        rf = rPr.find(_w("rFonts"))
        if rf is None:
            rf = etree.SubElement(rPr, _w("rFonts"))
        apply_explicit_rfonts(
            rPr,
            font_cn=target_font_name,
            font_en=target_font_name,
            font_cs=target_font_name,
        )
        rf.set(_w("ascii"), target_font_name)
        rf.set(_w("hAnsi"), target_font_name)
        rf.set(_w("eastAsia"), target_font_name)
        for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            rf.attrib.pop(_w(attr), None)

        for sz_tag in ("sz", "szCs"):
            sz = rPr.find(_w(sz_tag))
            if sz is None:
                sz = etree.SubElement(rPr, _w(sz_tag))
            sz.set(_w("val"), target_size_half_pt)


def _set_tc_borders_none(tcPr) -> None:
    """将单元格边框显式设为 none，避免表格样式继承出框线。"""
    old_tc_borders = tcPr.find(_w("tcBorders"))
    if old_tc_borders is not None:
        tcPr.remove(old_tc_borders)
    tc_borders = etree.SubElement(tcPr, _w("tcBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = etree.SubElement(tc_borders, _w(side))
        b.set(_w("val"), "none")
        b.set(_w("sz"), "0")
        b.set(_w("space"), "0")


def _set_cell_paragraph_align(tc, align: str) -> None:
    """设置单元格内段落水平对齐。align: left/center/right"""
    for p in _iter_cell_paragraphs(tc):
        pPr = p.find(_w("pPr"))
        if pPr is None:
            pPr = etree.SubElement(p, _w("pPr"))
            p.insert(0, pPr)
        jc = pPr.find(_w("jc"))
        if jc is None:
            jc = etree.SubElement(pPr, _w("jc"))
        jc.set(_w("val"), align)


def _set_table_fixed_layout(tbl_el) -> None:
    """将表格布局设置为 fixed，保证各行使用统一列宽。"""
    tblPr = tbl_el.find(_w("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(tbl_el, _w("tblPr"))
        tbl_el.insert(0, tblPr)
    layout = tblPr.find(_w("tblLayout"))
    if layout is None:
        layout = etree.SubElement(tblPr, _w("tblLayout"))
    layout.set(_w("type"), "fixed")


def _equation_number_col_target_width(tbl_el, total_w: int) -> int:
    """估算公式表格右侧编号列目标宽度（仅容纳“(3.2)”这类编号）。"""
    target = max(_text_width_twips("(3.2)"), _text_width_twips("（3.2）"))
    has_number = False

    for tr in tbl_el.findall(_w("tr")):
        number_id, _, number_cell = _find_existing_equation_number_in_row(tr)
        if not number_id:
            continue
        has_number = True
        display = None
        if number_cell is not None:
            display = _extract_equation_number_display(_get_cell_text(number_cell))
        if display:
            target = max(target, _text_width_twips(display))
        else:
            target = max(target, _text_width_twips(f"({number_id})"))

    # 新生成的公式表格在“右列空占位”场景下，允许极限压缩右列，
    # 避免左侧公式显示在页面偏左位置。
    if not has_number:
        rows = tbl_el.findall(_w("tr"))
        right_cells = []
        for tr in rows:
            cells = tr.findall(_w("tc"))
            if cells:
                right_cells.append(cells[-1])
        if right_cells and all(_cell_is_effectively_empty(tc) for tc in right_cells):
            # 约 0.4~0.7cm：足够占位，后续若插入编号可再按编号内容扩展。
            return max(
                _MIN_EQ_PLACEHOLDER_COL_WIDTH,
                min(int(total_w * 0.06), max(int(total_w * 0.03), _MIN_EQ_PLACEHOLDER_COL_WIDTH)),
            )

    # 预留少量安全余量，并限制编号列不应过宽
    target += 80
    target = max(_MIN_COL_WIDTH, target)
    target = min(target, int(total_w * 0.28))
    return target


def _shrink_equation_number_column(tbl_el, total_w: int) -> None:
    """压缩公式表格最右侧编号列宽度，其余列分配剩余空间。"""
    if total_w <= 0:
        return

    grid_widths = _get_grid_cols(tbl_el)
    num_cols = len(grid_widths)
    if num_cols == 0:
        num_cols = _count_cols_from_rows(tbl_el)
        if num_cols == 0:
            return
        grid_widths = [1] * num_cols

    if num_cols < 2:
        _set_table_width(tbl_el, total_w)
        _set_table_fixed_layout(tbl_el)
        return

    right_w = _equation_number_col_target_width(tbl_el, total_w)
    min_left_total = _MIN_COL_WIDTH * (num_cols - 1)
    if total_w - right_w < min_left_total:
        right_w = max(_MIN_COL_WIDTH, total_w - min_left_total)

    left_total = total_w - right_w
    if left_total <= 0:
        return

    left_weights = [max(1.0, float(w)) for w in grid_widths[:-1]]
    left_widths = _distribute_widths(left_weights, left_total)
    new_widths = left_widths + [right_w]

    diff = total_w - sum(new_widths)
    if diff != 0 and new_widths:
        new_widths[0] += diff

    _set_table_width(tbl_el, total_w)
    _set_grid_cols(tbl_el, new_widths)
    _update_cell_widths(tbl_el, new_widths)
    _clear_row_width_exceptions(tbl_el)
    _set_table_fixed_layout(tbl_el)


def _format_equation_table(
        tbl_el,
        total_w: int | None = None,
        *,
        table_alignment: str = "center",
        formula_cell_alignment: str = "center",
        number_alignment: str = "right",
        number_font_name: str = "Times New Roman",
        number_font_size_pt: float = 10.5,
        auto_shrink_number_column: bool = True) -> int:
    """公式表格样式：无边框，单元格纵向居中，纯编号单元格右对齐+TNR 10.5pt。

    返回值：本次格式化到的纯编号单元格数量。
    """
    formatted_number_cells = 0
    table_alignment = _normalize_horizontal_alignment(table_alignment, default="center")
    formula_cell_alignment = _normalize_horizontal_alignment(formula_cell_alignment, default="center")
    number_alignment = _normalize_horizontal_alignment(number_alignment, default="right")
    tblPr = tbl_el.find(_w("tblPr"))
    if tblPr is None:
        tblPr = etree.SubElement(tbl_el, _w("tblPr"))
        tbl_el.insert(0, tblPr)
    # 移除表格样式，避免样式继承重绘边框
    old_style = tblPr.find(_w("tblStyle"))
    if old_style is not None:
        tblPr.remove(old_style)
    # 移除边框
    old_borders = tblPr.find(_w("tblBorders"))
    if old_borders is not None:
        tblPr.remove(old_borders)
    borders = etree.SubElement(tblPr, _w("tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = etree.SubElement(borders, _w(side))
        b.set(_w("val"), "none")
        b.set(_w("sz"), "0")
        b.set(_w("space"), "0")
    # 单元格边框也清除 + 单元格纵向居中
    for tc in tbl_el.iter(_w("tc")):
        tcPr = tc.find(_w("tcPr"))
        if tcPr is None:
            tcPr = etree.SubElement(tc, _w("tcPr"))
            tc.insert(0, tcPr)
        _set_tc_borders_none(tcPr)
        vAlign = tcPr.find(_w("vAlign"))
        if vAlign is None:
            vAlign = etree.SubElement(tcPr, _w("vAlign"))
        vAlign.set(_w("val"), "center")

    _set_table_alignment(tbl_el, table_alignment)

    # 若给定总可用宽度，压缩最右编号列宽度
    if auto_shrink_number_column and total_w is not None and total_w > 0:
        _shrink_equation_number_column(tbl_el, total_w)

    # 逐行规范纯编号单元格，并将左侧公式单元格居中（避免仅处理首行）
    for tr in tbl_el.findall(_w("tr")):
        _, pure_number_cell, matched_cell = _find_existing_equation_number_in_row(tr)
        if pure_number_cell is not None:
            _format_equation_number_cell(
                pure_number_cell,
                alignment=number_alignment,
                font_name=number_font_name,
                font_size_pt=number_font_size_pt,
            )
            formatted_number_cells += 1
        elif matched_cell is not None:
            # 兼容“表达式+(编号)同格”场景：纯文本单元格靠右；
            # 含公式/对象时避免强制改对齐，防止公式显示异常。
            if _cell_safe_for_plain_text_rewrite(matched_cell):
                _set_cell_paragraph_align(matched_cell, number_alignment)

        for tc in tr.findall(_w("tc")):
            if tc is pure_number_cell or tc is matched_cell:
                continue
            _set_cell_paragraph_align(tc, formula_cell_alignment)

    return formatted_number_cells


def _bold_first_row(tbl_el):
    """将表格第一行所有 run 设为加粗。"""
    rows = tbl_el.findall(_w("tr"))
    if not rows:
        return
    tr = rows[0]
    for r in tr.iter(_w("r")):
        rPr = r.find(_w("rPr"))
        if rPr is None:
            rPr = etree.SubElement(r, _w("rPr"))
            r.insert(0, rPr)
        b = rPr.find(_w("b"))
        if b is None:
            etree.SubElement(rPr, _w("b"))
        bCs = rPr.find(_w("bCs"))
        if bCs is None:
            etree.SubElement(rPr, _w("bCs"))


def _set_repeat_header_first_row(tbl_el) -> bool:
    """为首行开启 Word“重复标题行”（w:tblHeader）。"""
    rows = tbl_el.findall(_w("tr"))
    if not rows:
        return False
    tr = rows[0]
    trPr = tr.find(_w("trPr"))
    if trPr is None:
        trPr = etree.SubElement(tr, _w("trPr"))
        tr.insert(0, trPr)
    tbl_header = trPr.find(_w("tblHeader"))
    if tbl_header is None:
        tbl_header = etree.SubElement(trPr, _w("tblHeader"))
    tbl_header.set(_w("val"), "1")
    return True


class TableFormatRule(BaseRule):
    name = "table_format"
    description = "表格宽度自适应与字体格式"

    # 表内文字：5号（10.5pt），宋体 / Times New Roman
    FONT_CN = "宋体"
    FONT_EN = "Times New Roman"
    SIZE_PT = 10.5

    def apply(self, doc: Document, config: SceneConfig,
              tracker: ChangeTracker, context: dict) -> None:
        avail_w = _available_width(config)
        table_layout_mode = str(
            getattr(config, "normal_table_layout_mode", "smart")
        ).strip().lower()
        table_border_mode = str(
            getattr(config, "normal_table_border_mode", "full_grid")
        ).strip().lower()
        table_line_spacing_mode = str(
            getattr(config, "normal_table_line_spacing_mode", "single")
        ).strip().lower()
        try:
            smart_levels = int(getattr(config, "normal_table_smart_levels", 4))
        except (TypeError, ValueError):
            smart_levels = 4
        repeat_header_enabled = bool(
            getattr(config, "normal_table_repeat_header", False)
        )
        if table_layout_mode not in {"smart", "compact", "full"}:
            table_layout_mode = "smart"
        if table_border_mode not in {"full_grid", "three_line", "keep"}:
            table_border_mode = "full_grid"
        if table_line_spacing_mode not in {"single", "one_half", "double"}:
            table_line_spacing_mode = "single"
        if smart_levels not in {3, 4, 5, 6}:
            smart_levels = 4
        count = 0
        header_span_fix_count = 0
        repeat_header_count = 0
        border_applied_count = 0
        header_unit_break_count = 0
        body_paren_break_count = 0
        rich_content_table_skip_count = 0
        # 获取正文分区边界，仅格式化正文区域内的表格
        body_range = None
        target_indices = context.get("target_paragraph_indices")
        if target_indices is not None:
            target_indices = set(target_indices)
        doc_tree = context.get("doc_tree")
        has_doc_tree_scope = bool(getattr(doc_tree, "sections", None))
        if doc_tree:
            body_sec = doc_tree.get_section("body")
            if body_sec:
                start, end = body_sec.start_index, body_sec.end_index
                # 某些前置规则可能改动段落结构，导致 doc_tree 区间暂时失效（end < start）。
                # 保留该边界仅作为缺省回退；有 doc_tree 时优先按段落所属分区判定。
                if end >= start:
                    body_range = (start, end)

        def _table_section_type(pos: int) -> str | None:
            if not has_doc_tree_scope or pos < 0:
                return None
            try:
                return doc_tree.get_section_for_paragraph(pos)
            except Exception:
                return None

        def _is_table_in_scope(pos: int) -> bool:
            if target_indices is not None:
                if pos < 0 or pos not in target_indices:
                    return False
            sec_type = _table_section_type(pos)
            if sec_type is not None:
                return sec_type == "body"
            if body_range is not None:
                return pos >= 0 and body_range[0] <= pos <= body_range[1]
            return not has_doc_tree_scope

        # 构建表格（按文档顺序）→ 前置段落索引映射
        tbl_para_pos = top_level_table_anchor_positions(doc)

        # 智能总宽：先按 compact 目标估算每个常规表格宽度，再聚类到少量代表宽度。
        smart_width_plan: dict[int, int] = {}
        if table_layout_mode == "smart":
            raw_targets: list[tuple[int, int]] = []
            for tbl_idx, tbl in enumerate(doc.tables):
                tbl_el = tbl._tbl
                pos = tbl_para_pos[tbl_idx] if tbl_idx < len(tbl_para_pos) else -1
                if not _is_table_in_scope(pos):
                    continue
                if _table_has_drawing_or_object(tbl_el):
                    continue
                if _is_equation_table(tbl_el):
                    continue

                grid_widths = _get_grid_cols(tbl_el)
                num_cols = len(grid_widths)
                if num_cols == 0:
                    num_cols = _count_cols_from_rows(tbl_el)
                if num_cols <= 0:
                    continue
                min_ws = _first_row_min_widths(tbl_el, num_cols)
                body_min_ws = _non_first_row_min_widths(tbl_el, num_cols)
                raw_w = _target_normal_table_width(
                    avail_w, num_cols, min_ws, body_min_ws
                )
                raw_targets.append((tbl_idx, raw_w))
            smart_width_plan = _build_smart_width_plan(
                raw_targets, avail_w, max_levels=smart_levels
            )

        for tbl_idx, tbl in enumerate(doc.tables):
            tbl_el = tbl._tbl
            pos = tbl_para_pos[tbl_idx] if tbl_idx < len(tbl_para_pos) else -1

            # 跳过不在正文区域内的表格
            if not _is_table_in_scope(pos):
                continue
            if _table_has_drawing_or_object(tbl_el):
                rich_content_table_skip_count += 1
                continue
            # 公式表格由公式规则集处理（formula_to_table / equation_table_format / formula_style）。
            if _is_equation_table(tbl_el):
                continue

            # 常规表格：优先修复“首行跨列模式与后续行错位”
            if _normalize_first_row_span_pattern(tbl_el):
                header_span_fix_count += 1

            grid_widths = _get_grid_cols(tbl_el)
            num_cols = len(grid_widths)
            if num_cols == 0:
                num_cols = _count_cols_from_rows(tbl_el)
                if num_cols == 0:
                    continue
                # 无 tblGrid 时均分宽度
                grid_widths = [avail_w // num_cols] * num_cols

            # 1. 分析内容权重 + 首行最小宽度，分配列宽
            weights = _analyze_col_weights(tbl_el, num_cols)
            min_ws = _first_row_min_widths(tbl_el, num_cols)
            body_min_ws = _non_first_row_min_widths(tbl_el, num_cols)
            table_w = avail_w
            if table_layout_mode == "compact":
                table_w = _target_normal_table_width(
                    avail_w, num_cols, min_ws, body_min_ws
                )
            elif table_layout_mode == "smart":
                table_w = smart_width_plan.get(
                    tbl_idx,
                    _target_normal_table_width(avail_w, num_cols, min_ws, body_min_ws),
                )
            new_widths = _distribute_widths(weights, table_w, min_ws)
            # 1.1 二次调整：首行多行单元格收窄，释放宽度给其他列
            new_widths = _second_pass_rebalance_first_row(
                tbl_el, new_widths, weights, body_min_widths=body_min_ws
            )
            # 1.2 标题单位换行：若 "Title (unit)" 明显超宽，改为两行并重算列宽
            unit_breaks = _normalize_first_row_unit_breaks(tbl_el, new_widths)
            body_breaks = _normalize_body_paren_breaks(tbl_el, new_widths)
            if unit_breaks:
                header_unit_break_count += unit_breaks
            if body_breaks:
                body_paren_break_count += body_breaks
            if unit_breaks or body_breaks:
                weights = _analyze_col_weights(tbl_el, num_cols)
                min_ws = _first_row_min_widths(tbl_el, num_cols)
                body_min_ws = _non_first_row_min_widths(tbl_el, num_cols)
                new_widths = _distribute_widths(weights, table_w, min_ws)
                new_widths = _second_pass_rebalance_first_row(
                    tbl_el, new_widths, weights, body_min_widths=body_min_ws
                )

            # 2. 设置表格总宽、网格列宽、单元格宽
            _set_table_width(tbl_el, table_w)
            _set_grid_cols(tbl_el, new_widths)
            _update_cell_widths(tbl_el, new_widths)
            _set_table_alignment(
                tbl_el,
                "center" if table_layout_mode in {"smart", "compact"} and table_w < avail_w else "left"
            )

            # 3. 固定布局，避免 Word 对跨页行再次重排导致列宽不一致
            _clear_row_width_exceptions(tbl_el)
            self._set_fixed_layout(tbl_el)

            # 4. 设置表内字体
            _format_cell_fonts(
                tbl_el, self.FONT_CN, self.FONT_EN, self.SIZE_PT)

            # 5. 单元格段落格式：行距可配置，无段前段后
            _format_cell_paragraphs(tbl_el, table_line_spacing_mode)

            # 6. 所有单元格水平+垂直居中
            _center_all_cells(tbl_el)

            # 7. 常规表格边框样式（线宽从config读取，转换为 1/8pt 单位）
            if table_border_mode == "full_grid":
                grid_sz = int(
                    getattr(config, "table_border_width_pt", 0.5) * 8
                )
                _clear_cell_border_overrides(tbl_el)
                _set_table_borders(tbl_el, size_eighth_pt=max(1, grid_sz))
                border_applied_count += 1
            elif table_border_mode == "three_line":
                header_sz = int(
                    getattr(config, "three_line_header_width_pt", 1.0) * 8
                )
                bottom_sz = int(
                    getattr(config, "three_line_bottom_width_pt", 0.5) * 8
                )
                _set_three_line_borders(
                    tbl_el,
                    header_sz_eighth_pt=max(1, header_sz),
                    bottom_sz_eighth_pt=max(1, bottom_sz),
                )
                border_applied_count += 1

            # 8. 首行加粗
            _bold_first_row(tbl_el)

            # 9. 常规表格可选开启“重复标题行”（跨页重复表头）
            if repeat_header_enabled and _set_repeat_header_first_row(tbl_el):
                repeat_header_count += 1

            count += 1

        if count:
            mode_desc = (
                f"{table_layout_mode}(levels={smart_levels})"
                if table_layout_mode == "smart"
                else table_layout_mode
            )
            tracker.record(
                rule_name=self.name,
                target=f"{count} 个表格",
                section="global",
                change_type="format",
                before="(mixed widths)",
                after=f"mode={mode_desc}, border={table_border_mode}, line_spacing={table_line_spacing_mode}, max_width={avail_w}dxa, 5号宋体/TNR",
                paragraph_index=-1,
            )
        if header_span_fix_count:
            tracker.record(
                rule_name=self.name,
                target=f"{header_span_fix_count} 个表格首行",
                section="body",
                change_type="format",
                before="首行跨列与后续行错位",
                after="已按后续主模式对齐首行跨列",
                paragraph_index=-1,
            )
        if repeat_header_count:
            tracker.record(
                rule_name=self.name,
                target=f"{repeat_header_count} 个常规表格",
                section="body",
                change_type="format",
                before="首行未开启跨页重复",
                after="已开启 Word 重复标题行（跨页重复表头）",
                paragraph_index=-1,
            )
        if border_applied_count and table_border_mode in {"full_grid", "three_line"}:
            tracker.record(
                rule_name=self.name,
                target=f"{border_applied_count} 个常规表格",
                section="body",
                change_type="format",
                before="边框样式不统一",
                after=f"已强制统一为 {table_border_mode}",
                paragraph_index=-1,
            )
        if header_unit_break_count:
            tracker.record(
                rule_name=self.name,
                target=f"{header_unit_break_count} 个表头单元格",
                section="body",
                change_type="format",
                before="单位括号换行位置不稳定",
                after="已规范为“标题\\n(单位)”换行",
                paragraph_index=-1,
            )
        if body_paren_break_count:
            tracker.record(
                rule_name=self.name,
                target=f"{body_paren_break_count} 个正文单元格",
                section="body",
                change_type="format",
                before="尾括号换行位置不稳定",
                after="已规范为“主名\\n(括号内容)”换行",
                paragraph_index=-1,
            )
        if rich_content_table_skip_count:
            tracker.record(
                rule_name=self.name,
                target=f"{rich_content_table_skip_count} 个含图片/对象表格",
                section="body",
                change_type="skip",
                before="常规表格样式处理",
                after="已跳过（避免破坏非标准布局表格）",
                paragraph_index=-1,
            )

    @staticmethod
    def _set_fixed_layout(tbl_el):
        """设为 fixed 布局，保证各行沿用统一列宽。"""
        tblPr = tbl_el.find(_w("tblPr"))
        if tblPr is None:
            tblPr = etree.SubElement(tbl_el, _w("tblPr"))
            tbl_el.insert(0, tblPr)
        layout = tblPr.find(_w("tblLayout"))
        if layout is None:
            layout = etree.SubElement(tblPr, _w("tblLayout"))
        layout.set(_w("type"), "fixed")
