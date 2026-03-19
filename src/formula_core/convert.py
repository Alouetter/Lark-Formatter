"""Formula AST conversion helpers."""

from __future__ import annotations

import os
import re
import unicodedata
from copy import deepcopy
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

from lxml import etree

from .ast import FormulaNode
from .normalize import normalize_formula_node
from .semantics import BIG_OPERATOR_NAMES, FUNCTION_NAMES
from .symbols import LATEX_COMMAND_TO_UNICODE

try:
    from latex2mathml import converter as _latex2mathml_converter
except Exception:  # pragma: no cover - optional dependency fallback
    _latex2mathml_converter = None

_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_FUNCTION_PATTERN = "|".join(
    re.escape(name)
    for name in sorted(FUNCTION_NAMES, key=len, reverse=True)
)
_BIG_OPERATOR_PATTERN = "|".join(
    re.escape(name)
    for name in sorted(BIG_OPERATOR_NAMES, key=len, reverse=True)
)

_RE_SIMPLE_LATEX = re.compile(r"^[A-Za-z0-9\s+\-*/=().,:;\\^_{}]+$")
_RE_FRAC = re.compile(r"^\\frac\{(?P<num>.+)\}\{(?P<den>.+)\}$")
_RE_SQRT = re.compile(r"^\\sqrt\{(?P<body>.+)\}$")
_RE_SUB_SUP_A = re.compile(r"^(?P<base>[A-Za-z0-9]+)_\{(?P<sub>[^{}]+)\}\^\{(?P<sup>[^{}]+)\}$")
_RE_SUB_SUP_B = re.compile(r"^(?P<base>[A-Za-z0-9]+)\^\{(?P<sup>[^{}]+)\}_\{(?P<sub>[^{}]+)\}$")
_RE_SUP = re.compile(r"^(?P<base>[A-Za-z0-9]+)\^\{(?P<sup>[^{}]+)\}$")
_RE_SUB = re.compile(r"^(?P<base>[A-Za-z0-9]+)_\{(?P<sub>[^{}]+)\}$")
_RE_BIG_OPERATOR = re.compile(
    r"^\\(?P<op>" + _BIG_OPERATOR_PATTERN + r")(?![A-Za-z])"
    r"(?:_\{(?P<sub>[^{}]+)\})?"
    r"(?:\^\{(?P<sup>[^{}]+)\})?"
    r"\s*(?P<body>[\s\S]*)$"
)
_RE_COMPOUND_MARKER = re.compile(r"[\^_{]|\\[A-Za-z]+\b")
_RE_MATRIX = re.compile(
    r"^\\begin\{(?P<env>[pbvV]?matrix)\}(?P<body>[\s\S]+)\\end\{(?P=env)\}$"
)
_RE_ALIGNED = re.compile(
    r"^\\begin\{(?P<env>aligned)\}(?P<body>[\s\S]+)\\end\{(?P=env)\}$"
)
_RE_DELIMITED = re.compile(
    r"^\\left(?P<left>\\[A-Za-z]+|[^\s])\s*(?P<body>[\s\S]+)\s*\\right(?P<right>\\[A-Za-z]+|[^\s])$"
)
_RE_FUNCTION = re.compile(
    r"^\\(?P<fn>" + _FUNCTION_PATTERN + r")(?![A-Za-z])\s*(?P<arg>[\s\S]+)$"
)
_RE_DIRECT_STYLE_TEXT = re.compile(r"^[^\s\\^_+=\-*/=()[\]{}]+(?:\s+[^\s\\^_+=\-*/=()[\]{}]+)*$")

_LATEX_DELIM_TO_CHAR = {
    "(": "(",
    ")": ")",
    "[": "[",
    "]": "]",
    "{": "{",
    "}": "}",
    "|": "|",
    ".": ".",
    r"\{": "{",
    r"\}": "}",
    r"\|": "|",
    r"\langle": "⟨",
    r"\rangle": "⟩",
    r"\lfloor": "⌊",
    r"\rfloor": "⌋",
    r"\lceil": "⌈",
    r"\rceil": "⌉",
}

_CHAR_TO_LATEX_DELIM = {
    "{": r"\{",
    "}": r"\}",
    "⟨": r"\langle",
    "⟩": r"\rangle",
    "⌊": r"\lfloor",
    "⌋": r"\rfloor",
    "⌈": r"\lceil",
    "⌉": r"\rceil",
}

_SOURCE_USING_LATEX_CANDIDATE = {
    "latex",
    "plain_text",
    "ocr_fragment",
    "unicode_text",
    "old_equation",
    "mathtype",
    "ole_equation",
}

_SYMBOL_COMMANDS = dict(LATEX_COMMAND_TO_UNICODE)


@dataclass
class ConversionOutcome:
    """Unified conversion result used by formula rules."""

    success: bool
    target_mode: str
    confidence: float
    reason: str = ""
    warnings: list[str] = field(default_factory=list)
    transformed: bool = False
    latex_text: str = ""
    omml_element: etree._Element | None = None


def _clamp(value: float) -> float:
    try:
        n = float(value)
    except (TypeError, ValueError):
        n = 0.0
    if n < 0.0:
        return 0.0
    if n > 1.0:
        return 1.0
    return n


def _decode_latex_delim(token: str) -> str:
    value = str(token or "").strip()
    if not value:
        return "."
    return _LATEX_DELIM_TO_CHAR.get(value, value)


def _encode_latex_delim(char: str) -> str:
    value = str(char or "").strip()
    if not value:
        return "."
    return _CHAR_TO_LATEX_DELIM.get(value, value)


def _build_math_text_run(
        text: str,
        *,
        style: str | None = None,
        script: str | None = None) -> etree._Element:
    m_r = etree.Element(f"{{{_M_NS}}}r")
    if style or script:
        m_rpr = etree.SubElement(m_r, f"{{{_M_NS}}}rPr")
        if script:
            scr = etree.SubElement(m_rpr, f"{{{_M_NS}}}scr")
            scr.set(f"{{{_M_NS}}}val", str(script))
        if style:
            sty = etree.SubElement(m_rpr, f"{{{_M_NS}}}sty")
            sty.set(f"{{{_M_NS}}}val", str(style))
    m_t = etree.SubElement(m_r, f"{{{_M_NS}}}t")
    m_t.text = str(text or "")
    return m_r


def _append_math_arg_content(target: etree._Element, content) -> None:
    if isinstance(content, etree._Element):
        target.append(deepcopy(content))
        return

    if isinstance(content, (list, tuple)):
        for item in content:
            if isinstance(item, etree._Element):
                target.append(deepcopy(item))
            elif str(item or ""):
                target.append(_build_math_text_run(str(item or "")))
        return

    raw = str(content or "")
    structured = bool(re.search(r"[\\^_+=\-*/]", raw))
    inner_elements = _tokenize_latex(raw) if structured else []
    if inner_elements:
        for item in inner_elements:
            target.append(item)
        return
    target.append(_build_math_text_run(raw))


def _build_math_arg(local_name: str, text) -> etree._Element:
    elem = etree.Element(f"{{{_M_NS}}}{local_name}")
    _append_math_arg_content(elem, text)
    return elem


def _wrap_math(content: etree._Element, *, block: bool) -> etree._Element:
    if block:
        root = etree.Element(f"{{{_M_NS}}}oMathPara")
        omath = etree.SubElement(root, f"{{{_M_NS}}}oMath")
    else:
        root = etree.Element(f"{{{_M_NS}}}oMath")
        omath = root
    omath.append(content)
    return root


def _build_literal_omml(expr: str, *, block: bool) -> etree._Element:
    return _wrap_math(_build_math_text_run(expr), block=block)


def _clone_omml_element(element: etree._Element) -> etree._Element:
    return deepcopy(element)


def _normalize_omml_block_mode(omml: etree._Element, *, block: bool) -> etree._Element:
    local = etree.QName(omml).localname

    if block:
        if local == "oMathPara":
            return _clone_omml_element(omml)
        if local == "oMath":
            root = etree.Element(f"{{{_M_NS}}}oMathPara")
            wrapped = etree.SubElement(root, f"{{{_M_NS}}}oMath")
            for child in omml:
                wrapped.append(_clone_omml_element(child))
            return root
        return _wrap_math(_clone_omml_element(omml), block=True)

    if local == "oMath":
        return _clone_omml_element(omml)
    if local == "oMathPara":
        inner = next(
            (
                child for child in omml
                if etree.QName(child).localname == "oMath"
            ),
            None,
        )
        if inner is not None:
            return _clone_omml_element(inner)
    return _clone_omml_element(omml)


def _omml_contains_latex_literal(omml: etree._Element | None) -> bool:
    if omml is None:
        return False
    try:
        text_nodes = omml.xpath(
            ".//*[namespace-uri()='%s' and local-name()='t']/text()" % _M_NS
        )
    except Exception:
        xml = etree.tostring(omml, encoding="unicode")
        return "\\" in xml
    return any("\\" in str(text or "") for text in text_nodes)


def _candidate_mml2omml_xsl_paths() -> list[Path]:
    candidates: list[Path] = []

    env_path = str(os.getenv("DOCX_MML2OMML_XSL_PATH", "") or "").strip()
    if env_path:
        candidates.append(Path(env_path))

    program_dirs = [
        os.getenv("ProgramFiles"),
        os.getenv("ProgramFiles(x86)"),
    ]
    office_roots = [
        ("Microsoft Office", "root", "Office16", "MML2OMML.XSL"),
        ("Microsoft Office", "root", "Office15", "MML2OMML.XSL"),
        ("Microsoft Office", "Office16", "MML2OMML.XSL"),
        ("Microsoft Office", "Office15", "MML2OMML.XSL"),
    ]
    for base in program_dirs:
        if not base:
            continue
        base_path = Path(base)
        for suffix in office_roots:
            candidates.append(base_path.joinpath(*suffix))

    seen: set[str] = set()
    ordered: list[Path] = []
    for path in candidates:
        key = str(path).lower()
        if key in seen:
            continue
        seen.add(key)
        ordered.append(path)
    return ordered


@lru_cache(maxsize=1)
def _get_mathml_to_omml_transform() -> etree.XSLT | None:
    for path in _candidate_mml2omml_xsl_paths():
        try:
            if not path.is_file():
                continue
            return etree.XSLT(etree.parse(str(path)))
        except Exception:
            continue
    return None


def _has_mathml_to_omml_support() -> bool:
    return _latex2mathml_converter is not None and _get_mathml_to_omml_transform() is not None


def _convert_latex_via_mathml(expr: str, *, block: bool) -> etree._Element | None:
    if not _has_mathml_to_omml_support():
        return None
    transform = _get_mathml_to_omml_transform()

    raw = str(expr or "").strip()
    if not raw:
        return None

    try:
        mathml_text = _latex2mathml_converter.convert(raw)
        mathml_root = etree.fromstring(mathml_text.encode("utf-8"))
        omml_root = transform(mathml_root).getroot()
        if omml_root is None:
            return None
        return _normalize_omml_block_mode(omml_root, block=block)
    except Exception:
        return None


def convert_mathml_to_omml(mathml_text: str, *, block: bool) -> etree._Element | None:
    """Convert MathML/XML text to OMML, if Office transform support is available."""
    transform = _get_mathml_to_omml_transform()
    if transform is None:
        return None

    raw = str(mathml_text or "").strip()
    if not raw:
        return None

    try:
        root = etree.fromstring(raw.encode("utf-8"))
    except Exception:
        return None

    try:
        local = etree.QName(root).localname
    except Exception:
        local = ""

    math_root = root
    if local != "math":
        math_nodes = root.xpath(".//*[local-name()='math']")
        math_root = math_nodes[0] if math_nodes else None
    if math_root is None:
        return None

    try:
        omml_root = transform(math_root).getroot()
        if omml_root is None:
            return None
        return _normalize_omml_block_mode(omml_root, block=block)
    except Exception:
        return None


# ── Inner builders (no oMath wrapping, for composition) ──

def _inner_fraction(num: str, den: str) -> etree._Element:
    frac = etree.Element(f"{{{_M_NS}}}f")
    frac.append(_build_math_arg("num", num))
    frac.append(_build_math_arg("den", den))
    return frac


def _inner_sqrt(body: str) -> etree._Element:
    rad = etree.Element(f"{{{_M_NS}}}rad")
    deg_hide = etree.SubElement(rad, f"{{{_M_NS}}}degHide")
    deg_hide.set(f"{{{_M_NS}}}val", "1")
    rad.append(_build_math_arg("e", body))
    return rad


def _inner_sup(base: str, sup: str) -> etree._Element:
    sup_el = etree.Element(f"{{{_M_NS}}}sSup")
    sup_el.append(_build_math_arg("e", base))
    sup_el.append(_build_math_arg("sup", sup))
    return sup_el


def _inner_sub(base: str, sub: str) -> etree._Element:
    sub_el = etree.Element(f"{{{_M_NS}}}sSub")
    sub_el.append(_build_math_arg("e", base))
    sub_el.append(_build_math_arg("sub", sub))
    return sub_el


def _inner_sub_sup(base: str, sub: str, sup: str) -> etree._Element:
    el = etree.Element(f"{{{_M_NS}}}sSubSup")
    el.append(_build_math_arg("e", base))
    el.append(_build_math_arg("sub", sub))
    el.append(_build_math_arg("sup", sup))
    return el


def _inner_pre_sub_sup(base, sub: str, sup: str) -> etree._Element:
    el = etree.Element(f"{{{_M_NS}}}sPre")
    el.append(_build_math_arg("sub", sub))
    el.append(_build_math_arg("sup", sup))
    el.append(_build_math_arg("e", base))
    return el


def _inner_big_operator(operator: str, sub: str, sup: str, body: str) -> etree._Element:
    symbol = {
        "sum": "∑",
        "int": "∫",
        "iint": "∬",
        "iiint": "∭",
        "oint": "∮",
        "prod": "∏",
        "lim": "lim",
    }.get(str(operator or "").strip().lower(), "∑")
    nary = etree.Element(f"{{{_M_NS}}}nary")
    nary_pr = etree.SubElement(nary, f"{{{_M_NS}}}naryPr")
    chr_el = etree.SubElement(nary_pr, f"{{{_M_NS}}}chr")
    chr_el.set(f"{{{_M_NS}}}val", symbol)
    if sub:
        nary.append(_build_math_arg("sub", sub))
    if sup:
        nary.append(_build_math_arg("sup", sup))
    nary.append(_build_math_arg("e", body or ""))
    return nary


def _inner_delimited(left: str, right: str, body: str) -> etree._Element:
    d = etree.Element(f"{{{_M_NS}}}d")
    dpr = etree.SubElement(d, f"{{{_M_NS}}}dPr")
    beg = etree.SubElement(dpr, f"{{{_M_NS}}}begChr")
    end = etree.SubElement(dpr, f"{{{_M_NS}}}endChr")
    beg.set(f"{{{_M_NS}}}val", left or ".")
    end.set(f"{{{_M_NS}}}val", right or ".")
    d.append(_build_math_arg("e", body or ""))
    return d


def _inner_matrix(rows: list[list[str]]) -> etree._Element:
    matrix = etree.Element(f"{{{_M_NS}}}m")
    etree.SubElement(matrix, f"{{{_M_NS}}}mPr")
    for row in rows:
        mr = etree.SubElement(matrix, f"{{{_M_NS}}}mr")
        for cell in row:
            mr.append(_build_math_arg("e", cell or ""))
    return matrix


def _inner_function(name: str, arg: str) -> etree._Element:
    func = etree.Element(f"{{{_M_NS}}}func")
    fname = etree.SubElement(func, f"{{{_M_NS}}}fName")
    fname.append(_build_math_text_run(name or "f"))
    func.append(_build_math_arg("e", arg or ""))
    return func


def _inner_lim_upp_accent(body, accent_text: str) -> etree._Element:
    lim_upp = etree.Element(f"{{{_M_NS}}}limUpp")
    lim_upp.append(_build_math_arg("e", body))
    lim = etree.SubElement(lim_upp, f"{{{_M_NS}}}lim")
    lim.append(_build_math_text_run(accent_text))
    return lim_upp


def _inner_bar_top(body) -> etree._Element:
    bar = etree.Element(f"{{{_M_NS}}}bar")
    bar_pr = etree.SubElement(bar, f"{{{_M_NS}}}barPr")
    pos = etree.SubElement(bar_pr, f"{{{_M_NS}}}pos")
    pos.set(f"{{{_M_NS}}}val", "top")
    bar.append(_build_math_arg("e", body))
    return bar


def _inner_acc_char(body, chr_val: str) -> etree._Element:
    acc = etree.Element(f"{{{_M_NS}}}acc")
    acc_pr = etree.SubElement(acc, f"{{{_M_NS}}}accPr")
    chr_el = etree.SubElement(acc_pr, f"{{{_M_NS}}}chr")
    chr_el.set(f"{{{_M_NS}}}val", chr_val)
    acc.append(_build_math_arg("e", body))
    return acc


# ── Wrapped builders (backward compatible) ──

def _build_fraction_omml(num: str, den: str, *, block: bool) -> etree._Element:
    return _wrap_math(_inner_fraction(num, den), block=block)


def _build_sqrt_omml(body: str, *, block: bool) -> etree._Element:
    return _wrap_math(_inner_sqrt(body), block=block)


def _build_sup_omml(base: str, sup: str, *, block: bool) -> etree._Element:
    return _wrap_math(_inner_sup(base, sup), block=block)


def _build_sub_omml(base: str, sub: str, *, block: bool) -> etree._Element:
    return _wrap_math(_inner_sub(base, sub), block=block)


def _build_sub_sup_omml(base: str, sub: str, sup: str, *, block: bool) -> etree._Element:
    return _wrap_math(_inner_sub_sup(base, sub, sup), block=block)


def _build_big_operator_omml(operator: str, sub: str, sup: str, body: str, *, block: bool) -> etree._Element:
    return _wrap_math(_inner_big_operator(operator, sub, sup, body), block=block)


def _build_delimited_omml(left: str, right: str, body: str, *, block: bool) -> etree._Element:
    return _wrap_math(_inner_delimited(left, right, body), block=block)


def _build_matrix_omml(rows: list[list[str]], *, block: bool) -> etree._Element:
    return _wrap_math(_inner_matrix(rows), block=block)


def _build_function_omml(name: str, arg: str, *, block: bool) -> etree._Element:
    return _wrap_math(_inner_function(name, arg), block=block)


# ── Compound expression tokenizer ──

_FUNCTION_NAMES = FUNCTION_NAMES
_BIG_OP_NAMES = BIG_OPERATOR_NAMES


def _parse_brace_group(expr: str, pos: int) -> tuple[str, int]:
    """Parse {content} or a single character at *pos*."""
    if pos >= len(expr):
        return "", pos
    if expr[pos] == "{":
        depth = 1
        i = pos + 1
        while i < len(expr) and depth > 0:
            if expr[i] == "{":
                depth += 1
            elif expr[i] == "}":
                depth -= 1
            i += 1
        return expr[pos + 1 : i - 1], i
    return expr[pos], pos + 1


def _parse_command_name(expr: str, pos: int) -> tuple[str, int]:
    """Parse \\commandname starting after the backslash."""
    i = pos
    while i < len(expr) and expr[i].isalpha():
        i += 1
    return expr[pos:i], i


def _skip_spaces(expr: str, pos: int) -> int:
    while pos < len(expr) and expr[pos].isspace():
        pos += 1
    return pos


def _parse_optional_scripts(expr: str, pos: int) -> tuple[str, str, int]:
    cur = pos
    sub_text = ""
    sup_text = ""
    for _ in range(2):
        cur = _skip_spaces(expr, cur)
        if cur >= len(expr):
            break
        if expr[cur] == "_":
            sub_text, cur = _parse_brace_group(expr, cur + 1)
            continue
        if expr[cur] == "^":
            sup_text, cur = _parse_brace_group(expr, cur + 1)
            continue
        break
    return sub_text, sup_text, cur


def _build_textish_base(text: str, sub_text: str, sup_text: str) -> etree._Element:
    base = str(text or "")
    if sub_text and sup_text:
        return _inner_sub_sup(base, sub_text, sup_text)
    if sup_text:
        return _inner_sup(base, sup_text)
    if sub_text:
        return _inner_sub(base, sub_text)
    return _build_math_text_run(base)


_STYLEABLE_MATH_SYMBOLS = {
    "∂",
    "∇",
}


def _merge_math_style(existing: str | None, added: str | None) -> str | None:
    current = str(existing or "").strip() or None
    incoming = str(added or "").strip() or None
    if incoming is None:
        return current
    if current is None or current == incoming:
        return incoming
    if "bi" in {current, incoming}:
        return "bi"
    if {current, incoming} == {"b", "i"}:
        return "bi"
    if incoming == "p":
        return current
    if current == "p":
        return incoming
    return incoming


def _merge_math_script(existing: str | None, added: str | None) -> str | None:
    incoming = str(added or "").strip() or None
    current = str(existing or "").strip() or None
    return incoming if incoming is not None else current


def _math_run_prop_value(run: etree._Element, local_name: str) -> str | None:
    rpr = next(
        (
            child for child in run
            if etree.QName(child).localname == "rPr"
        ),
        None,
    )
    if rpr is None:
        return None
    prop = next(
        (
            child for child in rpr
            if etree.QName(child).localname == local_name
        ),
        None,
    )
    if prop is None:
        return None
    value = prop.get(f"{{{_M_NS}}}val")
    text = str(value or "").strip()
    return text or None


def _is_styleable_math_char(ch: str) -> bool:
    if not ch or ch.isspace():
        return False
    if ch.isalnum():
        return True
    if ch in _STYLEABLE_MATH_SYMBOLS:
        return True
    category = unicodedata.category(ch)
    return category.startswith("L")


def _split_math_text_for_style(text: str) -> list[tuple[str, bool]]:
    raw = str(text or "")
    if not raw:
        return []
    chunks: list[tuple[str, bool]] = []
    buf = raw[0]
    last_flag = _is_styleable_math_char(raw[0])
    for ch in raw[1:]:
        flag = _is_styleable_math_char(ch)
        if flag == last_flag:
            buf += ch
            continue
        chunks.append((buf, last_flag))
        buf = ch
        last_flag = flag
    chunks.append((buf, last_flag))
    return chunks


def _style_math_run(
        run: etree._Element,
        style: str | None,
        script: str | None) -> list[etree._Element]:
    text = "".join(
        str(child.text or "")
        for child in run
        if etree.QName(child).localname == "t"
    )
    existing_style = _math_run_prop_value(run, "sty")
    existing_script = _math_run_prop_value(run, "scr")
    merged_style = _merge_math_style(existing_style, style)
    merged_script = _merge_math_script(existing_script, script)

    if not text:
        return [_build_math_text_run("", style=merged_style, script=merged_script)]

    segments = _split_math_text_for_style(text)
    if not segments:
        return [_build_math_text_run(text, style=merged_style, script=merged_script)]

    styled_runs: list[etree._Element] = []
    for segment_text, should_style in segments:
        seg_style = merged_style if should_style else existing_style
        seg_script = merged_script if should_style else existing_script
        styled_runs.append(
            _build_math_text_run(
                segment_text,
                style=seg_style,
                script=seg_script,
            )
        )
    return styled_runs


def _apply_math_style_to_element_inplace(
        element: etree._Element,
        style: str | None,
        script: str | None) -> list[etree._Element]:
    if etree.QName(element).localname == "r":
        return _style_math_run(element, style, script)

    for child in list(element):
        replacement_children = _apply_math_style_to_element_inplace(child, style, script)
        if len(replacement_children) == 1 and replacement_children[0] is child:
            continue
        insert_at = element.index(child)
        element.remove(child)
        for new_child in replacement_children:
            element.insert(insert_at, new_child)
            insert_at += 1
    return [element]


def _apply_math_style_to_elements(
        elements: list[etree._Element],
        style: str | None,
        script: str | None) -> list[etree._Element]:
    styled: list[etree._Element] = []
    for element in elements:
        styled.extend(
            _apply_math_style_to_element_inplace(
                deepcopy(element),
                style,
                script,
            )
        )
    return styled


def _parse_group_or_token_elements(
        expr: str,
        pos: int) -> tuple[list[etree._Element], int] | None:
    cur = _skip_spaces(expr, pos)
    if cur >= len(expr):
        return None

    if expr[cur] == "{":
        group, after = _parse_brace_group(expr, cur)
        inner = _tokenize_latex(group)
        if inner:
            return inner, after
        return [_build_math_text_run(group)], after

    if expr[cur] == "\\":
        cmd, cmd_end = _parse_command_name(expr, cur + 1)
        if cmd:
            fragment = expr[cur:cmd_end]
            inner = _tokenize_latex(fragment)
            if inner:
                return inner, cmd_end
            return [_build_math_text_run(fragment)], cmd_end
        if cur + 1 < len(expr):
            return [_build_math_text_run(expr[cur + 1])], cur + 2

    return [_build_math_text_run(expr[cur])], cur + 1


def _styled_command_spec(cmd_lower: str) -> tuple[str | None, str | None] | None:
    mapping = {
        "mathbf": ("b", None),
        "boldsymbol": ("bi", None),
        "mathrm": ("p", None),
        "operatorname": ("p", None),
        "mathit": ("i", None),
        "mathbb": ("p", "double-struck"),
        "mathcal": ("p", "script"),
        "mathscr": ("p", "script"),
        "mathfrak": ("p", "fraktur"),
        "mathsf": ("p", "sans-serif"),
        "mathtt": ("p", "monospace"),
    }
    return mapping.get(str(cmd_lower or "").strip().lower())


def _parse_direct_style_argument(expr: str, pos: int) -> tuple[str, int] | None:
    cur = _skip_spaces(expr, pos)
    if cur >= len(expr):
        return None

    if expr[cur] == "{":
        arg, after = _parse_brace_group(expr, cur)
        raw = str(arg or "").strip()
        if raw and _RE_DIRECT_STYLE_TEXT.fullmatch(raw):
            return raw, after
        return None

    token = expr[cur]
    if token in {"\\", "^", "_", "{", "}"}:
        return None
    return token, cur + 1


def _has_top_level_operator(expr: str) -> bool:
    raw = str(expr or "").strip()
    if not raw:
        return False

    brace_depth = 0
    paren_depth = 0
    bracket_depth = 0
    prev_non_space = ""
    for idx, ch in enumerate(raw):
        if ch == "\\":
            continue
        if ch == "{":
            brace_depth += 1
            continue
        if ch == "}":
            brace_depth = max(0, brace_depth - 1)
            continue
        if ch == "(":
            paren_depth += 1
            continue
        if ch == ")":
            paren_depth = max(0, paren_depth - 1)
            continue
        if ch == "[":
            bracket_depth += 1
            continue
        if ch == "]":
            bracket_depth = max(0, bracket_depth - 1)
            continue
        if brace_depth or paren_depth or bracket_depth:
            if not ch.isspace():
                prev_non_space = ch
            continue
        if ch == "=":
            return True
        if ch in "+-":
            if idx == 0:
                prev_non_space = ch
                continue
            if prev_non_space in {"", "=", "+", "-", "*", "/", "^", "_", "(", "[", "{", "\\"}:
                prev_non_space = ch
                continue
            return True
        if not ch.isspace():
            prev_non_space = ch
    return False


def _parse_single_fraction(expr: str) -> tuple[dict, str] | None:
    raw = str(expr or "").strip()
    if not raw.startswith(r"\frac"):
        return None
    pos = _skip_spaces(raw, len(r"\frac"))
    if pos >= len(raw) or raw[pos] != "{":
        return None
    num, pos = _parse_brace_group(raw, pos)
    pos = _skip_spaces(raw, pos)
    if pos >= len(raw) or raw[pos] != "{":
        return None
    den, pos = _parse_brace_group(raw, pos)
    if raw[pos:].strip():
        return None
    return {"num": num.strip(), "den": den.strip()}, "fraction"


def _parse_single_sqrt(expr: str) -> tuple[dict, str] | None:
    raw = str(expr or "").strip()
    if not raw.startswith(r"\sqrt"):
        return None
    pos = _skip_spaces(raw, len(r"\sqrt"))
    if pos >= len(raw) or raw[pos] != "{":
        return None
    body, pos = _parse_brace_group(raw, pos)
    if raw[pos:].strip():
        return None
    return {"body": body.strip()}, "sqrt"


def _parse_balanced_delimiter(expr: str, pos: int, open_ch: str, close_ch: str) -> tuple[str, int]:
    if pos >= len(expr) or expr[pos] != open_ch:
        return "", pos
    depth = 1
    idx = pos + 1
    while idx < len(expr) and depth > 0:
        if expr[idx] == open_ch:
            depth += 1
        elif expr[idx] == close_ch:
            depth -= 1
        idx += 1
    if depth != 0:
        return "", pos
    return expr[pos:idx], idx


def _flush_text_buf(buf: str, elements: list[etree._Element]) -> None:
    if buf:
        elements.append(_build_math_text_run(buf))


def _take_script_base(
        elements: list[etree._Element],
        text_buf: str) -> tuple[str | etree._Element, str]:
    if text_buf:
        base = text_buf[-1]
        prefix = text_buf[:-1]
        if prefix:
            _flush_text_buf(prefix, elements)
        return base, ""

    if elements:
        return elements.pop(), ""

    return "", ""


def _has_script_base(base) -> bool:
    if isinstance(base, str):
        return base != ""
    return base is not None


def _tokenize_latex(expr: str) -> list[etree._Element]:
    """Tokenize a LaTeX expression into a list of inner OMML elements."""
    elements: list[etree._Element] = []
    text_buf = ""
    pos = 0
    n = len(expr)

    while pos < n:
        ch = expr[pos]

        # ── LaTeX command ──
        if ch == "\\":
            cmd, cmd_end = _parse_command_name(expr, pos + 1)
            cmd_lower = cmd.lower()

            if cmd_lower == "frac" and cmd_end < n and expr[cmd_end] == "{":
                _flush_text_buf(text_buf, elements)
                text_buf = ""
                num, after_num = _parse_brace_group(expr, cmd_end)
                den, after_den = _parse_brace_group(expr, after_num)
                elements.append(_inner_fraction(num, den))
                pos = after_den
                continue

            if cmd_lower == "sqrt" and cmd_end < n and expr[cmd_end] == "{":
                _flush_text_buf(text_buf, elements)
                text_buf = ""
                body, after = _parse_brace_group(expr, cmd_end)
                elements.append(_inner_sqrt(body))
                pos = after
                continue

            if cmd_lower == "prescript":
                cur = _skip_spaces(expr, cmd_end)
                if cur < n and expr[cur] == "{":
                    sup_text, cur = _parse_brace_group(expr, cur)
                    cur = _skip_spaces(expr, cur)
                    if cur < n and expr[cur] == "{":
                        sub_text, cur = _parse_brace_group(expr, cur)
                        cur = _skip_spaces(expr, cur)
                        if cur < n and expr[cur] == "{":
                            base_text, cur = _parse_brace_group(expr, cur)
                            _flush_text_buf(text_buf, elements)
                            text_buf = ""
                            base_elements = _tokenize_latex(base_text)
                            base_content = base_elements if base_elements else base_text
                            elements.append(_inner_pre_sub_sup(base_content, sub_text, sup_text))
                            pos = cur
                            continue

            if cmd_lower in _BIG_OP_NAMES:
                _flush_text_buf(text_buf, elements)
                text_buf = ""
                cur = cmd_end
                sub_text, sup_text, cur = _parse_optional_scripts(expr, cur)
                cur = _skip_spaces(expr, cur)
                # Collect body until next command or top-level operator
                body_start = cur
                depth = 0
                paren_depth = 0
                bracket_depth = 0
                while cur < n:
                    ch_cur = expr[cur]
                    if ch_cur == "{":
                        depth += 1
                    elif ch_cur == "}":
                        if depth == 0:
                            break
                        depth -= 1
                    elif ch_cur == "(":
                        paren_depth += 1
                    elif ch_cur == ")":
                        if paren_depth == 0:
                            break
                        paren_depth -= 1
                    elif ch_cur == "[":
                        bracket_depth += 1
                    elif ch_cur == "]":
                        if bracket_depth == 0:
                            break
                        bracket_depth -= 1
                    elif depth == 0 and paren_depth == 0 and bracket_depth == 0:
                        if ch_cur == "\\" or ch_cur in "+-=":
                            break
                    cur += 1
                body_text = expr[body_start:cur].strip()
                elements.append(_inner_big_operator(cmd_lower, sub_text, sup_text, body_text))
                pos = cur
                continue

            if cmd_lower == "left":
                _flush_text_buf(text_buf, elements)
                text_buf = ""
                # Parse left delimiter
                if cmd_end < n and expr[cmd_end] == "\\":
                    delim_cmd, delim_end = _parse_command_name(expr, cmd_end + 1)
                    left_delim = _decode_latex_delim("\\" + delim_cmd)
                    cur = delim_end
                elif cmd_end < n:
                    left_delim = _decode_latex_delim(expr[cmd_end])
                    cur = cmd_end + 1
                else:
                    left_delim = "("
                    cur = cmd_end
                # Find matching \right
                right_pos = expr.find("\\right", cur)
                if right_pos >= 0:
                    body_text = expr[cur:right_pos].strip()
                    r_cmd_end = right_pos + 6  # len("\\right")
                    if r_cmd_end < n and expr[r_cmd_end] == "\\":
                        r_delim_cmd, r_delim_end = _parse_command_name(expr, r_cmd_end + 1)
                        right_delim = _decode_latex_delim("\\" + r_delim_cmd)
                        cur = r_delim_end
                    elif r_cmd_end < n:
                        right_delim = _decode_latex_delim(expr[r_cmd_end])
                        cur = r_cmd_end + 1
                    else:
                        right_delim = ")"
                        cur = r_cmd_end
                    elements.append(_inner_delimited(left_delim, right_delim, body_text))
                    pos = cur
                    continue
                # No matching \right found, treat as text
                text_buf += expr[pos:cmd_end]
                pos = cmd_end
                continue

            if cmd_lower in _FUNCTION_NAMES:
                _flush_text_buf(text_buf, elements)
                text_buf = ""
                cur = cmd_end
                sub_text, sup_text, cur = _parse_optional_scripts(expr, cur)
                cur = _skip_spaces(expr, cur)
                if cur < n and expr[cur] == "{":
                    arg, cur = _parse_brace_group(expr, cur)
                elif cur < n and expr[cur] == "(":
                    arg, cur = _parse_balanced_delimiter(expr, cur, "(", ")")
                elif cur < n and expr[cur] == "[":
                    arg, cur = _parse_balanced_delimiter(expr, cur, "[", "]")
                else:
                    arg_start = cur
                    depth = 0
                    paren_depth = 0
                    bracket_depth = 0
                    while cur < n:
                        if expr[cur] == "{":
                            depth += 1
                        elif expr[cur] == "}":
                            if depth == 0:
                                break
                            depth -= 1
                        elif expr[cur] == "(":
                            paren_depth += 1
                        elif expr[cur] == ")":
                            if paren_depth == 0:
                                break
                            paren_depth -= 1
                        elif expr[cur] == "[":
                            bracket_depth += 1
                        elif expr[cur] == "]":
                            if bracket_depth == 0:
                                break
                            bracket_depth -= 1
                        elif depth == 0 and paren_depth == 0 and bracket_depth == 0 and expr[cur] in "+-=":
                            break
                        cur += 1
                    arg = expr[arg_start:cur].strip()
                if sub_text or sup_text:
                    elements.append(_build_textish_base(cmd_lower, sub_text, sup_text))
                    if arg:
                        elements.append(_build_math_text_run(arg))
                else:
                    elements.append(_inner_function(cmd_lower, arg))
                pos = cur
                continue

            lim_upp_accent = {
                "vec": "→",
                "overrightarrow": "→",
                "hat": "^",
            }.get(cmd_lower)
            if lim_upp_accent is not None:
                parsed_body = _parse_group_or_token_elements(expr, cmd_end)
                if parsed_body is not None:
                    _flush_text_buf(text_buf, elements)
                    text_buf = ""
                    body_elements, cur = parsed_body
                    elements.append(_inner_lim_upp_accent(body_elements, lim_upp_accent))
                    pos = cur
                    continue

            if cmd_lower == "bar":
                parsed_body = _parse_group_or_token_elements(expr, cmd_end)
                if parsed_body is not None:
                    _flush_text_buf(text_buf, elements)
                    text_buf = ""
                    body_elements, cur = parsed_body
                    elements.append(_inner_bar_top(body_elements))
                    pos = cur
                    continue

            if cmd_lower == "overline":
                parsed_body = _parse_group_or_token_elements(expr, cmd_end)
                if parsed_body is not None:
                    _flush_text_buf(text_buf, elements)
                    text_buf = ""
                    body_elements, cur = parsed_body
                    elements.append(_inner_acc_char(body_elements, "―"))
                    pos = cur
                    continue

            styled_spec = _styled_command_spec(cmd_lower)
            if styled_spec is not None:
                parsed_body = _parse_group_or_token_elements(expr, cmd_end)
                if parsed_body is not None:
                    _flush_text_buf(text_buf, elements)
                    text_buf = ""
                    body_elements, cur = parsed_body
                    elements.extend(
                        _apply_math_style_to_elements(
                            body_elements,
                            styled_spec[0],
                            styled_spec[1],
                        )
                    )
                    pos = cur
                    continue

            symbol = _SYMBOL_COMMANDS.get(cmd) or _SYMBOL_COMMANDS.get(cmd_lower)
            if symbol is not None:
                text_buf += symbol
                pos = cmd_end
                continue

            # Unknown command → pass through as text
            text_buf += expr[pos:cmd_end]
            pos = cmd_end
            continue

        # ── Superscript ──
        if ch == "^":
            group, after = _parse_brace_group(expr, pos + 1)
            # Check if the next thing after this is a subscript (for sub_sup)
            base_char, text_buf = _take_script_base(elements, text_buf)

            if after < n and expr[after] == "_":
                sub_group, after2 = _parse_brace_group(expr, after + 1)
                if _has_script_base(base_char):
                    elements.append(_inner_sub_sup(base_char, sub_group, group))
                else:
                    text_buf += "^{" + group + "}_{" + sub_group + "}"
                pos = after2
            else:
                if _has_script_base(base_char):
                    elements.append(_inner_sup(base_char, group))
                else:
                    text_buf += "^{" + group + "}"
                pos = after
            continue

        # ── Subscript ──
        if ch == "_":
            group, after = _parse_brace_group(expr, pos + 1)
            base_char, text_buf = _take_script_base(elements, text_buf)

            if after < n and expr[after] == "^":
                sup_group, after2 = _parse_brace_group(expr, after + 1)
                if _has_script_base(base_char):
                    elements.append(_inner_sub_sup(base_char, group, sup_group))
                else:
                    text_buf += "_{" + group + "}^{" + sup_group + "}"
                pos = after2
            else:
                if _has_script_base(base_char):
                    elements.append(_inner_sub(base_char, group))
                else:
                    text_buf += "_{" + group + "}"
                pos = after
            continue

        # ── Regular character ──
        text_buf += ch
        pos += 1

    _flush_text_buf(text_buf, elements)
    return elements


def _wrap_math_elements(elements: list[etree._Element], *, block: bool) -> etree._Element:
    """Wrap multiple inner elements in a single oMath/oMathPara."""
    if block:
        root = etree.Element(f"{{{_M_NS}}}oMathPara")
        omath = etree.SubElement(root, f"{{{_M_NS}}}oMath")
    else:
        root = etree.Element(f"{{{_M_NS}}}oMath")
        omath = root
    for el in elements:
        omath.append(el)
    return root


def _build_compound_omml(expr: str, *, block: bool) -> etree._Element:
    """Tokenize a compound LaTeX expression and build structured OMML."""
    elements = _tokenize_latex(expr)
    if not elements:
        return _build_literal_omml(expr, block=block)
    return _wrap_math_elements(elements, block=block)


def _parse_matrix_rows(body: str) -> list[list[str]]:
    rows: list[list[str]] = []
    for row_text in re.split(r"\\\\", str(body or "").strip()):
        if not row_text.strip():
            continue
        cols = [col.strip() for col in row_text.split("&")]
        if cols:
            rows.append(cols)
    return rows


def _shape_from_latex(expr: str) -> tuple[str, dict]:
    raw = str(expr or "").strip()
    if not raw:
        return "empty", {}

    m = _RE_MATRIX.fullmatch(raw)
    if m:
        rows = _parse_matrix_rows(m.group("body"))
        if rows:
            return "matrix", {"env": m.group("env"), "rows": rows}

    m = _RE_ALIGNED.fullmatch(raw)
    if m:
        rows = _parse_matrix_rows(m.group("body"))
        if rows:
            return "matrix", {"env": m.group("env"), "rows": rows}

    m = _RE_DELIMITED.fullmatch(raw)
    if m:
        left = _decode_latex_delim(m.group("left"))
        right = _decode_latex_delim(m.group("right"))
        body = m.group("body").strip()
        return "delimited", {"left": left, "right": right, "body": body}

    if _has_top_level_operator(raw) and (raw.startswith("\\") or any(ch in raw for ch in "^_")):
        return "compound", {"expr": raw}

    parsed_fraction = _parse_single_fraction(raw)
    if parsed_fraction is not None:
        payload, shape = parsed_fraction
        return shape, payload

    parsed_sqrt = _parse_single_sqrt(raw)
    if parsed_sqrt is not None:
        payload, shape = parsed_sqrt
        return shape, payload

    m = _RE_SUB_SUP_A.fullmatch(raw) or _RE_SUB_SUP_B.fullmatch(raw)
    if m:
        return "sub_sup", {
            "base": m.group("base").strip(),
            "sub": m.group("sub").strip(),
            "sup": m.group("sup").strip(),
        }

    m = _RE_SUP.fullmatch(raw)
    if m:
        return "superscript", {"base": m.group("base").strip(), "sup": m.group("sup").strip()}

    m = _RE_SUB.fullmatch(raw)
    if m:
        return "subscript", {"base": m.group("base").strip(), "sub": m.group("sub").strip()}

    m = _RE_BIG_OPERATOR.fullmatch(raw)
    if m:
        op = m.group("op").strip()
        sub = (m.group("sub") or "").strip()
        sup = (m.group("sup") or "").strip()
        body = (m.group("body") or "").strip()
        if body and (body.startswith("\\") or _has_top_level_operator(body) or _RE_COMPOUND_MARKER.search(body)):
            return "compound", {"expr": raw}
        return "big_operator", {"operator": op, "sub": sub, "sup": sup, "body": body}

    m = _RE_FUNCTION.fullmatch(raw)
    if m:
        return "function", {"name": m.group("fn").strip(), "arg": m.group("arg").strip()}

    if _RE_COMPOUND_MARKER.search(raw):
        return "compound", {"expr": raw}

    if _RE_SIMPLE_LATEX.fullmatch(raw):
        return "text", {"text": raw}

    return "unsupported", {"text": raw}


def _latex_from_shape(shape: str, payload: dict) -> str:
    shape_name = str(shape or "").strip().lower()
    if shape_name == "fraction":
        num = str(payload.get("num", "")).strip()
        den = str(payload.get("den", "")).strip()
        if num and den:
            return f"\\frac{{{num}}}{{{den}}}"
    if shape_name == "sqrt":
        body = str(payload.get("body", "")).strip()
        if body:
            return f"\\sqrt{{{body}}}"
    if shape_name == "superscript":
        base = str(payload.get("base", "")).strip()
        sup = str(payload.get("sup", "")).strip()
        if base and sup:
            return f"{base}^{{{sup}}}"
    if shape_name == "subscript":
        base = str(payload.get("base", "")).strip()
        sub = str(payload.get("sub", "")).strip()
        if base and sub:
            return f"{base}_{{{sub}}}"
    if shape_name == "sub_sup":
        base = str(payload.get("base", "")).strip()
        sub = str(payload.get("sub", "")).strip()
        sup = str(payload.get("sup", "")).strip()
        if base and sub and sup:
            return f"{base}_{{{sub}}}^{{{sup}}}"
    if shape_name == "prescript":
        base = str(payload.get("base", "")).strip()
        sub = str(payload.get("sub", "")).strip()
        sup = str(payload.get("sup", "")).strip()
        if base and (sub or sup):
            return f"\\prescript{{{sup}}}{{{sub}}}{{{base}}}"
    if shape_name == "big_operator":
        op = str(payload.get("operator", "sum")).strip()
        cmd = f"\\{op}" if op else "\\sum"
        sub = str(payload.get("sub", "")).strip()
        sup = str(payload.get("sup", "")).strip()
        body = str(payload.get("body", "")).strip()
        tail = ""
        if sub:
            tail += f"_{{{sub}}}"
        if sup:
            tail += f"^{{{sup}}}"
        if body:
            tail += f" {body}"
        return f"{cmd}{tail}".strip()
    if shape_name == "function":
        name = str(payload.get("name", "f")).strip()
        arg = str(payload.get("arg", "")).strip()
        if arg:
            return f"\\{name} {arg}"
        return f"\\{name}"
    if shape_name == "delimited":
        left = _encode_latex_delim(str(payload.get("left", "(")).strip())
        right = _encode_latex_delim(str(payload.get("right", ")")).strip())
        body = str(payload.get("body", "")).strip()
        if body:
            return f"\\left{left} {body} \\right{right}"
    if shape_name == "matrix":
        rows = payload.get("rows")
        env = str(payload.get("env", "matrix")).strip() or "matrix"
        if isinstance(rows, list) and rows:
            row_chunks = []
            for row in rows:
                if isinstance(row, list):
                    row_chunks.append(" & ".join(str(col or "").strip() for col in row))
            if row_chunks:
                body = " \\\\ ".join(row_chunks)
                return f"\\begin{{{env}}}{body}\\end{{{env}}}"
    if shape_name == "text":
        return str(payload.get("text", "")).strip()
    return ""


def _resolve_source_expression(node: FormulaNode) -> str:
    payload = node.payload or {}
    source = str(node.source_type or "").strip().lower()
    if source == "latex":
        return str(payload.get("latex", "")).strip()
    if source in _SOURCE_USING_LATEX_CANDIDATE:
        return str(
            payload.get("latex")
            or payload.get("normalized_latex")
            or payload.get("text")
            or payload.get("linear_text")
            or ""
        ).strip()
    return ""


def _should_prefer_mathml_primary(node: FormulaNode, *, block: bool) -> bool:
    if not block:
        return False
    payload = node.payload if isinstance(node.payload, dict) else {}
    latex_source = str(payload.get("latex_source", "") or "").strip().lower()
    return latex_source == "multiline_display_block"


def convert_formula_node(node: FormulaNode, target_mode: str, *, block: bool = True) -> ConversionOutcome:
    """Convert one AST node to the target output mode."""

    mode = str(target_mode or "word_native").strip().lower()
    if mode not in {"word_native", "latex"}:
        mode = "word_native"

    normalized = normalize_formula_node(node)
    source = str(normalized.source_type or "").strip().lower()
    confidence = _clamp(normalized.confidence)
    warnings = list(normalized.warnings or [])

    if mode == "word_native":
        if source == "word_native":
            # Keep native OMML untouched for high-fidelity no-op.
            return ConversionOutcome(
                success=True,
                target_mode=mode,
                confidence=max(confidence, 0.85),
                reason="already_word_native",
                warnings=warnings,
                transformed=False,
            )

        if source in _SOURCE_USING_LATEX_CANDIDATE:
            expr = _resolve_source_expression(normalized)
            if not expr:
                return ConversionOutcome(
                    success=False,
                    target_mode=mode,
                    confidence=0.0,
                    reason="empty_formula_source",
                    warnings=warnings + ["empty_formula_source"],
                )

            if _should_prefer_mathml_primary(normalized, block=block):
                mathml_omml = _convert_latex_via_mathml(expr, block=block)
                if mathml_omml is not None and not _omml_contains_latex_literal(mathml_omml):
                    return ConversionOutcome(
                        success=True,
                        target_mode=mode,
                        confidence=_clamp(max(confidence, 0.90)),
                        reason="latex_to_word_mathml_primary",
                        warnings=warnings + ["mathml_omml_primary"],
                        transformed=True,
                        omml_element=mathml_omml,
                    )

            shape, shape_payload = _shape_from_latex(expr)
            if shape == "empty":
                return ConversionOutcome(
                    success=False,
                    target_mode=mode,
                    confidence=0.0,
                    reason="empty_formula_source",
                    warnings=warnings + ["empty_formula_source"],
                )

            if shape == "fraction":
                omml = _build_fraction_omml(shape_payload.get("num", ""), shape_payload.get("den", ""), block=block)
                conf = 0.90
                reason = "latex_to_word_fraction"
            elif shape == "sqrt":
                omml = _build_sqrt_omml(shape_payload.get("body", ""), block=block)
                conf = 0.88
                reason = "latex_to_word_sqrt"
            elif shape == "superscript":
                omml = _build_sup_omml(shape_payload.get("base", ""), shape_payload.get("sup", ""), block=block)
                conf = 0.87
                reason = "latex_to_word_superscript"
            elif shape == "subscript":
                omml = _build_sub_omml(shape_payload.get("base", ""), shape_payload.get("sub", ""), block=block)
                conf = 0.87
                reason = "latex_to_word_subscript"
            elif shape == "sub_sup":
                omml = _build_sub_sup_omml(
                    shape_payload.get("base", ""),
                    shape_payload.get("sub", ""),
                    shape_payload.get("sup", ""),
                    block=block,
                )
                conf = 0.86
                reason = "latex_to_word_subsup"
            elif shape == "big_operator":
                omml = _build_big_operator_omml(
                    shape_payload.get("operator", ""),
                    shape_payload.get("sub", ""),
                    shape_payload.get("sup", ""),
                    shape_payload.get("body", ""),
                    block=block,
                )
                conf = 0.82
                reason = "latex_to_word_big_operator"
            elif shape == "matrix":
                omml = _build_matrix_omml(shape_payload.get("rows", []), block=block)
                conf = 0.80
                reason = "latex_to_word_matrix"
            elif shape == "delimited":
                omml = _build_delimited_omml(
                    shape_payload.get("left", "("),
                    shape_payload.get("right", ")"),
                    shape_payload.get("body", ""),
                    block=block,
                )
                conf = 0.80
                reason = "latex_to_word_delimited"
            elif shape == "function":
                omml = _build_function_omml(
                    shape_payload.get("name", "f"),
                    shape_payload.get("arg", ""),
                    block=block,
                )
                conf = 0.80
                reason = "latex_to_word_function"
            elif shape == "compound":
                omml = _build_compound_omml(shape_payload.get("expr", expr), block=block)
                conf = 0.85 if source != "ocr_fragment" else 0.68
                reason = "latex_to_word_compound"
            elif shape == "text":
                omml = _build_literal_omml(expr, block=block)
                conf = 0.78 if source != "ocr_fragment" else 0.60
                reason = "text_to_word_literal"
            else:
                # Keep conversion path available and mark fallback in warnings.
                omml = _build_literal_omml(expr, block=block)
                conf = 0.66 if source != "ocr_fragment" else 0.55
                reason = "fallback_literal_conversion"
                warnings = warnings + ["fallback_literal_conversion"]

            if _omml_contains_latex_literal(omml):
                mathml_omml = _convert_latex_via_mathml(expr, block=block)
                if mathml_omml is not None and not _omml_contains_latex_literal(mathml_omml):
                    omml = mathml_omml
                    conf = max(conf, confidence)
                    reason = "latex_to_word_mathml"
                    warnings = warnings + ["mathml_omml_fallback"]

            return ConversionOutcome(
                success=True,
                target_mode=mode,
                confidence=_clamp(conf),
                reason=reason,
                warnings=warnings,
                transformed=True,
                omml_element=omml,
            )

        return ConversionOutcome(
            success=False,
            target_mode=mode,
            confidence=0.0,
            reason=f"unsupported_source:{source or 'unknown'}",
            warnings=warnings,
        )

    # mode == latex
    if source == "latex":
        latex = _resolve_source_expression(normalized)
        return ConversionOutcome(
            success=True,
            target_mode=mode,
            confidence=confidence if latex else 0.0,
            reason="already_latex",
            warnings=warnings,
            transformed=False,
            latex_text=latex,
        )

    if source == "word_native":
        shape = str(normalized.payload.get("shape", "text")).strip().lower()
        shape_payload = normalized.payload.get("shape_payload", {})
        if isinstance(shape_payload, dict):
            latex = _latex_from_shape(shape, shape_payload)
            if latex:
                conf = 0.90 if shape not in {"text", "unknown"} else 0.84
                return ConversionOutcome(
                    success=True,
                    target_mode=mode,
                    confidence=conf,
                    reason="word_to_latex_structured",
                    warnings=warnings,
                    transformed=True,
                    latex_text=latex,
                )

        linear = str(normalized.payload.get("linear_text", "")).strip()
        if linear:
            simple_omml = bool(normalized.payload.get("is_simple_omml", False))
            return ConversionOutcome(
                success=True,
                target_mode=mode,
                confidence=0.88 if simple_omml else 0.72,
                reason="word_to_latex_linearized",
                warnings=warnings + ([] if simple_omml else ["complex_omml_linearized"]),
                transformed=True,
                latex_text=linear,
            )
        return ConversionOutcome(
            success=False,
            target_mode=mode,
            confidence=0.30,
            reason="missing_word_formula_text",
            warnings=warnings + ["missing_word_formula_text"],
        )

    if source in _SOURCE_USING_LATEX_CANDIDATE:
        expr = _resolve_source_expression(normalized)
        if not expr:
            return ConversionOutcome(
                success=False,
                target_mode=mode,
                confidence=0.0,
                reason="empty_formula_source",
                warnings=warnings + ["empty_formula_source"],
            )
        shape, payload = _shape_from_latex(expr)
        if shape == "compound":
            latex = payload.get("expr", expr)
            conf = 0.80 if source != "ocr_fragment" else 0.56
            reason = "normalized_source_to_latex"
        elif shape != "unsupported":
            latex = _latex_from_shape(shape, payload) or expr
            conf = 0.80 if source != "ocr_fragment" else 0.56
            reason = "normalized_source_to_latex"
        else:
            latex = expr
            conf = 0.64 if source != "ocr_fragment" else 0.52
            reason = "source_literal_to_latex"
            warnings = warnings + ["literal_latex_passthrough"]

        return ConversionOutcome(
            success=True,
            target_mode=mode,
            confidence=conf,
            reason=reason,
            warnings=warnings,
            transformed=True,
            latex_text=latex,
        )

    return ConversionOutcome(
        success=False,
        target_mode=mode,
        confidence=0.0,
        reason=f"unsupported_source:{source or 'unknown'}",
        warnings=warnings,
    )
