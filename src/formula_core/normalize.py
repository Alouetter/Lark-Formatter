"""Formula AST normalization and text-to-LaTeX helpers."""

from __future__ import annotations

import re
from copy import deepcopy

from .ast import FormulaNode
from .repair import repair_formula_text
from .semantics import BIG_OPERATOR_NAMES, FUNCTION_NAMES
from .symbols import (
    UNICODE_GREEK_TO_LATEX,
    UNICODE_OPERATOR_TO_LATEX,
    ensure_latex_command_word_boundaries,
)
from src.utils.toc_entry import looks_like_bibliographic_reference_line

_RE_LATEX_COMMAND = re.compile(r"\\[A-Za-z]+\b")
_RE_CAPTION_PREFIX = re.compile(
    r"^\s*(?:图|表|Figure|Fig\.?|Table)\s*"
    r"(?:"
    r"[一二三四五六七八九十百零]+"
    r"|[（(]\d+(?:[.\-]\d+)*[）)]"
    r"|\d+(?:[.\-]\d+)*"
    r")"
    r"(?=\s|[（(A-Za-z\u4e00-\u9fff]|$)",
    re.IGNORECASE,
)
_RE_MATH_SYMBOL = re.compile(r"[=+\-*/^_→←↔∑∫∏√∞≤≥≠≈±∂∇∥]")
_RE_ALPHA = re.compile(r"[A-Za-zΑ-Ωα-ωϵϑϖϱϕ]")
_RE_CHINESE = re.compile(r"[\u4e00-\u9fff]")
_RE_PLAIN_SNAKE_IDENTIFIER = re.compile(r"^[A-Za-z]+(?:_[A-Za-z]+)+$")
_RE_OCR_SPACED_POWER = re.compile(
    r"\b([A-Za-z])\s+([0-9])\b(?:\s*[+\-]\s*\b([A-Za-z])\s+([0-9])\b)*"
)
_RE_BARE_CARET = re.compile(r"\^([A-Za-z0-9])(?![A-Za-z0-9{])")
_RE_BARE_UNDERSCORE = re.compile(r"_([A-Za-z0-9])(?![A-Za-z0-9{])")
_RE_LOST_SUPERSCRIPT = re.compile(r"(?<=[A-Za-z)])([0-9])(?![0-9A-Za-z])")
_RE_TRAILING_CN_NOTE = re.compile(r"\s*[（(](?=[^（）()]{1,24}[）)])(?=[^（）()]*[\u4e00-\u9fff])[^（）()]{1,24}[）)]\s*$")
_RE_INVISIBLE_FORMULA_NOISE = re.compile(r"[\u200b\u200c\u200d\ufeff]")
_RE_ESCAPE_PLACEHOLDER = re.compile(
    r"(?:__)?ESC_?(?P<hex>[0-9A-Fa-f]{8})(?:__)?(?P<tail>[A-Za-z]+)"
)
_RE_ESCAPE_ARTIFACT = re.compile(r"(?:__)?ESC_?[0-9A-Fa-f]{8}(?:__)?")
_RE_LATEX_COMMAND_NAME = re.compile(r"\\([A-Za-z]+)")
_BIG_OPERATOR_PATTERN = "|".join(
    re.escape(name)
    for name in sorted(BIG_OPERATOR_NAMES, key=len, reverse=True)
)

_SUP_DICT = {
    "⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4",
    "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9",
    "⁺": "+", "⁻": "-", "⁼": "=", "⁽": "(", "⁾": ")",
    "ⁿ": "n", "ᵢ": "i",
}
_SUB_DICT = {
    "₀": "0", "₁": "1", "₂": "2", "₃": "3", "₄": "4",
    "₅": "5", "₆": "6", "₇": "7", "₈": "8", "₉": "9",
    "₊": "+", "₋": "-", "₌": "=", "₍": "(", "₎": ")",
    "ᵢ": "i",
}
_SUP_MAP = str.maketrans(_SUP_DICT)
_SUB_MAP = str.maketrans(_SUB_DICT)
_SUP_CHARS = "".join(_SUP_DICT.keys())
_SUB_CHARS = "".join(_SUB_DICT.keys())

_GREEK_TO_LATEX = dict(UNICODE_GREEK_TO_LATEX)

_ESCAPE_TAIL_DIRECT_MAP = {
    "rac": "frac",
    "qrt": "sqrt",
    "um": "sum",
    "nt": "int",
    "rod": "prod",
    "og": "log",
    "n": "ln",
    "xp": "exp",
    "inh": "sinh",
    "osh": "cosh",
    "anh": "tanh",
    "sc": "csc",
    "ot": "cot",
    "egin": "begin",
    "nd": "end",
    "artial": "partial",
    "abla": "nabla",
    "igma": "sigma",
    "u": "mu",
    "et": "det",
    "dot": "cdot",
    "anfty": "infty",
    "nfty": "infty",
    "m": "pm",
    "athbb": "mathbb",
    "athcal": "mathcal",
    "athfrak": "mathfrak",
    "athscr": "mathscr",
    "athsf": "mathsf",
    "athtt": "mathtt",
}

_REPAIRED_COMMAND_GROUPS = {
    "frac": 2,
    "sqrt": 1,
    "vec": 1,
    "begin": 1,
    "end": 1,
    "operatorname": 1,
    "mathbb": 1,
    "mathcal": 1,
    "mathfrak": 1,
    "mathscr": 1,
    "mathsf": 1,
    "mathtt": 1,
}


def _normalize_doubled_backslashes(text: str) -> str:
    """Normalize doubled backslashes in LaTeX source text.

    Word paragraph text often stores ``\\\\frac`` instead of ``\\frac``.
    This collapses ``\\\\ <command>`` → ``\\<command>`` while preserving
    the LaTeX line-break ``\\\\`` used inside matrices and cases.
    """
    value = str(text or "")
    if "\\\\" not in value:
        return value
    # Collapse \\<alpha> → \<alpha>  (doubled-backslash before command)
    value = re.sub(r"\\\\([A-Za-z])", r"\\\1", value)
    # Collapse \\, → \, and \\; → \; etc. (LaTeX spacing commands)
    value = re.sub(r"\\\\([,;:!])", r"\\\1", value)
    return value


def _latexify_plain_function_names(text: str) -> str:
    """Convert plain-text function names to LaTeX commands.

    Examples: ``sin x`` → ``\\sin x``, ``sinx`` → ``\\sin x``,
    ``cos(x)`` → ``\\cos(x)``.
    Only replaces when the name is not already preceded by a backslash.
    """
    _FUNC_NAMES = tuple(sorted(FUNCTION_NAMES, key=len, reverse=True))
    value = str(text or "")
    for fn in _FUNC_NAMES:
        # Match word-boundary function name NOT preceded by backslash
        # Group 1 captures an optional directly-following letter (e.g. sinx)
        _pat = re.compile(r"(?<!\\)\b" + fn + r"(?=[\s({[^_]|([A-Za-z])|$)")
        def _repl(m, _fn=fn):
            if m.group(1):
                # Letter follows directly: sinx → \sin x
                return "\\" + _fn + " "
            return "\\" + _fn
        value = _pat.sub(_repl, value)
    return value


def _clamp(value: float) -> float:
    try:
        n = float(value)
    except (TypeError, ValueError):
        return 0.0
    if n < 0.0:
        return 0.0
    if n > 1.0:
        return 1.0
    return n


def looks_like_caption_text(text: str) -> bool:
    raw = str(text or "").strip()
    if not raw:
        return False
    return bool(_RE_CAPTION_PREFIX.match(raw))


def _looks_like_plain_snake_identifier(text: str) -> bool:
    raw = str(text or "").strip()
    if not _RE_PLAIN_SNAKE_IDENTIFIER.fullmatch(raw):
        return False
    parts = raw.split("_")
    return all(len(part) > 1 and part.isalpha() for part in parts)


def _standardize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def _strip_invisible_formula_noise(text: str) -> str:
    return _RE_INVISIBLE_FORMULA_NOISE.sub("", str(text or ""))


def _skip_spaces(text: str, pos: int) -> int:
    while pos < len(text) and text[pos].isspace():
        pos += 1
    return pos


def _consume_brace_group(text: str, pos: int) -> int:
    if pos >= len(text) or text[pos] != "{":
        return pos
    depth = 1
    idx = pos + 1
    while idx < len(text) and depth > 0:
        if text[idx] == "{":
            depth += 1
        elif text[idx] == "}":
            depth -= 1
        idx += 1
    return idx if depth == 0 else pos


def _consume_optional_scripts(text: str, pos: int) -> int:
    cur = pos
    while True:
        probe = _skip_spaces(text, cur)
        if probe >= len(text) or text[probe] not in {"^", "_"}:
            return cur
        probe = _skip_spaces(text, probe + 1)
        if probe < len(text) and text[probe] == "{":
            next_pos = _consume_brace_group(text, probe)
            if next_pos == probe:
                return cur
            cur = next_pos
            continue
        if probe < len(text) and (text[probe].isalnum() or text[probe] in {"(", ")"}):
            cur = probe + 1
            continue
        return cur


def _resolve_escape_command_tail(tail: str, following: str) -> str | None:
    key = str(tail or "").strip().lower()
    if not key:
        return None
    if key in _ESCAPE_TAIL_DIRECT_MAP:
        return _ESCAPE_TAIL_DIRECT_MAP[key]
    if key == "im":
        next_char = (following or "").lstrip()[:1]
        return "lim" if next_char in {"", "{", "_", "^"} else "sim"
    if key == "in":
        return "sin"
    if key == "os":
        return "cos"
    if key == "an":
        return "tan"
    if key == "ec":
        next_char = (following or "").lstrip()[:1]
        if next_char == "{":
            return "vec"
        return "sec"
    return None


def _recover_escaped_latex_commands(text: str) -> tuple[str, list[str]]:
    raw = str(text or "")
    if not raw:
        return "", []

    changed = False

    def _replace(match: re.Match[str]) -> str:
        nonlocal changed
        tail = match.group("tail")
        cmd = _resolve_escape_command_tail(tail, raw[match.end():])
        if not cmd:
            return match.group(0)
        changed = True
        return "\\" + cmd

    repaired = _RE_ESCAPE_PLACEHOLDER.sub(_replace, raw)
    warnings: list[str] = []
    if changed:
        warnings.append("escape_placeholder_command_recovered")
    if "ESC" in repaired and _RE_ESCAPE_ARTIFACT.search(repaired):
        warnings.append("escape_placeholder_unresolved")
    return repaired, warnings


def _iter_repaired_command_spans(text: str) -> list[tuple[int, int]]:
    spans: list[tuple[int, int]] = []
    idx = 0
    while idx < len(text):
        if text[idx] != "\\":
            idx += 1
            continue
        match = _RE_LATEX_COMMAND_NAME.match(text, idx)
        if not match:
            idx += 1
            continue
        cmd = match.group(1).strip().lower()
        end = _consume_optional_scripts(text, match.end())
        group_count = _REPAIRED_COMMAND_GROUPS.get(cmd, 0)
        for _ in range(group_count):
            probe = _skip_spaces(text, end)
            next_end = _consume_brace_group(text, probe)
            if next_end == probe:
                break
            end = next_end
        spans.append((idx, max(end, match.end())))
        idx = max(end, match.end())
    return spans


def _keep_repaired_command_prefix(prefix: str) -> str:
    value = _strip_invisible_formula_noise(prefix)
    compact = _standardize_spaces(value)
    if not compact:
        return ""
    if len(compact) > 12:
        return ""
    if compact.count("=") > 1:
        return ""
    if compact[-1] not in {"=", "+", "-", "*", "/", "(", "["}:
        return ""
    return compact


def _looks_like_formula_restart(text: str, idx: int) -> bool:
    candidate = _standardize_spaces(text[idx:])
    if not candidate:
        return False
    head = candidate[:1]
    if not (
        head.isalpha()
        or head in {"(", "[", "\u222b", "\u2211", "\u2202", "\u2207", "\u221a"}
    ):
        return False
    if (
        head.isalpha()
        and len(candidate) >= 2
        and not candidate[1].isalnum()
        and candidate[1] not in {"(", "["}
    ):
        return False
    if (
        head.isalpha()
        and len(candidate) >= 2
        and candidate[1] in {"\u222b", "\u2211", "\u2202", "\u2207", "\u221a"}
    ):
        return False
    eq_pos = candidate.find("=")
    return 0 <= eq_pos <= 40


def _trim_repaired_command_tail(tail: str, *, single_command: bool) -> str:
    value = _strip_invisible_formula_noise(tail)
    if not value:
        return ""
    if single_command:
        probe = value.lstrip()
        if probe and probe[:1] not in {"=", "+", "-", "*", "/", "^", "_", ",", ";", ":"}:
            return ""

    restart_idx: int | None = None
    for idx in range(len(value)):
        if _looks_like_formula_restart(value, idx):
            restart_idx = idx
            break
    if restart_idx is not None:
        value = value[:restart_idx]

    value = _RE_ESCAPE_PLACEHOLDER.sub("", value)
    value = re.sub(r"(?:^|[,\s])ESC(?:$|[,\s])", " ", value)
    return _standardize_spaces(value).rstrip(",;")


def _extract_repaired_command_formula(text: str) -> tuple[str | None, list[str]]:
    value = _strip_invisible_formula_noise(text)
    spans = _iter_repaired_command_spans(value)
    if not spans:
        return None, []

    first_start = spans[0][0]
    last_end = spans[-1][1]
    prefix = _keep_repaired_command_prefix(value[:first_start])
    middle = value[first_start:last_end]
    tail = _trim_repaired_command_tail(
        value[last_end:],
        single_command=len(spans) == 1,
    )
    extracted = f"{prefix}{middle}{tail}".strip()
    if not extracted:
        return None, []
    warnings = ["escape_placeholder_formula_extracted"]
    return extracted, warnings


def _replace_super_sub_chars(text: str) -> str:
    value = str(text or "")
    value = re.sub(
        rf"([A-Za-z0-9\)\]])([{re.escape(_SUP_CHARS)}]+)",
        lambda m: f"{m.group(1)}^{{{m.group(2).translate(_SUP_MAP)}}}",
        value,
    )
    value = re.sub(
        rf"([A-Za-z0-9\)\]])([{re.escape(_SUB_CHARS)}]+)",
        lambda m: f"{m.group(1)}_{{{m.group(2).translate(_SUB_MAP)}}}",
        value,
    )
    return value


def _normalize_bare_scripts(text: str) -> str:
    """Wrap bare ^X / _X with braces: a^2 → a^{2}, x_i → x_{i}."""
    value = str(text or "")
    value = _RE_BARE_CARET.sub(r"^{\1}", value)
    value = _RE_BARE_UNDERSCORE.sub(r"_{\1}", value)
    return value


def _replace_greek_chars(text: str) -> str:
    value = str(text or "")
    for src, dst in _GREEK_TO_LATEX.items():
        value = value.replace(src, dst)
    return ensure_latex_command_word_boundaries(value)


def _normalize_arrow_and_symbols(text: str) -> str:
    value = str(text or "")
    value = value.replace("−", "-")
    for src, dst in UNICODE_OPERATOR_TO_LATEX.items():
        value = value.replace(src, dst)
    value = ensure_latex_command_word_boundaries(value)
    # OCR common errors
    value = value.replace("一", "-")  # Chinese dash → minus
    # \Sigma → \sum (capital Sigma symbol used as summation)
    value = re.sub(r"\\Sigma(?=\s|_|\^|$|\{)", r"\\sum", value)
    return value


def _normalize_implicit_big_operator_subscript(text: str) -> tuple[str, bool]:
    value = str(text or "")
    if not value:
        return "", False

    out: list[str] = []
    pos = 0
    changed = False
    pattern = re.compile(r"\\(" + _BIG_OPERATOR_PATTERN + r")\b")

    while pos < len(value):
        match = pattern.search(value, pos)
        if not match:
            out.append(value[pos:])
            break

        out.append(value[pos:match.start()])
        cmd = match.group(1)
        cur = match.end()
        cur = _skip_spaces(value, cur)
        if cur < len(value) and value[cur] == "{":
            group_end = _consume_brace_group(value, cur)
            if group_end > cur:
                rest = value[group_end:].lstrip()
                if rest:
                    out.append(f"\\{cmd}_{{{value[cur + 1:group_end - 1]}}}")
                    pos = group_end
                    changed = True
                    continue
        out.append(match.group(0))
        pos = match.end()

    return "".join(out), changed


def _should_recover_lost_superscript(value: str) -> bool:
    if _RE_LATEX_COMMAND.search(value):
        return False

    hits = list(_RE_LOST_SUPERSCRIPT.finditer(value))
    if not hits:
        return False

    lower_or_paren_hit = any(
        match.start() > 0 and (value[match.start() - 1].islower() or value[match.start() - 1] == ")")
        for match in hits
    )
    if not lower_or_paren_hit:
        return False

    has_additive_context = any(ch in value for ch in "+-")
    has_paren_power = bool(
        re.search(r"\)[0-9](?=$|[=+\-*/,])", value)
        and re.search(r"\([^)]*[+\-][^)]*\)", value)
    )
    has_polynomial_term = bool(
        re.search(r"(?:^|[=+\-(])(?:[A-Za-z]+)?[A-Za-z][0-9](?=$|[=+\-*/,)])", value)
    )

    if "=" in value:
        if len(hits) >= 2:
            return True
        return has_additive_context and (has_paren_power or has_polynomial_term)

    if len(hits) >= 2 and has_additive_context:
        return True
    if has_paren_power:
        return True
    return has_additive_context and has_polynomial_term and len(hits) >= 1


def _apply_plain_rules(text: str) -> tuple[str, list[str]]:
    value = str(text or "")
    warnings: list[str] = []

    # sum_(i=1)^n => \sum_{i=1}^{n}
    # Handles both sum_(i=1)^n and sum_(i=1)^{n} patterns.
    value = re.sub(
        r"(?<!\\)\bsum_\(\s*([^)]+?)\s*\)\s*\^\s*\{?([A-Za-z0-9]+)\}?",
        lambda m: rf"\sum_{{{m.group(1).strip()}}}^{{{m.group(2).strip()}}}",
        value,
        flags=re.IGNORECASE,
    )
    # lim x->0 => \lim_{x\to0}
    # Handles: lim x->0, lim x - > 0, lim x-->0, lim x \to 0
    value = re.sub(
        r"(?<!\\)\blim\s+([A-Za-z][A-Za-z0-9]*)\s*(?:-\s*-?\s*>|\\to)\s*([A-Za-z0-9+\-]+)",
        lambda m: rf"\lim_{{{m.group(1)}\to{m.group(2)}}}",
        value,
        flags=re.IGNORECASE,
    )
    # sqrt x => \sqrt{x}
    value = re.sub(
        r"\\sqrt\s+([A-Za-z0-9][A-Za-z0-9+\-*/^_()]*)",
        lambda m: rf"\sqrt{{{m.group(1)}}}",
        value,
    )
    # Plain text function names → LaTeX commands
    value = _latexify_plain_function_names(value)

    value, changed_big_op = _normalize_implicit_big_operator_subscript(value)
    if changed_big_op:
        warnings.append("implicit_big_operator_subscript_normalized")

    # OCR-like spaced superscript pattern: x 2 + y 2 = z 2
    if _RE_OCR_SPACED_POWER.search(value):
        value = re.sub(
            r"\b([A-Za-z])\s+([0-9])\b",
            lambda m: f"{m.group(1)}^{{{m.group(2)}}}",
            value,
        )
        warnings.append("ocr_spaced_power_normalized")

    # Lost superscript recovery: a2 → a^{2}, )2 → )^{2}
    # Applies when copy/paste stripped the caret entirely, leaving digits
    # directly adjacent to letters or closing parens.  Requires '=' and at
    # least 2 matches to avoid false positives on variable names like "x2".
    if _should_recover_lost_superscript(value):
        value = _RE_LOST_SUPERSCRIPT.sub(r"^{\1}", value)
        warnings.append("lost_superscript_recovered")

    value = re.sub(r"\bei\\pi\b", r"e^{i\\pi}", value)
    value = re.sub(r"\be([A-Za-z])=(?=\\sum\b)", r"e^{\1}=", value)

    # Remove trailing explanation note like "(上标丢失)".
    trimmed = _RE_TRAILING_CN_NOTE.sub("", value).strip()
    if trimmed != value:
        value = trimmed
        warnings.append("trimmed_trailing_note")

    return value, warnings


def looks_like_bibliographic_reference_text(text: str) -> bool:
    raw = str(text or "").strip()
    if not raw:
        return False
    return looks_like_bibliographic_reference_line(raw)


def looks_like_formula_text(text: str) -> tuple[bool, float, str]:
    raw = str(text or "").strip()
    if not raw:
        return False, 0.0, "plain_text"

    if looks_like_caption_text(raw):
        return False, 0.0, "plain_text"

    if looks_like_bibliographic_reference_text(raw):
        return False, 0.0, "plain_text"

    if _RE_LATEX_COMMAND.search(raw):
        return True, 0.86, "plain_text"

    if _RE_OCR_SPACED_POWER.search(raw) or "OCR" in raw.upper():
        return True, 0.52, "ocr_fragment"

    # Reject plain snake_case / SCREAMING_SNAKE identifiers such as BODY_KEEP.
    # Legitimate formulas like x_i, A_n, foo_1 are preserved because they contain
    # single-letter or numeric segments instead of prose-like word chunks.
    if _looks_like_plain_snake_identifier(raw):
        return False, 0.0, "plain_text"

    symbol_hits = len(_RE_MATH_SYMBOL.findall(raw))
    alpha_hits = len(_RE_ALPHA.findall(raw))
    if symbol_hits >= 2 and alpha_hits >= 2:
        return True, 0.72, "plain_text"

    if symbol_hits >= 1 and alpha_hits >= 1 and any(ch in raw for ch in ("^", "_", "=", "->", "√", "∫", "∑")):
        return True, 0.66, "plain_text"

    if any(ch in raw for ch in _GREEK_TO_LATEX):
        return True, 0.62, "plain_text"

    if _should_recover_lost_superscript(raw):
        return True, 0.66, "plain_text"

    repair = repair_formula_text(raw)
    repaired = str(repair.text or "").strip()
    if repaired and repaired != raw:
        if _RE_LATEX_COMMAND.search(repaired):
            return True, max(0.72, float(repair.confidence)), "plain_text"
        repaired_symbol_hits = len(_RE_MATH_SYMBOL.findall(repaired))
        repaired_alpha_hits = len(_RE_ALPHA.findall(repaired))
        if repaired_symbol_hits >= 1 and repaired_alpha_hits >= 1:
            return True, max(0.62, float(repair.confidence)), "plain_text"

    return False, 0.0, "plain_text"


def text_formula_to_latex(text: str, *, source_hint: str = "plain_text") -> tuple[str, float, list[str]]:
    raw = str(text or "").strip()
    if not raw:
        return "", 0.0, ["empty_formula_text"]

    warnings: list[str] = []
    confidence = 0.64 if source_hint != "ocr_fragment" else 0.52
    repair = repair_formula_text(raw)
    expr = str(repair.text or raw)
    warnings.extend(repair.warnings)
    if repair.confidence > 0:
        confidence = max(confidence, float(repair.confidence))

    expr = _normalize_arrow_and_symbols(expr)
    expr = _replace_super_sub_chars(expr)
    expr = _replace_greek_chars(expr)
    expr = re.sub(r"\\Sigma(?=\s|_|\^|$|\{)", r"\\sum", expr)
    expr = _normalize_bare_scripts(expr)
    expr, plain_warnings = _apply_plain_rules(expr)
    warnings.extend(plain_warnings)

    expr = _standardize_spaces(expr)

    if _RE_LATEX_COMMAND.search(expr):
        confidence += 0.18
    if _RE_MATH_SYMBOL.search(expr):
        confidence += 0.08
    if "escape_placeholder_formula_extracted" in warnings:
        confidence += 0.08
    if "escape_placeholder_unresolved" in warnings:
        confidence -= 0.18
    if _RE_CHINESE.search(expr):
        confidence -= 0.18
        warnings.append("contains_non_formula_text")

    if not _RE_ALPHA.search(expr) and not _RE_MATH_SYMBOL.search(expr) and not _RE_LATEX_COMMAND.search(expr):
        return "", 0.35, warnings + ["unable_to_structure_formula"]

    return expr, _clamp(confidence), warnings


def normalize_formula_node(node: FormulaNode) -> FormulaNode:
    normalized = deepcopy(node)
    payload = dict(normalized.payload or {})
    warnings = list(normalized.warnings or [])
    source = str(normalized.source_type or "").strip().lower()

    if source == "latex":
        latex = _standardize_spaces(str(payload.get("latex", "")).strip())
        if latex:
            latex = _normalize_doubled_backslashes(latex)
            latex = ensure_latex_command_word_boundaries(latex)
            latex = _normalize_bare_scripts(latex)
            payload["latex"] = latex
    elif source in {"plain_text", "ocr_fragment", "unicode_text", "old_equation", "mathtype", "ole_equation"}:
        seed = str(
            payload.get("latex")
            or payload.get("normalized_latex")
            or payload.get("text")
            or payload.get("linear_text")
            or ""
        ).strip()
        if seed:
            latex, conf, extra = text_formula_to_latex(seed, source_hint=source)
            if latex:
                payload["latex"] = latex
                payload["normalized_latex"] = latex
            warnings.extend(extra)
            if conf > 0:
                normalized.confidence = _clamp((normalized.confidence + conf) / 2.0)

    normalized.payload = payload
    normalized.warnings = list(dict.fromkeys(warnings))
    return normalized
