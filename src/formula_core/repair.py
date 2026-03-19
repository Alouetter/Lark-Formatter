"""Formula repair engine for polluted copied formula text."""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field

from .escape_features import (
    EscapeFeatureProfile,
    _RE_ESCAPE_ARTIFACT,
    _RE_ESCAPE_PLACEHOLDER,
    build_escape_feature_profile,
    recover_escaped_latex_commands,
)
from .symbols import (
    UNICODE_GREEK_TO_LATEX,
    UNICODE_OPERATOR_TO_LATEX,
    ensure_latex_command_word_boundaries,
)

_RE_INVISIBLE_FORMULA_NOISE = re.compile(r"[\u200b\u200c\u200d\ufeff]")
_RE_LATEX_COMMAND_NAME = re.compile(r"\\([A-Za-z]+)")
_RE_LATEX_COMMAND = re.compile(r"\\[A-Za-z]+")
_RE_FORMULA_OPERATOR = re.compile(r"(=|[+\-*/^_]|→|←|↔|∞|≤|≥|≠|≈|±|∩|∪|×|·|∑|∫|∏|√|∂|∇|∥)")
_RE_LETTER = re.compile(r"[A-Za-zΑ-Ωα-ωϵϑϖϱϕ]")

_REPAIRED_COMMAND_GROUPS = {
    "frac": 2,
    "sqrt": 1,
    "vec": 1,
    "bar": 1,
    "hat": 1,
    "begin": 1,
    "end": 1,
    "text": 1,
    "operatorname": 1,
    "mathbb": 1,
    "mathcal": 1,
    "mathfrak": 1,
    "mathscr": 1,
    "mathsf": 1,
    "mathtt": 1,
    "binom": 2,
    "overset": 2,
    "underline": 1,
    "widetilde": 1,
    "overrightarrow": 1,
    "overleftarrow": 1,
    "overline": 1,
    "underset": 2,
    "tilde": 1,
}

_ROW_SPLIT_ENVS = {"matrix", "bmatrix", "pmatrix", "vmatrix", "Vmatrix", "cases", "aligned"}
_BINARY_OPERATOR_COMMANDS = {
    "cap", "cup", "subset", "subseteq", "setminus",
    "land", "lor", "Rightarrow", "Leftrightarrow",
    "times", "cdot", "to", "ge", "le",
}
_PREFIX_OPERATOR_COMMANDS = {"forall", "exists", "neg"}
_FUNCTION_LIKE_COMMANDS = {
    "sin", "cos", "tan", "sec", "csc", "cot", "sinh", "cosh", "tanh",
    "log", "ln", "exp", "max", "min",
    "det", "arcsin", "arccos", "arctan",
}
_SYMBOL_CONTEXT_COMMANDS = {
    str(value).strip()[1:].lower()
    for value in UNICODE_GREEK_TO_LATEX.values()
    if str(value).strip().startswith("\\")
} | {"nabla", "partial", "hbar"}
_BIG_OPERATOR_COMMANDS = {"sum", "int", "iint", "iiint", "oint", "prod", "lim"}
_RENDERED_HINT_CHARS = "".join(
    sorted(
        {
            "\u2061",
            *UNICODE_OPERATOR_TO_LATEX.keys(),
            *UNICODE_GREEK_TO_LATEX.keys(),
        }
    )
)
_RE_RENDERED_MATH_HINT = re.compile(r"[" + re.escape(_RENDERED_HINT_CHARS) + r"]")

_RENDERED_SYMBOL_TO_LATEX = {
    "\u2061": " ",
    "−": "-",
    **UNICODE_OPERATOR_TO_LATEX,
}
_RENDERED_GREEK_TO_LATEX = dict(UNICODE_GREEK_TO_LATEX)
_FUNCTION_NAME_PATTERN = "|".join(
    re.escape(name)
    for name in sorted(_FUNCTION_LIKE_COMMANDS, key=len, reverse=True)
)
_EXACT_UNICODE_COMMANDS = {
    str(value).strip()[1:]
    for value in (*UNICODE_GREEK_TO_LATEX.values(), *UNICODE_OPERATOR_TO_LATEX.values())
    if str(value).strip().startswith("\\")
}
_RE_RENDERED_FUNCTION_NAME = re.compile(
    r"(?<!\\)\b(" + _FUNCTION_NAME_PATTERN + r")\b"
)
_RE_RENDERED_FUNCTION_QUESTION = re.compile(
    r"(\\(?:" + _FUNCTION_NAME_PATTERN + r"))\?\s*"
)
_RE_RENDERED_QUESTION_HINT = re.compile(
    r"(?:" + _FUNCTION_NAME_PATTERN + r")\?"
)
_RE_LOG_BASE_PREFIX = re.compile(r"\\log\s*([A-Za-z])([A-Za-z])\s*=")
_SIMPLE_RENDERED_COMMANDS = (
    _FUNCTION_LIKE_COMMANDS
    | _SYMBOL_CONTEXT_COMMANDS
    | _BIG_OPERATOR_COMMANDS
    | _EXACT_UNICODE_COMMANDS
    | {
        "sqrt",
        "infty",
        "to",
        "leftarrow",
        "leftrightarrow",
        "Rightarrow",
        "Leftarrow",
        "Leftrightarrow",
        "times",
        "cdot",
        "cap",
        "cup",
        "subset",
        "subseteq",
        "setminus",
        "in",
        "forall",
        "exists",
        "neg",
        "land",
        "lor",
        "Vert",
        "ge",
        "le",
        "pm",
        "ne",
        "approx",
    }
)
_RE_SIMPLE_RENDERED_COMMAND = re.compile(
    r"\\(?:"
    + "|".join(re.escape(name) for name in sorted(_SIMPLE_RENDERED_COMMANDS, key=len, reverse=True))
    + r")\b"
)
_RE_TRAILING_COMMAND_ONLY = re.compile(r"\\[A-Za-z]+$")
_RE_TRAILING_FUNCTION_LIKE = re.compile(
    r"\\(?:" + _FUNCTION_NAME_PATTERN + r")"
    r"(?:_(?:\{[^{}]+\}|[A-Za-z0-9])|\^(?:\{[^{}]+\}|[A-Za-z0-9]))*\s*$"
)


@dataclass
class RepairCandidate:
    text: str
    source: str
    score: float = 0.0
    warnings: list[str] = field(default_factory=list)
    metadata: dict[str, object] = field(default_factory=dict)


@dataclass
class RepairOutcome:
    text: str
    confidence: float
    source: str = "raw"
    warnings: list[str] = field(default_factory=list)
    candidates: list[RepairCandidate] = field(default_factory=list)


def _clamp(value: float) -> float:
    try:
        num = float(value)
    except (TypeError, ValueError):
        return 0.0
    if num < 0.0:
        return 0.0
    if num > 1.0:
        return 1.0
    return num


def standardize_formula_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def strip_invisible_formula_noise(text: str) -> str:
    return _RE_INVISIBLE_FORMULA_NOISE.sub("", str(text or ""))


def _compact_formula_text(text: str) -> str:
    return re.sub(r"\s+", "", strip_invisible_formula_noise(text))


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


def _normalize_rendered_segment(text: str) -> str:
    value = strip_invisible_formula_noise(text)
    if not value:
        return ""
    for src, dst in _RENDERED_SYMBOL_TO_LATEX.items():
        value = value.replace(src, dst)
    for src, dst in _RENDERED_GREEK_TO_LATEX.items():
        value = value.replace(src, dst)
    value = ensure_latex_command_word_boundaries(value)
    value = standardize_formula_spaces(value)
    if not value:
        return ""
    value = _RE_RENDERED_FUNCTION_NAME.sub(lambda m: "\\" + m.group(1), value)

    def _replace_function_question(match: re.Match[str]) -> str:
        next_char = match.string[match.end():match.end() + 1]
        return match.group(1) if next_char in {"(", "["} else match.group(1) + " "

    value = _RE_RENDERED_FUNCTION_QUESTION.sub(_replace_function_question, value)
    value = value.replace("?", "")
    value = re.sub(r"\\([A-Za-z]+)\s+([\(\[\{])", r"\\\1\2", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def _looks_like_rendered_formula(text: str) -> bool:
    value = str(text or "").strip()
    if not value or len(value) > 120:
        return False
    if _RE_ESCAPE_ARTIFACT.search(value):
        return False
    if _RE_LATEX_COMMAND.search(value):
        return True
    return bool(_RE_FORMULA_OPERATOR.search(value) and _RE_LETTER.search(value))


def extract_rendered_segment_formulas(text: str) -> list[tuple[str, list[str]]]:
    raw = str(text or "")
    if not raw:
        return []

    out: list[tuple[str, list[str]]] = []
    seen: set[str] = set()
    for segment in _RE_ESCAPE_ARTIFACT.split(raw):
        if not (
            _RE_RENDERED_MATH_HINT.search(segment)
            or any(ch in segment for ch in _RENDERED_GREEK_TO_LATEX)
            or "\u2061" in segment
            or _RE_RENDERED_QUESTION_HINT.search(segment)
        ):
            continue
        normalized = _normalize_rendered_segment(segment)
        if not _looks_like_rendered_formula(normalized):
            continue
        if len(normalized) > 160:
            continue
        if not (
            _RE_SIMPLE_RENDERED_COMMAND.search(normalized)
            or normalized.startswith(r"\nabla \cdot")
            or normalized.startswith(r"\nabla \times")
            or "=" in normalized
            or normalized.count("+") >= 1
            or normalized.count("-") >= 1
        ):
            continue
        if normalized in seen:
            continue
        seen.add(normalized)
        out.append((normalized, ["rendered_segment_extracted"]))
    return out


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
    value = standardize_formula_spaces(strip_invisible_formula_noise(prefix))
    if not value:
        return ""
    if len(value) > 12:
        end_candidates: list[str] = []
        if value.endswith("="):
            start_min = max(0, len(value) - 16)
            for start in range(start_min, len(value)):
                seg = value[start:].strip()
                if not seg or not seg.endswith("="):
                    continue
                if "=" in seg[:-1]:
                    continue
                end_candidates.append(seg)
        structured = [
            seg for seg in end_candidates
            if any(ch in seg for ch in "^_({\\")
        ]
        structured_valid = [
            seg for seg in structured
            if seg[:1].isalnum() or seg[:1] in {"\\", "("}
        ]
        end_valid = [
            seg for seg in end_candidates
            if seg[:1].isalnum() or seg[:1] in {"\\", "("}
        ]
        if structured_valid:
            value = min(structured_valid, key=len)
        elif end_valid:
            value = min(end_valid, key=len)
        else:
            tail_segment = value.split("=")[-1]
            additive_match = re.search(r"([A-Za-z0-9]\s*[+\-*/]\s*)$", tail_segment)
            if additive_match:
                value = standardize_formula_spaces(additive_match.group(1))
            else:
                return ""
    if value.count("=") > 1:
        return ""
    if value[-1] not in {"=", "+", "-", "*", "/", "(", "[", "{", ","}:
        return ""
    return value


def _looks_like_formula_restart(text: str, idx: int) -> bool:
    candidate = standardize_formula_spaces(text[idx:])
    if not candidate:
        return False
    head = candidate[:1]
    if not (
        head.isalpha()
        or head in {"(", "[", "∫", "∑", "∂", "∇", "√"}
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
        and candidate[1] in {"∫", "∑", "∂", "∇", "√"}
    ):
        return False
    eq_pos = candidate.find("=")
    return 0 <= eq_pos <= 40


def _trim_repaired_command_tail(tail: str, *, single_command: bool) -> str:
    value = strip_invisible_formula_noise(tail)
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
    return standardize_formula_spaces(value).rstrip(",;")


def _extract_trailing_argument_hint(suffix: str) -> tuple[str, int]:
    value = strip_invisible_formula_noise(suffix)
    value = _RE_ESCAPE_ARTIFACT.sub("", value)
    value = value.replace("?", "")
    leading_len = len(value) - len(value.lstrip())
    value = value.lstrip()
    if not value:
        return "", 0
    if value[:1] == "{":
        end = _consume_brace_group(value, 0)
        return (value[:end], leading_len + end) if end > 0 else ("", 0)
    if value[:1] == "(":
        depth = 1
        idx = 1
        while idx < len(value) and depth > 0:
            if value[idx] == "(":
                depth += 1
            elif value[idx] == ")":
                depth -= 1
            idx += 1
        return (value[:idx], leading_len + idx) if depth == 0 else ("", 0)
    if value[:1] == "[":
        depth = 1
        idx = 1
        while idx < len(value) and depth > 0:
            if value[idx] == "[":
                depth += 1
            elif value[idx] == "]":
                depth -= 1
            idx += 1
        return (value[:idx], leading_len + idx) if depth == 0 else ("", 0)
    if value[:1] == "\\":
        atom = _extract_first_formula_atom(value)
        return atom, leading_len + len(atom)
    return value[:1], leading_len + 1


def _complete_trailing_command(text: str, suffix: str) -> tuple[str, str]:
    value = str(text or "").rstrip()
    if not value:
        return value, suffix
    if not (_RE_TRAILING_COMMAND_ONLY.search(value) or _RE_TRAILING_FUNCTION_LIKE.search(value)):
        return value, suffix
    hint, consumed = _extract_trailing_argument_hint(suffix)
    if not hint:
        return value, suffix
    joiner = "" if hint[:1] in {"(", "[", "{"} else " "
    return f"{value}{joiner}{hint}".rstrip(), suffix[consumed:]


def extract_repaired_command_formula(text: str) -> tuple[str | None, list[str]]:
    value = strip_invisible_formula_noise(text)
    spans = _iter_repaired_command_spans(value)
    if not spans:
        return None, []

    first_start = spans[0][0]
    last_end = spans[-1][1]
    prefix = _keep_repaired_command_prefix(value[:first_start])
    suffix = value[last_end:]
    middle, tail_seed = _complete_trailing_command(value[first_start:last_end], suffix)
    tail = _trim_repaired_command_tail(
        tail_seed,
        single_command=len(spans) == 1,
    )

    # Improved duplicate removal: handle both exact and partial duplicates
    if tail and prefix:
        # Remove exact duplicate
        if tail == prefix and tail[-1] in {"+", "-", "*", "/", "="}:
            tail = ""
        # Remove partial duplicate suffix (e.g., "i2" in "i=1ni2")
        elif prefix.endswith(tail) and len(tail) <= 8:
            tail = ""
        # Remove partial duplicate prefix (e.g., "i=" in "i=1")
        elif tail.startswith(prefix) and len(prefix) <= 8:
            tail = ""

    if tail and len(tail) <= 3 and tail[-1] in {"+", "-", "*", "/"} and "=" in middle:
        tail = ""
    extracted = f"{prefix}{middle}{tail}".strip()
    if not extracted:
        return None, []
    return extracted, ["escape_placeholder_formula_extracted"]


def _extract_assignment_prefix(prefix: str) -> str:
    value = standardize_formula_spaces(strip_invisible_formula_noise(prefix))
    if not value:
        return ""
    match = re.search(
        r"([A-Za-zΑ-Ωα-ω][A-Za-z0-9Α-Ωα-ω\(\)\[\],]{0,24}\s*=\s*)$",
        value,
    )
    if match:
        return standardize_formula_spaces(match.group(1))
    return _keep_repaired_command_prefix(value)


def _clean_environment_tail(tail: str) -> str:
    value = standardize_formula_spaces(strip_invisible_formula_noise(tail))
    if not value:
        return ""
    if value[:1] in {"=", "+", "-"}:
        return _trim_repaired_command_tail(value, single_command=False)
    return ""


def extract_environment_formula(text: str) -> tuple[str | None, list[str]]:
    raw = str(text or "")
    match = re.search(
        r"\\begin\{(?P<env>[A-Za-z]+)\}(?P<body>[\s\S]+?)\\end\{(?P=env)\}",
        raw,
    )
    if not match:
        return None, []

    env = match.group("env").strip()
    body = match.group("body")
    if env in _ROW_SPLIT_ENVS:
        body = _RE_ESCAPE_ARTIFACT.sub(lambda _: r"\\", body)
        body = re.sub(r"(?:\\\s*){3,}", r"\\\\", body)
        body = re.sub(r"\s*\\\\\s*", r"\\\\", body)
    body = standardize_formula_spaces(body)

    begin = rf"\begin{{{env}}}"
    end = rf"\end{{{env}}}"
    prefix = _extract_assignment_prefix(raw[:match.start()])
    tail = _clean_environment_tail(raw[match.end():])
    rebuilt = f"{prefix}{begin}{body}{end}{tail}".strip()
    if not rebuilt:
        return None, []
    return rebuilt, ["environment_formula_extracted"]


def _paren_balance(text: str) -> int:
    return (
        text.count("(") - text.count(")")
        + text.count("[") - text.count("]")
        + text.count("{") - text.count("}")
    )


def _is_valid_context_start(ch: str) -> bool:
    return bool(ch) and (ch.isalnum() or ch in {"(", "[", "{", "\\"})


def _is_valid_context_end(ch: str) -> bool:
    return bool(ch) and (ch.isalnum() or ch in {")", "]", "}"})


def _join_operator_formula(left: str, core: str, right: str) -> str:
    left_part = str(left or "").rstrip()
    core_part = str(core or "").strip()
    right_part = str(right or "").lstrip()
    if not core_part:
        return f"{left_part}{right_part}".strip()
    if left_part and left_part[-1].isalnum():
        core_part = " " + core_part
    if right_part and (right_part[:1].isalnum() or right_part[:1] == "\\"):
        core_part = core_part + " "
    return f"{left_part}{core_part}{right_part}".strip()


def _collect_left_context_options(prefix: str) -> list[str]:
    value = standardize_formula_spaces(strip_invisible_formula_noise(prefix))
    if not value:
        return []
    options: list[str] = []
    seen: set[str] = set()
    start_min = max(0, len(value) - 18)
    for start in range(len(value) - 1, start_min - 1, -1):
        seg = value[start:].strip()
        if not seg or seg in seen:
            continue
        if not _is_valid_context_start(seg[:1]):
            continue
        if seg.startswith((")", "]", "}")) or seg.endswith("\\"):
            continue
        if "ESC_" in seg or "__ESC" in seg or seg.count("=") > 1:
            continue
        if _paren_balance(seg) < -1:
            continue
        seen.add(seg)
        options.append(seg)
    options.sort(key=len)
    return options[:6]


def _collect_right_context_options(suffix: str, open_need: int = 0) -> list[str]:
    value = standardize_formula_spaces(strip_invisible_formula_noise(suffix)).lstrip()
    if not value:
        return []
    options: list[str] = []
    seen: set[str] = set()
    end_max = min(len(value), 18)
    for end in range(1, end_max + 1):
        seg = value[:end].strip()
        if not seg or seg in seen:
            continue
        if not _is_valid_context_start(seg[:1]) or not _is_valid_context_end(seg[-1]):
            continue
        if "ESC_" in seg or "__ESC" in seg or seg.count("=") > 1:
            continue
        if open_need > 0 and _paren_balance(seg) > -open_need:
            continue
        seen.add(seg)
        options.append(seg)
    options.sort(key=len)
    return options[:6]


def _consume_balanced_delimiter(text: str, open_ch: str, close_ch: str) -> tuple[str, int]:
    if not text or text[:1] != open_ch:
        return "", 0
    depth = 1
    idx = 1
    while idx < len(text) and depth > 0:
        if text[idx] == open_ch:
            depth += 1
        elif text[idx] == close_ch:
            depth -= 1
        idx += 1
    if depth != 0:
        return "", 0
    return text[:idx], idx


def _extract_first_formula_atom(suffix: str) -> str:
    value = strip_invisible_formula_noise(suffix).lstrip()
    if not value:
        return ""
    head = value[:1]
    if head == "{":
        group, _ = _consume_balanced_delimiter(value, "{", "}")
        return group
    if head == "(":
        group, _ = _consume_balanced_delimiter(value, "(", ")")
        return group
    if head == "[":
        group, _ = _consume_balanced_delimiter(value, "[", "]")
        return group
    if head == "\\":
        match = _RE_LATEX_COMMAND_NAME.match(value)
        if not match:
            return ""
        end = _consume_optional_scripts(value, match.end())
        group_count = _REPAIRED_COMMAND_GROUPS.get(match.group(1).strip().lower(), 0)
        for _ in range(group_count):
            probe = _skip_spaces(value, end)
            next_end = _consume_brace_group(value, probe)
            if next_end == probe:
                break
            end = next_end
        return value[:end].strip()
    match = re.match(
        r"[A-Za-zΑ-Ωα-ω0-9](?:_(?:\{[^{}]+\}|[A-Za-z0-9])|\^(?:\{[^{}]+\}|[A-Za-z0-9]))*",
        value,
    )
    if match:
        return match.group(0).strip()
    return ""


def extract_command_context_formulas(text: str) -> list[tuple[str, list[str]]]:
    raw = str(text or "")
    out: list[tuple[str, list[str]]] = []
    spans = _iter_repaired_command_spans(raw)
    for start, end in spans:
        match = _RE_LATEX_COMMAND_NAME.match(raw, start)
        if not match:
            continue
        cmd = match.group(1).strip()
        cmd_lower = cmd.lower()
        if cmd_lower not in _FUNCTION_LIKE_COMMANDS and cmd_lower not in _SYMBOL_CONTEXT_COMMANDS:
            continue
        core = raw[start:end]
        atom = _extract_first_formula_atom(raw[end:])
        if not atom:
            continue
        joiner = "" if atom[:1] in {"(", "[", "{"} else " "
        out.append((f"{core}{joiner}{atom}".strip(), ["command_context_extracted"]))
    return out


def _sanitize_escape_artifacts(text: str) -> tuple[str, list[str]]:
    value = str(text or "")
    cleaned = _RE_ESCAPE_ARTIFACT.sub(" ", value)
    cleaned = standardize_formula_spaces(cleaned)
    if cleaned == standardize_formula_spaces(value):
        return standardize_formula_spaces(value), []
    return cleaned, ["escape_placeholder_artifact_stripped"]


def extract_big_operator_formulas(text: str) -> list[tuple[str, list[str]]]:
    raw = str(text or "")
    out: list[tuple[str, list[str]]] = []
    for match in re.finditer(r"\\(?P<op>sum|int|prod|lim)\b", raw):
        op = match.group("op")
        pos = match.end()
        pos = _skip_spaces(raw, pos)

        subscript = ""
        if pos < len(raw) and raw[pos] == "{":
            end = _consume_brace_group(raw, pos)
            if end > pos:
                raw_group = raw[pos + 1:end - 1]
                cleaned_group, _ = _sanitize_escape_artifacts(raw_group)
                cleaned_group = cleaned_group.strip()
                if cleaned_group:
                    subscript = cleaned_group
                pos = end

        while True:
            artifact = _RE_ESCAPE_ARTIFACT.match(raw, pos)
            if artifact is None:
                break
            pos = _skip_spaces(raw, artifact.end())

        body = _extract_first_formula_atom(raw[pos:])
        if not body:
            continue

        cleaned_body, body_warnings = _sanitize_escape_artifacts(body)
        rebuilt = f"\\{op}"
        warnings = ["big_operator_context_extracted"]
        if subscript:
            rebuilt += f"_{{{subscript}}}"
        rebuilt += cleaned_body if cleaned_body[:1] in {"(", "[", "{"} else f" {cleaned_body}"
        warnings.extend(body_warnings)
        out.append((rebuilt.strip(), list(dict.fromkeys(warnings))))
    return out


def extract_log_base_formula(raw_text: str, repaired_text: str) -> tuple[str | None, list[str]]:
    raw = str(raw_text or "")
    if "log" not in raw.lower():
        return None, []

    prefix_segment = _RE_ESCAPE_ARTIFACT.split(raw, maxsplit=1)[0] if raw else ""
    normalized_prefix = _normalize_rendered_segment(prefix_segment)
    prefix_match = _RE_LOG_BASE_PREFIX.search(normalized_prefix)
    if not prefix_match:
        return None, []

    rhs = ""
    extracted, _ = extract_repaired_command_formula(repaired_text)
    if extracted and "=" in extracted:
        rhs = extracted.split("=", 1)[1].strip()
    if not rhs:
        frac_match = re.search(r"(\\frac\{[\s\S]+?\}\{[\s\S]+?\})", repaired_text)
        if frac_match:
            rhs = frac_match.group(1).strip()
    if not rhs:
        return None, []

    base = prefix_match.group(1).strip()
    arg = prefix_match.group(2).strip()
    rebuilt = rf"\log_{{{base}}} {arg}={rhs}"
    return rebuilt, ["log_base_context_extracted"]


def extract_operator_centered_formulas(text: str) -> list[tuple[str, list[str]]]:
    raw = str(text or "")
    out: list[tuple[str, list[str]]] = []
    spans = _iter_repaired_command_spans(raw)
    for start, end in spans:
        match = _RE_LATEX_COMMAND_NAME.match(raw, start)
        if not match:
            continue
        cmd = match.group(1).strip()
        core = raw[start:end]
        prefix = raw[:start]
        suffix = raw[end:]

        if cmd in _BINARY_OPERATOR_COMMANDS:
            left_options = _collect_left_context_options(prefix)
            for left in left_options:
                need = max(0, _paren_balance(left))
                right_options = _collect_right_context_options(suffix, need)
                for right in right_options:
                    if not any(ch.isalpha() for ch in left + right):
                        continue
                    out.append((_join_operator_formula(left, core, right), ["operator_context_extracted"]))

        if cmd in _PREFIX_OPERATOR_COMMANDS:
            right_options = _collect_right_context_options(suffix, 0)
            for right in right_options:
                if not any(ch.isalpha() for ch in right):
                    continue
                out.append((_join_operator_formula("", core, right), ["operator_context_extracted"]))

    return out


def _is_formula_segment_candidate(seg: str) -> bool:
    if not seg:
        return False
    if len(seg) < 5 or len(seg) > 120:
        return False
    if seg.count("=") != 1:
        return False
    if re.search(r"[\u4e00-\u9fff]", seg):
        return False

    left, right = seg.split("=", 1)
    if not left or not right:
        return False
    if not (left[0].isalpha() or left[0] in {"(", "\\"}):
        return False
    if not (right[0].isalpha() or right[0].isdigit() or right[0] in {"(", "\\"}):
        return False
    if not (left[-1].isalnum() or left[-1] in {")", "]", "}"}):
        return False
    if not (right[-1].isalnum() or right[-1] in {")", "]", "}"}):
        return False

    for ch in seg:
        if ch.isalnum():
            continue
        if ch in {"+", "-", "*", "/", "^", "_", "=", "(", ")", "[", "]", "{", "}", ".", ",", "\\", "→", "←"}:
            continue
        return False

    if not any(ch.isalpha() for ch in seg):
        return False
    if not re.search(r"[+\-*/^_=]|\d|\\", seg):
        return False
    return True


def _formula_skeleton_key(text: str) -> str:
    return re.sub(r"[\^_{}\(\)\s]", "", text.lower())


def _formula_detail_score(text: str) -> int:
    return (
        text.count("^") * 4
        + text.count("_") * 4
        + text.count("{")
        + text.count("}")
        + text.count("\\") * 2
        + len(text) // 20
    )


def _is_formula_expression_candidate(seg: str) -> bool:
    if not seg:
        return False
    if len(seg) < 4 or len(seg) > 120:
        return False
    if re.search(r"[\u4e00-\u9fff]", seg):
        return False
    if seg.count("=") > 1:
        return False

    if not (seg[0].isalpha() or seg[0] in {"(", "\\", "1"}):
        return False
    if not (seg[-1].isalnum() or seg[-1] in {")", "]", "}"}):
        return False

    for ch in seg:
        if ch.isalnum():
            continue
        if ch in {
            "+", "-", "*", "/", "^", "_", "=", "(", ")", "[", "]", "{", "}",
            ".", ",", "\\", "→", "←", "±",
        }:
            continue
        return False

    if not any(ch.isalpha() for ch in seg):
        return False
    if not (
        re.search(r"[+\-*/^_=\\]", seg)
        or re.search(r"(?<=[A-Za-z)])\d", seg)
        or re.search(r"[A-Za-z]\(", seg)
    ):
        return False
    return True


def _extract_exact_repeated_expression_candidate(s: str) -> str | None:
    if not s or len(s) > 240:
        return None

    best: str | None = None
    best_score = (-1, -1, -1)
    for repeat in range(2, min(8, len(s)) + 1):
        if len(s) % repeat != 0:
            continue
        seg_len = len(s) // repeat
        if seg_len < 4 or seg_len > 120:
            continue
        seg = s[:seg_len]
        if seg * repeat != s:
            continue
        if not _is_formula_expression_candidate(seg):
            continue
        score = (
            _formula_detail_score(seg),
            len(seg),
            repeat,
        )
        if best is None or score > best_score:
            best = seg
            best_score = score
    return best


def extract_repeated_formula_candidate(text: str) -> str | None:
    s = _compact_formula_text(text)
    if not s or len(s) > 240:
        return None

    if s.count("=") < 2:
        return _extract_exact_repeated_expression_candidate(s)

    cache: dict[int, tuple[str, ...] | None] = {}

    def _dfs(start: int) -> tuple[str, ...] | None:
        if start == len(s):
            return ()
        if start in cache:
            return cache[start]
        if not (s[start].isalpha() or s[start] in {"(", "\\"}):
            cache[start] = None
            return None

        best: tuple[str, ...] | None = None
        best_score = (-1, -1, -1)
        max_end = min(len(s), start + 140)
        for end in range(start + 5, max_end + 1):
            seg = s[start:end]
            eq_count = seg.count("=")
            if eq_count == 0:
                continue
            if eq_count > 1:
                break
            if not _is_formula_segment_candidate(seg):
                continue
            rest = _dfs(end)
            if rest is None:
                continue
            candidate = (seg,) + rest
            lengths = [len(item) for item in candidate]
            avg_len = sum(lengths) / len(lengths)
            variance = sum((ln - avg_len) ** 2 for ln in lengths)
            score = (
                len(candidate),
                -variance,
                sum(_formula_detail_score(item) for item in candidate),
            )
            if best is None or score > best_score:
                best = candidate
                best_score = score

        cache[start] = best
        return best

    result = _dfs(0)
    if result is None or len(result) < 2:
        return None

    counts: dict[str, int] = {}
    for seg in result:
        key = _formula_skeleton_key(seg)
        counts[key] = counts.get(key, 0) + 1
    dominant_key = max(counts, key=counts.get)
    dominant = [seg for seg in result if _formula_skeleton_key(seg) == dominant_key]
    if len(dominant) < 2:
        return None
    return max(dominant, key=lambda seg: (_formula_detail_score(seg), len(seg)))


def extract_rendered_fraction_formula(text: str) -> tuple[str | None, list[str]]:
    s = _compact_formula_text(text).replace("−", "-")
    if not s or "=" in s or "\\" in s:
        return None, []

    best_den: str | None = None
    best_score = (-1, 0, 0)
    for idx in range(len(s) - 2):
        if s[idx] != "1" or s[idx + 1] not in "+-":
            continue
        max_end = min(len(s), idx + 18)
        for end in range(idx + 3, max_end + 1):
            den = s[idx:end]
            if not re.fullmatch(r"1[+\-][A-Za-z0-9()]+", den):
                continue
            count = s.count(den)
            if count < 2:
                continue
            remainder = s.replace(den, "").replace("1", "")
            if remainder:
                continue
            score = (count, -len(den), len(den))
            if best_den is None or score > best_score:
                best_den = den
                best_score = score

    if not best_den:
        return None, []
    return rf"\frac{{1}}{{{best_den}}}", ["rendered_fraction_reconstructed"]


_RE_CROSS_FRACTION_IDENTITY = re.compile(
    r"^(?P<a>[A-Za-z])(?P<b>[A-Za-z])\+(?P<c>[A-Za-z])(?P<d>[A-Za-z])="
    r"(?P=a)(?P=d)\+(?P=b)(?P=c)(?P=b)(?P=d)$"
)


def extract_cross_fraction_identity_formula(text: str) -> tuple[str | None, list[str]]:
    s = _compact_formula_text(text)
    if not s:
        return None, []
    match = _RE_CROSS_FRACTION_IDENTITY.fullmatch(s)
    if not match:
        return None, []
    a = match.group("a")
    b = match.group("b")
    c = match.group("c")
    d = match.group("d")
    latex = rf"\frac{{{a}}}{{{b}}}+\frac{{{c}}}{{{d}}}=\frac{{{a}{d}+{b}{c}}}{{{b}{d}}}"
    return latex, ["rendered_fraction_identity_reconstructed"]


_RE_QUADRATIC_COPY_SEQ = re.compile(
    r"x=.*?-?b.*?[±\\pm].*?b2.*?4ac.*?2a",
    flags=re.IGNORECASE,
)


def extract_quadratic_formula_copy(text: str) -> tuple[str | None, list[str]]:
    s = _compact_formula_text(text).replace("−", "-")
    if not s:
        return None, []
    if not _RE_QUADRATIC_COPY_SEQ.search(s):
        return None, []
    if not all(token in s for token in ("x=", "4ac", "2a")):
        return None, []
    latex = r"x=\frac{-b\pm\sqrt{b^{2}-4ac}}{2a}"
    return latex, ["quadratic_formula_copy_reconstructed"]


def _brace_balance_score(text: str) -> int:
    score = 0
    if text.count("{") == text.count("}"):
        score += 4
    else:
        score -= abs(text.count("{") - text.count("}")) * 3
    if text.count("\\begin") == text.count("\\end"):
        score += 4
    if text.count("\\left") == text.count("\\right"):
        score += 3
    return score


def _candidate_score(candidate: RepairCandidate) -> float:
    text = str(candidate.text or "").strip()
    if not text:
        return -1e6

    unresolved_count = len(_RE_ESCAPE_ARTIFACT.findall(text))
    command_stripped = _RE_LATEX_COMMAND.sub("", text)
    semantic_payload = re.sub(r"[\s{}_^]", "", command_stripped)
    profile = candidate.metadata.get("escape_profile")
    if not isinstance(profile, EscapeFeatureProfile):
        profile = None
    score = 0.0
    score += _brace_balance_score(text)
    score += min(16, len(_RE_LATEX_COMMAND.findall(text)) * 2)
    score += 4 if _RE_FORMULA_OPERATOR.search(text) else 0
    score += 2 if _RE_LETTER.search(text) else -4
    score += 2 if any(ch in text for ch in ("∫", "∑", "∏", "√", "∞", "∩", "∪", "×")) else 0
    score += 6 if candidate.source == "escaped_extracted" else 0
    score += 7 if candidate.source == "environment_extracted" else 0
    score += 4 if candidate.source == "operator_context" else 0
    score += 6 if candidate.source == "command_context" else 0
    score += 7 if candidate.source == "big_operator_context" else 0
    score += 4 if candidate.source == "big_operator_context" and any(mark in text for mark in (r"\frac", "^", "_")) else 0
    score += 14 if candidate.source == "log_base_context" else 0
    score += 8 if candidate.source == "rendered_segment" else 0
    score += 11 if candidate.source == "rendered_fraction" else 0
    score += 12 if candidate.source == "fraction_identity" else 0
    score += 12 if candidate.source == "quadratic_formula_copy" else 0
    score += 5 if candidate.source == "escaped_extracted" and any(mark in text for mark in (r"\frac", "^", "_")) else 0
    score += 3 if candidate.source == "escaped_extracted" and text.count("\\") >= 2 and "=" in text else 0
    score += 3 if candidate.source == "escaped_recovered" else 0
    score += 2 if candidate.source == "repeated_segment" else 0
    if profile and profile.command_count:
        if candidate.source in {"escaped_extracted", "big_operator_context", "log_base_context"}:
            score += min(6, profile.command_count * 1.5)
        if profile.strong_structural and candidate.source == "rendered_segment":
            score -= 7
        if profile.strong_structural and candidate.source == "command_context":
            score -= 3
        if "frac" in profile.commands and r"\frac" not in text:
            score -= 6
        if "sqrt" in profile.commands and r"\sqrt" not in text:
            score -= 5
        if any(cmd in {"sum", "prod", "int"} for cmd in profile.commands) and not any(
            mark in text for mark in (r"\sum", r"\prod", r"\int")
        ):
            score -= 6
    score -= unresolved_count * 8
    # Penalty for mixing LaTeX commands with rendered symbols
    # Reduced penalty for intermediate repair states that naturally mix both
    if _RE_LATEX_COMMAND.search(text) and _RE_RENDERED_MATH_HINT.search(text):
        if candidate.source == "escaped_extracted":
            score -= 8
        latex_count = len(_RE_LATEX_COMMAND.findall(text))
        if latex_count < 2:
            score -= 6  # Heavy penalty if very few LaTeX commands
        else:
            score -= 2  # Light penalty for intermediate states with many commands
    if _RE_TRAILING_COMMAND_ONLY.search(text):
        score -= 12
    if _RE_TRAILING_FUNCTION_LIKE.search(text):
        score -= 10
    if not re.search(r"[A-Za-zΑ-Ωα-ω0-9]", semantic_payload):
        score -= 8
    if text.endswith(("+", "-", "=", "\\", "{", "_", "^", ",")):
        score -= 6
    if len(text) > 24:
        score -= min(4.0, (len(text) - 24) * 0.05)
    if len(text) > 180:
        score -= (len(text) - 180) * 0.1

    candidate.metadata["unresolved_escape_count"] = unresolved_count
    return score


def _make_candidate(
        text: str,
        source: str,
        warnings: list[str] | None = None,
        *,
        escape_profile: EscapeFeatureProfile | None = None) -> RepairCandidate | None:
    value = str(text or "")
    value = re.sub(
        r"(?<=[A-Za-z0-9}])\\{2,}(?=\\(?:text|operatorname|mathbb|mathcal|mathfrak|mathscr|mathsf|mathtt)\{)",
        " ",
        value,
    )
    value = standardize_formula_spaces(value)
    if not value:
        return None
    candidate = RepairCandidate(
        text=value,
        source=source,
        warnings=list(dict.fromkeys(warnings or [])),
    )
    if escape_profile is not None:
        candidate.metadata["escape_profile"] = escape_profile
    candidate.score = _candidate_score(candidate)
    return candidate


def repair_formula_text(text: str) -> RepairOutcome:
    raw = strip_invisible_formula_noise(text)
    escape_profile = build_escape_feature_profile(raw)
    candidates: list[RepairCandidate] = []
    seen: set[tuple[str, str]] = set()

    def _push(candidate: RepairCandidate | None) -> None:
        if candidate is None:
            return
        key = (candidate.source, candidate.text)
        if key in seen:
            return
        seen.add(key)
        candidates.append(candidate)

    _push(_make_candidate(raw, "raw", escape_profile=escape_profile))

    repeated = extract_repeated_formula_candidate(raw)
    if repeated:
        _push(_make_candidate(
            repeated,
            "repeated_segment",
            ["formula_repeated_segment_extracted"],
            escape_profile=escape_profile,
        ))
        repeated_fraction_identity, repeated_fraction_warnings = extract_cross_fraction_identity_formula(repeated)
        if repeated_fraction_identity:
            _push(_make_candidate(
                repeated_fraction_identity,
                "fraction_identity",
                repeated_fraction_warnings,
                escape_profile=escape_profile,
            ))

    rendered_fraction, rendered_fraction_warnings = extract_rendered_fraction_formula(raw)
    if rendered_fraction:
        _push(_make_candidate(
            rendered_fraction,
            "rendered_fraction",
            rendered_fraction_warnings,
            escape_profile=escape_profile,
        ))

    fraction_identity, fraction_identity_warnings = extract_cross_fraction_identity_formula(raw)
    if fraction_identity:
        _push(_make_candidate(
            fraction_identity,
            "fraction_identity",
            fraction_identity_warnings,
            escape_profile=escape_profile,
        ))

    quadratic_formula, quadratic_warnings = extract_quadratic_formula_copy(raw)
    if quadratic_formula:
        _push(_make_candidate(
            quadratic_formula,
            "quadratic_formula_copy",
            quadratic_warnings,
            escape_profile=escape_profile,
        ))

    for rendered_formula, rendered_warnings in extract_rendered_segment_formulas(raw):
        _push(_make_candidate(
            rendered_formula,
            "rendered_segment",
            rendered_warnings,
            escape_profile=escape_profile,
        ))

    recovered, recovered_warnings = recover_escaped_latex_commands(raw)
    if recovered != raw or recovered_warnings:
        _push(_make_candidate(
            recovered,
            "escaped_recovered",
            recovered_warnings,
            escape_profile=escape_profile,
        ))
        extracted, extracted_warnings = extract_repaired_command_formula(recovered)
        if extracted:
            _push(_make_candidate(
                extracted,
                "escaped_extracted",
                list(dict.fromkeys(recovered_warnings + extracted_warnings)),
                escape_profile=escape_profile,
            ))
        env_formula, env_warnings = extract_environment_formula(recovered)
        if env_formula:
            _push(_make_candidate(
                env_formula,
                "environment_extracted",
                list(dict.fromkeys(recovered_warnings + env_warnings)),
                escape_profile=escape_profile,
            ))
        for operator_formula, operator_warnings in extract_operator_centered_formulas(recovered):
            _push(_make_candidate(
                operator_formula,
                "operator_context",
                list(dict.fromkeys(recovered_warnings + operator_warnings)),
                escape_profile=escape_profile,
            ))
        for command_formula, command_warnings in extract_command_context_formulas(recovered):
            _push(_make_candidate(
                command_formula,
                "command_context",
                list(dict.fromkeys(recovered_warnings + command_warnings)),
                escape_profile=escape_profile,
            ))
        for bigop_formula, bigop_warnings in extract_big_operator_formulas(recovered):
            _push(_make_candidate(
                bigop_formula,
                "big_operator_context",
                list(dict.fromkeys(recovered_warnings + bigop_warnings)),
                escape_profile=escape_profile,
            ))
        log_base_formula, log_base_warnings = extract_log_base_formula(raw, recovered)
        if log_base_formula:
            _push(_make_candidate(
                log_base_formula,
                "log_base_context",
                list(dict.fromkeys(recovered_warnings + log_base_warnings)),
                escape_profile=escape_profile,
            ))
        repeated_recovered = extract_repeated_formula_candidate(recovered)
        if repeated_recovered:
            _push(_make_candidate(
                repeated_recovered,
                "repeated_segment",
                list(dict.fromkeys(recovered_warnings + ["formula_repeated_segment_extracted"])),
                escape_profile=escape_profile,
            ))
            repeated_fraction_identity, repeated_fraction_warnings = extract_cross_fraction_identity_formula(repeated_recovered)
            if repeated_fraction_identity:
                _push(_make_candidate(
                    repeated_fraction_identity,
                    "fraction_identity",
                    list(dict.fromkeys(recovered_warnings + repeated_fraction_warnings)),
                    escape_profile=escape_profile,
                ))

    if not candidates:
        return RepairOutcome(text=standardize_formula_spaces(raw), confidence=0.45)

    best = max(candidates, key=lambda item: item.score)
    unresolved = int(best.metadata.get("unresolved_escape_count", 0) or 0)
    confidence = 0.48 + max(0.0, min(best.score, 28.0)) / 40.0 - unresolved * 0.08
    if best.source == "escaped_extracted":
        confidence += 0.10
    elif best.source == "escaped_recovered":
        confidence += 0.04
    elif best.source in {
        "environment_extracted",
        "operator_context",
        "command_context",
        "big_operator_context",
        "rendered_segment",
        "rendered_fraction",
        "fraction_identity",
        "quadratic_formula_copy",
    }:
        confidence += 0.08

    warnings = list(dict.fromkeys(best.warnings))
    if "ESC_" not in best.text and "__ESC" not in best.text:
        warnings = [item for item in warnings if item != "escape_placeholder_unresolved"]

    # Debug output for development/troubleshooting
    if os.getenv("FORMULA_REPAIR_DEBUG"):
        try:
            print(f"\n[FORMULA_REPAIR_DEBUG] Input: {raw[:80]}...")
            print(f"[FORMULA_REPAIR_DEBUG] Best candidate: {best.source} (score={best.score:.2f}, confidence={confidence:.2f})")
            print(f"[FORMULA_REPAIR_DEBUG] Result: {best.text}")
            print(f"[FORMULA_REPAIR_DEBUG] Top 5 candidates:")
            for i, c in enumerate(sorted(candidates, key=lambda x: x.score, reverse=True)[:5], 1):
                print(f"  {i}. [{c.source:25s}] score={c.score:6.2f} | {c.text[:60]}")
            if warnings:
                print(f"[FORMULA_REPAIR_DEBUG] Warnings: {', '.join(warnings)}")
            print()
        except (UnicodeEncodeError, UnicodeDecodeError):
            # Fallback for systems with encoding issues
            print(f"\n[FORMULA_REPAIR_DEBUG] Best: {best.source} (score={best.score:.2f})")
            print(f"[FORMULA_REPAIR_DEBUG] Candidates: {len(candidates)}\n")

    return RepairOutcome(
        text=best.text,
        confidence=_clamp(confidence),
        source=best.source,
        warnings=warnings,
        candidates=sorted(candidates, key=lambda item: item.score, reverse=True),
    )
