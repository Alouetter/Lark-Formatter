"""ESC placeholder feature extraction and recovery helpers."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

_RE_ESCAPE_PLACEHOLDER = re.compile(
    r"(?:__)?ESC_?(?P<hex>[0-9A-Fa-f]{8})(?:__)?(?P<tail>[A-Za-z]*?)"
    r"(?=\{|(?:__)?ESC_?[0-9A-Fa-f]{8}|[^A-Za-z]|$)"
)
_RE_ESCAPE_ARTIFACT = re.compile(r"(?:__)?ESC_?[0-9A-Fa-f]{8}(?:__)?")

_DIRECT_ESCAPE_COMMAND_MAP = {
    "rac": "frac",
    "qrt": "sqrt",
    "um": "sum",
    "nt": "int",
    "rod": "prod",
    "og": "log",
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
    "ext": "text",
    "peratorname": "operatorname",
    "athbb": "mathbb",
    "athcal": "mathcal",
    "athfrak": "mathfrak",
    "athscr": "mathscr",
    "athsf": "mathsf",
    "athtt": "mathtt",
    "inom": "binom",
    "eft": "left",
    "ight": "right",
    "verset": "overset",
    "elta": "delta",
    "ambda": "lambda",
    "imes": "times",
    "ap": "cap",
    "up": "cup",
    "bar": "bar",
    "hat": "hat",
    "verline": "overline",
    "nderline": "underline",
    "idetilde": "widetilde",
    "verrightarrow": "overrightarrow",
    "verleftarrow": "overleftarrow",
    "rcsin": "arcsin",
    "rccos": "arccos",
    "rctan": "arctan",
    "iint": "iint",
    "iiint": "iiint",
    "ubset": "subset",
    "ubseteq": "subseteq",
    "etminus": "setminus",
    "orall": "forall",
    "xists": "exists",
    "eg": "neg",
    "and": "land",
    "or": "lor",
    "ightarrow": "Rightarrow",
    "eftrightarrow": "Leftrightarrow",
    "nderset": "underset",
    "ilde": "tilde",
    "si": "psi",
    "sii": "psi",
    "hbar": "hbar",
    "ar": "bar",
    "vline": "Vert",
    "intc": "oint_C",
}

_ESCAPE_COMMAND_LEXICON = (
    "alpha", "beta", "gamma", "delta", "epsilon", "theta", "lambda", "mu",
    "pi", "sigma", "omega", "phi", "psi",
    "frac", "sqrt", "sum", "int", "iint", "iiint", "oint", "prod", "lim", "sim",
    "sin", "cos", "tan", "sec", "csc", "cot", "sinh", "cosh", "tanh", "log", "ln", "exp",
    "begin", "end", "left", "right", "text", "operatorname",
    "mathbb", "mathcal", "mathfrak", "mathscr", "mathsf", "mathtt",
    "partial", "nabla", "vec", "bar", "hat", "binom", "overset",
    "cdot", "times", "cap", "cup", "infty", "pm", "ge", "le", "to",
    "arcsin", "arccos", "arctan", "subset", "subseteq", "setminus",
    "forall", "exists", "neg", "land", "lor", "Rightarrow", "Leftrightarrow",
    "underline", "widetilde", "overrightarrow", "overleftarrow", "overline",
    "underset", "tilde", "psi", "hbar", "Vert",
)


@dataclass
class EscapeFeatureProfile:
    commands: list[str] = field(default_factory=list)
    command_count: int = 0
    strong_structural: bool = False


def _normalized_escape_tail(tail: str) -> str:
    key = str(tail or "").strip().lower()
    if not key:
        return ""
    while "esc" in key:
        key = key.replace("esc", "")
    return key


def _fallback_escape_candidates(key: str) -> list[str]:
    if not key:
        return []
    if len(key) < 2:
        return []

    ranked: list[tuple[int, str]] = []
    for cmd in _ESCAPE_COMMAND_LEXICON:
        if cmd.endswith(key):
            ranked.append((len(cmd) - len(key), cmd))
        elif key.endswith(cmd):
            ranked.append((len(key) - len(cmd), cmd))
    ranked.sort(key=lambda item: (item[0], len(item[1])))
    return [cmd for _, cmd in ranked[:3]]


def _resolve_escape_command_tail(tail: str, following: str, prefix: str) -> str | None:
    key = _normalized_escape_tail(tail)

    # 空tail通常是 \\ 被转义产生的（如矩阵中的行分隔符）
    if not key:
        # 检查后续字符，如果是数字或&，很可能是矩阵中的 \\
        next_char = (following or "").lstrip()[:1]
        if next_char in {"&", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "\\"}:
            return "\\"  # 返回反斜杠本身
        return None

    if key in _DIRECT_ESCAPE_COMMAND_MAP:
        return _DIRECT_ESCAPE_COMMAND_MAP[key]
    if key == "im":
        next_char = (following or "").lstrip()[:1]
        return "lim" if next_char in {"", "{", "_", "^"} else "sim"
    if key == "n":
        next_text = (following or "").lstrip()
        prefix_tail = str(prefix or "").rstrip()
        if next_text.startswith("("):
            return "ln"
        if next_text[:1].isupper():
            return "in"
        if (
            "\\mathbb" in next_text
            or "athbb" in next_text
            or "\\forall" in prefix_tail[-16:]
            or "\\exists" in prefix_tail[-16:]
        ):
            return "in"
        return "ln"
    if key == "in":
        return "sin"
    if key == "os":
        return "cos"
    if key == "an":
        return "tan"
    if key == "ec":
        next_text = (following or "").lstrip()
        next_char = next_text[:1]
        if next_char == "{" or (next_char.isalpha() and next_char.isupper()):
            return "vec"
        return "sec"
    if key == "o":
        return "to"
    if key == "e":
        next_char = (following or "").lstrip()[:1]
        if next_char.isdigit() or next_char in {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "-", "="}:
            return "ge"
    if key == "i":
        prefix_tail = str(prefix or "").rstrip()
        if prefix_tail and prefix_tail[-1] in "0123456789^_{}":
            return "pi"

    candidates = _fallback_escape_candidates(key)
    if candidates:
        return candidates[0]
    return None


def build_escape_feature_profile(text: str) -> EscapeFeatureProfile:
    raw = str(text or "")
    commands: list[str] = []
    for match in _RE_ESCAPE_PLACEHOLDER.finditer(raw):
        cmd = _resolve_escape_command_tail(
            match.group("tail"),
            raw[match.end():],
            raw[:match.start()],
        )
        if cmd:
            commands.append(cmd)

    strong_commands = {
        "frac", "sqrt", "sum", "prod", "int", "iint", "iiint", "oint",
        "begin", "end", "binom", "overset", "underset",
        "mathbb", "mathcal", "mathfrak", "mathscr", "mathsf", "mathtt",
        "overline", "underline", "widetilde",
        "overrightarrow", "overleftarrow",
    }
    strong_structural = any(cmd in strong_commands for cmd in commands) or len(commands) >= 2
    return EscapeFeatureProfile(
        commands=commands,
        command_count=len(commands),
        strong_structural=strong_structural,
    )


def recover_escaped_latex_commands(text: str, *, max_rounds: int = 6) -> tuple[str, list[str]]:
    current = str(text or "")
    changed = False

    for round_num in range(max_rounds):
        local_changed = False
        source_text = current

        def _replace(match: re.Match[str]) -> str:
            nonlocal local_changed
            tail = match.group("tail")
            cmd = _resolve_escape_command_tail(
                tail,
                source_text[match.end():],
                source_text[:match.start()],
            )
            if not cmd:
                return match.group(0)
            local_changed = True
            return "\\" + cmd

        updated = _RE_ESCAPE_PLACEHOLDER.sub(_replace, source_text)
        if not local_changed:
            break
        current = updated
        changed = True

    warnings: list[str] = []
    if changed:
        warnings.append("escape_placeholder_command_recovered")
    if "ESC" in current and _RE_ESCAPE_ARTIFACT.search(current):
        warnings.append("escape_placeholder_unresolved")
    return current, warnings


__all__ = [
    "_RE_ESCAPE_ARTIFACT",
    "_RE_ESCAPE_PLACEHOLDER",
    "EscapeFeatureProfile",
    "build_escape_feature_profile",
    "recover_escaped_latex_commands",
]
