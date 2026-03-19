"""Shared formula symbol registries and LaTeX command boundary helpers."""

from __future__ import annotations

import re
from functools import lru_cache
from typing import Iterable

UNICODE_GREEK_TO_LATEX: dict[str, str] = {
    "α": r"\alpha",
    "β": r"\beta",
    "γ": r"\gamma",
    "δ": r"\delta",
    "ε": r"\epsilon",
    "ζ": r"\zeta",
    "η": r"\eta",
    "θ": r"\theta",
    "ι": r"\iota",
    "κ": r"\kappa",
    "λ": r"\lambda",
    "μ": r"\mu",
    "ν": r"\nu",
    "ξ": r"\xi",
    "π": r"\pi",
    "ρ": r"\rho",
    "σ": r"\sigma",
    "τ": r"\tau",
    "υ": r"\upsilon",
    "φ": r"\phi",
    "χ": r"\chi",
    "ψ": r"\psi",
    "ω": r"\omega",
    "ϵ": r"\varepsilon",
    "ϑ": r"\vartheta",
    "ϖ": r"\varpi",
    "ϱ": r"\varrho",
    "ς": r"\varsigma",
    "ϕ": r"\varphi",
    "Γ": r"\Gamma",
    "Δ": r"\Delta",
    "Θ": r"\Theta",
    "Λ": r"\Lambda",
    "Ξ": r"\Xi",
    "Π": r"\Pi",
    "Σ": r"\Sigma",
    "Υ": r"\Upsilon",
    "Φ": r"\Phi",
    "Ψ": r"\Psi",
    "Ω": r"\Omega",
}

UNICODE_OPERATOR_TO_LATEX: dict[str, str] = {
    "×": r"\times",
    "·": r"\cdot",
    "⋅": r"\cdot",
    "→": r"\to",
    "←": r"\leftarrow",
    "↔": r"\leftrightarrow",
    "⇒": r"\Rightarrow",
    "⇐": r"\Leftarrow",
    "⇔": r"\Leftrightarrow",
    "∞": r"\infty",
    "∑": r"\sum",
    "∫": r"\int",
    "∮": r"\oint",
    "∏": r"\prod",
    "√": r"\sqrt",
    "∩": r"\cap",
    "∪": r"\cup",
    "⊂": r"\subset",
    "⊆": r"\subseteq",
    "∖": r"\setminus",
    "∈": r"\in",
    "∀": r"\forall",
    "∃": r"\exists",
    "¬": r"\neg",
    "∧": r"\land",
    "∨": r"\lor",
    "∂": r"\partial",
    "∇": r"\nabla",
    "∥": r"\Vert",
    "ℏ": r"\hbar",
    "≥": r"\ge",
    "≤": r"\le",
    "±": r"\pm",
    "≠": r"\ne",
    "≈": r"\approx",
}

UNICODE_MATH_TO_LATEX: dict[str, str] = {
    **UNICODE_OPERATOR_TO_LATEX,
    **UNICODE_GREEK_TO_LATEX,
}

_STRUCTURAL_LATEX_COMMANDS = {"sqrt"}

LATEX_COMMAND_TO_UNICODE: dict[str, str] = {
    str(value).strip()[1:]: key
    for key, value in UNICODE_MATH_TO_LATEX.items()
    if str(value).strip().startswith("\\")
    and str(value).strip()[1:] not in _STRUCTURAL_LATEX_COMMANDS
}

# Keep a few compatible aliases used elsewhere in the project.
LATEX_COMMAND_TO_UNICODE.setdefault("rightarrow", "→")
LATEX_COMMAND_TO_UNICODE.setdefault("leftrightarrow", "↔")
LATEX_COMMAND_TO_UNICODE.setdefault("vert", "∥")

LATEX_COMMAND_BOUNDARY_NAMES: tuple[str, ...] = tuple(
    sorted(LATEX_COMMAND_TO_UNICODE.keys(), key=len, reverse=True)
)
LATEX_COMMAND_BOUNDARY_NAMES = tuple(
    sorted(
        {
            *LATEX_COMMAND_BOUNDARY_NAMES,
            *(
                str(value).strip()[1:]
                for value in UNICODE_MATH_TO_LATEX.values()
                if str(value).strip().startswith("\\")
            ),
        },
        key=len,
        reverse=True,
    )
)

KNOWN_LATEX_COMMAND_NAMES = frozenset(
    {
        *LATEX_COMMAND_BOUNDARY_NAMES,
        "frac",
        "sqrt",
        "left",
        "right",
        "begin",
        "end",
        "lim",
        "sin",
        "cos",
        "tan",
        "sec",
        "csc",
        "cot",
        "sinh",
        "cosh",
        "tanh",
        "log",
        "ln",
        "exp",
        "max",
        "min",
        "det",
        "arcsin",
        "arccos",
        "arctan",
        "vec",
        "overrightarrow",
        "overleftarrow",
        "hat",
        "bar",
        "overline",
        "underline",
        "boxed",
        "tilde",
        "overbrace",
        "underbrace",
        "mathbf",
        "boldsymbol",
        "mathrm",
        "operatorname",
        "mathit",
        "mathbb",
        "mathcal",
        "mathscr",
        "mathfrak",
        "mathsf",
        "mathtt",
    }
)


_RE_LATEX_COMMAND_TOKEN = re.compile(r"\\([A-Za-z]+)")


@lru_cache(maxsize=None)
def _boundary_name_order(command_names: tuple[str, ...]) -> tuple[str, ...]:
    return tuple(sorted(command_names, key=len, reverse=True))


def ensure_latex_command_word_boundaries(
    text: str,
    *,
    command_names: Iterable[str] | None = None,
) -> str:
    """Insert a separator after known LaTeX commands when glued to word chars."""

    value = str(text or "")
    if not value:
        return value

    names = tuple(
        sorted(
            {
                str(name or "").strip().lstrip("\\")
                for name in (
                    LATEX_COMMAND_BOUNDARY_NAMES if command_names is None else command_names
                )
                if str(name or "").strip()
            },
            key=len,
            reverse=True,
        )
    )
    if not names:
        return value
    names_set = set(names)
    ordered = _boundary_name_order(names)

    def _replace(match: re.Match[str]) -> str:
        raw = str(match.group(1) or "")
        if not raw or raw in names_set or raw in KNOWN_LATEX_COMMAND_NAMES:
            return match.group(0)
        for name in ordered:
            if raw.startswith(name):
                suffix = raw[len(name):]
                if suffix:
                    return f"\\{name} {suffix}"
                return f"\\{name}"
        return match.group(0)

    return _RE_LATEX_COMMAND_TOKEN.sub(_replace, value)
