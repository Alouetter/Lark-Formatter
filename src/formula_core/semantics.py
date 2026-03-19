"""Shared formula semantic constants used across parse/convert/style."""

from __future__ import annotations

FUNCTION_NAMES = frozenset(
    {
        "sin", "cos", "tan", "sec", "csc", "cot",
        "sinh", "cosh", "tanh",
        "log", "ln", "exp",
        "max", "min", "det", "lim",
        "arcsin", "arccos", "arctan",
    }
)

BIG_OPERATOR_NAMES = frozenset({"sum", "int", "iint", "iiint", "oint", "prod", "lim"})

GREEK_COMMAND_NAMES = frozenset(
    {
        "alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta",
        "theta", "iota", "kappa", "lambda", "mu", "nu", "xi",
        "pi", "rho", "sigma", "tau", "upsilon", "phi", "chi", "psi", "omega",
        "varepsilon", "vartheta", "varpi", "varrho", "varsigma", "varphi",
    }
)

BOLD_ITALIC_MARKER_COMMANDS = frozenset({"vec", "overrightarrow", "boldsymbol"})
BOLD_ROMAN_MARKER_COMMANDS = frozenset({"mathbf"})
ROMAN_MARKER_COMMANDS = frozenset({"mathrm", "operatorname"})
ITALIC_MARKER_COMMANDS = frozenset({"mathit"})
UPRIGHT_FAMILY_MARKER_COMMANDS = frozenset(
    {
        "mathbb",
        "mathcal",
        "mathfrak",
        "mathscr",
        "mathsf",
        "mathtt",
    }
)
