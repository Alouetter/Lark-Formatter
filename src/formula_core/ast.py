"""Formula AST models used by formula rules."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


def _clamp_confidence(value: float) -> float:
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        numeric = 0.0
    if numeric < 0.0:
        return 0.0
    if numeric > 1.0:
        return 1.0
    return numeric


@dataclass
class FormulaNode:
    """Generic formula AST node.

    Required core fields follow the agreed schema:
    - kind
    - payload
    - children
    - confidence
    - warnings
    """

    kind: str
    payload: dict[str, Any] = field(default_factory=dict)
    children: list["FormulaNode"] = field(default_factory=list)
    source_type: str = ""
    confidence: float = 1.0
    warnings: list[str] = field(default_factory=list)

    def __post_init__(self) -> None:
        self.kind = str(self.kind or "").strip() or "unknown"
        self.source_type = str(self.source_type or "").strip().lower()
        self.confidence = _clamp_confidence(self.confidence)
        if not isinstance(self.payload, dict):
            self.payload = {}
        if not isinstance(self.children, list):
            self.children = []
        if not isinstance(self.warnings, list):
            self.warnings = []


@dataclass
class FormulaOccurrence:
    """A formula occurrence in the source document."""

    node: FormulaNode
    paragraph: Any
    paragraph_index: int
    location: str
    source_type: str
    is_block: bool
    is_formula_only: bool
    source_text: str = ""
    source_range: tuple[int, int] | None = None
    in_table: bool = False
    source_paragraph_indices: list[int] = field(default_factory=list)


@dataclass
class FormulaParseResult:
    """Unified parse result passed between formula rules."""

    occurrences: list[FormulaOccurrence] = field(default_factory=list)

    @property
    def total(self) -> int:
        return len(self.occurrences)

    @property
    def by_source(self) -> dict[str, int]:
        stats: dict[str, int] = {}
        for occ in self.occurrences:
            key = str(occ.source_type or "unknown")
            stats[key] = stats.get(key, 0) + 1
        return stats
