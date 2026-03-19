"""Runtime counters for formula pipeline rules."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class FormulaRuleStats:
    matched: int = 0
    converted: int = 0
    skipped_low_confidence: int = 0
    skipped_unsupported: int = 0
    skipped_dependency: int = 0
    errors: int = 0
    high_confidence: int = 0
    medium_confidence: int = 0
    low_confidence: int = 0

    def note_confidence(self, value: float) -> None:
        try:
            score = float(value)
        except (TypeError, ValueError):
            score = 0.0
        if score >= 0.85:
            self.high_confidence += 1
        elif score >= 0.60:
            self.medium_confidence += 1
        else:
            self.low_confidence += 1

    def confidence_summary(self) -> str:
        return (
            f"high={self.high_confidence}, "
            f"medium={self.medium_confidence}, "
            f"low={self.low_confidence}"
        )

    def to_summary(self) -> str:
        return (
            f"matched={self.matched}, converted={self.converted}, "
            f"skip_low_conf={self.skipped_low_confidence}, "
            f"skip_unsupported={self.skipped_unsupported}, "
            f"skip_dependency={self.skipped_dependency}, errors={self.errors}, "
            f"{self.confidence_summary()}"
        )
