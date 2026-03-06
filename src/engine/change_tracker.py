"""变更追踪器：记录排版过程中的每次修改"""

from dataclasses import dataclass, field
from typing import Any


@dataclass
class ChangeRecord:
    rule_name: str
    target: str
    section: str
    change_type: str  # "text" | "format" | "style" | "numbering"
    before: str
    after: str
    paragraph_index: int
    success: bool = True
    failure_reason: str | None = None


class ChangeTracker:
    def __init__(self):
        self.records: list[ChangeRecord] = []

    def record(self, rule_name: str, target: str, section: str,
               change_type: str, before: str, after: str,
               paragraph_index: int, success: bool = True,
               failure_reason: str | None = None):
        self.records.append(ChangeRecord(
            rule_name=rule_name, target=target, section=section,
            change_type=change_type, before=before, after=after,
            paragraph_index=paragraph_index, success=success,
            failure_reason=failure_reason,
        ))

    def get_by_rule(self, rule_name: str) -> list[ChangeRecord]:
        return [r for r in self.records if r.rule_name == rule_name]

    def get_failures(self) -> list[ChangeRecord]:
        return [r for r in self.records if not r.success]

    def summary(self) -> dict:
        by_rule: dict[str, dict[str, int]] = {}
        total: int = 0
        failures: int = 0
        for r in self.records:
            if r.rule_name not in by_rule:
                by_rule[r.rule_name] = {"success": 0, "failure": 0, "total": 0}
            by_rule[r.rule_name]["total"] += 1
            total += 1
            if r.success:
                by_rule[r.rule_name]["success"] += 1
            else:
                by_rule[r.rule_name]["failure"] += 1
                failures += 1
        return {
            "by_rule": by_rule,
            "total": total,
            "failures": failures,
        }
