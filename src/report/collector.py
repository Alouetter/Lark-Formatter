"""报告收集器：从 ChangeTracker 和 ValidationRule 收集结构化报告数据"""

from dataclasses import dataclass, field
from src.engine.change_tracker import ChangeTracker


@dataclass
class ReportData:
    """结构化报告数据"""
    scene_name: str = ""
    input_file: str = ""
    total_changes: int = 0
    total_failures: int = 0
    changes_by_rule: dict = field(default_factory=dict)
    validation_issues: list = field(default_factory=list)
    summary: dict = field(default_factory=dict)
    formula_stats: dict = field(default_factory=dict)
    low_confidence_items: list = field(default_factory=list)
    formula_diagnostics: list = field(default_factory=list)


def collect_report(tracker: ChangeTracker,
                   scene_name: str = "",
                   input_file: str = "",
                   validation_issues: list = None) -> ReportData:
    """从 ChangeTracker 收集报告数据"""
    report = ReportData(
        scene_name=scene_name,
        input_file=input_file,
    )

    # 按规则分组
    rules_seen = {}
    for rec in tracker.records:
        rule = rec.rule_name
        if rule not in rules_seen:
            rules_seen[rule] = []
        rules_seen[rule].append({
            "target": rec.target,
            "section": rec.section,
            "type": rec.change_type,
            "before": rec.before,
            "after": rec.after,
            "success": rec.success,
            "failure": rec.failure_reason,
        })

    report.changes_by_rule = rules_seen
    report.total_changes = len(tracker.records)
    report.total_failures = len(tracker.get_failures())
    report.validation_issues = [
        {"level": v.level, "rule": v.rule_name,
         "message": v.message, "location": v.location}
        for v in (validation_issues or [])
    ]
    report.summary = tracker.summary()

    formula_rules = {"formula_convert", "formula_to_table", "equation_table_format", "formula_style"}
    formula_stats: dict[str, str] = {}
    low_confidence_items: list[dict] = []
    formula_diagnostics: list[dict] = []
    for rec in tracker.records:
        if rec.rule_name in formula_rules and rec.target == "summary":
            formula_stats[rec.rule_name] = rec.after or ""
        if rec.rule_name in formula_rules and rec.change_type == "skip" and "低置信" in (rec.after or ""):
            low_confidence_items.append({
                "rule": rec.rule_name,
                "target": rec.target,
                "section": rec.section,
                "paragraph_index": rec.paragraph_index,
                "reason": rec.after,
                "suggestion": "建议人工复核该公式后再处理",
            })
        if rec.rule_name in formula_rules and rec.change_type == "diagnostic":
            formula_diagnostics.append({
                "rule": rec.rule_name,
                "target": rec.target,
                "section": rec.section,
                "paragraph_index": rec.paragraph_index,
                "reason": rec.after,
            })
    report.formula_stats = formula_stats
    report.low_confidence_items = low_confidence_items
    report.formula_diagnostics = formula_diagnostics
    return report
