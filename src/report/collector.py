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
    return report
