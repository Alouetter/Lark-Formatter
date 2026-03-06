"""JSON 报告生成"""

import json
from src.report.collector import ReportData


def generate_json_report(report: ReportData, output_path: str) -> None:
    """将报告数据输出为 JSON 文件"""
    data = {
        "scene": report.scene_name,
        "input_file": report.input_file,
        "total_changes": report.total_changes,
        "total_failures": report.total_failures,
        "summary": report.summary,
        "changes_by_rule": report.changes_by_rule,
        "validation_issues": report.validation_issues,
    }
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
