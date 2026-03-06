"""Markdown 报告生成"""

from src.report.collector import ReportData


def generate_markdown_report(report: ReportData, output_path: str) -> None:
    """将报告数据输出为 Markdown 文件"""
    lines = []
    lines.append(f"# 排版报告")
    lines.append("")
    lines.append(f"- 场景: {report.scene_name}")
    lines.append(f"- 输入文件: {report.input_file}")
    lines.append(f"- 总修改数: {report.total_changes}")
    lines.append(f"- 失败数: {report.total_failures}")
    lines.append("")

    # 按规则输出修改详情
    lines.append("## 修改详情")
    lines.append("")
    for rule_name, changes in report.changes_by_rule.items():
        lines.append(f"### {rule_name} ({len(changes)} 项)")
        lines.append("")
        lines.append("| 目标 | 类型 | 修改前 | 修改后 | 状态 |")
        lines.append("|------|------|--------|--------|------|")
        for c in changes:
            status = "✓" if c["success"] else f"✗ {c['failure'] or ''}"
            before = (c["before"] or "")[:40].replace("|", "\\|")
            after = (c["after"] or "")[:40].replace("|", "\\|")
            lines.append(
                f"| {c['target']} | {c['type']} "
                f"| {before} | {after} | {status} |"
            )
        lines.append("")

    # 校验问题
    if report.validation_issues:
        lines.append("## 校验问题")
        lines.append("")
        for v in report.validation_issues:
            icon = "⚠️" if v["level"] == "warning" else "❌"
            lines.append(f"- {icon} [{v['level']}] {v['message']} @ {v['location']}")
        lines.append("")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
