"""Markdown 报告生成"""

from src.report.collector import ReportData


def _escape_md_table_cell(value, *, limit: int | None = None) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("|", "\\|")
    text = text.replace("\n", "<br>")
    if limit is not None:
        text = text[:limit]
    return text


def generate_markdown_report(report: ReportData, output_path: str) -> None:
    """将报告数据输出为 Markdown 文件"""
    lines = []
    lines.append("# 排版报告")
    lines.append("")
    lines.append(f"- 场景: {report.scene_name}")
    lines.append(f"- 输入文件: {report.input_file}")
    lines.append(f"- 总修改数: {report.total_changes}")
    lines.append(f"- 失败数: {report.total_failures}")
    lines.append("")

    lines.append("## 修改详情")
    lines.append("")
    for rule_name, changes in report.changes_by_rule.items():
        lines.append(f"### {rule_name} ({len(changes)} 项)")
        lines.append("")
        lines.append("| 目标 | 类型 | 修改前 | 修改后 | 状态 |")
        lines.append("|------|------|--------|--------|------|")
        for c in changes:
            status = "✓" if c["success"] else f"✗ {c['failure'] or ''}"
            target = _escape_md_table_cell(c["target"])
            change_type = _escape_md_table_cell(c["type"])
            before = _escape_md_table_cell(c["before"], limit=40)
            after = _escape_md_table_cell(c["after"], limit=40)
            status = _escape_md_table_cell(status)
            lines.append(
                f"| {target} | {change_type} "
                f"| {before} | {after} | {status} |"
            )
        lines.append("")

    if report.formula_stats:
        lines.append("## 公式处理统计")
        lines.append("")
        for rule_name, summary in report.formula_stats.items():
            lines.append(f"- {rule_name}: {summary}")
        lines.append("")

    if report.low_confidence_items:
        lines.append("## 低置信公式清单")
        lines.append("")
        for item in report.low_confidence_items:
            lines.append(
                f"- [{item['rule']}] {item['target']} (段落 {item['paragraph_index']}): "
                f"{item['reason']}；{item['suggestion']}"
            )
        lines.append("")

    if report.formula_diagnostics:
        lines.append("## 公式诊断")
        lines.append("")
        for item in report.formula_diagnostics:
            lines.append(
                f"- [{item['rule']}] {item['target']} (段落 {item['paragraph_index']}): "
                f"{item['reason']}"
            )
        lines.append("")

    if report.validation_issues:
        lines.append("## 校验问题")
        lines.append("")
        for v in report.validation_issues:
            icon = "⚠️" if v["level"] == "warning" else "❌"
            lines.append(f"- {icon} [{v['level']}] {v['message']} @ {v['location']}")
        lines.append("")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
