from pathlib import Path

from src.report.collector import ReportData
from src.report.markdown_report import generate_markdown_report


def test_markdown_report_escapes_multiline_table_cells(tmp_path):
    report = ReportData(
        scene_name="scene",
        input_file="input.docx",
        changes_by_rule={
            "rule": [
                {
                    "target": "段落|1\n第二行",
                    "section": "global",
                    "type": "format\nfix",
                    "before": "old|\nvalue",
                    "after": "new\r\nvalue",
                    "success": False,
                    "failure": "line1\nline2",
                }
            ]
        },
    )
    output_path = tmp_path / "report.md"

    generate_markdown_report(report, str(output_path))

    text = output_path.read_text(encoding="utf-8")
    assert "| 段落\\|1<br>第二行 | format<br>fix | old\\|<br>value | new<br>value | ✗ line1<br>line2 |" in text
