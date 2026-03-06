from pathlib import Path

from docx import Document

from src.engine.pipeline import Pipeline
from src.scene.schema import SceneConfig


def test_pipeline_progress_is_monotonic_and_chinese(tmp_path: Path, monkeypatch):
    monkeypatch.setenv("DOCX_DISABLE_FIELD_REFRESH", "1")

    doc_path = tmp_path / "input.docx"
    doc = Document()
    doc.add_paragraph("测试正文")
    doc.save(doc_path)

    cfg = SceneConfig()
    cfg.pipeline = []
    cfg.output.final_docx = False
    cfg.output.compare_docx = False
    cfg.output.report_json = False
    cfg.output.report_markdown = False

    progress_events: list[tuple[int, int, str]] = []
    result = Pipeline(
        cfg,
        progress_callback=lambda current, total, message: progress_events.append(
            (current, total, message)
        ),
    ).run(str(doc_path))

    assert result.success is True
    assert progress_events
    current_values = [event[0] for event in progress_events]
    assert current_values == sorted(current_values)
    assert progress_events[-1][2] == "正在生成处理报告…"
    assert all("正在" in message for _, _, message in progress_events)
