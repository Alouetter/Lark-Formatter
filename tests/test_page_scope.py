from docx import Document

from src.engine.page_scope import (
    page_number_to_start_paragraph_index,
    page_ranges_to_paragraph_ranges,
)


def test_page_ranges_use_word_page_spans_with_cross_page_paragraph(monkeypatch):
    monkeypatch.setattr(
        "src.engine.page_scope._resolve_word_paragraph_page_spans",
        lambda doc, source_doc_path=None, timeout_sec=None: [
            (1, 1),
            (1, 2),
            (2, 2),
            (3, 3),
        ],
    )

    doc = Document()
    for text in ("p0", "p1", "p2", "p3"):
        doc.add_paragraph(text)

    assert page_ranges_to_paragraph_ranges(doc, [(2, 2)], require_word=True) == [(1, 2)]
    assert page_number_to_start_paragraph_index(doc, 2, require_word=True) == 1
