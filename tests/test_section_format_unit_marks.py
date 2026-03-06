from __future__ import annotations

import pytest

from src.engine.rules.section_format import _build_chem_style_marks


def _superscript_positions(text: str) -> list[int]:
    marks = _build_chem_style_marks(text)
    return [idx for idx, style in enumerate(marks) if style == "superscript"]


def test_unit_exponents_preserved_with_cjk_tail() -> None:
    text = "mmol h-1 g-1\u8fd9\u4e2a\u53c2\u6570"
    assert _superscript_positions(text) == [6, 7, 10, 11]

    text_single = "cm-2\u8fd9\u4e2a"
    assert _superscript_positions(text_single) == [2, 3]


@pytest.mark.parametrize("text", ["(cm-2)", "\uff08cm-2\uff09", "[cm-2]", "{cm-2}"])
def test_wrapped_unit_exponents_are_detected(text: str) -> None:
    assert _superscript_positions(text) == [3, 4]


def test_multiple_unit_exponents_are_detected() -> None:
    assert _superscript_positions("mol m-2 s-1") == [5, 6, 9, 10]


@pytest.mark.parametrize(
    "text",
    [
        "D1-D4",
        "200k-1.1M",
        "Buck2 Documentation",
        "(D1-D4)",
        "[200k-1.1M]",
        "(A-1)",
    ],
)
def test_non_unit_tokens_remain_plain_text(text: str) -> None:
    assert _superscript_positions(text) == []
