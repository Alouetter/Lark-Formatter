from types import SimpleNamespace

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from src.engine.change_tracker import ChangeTracker
from src.engine.rules.heading_numbering import HeadingNumberingRule
from src.engine.rules.style_manager import _apply_style_config
from src.scene.manager import load_scene_from_data
from src.scene.schema import HeadingLevelConfig, StyleConfig

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _seed_heading1_numbering(cfg) -> None:
    cfg.heading_numbering.levels = {
        "heading1": HeadingLevelConfig(format="arabic", template="{current}", separator=" "),
    }
    binding = cfg.heading_numbering_v2.level_bindings["heading1"]
    binding.enabled = True
    binding.display_shell = "plain"
    binding.display_core_style = "arabic"
    binding.reference_core_style = "arabic"
    binding.chain = "current_only"
    binding.title_separator = " "
    binding.start_at = 1
    binding.restart_on = None


@pytest.mark.parametrize("mode", ["A", "B"])
def test_heading_numbering_rule_clears_direct_paragraph_and_run_overrides(mode):
    doc = Document()
    _apply_style_config(
        doc.styles["Heading 1"],
        StyleConfig(
            font_cn="宋体",
            font_en="Times New Roman",
            size_pt=12,
            bold=True,
            italic=True,
            alignment="left",
            space_before_pt=0,
            space_after_pt=0,
        ),
    )

    para = doc.add_paragraph("1 标题")
    para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    para.paragraph_format.space_before = Pt(24)
    para.paragraph_format.space_after = Pt(18)
    para.paragraph_format.left_indent = Pt(36)
    run = para.runs[0]
    run.bold = False
    run.italic = False

    cfg = load_scene_from_data({})
    cfg.heading_numbering.mode = mode
    _seed_heading1_numbering(cfg)

    HeadingNumberingRule().apply(
        doc,
        cfg,
        ChangeTracker(),
        {"headings": [SimpleNamespace(para_index=0, level="heading1", text="1 标题")]},
    )

    ppr = para._element.find(_w("pPr"))
    assert ppr is not None
    assert ppr.find(_w("spacing")) is None
    assert ppr.find(_w("jc")) is None
    assert ppr.find(_w("ind")) is None

    rpr = run._element.find(_w("rPr"))
    if rpr is not None:
        for tag in ("rFonts", "color", "b", "bCs", "i", "iCs", "sz", "szCs"):
            assert rpr.find(_w(tag)) is None, (mode, tag)

    assert para.style.name.lower() == "heading 1"
