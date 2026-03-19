from docx import Document

from src.engine.rules.style_manager import _apply_style_config
from src.scene.schema import StyleConfig

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def test_apply_style_config_writes_explicit_off_for_bold_and_italic():
    doc = Document()
    style = doc.styles["Heading 1"]

    _apply_style_config(
        style,
        StyleConfig(
            font_cn="宋体",
            font_en="Times New Roman",
            size_pt=12,
            bold=False,
            italic=False,
        ),
    )

    rpr = style.element.find(_w("rPr"))
    assert rpr is not None
    rfonts = rpr.find(_w("rFonts"))
    assert rfonts is not None
    assert rfonts.get(_w("hint")) == "default"
    for tag in ("b", "bCs", "i", "iCs"):
        el = rpr.find(_w(tag))
        assert el is not None, tag
        assert el.get(_w("val")) == "0", tag
