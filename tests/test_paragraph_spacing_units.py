from lxml import etree

from src.docx_io.style_clone import _read_xml_spacing_info
from src.scene.manager import load_scene_from_data
from src.scene.schema import StyleConfig
from src.utils.line_spacing import sync_spacing_ooxml, sync_style_spacing_ooxml

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _spacing(el):
    return el.find(_w("pPr")).find(_w("spacing"))


def test_sync_spacing_ooxml_writes_physical_units_as_twips():
    p = etree.Element(_w("p"))

    sync_spacing_ooxml(
        p,
        space_before_value=0.5,
        space_before_unit="cm",
        space_after_value=3,
        space_after_unit="mm",
        line_spacing_type="single",
        line_spacing_value=1.0,
    )

    spacing = _spacing(p)
    assert spacing.get(_w("before")) == "283"
    assert spacing.get(_w("after")) == "170"
    assert spacing.get(_w("beforeLines")) is None
    assert spacing.get(_w("afterAutospacing")) is None


def test_sync_spacing_ooxml_writes_line_and_auto_units():
    p = etree.Element(_w("p"))

    sync_spacing_ooxml(
        p,
        space_before_value=1.5,
        space_before_unit="line",
        space_after_value=0,
        space_after_unit="auto",
        line_spacing_type="single",
        line_spacing_value=1.0,
    )

    spacing = _spacing(p)
    assert spacing.get(_w("beforeLines")) == "150"
    assert spacing.get(_w("before")) is None
    assert spacing.get(_w("afterAutospacing")) == "1"
    assert spacing.get(_w("after")) is None


def test_style_config_legacy_spacing_defaults_to_pt():
    cfg = load_scene_from_data(
        {
            "styles": {
                "normal": {
                    "space_before_pt": 12.0,
                    "space_after_pt": 6.0,
                }
            }
        }
    )

    sc = cfg.styles["normal"]
    assert sc.space_before_value == 12.0
    assert sc.space_before_unit == "pt"
    assert sc.space_after_value == 6.0
    assert sc.space_after_unit == "pt"


def test_style_spacing_ooxml_preserves_line_unit():
    sc = StyleConfig(space_before_value=1.0, space_before_unit="line")
    p = etree.Element(_w("p"))

    sync_style_spacing_ooxml(p, sc, line_spacing_type="single", line_spacing_value=1.0)

    spacing = _spacing(p)
    assert spacing.get(_w("beforeLines")) == "100"
    assert spacing.get(_w("before")) is None


def test_clone_spacing_reader_preserves_line_and_auto_units():
    p = etree.Element(_w("p"))
    ppr = etree.SubElement(p, _w("pPr"))
    spacing = etree.SubElement(ppr, _w("spacing"))
    spacing.set(_w("beforeLines"), "125")
    spacing.set(_w("afterAutospacing"), "1")

    info = _read_xml_spacing_info(type("Obj", (), {"_element": p})())

    assert info["before_explicit"] is True
    assert info["before_unit"] == "line"
    assert info["before_value"] == 1.25
    assert info["after_explicit"] is True
    assert info["after_unit"] == "auto"
