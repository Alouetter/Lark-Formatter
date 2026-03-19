from src.scene.schema import HeadingLevelBindingConfig, HeadingLevelConfig
from src.utils.chinese import int_to_chinese
from src.utils.heading_numbering_template import (
    build_ooxml_lvl_text,
    render_heading_numbering,
    template_uses_legacy_parent,
    validate_heading_numbering_template,
)
from src.utils.ooxml import build_abstract_num

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def test_validate_heading_template_rejects_future_level_reference():
    errors = validate_heading_numbering_template(
        "heading2",
        "{level3}.{current}",
    )
    assert any("level3" in err for err in errors)


def test_render_heading_numbering_uses_referenced_level_native_format():
    levels_config = {
        "heading1": HeadingLevelConfig(
            format="chinese_chapter",
            template="第{current}章",
        ),
        "heading2": HeadingLevelConfig(
            format="arabic",
            template="{level1}.{current}",
        ),
    }
    rendered = render_heading_numbering(
        "heading2",
        levels_config["heading2"].template,
        levels_config,
        {"heading1": 3, "heading2": 2},
    )
    assert rendered == f"{int_to_chinese(3)}.2"


def test_legacy_parent_keeps_decimal_chain_behavior():
    levels_config = {
        "heading1": HeadingLevelConfig(format="chinese_chapter"),
        "heading2": HeadingLevelConfig(format="chinese_section"),
        "heading3": HeadingLevelConfig(
            format="arabic_dotted",
            template="{parent}.{current}",
        ),
    }
    assert template_uses_legacy_parent("heading3", "{parent}.{current}") is True
    assert build_ooxml_lvl_text("heading3", "{parent}.{current}") == "%1.%2.%3"
    rendered = render_heading_numbering(
        "heading3",
        levels_config["heading3"].template,
        levels_config,
        {"heading1": 2, "heading2": 4, "heading3": 1},
    )
    assert rendered == "2.4.1"


def test_build_abstract_num_keeps_explicit_level_refs_in_native_mode():
    levels_config = {
        "heading3": HeadingLevelConfig(
            format="arabic_dotted",
            template="{level1}.{level2}.{current}",
            separator="\u3000",
        )
    }

    abs_num = build_abstract_num(1, levels_config, {})
    lvl = abs_num.find(_w("lvl"))
    assert lvl is not None
    assert lvl.find(_w("isLgl")) is None
    lvl_text = lvl.find(_w("lvlText"))
    assert lvl_text is not None
    assert lvl_text.get(_w("val")) == "%1.%2.%3\u3000"


def test_build_abstract_num_marks_legacy_parent_refs_as_legal():
    levels_config = {
        "heading3": HeadingLevelConfig(
            format="arabic_dotted",
            template="{parent}.{current}",
            separator="\u3000",
        )
    }

    abs_num = build_abstract_num(1, levels_config, {})
    lvl = abs_num.find(_w("lvl"))
    assert lvl is not None
    assert lvl.find(_w("isLgl")) is not None


def test_build_abstract_num_respects_start_at_and_restart_on_from_v2_bindings():
    levels_config = {
        "heading1": HeadingLevelConfig(format="arabic", template="{current}", separator=" "),
        "heading2": HeadingLevelConfig(format="arabic_dotted", template="{parent}.{current}", separator=" "),
        "heading3": HeadingLevelConfig(format="arabic_dotted", template="{parent}.{current}", separator=" "),
    }
    level_bindings = {
        "heading1": HeadingLevelBindingConfig(start_at=3, restart_on=None),
        "heading2": HeadingLevelBindingConfig(start_at=5, restart_on=None),
        "heading3": HeadingLevelBindingConfig(start_at=7, restart_on="heading1"),
    }

    abs_num = build_abstract_num(1, levels_config, level_bindings=level_bindings)
    lvls = abs_num.findall(_w("lvl"))
    assert len(lvls) == 3

    lvl1 = lvls[0]
    lvl2 = lvls[1]
    lvl3 = lvls[2]

    assert lvl1.find(_w("start")).get(_w("val")) == "3"
    assert lvl2.find(_w("start")).get(_w("val")) == "5"
    assert lvl3.find(_w("start")).get(_w("val")) == "7"

    assert lvl1.find(_w("lvlRestart")) is None
    assert lvl2.find(_w("lvlRestart")).get(_w("val")) == "0"
    assert lvl3.find(_w("lvlRestart")).get(_w("val")) == "1"
