from src.scene.manager import load_scene_from_data
from src.scene.schema import default_chain_id_for_level


def test_scene_manager_derives_heading_numbering_v2_from_legacy_parent_templates():
    cfg = load_scene_from_data(
        {
            "heading_numbering": {
                "levels": {
                    "heading1": {
                        "format": "chinese_chapter",
                        "template": "第{cn}章",
                        "separator": "\u3000",
                    },
                    "heading2": {
                        "format": "arabic_dotted",
                        "template": "{parent}.{current}",
                        "separator": " ",
                    },
                }
            }
        }
    )

    bindings = cfg.heading_numbering_v2.level_bindings
    assert bindings["heading1"].enabled is True
    assert bindings["heading1"].display_shell == "chapter_cn"
    assert bindings["heading1"].display_core_style == "cn_lower"
    assert bindings["heading1"].reference_core_style == "arabic"

    assert bindings["heading2"].enabled is True
    assert bindings["heading2"].display_shell == "plain"
    assert bindings["heading2"].display_core_style == "arabic"
    assert bindings["heading2"].chain == default_chain_id_for_level("heading2")
    assert bindings["heading2"].title_separator == " "

    assert bindings["heading3"].enabled is False


def test_scene_manager_explicit_level_refs_keep_parent_reference_style():
    cfg = load_scene_from_data(
        {
            "heading_numbering": {
                "levels": {
                    "heading1": {
                        "format": "chinese_chapter",
                        "template": "第{current}章",
                        "separator": "\u3000",
                    },
                    "heading2": {
                        "format": "arabic_dotted",
                        "template": "{level1}.{current}",
                        "separator": "\u3000",
                    },
                }
            }
        }
    )

    bindings = cfg.heading_numbering_v2.level_bindings
    assert bindings["heading1"].reference_core_style == "cn_lower"
    assert bindings["heading2"].chain == default_chain_id_for_level("heading2")


def test_scene_manager_creates_custom_shell_entry_during_v2_migration():
    cfg = load_scene_from_data(
        {
            "heading_numbering": {
                "levels": {
                    "heading1": {
                        "format": "arabic",
                        "template": "【{current}】",
                        "separator": "\u3000",
                    }
                }
            }
        }
    )

    binding = cfg.heading_numbering_v2.level_bindings["heading1"]
    assert binding.display_shell.startswith("custom_shell_heading1")

    shell = cfg.heading_numbering_v2.shell_catalog[binding.display_shell]
    assert shell.prefix == "【"
    assert shell.suffix == "】"
    assert shell.label == "【{}】"


def test_scene_manager_falls_back_removed_legacy_core_style_to_arabic():
    cfg = load_scene_from_data(
        {
            "heading_numbering": {
                "levels": {
                    "heading1": {
                        "format": "alpha_lower",
                        "template": "{current}",
                        "separator": "\u3000",
                    }
                }
            }
        }
    )

    binding = cfg.heading_numbering_v2.level_bindings["heading1"]
    assert binding.display_core_style == "arabic"
    assert binding.reference_core_style == "arabic"


def test_scene_manager_prefers_explicit_heading_numbering_v2_payload():
    cfg = load_scene_from_data(
        {
            "heading_numbering": {
                "levels": {
                    "heading1": {
                        "format": "chinese_chapter",
                        "template": "第{current}章",
                        "separator": "\u3000",
                    }
                }
            },
            "heading_numbering_v2": {
                "core_style_catalog": {
                    "arabic_fullwidth": {
                        "label": "１ ２ ３",
                        "sample": "１",
                    },
                    "alpha_lower": {
                        "label": "a b c",
                        "sample": "a",
                    }
                },
                "level_bindings": {
                    "heading1": {
                        "enabled": True,
                        "display_shell": "plain",
                        "display_core_style": "arabic_fullwidth",
                        "reference_core_style": "alpha_lower",
                        "chain": "current_only",
                    }
                }
            },
        }
    )

    binding = cfg.heading_numbering_v2.level_bindings["heading1"]
    assert binding.display_shell == "plain"
    assert binding.display_core_style == "arabic"
    assert binding.reference_core_style == "arabic"
    assert "arabic_fullwidth" not in cfg.heading_numbering_v2.core_style_catalog
    assert "alpha_lower" not in cfg.heading_numbering_v2.core_style_catalog


def test_scene_manager_projects_explicit_v2_payload_back_to_runtime_legacy_levels():
    cfg = load_scene_from_data(
        {
            "heading_numbering_v2": {
                "shell_catalog": {
                    "chapter_custom": {
                        "label": "第{}章",
                        "prefix": "第",
                        "suffix": "章",
                    }
                },
                "level_bindings": {
                    "heading1": {
                        "enabled": True,
                        "display_shell": "chapter_custom",
                        "display_core_style": "arabic",
                        "reference_core_style": "arabic",
                        "chain": "current_only",
                        "title_separator": " ",
                    },
                    "heading2": {
                        "enabled": True,
                        "display_shell": "plain",
                        "display_core_style": "arabic",
                        "reference_core_style": "arabic",
                        "chain": default_chain_id_for_level("heading2"),
                        "title_separator": " ",
                    },
                },
            }
        }
    )

    assert cfg.heading_numbering.levels["heading1"].format == "arabic"
    assert cfg.heading_numbering.levels["heading1"].template == "第{current}章"
    assert cfg.heading_numbering.levels["heading1"].separator == " "

    assert cfg.heading_numbering.levels["heading2"].format == "arabic_dotted"
    assert cfg.heading_numbering.levels["heading2"].template == "{parent}.{current}"
    assert cfg.heading_numbering.levels["heading2"].separator == " "


def test_scene_manager_preserves_legacy_heading_layout_fields_when_v2_payload_exists():
    cfg = load_scene_from_data(
        {
            "heading_numbering": {
                "levels": {
                    "heading1": {
                        "format": "chinese_chapter",
                        "template": "第{cn}章",
                        "separator": "\u3000",
                        "alignment": "center",
                        "left_indent_chars": 2.0,
                    }
                }
            },
            "heading_numbering_v2": {
                "level_bindings": {
                    "heading1": {
                        "enabled": True,
                        "display_shell": "plain",
                        "display_core_style": "arabic",
                        "reference_core_style": "arabic",
                        "chain": "current_only",
                        "title_separator": " ",
                    }
                }
            },
        }
    )

    level = cfg.heading_numbering.levels["heading1"]
    assert level.format == "arabic"
    assert level.template == "{current}"
    assert level.separator == " "
    assert level.alignment == "center"
    assert level.left_indent_chars == 2.0


def test_scene_manager_normalizes_invalid_v2_shell_and_chain_refs():
    cfg = load_scene_from_data(
        {
            "heading_numbering_v2": {
                "level_bindings": {
                    "heading1": {
                        "enabled": True,
                        "display_shell": "missing_shell",
                        "display_core_style": "arabic",
                        "reference_core_style": "arabic",
                        "chain": "missing_chain",
                        "title_separator": " ",
                    }
                }
            }
        }
    )

    binding = cfg.heading_numbering_v2.level_bindings["heading1"]
    assert binding.display_shell == "plain"
    assert binding.chain == "current_only"
    assert cfg.heading_numbering.levels["heading1"].format == "arabic"
    assert cfg.heading_numbering.levels["heading1"].template == "{current}"
    assert cfg.heading_numbering.levels["heading1"].separator == " "
