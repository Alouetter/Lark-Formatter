from src.scene.manager import load_scene_from_data


def test_scene_manager_loads_md_cleanup_list_style_fields():
    cfg = load_scene_from_data(
        {
            "md_cleanup": {
                "enabled": True,
                "list_marker_separator": "full_space",
                "ordered_list_style": "decimal_cn_dun",
                "unordered_list_style": "bullet_square",
            }
        }
    )

    assert cfg.md_cleanup.enabled is True
    assert cfg.md_cleanup.list_marker_separator == "full_space"
    assert cfg.md_cleanup.ordered_list_style == "decimal_cn_dun"
    assert cfg.md_cleanup.unordered_list_style == "bullet_square"


def test_scene_manager_invalid_md_cleanup_list_style_fields_fallback_to_default():
    cfg = load_scene_from_data(
        {
            "md_cleanup": {
                "ordered_list_style": "unknown_style",
                "unordered_list_style": "unknown_style",
            }
        }
    )

    assert cfg.md_cleanup.ordered_list_style == "mixed"
    assert cfg.md_cleanup.unordered_list_style == "word_default"


def test_scene_manager_loads_md_cleanup_formula_noise_switches():
    cfg = load_scene_from_data(
        {
            "md_cleanup": {
                "formula_copy_noise_cleanup": False,
                "suppress_formula_fake_lists": False,
            }
        }
    )

    assert cfg.md_cleanup.formula_copy_noise_cleanup is False
    assert cfg.md_cleanup.suppress_formula_fake_lists is False

    cfg_default = load_scene_from_data({"md_cleanup": {}})
    assert cfg_default.md_cleanup.formula_copy_noise_cleanup is True
    assert cfg_default.md_cleanup.suppress_formula_fake_lists is True
