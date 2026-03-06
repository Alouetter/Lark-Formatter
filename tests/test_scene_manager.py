from pathlib import Path

import pytest

from src.scene import manager
from src.scene.schema import SceneConfig


def test_safe_filename_replaces_illegal_chars():
    assert manager._safe_filename('  a<>:"/\\|?*b..  ') == "a_b"
    assert manager._safe_filename("   ") == "unnamed"


def test_load_scene_rejects_outside_format_file(tmp_path: Path):
    with pytest.raises(ValueError, match="outside preset directory"):
        manager.load_scene_from_data(
            {"name": "x", "format_file": "../outside.json"},
            base_dir=tmp_path,
        )


def test_load_scene_merges_relative_format_file(tmp_path: Path):
    fmt = tmp_path / "base.json"
    fmt.write_text(
        '{"name":"基础模板","styles":{"normal":{"font_cn":"黑体"}}}',
        encoding="utf-8",
    )
    cfg = manager.load_scene_from_data(
        {"name": "覆盖模板", "format_file": "base.json"},
        base_dir=tmp_path,
    )
    assert cfg.name == "覆盖模板"
    assert cfg.styles["normal"].font_cn == "黑体"


def test_save_scene_creates_parent_directory(tmp_path: Path):
    scene_path = tmp_path / "nested" / "custom_scene.json"
    cfg = SceneConfig(name="新场景")
    manager.save_scene(cfg, scene_path)
    assert scene_path.exists()
