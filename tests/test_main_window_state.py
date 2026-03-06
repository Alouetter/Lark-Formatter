import sys
from pathlib import Path

import src.ui.main_window as main_window_module
from src.ui.main_window import MainWindow
from src.scene.schema import SceneConfig


def test_normalized_scene_state_migrates_legacy_source_preset(
    tmp_path: Path, monkeypatch
):
    templates_dir = tmp_path / "templates"
    templates_dir.mkdir(parents=True, exist_ok=True)
    preset_path = templates_dir / "default_format.json"
    preset_path.write_text("{}", encoding="utf-8")

    win = MainWindow.__new__(MainWindow)
    win._scene_item_paths = {0: preset_path}
    win._current_scene_path = str(
        tmp_path / "project" / "src" / "scene" / "presets" / "default_format.json"
    )
    win._current_scene_is_custom = True

    monkeypatch.setattr(sys, "frozen", True, raising=False)
    path, is_custom = win._normalized_scene_state_for_storage()

    assert path == str(preset_path)
    assert is_custom is False


def test_canonicalize_scene_path_migrates_legacy_file(
    tmp_path: Path, monkeypatch
):
    templates_dir = tmp_path / "templates"
    templates_dir.mkdir(parents=True, exist_ok=True)
    legacy_file = tmp_path / "repo" / "src" / "scene" / "presets" / "custom_a.json"
    legacy_file.parent.mkdir(parents=True, exist_ok=True)
    legacy_file.write_text('{"name":"custom_a"}', encoding="utf-8")

    win = MainWindow.__new__(MainWindow)
    win._scene_item_paths = {}
    win._log = lambda *_args, **_kwargs: None

    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(main_window_module, "PRESETS_DIR", templates_dir)

    target = win._canonicalize_scene_path(legacy_file, migrate=True)
    assert target == templates_dir / "custom_a.json"
    assert target.exists()


def test_save_current_scene_config_respects_suspend_flag(tmp_path: Path):
    out_file = tmp_path / "scene.json"

    win = MainWindow.__new__(MainWindow)
    win._config = SceneConfig(name="demo")
    win._current_scene_path = str(out_file)
    win._is_restoring_state = False
    win._suspend_scene_autosave = True
    win._sync_all_controls_to_config = lambda: None
    win._canonicalize_scene_path = lambda p, migrate: p
    win._log = lambda *_args, **_kwargs: None

    win._save_current_scene_config()
    assert out_file.exists() is False
