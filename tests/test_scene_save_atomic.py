from __future__ import annotations

import json
from pathlib import Path

import pytest

from src.scene.manager import load_scene_from_data, rename_scene, save_scene


def test_save_scene_writes_valid_json_atomically(tmp_path):
    scene_path = tmp_path / "atomic_scene.json"
    cfg = load_scene_from_data(
        {
            "name": "原子保存测试",
            "styles": {
                "normal": {
                    "font_cn": "仿宋",
                    "bold": True,
                }
            },
        }
    )

    save_scene(cfg, scene_path)

    payload = json.loads(scene_path.read_text(encoding="utf-8"))
    assert payload["name"] == "原子保存测试"
    assert payload["styles"]["normal"]["font_cn"] == "仿宋"
    assert payload["styles"]["normal"]["bold"] is True
    assert not list(tmp_path.glob("*.tmp"))


def test_save_scene_replace_failure_keeps_original_file(monkeypatch, tmp_path):
    scene_path = tmp_path / "atomic_scene.json"
    original_payload = {"name": "旧内容", "styles": {"normal": {"font_cn": "宋体"}}}
    scene_path.write_text(json.dumps(original_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    cfg = load_scene_from_data(
        {
            "name": "新内容",
            "styles": {
                "normal": {
                    "font_cn": "黑体",
                    "bold": True,
                }
            },
        }
    )

    def _fail_replace(_src, _dst):
        raise PermissionError("replace blocked")

    monkeypatch.setattr("src.scene.manager.os.replace", _fail_replace)

    with pytest.raises(PermissionError):
        save_scene(cfg, scene_path)

    payload = json.loads(scene_path.read_text(encoding="utf-8"))
    assert payload == original_payload
    assert not [p for p in tmp_path.iterdir() if p.name.endswith(".tmp")]


def test_rename_scene_write_failure_keeps_original_file(monkeypatch, tmp_path):
    scene_path = tmp_path / "original_scene.json"
    original_payload = {"name": "original_scene", "styles": {"normal": {"font_cn": "宋体"}}}
    scene_path.write_text(json.dumps(original_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def _fail_replace(_src, _dst):
        raise PermissionError("replace blocked")

    monkeypatch.setattr("src.scene.manager.os.replace", _fail_replace)

    with pytest.raises(PermissionError):
        rename_scene(scene_path, "renamed_scene")

    assert scene_path.exists()
    assert json.loads(scene_path.read_text(encoding="utf-8")) == original_payload
    assert not (tmp_path / "renamed_scene.json").exists()
    assert not [p for p in tmp_path.iterdir() if p.name.endswith(".tmp")]
