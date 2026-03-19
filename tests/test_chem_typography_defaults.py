from __future__ import annotations

import json

from src.scene.manager import load_scene
from src.scene.schema import SceneConfig


def test_scene_config_chem_typography_defaults_are_safe() -> None:
    cfg = SceneConfig()

    assert cfg.chem_typography.enabled is False
    assert cfg.chem_typography.scopes == {
        "references": False,
        "body": False,
        "headings": False,
        "abstract_cn": False,
        "abstract_en": False,
        "captions": False,
        "tables": False,
    }


def test_legacy_scene_without_chem_typography_field_does_not_enable_restore(tmp_path) -> None:
    scene_path = tmp_path / "legacy_scene.json"
    payload = {
        "name": "legacy",
        "page_setup": {},
        "heading_numbering": {},
        "heading_model": {},
        "styles": {},
        "format_scope": {"mode": "auto", "sections": {}},
    }
    scene_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

    cfg = load_scene(scene_path)

    assert cfg.chem_typography.enabled is False
    assert cfg.chem_typography.scopes["references"] is False
    assert cfg.chem_typography.scopes["captions"] is False
    assert cfg.chem_typography.scopes["tables"] is False
