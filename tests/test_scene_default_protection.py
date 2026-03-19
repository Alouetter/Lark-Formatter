import json

from src.scene.manager import _ensure_templates, load_scene_from_data, save_scene


def test_ensure_templates_repairs_semantically_empty_default_scene(tmp_path):
    builtin_dir = tmp_path / "builtin"
    templates_dir = tmp_path / "templates"
    builtin_dir.mkdir()
    templates_dir.mkdir()

    seed_payload = {
        "name": "默认格式",
        "heading_numbering": {
            "levels": {
                "heading1": {"format": "arabic", "template": "{current}", "separator": " "}
            }
        },
        "styles": {
            "normal": {"font_cn": "宋体", "font_en": "Times New Roman", "size_pt": 12}
        },
    }
    broken_payload = {
        "name": "",
        "heading_numbering": {"levels": {}},
        "styles": {},
    }

    (builtin_dir / "default_format.json").write_text(
        json.dumps(seed_payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (templates_dir / "default_format.json").write_text(
        json.dumps(broken_payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    _ensure_templates(builtin_dir, templates_dir)

    repaired = json.loads((templates_dir / "default_format.json").read_text(encoding="utf-8"))
    assert repaired == seed_payload
    assert list(templates_dir.glob("default_format.json.broken*"))


def test_default_scene_hides_cover_option_and_keeps_cover_protected() -> None:
    cfg = load_scene_from_data({})

    assert "cover" not in cfg.available_sections
    assert cfg.format_scope.sections.get("cover") is None
    assert cfg.format_scope.is_section_enabled("cover") is False


def test_loading_or_saving_scene_cannot_reenable_cover_scope(tmp_path) -> None:
    cfg = load_scene_from_data(
        {
            "available_sections": ["body", "cover", "references"],
            "format_scope": {
                "sections": {
                    "body": True,
                    "cover": True,
                    "references": True,
                }
            },
        }
    )

    assert cfg.available_sections == ["body", "references"]
    assert cfg.format_scope.sections.get("cover") is None
    assert cfg.format_scope.is_section_enabled("cover") is False

    cfg.available_sections.append("cover")
    cfg.format_scope.sections["cover"] = True
    output_path = tmp_path / "scene.json"
    save_scene(cfg, output_path)

    saved = json.loads(output_path.read_text(encoding="utf-8"))
    assert "cover" not in saved["available_sections"]
    assert "cover" not in saved["format_scope"]["sections"]
