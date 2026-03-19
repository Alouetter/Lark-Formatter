import json
from pathlib import Path

from src.scene.manager import load_default_scene


def _load_formula_table(path: Path) -> dict:
    payload = json.loads(path.read_text(encoding="utf-8"))
    return dict(payload.get("formula_table", {}))


def test_default_scene_formula_table_matches_builtin_preset() -> None:
    template_formula_table = _load_formula_table(Path("templates/default_format.json"))
    builtin_formula_table = _load_formula_table(Path("src/scene/presets/default_format.json"))

    assert template_formula_table == builtin_formula_table


def test_load_default_scene_uses_expected_formula_math_defaults() -> None:
    cfg = load_default_scene()

    assert cfg.formula_table.formula_font_name == "Cambria Math"
    assert cfg.formula_table.formula_font_size_pt == 12.0
    assert cfg.formula_table.formula_space_before_pt == 0.0
    assert cfg.formula_table.formula_space_after_pt == 0.0
