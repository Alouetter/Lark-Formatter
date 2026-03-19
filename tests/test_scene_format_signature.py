from pathlib import Path

from src.scene.manager import load_scene, load_scene_from_data, save_scene


def test_scene_format_signature_round_trip(tmp_path: Path):
    scene_path = tmp_path / "signed_scene.json"
    cfg = load_scene_from_data(
        {
            "name": "\u6d4b\u8bd5\u683c\u5f0f",
            "format_signature": "\u5f20\u4e09",
        }
    )

    save_scene(cfg, scene_path)
    reloaded = load_scene(scene_path)

    assert reloaded.name == "\u6d4b\u8bd5\u683c\u5f0f"
    assert reloaded.format_signature == "\u5f20\u4e09"
