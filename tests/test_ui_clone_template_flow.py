import copy
import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QFileDialog, QInputDialog, QMessageBox

from src.scene.manager import load_scene_from_data
from src.ui.main_window import MainWindow


def _app():
    return QApplication.instance() or QApplication([])


def _make_scene(font_cn: str, *, bold: bool) -> object:
    return load_scene_from_data(
        {
            "styles": {
                "normal": {
                    "font_cn": font_cn,
                    "font_en": f"{font_cn} EN",
                    "bold": bold,
                    "alignment": "justify",
                }
            }
        }
    )


def _stub_clone_dialog_runtime(window: MainWindow, monkeypatch, *, name: str = "克隆自 示例文档"):
    monkeypatch.setattr(
        QFileDialog,
        "getOpenFileName",
        staticmethod(lambda *args, **kwargs: ("C:/tmp/source.docx", "Word 文档 (*.docx)")),
    )
    monkeypatch.setattr(
        QInputDialog,
        "getText",
        staticmethod(lambda *args, **kwargs: (name, True)),
    )
    monkeypatch.setattr(
        QMessageBox,
        "information",
        staticmethod(lambda *args, **kwargs: QMessageBox.Ok),
    )
    monkeypatch.setattr(
        QMessageBox,
        "critical",
        staticmethod(lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError(args[2] if len(args) > 2 else "unexpected critical"))),
    )
    monkeypatch.setattr(window, "_apply_config_to_controls", lambda: None)
    monkeypatch.setattr(window, "_refresh_ui_panels", lambda: None)
    monkeypatch.setattr(window, "_populate_scene_combo", lambda: None)
    monkeypatch.setattr(window, "_save_ui_state", lambda: None)


def test_clone_template_uses_default_scene_as_base(monkeypatch):
    _app()
    window = MainWindow()

    current_scene = _make_scene("DirtyFont", bold=True)
    default_scene = _make_scene("BaseFont", bold=False)
    default_scene.format_signature = "已署名"
    window._config = current_scene

    _stub_clone_dialog_runtime(window, monkeypatch)

    monkeypatch.setattr(
        "src.ui.main_window.load_default_scene",
        lambda: copy.deepcopy(default_scene),
    )

    def _fake_clone(config, _path):
        config.styles["normal"].alignment = "center"
        return {
            "styles_updated": ["normal"],
            "styles_missing": [],
            "styles_added": [],
            "page_setup_updated": False,
            "numbering_updated": False,
            "numbering_levels_updated": [],
            "numbering_fallback_used": False,
            "numbering_inferred": False,
        }

    monkeypatch.setattr("src.ui.main_window.clone_scene_style_from_docx", _fake_clone)

    saved = {}

    def _fake_save(config, path):
        saved["config"] = copy.deepcopy(config)
        saved["path"] = path

    monkeypatch.setattr("src.ui.main_window.save_scene", _fake_save)

    window._clone_word_template_style()

    assert "config" in saved
    assert saved["config"].styles["normal"].font_cn == "BaseFont"
    assert saved["config"].styles["normal"].bold is False
    assert saved["config"].styles["normal"].alignment == "center"
    assert saved["config"].format_signature == ""
    assert "DirtyFont" not in saved["config"].styles["normal"].font_cn

    window.close()


def test_clone_template_reentrant_call_saves_only_once(monkeypatch):
    _app()
    window = MainWindow()

    window._config = _make_scene("DirtyFont", bold=True)
    default_scene = _make_scene("BaseFont", bold=False)

    open_calls = {"count": 0}
    monkeypatch.setattr(
        QFileDialog,
        "getOpenFileName",
        staticmethod(lambda *args, **kwargs: (open_calls.__setitem__("count", open_calls["count"] + 1) or "C:/tmp/source.docx", "Word 文档 (*.docx)")),
    )
    monkeypatch.setattr(
        QInputDialog,
        "getText",
        staticmethod(lambda *args, **kwargs: ("克隆自 示例文档", True)),
    )
    monkeypatch.setattr(
        QMessageBox,
        "information",
        staticmethod(lambda *args, **kwargs: QMessageBox.Ok),
    )
    monkeypatch.setattr(
        QMessageBox,
        "critical",
        staticmethod(lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError(args[2] if len(args) > 2 else "unexpected critical"))),
    )
    monkeypatch.setattr(window, "_apply_config_to_controls", lambda: None)
    monkeypatch.setattr(window, "_refresh_ui_panels", lambda: None)
    monkeypatch.setattr(window, "_populate_scene_combo", lambda: None)
    monkeypatch.setattr(window, "_save_ui_state", lambda: None)
    monkeypatch.setattr(
        "src.ui.main_window.load_default_scene",
        lambda: copy.deepcopy(default_scene),
    )

    nested = {"done": False}

    def _fake_clone(config, _path):
        if not nested["done"]:
            nested["done"] = True
            window._clone_word_template_style()
        config.styles["normal"].alignment = "center"
        return {
            "styles_updated": ["normal"],
            "styles_missing": [],
            "styles_added": [],
            "page_setup_updated": False,
            "numbering_updated": False,
            "numbering_levels_updated": [],
            "numbering_fallback_used": False,
            "numbering_inferred": False,
        }

    monkeypatch.setattr("src.ui.main_window.clone_scene_style_from_docx", _fake_clone)

    save_calls = {"count": 0}

    def _fake_save(_config, _path):
        save_calls["count"] += 1

    monkeypatch.setattr("src.ui.main_window.save_scene", _fake_save)

    window._clone_word_template_style()

    assert open_calls["count"] == 1
    assert save_calls["count"] == 1
    assert window._clone_in_progress is False

    window.close()
