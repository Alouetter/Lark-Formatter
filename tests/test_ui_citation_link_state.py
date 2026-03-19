import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from src.scene.manager import load_scene_from_data
from src.ui.main_window import MainWindow


def _app():
    return QApplication.instance() or QApplication([])


def test_main_window_applies_citation_link_controls_from_scene_config(monkeypatch):
    _app()
    monkeypatch.setattr(MainWindow, "_load_ui_state", lambda self: {})
    monkeypatch.setattr(MainWindow, "_save_ui_state", lambda self: None)

    window = MainWindow()
    try:
        cfg = load_scene_from_data(
            {
                "pipeline": ["page_setup", "table_format", "citation_link", "validation"],
                "citation_link": {
                    "enabled": True,
                    "auto_number_reference_entries": False,
                    "superscript_outer_page_numbers": True,
                },
            }
        )
        window._config = cfg
        window._apply_config_to_controls()

        assert window._citation_link_check.isChecked() is True
        assert window._citation_ref_auto_number_check.isChecked() is False
        assert window._citation_outer_page_sup_check.isChecked() is True
    finally:
        window.close()


def test_ui_state_controls_do_not_override_citation_link_scene_config(monkeypatch):
    _app()
    monkeypatch.setattr(MainWindow, "_load_ui_state", lambda self: {})
    monkeypatch.setattr(MainWindow, "_save_ui_state", lambda self: None)

    window = MainWindow()
    try:
        cfg = load_scene_from_data(
            {
                "pipeline": ["page_setup", "table_format", "citation_link", "validation"],
                "citation_link": {
                    "enabled": True,
                    "auto_number_reference_entries": True,
                    "superscript_outer_page_numbers": False,
                },
            }
        )
        window._config = cfg
        window._apply_config_to_controls()

        window._apply_ui_state_controls(
            {
                "citation_link_restore": False,
                "citation_link_options": {
                    "auto_number_reference_entries": False,
                    "superscript_outer_page_numbers": True,
                },
            }
        )

        assert window._citation_link_check.isChecked() is True
        assert window._citation_ref_auto_number_check.isChecked() is True
        assert window._citation_outer_page_sup_check.isChecked() is False
    finally:
        window.close()
