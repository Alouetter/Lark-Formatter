import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from src.scene.manager import load_scene_from_data
from src.ui.main_window import FormatConfigDialog
from src.ui.font_sizes import (
    NUMERIC_FONT_SIZE_OPTIONS,
    font_size_display_text,
    format_font_size_pt,
    is_named_font_size_input,
    parse_font_size_input,
)


def test_parse_named_font_sizes():
    assert parse_font_size_input("小四") == 12.0
    assert parse_font_size_input("五号") == 10.5
    assert parse_font_size_input("初号") == 42.0


def test_parse_numeric_font_sizes():
    assert parse_font_size_input("12") == 12.0
    assert parse_font_size_input("10.5pt") == 10.5
    assert parse_font_size_input("１０．５") == 10.5


def test_parse_named_font_size_with_hint_suffix():
    assert parse_font_size_input("小四(12pt)") == 12.0


def test_font_size_display_keeps_numeric_value():
    assert font_size_display_text(12) == "12"
    assert font_size_display_text(10.5) == "10.5"
    assert font_size_display_text(13) == "13"


def test_numeric_font_size_options_use_word_like_steps():
    assert NUMERIC_FONT_SIZE_OPTIONS == (
        5.0, 5.5, 6.5, 7.5, 8.0, 9.0, 10.0, 10.5, 11.0, 12.0,
        14.0, 16.0, 18.0, 20.0, 22.0, 26.0, 28.0, 36.0, 48.0, 56.0, 72.0,
    )


def test_named_font_size_input_detection():
    assert is_named_font_size_input("小四")
    assert is_named_font_size_input("小四(12pt)")
    assert not is_named_font_size_input("12")


def test_format_font_size_pt_trims_trailing_zero():
    assert format_font_size_pt(12.0) == "12"
    assert format_font_size_pt(10.5) == "10.5"


def test_style_font_size_preserves_named_display_after_save_roundtrip():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    dlg = FormatConfigDialog(cfg)

    size_col = dlg._style_table_col_index("size_pt")
    combo = dlg._style_cell_widget(0, size_col)
    combo.setCurrentText("初号")
    dlg._sync_config_from_ui()

    assert cfg.styles["normal"].size_pt == 42.0
    assert cfg.styles["normal"].size_display == "初号"

    dlg2 = FormatConfigDialog(cfg)
    combo2 = dlg2._style_cell_widget(0, size_col)
    assert combo2.currentText() == "初号"
