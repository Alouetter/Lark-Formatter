import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest
from PySide6.QtCore import Qt, QRect
from PySide6.QtGui import QKeyEvent
from PySide6.QtWidgets import QApplication, QDialog, QMessageBox, QGroupBox, QLabel

from src.scene.manager import PRESETS_DIR, delete_scene, load_scene, load_scene_from_data, rename_scene
from src.scene.schema import default_chain_id_for_level
from src.ui.main_window import FormatConfigDialog, MainWindow


def _app():
    return QApplication.instance() or QApplication([])


def _style_size_combo(dialog: FormatConfigDialog):
    row = dialog._style_keys.index("normal")
    col = dialog._style_table_col_index("size_pt")
    return dialog._style_cell_widget(row, col)


def test_cancelled_incremental_save_does_not_mutate_live_config(tmp_path):
    _app()
    cfg = load_scene_from_data({})
    scene_path = tmp_path / "demo_scene.json"
    dlg = FormatConfigDialog(cfg, scene_path=str(scene_path))

    original_name = cfg.name
    original_report_json = cfg.output.report_json

    dlg._output_checks["report_json"].setChecked(False)
    dlg._prompt_format_name = lambda *args, **kwargs: ("", False)
    dlg._save_format()

    assert cfg.name == original_name
    assert cfg.output.report_json is original_report_json
    assert dlg.result() == 0


def test_apply_with_invalid_input_keeps_dialog_open_and_config_unmodified(monkeypatch):
    _app()
    cfg = load_scene_from_data({})
    dlg = FormatConfigDialog(cfg)

    warnings = []
    monkeypatch.setattr(
        QMessageBox,
        "warning",
        staticmethod(lambda *args, **kwargs: warnings.append(args[2]) or QMessageBox.Ok),
    )

    combo = _style_size_combo(dlg)
    combo.setCurrentText("abc")
    dlg._apply_changes()

    assert warnings
    assert dlg.result() == 0
    assert cfg.styles["normal"].size_pt == 12.0
    assert cfg.styles["normal"].size_display == ""


def test_incremental_save_with_invalid_input_does_not_prompt_or_close(monkeypatch, tmp_path):
    _app()
    cfg = load_scene_from_data({})
    scene_path = tmp_path / "demo_scene.json"
    dlg = FormatConfigDialog(cfg, scene_path=str(scene_path))

    warnings = []
    prompt_called = {"value": False}
    monkeypatch.setattr(
        QMessageBox,
        "warning",
        staticmethod(lambda *args, **kwargs: warnings.append(args[2]) or QMessageBox.Ok),
    )

    combo = _style_size_combo(dlg)
    combo.setCurrentText("abc")

    def _prompt(*args, **kwargs):
        prompt_called["value"] = True
        return "ShouldNotSave", True

    dlg._prompt_format_name = _prompt
    dlg._save_format()

    assert warnings
    assert prompt_called["value"] is False
    assert dlg.result() == 0
    assert cfg.styles["normal"].size_pt == 12.0
    assert cfg.styles["normal"].size_display == ""


def test_incremental_save_failure_does_not_mutate_live_config(monkeypatch, tmp_path):
    _app()
    cfg = load_scene_from_data({})
    scene_path = tmp_path / "demo_scene.json"
    dlg = FormatConfigDialog(cfg, scene_path=str(scene_path))

    criticals = []
    monkeypatch.setattr(
        QMessageBox,
        "critical",
        staticmethod(lambda *args, **kwargs: criticals.append(args[2]) or QMessageBox.Ok),
    )
    monkeypatch.setattr("src.scene.manager.save_scene", lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("boom")))

    original_name = cfg.name
    original_report_json = cfg.output.report_json

    dlg._output_checks["report_json"].setChecked(False)
    dlg._prompt_format_name = lambda *args, **kwargs: ("New Format", True)
    dlg._save_format()

    assert criticals
    assert cfg.name == original_name
    assert cfg.output.report_json is original_report_json
    assert dlg.result() == 0


def test_incremental_save_resets_format_signature_for_new_file(tmp_path):
    _app()
    cfg = load_scene_from_data(
        {
            "name": "\u539f\u683c\u5f0f",
            "format_signature": "\u65e7\u7f72\u540d",
        }
    )
    scene_path = tmp_path / "demo_scene.json"
    dlg = FormatConfigDialog(cfg, scene_path=str(scene_path))

    dlg._prompt_format_name = lambda *args, **kwargs: ("\u65b0\u683c\u5f0f", True)
    dlg._save_format()

    saved_path = tmp_path / "\u65b0\u683c\u5f0f.json"
    reloaded = load_scene(saved_path)

    assert dlg.result() == QDialog.Accepted
    assert cfg.name == "\u65b0\u683c\u5f0f"
    assert cfg.format_signature == ""
    assert saved_path.exists()
    assert reloaded.format_signature == ""


def test_main_window_rejects_overwriting_permanent_format_signature(monkeypatch):
    _app()
    monkeypatch.setattr(MainWindow, "_load_ui_state", lambda self: {})
    monkeypatch.setattr(MainWindow, "_save_ui_state", lambda self: None)
    monkeypatch.setattr(MainWindow, "_persist_current_scene_config", lambda self: None)

    window = MainWindow()
    try:
        window._config = load_scene_from_data(
            {
                "name": "\u9ed8\u8ba4\u683c\u5f0f",
                "format_signature": "\u5df2\u7f72\u540d",
            }
        )
        window._current_scene_path = str(PRESETS_DIR / "custom_scene.json")

        with pytest.raises(RuntimeError):
            window._update_current_format_signature("\u65b0\u7f72\u540d")
    finally:
        window.close()


def test_main_window_disables_signature_for_protected_default_scene(monkeypatch, tmp_path):
    _app()
    monkeypatch.setattr(MainWindow, "_load_ui_state", lambda self: {})
    monkeypatch.setattr(MainWindow, "_save_ui_state", lambda self: None)

    window = MainWindow()
    try:
        protected_path = tmp_path / "default_format.json"
        protected_path.write_text("{}", encoding="utf-8")
        window._config = load_scene_from_data({"name": "\u9ed8\u8ba4\u683c\u5f0f"})
        window._current_scene_path = str(protected_path)

        monkeypatch.setattr("src.ui.main_window.is_protected_scene_path", lambda path: True)

        assert window._current_format_signable() is False
        assert window._current_format_signature_block_label() == "\u9ed8\u8ba4\u683c\u5f0f\u4e0d\u53ef\u7f72\u540d"
        assert "\u9ed8\u8ba4\u683c\u5f0f" in window._current_format_signature_block_reason()

        with pytest.raises(RuntimeError):
            window._update_current_format_signature("\u65b0\u7f72\u540d")
    finally:
        window.close()


def test_main_window_disables_signature_without_scene_file(monkeypatch):
    _app()
    monkeypatch.setattr(MainWindow, "_load_ui_state", lambda self: {})
    monkeypatch.setattr(MainWindow, "_save_ui_state", lambda self: None)

    window = MainWindow()
    try:
        window._config = load_scene_from_data({"name": "\u4e34\u65f6\u683c\u5f0f"})
        window._current_scene_path = ""

        assert window._current_format_signable() is False
        assert window._current_format_signature_block_label() == "\u672a\u4fdd\u5b58\u683c\u5f0f\u4e0d\u53ef\u7f72\u540d"
        assert "\u672a\u5173\u8054\u5230\u914d\u7f6e\u6587\u4ef6" in window._current_format_signature_block_reason()

        with pytest.raises(RuntimeError):
            window._update_current_format_signature("\u65b0\u7f72\u540d")
    finally:
        window.close()


def test_numbering_mode_restores_matching_preset_on_open():
    _app()
    cfg = load_scene_from_data(
        {
            "heading_numbering_v2": {
                "level_bindings": {
                    "heading1": {
                        "enabled": True,
                        "display_shell": "chapter_cn",
                        "display_core_style": "cn_lower",
                        "reference_core_style": "cn_lower",
                        "chain": "current_only",
                    },
                    "heading2": {
                        "enabled": True,
                        "display_shell": "section_cn",
                        "display_core_style": "cn_lower",
                        "reference_core_style": "cn_lower",
                        "chain": "current_only",
                    },
                    "heading3": {
                        "enabled": True,
                        "display_shell": "dunhao_cn",
                        "display_core_style": "cn_lower",
                        "reference_core_style": "cn_lower",
                        "chain": "current_only",
                    },
                    "heading4": {
                        "enabled": True,
                        "display_shell": "paren_cn",
                        "display_core_style": "cn_lower",
                        "reference_core_style": "cn_lower",
                        "chain": "current_only",
                    },
                    "heading5": {
                        "enabled": False,
                        "chain": default_chain_id_for_level("heading5"),
                    },
                    "heading6": {
                        "enabled": False,
                        "chain": default_chain_id_for_level("heading6"),
                    },
                    "heading7": {
                        "enabled": False,
                        "chain": default_chain_id_for_level("heading7"),
                    },
                    "heading8": {
                        "enabled": False,
                        "chain": default_chain_id_for_level("heading8"),
                    },
                }
            }
        }
    )
    dlg = FormatConfigDialog(cfg)

    assert dlg._num_mode_combo.currentData() == "preset_1"
    assert dlg._num_detail_container.isEnabled() is False
    effect = dlg._num_detail_container.graphicsEffect()
    assert effect is not None
    assert round(effect.opacity(), 2) == 0.55


def test_numbering_mode_custom_reenables_detail_area_without_dimming():
    app = _app()
    cfg = load_scene_from_data({})
    dlg = FormatConfigDialog(cfg)

    dlg._set_combo_current_data(dlg._num_mode_combo, "preset_1")
    app.processEvents()
    assert dlg._num_detail_container.isEnabled() is False
    assert dlg._num_detail_container.graphicsEffect() is not None

    dlg._set_combo_current_data(dlg._num_mode_combo, "custom")
    app.processEvents()

    assert dlg._num_detail_container.isEnabled() is True
    assert dlg._num_detail_container.graphicsEffect() is None


def test_formula_table_controls_roundtrip_into_draft_config():
    _app()
    cfg = load_scene_from_data({})
    dlg = FormatConfigDialog(cfg)

    dlg._formula_table_unify_font_check.setChecked(False)
    dlg._formula_table_unify_size_check.setChecked(False)
    dlg._formula_table_unify_spacing_check.setChecked(True)
    dlg._formula_table_font_combo.setCurrentText("XITS Math")
    dlg._formula_table_font_size_combo.setCurrentText("11.5pt")
    dlg._formula_table_line_spacing_spin.setValue(1.25)
    dlg._formula_table_space_before_spin.setValue(2.0)
    dlg._formula_table_space_after_spin.setValue(3.0)
    dlg._set_combo_current_data(dlg._formula_table_block_alignment_combo, "right")
    dlg._set_combo_current_data(dlg._formula_table_table_alignment_combo, "left")
    dlg._set_combo_current_data(dlg._formula_table_cell_alignment_combo, "left")
    dlg._set_combo_current_data(dlg._formula_table_number_alignment_combo, "center")
    dlg._formula_table_number_font_combo.setCurrentText("Arial")
    dlg._formula_table_number_size_combo.setCurrentText("小五")
    dlg._formula_table_auto_shrink_check.setChecked(False)

    draft = dlg._build_config_draft_from_ui()

    assert draft is not None
    assert draft.formula_style.unify_font is False
    assert draft.formula_style.unify_size is False
    assert draft.formula_style.unify_spacing is True
    assert draft.formula_table.formula_font_name == "XITS Math"
    assert draft.formula_table.formula_font_size_pt == 11.5
    assert draft.formula_table.formula_font_size_display == "11.5pt"
    assert draft.formula_table.formula_line_spacing == 1.25
    assert draft.formula_table.formula_space_before_pt == 2.0
    assert draft.formula_table.formula_space_after_pt == 3.0
    assert draft.formula_table.block_alignment == "right"
    assert draft.formula_table.table_alignment == "left"
    assert draft.formula_table.formula_cell_alignment == "left"
    assert draft.formula_table.number_alignment == "center"
    assert draft.formula_table.number_font_name == "Arial"
    assert draft.formula_table.number_font_size_pt == 9.0
    assert draft.formula_table.number_font_size_display == "小五"
    assert draft.formula_table.auto_shrink_number_column is False


def test_caption_panel_equation_numbering_format_roundtrip_into_draft_config():
    _app()
    cfg = load_scene_from_data(
        {
            "caption": {"numbering_format": "chapter-seq"},
            "equation_table_format": {"enabled": True, "numbering_format": "seq"},
        }
    )
    dlg = FormatConfigDialog(cfg)

    assert dlg._numbering_format_combo.currentData() == "chapter-seq"
    assert dlg._equation_numbering_format_combo.currentData() == "seq"

    dlg._set_combo_current_data(dlg._numbering_format_combo, "chapter.seq")
    dlg._set_combo_current_data(dlg._equation_numbering_format_combo, "chapter-seq")

    draft = dlg._build_config_draft_from_ui()

    assert draft is not None
    assert draft.caption.numbering_format == "chapter.seq"
    assert draft.equation_table_format.numbering_format == "chapter-seq"


def test_formula_table_section_uses_divider_title_instead_of_group_box():
    _app()
    cfg = load_scene_from_data({})
    dlg = FormatConfigDialog(cfg)

    group_titles = {box.title() for box in dlg.findChildren(QGroupBox)}
    section_labels = [label.text() for label in dlg.findChildren(QLabel)]

    assert "公式表格配置" not in group_titles
    assert "公式表格配置" in section_labels


def test_formula_table_invalid_font_size_input_keeps_dialog_open(monkeypatch):
    _app()
    cfg = load_scene_from_data({})
    dlg = FormatConfigDialog(cfg)

    warnings = []
    monkeypatch.setattr(
        QMessageBox,
        "warning",
        staticmethod(lambda *args, **kwargs: warnings.append(args[2]) or QMessageBox.Ok),
    )

    dlg._formula_table_font_size_combo.setCurrentText("abc")
    dlg._apply_changes()

    assert warnings
    assert dlg.result() == 0
    assert cfg.formula_table.formula_font_size_pt == 12.0
    assert cfg.formula_table.formula_font_size_display == ""


def test_delete_scene_blocks_default_scene(monkeypatch, tmp_path):
    default_path = tmp_path / "default_format.json"
    default_path.write_text("{}", encoding="utf-8")

    monkeypatch.setattr("src.scene.manager._default_scene_candidate_paths", lambda: [default_path])

    with pytest.raises(PermissionError):
        delete_scene(default_path)

    assert default_path.exists()


def test_delete_button_disabled_for_default_scene(monkeypatch, tmp_path):
    _app()
    cfg = load_scene_from_data({})
    default_path = tmp_path / "default_format.json"
    default_path.write_text("{}", encoding="utf-8")

    monkeypatch.setattr("src.ui.main_window.is_protected_scene_path", lambda path: True)

    dlg = FormatConfigDialog(cfg, scene_path=str(default_path))

    assert dlg._delete_btn.isEnabled() is False


def test_rename_scene_blocks_default_scene(monkeypatch, tmp_path):
    default_path = tmp_path / "default_format.json"
    default_path.write_text("{}", encoding="utf-8")

    monkeypatch.setattr("src.scene.manager._default_scene_candidate_paths", lambda: [default_path])

    with pytest.raises(PermissionError):
        rename_scene(default_path, "新默认格式")

    assert default_path.exists()


def test_rename_button_disabled_for_default_scene(monkeypatch, tmp_path):
    _app()
    cfg = load_scene_from_data({})
    default_path = tmp_path / "default_format.json"
    default_path.write_text("{}", encoding="utf-8")

    monkeypatch.setattr("src.ui.main_window.is_protected_scene_path", lambda path: True)

    dlg = FormatConfigDialog(cfg, scene_path=str(default_path))

    assert dlg._rename_btn.isEnabled() is False


def test_main_window_autosave_writes_back_default_format_scene(monkeypatch, tmp_path):
    _app()
    monkeypatch.setattr(MainWindow, "_load_ui_state", lambda self: {})
    monkeypatch.setattr(MainWindow, "_save_ui_state", lambda self: None)
    monkeypatch.setattr("src.ui.main_window.is_protected_scene_path", lambda path: True)

    saved = []
    monkeypatch.setattr("src.ui.main_window.save_scene", lambda *args, **kwargs: saved.append(args))

    window = MainWindow()
    try:
        protected_path = tmp_path / "default_format.json"
        protected_path.write_text("{}", encoding="utf-8")
        window._config = load_scene_from_data({"name": "默认格式"})
        window._current_scene_path = str(protected_path)
        window._current_scene_is_custom = False

        window._save_current_scene_config()

        assert len(saved) == 1
        assert saved[0][1] == PRESETS_DIR / "default_format.json"
    finally:
        window.close()


def test_reference_hanging_indent_controls_roundtrip_into_draft_config():
    _app()
    cfg = load_scene_from_data({})
    dlg = FormatConfigDialog(cfg)

    row = dlg._style_keys.index("references_body")

    left_widget = dlg._style_cell_widget(row, dlg._style_table_col_index("left_indent_chars"))
    left_widget.setUnit("chars")
    left_widget._edit.setText("2")

    special_widget = dlg._style_cell_widget(row, dlg._style_table_col_index("special_indent_value"))
    special_widget.setMode("hanging")
    special_widget.setUnit("chars")
    special_widget._value_widget._edit.setText("2")

    draft = dlg._build_config_draft_from_ui()

    assert draft is not None
    assert draft.styles["references_body"].left_indent_chars == 2.0
    assert draft.styles["references_body"].left_indent_unit == "chars"
    assert draft.styles["references_body"].special_indent_mode == "hanging"
    assert draft.styles["references_body"].special_indent_value == 2.0
    assert draft.styles["references_body"].special_indent_unit == "chars"
    assert draft.styles["references_body"].hanging_indent_chars == 2.0
    assert draft.styles["references_body"].hanging_indent_unit == "chars"


def test_misc_format_controls_roundtrip_into_draft_config():
    _app()
    cfg = load_scene_from_data({})
    dlg = FormatConfigDialog(cfg)

    dlg._enforcement_checks["ban_tab"].setChecked(False)
    dlg._enforcement_checks["auto_fix"].setChecked(False)
    dlg._risk_guard_controls["enabled"].setChecked(False)
    dlg._risk_guard_controls["no_body_min_candidates"].setValue(5)
    dlg._risk_guard_controls["tiny_body_max_ratio"].setValue(0.125)
    dlg._risk_guard_controls["keep_after_first_chapter"].setChecked(False)

    dlg._cap_edits["figure_prefix"].setText("Fig")
    dlg._cap_edits["table_prefix"].setText("Tbl")
    dlg._cap_edits["placeholder"].setText("[PENDING]")
    dlg._set_combo_current_data(dlg._cap_separator_mode_combo, "custom")
    dlg._cap_custom_separator_edit.setText("::")
    dlg._set_combo_current_data(dlg._numbering_format_combo, "chapter-seq")
    dlg._set_combo_current_data(dlg._equation_numbering_format_combo, "seq")
    dlg._cap_enabled.setChecked(False)

    dlg._set_combo_current_data(dlg._paper_size_combo, "Letter")
    dlg._page_spins["top_cm"].setValue(2.6)
    dlg._page_spins["bottom_cm"].setValue(2.4)
    dlg._page_spins["left_cm"].setValue(2.8)
    dlg._page_spins["right_cm"].setValue(2.2)
    dlg._page_spins["gutter_cm"].setValue(0.4)
    dlg._page_spins["header_distance_cm"].setValue(1.5)
    dlg._page_spins["footer_distance_cm"].setValue(1.2)

    dlg._output_checks["compare_docx"].setChecked(False)
    dlg._output_checks["report_json"].setChecked(False)
    dlg._output_checks["report_markdown"].setChecked(False)

    dlg._pipeline_checks["caption_format"].setChecked(False)
    dlg._pipeline_checks["validation"].setChecked(False)

    draft = dlg._build_config_draft_from_ui()

    assert draft is not None
    assert draft.heading_numbering.enforcement.ban_tab is False
    assert draft.heading_numbering.enforcement.auto_fix is False
    assert draft.heading_numbering.risk_guard.enabled is False
    assert draft.heading_numbering.risk_guard.no_body_min_candidates == 5
    assert draft.heading_numbering.risk_guard.tiny_body_max_ratio == 0.125
    assert draft.heading_numbering.risk_guard.keep_after_first_chapter is False

    assert draft.caption.enabled is False
    assert draft.caption.figure_prefix == "Fig"
    assert draft.caption.table_prefix == "Tbl"
    assert draft.caption.placeholder == "[PENDING]"
    assert draft.caption.separator == "::"
    assert draft.caption.numbering_format == "chapter-seq"
    assert draft.equation_table_format.numbering_format == "seq"

    assert draft.page_setup.paper_size == "Letter"
    assert draft.page_setup.margin.top_cm == 2.6
    assert draft.page_setup.margin.bottom_cm == 2.4
    assert draft.page_setup.margin.left_cm == 2.8
    assert draft.page_setup.margin.right_cm == 2.2
    assert draft.page_setup.gutter_cm == 0.4
    assert draft.page_setup.header_distance_cm == 1.5
    assert draft.page_setup.footer_distance_cm == 1.2

    assert draft.output.compare_docx is False
    assert draft.output.report_json is False
    assert draft.output.report_markdown is False
    assert "caption_format" not in draft.pipeline
    assert "validation" not in draft.pipeline
    assert draft.pipeline_critical_rules == draft.pipeline


def test_special_indent_mode_switches_cleanly_in_ui():
    _app()
    cfg = load_scene_from_data({})
    dlg = FormatConfigDialog(cfg)

    row = dlg._style_keys.index("references_body")
    special_widget = dlg._style_cell_widget(row, dlg._style_table_col_index("special_indent_value"))

    special_widget.setMode("hanging")
    special_widget.setUnit("chars")
    special_widget._value_widget._edit.setText("2")

    draft = dlg._build_config_draft_from_ui()
    assert draft.styles["references_body"].special_indent_mode == "hanging"
    assert draft.styles["references_body"].hanging_indent_chars == 2.0
    assert draft.styles["references_body"].first_line_indent_chars == 0.0

    special_widget.setMode("first_line")
    draft = dlg._build_config_draft_from_ui()
    assert draft.styles["references_body"].special_indent_mode == "first_line"
    assert draft.styles["references_body"].first_line_indent_chars == 2.0
    assert draft.styles["references_body"].hanging_indent_chars == 0.0

    special_widget.setMode("none")
    draft = dlg._build_config_draft_from_ui()
    assert draft.styles["references_body"].special_indent_mode == "none"
    assert draft.styles["references_body"].special_indent_value == 0.0
    assert draft.styles["references_body"].first_line_indent_chars == 0.0
    assert draft.styles["references_body"].hanging_indent_chars == 0.0


def test_style_table_indent_headers_are_compact_and_empty_indent_cells_show_blank():
    _app()
    cfg = load_scene_from_data({})
    dlg = FormatConfigDialog(cfg)

    left_header = dlg._style_table.horizontalHeaderItem(dlg._style_table_col_index("left_indent_chars"))
    line_header = dlg._style_table.horizontalHeaderItem(dlg._style_table_col_index("line_spacing_pt"))
    assert left_header.text() == "左缩进"
    assert line_header.text() == "行距"

    row = dlg._style_keys.index("heading1")
    left_widget = dlg._style_cell_widget(row, dlg._style_table_col_index("left_indent_chars"))
    special_widget = dlg._style_cell_widget(row, dlg._style_table_col_index("special_indent_value"))

    assert left_widget.text() == ""
    assert special_widget.text() == ""
    assert "#f5f6f7" in left_widget._edit.styleSheet()


def test_indent_unit_toggle_keeps_numeric_text_without_auto_conversion():
    _app()
    cfg = load_scene_from_data({})
    dlg = FormatConfigDialog(cfg)

    row = dlg._style_keys.index("normal")
    left_widget = dlg._style_cell_widget(row, dlg._style_table_col_index("left_indent_chars"))
    right_widget = dlg._style_cell_widget(row, dlg._style_table_col_index("right_indent_chars"))

    left_widget.setUnit("chars")
    left_widget._edit.setText("2")
    left_widget._toggle_unit()
    assert left_widget.text() == "2"
    assert left_widget.unit() == "cm"

    right_widget.setUnit("pt")
    right_widget._edit.setText("18")
    right_widget._toggle_unit()
    assert right_widget.text() == "18"
    assert right_widget.unit() == "chars"


def test_enter_key_does_not_apply_or_close_dialog():
    app = _app()
    cfg = load_scene_from_data({})
    dlg = FormatConfigDialog(cfg)
    dlg.show()
    app.processEvents()

    row = dlg._style_keys.index("normal")
    size_combo = dlg._style_cell_widget(row, dlg._style_table_col_index("size_pt"))
    size_combo.setFocus()
    app.processEvents()

    event = QKeyEvent(QKeyEvent.KeyPress, Qt.Key_Return, Qt.NoModifier)
    QApplication.sendEvent(dlg, event)
    app.processEvents()

    assert dlg.result() == 0


def test_format_config_dialog_initial_width_expands_to_fit_style_table(monkeypatch):
    _app()

    class _FakeScreen:
        @staticmethod
        def availableGeometry():
            return QRect(0, 0, 1800, 1200)

    monkeypatch.setattr(FormatConfigDialog, "screen", lambda self: _FakeScreen())

    cfg = load_scene_from_data({})
    dlg = FormatConfigDialog(cfg)

    expected_w, _ = dlg._compute_initial_dialog_size()
    assert dlg.width() == expected_w
    assert dlg.width() > 960
