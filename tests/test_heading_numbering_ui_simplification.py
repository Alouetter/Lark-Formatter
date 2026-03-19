import copy
import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QAbstractSpinBox

from src.scene.manager import load_scene_from_data
from src.scene.schema import HeadingLevelConfig, StyleConfig, default_chain_id_for_level
from src.ui.main_window import FormatConfigDialog
from src.ui.theme_manager import ThemeManager
from src.utils.ooxml import build_abstract_num

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _combo_select_by_data(combo, value) -> None:
    idx = combo.findData(value)
    assert idx >= 0
    combo.setCurrentIndex(idx)


def test_numbering_tab_prefers_scheme_2_and_uses_structured_columns():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg.heading_numbering.scheme = "1"
    cfg.heading_numbering.levels = {
        "heading1": HeadingLevelConfig(format="arabic", template="{n}", separator=" "),
    }
    cfg.heading_numbering.schemes = {
        "1": {
            "heading1": HeadingLevelConfig(format="arabic", template="{n}", separator=" "),
        },
        "2": {
            "heading1": HeadingLevelConfig(format="chinese_chapter", template="第{cn}章", separator="\u3000"),
        },
    }

    dlg = FormatConfigDialog(cfg)

    attrs = [attr for attr, _, _ in dlg._NUM_COLS]
    assert attrs[0] == "preview"
    assert "display_shell" in attrs
    assert "chain_summary" in attrs
    assert "display_core_style" in attrs
    assert "reference_core_style" in attrs
    assert "start_at" in attrs
    assert "restart_on" in attrs
    assert "include_in_toc" in attrs
    assert "format" not in attrs
    assert "template" not in attrs
    assert dlg._numbering_scheme_id == "2"

    preview_col = dlg._num_col_index("preview")
    assert dlg._num_table.columnWidth(preview_col) == dlg._NUM_PREVIEW_BASE_WIDTH
    assert "第一章" in dlg._num_table.item(0, preview_col).text()
    assert dlg._num_table.item(0, preview_col).text().endswith("示例")
    assert dlg._num_table.item(0, preview_col).textAlignment() == int(Qt.AlignCenter)

    shell_combo = dlg._num_cell_widget(0, "display_shell")
    core_style_combo = dlg._num_cell_widget(0, "display_core_style")
    assert shell_combo.currentData() == "chapter_cn"
    assert core_style_combo.currentData() == "cn_lower"

    heading2_core_style_combo = dlg._num_cell_widget(1, "display_core_style")
    assert heading2_core_style_combo.currentText() == "1"
    assert heading2_core_style_combo.itemText(heading2_core_style_combo.findData("arabic")) == "1 2 3"

    del app


def test_numbering_tab_hides_removed_core_style_options():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    core_style_combo = dlg._num_cell_widget(0, "display_core_style")
    option_ids = {core_style_combo.itemData(index) for index in range(core_style_combo.count())}

    assert "arabic_fullwidth" not in option_ids
    assert "arabic_pad3" not in option_ids
    assert "roman_lower" not in option_ids
    assert "alpha_upper" not in option_ids
    assert "alpha_lower" not in option_ids

    del app


def test_numbering_tab_applies_top_level_preset_scheme():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    _combo_select_by_data(dlg._num_mode_combo, "preset_2")

    preview_col = dlg._num_col_index("preview")
    assert dlg._num_table.item(0, preview_col).text().startswith("1")
    assert dlg._num_table.item(1, preview_col).text().startswith("1.1")
    assert dlg._num_table.item(2, preview_col).text().startswith("1.1.1")
    assert "一级起" in dlg._num_cell_widget(2, "chain_summary").text()
    assert dlg._num_cell_widget(0, "display_shell").isEnabled() is False
    assert dlg._num_detail_container.isEnabled() is False

    draft = copy.deepcopy(cfg)
    assert dlg._sync_config_from_ui(target_config=draft) is True
    assert draft.heading_numbering_v2.level_bindings["heading1"].display_shell == "plain"
    assert draft.heading_numbering_v2.level_bindings["heading1"].reference_core_style == "arabic"
    assert draft.heading_numbering_v2.level_bindings["heading3"].chain == default_chain_id_for_level("heading3")

    del app


def test_numbering_tab_supports_trailing_dot_preset():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    _combo_select_by_data(dlg._num_mode_combo, "preset_3")

    preview_col = dlg._num_col_index("preview")
    assert dlg._num_table.item(0, preview_col).text().startswith("1.")
    assert dlg._num_table.item(1, preview_col).text().startswith("1.1.")
    assert dlg._num_table.item(2, preview_col).text().startswith("1.1.1.")

    draft = copy.deepcopy(cfg)
    assert dlg._sync_config_from_ui(target_config=draft) is True
    assert draft.heading_numbering_v2.level_bindings["heading1"].display_shell == "dot_suffix"

    del app


def test_numbering_compact_combo_display_survives_preset_switches():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    for mode in ("preset_1", "preset_2", "preset_3", "custom"):
        _combo_select_by_data(dlg._num_mode_combo, mode)
        for row in range(8):
            core_combo = dlg._num_cell_widget(row, "display_core_style")
            ref_combo = dlg._num_cell_widget(row, "reference_core_style")
            if core_combo.lineEdit() is not None:
                assert len(core_combo.lineEdit().text()) <= 2
            if ref_combo.lineEdit() is not None:
                assert len(ref_combo.lineEdit().text()) <= 2

    del app


def test_numbering_custom_state_restores_after_switching_presets():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    _combo_select_by_data(dlg._num_mode_combo, "custom")
    dlg._num_cell_widget(0, "display_shell").setCurrentText("【{}】")

    _combo_select_by_data(dlg._num_mode_combo, "preset_2")
    assert dlg._num_cell_widget(0, "display_shell").currentData() == "plain"

    _combo_select_by_data(dlg._num_mode_combo, "custom")
    assert dlg._num_cell_widget(0, "display_shell").currentText() == "【{}】"

    del app


def test_numbering_tab_edit_round_trips_to_v2_and_legacy_config():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"

    dlg = FormatConfigDialog(cfg)
    _combo_select_by_data(dlg._num_cell_widget(0, "display_shell"), "plain")
    _combo_select_by_data(dlg._num_cell_widget(0, "display_core_style"), "arabic")
    _combo_select_by_data(
        dlg._num_cell_widget(0, "reference_core_style"),
        dlg._NUM_REFERENCE_SAME_AS_DISPLAY,
    )
    dlg._num_cell_widget(0, "title_separator").setRawText(" ")

    draft = copy.deepcopy(cfg)
    assert dlg._sync_config_from_ui(target_config=draft) is True

    binding = draft.heading_numbering_v2.level_bindings["heading1"]
    assert binding.display_shell == "plain"
    assert binding.display_core_style == "arabic"
    assert binding.reference_core_style == "arabic"
    assert binding.chain == "current_only"
    assert binding.title_separator == " "

    assert draft.heading_numbering.levels["heading1"].format == "arabic"
    assert draft.heading_numbering.levels["heading1"].template == "{current}"
    assert draft.heading_numbering.levels["heading1"].separator == " "

    del app


def test_numbering_title_separator_supports_preset_and_custom():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    separator_widget = dlg._num_cell_widget(0, "title_separator")
    _combo_select_by_data(separator_widget._mode_combo, "halfwidth_space")
    assert separator_widget.text() == " "
    assert separator_widget._mode_combo.currentText() == "·"
    assert separator_widget._mode_combo.itemText(
        separator_widget._mode_combo.findData("fullwidth_space")
    ) == "全角空格(□ / U+3000)"
    assert separator_widget._mode_combo.itemText(
        separator_widget._mode_combo.findData("tab")
    ) == "制表符(➡ / U+0009)"

    _combo_select_by_data(separator_widget._mode_combo, "tab")
    assert separator_widget.text() == "\t"
    assert separator_widget._mode_combo.currentText() == "➡"
    draft_tab = copy.deepcopy(cfg)
    assert dlg._sync_config_from_ui(target_config=draft_tab) is True
    tab_binding = draft_tab.heading_numbering_v2.level_bindings["heading1"]
    assert tab_binding.title_separator == "\t"
    assert tab_binding.ooxml_separator_mode == "suff"
    assert tab_binding.ooxml_suff == "tab"

    _combo_select_by_data(separator_widget._mode_combo, "custom")
    separator_widget.setRawText("__")
    assert separator_widget.text() == "__"

    draft = copy.deepcopy(cfg)
    assert dlg._sync_config_from_ui(target_config=draft) is True
    assert draft.heading_numbering_v2.level_bindings["heading1"].title_separator == "__"
    assert draft.heading_numbering.levels["heading1"].separator == "__"

    del app


def test_numbering_preview_renders_tab_separator_as_three_spaces():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    separator_widget = dlg._num_cell_widget(0, "title_separator")
    _combo_select_by_data(separator_widget._mode_combo, "tab")
    dlg._refresh_numbering_previews()

    preview_col = dlg._num_col_index("preview")
    preview_text = dlg._num_table.item(0, preview_col).text()
    assert "   示例" in preview_text
    assert "\t" not in preview_text

    del app


def test_numbering_title_separator_widget_hides_custom_editor_for_presets():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)
    dlg.show()
    QApplication.processEvents()
    QApplication.processEvents()

    separator_widget = dlg._num_cell_widget(0, "title_separator")
    separator_col = dlg._num_col_index("title_separator")
    separator_wrap = dlg._num_table.cellWidget(0, separator_col)
    assert separator_widget._edit.isHidden() is True
    assert separator_widget._mode_combo.height() == dlg._NUM_TABLE_CONTROL_HEIGHT
    assert dlg._num_table.columnWidth(separator_col) == dlg._NUM_TITLE_SEPARATOR_EXPANDED_WIDTH
    assert separator_widget.width() <= separator_wrap.width()
    assert separator_widget._mode_combo.width() == separator_widget.width()
    assert separator_widget._mode_combo.currentText() == "\u25A1"

    _combo_select_by_data(separator_widget._mode_combo, "custom")
    QApplication.processEvents()
    QApplication.processEvents()
    assert separator_widget._edit.isHidden() is False
    assert dlg._num_table.columnWidth(separator_col) == dlg._NUM_TITLE_SEPARATOR_EXPANDED_WIDTH
    assert separator_widget.width() <= separator_wrap.width()
    assert separator_widget._mode_combo.width() >= dlg._NUM_TITLE_SEPARATOR_COMPACT_CONTROL_WIDTH
    assert (
        separator_widget._mode_combo.width()
        + separator_widget._edit.width()
        + separator_widget._custom_layout_spacing
        <= separator_widget.width()
    )
    assert separator_widget._mode_combo.currentText() == "\u81ea\u5b9a\u4e49"

    _combo_select_by_data(separator_widget._mode_combo, "halfwidth_space")
    QApplication.processEvents()
    QApplication.processEvents()
    assert separator_widget._mode_combo.currentData() == "halfwidth_space"
    assert separator_widget._edit.isHidden() is True
    assert separator_widget._edit.displayText() == "\u00B7"
    assert separator_widget._mode_combo.currentText() == "\u00B7"
    assert dlg._num_table.columnWidth(separator_col) == dlg._NUM_TITLE_SEPARATOR_EXPANDED_WIDTH

    del app

def test_numbering_restart_column_uses_clearer_label():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    assert dlg._num_table.horizontalHeaderItem(dlg._num_col_index("restart_on")).text() == "遇级"

    del app


def test_numbering_start_spin_keeps_native_arrow_buttons():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)
    dlg.show()
    QApplication.processEvents()
    QApplication.processEvents()

    start_spin = dlg._num_cell_widget(0, "start_at")
    start_wrap = dlg._num_table.cellWidget(0, dlg._num_col_index("start_at"))
    assert start_spin.buttonSymbols() == QAbstractSpinBox.UpDownArrows
    assert start_spin.width() == dlg._NUM_START_AT_CONTROL_WIDTH
    assert start_spin.width() <= start_wrap.width()
    assert start_spin.styleSheet().strip() == ""

    del app

def test_numbering_inner_widgets_stay_within_cell_widths():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)
    dlg.show()
    QApplication.processEvents()
    QApplication.processEvents()

    for attr in ("reference_core_style", "title_separator", "start_at", "include_in_toc", "enabled"):
        col = dlg._num_col_index(attr)
        wrap = dlg._num_table.cellWidget(0, col)
        inner = dlg._num_cell_widget(0, attr)
        assert inner.width() <= wrap.width()

    separator_widget = dlg._num_cell_widget(0, "title_separator")
    _combo_select_by_data(separator_widget._mode_combo, "custom")
    QApplication.processEvents()
    QApplication.processEvents()
    assert separator_widget.width() <= dlg._num_table.cellWidget(
        0, dlg._num_col_index("title_separator")
    ).width()
    assert (
        separator_widget._mode_combo.width()
        + separator_widget._edit.width()
        + separator_widget._custom_layout_spacing
        <= separator_widget.width()
    )

    del app


def test_numbering_table_uses_compact_style_table_object_name():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    assert dlg._num_table.objectName() == "StyleConfigTable"

    del app


def test_numbering_table_controls_match_separator_height_under_theme():
    app = QApplication.instance() or QApplication([])
    ThemeManager.apply_theme(app, "light")
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)
    dlg.show()
    QApplication.processEvents()

    expected_height = dlg._num_cell_widget(0, "title_separator").height()
    assert expected_height == dlg._NUM_TABLE_CONTROL_HEIGHT

    for attr in (
        "display_shell",
        "chain_summary",
        "display_core_style",
        "reference_core_style",
        "start_at",
        "restart_on",
        "include_in_toc",
        "enabled",
    ):
        assert dlg._num_cell_widget(0, attr).height() == expected_height

    assert dlg._num_table.columnWidth(dlg._num_col_index("start_at")) >= 70

    del app


def test_numbering_compact_combo_display_recovers_after_reselecting_same_item():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    ref_combo = dlg._num_cell_widget(1, "reference_core_style")
    same_idx = ref_combo.findData(dlg._NUM_REFERENCE_SAME_AS_DISPLAY)
    assert same_idx >= 0
    ref_combo.lineEdit().setText("沿用本体")
    ref_combo.activated.emit(same_idx)
    QApplication.processEvents()
    assert ref_combo.lineEdit().text() == "沿用"

    separator_widget = dlg._num_cell_widget(0, "title_separator")
    separator_widget._mode_combo.lineEdit().setText("全角空格(□ / U+3000)")
    fullwidth_idx = separator_widget._mode_combo.findData("fullwidth_space")
    separator_widget._mode_combo.activated.emit(fullwidth_idx)
    QApplication.processEvents()
    assert separator_widget._mode_combo.currentText() == "□"

    del app


def test_numbering_preview_reflects_parent_reference_style():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    preview_col = dlg._num_col_index("preview")
    assert dlg._num_table.item(1, preview_col).text().startswith("1.1")
    assert "一级起" in dlg._num_cell_widget(1, "chain_summary").text()

    _combo_select_by_data(
        dlg._num_cell_widget(0, "reference_core_style"),
        dlg._NUM_REFERENCE_SAME_AS_DISPLAY,
    )
    preview_text = dlg._num_table.item(1, preview_col).text()
    assert preview_text.startswith("1.1")

    del app


def test_numbering_tab_round_trips_behavior_fields_and_preview_uses_start_at():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    dlg._num_cell_widget(1, "start_at").setValue(5)
    _combo_select_by_data(dlg._num_cell_widget(2, "restart_on"), "heading1")
    dlg._num_cell_widget(2, "include_in_toc").setChecked(False)
    dlg._refresh_numbering_previews()

    preview_col = dlg._num_col_index("preview")
    assert dlg._num_table.item(1, preview_col).text().startswith("1.5")

    draft = copy.deepcopy(cfg)
    assert dlg._sync_config_from_ui(target_config=draft) is True
    assert draft.heading_numbering_v2.level_bindings["heading2"].start_at == 5
    assert draft.heading_numbering_v2.level_bindings["heading3"].restart_on == "heading1"
    assert draft.heading_numbering_v2.level_bindings["heading3"].include_in_toc is False

    del app


def test_numbering_preview_shows_only_disabled_marker_when_disabled():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    dlg._num_cell_widget(0, "enabled").setChecked(False)
    preview_col = dlg._num_col_index("preview")
    assert dlg._num_table.item(0, preview_col).text() == "[停用]"

    del app


def test_numbering_preview_uses_example_text_for_enabled_higher_levels():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    dlg._num_cell_widget(4, "enabled").setChecked(True)
    preview_col = dlg._num_col_index("preview")
    preview_item = dlg._num_table.item(4, preview_col)
    assert preview_item.text().endswith("示例")

    del app


def test_numbering_tab_syncs_decimal_parent_chain_back_to_legacy_template():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    _combo_select_by_data(dlg._num_cell_widget(1, "display_shell"), "plain")
    _combo_select_by_data(dlg._num_cell_widget(1, "display_core_style"), "arabic")
    dlg._set_chain_button_state(
        dlg._num_cell_widget(1, "chain_summary"),
        "heading2",
        dlg._resolve_or_create_chain_id_from_text("heading2", "{level1}.{current}"),
    )

    draft = copy.deepcopy(cfg)
    assert dlg._sync_config_from_ui(target_config=draft) is True
    assert draft.heading_numbering.levels["heading2"].template == "{parent}.{current}"

    del app


def test_numbering_tab_keeps_non_decimal_parent_chain_explicit():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    _combo_select_by_data(dlg._num_cell_widget(1, "display_core_style"), "cn_lower")
    dlg._set_chain_button_state(
        dlg._num_cell_widget(1, "chain_summary"),
        "heading2",
        dlg._resolve_or_create_chain_id_from_text("heading2", "{level1}.{current}"),
    )

    draft = copy.deepcopy(cfg)
    assert dlg._sync_config_from_ui(target_config=draft) is True
    assert draft.heading_numbering.levels["heading2"].template == "{level1}.{current}"

    del app


def test_numbering_tab_syncs_decimal_parent_chain_for_heading3():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    _combo_select_by_data(dlg._num_cell_widget(2, "display_shell"), "plain")
    _combo_select_by_data(dlg._num_cell_widget(2, "display_core_style"), "arabic")
    dlg._set_chain_button_state(
        dlg._num_cell_widget(2, "chain_summary"),
        "heading3",
        dlg._resolve_or_create_chain_id_from_text("heading3", "{level1}.{level2}.{current}"),
    )

    draft = copy.deepcopy(cfg)
    assert dlg._sync_config_from_ui(target_config=draft) is True
    assert draft.heading_numbering.levels["heading3"].template == "{parent}.{current}"

    del app


def test_numbering_preview_uses_decimal_chain_for_enabled_heading5():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    dlg._num_cell_widget(4, "enabled").setChecked(True)
    _combo_select_by_data(dlg._num_cell_widget(4, "display_core_style"), "arabic")
    dlg._set_chain_button_state(
        dlg._num_cell_widget(4, "chain_summary"),
        "heading5",
        dlg._resolve_or_create_chain_id_from_text("heading5", "{level1}.{level2}.{level3}.{level4}.{current}"),
    )

    preview_col = dlg._num_col_index("preview")
    assert dlg._num_table.item(4, preview_col).text().startswith("1.1.1.1.1")

    del app


def test_numbering_tab_accepts_custom_shell_text():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    dlg._num_cell_widget(0, "display_shell").setCurrentText("【{}】")

    draft = copy.deepcopy(cfg)
    assert dlg._sync_config_from_ui(target_config=draft) is True
    binding = draft.heading_numbering_v2.level_bindings["heading1"]
    shell = draft.heading_numbering_v2.shell_catalog[binding.display_shell]

    assert binding.display_shell.startswith("custom_shell_heading1")
    assert shell.label == "【{}】"
    assert shell.prefix == "【"
    assert shell.suffix == "】"
    assert draft.heading_numbering.levels["heading1"].template == "【{current}】"

    del app


def test_numbering_preview_does_not_persist_custom_shell_before_save():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    dlg._num_cell_widget(0, "display_shell").setCurrentText("【{}】")
    dlg._refresh_numbering_previews()

    assert dlg._find_shell_id_by_text("【{}】") is None
    preview_col = dlg._num_col_index("preview")
    assert dlg._num_table.item(0, preview_col).text().startswith("【一】")

    draft = copy.deepcopy(cfg)
    assert dlg._sync_config_from_ui(target_config=draft) is True
    binding = draft.heading_numbering_v2.level_bindings["heading1"]
    assert draft.heading_numbering_v2.shell_catalog[binding.display_shell].label == "【{}】"

    del app


def test_numbering_tab_allocates_distinct_custom_catalog_ids():
    app = QApplication.instance() or QApplication([])
    cfg = load_scene_from_data({})
    cfg._heading_numbering_v2_source = "payload"
    dlg = FormatConfigDialog(cfg)

    first_shell_id = dlg._resolve_or_create_shell_id_from_text("heading1", "【{}】")
    second_shell_id = dlg._resolve_or_create_shell_id_from_text("heading1", "〔{}〕")
    assert first_shell_id != second_shell_id
    assert dlg._num_catalog_source.shell_catalog[first_shell_id].label == "【{}】"
    assert dlg._num_catalog_source.shell_catalog[second_shell_id].label == "〔{}〕"

    first_chain_id = dlg._resolve_or_create_chain_id_from_text("heading3", "{level1}-{current}")
    second_chain_id = dlg._resolve_or_create_chain_id_from_text("heading3", "{level2}/{current}")
    assert first_chain_id != second_chain_id
    assert dlg._num_catalog_source.chain_catalog[first_chain_id].label == "{level1}-{current}"
    assert dlg._num_catalog_source.chain_catalog[second_chain_id].label == "{level2}/{current}"

    del app


def test_heading_numbering_ooxml_uses_style_alignment_and_indent_as_single_source():
    levels_config = {
        "heading1": HeadingLevelConfig(
            format="arabic",
            template="{n}",
            separator="\u3000",
            alignment="left",
            left_indent_chars=0.0,
        )
    }
    styles_config = {
        "heading1": StyleConfig(
            alignment="right",
            left_indent_chars=2.0,
            space_before_pt=6.0,
            space_after_pt=8.0,
        )
    }

    abs_num = build_abstract_num(1, levels_config, styles_config)
    lvl = abs_num.find(_w("lvl"))
    assert lvl is not None

    ppr = lvl.find(_w("pPr"))
    assert ppr is not None
    jc = ppr.find(_w("jc"))
    assert jc is not None
    assert jc.get(_w("val")) == "right"

    ind = ppr.find(_w("ind"))
    assert ind is not None
    assert ind.get(_w("leftChars")) == "200"
    assert ind.get(_w("hanging")) == "0"
