"""主窗口：docx 一键排版桌面应用"""

import os
import json
import copy
import html
import re
import shutil
import sys
from dataclasses import asdict
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QPushButton, QFileDialog,
    QGroupBox, QTextEdit, QMessageBox, QRadioButton,
    QButtonGroup, QLineEdit, QCheckBox,
    QDialog, QTabWidget, QTableWidget, QTableWidgetItem,
    QHeaderView, QDialogButtonBox, QFormLayout, QFrame, QAbstractItemView, QCompleter,
    QStyleOptionHeader,
    QDoubleSpinBox, QSpinBox, QAbstractSpinBox, QApplication, QSizePolicy,
    QProgressBar, QInputDialog, QLayout, QScrollArea, QStyle, QGraphicsOpacityEffect
)
from PySide6.QtCore import Qt, QThread, Signal, QSortFilterProxyModel, QStringListModel, QEvent, QPoint, QSize
from PySide6.QtCore import QTimer

from src.scene.manager import (
    PRESETS_DIR,
    _derive_heading_numbering_v2_from_legacy,
    _sync_pipeline_critical_rules,
    get_scene_upgrade_notes,
    is_protected_scene_path,
    list_presets,
    load_default_scene,
    load_scene,
    load_scene_from_data,
    save_scene,
)
from src.scene.schema import SceneConfig, HeadingLevelBindingConfig, HeadingLevelConfig, HeadingNumberingV2Config
from src.scene.schema import (
    NumberShellConfig,
    NumberChainConfig,
    NumberChainSegmentConfig,
    default_chain_id_for_level,
    heading_level_index,
)
from src.engine.pipeline import Pipeline, PipelineResult
from src.engine.page_scope import parse_page_ranges_text, format_page_ranges_text
from src.report.collector import collect_report
from src.report.json_report import generate_json_report
from src.report.markdown_report import generate_markdown_report
from src.ui.progress_dialog import ProgressDialog
from src.converter.md_to_docx import convert_md_to_docx
from src.docx_io.style_clone import clone_scene_style_from_docx
from src.ui.alignment_options import (
    ALIGNMENT_OPTIONS,
    alignment_display_label,
    normalize_alignment_value,
)
from src.ui.line_spacing_options import (
    line_spacing_display_label,
    line_spacing_unit_label,
    normalize_line_spacing_type,
    resolve_line_spacing_value,
)
from src.ui.font_sizes import (
    NUMERIC_FONT_SIZE_OPTIONS,
    WORD_NAMED_FONT_SIZES,
    font_size_display_text,
    format_font_size_pt,
    is_named_font_size_input,
    parse_font_size_input,
)
from src.ui.font_search import font_matches_query
from src.ui.theme_manager import ThemeManager
from src.ui.wheel_guard import install_global_wheel_guard
from src.utils.heading_numbering_template import (
    validate_heading_numbering_template,
)
from src.utils.heading_numbering_v2 import (
    advance_heading_counters,
    explicit_chain_template_from_binding,
    format_core_style_value,
    legacy_format_from_binding,
    legacy_levels_from_v2,
    legacy_template_from_binding,
    merged_level_binding,
    normalize_start_at,
    uses_decimal_parent_chain,
)
from src.utils.indent import (
    normalize_indent_unit,
    normalize_special_indent_mode,
    resolve_pt_indent_value,
    sync_style_config_indent_fields,
)


class ScrollSafeComboBox(QComboBox):
    """QComboBox that ignores mouse wheel events to prevent accidental value changes."""
    def wheelEvent(self, event):
        event.ignore()

    def showPopup(self):
        self._ensure_popup_width()
        super().showPopup()
        self._ensure_popup_width()

    def _popup_display_texts(self) -> list[str]:
        texts: list[str] = []
        for index in range(self.count()):
            text = self.itemText(index).strip()
            if text:
                texts.append(text)
        current_text = self.currentText().strip()
        if current_text:
            texts.append(current_text)
        return texts

    def _ensure_popup_width(self) -> None:
        from PySide6.QtGui import QFontMetrics

        view = self.view()
        if view is None:
            return

        view.setTextElideMode(Qt.ElideNone)
        metrics = QFontMetrics(view.font())
        max_text_width = 0
        for text in self._popup_display_texts():
            max_text_width = max(max_text_width, metrics.horizontalAdvance(text))

        popup_width = max(self.width(), max_text_width + 56)
        view.setMinimumWidth(popup_width)
        popup_window = view.window()
        if popup_window is not None:
            popup_window.setMinimumWidth(popup_width)


class ScrollSafeSpinBox(QSpinBox):
    """QSpinBox that ignores mouse wheel events to prevent accidental value changes."""

    def wheelEvent(self, event):
        event.ignore()


class ScrollSafeDoubleSpinBox(QDoubleSpinBox):
    """QDoubleSpinBox that ignores mouse wheel events to prevent accidental value changes."""

    def wheelEvent(self, event):
        event.ignore()


class FontSizeComboBox(ScrollSafeComboBox):
    """Editable font-size combo that opens at the current exact value when possible."""

    _SECTION_ROLE = Qt.UserRole + 1
    _CUSTOM_ITEM_ROLE = Qt.UserRole + 2
    _SECTION_NAMED = "named"
    _SECTION_NUMERIC = "numeric"

    def addFontSizeItem(self, text: str, pt: float, *, section: str) -> None:
        self.addItem(text, pt)
        self.setItemData(self.count() - 1, section, self._SECTION_ROLE)

    def showPopup(self):
        index = self._find_popup_target_index()
        super().showPopup()
        if index >= 0:
            model_index = self.model().index(index, 0)
            self.view().setCurrentIndex(model_index)
            self.view().scrollTo(model_index, QAbstractItemView.PositionAtCenter)

    def _find_popup_target_index(self) -> int:
        self._clear_custom_font_size_item()
        text = self.currentText().strip()
        if not text:
            return self.currentIndex()

        target_section = (
            self._SECTION_NAMED
            if is_named_font_size_input(text)
            else self._SECTION_NUMERIC
        )

        exact_index = self._find_exact_index(text, target_section)
        if exact_index >= 0:
            return exact_index

        try:
            target_value = parse_font_size_input(text)
        except ValueError:
            return self.currentIndex()

        value_index = self._find_value_index(target_value, target_section)
        if value_index >= 0:
            return value_index

        if target_section == self._SECTION_NUMERIC:
            return self._insert_custom_numeric_value(target_value)

        return self.currentIndex()

    def _find_exact_index(self, text: str, section: str) -> int:
        normalized_text = text.strip()
        for index in self._font_size_indices(section):
            if self.itemText(index).strip() == normalized_text:
                return index
        return -1

    def _font_size_indices(self, section: str) -> list[int]:
        indices = []
        for index in range(self.count()):
            if self.itemData(index, self._SECTION_ROLE) == section:
                indices.append(index)
        return indices

    def _find_value_index(self, target_value: float, section: str) -> int:
        for index in self._font_size_indices(section):
            item_value = self.itemData(index)
            if item_value is None:
                continue
            if abs(float(item_value) - float(target_value)) < 0.01:
                return index
        return -1

    def _clear_custom_font_size_item(self) -> None:
        for index in range(self.count() - 1, -1, -1):
            if self.itemData(index, self._CUSTOM_ITEM_ROLE):
                self.removeItem(index)

    def _insert_custom_numeric_value(self, value: float) -> int:
        numeric_indices = self._font_size_indices(self._SECTION_NUMERIC)
        if not numeric_indices:
            self.addFontSizeItem(format_font_size_pt(value), value, section=self._SECTION_NUMERIC)
            index = self.count() - 1
            self.setItemData(index, True, self._CUSTOM_ITEM_ROLE)
            return index

        insert_index = numeric_indices[-1] + 1
        for index in numeric_indices:
            item_value = self.itemData(index)
            if item_value is None:
                continue
            if float(value) < float(item_value):
                insert_index = index
                break

        self.insertItem(insert_index, format_font_size_pt(value), value)
        self.setItemData(insert_index, self._SECTION_NUMERIC, self._SECTION_ROLE)
        self.setItemData(insert_index, True, self._CUSTOM_ITEM_ROLE)
        return insert_index


class OnOffToggleButton(QPushButton):
    """A small checkable button that toggles between 开启 / 关闭."""

    def __init__(self, checked: bool = False, parent=None, *, is_placeholder: bool = False):
        super().__init__(parent)
        self._is_placeholder = bool(is_placeholder)
        self._control_height = 24
        self.setCheckable(True)
        self.setChecked(bool(checked))
        self.setMinimumWidth(34)
        self.setMaximumWidth(38)
        self.setFixedHeight(24)
        self.toggled.connect(self._sync_label)
        self._sync_label(self.isChecked())

    def setFixedHeight(self, h: int) -> None:
        self._control_height = max(22, int(h or 24))
        super().setFixedHeight(self._control_height)
        self._sync_style(self.isChecked())

    def _sync_label(self, checked: bool) -> None:
        self.setText("开" if checked else "关")
        self._sync_style(checked)

    def _sync_style(self, checked: bool) -> None:
        text_color = "#7a7a7a" if self._is_placeholder else "#2e5f35"
        inner_height = max(18, self._control_height - 2)
        common = (
            f"min-height: {inner_height}px;"
            f"max-height: {inner_height}px;"
            "padding: 0 6px;"
        )
        if checked:
            self.setStyleSheet(
                f"""
                QPushButton {{
                    background-color: #dff3df;
                    border: 1px solid #b7d8b7;
                    border-radius: 4px;
                    {common}
                    color: {text_color};
                }}
                """
            )
            return

        placeholder_style = "color: #7a7a7a;" if self._is_placeholder else ""
        self.setStyleSheet(
            f"""
            QPushButton {{
                border-radius: 4px;
                {common}
                {placeholder_style}
            }}
            """
        )


class PreviewGroupHeaderView(QHeaderView):
    """Header view that paints a merged-looking '预览' group over the first two columns."""

    def paintSection(self, painter, rect, logicalIndex):
        if logicalIndex == 0:
            span_width = rect.width()
            if self.isSectionHidden(1):
                span_width = rect.width()
            else:
                span_width += self.sectionSize(1)
            merged_rect = rect.adjusted(0, 0, span_width - rect.width(), 0)
            option = QStyleOptionHeader()
            self.initStyleOption(option)
            option.rect = merged_rect
            option.text = "预览"
            option.textAlignment = Qt.AlignCenter
            option.position = QStyleOptionHeader.OnlyOneSection
            self.style().drawControl(QStyle.CE_Header, option, painter, self)
            return
        if logicalIndex == 1:
            return
        super().paintSection(painter, rect, logicalIndex)


class StyleHoverPopup(QLabel):
    """Non-blocking hover popup for style details."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.ToolTip | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_ShowWithoutActivating, True)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.setWordWrap(True)
        self.setMargin(8)
        self.setMinimumWidth(220)
        self.setMaximumWidth(320)
        self.setFrameShape(QFrame.Box)
        self.setFrameShadow(QFrame.Plain)
        self.setStyleSheet(
            """
            QLabel {
                background-color: rgba(255, 255, 255, 245);
                border: 1px solid #dcdfe6;
                border-radius: 6px;
                padding: 4px;
            }
            """
        )


class IndentValueWidget(QWidget):
    """Indent editor with a switchable unit (chars / cm / pt)."""

    valueChanged = Signal(str)
    unitChanged = Signal(str)
    _UNIT_ORDER = ("chars", "cm", "pt")
    _UNIT_LABELS = {"chars": "字", "cm": "cm", "pt": "磅"}

    def __init__(self, value, size_pt, parent=None, *, unit: str = "chars", is_placeholder: bool = False):
        super().__init__(parent)
        self._size_pt = self._normalize_size_pt(size_pt)
        self._active_unit = normalize_indent_unit(unit)
        self._canonical_pt_value = self._unit_value_to_pt(self._coerce_numeric(value, default=0.0), self._active_unit)
        self._is_placeholder = bool(is_placeholder)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 0, 2, 0)
        layout.setSpacing(2)

        self._edit = QLineEdit(self)
        self._edit.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self._edit.setFixedWidth(42)
        self._edit.setFixedHeight(24)
        self._edit.textChanged.connect(self._on_text_changed)
        self._edit.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        self._unit_button = QPushButton(self)
        self._unit_button.setCursor(Qt.PointingHandCursor)
        self._unit_button.setFixedSize(34, 24)
        self._unit_button.setFocusPolicy(Qt.NoFocus)
        self._unit_button.clicked.connect(self._toggle_unit)

        if is_placeholder:
            color = "rgb(180, 180, 180)"
            self._edit.setStyleSheet(f"color: {color};")

        layout.addWidget(self._edit)
        layout.addWidget(self._unit_button)
        self.setFixedHeight(24)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        self.setConfigValue(value, unit=self._active_unit)

    @staticmethod
    def _coerce_numeric(value, *, default: float = 0.0) -> float:
        try:
            return float(value)
        except (TypeError, ValueError):
            return float(default)

    @staticmethod
    def _normalize_size_pt(size_pt) -> float:
        try:
            if isinstance(size_pt, str):
                value = parse_font_size_input(size_pt)
            else:
                value = float(size_pt)
        except (TypeError, ValueError):
            value = 12.0
        return value if value > 0 else 12.0

    def _unit_value_to_pt(self, value: float, unit: str) -> float:
        normalized_unit = normalize_indent_unit(unit)
        if normalized_unit == "pt":
            return value
        if normalized_unit == "cm":
            return value * 72.0 / 2.54
        return value * self._size_pt

    def _pt_to_unit_value(self, value_pt: float, unit: str) -> float:
        return resolve_pt_indent_value(value_pt, normalize_indent_unit(unit), self._size_pt)

    def text(self) -> str:
        return self._edit.text().strip()

    def unit(self) -> str:
        return self._active_unit

    def _is_empty_value(self) -> bool:
        return self.text() == "" and self._canonical_pt_value <= 0

    def displayValue(self) -> float:
        value = self._parse_current_value(None)
        if value is not None:
            return value
        return self._pt_to_unit_value(self._canonical_pt_value, self._active_unit)

    def setUnit(self, unit: str) -> None:
        self._active_unit = normalize_indent_unit(unit)
        self._refresh_unit_chrome()

    def setReferenceSize(self, size_pt) -> None:
        try:
            if isinstance(size_pt, str):
                new_size_pt = parse_font_size_input(size_pt)
            else:
                new_size_pt = float(size_pt)
        except (TypeError, ValueError):
            return
        if new_size_pt <= 0:
            return
        self._size_pt = new_size_pt
        self._refresh_unit_chrome()

    def setConfigValue(self, value, *, unit: str | None = None) -> None:
        if unit is not None:
            self._active_unit = normalize_indent_unit(unit)
        self._canonical_pt_value = self._unit_value_to_pt(self._coerce_numeric(value, default=0.0), self._active_unit)
        self._refresh_text_from_cache()

    def charsValue(self) -> float:
        raw = self.text()
        if not raw:
            return 0.0
        try:
            value = float(raw)
        except (TypeError, ValueError) as exc:
            raise ValueError("invalid indent value") from exc
        return self._unit_value_to_pt(value, self._active_unit) / self._size_pt if self._size_pt else value

    def configValue(self) -> float:
        raw = self.text()
        if not raw:
            return 0.0
        try:
            return float(raw)
        except (TypeError, ValueError) as exc:
            raise ValueError("invalid indent value") from exc

    def _parse_current_value(self, fallback: float | None) -> float | None:
        text = self.text()
        if not text:
            return fallback
        try:
            return float(text)
        except (TypeError, ValueError):
            return fallback

    def _refresh_text_from_cache(self) -> None:
        value_text = ""
        if self._canonical_pt_value > 0:
            value_text = format_font_size_pt(self._pt_to_unit_value(self._canonical_pt_value, self._active_unit))
        was_blocked = self._edit.blockSignals(True)
        self._edit.setText(value_text)
        self._edit.blockSignals(was_blocked)
        self._refresh_unit_chrome()

    def _refresh_unit_chrome(self) -> None:
        unit_label = self._UNIT_LABELS.get(self._active_unit, "字")
        self._edit.setToolTip(f"当前单位：{unit_label}")
        self._unit_button.setText(unit_label)
        self._unit_button.setToolTip(f"点击切换单位：当前为{unit_label}")
        self._sync_edit_style()
        self._sync_unit_button_style()

    def _on_text_changed(self, _text: str) -> None:
        value = self._parse_current_value(None)
        if self.text() == "":
            self._canonical_pt_value = 0.0
        elif value is not None:
            self._canonical_pt_value = self._unit_value_to_pt(value, self._active_unit)
        self._sync_edit_style()
        self._sync_unit_button_style()
        self.valueChanged.emit(self.text())

    def _toggle_unit(self) -> None:
        try:
            idx = self._UNIT_ORDER.index(self._active_unit)
        except ValueError:
            idx = 0
        self._active_unit = self._UNIT_ORDER[(idx + 1) % len(self._UNIT_ORDER)]
        self._refresh_unit_chrome()
        self.unitChanged.emit(self._active_unit)

    def _sync_unit_button_style(self) -> None:
        is_empty = self._is_empty_value()
        text_color = "#9ca3af" if (self._is_placeholder or is_empty) else "#111111"
        bg_color = "#f3f4f6" if not is_empty else "#f5f6f7"
        border_color = "#d1d5db" if not is_empty else "#e5e7eb"
        self._unit_button.setStyleSheet(
            f"""
            QPushButton {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-radius: 4px;
                padding: 0 4px;
                color: {text_color};
                font-weight: 600;
            }}
            QPushButton:pressed {{
                background-color: #e5e7eb;
            }}
            """
        )

    def _sync_edit_style(self) -> None:
        is_empty = self._is_empty_value()
        if self._is_placeholder:
            text_color = "rgb(180, 180, 180)"
        elif is_empty:
            text_color = "#9ca3af"
        else:
            text_color = "#111111"
        background = "#f5f6f7" if is_empty else "#ffffff"
        border = "#e5e7eb" if is_empty else "#d1d5db"
        self._edit.setStyleSheet(
            f"""
            QLineEdit {{
                color: {text_color};
                background-color: {background};
                border: 1px solid {border};
                border-radius: 4px;
                padding: 0 4px;
            }}
            """
        )


class SpecialIndentWidget(QWidget):
    """Composite widget: mode button + indent value/unit editor."""

    valueChanged = Signal(str)
    unitChanged = Signal(str)
    modeChanged = Signal(str)
    _MODE_ORDER = ("none", "first_line", "hanging")
    _MODE_LABELS = {"none": "无", "first_line": "首行", "hanging": "悬挂"}

    def __init__(self, mode, value, size_pt, parent=None, *, unit: str = "chars", is_placeholder: bool = False):
        super().__init__(parent)
        self._mode = normalize_special_indent_mode(mode)
        self._is_placeholder = bool(is_placeholder)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 0, 2, 0)
        layout.setSpacing(2)

        self._mode_button = QPushButton(self)
        self._mode_button.setCursor(Qt.PointingHandCursor)
        self._mode_button.setFixedSize(40, 24)
        self._mode_button.setFocusPolicy(Qt.NoFocus)
        self._mode_button.clicked.connect(self._toggle_mode)

        self._value_widget = IndentValueWidget(
            value,
            size_pt,
            self,
            unit=unit,
            is_placeholder=is_placeholder,
        )
        self._value_widget.valueChanged.connect(self.valueChanged.emit)
        self._value_widget.unitChanged.connect(self.unitChanged.emit)

        layout.addWidget(self._mode_button)
        layout.addWidget(self._value_widget)
        self.setFixedHeight(24)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self._sync_mode_ui()

    def text(self) -> str:
        return self._value_widget.text()

    def mode(self) -> str:
        return self._mode

    def unit(self) -> str:
        return self._value_widget.unit()

    def configValue(self) -> float:
        if self._mode == "none":
            return 0.0
        return self._value_widget.configValue()

    def setConfigValue(self, value, *, unit: str | None = None) -> None:
        self._value_widget.setConfigValue(value, unit=unit)

    def setUnit(self, unit: str) -> None:
        self._value_widget.setUnit(unit)

    def setReferenceSize(self, size_pt) -> None:
        self._value_widget.setReferenceSize(size_pt)

    def setMode(self, mode: str) -> None:
        normalized = normalize_special_indent_mode(mode)
        changed = normalized != self._mode
        self._mode = normalized
        self._sync_mode_ui()
        if changed:
            self.modeChanged.emit(self._mode)

    def _toggle_mode(self) -> None:
        try:
            idx = self._MODE_ORDER.index(self._mode)
        except ValueError:
            idx = 0
        self.setMode(self._MODE_ORDER[(idx + 1) % len(self._MODE_ORDER)])

    def _sync_mode_ui(self) -> None:
        label = self._MODE_LABELS.get(self._mode, "无")
        enabled = self._mode != "none"
        self._mode_button.setText(label)
        self._mode_button.setToolTip(f"点击切换特殊缩进模式：当前为{label}")
        text_color = "#6b7280" if self._is_placeholder else "#111111"
        self._mode_button.setStyleSheet(
            f"""
            QPushButton {{
                background-color: #f3f4f6;
                border: 1px solid #d1d5db;
                border-radius: 4px;
                padding: 0 4px;
                color: {text_color};
                font-weight: 600;
            }}
            QPushButton:pressed {{
                background-color: #e5e7eb;
            }}
            """
        )
        self._value_widget.setEnabled(enabled)


class LineSpacingValueWidget(QWidget):
    """Line-spacing value editor with a dynamic unit label."""

    valueChanged = Signal(str)
    spacingTypeChanged = Signal(str)

    def __init__(self, spacing_type: str, value, parent=None, *, is_placeholder: bool = False):
        super().__init__(parent)
        self._last_exact_value = resolve_line_spacing_value("exact", value)
        normalized_kind = normalize_line_spacing_type(spacing_type)
        self._last_multiple_value = resolve_line_spacing_value(
            normalized_kind if normalized_kind != "exact" else "multiple",
            value,
        )
        self._is_placeholder = bool(is_placeholder)
        self._active_kind = "exact"

        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 0, 2, 0)
        layout.setSpacing(2)

        self._edit = QLineEdit(self)
        self._edit.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self._edit.setFixedWidth(42)
        self._edit.setFixedHeight(24)
        self._edit.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self._edit.textChanged.connect(self._on_text_changed)
        self._unit_badge = QPushButton(self)
        self._unit_badge.setFocusPolicy(Qt.NoFocus)
        self._unit_badge.setCursor(Qt.PointingHandCursor)
        self._unit_badge.setFixedSize(28, 24)
        self._unit_badge.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self._unit_badge.clicked.connect(self._toggle_spacing_type)

        if is_placeholder:
            color = "rgb(180, 180, 180)"
            self._edit.setStyleSheet(f"color: {color};")

        layout.addWidget(self._edit)
        layout.addWidget(self._unit_badge)
        self.setFixedHeight(24)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        self.setSpacingType(spacing_type, value)

    def text(self) -> str:
        return self._edit.text().strip()

    def spacingType(self) -> str:
        return self._active_kind

    def currentValue(self) -> float | None:
        try:
            value = float(self.text())
        except (TypeError, ValueError):
            return None
        return value if value > 0 else None

    def displayValue(self) -> float:
        value = self.currentValue()
        if value is not None:
            return value
        return self._last_exact_value if self._active_kind == "exact" else self._last_multiple_value

    def resolvedValue(self) -> float:
        value = self.currentValue()
        if value is None:
            raise ValueError("invalid line spacing value")
        return value

    def configSpacingType(self, *, use_fallback: bool = False) -> str:
        if self._active_kind == "exact":
            return "exact"

        value = self.displayValue() if use_fallback else self.resolvedValue()
        if abs(value - 1.0) < 0.01:
            return "single"
        if abs(value - 1.5) < 0.01:
            return "one_half"
        if abs(value - 2.0) < 0.01:
            return "double"
        return "multiple"

    def setSpacingType(self, spacing_type: str, value=None) -> None:
        kind = normalize_line_spacing_type(spacing_type)
        if kind == "exact":
            resolved_value = (
                resolve_line_spacing_value(kind, value)
                if value is not None else self._last_exact_value
            )
            self._last_exact_value = resolved_value
        else:
            resolved_value = (
                resolve_line_spacing_value(kind, value)
                if value is not None else self._last_multiple_value
            )
            self._last_multiple_value = resolved_value

        self._active_kind = "exact" if kind == "exact" else "multiple"

        was_blocked = self._edit.blockSignals(True)
        self._edit.setText(format_font_size_pt(resolved_value))
        self._edit.blockSignals(was_blocked)
        self._edit.setEnabled(True)
        unit_label = line_spacing_unit_label(self._active_kind)
        self._unit_badge.setText(unit_label)
        self._edit.setToolTip(f"当前单位：{unit_label}")
        self._unit_badge.setToolTip(f"点击切换类型：当前为{unit_label}")
        self._sync_unit_badge_style()

    def _sync_unit_badge_style(self) -> None:
        text_color = "#6b7280" if self._is_placeholder else "#111111"
        self._unit_badge.setStyleSheet(
            f"""
            QPushButton {{
                background-color: #f3f4f6;
                border: 1px solid #d1d5db;
                border-radius: 4px;
                padding: 0 4px;
                color: {text_color};
                font-weight: 600;
            }}
            QPushButton:pressed {{
                background-color: #e5e7eb;
            }}
            """
        )

    def _on_text_changed(self, _text: str) -> None:
        try:
            value = float(self.text())
        except (TypeError, ValueError):
            self.valueChanged.emit(self.text())
            return
        if self._active_kind == "exact":
            self._last_exact_value = value
        else:
            self._last_multiple_value = value
        self.valueChanged.emit(self.text())

    def _toggle_spacing_type(self) -> None:
        if self._active_kind == "exact":
            try:
                self._last_exact_value = float(self.text())
            except (TypeError, ValueError):
                pass
            self._active_kind = "multiple"
            value = self._last_multiple_value
        else:
            try:
                self._last_multiple_value = float(self.text())
            except (TypeError, ValueError):
                pass
            self._active_kind = "exact"
            value = self._last_exact_value
        was_blocked = self._edit.blockSignals(True)
        self._edit.setText(format_font_size_pt(value))
        self._edit.blockSignals(was_blocked)
        unit_label = line_spacing_unit_label(self._active_kind)
        self._unit_badge.setText(unit_label)
        self._edit.setToolTip(f"当前单位：{unit_label}")
        self._unit_badge.setToolTip(f"点击切换类型：当前为{unit_label}")
        self._sync_unit_badge_style()
        self.spacingTypeChanged.emit(self._active_kind)


class FontSearchFilterProxyModel(QSortFilterProxyModel):
    """Filter installed fonts with case-insensitive fuzzy matching."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._query = ""

    def setQuery(self, query: str) -> None:
        self._query = str(query or "")
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row: int, source_parent) -> bool:
        if not self._query.strip():
            return True
        index = self.sourceModel().index(source_row, 0, source_parent)
        font_name = self.sourceModel().data(index, Qt.DisplayRole)
        return font_matches_query(font_name or "", self._query)


class FontFamilyComboBox(ScrollSafeComboBox):
    """Editable font-family combo with installed-font dropdown and fuzzy search."""

    _SHARED_FONT_MODEL: QStringListModel | None = None
    _SHARED_POPUP_MIN_WIDTH: int | None = None

    def __init__(self, font_families: list[str], parent=None, *, is_placeholder: bool = False):
        super().__init__(parent)
        self._font_families = list(font_families)
        self._font_filter_model: FontSearchFilterProxyModel | None = None
        self._font_completer: QCompleter | None = None
        self._items_loaded = False
        self._user_is_editing_query = False
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)
        self.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.setMinimumContentsLength(4)
        self.setMaxVisibleItems(16)

        line_edit = self.lineEdit()
        if line_edit is not None:
            if is_placeholder:
                line_edit.setStyleSheet("color: rgb(180, 180, 180);")
            line_edit.textEdited.connect(self._on_text_edited)
            line_edit.editingFinished.connect(self._on_editing_finished)
        self.activated.connect(self._on_item_activated)

    def mousePressEvent(self, event):
        self._ensure_font_items_loaded()
        super().mousePressEvent(event)

    def showPopup(self):
        self._ensure_font_items_loaded()
        self._ensure_popup_width()
        if self._font_filter_model is not None:
            self._font_filter_model.setQuery("")
        super().showPopup()
        self._ensure_popup_width()
        self._center_popup_on_current_value()

    def _on_text_edited(self, text: str) -> None:
        self._user_is_editing_query = True
        self._ensure_font_search_loaded()
        if self._font_filter_model is None or self._font_completer is None:
            return
        self._font_filter_model.setQuery(text)
        self._ensure_completer_popup_width()
        if not text.strip():
            return
        self._font_completer.setCompletionPrefix(text)
        self._font_completer.complete()
        QTimer.singleShot(0, lambda text=text: self._restore_query_text_state(text))

    def _on_item_activated(self, *_args) -> None:
        self._user_is_editing_query = False
        QTimer.singleShot(0, self.reveal_text_start)

    def _on_editing_finished(self) -> None:
        self._user_is_editing_query = False

    def _ensure_font_items_loaded(self) -> None:
        if not self._items_loaded:
            current_text = self.currentText()
            self.addItems(self._font_families)
            if current_text:
                self.setCurrentText(current_text)
                self.reveal_text_start()
            self._items_loaded = True

    def _ensure_font_search_loaded(self) -> None:
        self._ensure_font_items_loaded()
        if self._font_filter_model is not None and self._font_completer is not None:
            return

        if self.__class__._SHARED_FONT_MODEL is None:
            self.__class__._SHARED_FONT_MODEL = QStringListModel(self._font_families)

        self._font_filter_model = FontSearchFilterProxyModel(self)
        self._font_filter_model.setSourceModel(self.__class__._SHARED_FONT_MODEL)
        self._font_completer = QCompleter(self._font_filter_model, self)
        self._font_completer.setCaseSensitivity(Qt.CaseInsensitive)
        self._font_completer.setCompletionMode(QCompleter.PopupCompletion)
        self._font_completer.activated.connect(self._on_item_activated)
        popup = self._font_completer.popup()
        if popup is not None:
            popup.setTextElideMode(Qt.ElideNone)
        self.setCompleter(self._font_completer)
        self._ensure_completer_popup_width()

    def _restore_query_text_state(self, text: str) -> None:
        if not self._user_is_editing_query:
            return
        line_edit = self.lineEdit()
        if line_edit is None or not line_edit.hasFocus():
            return

        if line_edit.text() != text:
            was_blocked = line_edit.blockSignals(True)
            line_edit.setText(text)
            line_edit.blockSignals(was_blocked)
        line_edit.deselect()
        line_edit.setCursorPosition(len(text))

    def reveal_text_start(self) -> None:
        line_edit = self.lineEdit()
        if line_edit is None:
            return
        line_edit.setCursorPosition(0)
        line_edit.home(False)

    def _ensure_popup_width(self) -> None:
        from PySide6.QtGui import QFontMetrics

        if self.__class__._SHARED_POPUP_MIN_WIDTH is None:
            metrics = QFontMetrics(self.view().font())
            max_text_width = 0
            for family in self._font_families:
                max_text_width = max(max_text_width, metrics.horizontalAdvance(family))
            self.__class__._SHARED_POPUP_MIN_WIDTH = max(max_text_width + 48, 220)

        popup_width = max(self.width(), self.__class__._SHARED_POPUP_MIN_WIDTH)
        self.view().setMinimumWidth(popup_width)
        self.view().window().setMinimumWidth(popup_width)
        self._ensure_completer_popup_width(popup_width)

    def _ensure_completer_popup_width(self, popup_width: int | None = None) -> None:
        if self._font_completer is None:
            return
        popup = self._font_completer.popup()
        if popup is None:
            return
        if popup_width is None:
            if self.__class__._SHARED_POPUP_MIN_WIDTH is None:
                self._ensure_popup_width()
                return
            popup_width = max(self.width(), self.__class__._SHARED_POPUP_MIN_WIDTH)
        popup.setMinimumWidth(popup_width)
        popup.window().setMinimumWidth(popup_width)
        popup.setTextElideMode(Qt.ElideNone)

    def _center_popup_on_current_value(self) -> None:
        index = self.findText(self.currentText().strip(), Qt.MatchFixedString)
        if index < 0:
            index = self.currentIndex()
        if index < 0:
            return
        model_index = self.model().index(index, 0)
        self.view().setCurrentIndex(model_index)
        self.view().scrollTo(model_index, QAbstractItemView.PositionAtCenter)

class WhitespaceVisibleLineEdit(QLineEdit):
    """QLineEdit that shows whitespace as visible symbols.

    Display: full-width space → □, regular space → ·, tab → ➡
    The actual raw value is returned by text() transparently.
    """

    _TO_SYMBOL = {
        "\u3000": "\u25A1",   # □  full-width space
        " ": "\u00B7",        # ·  regular space
        "\t": "\u27A1",       # ➡  tab
    }
    _FROM_SYMBOL = {
        "\u25A1": "\u3000",   # □ → full-width space
        "\u00B7": " ",        # · → regular space
        "\u27A1": "\t",       # ➡ → tab
    }

    def __init__(self, text="", parent=None):
        super().__init__(parent)
        self.setRawText(text)

    def setRawText(self, raw: str) -> None:
        """Set the actual value; display shows visible symbols."""
        display = "".join(self._TO_SYMBOL.get(ch, ch) for ch in raw)
        super().setText(display)

    def text(self) -> str:
        """Return the actual raw value with real whitespace chars."""
        display = super().text()
        return "".join(self._FROM_SYMBOL.get(ch, ch) for ch in display)


class WhitespacePresetWidget(QWidget):
    """Compact preset/custom whitespace editor used in table cells."""

    textChanged = Signal(str)

    def __init__(
        self,
        options: list[tuple[str, str, str | None]],
        value_to_key: dict[str, str],
        *,
        text: str = "",
        compact_labels: dict[str, str] | None = None,
        control_height: int = 24,
        compact_mode_width: int = 52,
        parent=None,
    ):
        super().__init__(parent)
        self._preset_values = {key: value for key, _, value in options if value is not None}
        self._preset_labels = {key: label for key, label, _ in options}
        self._value_to_key = dict(value_to_key)
        self._compact_labels = dict(compact_labels or {})
        self._custom_raw_text = ""
        self._current_mode_key = "custom"
        self._control_height = max(24, int(control_height or 24))
        self._compact_mode_width = max(48, int(compact_mode_width or 52))
        self._custom_layout_spacing = 2
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.setMinimumWidth(0)

        self._mode_combo = ScrollSafeComboBox(self)
        self._mode_combo.setFixedHeight(self._control_height)
        self._mode_combo.setMinimumWidth(0)
        self._mode_combo.setEditable(True)
        self._mode_combo.setInsertPolicy(QComboBox.NoInsert)
        self._mode_combo.setToolTip("选择常用分隔符，或切换到自定义。")
        self._mode_combo.setStyleSheet(
            """
            QComboBox {
                padding: 0 10px 0 2px;
            }
            QComboBox::drop-down {
                width: 10px;
            }
            """
        )
        self._mode_combo.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        combo_line_edit = self._mode_combo.lineEdit()
        if combo_line_edit is not None:
            combo_line_edit.setReadOnly(True)
            combo_line_edit.setAlignment(Qt.AlignCenter)
            combo_line_edit.setMinimumWidth(0)
            combo_line_edit.setTextMargins(0, 0, 0, 0)
        for key, label, _ in options:
            self._mode_combo.addItem(label, key)
            self._mode_combo.setItemData(self._mode_combo.count() - 1, label, Qt.ToolTipRole)

        self._edit = WhitespaceVisibleLineEdit(parent=self)
        self._edit.setFixedHeight(self._control_height)
        self._edit.setToolTip("输入时：全角空格=□，半角空格=·，Tab=➡。")
        self._edit.setVisible(False)
        self._mode_combo.view().setMinimumWidth(146)
        self._edit.setMinimumWidth(0)
        self._edit.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)

        self._mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        self._mode_combo.activated.connect(lambda *_args: QTimer.singleShot(0, self._on_mode_activated))
        self._edit.textChanged.connect(lambda _text: self._on_edit_text_changed())
        self.setRawText(text)

    def setRawText(self, raw: str) -> None:
        raw_text = str(raw or "")
        mode_key = self._value_to_key.get(raw_text, "custom")
        if mode_key == "custom":
            self._custom_raw_text = raw_text
        self._set_mode(mode_key, emit=False)

    def text(self) -> str:
        if self._current_mode_key == "custom":
            return self._edit.text()
        return str(self._preset_values.get(self._current_mode_key, "") or "")

    def isCustomMode(self) -> bool:
        return self._current_mode_key == "custom"

    def _set_mode(self, mode_key: str, *, emit: bool) -> None:
        target_mode = str(mode_key or "custom")
        index = self._mode_combo.findData(target_mode)
        blocked_combo = self._mode_combo.blockSignals(True)
        self._mode_combo.setCurrentIndex(max(index, 0))
        self._mode_combo.blockSignals(blocked_combo)

        raw_text = (
            self._custom_raw_text
            if target_mode == "custom"
            else str(self._preset_values.get(target_mode, "") or "")
        )
        blocked_edit = self._edit.blockSignals(True)
        self._edit.setRawText(raw_text)
        is_custom = target_mode == "custom"
        self._edit.setEnabled(is_custom)
        self._edit.setVisible(is_custom)
        self._edit.blockSignals(blocked_edit)
        self._current_mode_key = target_mode
        self._sync_mode_combo_display()
        self._sync_mode_combo_alignment()
        self._relayout_children()
        self.updateGeometry()
        if emit:
            self.textChanged.emit(self.text())

    def _on_mode_changed(self) -> None:
        if self._current_mode_key == "custom":
            self._custom_raw_text = self._edit.text()
        self._set_mode(str(self._mode_combo.currentData() or "custom"), emit=True)

    def _on_edit_text_changed(self) -> None:
        if self._current_mode_key == "custom":
            self._custom_raw_text = self._edit.text()
        self.textChanged.emit(self.text())

    def _on_mode_activated(self) -> None:
        self._set_mode(str(self._mode_combo.currentData() or "custom"), emit=True)

    def _sync_mode_combo_display(self) -> None:
        line_edit = self._mode_combo.lineEdit()
        if line_edit is None:
            return
        mode_key = str(self._mode_combo.currentData() or "custom")
        display_text = self._compact_labels.get(
            mode_key,
            self._preset_labels.get(mode_key, mode_key),
        )
        blocked = line_edit.blockSignals(True)
        line_edit.setText(display_text)
        line_edit.blockSignals(blocked)

    def _sync_mode_combo_alignment(self) -> None:
        line_edit = self._mode_combo.lineEdit()
        if line_edit is None:
            return
        if self._current_mode_key == "custom":
            line_edit.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        else:
            line_edit.setAlignment(Qt.AlignCenter)

    def _minimum_mode_combo_width(self) -> int:
        labels = [
            str(self._compact_labels.get(key, self._preset_labels.get(key, key)) or "")
            for key in self._preset_labels
        ]
        metrics = self._mode_combo.fontMetrics()
        widest_text_width = max((metrics.horizontalAdvance(label) for label in labels), default=0)
        return max(
            self._compact_mode_width,
            widest_text_width + 28,
        )

    def _relayout_children(self) -> None:
        rect = self.rect()
        if rect.width() <= 0 or rect.height() <= 0:
            return
        control_height = min(self._control_height, rect.height())
        y = rect.y() + max(0, (rect.height() - control_height) // 2)
        if self._current_mode_key != "custom":
            self._mode_combo.setGeometry(rect.x(), y, rect.width(), control_height)
            self._edit.setGeometry(rect.x(), y, 0, control_height)
            return

        spacing = self._custom_layout_spacing
        combo_width = min(self._minimum_mode_combo_width(), rect.width())
        if combo_width >= rect.width():
            spacing = 0
        edit_width = max(0, rect.width() - combo_width - spacing)
        self._mode_combo.setGeometry(rect.x(), y, combo_width, control_height)
        self._edit.setGeometry(rect.x() + combo_width + spacing, y, edit_width, control_height)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._relayout_children()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self._relayout_children()

    def sizeHint(self) -> QSize:
        base_width = self._minimum_mode_combo_width()
        if self._current_mode_key == "custom":
            base_width += self._custom_layout_spacing + 40
        return QSize(base_width, self._control_height)

    def minimumSizeHint(self) -> QSize:
        return QSize(self._minimum_mode_combo_width(), self._control_height)


class NumberingCellWidgetContainer(QWidget):
    """Cell wrapper that keeps the inner widget strictly inside the table column."""

    def __init__(self, widget: QWidget, *, center_horizontally: bool = False, parent=None):
        super().__init__(parent)
        self._num_inner_widget = widget
        self._center_horizontally = bool(center_horizontally)
        self._horizontal_margin = 0
        self._vertical_margin = 4
        widget.setParent(self)
        widget.show()
        QTimer.singleShot(0, self._relayout_child)

    def _content_rect(self):
        rect = self.contentsRect().adjusted(
            self._horizontal_margin,
            self._vertical_margin,
            -self._horizontal_margin,
            -self._vertical_margin,
        )
        if rect.width() < 0:
            rect.setWidth(0)
        if rect.height() < 0:
            rect.setHeight(0)
        return rect

    def _inner_width(self, available_width: int) -> int:
        if not self._center_horizontally:
            return max(0, int(available_width))
        widget = self._num_inner_widget
        max_width = widget.maximumWidth()
        min_width = widget.minimumWidth()
        if 0 < max_width < 16777215 and max_width == min_width:
            return min(available_width, max_width)
        preferred_width = int(getattr(widget, "_num_preferred_width", 0) or 0)
        if preferred_width <= 0:
            current_width = int(widget.width() or 0)
            preferred_width = current_width if current_width > 0 else int(widget.sizeHint().width() or 0)
        if preferred_width <= 0:
            preferred_width = available_width
        return min(available_width, preferred_width)

    def _relayout_child(self) -> None:
        widget = self._num_inner_widget
        rect = self._content_rect()
        if rect.width() <= 0 or rect.height() <= 0:
            widget.setGeometry(rect.x(), rect.y(), 0, 0)
            return

        target_width = self._inner_width(rect.width())
        target_height = rect.height()
        max_height = widget.maximumHeight()
        min_height = widget.minimumHeight()
        if 0 < max_height < 16777215:
            target_height = min(target_height, max_height)
        if min_height > 0 and min_height == max_height:
            target_height = min(rect.height(), min_height)
        x = rect.x()
        if self._center_horizontally:
            x += max(0, (rect.width() - target_width) // 2)
        y = rect.y() + max(0, (rect.height() - target_height) // 2)
        widget.setGeometry(x, y, target_width, target_height)
        relayout = getattr(widget, "_relayout_children", None)
        if callable(relayout):
            relayout()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._relayout_child()

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self._relayout_child()
        QTimer.singleShot(0, self._relayout_child)


class PipelineWorker(QThread):
    """后台线程执行排版 Pipeline"""
    progress = Signal(int, int, str)
    finished = Signal(object)

    def __init__(self, config: SceneConfig, doc_path: str):
        super().__init__()
        self.config = config
        self.doc_path = doc_path
        self.result: PipelineResult | None = None
        self._cancel_requested = False

    def run(self):
        try:
            pipeline = Pipeline(
                self.config,
                progress_callback=self._on_progress,
                cancel_requested=self._is_cancel_requested,
            )
            self.result = pipeline.run(self.doc_path)
        except Exception as e:
            self.result = PipelineResult(
                success=False,
                status="failed",
                error=f"后台任务异常: {e}",
            )
        finally:
            if self.result is None:
                self.result = PipelineResult(
                    success=False,
                    status="failed",
                    error="后台任务未返回结果",
                )
            self.finished.emit(self.result)

    def _on_progress(self, current, total, message):
        self.progress.emit(current, total, message)

    def cancel(self):
        self._cancel_requested = True
        self.requestInterruption()

    def _is_cancel_requested(self) -> bool:
        return self._cancel_requested or self.isInterruptionRequested()


# ── 样式名称映射 ──
_STYLE_DISPLAY_NAMES = {
    "normal": "正文",
    "heading1": "一级标题",
    "heading2": "二级标题",
    "heading3": "三级标题",
    "heading4": "四级标题",
    "heading5": "五级标题",
    "heading6": "六级标题",
    "heading7": "七级标题",
    "heading8": "八级标题",
    "abstract_title_cn": "中文摘要标题",
    "abstract_title_en": "英文摘要标题",
    "abstract_body": "中文摘要正文",
    "abstract_body_en": "英文摘要正文",
    "toc_title": "目录标题",
    "toc_chapter": "一级目录",
    "toc_level1": "二级目录",
    "toc_level2": "三级目录",
    "references_body": "参考文献",
    "acknowledgment_body": "致谢正文",
    "appendix_body": "附录正文",
    "resume_body": "个人简历正文",
    "symbol_table_body": "符号注释表",
    "code_block": "代码块",
    "figure_caption": "图题注",
    "table_caption": "表题注",
    "header_cn": "篇眉(中文)",
    "header_en": "篇眉(英文)",
    "page_number": "页码",
}

_STYLE_COLUMNS = [
    ("font_cn", "中文字体", "str"),
    ("font_en", "英文字体", "str"),
    ("size_pt", "字号", "float"),
    ("bold", "加粗", "bool"),
    ("italic", "斜体", "bool"),
    ("alignment", "对齐", "str"),
    ("special_indent_value", "特殊缩进", "float"),
    ("left_indent_chars", "左缩进", "float"),
    ("right_indent_chars", "右缩进", "float"),
    ("line_spacing_pt", "行距", "float"),
    ("space_before_pt", "段前(磅)", "float"),
    ("space_after_pt", "段后(磅)", "float"),
]

_STYLE_PREVIEW_COLUMNS: tuple[tuple[str, str], ...] = ()

SECTION_LABELS = {
    "body": "正文",
    "references": "参考文献",
    "errata": "勘误页",
    "acknowledgment": "致谢",
    "appendix": "附录",
    "resume": "个人简历",
    "abstract_cn": "中文摘要",
    "abstract_en": "英文摘要",
    "toc": "目录",
    "pricing": "报价",
    "qualification": "资质",
    "conclusion": "结论",
}

PAPER_SIZE_OPTIONS = [
    ("A4", "A4"),
    ("Letter", "Letter"),
    ("A3", "A3"),
]

PIPELINE_STEP_LABELS = {
    "page_setup": "页面设置",
    "md_cleanup": "Markdown 文本修复",
    "whitespace_normalize": "文本空白清洗",
    "style_manager": "样式管理",
    "heading_detect": "标题识别",
    "heading_numbering": "标题编号",
    "toc_format": "目录格式",
    "caption_format": "题注格式",
    "table_format": "表格格式",
    "formula_convert": "公式编码统一",
    "formula_to_table": "公式转公式表格",
    "equation_table_format": "公式表格编号",
    "formula_style": "公式格式统一",
    "section_format": "分区排版",
    "citation_link": "正文参考文献域关联",
    "header_footer": "页眉页脚",
    "validation": "校验",
}

PIPELINE_DEFAULT_ORDER = [
    "page_setup",
    "md_cleanup",
    "whitespace_normalize",
    "style_manager",
    "heading_detect",
    "heading_numbering",
    "toc_format",
    "caption_format",
    "table_format",
    "formula_convert",
    "formula_to_table",
    "equation_table_format",
    "formula_style",
    "section_format",
    "citation_link",
    "header_footer",
    "validation",
]


def _mark_pipeline_critical_rules_auto_managed(config: SceneConfig | None) -> None:
    """UI only exposes pipeline steps, so step edits should reset critical rules to auto mode."""
    if config is None:
        return
    config._pipeline_critical_rules_source = "auto"  # type: ignore[attr-defined]
    _sync_pipeline_critical_rules(config)

_MD_ORDERED_LIST_STYLE_OPTIONS = [
    ("mixed", "混合"),
    ("decimal_dot", "1."),
    ("decimal_paren_right", "1)"),
    ("decimal_cn_dun", "1、"),
    ("decimal_full_paren", "(1)"),
]

_MD_UNORDERED_LIST_STYLE_OPTIONS = [
    ("word_default", "混合"),
    ("bullet_dot", "•"),
    ("bullet_circle", "○"),
    ("bullet_square", "■"),
    ("bullet_dash", "-"),
]


class FormatConfigDialog(QDialog):
    """格式配置查看/编辑对话框"""

    _INSTALLED_FONT_FAMILIES_CACHE: list[str] | None = None
    _MIN_DIALOG_WIDTH = 960
    _MIN_DIALOG_HEIGHT = 560
    _PREFERRED_DIALOG_HEIGHT = 760
    _MAX_SCREEN_USAGE = 0.94
    _DIALOG_EXTRA_WIDTH = 40
    _STYLE_TABLE_COLUMN_WIDTHS = {
        "font_cn": 72,
        "font_en": 72,
        "size_pt": 84,
        "bold": 44,
        "italic": 44,
        "alignment": 100,
        "special_indent_value": 150,
        "left_indent_chars": 88,
        "right_indent_chars": 88,
        "line_spacing_pt": 96,
        "space_before_pt": 56,
        "space_after_pt": 56,
    }
    _STYLE_TABLE_ROW_HEIGHT = 32
    _STYLE_TABLE_MIN_ROW_HEADER_WIDTH = 80
    _STYLE_TABLE_MAX_ROW_HEADER_WIDTH = 100

    _LEGACY_HEADING_LEVEL_KEY_MAP = {
        "chapter": "heading1",
        "chapter_heading": "heading1",
        "level1": "heading2",
        "level2": "heading3",
        "level3": "heading4",
    }

    def __init__(self, config: SceneConfig, parent=None, *,
                 scene_path: str | None = None,
                 scene_items: dict[int, "Path"] | None = None):
        super().__init__(parent)
        install_global_wheel_guard(QApplication.instance())
        self.setWindowTitle("格式配置")
        self.setMinimumSize(self._MIN_DIALOG_WIDTH, self._MIN_DIALOG_HEIGHT)
        self._config = config
        self._normalize_heading_numbering_legacy_keys()
        self._scene_path = scene_path  # 当前场景文件路径
        self._scene_items = scene_items or {}  # combo index → path
        self._deleted_scene_path: str | None = None
        self._style_hover_popup: StyleHoverPopup | None = None
        self._style_header_viewport = None
        self._style_hover_row: int = -1
        self._build_ui()
        self._apply_initial_dialog_size()

    @classmethod
    def _normalize_heading_level_key_map(cls, levels: dict) -> tuple[dict, bool]:
        if not isinstance(levels, dict):
            return levels, False

        has_any_legacy = any(k in levels for k in cls._LEGACY_HEADING_LEVEL_KEY_MAP)
        if not has_any_legacy:
            return levels, False

        normalized = {}
        for raw_key, value in levels.items():
            key = str(raw_key or "").strip()
            if key in cls._LEGACY_HEADING_LEVEL_KEY_MAP:
                dst_key = cls._LEGACY_HEADING_LEVEL_KEY_MAP[key]
            elif has_any_legacy and key == "heading1":
                dst_key = "heading2"
            elif has_any_legacy and key == "heading2":
                dst_key = "heading3"
            elif has_any_legacy and key == "heading3":
                dst_key = "heading4"
            else:
                dst_key = key
            normalized[dst_key] = value

        return normalized, normalized != levels

    def _normalize_heading_numbering_legacy_keys(self) -> None:
        hn = getattr(self._config, "heading_numbering", None)
        if hn is None:
            return

        levels = getattr(hn, "levels", {})
        normalized_levels, changed = self._normalize_heading_level_key_map(levels)
        if changed:
            hn.levels = normalized_levels

        schemes = getattr(hn, "schemes", {})
        if not isinstance(schemes, dict):
            return
        for sid, scheme_levels in list(schemes.items()):
            normalized_scheme, scheme_changed = self._normalize_heading_level_key_map(scheme_levels)
            if scheme_changed:
                schemes[sid] = normalized_scheme

    def _build_ui(self):
        layout = QVBoxLayout(self)
        is_protected_scene = False
        if self._scene_path:
            try:
                is_protected_scene = is_protected_scene_path(Path(self._scene_path))
            except Exception:
                is_protected_scene = False

        # ── 格式管理工具栏 ──
        toolbar = QHBoxLayout()
        toolbar.addWidget(QLabel("当前格式:"))
        self._fmt_name_label = QLabel(
            getattr(self._config, "name", "") or "(未命名)")
        self._fmt_name_label.setStyleSheet("font-weight: bold;")
        toolbar.addWidget(self._fmt_name_label, stretch=1)

        self._rename_btn = QPushButton("重命名")
        self._rename_btn.setEnabled(bool(self._scene_path) and not is_protected_scene)
        self._rename_btn.clicked.connect(self._rename_format)
        toolbar.addWidget(self._rename_btn)

        self._delete_btn = QPushButton("删除")
        self._delete_btn.setEnabled(bool(self._scene_path) and not is_protected_scene)
        self._delete_btn.clicked.connect(self._delete_format)
        toolbar.addWidget(self._delete_btn)

        layout.addLayout(toolbar)

        # ── Tab 页 ──
        tabs = QTabWidget()
        self._config_tabs = tabs
        layout.addWidget(tabs)

        tabs.addTab(self._build_styles_tab(), "样式配置")
        tabs.addTab(self._build_numbering_tab(), "标题编号")
        tabs.addTab(self._build_heading_advanced_tab(), "标题高级")
        tabs.addTab(self._build_caption_tab(), "题注配置")
        tabs.addTab(self._build_table_tab(), "表格配置")
        tabs.addTab(self._build_page_tab(), "页面设置")
        tabs.addTab(self._build_output_tab(), "输出选项")
        tabs.addTab(self._build_pipeline_tab(), "执行流程")

        # 底部按钮
        btn_box = QDialogButtonBox()
        self._save_btn = btn_box.addButton("增量保存", QDialogButtonBox.ActionRole)
        self._prepare_dialog_action_button(self._save_btn)
        self._save_btn.setEnabled(bool(self._scene_path))
        self._save_btn.clicked.connect(self._save_format)
        self._apply_btn = btn_box.addButton("应用修改", QDialogButtonBox.AcceptRole)
        self._prepare_dialog_action_button(self._apply_btn)
        self._apply_btn.clicked.connect(self._apply_changes)
        self._close_btn = btn_box.addButton("关闭", QDialogButtonBox.RejectRole)
        self._prepare_dialog_action_button(self._close_btn)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)
        QTimer.singleShot(0, self._focus_config_tabs)

    @staticmethod
    def _prepare_dialog_action_button(button: QPushButton) -> None:
        button.setDefault(False)
        button.setAutoDefault(False)
        button.setFocusPolicy(Qt.ClickFocus)

    def _focus_config_tabs(self) -> None:
        tabs = getattr(self, "_config_tabs", None)
        if tabs is not None:
            tabs.setFocus(Qt.OtherFocusReason)

    def _style_table_total_width(self) -> int:
        table = getattr(self, "_style_table", None)
        if table is None:
            return self._MIN_DIALOG_WIDTH
        width = table.verticalHeader().width()
        for col in range(table.columnCount()):
            width += table.columnWidth(col)
        width += table.frameWidth() * 2
        width += table.style().pixelMetric(QStyle.PM_ScrollBarExtent)
        return width

    def _preferred_dialog_width(self) -> int:
        return max(self._MIN_DIALOG_WIDTH, self._style_table_total_width() + self._DIALOG_EXTRA_WIDTH)

    def _compute_initial_dialog_size(self) -> tuple[int, int]:
        preferred_w = self._preferred_dialog_width()
        preferred_h = max(self._MIN_DIALOG_HEIGHT, self._PREFERRED_DIALOG_HEIGHT)
        screen = self.screen() or QApplication.primaryScreen()
        if screen is None:
            return preferred_w, preferred_h
        avail = screen.availableGeometry()
        cap_w = max(self._MIN_DIALOG_WIDTH, int(avail.width() * self._MAX_SCREEN_USAGE))
        cap_h = max(self._MIN_DIALOG_HEIGHT, int(avail.height() * self._MAX_SCREEN_USAGE))
        return min(cap_w, preferred_w), min(cap_h, preferred_h)

    def _apply_initial_dialog_size(self) -> None:
        target_w, target_h = self._compute_initial_dialog_size()
        self.resize(target_w, target_h)

    def keyPressEvent(self, event) -> None:
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            event.ignore()
            return
        super().keyPressEvent(event)

    def _prompt_format_name(self, title: str, label: str, default_text: str) -> tuple[str, bool]:
        dialog = QInputDialog(self)
        dialog.setInputMode(QInputDialog.TextInput)
        dialog.setWindowTitle(title)
        dialog.setLabelText(label)
        dialog.setTextValue(default_text)
        dialog.setOkButtonText("确定")
        dialog.setCancelButtonText("取消")
        dialog.setMinimumWidth(440)
        dialog.resize(440, dialog.sizeHint().height())

        line_edit = dialog.findChild(QLineEdit)
        if line_edit is not None:
            QTimer.singleShot(0, line_edit.selectAll)

        accepted = dialog.exec() == QDialog.Accepted
        return dialog.textValue(), accepted

    # ── 格式管理操作 ──

    def _rename_format(self):
        from src.scene.manager import rename_scene
        if not self._scene_path:
            return
        from pathlib import Path
        path = Path(self._scene_path)
        if not path.exists():
            QMessageBox.warning(self, "重命名", "场景文件不存在，无法重命名。")
            return
        if is_protected_scene_path(path):
            QMessageBox.warning(self, "重命名", "默认场景模板不能重命名。")
            return
        old_name = getattr(self._config, "name", "") or path.stem
        new_name, ok = self._prompt_format_name("重命名格式", "新名称:", old_name)
        if ok and new_name.strip():
            new_name = new_name.strip()
            try:
                new_path = rename_scene(path, new_name)
                self._scene_path = str(new_path)
                self._config.name = new_name
                self._fmt_name_label.setText(new_name)
                QMessageBox.information(
                    self, "重命名",
                    f"已重命名为「{new_name}」\n文件: {new_path.name}")
            except Exception as e:
                QMessageBox.critical(self, "重命名失败", str(e))

    def _delete_format(self):
        from src.scene.manager import delete_scene
        if not self._scene_path:
            return
        from pathlib import Path
        path = Path(self._scene_path)
        if not path.exists():
            QMessageBox.warning(self, "删除", "场景文件不存在。")
            return
        if is_protected_scene_path(path):
            QMessageBox.warning(self, "删除", "默认场景模板不能删除。")
            return
        config_name = getattr(self._config, "name", "") or path.stem
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定要删除场景「{config_name}」吗？\n文件: {path.name}\n\n此操作不可撤销。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        try:
            delete_scene(path)
            self._deleted_scene_path = str(path)
            self._scene_path = None
            self._rename_btn.setEnabled(False)
            self._save_btn.setEnabled(False)
            self._delete_btn.setEnabled(False)
            self.reject()  # 关闭对话框
        except Exception as e:
            QMessageBox.critical(self, "删除失败", str(e))

    @staticmethod
    def _commit_scene_config(target: SceneConfig, source: SceneConfig) -> None:
        """Copy a prepared config back into the live shared SceneConfig object."""
        target.__dict__.clear()
        target.__dict__.update(copy.deepcopy(source.__dict__))

    def _build_config_draft_from_ui(self) -> SceneConfig | None:
        """Collect current dialog values into a temporary config draft.

        This keeps the live shared config unchanged until the user truly commits
        the operation (应用修改 / 增量保存成功).
        """
        draft = copy.deepcopy(self._config)
        if not self._sync_config_from_ui(target_config=draft):
            return None
        return draft

    def _save_format(self):
        from src.scene.manager import save_scene, _safe_filename
        from pathlib import Path
        if not self._scene_path:
            QMessageBox.warning(self, "增量保存", "无关联的场景文件，无法保存。")
            return
        src_path = Path(self._scene_path)
        draft = self._build_config_draft_from_ui()
        if draft is None:
            return
        # 生成默认名称
        old_config_name = getattr(draft, "name", "") or src_path.stem
        default_name = f"{old_config_name} v2"
        # 弹窗让用户确认或修改名称
        new_name, ok = self._prompt_format_name("增量保存", "新格式名称:", default_name)
        if not ok:
            return  # 用户取消
        new_name = new_name.strip() or default_name
        draft.name = new_name
        # 增量保存视为新格式文件，永久署名需要重新写入。
        draft.format_signature = ""
        # 文件名由 name 派生
        base_dir = src_path.parent
        safe_stem = _safe_filename(new_name)
        new_path = base_dir / f"{safe_stem}.json"
        # 避免覆盖已有文件
        if new_path.exists():
            idx = 1
            while True:
                candidate = base_dir / f"{safe_stem}_{idx}.json"
                if not candidate.exists():
                    new_path = candidate
                    break
                idx += 1
        try:
            save_scene(draft, new_path)
            self._commit_scene_config(self._config, draft)
            self._scene_path = str(new_path)
            self._fmt_name_label.setText(self._config.name)
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "增量保存失败", str(e))

    # ── 样式表格 ──

    def _build_styles_tab(self) -> QWidget:
        from PySide6.QtGui import QColor, QBrush
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self._style_table = QTableWidget()
        self._style_table.setObjectName("StyleConfigTable")
        cols = _STYLE_COLUMNS
        style_keys = list(self._config.styles.keys())
        backfilled = getattr(self._config, "_backfilled_styles", set())

        self._style_table.setShowGrid(True)
        self._style_table.setWordWrap(False)
        self._style_table.setAlternatingRowColors(False)
        self._style_table.setCornerButtonEnabled(False)
        self._style_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self._style_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self._style_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self._style_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)

        self._style_table.setColumnCount(len(cols))
        self._style_table.setRowCount(len(style_keys))
        self._style_table.setHorizontalHeaderLabels([c[1] for c in cols])

        horizontal_header = self._style_table.horizontalHeader()
        horizontal_header.setSectionResizeMode(QHeaderView.Fixed)
        horizontal_header.setStretchLastSection(False)
        horizontal_header.setDefaultAlignment(Qt.AlignCenter)
        horizontal_header.setHighlightSections(False)

        vertical_header = self._style_table.verticalHeader()
        vertical_header.setSectionResizeMode(QHeaderView.Fixed)
        vertical_header.setDefaultSectionSize(self._STYLE_TABLE_ROW_HEIGHT)
        vertical_header.setMinimumSectionSize(self._STYLE_TABLE_ROW_HEIGHT)
        vertical_header.setDefaultAlignment(Qt.AlignCenter)

        gray_brush = QBrush(QColor(180, 180, 180))
        for row, key in enumerate(style_keys):
            sc = self._config.styles[key]
            sync_style_config_indent_fields(sc)
            is_placeholder = key in backfilled
            header_item = QTableWidgetItem(_STYLE_DISPLAY_NAMES.get(key, key))
            if is_placeholder:
                header_item.setForeground(gray_brush)
            self._style_table.setVerticalHeaderItem(row, header_item)
            for col, (attr, _, typ) in enumerate(cols):
                val = getattr(sc, attr, "")
                table_col = self._style_table_col_index(attr)
                if attr in {"font_cn", "font_en"}:
                    combo = self._build_font_family_combo(val, is_placeholder=is_placeholder)
                    combo.currentTextChanged.connect(
                        lambda _text, row=row: self._update_style_row_header_visual(row)
                    )
                    self._style_table.setCellWidget(row, table_col, self._wrap_style_cell_widget(combo))
                    continue
                if attr == "alignment":
                    combo = self._build_alignment_combo(val, is_placeholder=is_placeholder)
                    combo.currentIndexChanged.connect(
                        lambda _index, row=row: self._update_style_row_header_visual(row)
                    )
                    self._style_table.setCellWidget(row, table_col, self._wrap_style_cell_widget(combo))
                    continue
                if attr == "line_spacing_pt":
                    widget = self._build_line_spacing_value_widget(
                        getattr(sc, "line_spacing_type", "exact"),
                        val,
                        is_placeholder=is_placeholder,
                    )
                    if isinstance(widget, LineSpacingValueWidget):
                        widget.valueChanged.connect(
                            lambda _text, row=row: self._update_style_row_header_visual(row)
                        )
                        widget.spacingTypeChanged.connect(
                            lambda _kind, row=row: self._update_style_row_header_visual(row)
                        )
                    self._style_table.setCellWidget(row, table_col, self._wrap_style_cell_widget(widget))
                    continue
                if attr == "size_pt":
                    combo = self._build_font_size_combo(
                        self._style_size_display_text(sc),
                        is_placeholder=is_placeholder,
                    )
                    combo.currentTextChanged.connect(
                        lambda _text, row=row: (
                            self._update_indent_value_widgets(row),
                            self._update_style_row_header_visual(row),
                        )
                    )
                    self._style_table.setCellWidget(row, table_col, self._wrap_style_cell_widget(combo))
                    continue
                if attr == "special_indent_value":
                    widget = self._build_special_indent_widget(
                        getattr(sc, "special_indent_mode", "none"),
                        getattr(sc, "special_indent_value", 0.0),
                        getattr(sc, "size_pt", 12),
                        getattr(sc, "special_indent_unit", "chars"),
                        is_placeholder=is_placeholder,
                    )
                    if isinstance(widget, SpecialIndentWidget):
                        widget.modeChanged.connect(
                            lambda _mode, row=row: self._update_style_row_header_visual(row)
                        )
                        widget.valueChanged.connect(
                            lambda _text, row=row: self._update_style_row_header_visual(row)
                        )
                        widget.unitChanged.connect(
                            lambda _unit, row=row: self._update_style_row_header_visual(row)
                        )
                    self._style_table.setCellWidget(row, table_col, self._wrap_style_cell_widget(widget))
                    continue
                if attr in {"left_indent_chars", "right_indent_chars"}:
                    widget = self._build_indent_value_widget(
                        val,
                        getattr(sc, "size_pt", 12),
                        getattr(sc, "left_indent_unit" if attr == "left_indent_chars" else "right_indent_unit", "chars"),
                        is_placeholder=is_placeholder,
                    )
                    if isinstance(widget, IndentValueWidget):
                        widget.valueChanged.connect(
                            lambda _text, row=row: self._update_style_row_header_visual(row)
                        )
                        widget.unitChanged.connect(
                            lambda _unit, row=row: self._update_style_row_header_visual(row)
                        )
                    self._style_table.setCellWidget(row, table_col, self._wrap_style_cell_widget(widget))
                    continue
                if typ == "bool":
                    toggle = self._build_bool_toggle(val, is_placeholder=is_placeholder)
                    toggle.toggled.connect(
                        lambda _checked, row=row: self._update_style_row_header_visual(row)
                    )
                    self._style_table.setCellWidget(
                        row,
                        table_col,
                        self._wrap_style_cell_widget(toggle, center_horizontally=True),
                    )
                    continue
                if typ == "float":
                    text = format_font_size_pt(val)
                else:
                    text = str(val)
                item = QTableWidgetItem(text)
                if is_placeholder:
                    item.setForeground(gray_brush)
                self._style_table.setItem(row, table_col, item)
            self._update_style_row_header_visual(row)

        self._style_hover_popup = StyleHoverPopup(self)
        self._style_hover_popup.hide()
        vertical_header.setMouseTracking(True)
        self._style_header_viewport = vertical_header.viewport()
        self._style_header_viewport.setMouseTracking(True)
        self._style_header_viewport.installEventFilter(self)
        self._style_table.verticalScrollBar().valueChanged.connect(self._hide_style_hover_popup)
        self._style_table.horizontalScrollBar().valueChanged.connect(self._hide_style_hover_popup)
        self._apply_style_table_column_widths()
        QTimer.singleShot(0, self._refresh_style_row_header_visuals)
        self._style_keys = style_keys
        layout.addWidget(self._style_table)
        return w

    @classmethod
    def _style_col_index(cls, attr_name: str) -> int:
        for idx, (attr, _, _) in enumerate(_STYLE_COLUMNS):
            if attr == attr_name:
                return idx
        raise KeyError(f"unknown style column: {attr_name}")

    @classmethod
    def _style_preview_col_index(cls, attr_name: str) -> int:
        raise KeyError(f"unknown style preview column: {attr_name}")

    @classmethod
    def _style_table_col_index(cls, attr_name: str) -> int:
        return cls._style_col_index(attr_name)

    @staticmethod
    def _font_size_display_text(pt_value, display_text: str = "") -> str:
        display_text = str(display_text or "").strip()
        if display_text:
            return font_size_display_text(display_text)
        return font_size_display_text(pt_value)

    @classmethod
    def _style_size_display_text(cls, style_config) -> str:
        return cls._font_size_display_text(
            getattr(style_config, "size_pt", ""),
            getattr(style_config, "size_display", ""),
        )

    def _wrap_style_cell_widget(self, widget: QWidget, *, center_horizontally: bool = False) -> QWidget:
        container = QWidget(self._style_table)
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 4, 0, 4)
        layout.setSpacing(0)
        if center_horizontally:
            layout.addStretch(1)
            layout.addWidget(widget)
            layout.addStretch(1)
        else:
            layout.addWidget(widget)
        container._style_inner_widget = widget
        return container

    def _style_cell_widget(self, row: int, col: int):
        widget = self._style_table.cellWidget(row, col)
        inner = getattr(widget, "_style_inner_widget", None)
        return inner if inner is not None else widget

    def _build_font_size_combo(self, value, *, is_placeholder: bool = False) -> QComboBox:
        combo = FontSizeComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        combo.setMinimumContentsLength(4)
        combo.setMaxVisibleItems(16)
        combo.setFixedHeight(24)
        combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        combo.setToolTip("支持直接输入磅值；展开时会自动定位到当前值或最近的预设字号。")

        line_edit = combo.lineEdit()
        if line_edit is not None:
            line_edit.setPlaceholderText("输入磅值或选择字号")
            if is_placeholder:
                line_edit.setStyleSheet("color: rgb(180, 180, 180);")

        for name, pt in WORD_NAMED_FONT_SIZES:
            combo.addFontSizeItem(name, pt, section=FontSizeComboBox._SECTION_NAMED)
        combo.insertSeparator(combo.count())
        for pt in NUMERIC_FONT_SIZE_OPTIONS:
            combo.addFontSizeItem(font_size_display_text(pt), pt, section=FontSizeComboBox._SECTION_NUMERIC)

        combo.setCurrentText(font_size_display_text(value))
        return combo

    def _build_font_family_combo(self, value, *, is_placeholder: bool = False) -> QComboBox:
        combo = FontFamilyComboBox(
            self._installed_font_families(),
            is_placeholder=is_placeholder,
        )
        combo.setFixedHeight(24)
        combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        combo.setToolTip("下拉显示本机已安装字体；支持输入后模糊搜索，例如 times → Times New Roman。")
        line_edit = combo.lineEdit()
        if line_edit is not None:
            line_edit.setPlaceholderText("输入字体名或下拉选择")
        combo.setCurrentText(str(value or ""))
        combo.reveal_text_start()
        return combo

    def _build_alignment_combo(self, value, *, is_placeholder: bool = False) -> QComboBox:
        combo = ScrollSafeComboBox()
        combo.setEditable(False)
        combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        combo.setMinimumContentsLength(6)
        combo.setMaxVisibleItems(8)
        combo.setFixedHeight(24)
        combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        combo.setToolTip("对齐支持左对齐、右对齐、居中对齐、两端对齐。")
        for internal_value, label in ALIGNMENT_OPTIONS:
            combo.addItem(label, internal_value)
        current_value = normalize_alignment_value(value)
        current_index = combo.findData(current_value)
        if current_index < 0:
            current_index = combo.findData("justify")
        combo.setCurrentIndex(max(current_index, 0))
        if is_placeholder:
            combo.setStyleSheet("color: rgb(180, 180, 180);")
        return combo

    def _build_choice_combo(
            self,
            choices: list[tuple[str, str]],
            value,
            *,
            is_placeholder: bool = False,
            tooltip: str = "") -> QComboBox:
        combo = ScrollSafeComboBox()
        combo.setEditable(False)
        combo.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        combo.setMinimumContentsLength(6)
        combo.setMaxVisibleItems(max(6, len(choices)))
        combo.setFixedHeight(24)
        combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        if tooltip:
            combo.setToolTip(tooltip)
        for internal_value, label in choices:
            combo.addItem(label, internal_value)
        current_value = str(value or "").strip().lower()
        current_index = combo.findData(current_value)
        if current_index < 0:
            current_index = 0
        combo.setCurrentIndex(current_index)
        if is_placeholder:
            combo.setStyleSheet("color: rgb(180, 180, 180);")
        return combo

    def _build_section_divider(self, title: str) -> QWidget:
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 10, 0, 2)
        layout.setSpacing(10)

        label = QLabel(title)
        label.setStyleSheet("font-weight: 600; color: #555555;")

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Plain)
        line.setStyleSheet("color: #e1e4e8; background-color: #e1e4e8;")

        layout.addWidget(label)
        layout.addWidget(line, 1)
        return container

    def _build_compact_option_row(
        self,
        *items: tuple[str, QWidget],
        add_stretch: bool = True,
        label_min_width: int = 96,
    ) -> QWidget:
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        expanding_policies = {
            QSizePolicy.Expanding,
            QSizePolicy.MinimumExpanding,
            QSizePolicy.Ignored,
        }
        for label_text, widget in items:
            label = QLabel(label_text)
            label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            if label_min_width > 0:
                label.setMinimumWidth(label_min_width)
            layout.addWidget(label)
            stretch = 1 if widget.sizePolicy().horizontalPolicy() in expanding_policies else 0
            layout.addWidget(widget, stretch)

        if add_stretch:
            layout.addStretch(1)
        return row

    def _build_line_spacing_value_widget(self, spacing_type, value, *, is_placeholder: bool = False) -> QWidget:
        widget = LineSpacingValueWidget(
            spacing_type,
            value,
            self._style_table,
            is_placeholder=is_placeholder,
        )
        return widget

    def _build_indent_value_widget(self, value, size_pt, unit, *, is_placeholder: bool = False) -> QWidget:
        return IndentValueWidget(
            value,
            size_pt,
            self._style_table,
            unit=unit,
            is_placeholder=is_placeholder,
        )

    def _build_special_indent_widget(
        self,
        mode,
        value,
        size_pt,
        unit,
        *,
        is_placeholder: bool = False,
    ) -> QWidget:
        return SpecialIndentWidget(
            mode,
            value,
            size_pt,
            self._style_table,
            unit=unit,
            is_placeholder=is_placeholder,
        )

    @classmethod
    def _installed_font_families(cls) -> list[str]:
        if cls._INSTALLED_FONT_FAMILIES_CACHE is not None:
            return cls._INSTALLED_FONT_FAMILIES_CACHE
        from PySide6.QtGui import QFontDatabase

        families = sorted(
            {str(name).strip() for name in QFontDatabase.families() if str(name).strip()},
            key=lambda item: item.casefold(),
        )
        cls._INSTALLED_FONT_FAMILIES_CACHE = families
        return families

    def _build_bool_toggle(self, value, *, is_placeholder: bool = False) -> QPushButton:
        toggle = OnOffToggleButton(
            self._coerce_bool_value(value),
            is_placeholder=is_placeholder,
        )
        toggle.setToolTip("点击切换开 / 关。")
        return toggle

    def _apply_style_table_column_widths(self) -> None:
        from PySide6.QtGui import QFontMetrics

        header = self._style_table.horizontalHeader()
        metrics = QFontMetrics(header.font())
        header.setSectionResizeMode(QHeaderView.Fixed)
        header.setMinimumSectionSize(28)

        for attr, label, _ in _STYLE_COLUMNS:
            col = self._style_table_col_index(attr)
            width = self._STYLE_TABLE_COLUMN_WIDTHS.get(attr)
            if width is None:
                width = metrics.horizontalAdvance(label) + 22
            self._style_table.setColumnWidth(col, width)

        vertical_header = self._style_table.verticalHeader()
        row_metrics = QFontMetrics(vertical_header.font())
        row_header_width = self._STYLE_TABLE_MIN_ROW_HEADER_WIDTH
        for row in range(self._style_table.rowCount()):
            item = self._style_table.verticalHeaderItem(row)
            if item is None:
                continue
            row_header_width = max(
                row_header_width,
                row_metrics.horizontalAdvance(item.text()) + 24,
            )
            self._style_table.setRowHeight(row, self._STYLE_TABLE_ROW_HEIGHT)
        row_header_width = min(row_header_width, self._STYLE_TABLE_MAX_ROW_HEADER_WIDTH)
        vertical_header.setMinimumWidth(row_header_width)
        vertical_header.setMaximumWidth(row_header_width)

    def _coerce_bool_value(self, value) -> bool:
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            try:
                return self._parse_bool_text(value)
            except ValueError:
                return False
        return bool(value)

    def _style_bool_value(self, row: int, attr_name: str) -> bool:
        col = self._style_table_col_index(attr_name)
        widget = self._style_cell_widget(row, col)
        if isinstance(widget, OnOffToggleButton):
            return widget.isChecked()
        item = self._style_table.item(row, col)
        if not item:
            return False
        return self._coerce_bool_value(item.text().strip())

    def _style_text_value(self, row: int, attr_name: str) -> str:
        col = self._style_table_col_index(attr_name)
        widget = self._style_cell_widget(row, col)
        if isinstance(widget, QComboBox):
            return widget.currentText().strip()
        if isinstance(widget, SpecialIndentWidget):
            return widget.text()
        if isinstance(widget, IndentValueWidget):
            return widget.text()
        if isinstance(widget, LineSpacingValueWidget):
            return widget.text()
        item = self._style_table.item(row, col)
        if not item:
            return ""
        return item.text().strip()

    def _update_indent_value_widgets(self, row: int) -> None:
        size_pt = self._style_size_value(row)
        for attr_name in ("special_indent_value", "left_indent_chars", "right_indent_chars"):
            col = self._style_table_col_index(attr_name)
            widget = self._style_cell_widget(row, col)
            if isinstance(widget, SpecialIndentWidget):
                widget.setReferenceSize(size_pt)
                continue
            if isinstance(widget, IndentValueWidget):
                widget.setReferenceSize(size_pt)

    def _style_line_spacing_state(self, row: int) -> tuple[str, float]:
        value_col = self._style_table_col_index("line_spacing_pt")
        value_widget = self._style_cell_widget(row, value_col)
        if isinstance(value_widget, LineSpacingValueWidget):
            return (
                value_widget.configSpacingType(use_fallback=True),
                value_widget.displayValue(),
            )

        fallback_type = "exact"
        fallback_value = 20.0
        if 0 <= row < len(getattr(self, "_style_keys", [])):
            sc = self._config.styles[self._style_keys[row]]
            fallback_type = normalize_line_spacing_type(getattr(sc, "line_spacing_type", "exact"))
            fallback_value = resolve_line_spacing_value(
                fallback_type,
                getattr(sc, "line_spacing_pt", 20.0),
            )

        raw_value = self._style_text_value(row, "line_spacing_pt")
        return fallback_type, resolve_line_spacing_value(fallback_type, raw_value)

    def _update_style_row_header_visual(self, row: int) -> None:
        if not hasattr(self, "_style_table"):
            return
        header_item = self._style_table.verticalHeaderItem(row)
        if header_item is None:
            return
        font = header_item.font()
        font_cn = self._style_text_value(row, "font_cn")
        if font_cn:
            font.setFamily(font_cn)
        font.setBold(self._style_bool_value(row, "bold"))
        font.setItalic(self._style_bool_value(row, "italic"))
        header_item.setFont(font)
        header_item.setToolTip("")
        if self._style_hover_row == row and self._style_hover_popup is not None and self._style_hover_popup.isVisible():
            self._show_style_hover_popup_for_row(row, force=True)

    def _refresh_style_row_header_visuals(self) -> None:
        for row in range(self._style_table.rowCount()):
            self._update_style_row_header_visual(row)

    def _hide_style_hover_popup(self) -> None:
        self._style_hover_row = -1
        popup = getattr(self, "_style_hover_popup", None)
        if popup is not None:
            popup.hide()

    def _show_style_hover_popup_for_row(self, row: int, *, force: bool = False) -> None:
        if row < 0 or not hasattr(self, "_style_table"):
            self._hide_style_hover_popup()
            return

        popup = getattr(self, "_style_hover_popup", None)
        if popup is None:
            return
        if not force and self._style_hover_row == row and popup.isVisible():
            return

        header = self._style_table.verticalHeader()
        section_top = header.sectionViewportPosition(row)
        if section_top < 0:
            self._hide_style_hover_popup()
            return

        popup.setText(self._build_style_row_tooltip(row))
        popup.adjustSize()

        popup_width = popup.width()
        popup_height = popup.height()
        x = header.viewport().width() + 12
        y = section_top + max(0, (header.sectionSize(row) - popup_height) // 2)
        global_pos = header.viewport().mapToGlobal(QPoint(x, y))

        screen = QApplication.screenAt(global_pos)
        if screen is None:
            screen = QApplication.primaryScreen()
        if screen is not None:
            geometry = screen.availableGeometry()
            max_x = geometry.right() - popup_width - 8
            max_y = geometry.bottom() - popup_height - 8
            clamped_x = min(max(global_pos.x(), geometry.left() + 8), max_x)
            clamped_y = min(max(global_pos.y(), geometry.top() + 8), max_y)
            global_pos = QPoint(clamped_x, clamped_y)

        popup.move(global_pos)
        popup.show()
        popup.raise_()
        self._style_hover_row = row

    def eventFilter(self, watched, event):
        if watched is getattr(self, "_style_header_viewport", None):
            event_type = event.type()
            if event_type == QEvent.ToolTip:
                return True
            if event_type == QEvent.MouseMove:
                header = self._style_table.verticalHeader()
                pos = event.position().toPoint() if hasattr(event, "position") else event.pos()
                row = header.logicalIndexAt(pos.y())
                if row >= 0:
                    self._show_style_hover_popup_for_row(row)
                else:
                    self._hide_style_hover_popup()
            elif event_type in {
                QEvent.Leave,
                QEvent.MouseButtonPress,
                QEvent.MouseButtonDblClick,
                QEvent.Wheel,
                QEvent.Hide,
            }:
                self._hide_style_hover_popup()
        return super().eventFilter(watched, event)

    def _build_style_row_tooltip(self, row: int) -> str:
        font_cn = self._style_text_value(row, "font_cn") or "默认"
        font_en = self._style_text_value(row, "font_en") or "默认"
        size_pt = self._style_size_value(row)
        alignment = alignment_display_label(self._style_text_value(row, "alignment"))
        line_spacing_type, line_spacing_value = self._style_line_spacing_state(row)
        return (
            f"中文字体: {font_cn}\n"
            f"英文字体: {font_en}\n"
            f"字号: {format_font_size_pt(size_pt) if size_pt is not None else '未设置'} 磅\n"
            f"加粗: {'开启' if self._style_bool_value(row, 'bold') else '关闭'}\n"
            f"斜体: {'开启' if self._style_bool_value(row, 'italic') else '关闭'}\n"
            f"对齐: {alignment}\n"
            f"行距: {line_spacing_display_label(line_spacing_type)} "
            f"{format_font_size_pt(line_spacing_value)}{line_spacing_unit_label(line_spacing_type)}"
        )

    def _style_size_value(self, row: int) -> float | None:
        text = self._style_text_value(row, "size_pt")
        try:
            value = parse_font_size_input(text)
        except (TypeError, ValueError):
            return None
        return value if value > 0 else None

    # ── 标题编号 ──

    _LEVEL_DISPLAY = {
        "heading1": "一级标题",
        "heading2": "二级标题",
        "heading3": "三级标题",
        "heading4": "四级标题",
        "heading5": "五级标题",
        "heading6": "六级标题",
        "heading7": "七级标题",
        "heading8": "八级标题",
    }
    _LEVEL_PREVIEW_TITLES = {
        "heading1": "示例",
        "heading2": "示例",
        "heading3": "示例",
        "heading4": "示例",
        "heading5": "示例",
        "heading6": "示例",
        "heading7": "示例",
        "heading8": "示例",
    }
    _NUM_COLS = [
        ("preview", "预览", "preview"),
        ("display_shell", "外壳", "choice"),
        ("chain_summary", "前级", "action"),
        ("display_core_style", "本体", "choice"),
        ("reference_core_style", "下引", "choice"),
        ("title_separator", "分隔", "str"),
        ("start_at", "起始", "int"),
        ("restart_on", "遇级", "choice"),
        ("include_in_toc", "目录", "bool"),
        ("enabled", "启用", "bool"),
    ]
    _NUM_COL_TOOLTIPS = {
        "preview": "编号最终显示效果预览",
        "display_shell": "编号外壳，如 第{}章、（{}）",
        "chain_summary": "父级如何参与组合；完整模板见悬浮提示",
        "display_core_style": "当前级编号本体样式",
        "reference_core_style": "被下级引用时采用的样式",
        "title_separator": "编号与标题正文之间的分隔符",
        "start_at": "该级编号起始值",
        "restart_on": "当遇到指定上级标题时，本级编号会从起始值重新开始",
        "include_in_toc": "该级是否进入目录",
        "enabled": "该级编号是否启用",
    }
    _NUM_REFERENCE_SAME_AS_DISPLAY = "__same_as_display__"
    _NUM_TABLE_ROW_HEIGHT = 42
    _NUM_TABLE_CONTROL_HEIGHT = 32
    _NUM_PREVIEW_BASE_WIDTH = 108
    _NUM_TABLE_WIDTH_PADDING = 28
    _NUM_TITLE_SEPARATOR_EXPANDED_WIDTH = 112
    _NUM_TITLE_SEPARATOR_COMPACT_WIDTH = _NUM_TITLE_SEPARATOR_EXPANDED_WIDTH
    _NUM_TITLE_SEPARATOR_COMPACT_CONTROL_WIDTH = 76
    _NUM_START_AT_CONTROL_WIDTH = 62
    _NUM_TABLE_COLUMN_WIDTHS = {
        "display_shell": 96,
        "chain_summary": 88,
        "display_core_style": 72,
        "reference_core_style": 92,
        "title_separator": _NUM_TITLE_SEPARATOR_EXPANDED_WIDTH,
        "start_at": 70,
        "restart_on": 86,
        "include_in_toc": 50,
        "enabled": 50,
    }
    _NUM_SEPARATOR_OPTIONS = [
        ("fullwidth_space", "全角空格(□ / U+3000)", "\u3000"),
        ("halfwidth_space", "半角空格(· / U+0020)", " "),
        ("tab", "制表符(➡ / U+0009)", "\t"),
        ("underscore", "下划线( _ / U+005F )", "_"),
        ("custom", "自定义", None),
    ]
    _NUM_SEPARATOR_COMPACT_LABELS = {
        "fullwidth_space": "□",
        "halfwidth_space": "·",
        "tab": "➡",
        "underscore": "_",
        "custom": "自定义",
    }
    _NUM_SEPARATOR_PRESET_VALUES = {
        key: value
        for key, _, value in _NUM_SEPARATOR_OPTIONS
        if value is not None
    }
    _NUM_SEPARATOR_VALUE_TO_KEY = {
        value: key for key, value in _NUM_SEPARATOR_PRESET_VALUES.items()
    }
    _NUM_CHAIN_TOKEN_RE = re.compile(r"\{(?P<name>[a-zA-Z][a-zA-Z0-9_]*)\}")
    _NUM_SCHEME_PRESETS = {
        "preset_1": {
            "label": "预设1",
            "description": "第一章 / 第一节 / 一、 / （一）",
        },
        "preset_2": {
            "label": "预设2",
            "description": "1 / 1.1 / 1.1.1",
        },
        "preset_3": {
            "label": "预设3",
            "description": "1. / 1.1. / 1.1.1.",
        },
    }

    @classmethod
    def _num_col_index(cls, attr_name: str) -> int:
        for idx, (attr, _, _) in enumerate(cls._NUM_COLS):
            if attr == attr_name:
                return idx
        raise KeyError(f"unknown numbering column: {attr_name}")

    @classmethod
    def _numbering_table_min_width(cls) -> int:
        return cls._NUM_PREVIEW_BASE_WIDTH + sum(cls._NUM_TABLE_COLUMN_WIDTHS.values()) + cls._NUM_TABLE_WIDTH_PADDING

    @staticmethod
    def _set_combo_current_data(combo: QComboBox, value) -> bool:
        idx = combo.findData(value)
        if idx < 0:
            return False
        combo.setCurrentIndex(idx)
        sync_compact = getattr(combo, "_num_sync_compact_text", None)
        if callable(sync_compact):
            sync_compact()
        return True

    @staticmethod
    def _set_whitespace_widget_raw_text(widget: QWidget, raw_text: str) -> None:
        if isinstance(widget, WhitespaceVisibleLineEdit):
            widget.setRawText(raw_text)
            return
        if hasattr(widget, "setRawText"):
            widget.setRawText(raw_text)

    def _numbering_separator_column_width(self) -> int:
        return self._NUM_TITLE_SEPARATOR_EXPANDED_WIDTH

    def _refresh_numbering_separator_column_width(self) -> None:
        if not hasattr(self, "_num_table"):
            return
        col = self._num_col_index("title_separator")
        width = self._numbering_separator_column_width()
        self._num_table.setColumnWidth(col, width)
        self._num_table.setMinimumWidth(self._numbering_table_min_width())

    @staticmethod
    def _whitespace_widget_text(widget: QWidget) -> str:
        if isinstance(widget, WhitespaceVisibleLineEdit):
            return widget.text()
        if hasattr(widget, "text"):
            return str(widget.text() or "")
        return ""

    @staticmethod
    def _numbering_binding_visible_signature(
        binding: HeadingLevelBindingConfig,
    ) -> tuple[str, str, str, str, str, str, str, int, str, bool, bool]:
        return (
            str(binding.display_shell or ""),
            str(binding.display_core_style or ""),
            str(binding.reference_core_style or ""),
            str(binding.chain or ""),
            str(binding.title_separator or ""),
            str(getattr(binding, "ooxml_separator_mode", "inline") or "inline"),
            str(getattr(binding, "ooxml_suff", "nothing") or ""),
            normalize_start_at(getattr(binding, "start_at", 1)),
            str(getattr(binding, "restart_on", "") or ""),
            bool(getattr(binding, "include_in_toc", True)),
            bool(binding.enabled),
        )

    @staticmethod
    def _default_numbering_binding(level_name: str) -> HeadingLevelBindingConfig:
        defaults = HeadingNumberingV2Config().level_bindings
        return copy.deepcopy(defaults.get(level_name, HeadingLevelBindingConfig()))

    def _decimal_numbering_binding(self, level_name: str, *, enabled: bool) -> HeadingLevelBindingConfig:
        binding = self._default_numbering_binding(level_name)
        binding.enabled = bool(enabled)
        binding.display_shell = "plain"
        binding.display_core_style = "arabic"
        binding.reference_core_style = "arabic"
        binding.chain = default_chain_id_for_level(level_name)
        binding.title_separator = "\u3000"
        return binding

    def _numbering_scheme_preset_bindings(self, preset_id: str) -> dict[str, HeadingLevelBindingConfig]:
        bindings = {
            level_name: self._default_numbering_binding(level_name)
            for level_name in self._ALL_HEADING_KEYS
        }
        for level_name in self._ALL_HEADING_KEYS:
            level_idx = heading_level_index(level_name)
            enabled = level_idx <= 4
            if preset_id == "preset_1":
                if level_idx == 1:
                    binding = self._default_numbering_binding(level_name)
                    binding.enabled = True
                    binding.display_shell = "chapter_cn"
                    binding.display_core_style = "cn_lower"
                    binding.reference_core_style = "cn_lower"
                    binding.chain = "current_only"
                    binding.title_separator = "\u3000"
                elif level_idx == 2:
                    binding = self._default_numbering_binding(level_name)
                    binding.enabled = True
                    binding.display_shell = "section_cn"
                    binding.display_core_style = "cn_lower"
                    binding.reference_core_style = "cn_lower"
                    binding.chain = "current_only"
                    binding.title_separator = "\u3000"
                elif level_idx == 3:
                    binding = self._default_numbering_binding(level_name)
                    binding.enabled = True
                    binding.display_shell = "dunhao_cn"
                    binding.display_core_style = "cn_lower"
                    binding.reference_core_style = "cn_lower"
                    binding.chain = "current_only"
                    binding.title_separator = "\u3000"
                elif level_idx == 4:
                    binding = self._default_numbering_binding(level_name)
                    binding.enabled = True
                    binding.display_shell = "paren_cn"
                    binding.display_core_style = "cn_lower"
                    binding.reference_core_style = "cn_lower"
                    binding.chain = "current_only"
                    binding.title_separator = "\u3000"
                else:
                    binding = self._default_numbering_binding(level_name)
                    binding.enabled = False
            elif preset_id == "preset_2":
                binding = self._decimal_numbering_binding(level_name, enabled=enabled)
            elif preset_id == "preset_3":
                binding = self._decimal_numbering_binding(level_name, enabled=enabled)
                binding.display_shell = "dot_suffix"
            else:
                binding = self._default_numbering_binding(level_name)
            bindings[level_name] = binding
        return bindings

    def _detect_matching_numbering_preset_id(
        self,
        bindings: dict[str, HeadingLevelBindingConfig],
    ) -> str | None:
        for preset_id in self._NUM_SCHEME_PRESETS:
            preset_bindings = self._numbering_scheme_preset_bindings(preset_id)
            if all(
                self._numbering_binding_visible_signature(bindings.get(level_name, HeadingLevelBindingConfig()))
                == self._numbering_binding_visible_signature(preset_bindings[level_name])
                for level_name in self._ALL_HEADING_KEYS
            ):
                return preset_id
        return None

    def _set_numbering_table_editable(self, editable: bool) -> None:
        detail_container = getattr(self, "_num_detail_container", None)
        if detail_container is not None:
            detail_container.setEnabled(bool(editable))
            if editable:
                if detail_container.graphicsEffect() is getattr(self, "_num_detail_disabled_effect", None):
                    detail_container.setGraphicsEffect(None)
            else:
                effect = getattr(self, "_num_detail_disabled_effect", None)
                if effect is None:
                    effect = QGraphicsOpacityEffect(detail_container)
                    effect.setOpacity(0.55)
                    self._num_detail_disabled_effect = effect
                detail_container.setGraphicsEffect(effect)
        table = getattr(self, "_num_table", None)
        if table is not None:
            table.setEnabled(bool(editable))
        for widgets in getattr(self, "_num_row_widgets", {}).values():
            for widget in widgets.values():
                widget.setEnabled(bool(editable))

    def _capture_numbering_ui_state(self) -> dict[str, object]:
        state: dict[str, object] = {
            "catalog_source": copy.deepcopy(getattr(self, "_num_catalog_source", None)),
            "rows": {},
        }
        row_states: dict[str, dict[str, object]] = {}
        for row, level_name in enumerate(getattr(self, "_num_level_keys", [])):
            shell_combo = self._num_cell_widget(row, "display_shell")
            chain_button = self._num_cell_widget(row, "chain_summary")
            core_style_combo = self._num_cell_widget(row, "display_core_style")
            ref_combo = self._num_cell_widget(row, "reference_core_style")
            separator_edit = self._num_cell_widget(row, "title_separator")
            start_spin = self._num_cell_widget(row, "start_at")
            restart_combo = self._num_cell_widget(row, "restart_on")
            toc_toggle = self._num_cell_widget(row, "include_in_toc")
            enabled_toggle = self._num_cell_widget(row, "enabled")
            row_states[level_name] = {
                "display_shell_text": shell_combo.currentText() if isinstance(shell_combo, QComboBox) else "",
                "chain_id": (
                    str(getattr(chain_button, "_num_chain_id", "current_only") or "current_only")
                    if isinstance(chain_button, QPushButton)
                    else "current_only"
                ),
                "display_core_style": (
                    str(core_style_combo.currentData() or "")
                    if isinstance(core_style_combo, QComboBox)
                    else ""
                ),
                "reference_core_style": (
                    str(ref_combo.currentData() or self._NUM_REFERENCE_SAME_AS_DISPLAY)
                    if isinstance(ref_combo, QComboBox)
                    else self._NUM_REFERENCE_SAME_AS_DISPLAY
                ),
                "title_separator": self._whitespace_widget_text(separator_edit),
                "start_at": int(start_spin.value()) if isinstance(start_spin, QSpinBox) else 1,
                "restart_on": (
                    str(restart_combo.currentData() or "")
                    if isinstance(restart_combo, QComboBox)
                    else ""
                ),
                "include_in_toc": bool(toc_toggle.isChecked()) if isinstance(toc_toggle, OnOffToggleButton) else False,
                "enabled": bool(enabled_toggle.isChecked()) if isinstance(enabled_toggle, OnOffToggleButton) else False,
            }
        state["rows"] = row_states
        return state

    def _restore_numbering_ui_state(self, state: dict[str, object] | None) -> None:
        if not isinstance(state, dict):
            return
        catalog_source = state.get("catalog_source")
        if catalog_source is not None:
            self._num_catalog_source = copy.deepcopy(catalog_source)
        rows = state.get("rows", {})
        if not isinstance(rows, dict):
            return

        for row, level_name in enumerate(getattr(self, "_num_level_keys", [])):
            row_state = rows.get(level_name)
            if not isinstance(row_state, dict):
                continue
            shell_combo = self._num_cell_widget(row, "display_shell")
            chain_button = self._num_cell_widget(row, "chain_summary")
            core_style_combo = self._num_cell_widget(row, "display_core_style")
            ref_combo = self._num_cell_widget(row, "reference_core_style")
            separator_edit = self._num_cell_widget(row, "title_separator")
            start_spin = self._num_cell_widget(row, "start_at")
            restart_combo = self._num_cell_widget(row, "restart_on")
            toc_toggle = self._num_cell_widget(row, "include_in_toc")
            enabled_toggle = self._num_cell_widget(row, "enabled")

            if isinstance(shell_combo, QComboBox):
                shell_combo.setCurrentText(str(row_state.get("display_shell_text", "") or ""))
            if isinstance(chain_button, QPushButton):
                self._set_chain_button_state(
                    chain_button,
                    level_name,
                    str(row_state.get("chain_id", "current_only") or "current_only"),
                )
            if isinstance(core_style_combo, QComboBox):
                self._set_combo_current_data(core_style_combo, row_state.get("display_core_style"))
            if isinstance(ref_combo, QComboBox):
                self._set_combo_current_data(ref_combo, row_state.get("reference_core_style"))
            self._set_whitespace_widget_raw_text(
                separator_edit,
                str(row_state.get("title_separator", "") or ""),
            )
            if isinstance(start_spin, QSpinBox):
                start_spin.setValue(max(1, int(row_state.get("start_at", 1) or 1)))
            if isinstance(restart_combo, QComboBox):
                self._set_combo_current_data(restart_combo, row_state.get("restart_on"))
            if isinstance(toc_toggle, OnOffToggleButton):
                toc_toggle.setChecked(bool(row_state.get("include_in_toc", True)))
            if isinstance(enabled_toggle, OnOffToggleButton):
                enabled_toggle.setChecked(bool(row_state.get("enabled", False)))
        self._refresh_numbering_separator_column_width()
        self._refresh_numbering_previews()

    def _update_numbering_preset_description(self) -> None:
        label = getattr(self, "_num_preset_desc_label", None)
        combo = getattr(self, "_num_mode_combo", None)
        if label is None or combo is None:
            return
        preset_id = str(combo.currentData() or "")
        meta = self._NUM_SCHEME_PRESETS.get(preset_id, {})
        description = str(meta.get("description", "") or "").strip()
        if description:
            label.setText(f"当前方案：{description}（预设模式，下方明细不可编辑）")
            return
        label.setText("当前方案：自定义。下方各列可逐级调整。")

    def _apply_numbering_scheme_preset_to_ui(self, preset_id: str) -> None:
        preset_bindings = self._numbering_scheme_preset_bindings(preset_id)
        for row, level_name in enumerate(getattr(self, "_num_level_keys", [])):
            binding = preset_bindings.get(level_name)
            if binding is None:
                continue
            shell_combo = self._num_cell_widget(row, "display_shell")
            chain_button = self._num_cell_widget(row, "chain_summary")
            core_style_combo = self._num_cell_widget(row, "display_core_style")
            ref_combo = self._num_cell_widget(row, "reference_core_style")
            separator_edit = self._num_cell_widget(row, "title_separator")
            start_spin = self._num_cell_widget(row, "start_at")
            restart_combo = self._num_cell_widget(row, "restart_on")
            toc_toggle = self._num_cell_widget(row, "include_in_toc")
            enabled_toggle = self._num_cell_widget(row, "enabled")

            if isinstance(shell_combo, QComboBox):
                self._set_combo_current_data(shell_combo, binding.display_shell)
            if isinstance(chain_button, QPushButton):
                self._set_chain_button_state(chain_button, level_name, binding.chain)
            if isinstance(core_style_combo, QComboBox):
                self._set_combo_current_data(core_style_combo, binding.display_core_style)
            if isinstance(ref_combo, QComboBox):
                ref_value = (
                    self._NUM_REFERENCE_SAME_AS_DISPLAY
                    if binding.reference_core_style == binding.display_core_style
                    else binding.reference_core_style
                )
                self._set_combo_current_data(ref_combo, ref_value)
            self._set_whitespace_widget_raw_text(separator_edit, binding.title_separator)
            if isinstance(start_spin, QSpinBox):
                start_spin.setValue(normalize_start_at(getattr(binding, "start_at", 1)))
            if isinstance(restart_combo, QComboBox):
                self._set_combo_current_data(
                    restart_combo,
                    str(getattr(binding, "restart_on", "") or ""),
                )
            if isinstance(toc_toggle, OnOffToggleButton):
                toc_toggle.setChecked(bool(getattr(binding, "include_in_toc", True)))
            if isinstance(enabled_toggle, OnOffToggleButton):
                enabled_toggle.setChecked(binding.enabled)
        self._refresh_numbering_separator_column_width()
        self._refresh_numbering_previews()

    def _sync_numbering_mode_controls(self, *, apply_selected_preset: bool = False) -> None:
        mode_combo = getattr(self, "_num_mode_combo", None)
        if mode_combo is None:
            return

        previous_mode = str(getattr(self, "_num_active_mode", "") or "")
        mode = str(mode_combo.currentData() or "custom")
        if previous_mode == "custom" and mode != "custom":
            self._num_custom_ui_state = self._capture_numbering_ui_state()
        self._update_numbering_preset_description()
        if mode != "custom":
            if apply_selected_preset and mode in self._NUM_SCHEME_PRESETS:
                self._apply_numbering_scheme_preset_to_ui(mode)
            self._set_numbering_table_editable(False)
            self._num_active_mode = mode
            return

        if previous_mode and previous_mode != "custom":
            self._restore_numbering_ui_state(getattr(self, "_num_custom_ui_state", None))
        self._set_numbering_table_editable(True)
        self._num_active_mode = mode

    def _numbering_v2_display_config(self, *, config: SceneConfig | None = None):
        cfg = config or self._config
        source = str(getattr(cfg, "_heading_numbering_v2_source", "") or "").strip().lower()
        scheme_id = str(
            getattr(self, "_numbering_scheme_id", "") or self._preferred_numbering_scheme_id()
        ).strip()
        if source != "payload" and scheme_id and scheme_id != "__levels__":
            levels = self._resolve_numbering_levels(scheme_id, config=cfg)
            if levels:
                legacy_copy = copy.deepcopy(cfg.heading_numbering)
                legacy_copy.levels = copy.deepcopy(levels)
                return _derive_heading_numbering_v2_from_legacy(
                    legacy_copy,
                    enabled=bool(cfg.capabilities.get("heading_numbering", True)),
                )
        return cfg.heading_numbering_v2

    def _build_numbering_choice_combo(
        self,
        options: list[tuple[str, str]],
        current_key: str,
        *,
        editable: bool = False,
        text_alignment=Qt.AlignLeft | Qt.AlignVCenter,
        compact_current_labels: dict[str, str] | None = None,
    ) -> QComboBox:
        combo = ScrollSafeComboBox()
        combo.setFixedHeight(self._NUM_TABLE_CONTROL_HEIGHT)
        combo.setMinimumWidth(0)
        combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        use_compact_current = bool(compact_current_labels) and not editable
        combo.setEditable(editable or use_compact_current)
        if editable or use_compact_current:
            combo.setInsertPolicy(QComboBox.NoInsert)
        line_edit = combo.lineEdit()
        if line_edit is not None:
            line_edit.setAlignment(text_alignment)
            line_edit.setMinimumWidth(0)
            if use_compact_current:
                line_edit.setReadOnly(True)
        for key, label in options:
            combo.addItem(label, key)
            combo.setItemData(combo.count() - 1, int(text_alignment), Qt.TextAlignmentRole)
            combo.setItemData(combo.count() - 1, label, Qt.ToolTipRole)
        idx = combo.findData(current_key)
        if idx < 0 and options:
            idx = 0
        combo.setCurrentIndex(max(idx, 0))
        if options:
            longest = max((len(label) for _, label in options), default=0)
            combo.view().setMinimumWidth(max(combo.view().minimumWidth(), 16 * longest))
        if use_compact_current and line_edit is not None:
            compact_map = {str(key): str(label) for key, label in (compact_current_labels or {}).items()}

            def _sync_compact_text(*_args) -> None:
                current_data = str(combo.currentData() or "")
                display_text = compact_map.get(current_data, combo.itemText(combo.currentIndex()))
                blocked = line_edit.blockSignals(True)
                line_edit.setText(display_text)
                line_edit.blockSignals(blocked)

            combo.currentIndexChanged.connect(_sync_compact_text)
            combo.activated.connect(lambda *_args: QTimer.singleShot(0, _sync_compact_text))
            combo._num_sync_compact_text = _sync_compact_text
            _sync_compact_text()
        combo.setStyleSheet(
            """
            QComboBox {
                padding: 0 14px 0 3px;
            }
            QComboBox::drop-down {
                width: 14px;
            }
            """
        )
        return combo

    def _wrap_num_cell_widget(self, widget: QWidget, *, center_horizontally: bool = False) -> QWidget:
        if not center_horizontally:
            policy = widget.sizePolicy()
            policy.setHorizontalPolicy(QSizePolicy.Expanding)
            widget.setSizePolicy(policy)
            widget.setMinimumWidth(0)
        return NumberingCellWidgetContainer(
            widget,
            center_horizontally=center_horizontally,
            parent=self._num_table,
        )

    def _num_cell_widget(self, row: int, col_or_attr):
        col = col_or_attr
        if isinstance(col_or_attr, str):
            col = self._num_col_index(col_or_attr)
        widget = self._num_table.cellWidget(row, col)
        inner = getattr(widget, "_num_inner_widget", None)
        return inner if inner is not None else widget

    @staticmethod
    def _numbering_choice_label(text: str, *, fallback: str) -> str:
        clean = str(text or "").strip()
        return clean or fallback

    @staticmethod
    def _allocate_numbering_catalog_id(prefix: str, level_name: str, catalog: dict[str, object]) -> str:
        base = f"{prefix}_{str(level_name or '').strip().lower() or 'heading'}"
        candidate = base
        idx = 1
        while candidate in catalog:
            candidate = f"{base}_{idx}"
            idx += 1
        return candidate

    def _numbering_shell_options(self) -> list[tuple[str, str]]:
        options = []
        for key, shell in getattr(self._num_catalog_source, "shell_catalog", {}).items():
            label = self._numbering_choice_label(
                getattr(shell, "label", ""),
                fallback=f"{getattr(shell, 'prefix', '')}{{}}{getattr(shell, 'suffix', '')}",
            )
            options.append((str(key), label))
        return options

    def _numbering_core_style_options(self) -> list[tuple[str, str]]:
        options = []
        for key, style in getattr(self._num_catalog_source, "core_style_catalog", {}).items():
            label = self._numbering_choice_label(
                getattr(style, "label", ""),
                fallback=getattr(style, "sample", "") or str(key),
            )
            options.append((str(key), label))
        return options

    def _numbering_core_style_compact_labels(self) -> dict[str, str]:
        compact: dict[str, str] = {}
        for key, style in getattr(self._num_catalog_source, "core_style_catalog", {}).items():
            compact[str(key)] = self._numbering_core_style_sample(str(key))
        return compact

    def _numbering_reference_options(self) -> list[tuple[str, str]]:
        options = [(self._NUM_REFERENCE_SAME_AS_DISPLAY, "沿用本体")]
        options.extend(self._numbering_core_style_options())
        return options

    def _numbering_reference_compact_labels(self) -> dict[str, str]:
        compact = {self._NUM_REFERENCE_SAME_AS_DISPLAY: "沿用"}
        compact.update(self._numbering_core_style_compact_labels())
        return compact

    def _numbering_restart_options(self, level_name: str) -> list[tuple[str, str]]:
        level_idx = heading_level_index(level_name)
        options = [("", "无")]
        for idx in range(1, level_idx):
            short_label = self._LEVEL_DISPLAY.get(f"heading{idx}", f"{idx}级").replace("标题", "")
            options.append((f"heading{idx}", short_label))
        return options

    def _numbering_core_style_sample(self, core_style_id: str, *, catalog_source=None) -> str:
        source_v2 = self._num_catalog_source if catalog_source is None else catalog_source
        style = getattr(source_v2, "core_style_catalog", {}).get(core_style_id)
        if style is None:
            return "1"
        sample = str(getattr(style, "sample", "") or "").strip()
        if sample:
            return sample
        label = str(getattr(style, "label", "") or "").strip()
        if label:
            return label.split()[0]
        return "1"

    def _numbering_core_style_preview_value(
        self,
        core_style_id: str,
        value: int,
        *,
        catalog_source=None,
    ) -> str:
        preview = format_core_style_value(core_style_id, value)
        if preview:
            return preview
        return self._numbering_core_style_sample(core_style_id, catalog_source=catalog_source)

    def _preview_counters_for_level(
        self,
        level_name: str,
        bindings: dict[str, HeadingLevelBindingConfig],
    ) -> dict[str, int]:
        counters = {f"heading{idx}": 0 for idx in range(1, 9)}
        target_idx = heading_level_index(level_name)
        for idx in range(1, target_idx + 1):
            counters = advance_heading_counters(f"heading{idx}", counters, bindings)
        return counters

    def _numbering_chain_text(self, chain: NumberChainConfig | None) -> str:
        if chain is None:
            return "{current}"
        parts: list[str] = []
        for segment in getattr(chain, "segments", None) or []:
            if str(getattr(segment, "type", "") or "") == "literal":
                parts.append(str(getattr(segment, "text", "") or ""))
                continue
            source = str(getattr(segment, "source", "") or "")
            if source == "current":
                parts.append("{current}")
            elif source.startswith("level"):
                parts.append(f"{{{source}}}")
        return "".join(parts) or "{current}"

    def _find_shell_id_by_text(self, shell_text: str, *, catalog_source=None) -> str | None:
        normalized = str(shell_text or "").strip()
        source_v2 = self._num_catalog_source if catalog_source is None else catalog_source
        for shell_id, shell in getattr(source_v2, "shell_catalog", {}).items():
            if self._numbering_choice_label(
                getattr(shell, "label", ""),
                fallback=f"{getattr(shell, 'prefix', '')}{{}}{getattr(shell, 'suffix', '')}",
            ) == normalized:
                return str(shell_id)
        return None

    def _find_chain_id_by_text(self, chain_text: str, *, catalog_source=None) -> str | None:
        normalized = str(chain_text or "").strip()
        source_v2 = self._num_catalog_source if catalog_source is None else catalog_source
        for chain_id, chain in getattr(source_v2, "chain_catalog", {}).items():
            if self._numbering_chain_text(chain) == normalized:
                return str(chain_id)
        return None

    def _resolve_or_create_shell_id_from_text(
        self,
        level_name: str,
        shell_text: str,
        *,
        catalog_source=None,
    ) -> str:
        normalized = str(shell_text or "").strip()
        if not normalized:
            raise ValueError("外壳不能为空")
        source_v2 = self._num_catalog_source if catalog_source is None else catalog_source
        existing = self._find_shell_id_by_text(normalized, catalog_source=source_v2)
        if existing:
            return existing
        if normalized.count("{}") != 1:
            raise ValueError("自定义外壳必须且只能包含一个 {} 占位符")
        prefix, suffix = normalized.split("{}", 1)
        catalog = getattr(source_v2, "shell_catalog", {})
        shell_id = self._allocate_numbering_catalog_id("custom_shell", level_name, catalog)
        catalog[shell_id] = NumberShellConfig(label=normalized, prefix=prefix, suffix=suffix)
        return shell_id

    def _chain_segments_from_text(
        self,
        level_name: str,
        chain_text: str,
    ) -> list[NumberChainSegmentConfig]:
        normalized = str(chain_text or "").strip()
        if not normalized:
            raise ValueError("多级组合不能为空")
        errors = validate_heading_numbering_template(level_name, normalized)
        if errors:
            raise ValueError("；".join(errors))

        level_idx = heading_level_index(level_name)
        current_count = 0
        segments: list[NumberChainSegmentConfig] = []
        cursor = 0
        for match in self._NUM_CHAIN_TOKEN_RE.finditer(normalized):
            literal = normalized[cursor:match.start()]
            if literal:
                segments.append(NumberChainSegmentConfig(type="literal", text=literal))

            token_name = str(match.group("name") or "").strip()
            if token_name in {"current", "n", "cn"}:
                segments.append(NumberChainSegmentConfig(type="value", source="current"))
                current_count += 1
            elif token_name == "parent":
                for idx in range(1, level_idx):
                    if segments and str(getattr(segments[-1], "type", "") or "") != "literal":
                        segments.append(NumberChainSegmentConfig(type="literal", text="."))
                    segments.append(NumberChainSegmentConfig(type="value", source=f"level{idx}"))
            else:
                segments.append(NumberChainSegmentConfig(type="value", source=token_name))
            cursor = match.end()

        tail = normalized[cursor:]
        if tail:
            segments.append(NumberChainSegmentConfig(type="literal", text=tail))

        if current_count != 1:
            raise ValueError("多级组合必须且只能包含一个 {current}")
        return segments

    def _resolve_or_create_chain_id_from_text(
        self,
        level_name: str,
        chain_text: str,
        *,
        catalog_source=None,
    ) -> str:
        normalized = str(chain_text or "").strip()
        if not normalized:
            raise ValueError("多级组合不能为空")
        source_v2 = self._num_catalog_source if catalog_source is None else catalog_source
        existing = self._find_chain_id_by_text(normalized, catalog_source=source_v2)
        if existing:
            return existing
        segments = self._chain_segments_from_text(level_name, normalized)
        catalog = getattr(source_v2, "chain_catalog", {})
        chain_id = self._allocate_numbering_catalog_id("custom_chain", level_name, catalog)
        catalog[chain_id] = NumberChainConfig(label=normalized, segments=copy.deepcopy(segments))
        return chain_id

    def _analyze_chain(
        self,
        level_name: str,
        chain_id: str,
        *,
        catalog_source=None,
    ) -> dict[str, object]:
        source_v2 = self._num_catalog_source if catalog_source is None else catalog_source
        chain = getattr(source_v2, "chain_catalog", {}).get(chain_id)
        if chain is None:
            return {
                "mode": "current_only",
                "refs": [],
                "connector": ".",
                "start_idx": None,
                "template_text": "{current}",
            }

        segments = list(getattr(chain, "segments", None) or [])
        template_text = self._numbering_chain_text(chain)

        value_sources: list[str] = []
        literals_between: list[str] = []
        literal_buffer = ""
        leading_literal = ""
        trailing_literal = ""
        saw_value = False

        for segment in segments:
            if str(getattr(segment, "type", "") or "") == "literal":
                if saw_value:
                    literal_buffer += str(getattr(segment, "text", "") or "")
                else:
                    leading_literal += str(getattr(segment, "text", "") or "")
                continue

            saw_value = True
            if value_sources:
                literals_between.append(literal_buffer)
            literal_buffer = ""
            value_sources.append(str(getattr(segment, "source", "") or ""))

        trailing_literal = literal_buffer

        refs: list[int] = []
        simple = not leading_literal and not trailing_literal and bool(value_sources)
        if not value_sources or value_sources[-1] != "current":
            simple = False

        prev_idx = 0
        for source in value_sources[:-1]:
            if not source.startswith("level"):
                simple = False
                continue
            try:
                idx = int(source[5:])
            except ValueError:
                simple = False
                continue
            refs.append(idx)
            if idx <= prev_idx or idx >= heading_level_index(level_name):
                simple = False
            prev_idx = idx

        connector = "."
        if literals_between:
            connector = literals_between[0]
            if any(literal != connector for literal in literals_between):
                simple = False
        elif refs:
            connector = ""

        if not refs:
            connector = ""

        level_idx = heading_level_index(level_name)
        continuous_start = refs[0] if refs and refs == list(range(refs[0], level_idx)) else None
        if simple and not refs:
            mode = "current_only"
        elif simple and continuous_start is not None:
            mode = "continuous"
        else:
            mode = "advanced"

        return {
            "simple": simple,
            "mode": mode,
            "refs": refs,
            "connector": connector,
            "start_idx": continuous_start,
            "template_text": template_text,
        }

    @staticmethod
    def _simple_chain_template(refs: list[int], connector: str) -> str:
        if not refs:
            return "{current}"
        parts = [f"{{level{idx}}}" for idx in refs] + ["{current}"]
        return str(connector).join(parts)

    def _continuous_chain_summary_text(self, start_idx: int, connector: str) -> str:
        start_label = self._LEVEL_DISPLAY.get(f"heading{start_idx}", f"{start_idx}级").replace("标题", "")
        connector_label = connector if connector else "直连"
        return f"{start_label}起[{connector_label}]"

    def _numbering_chain_summary(self, level_name: str, chain_id: str, *, catalog_source=None) -> tuple[str, str]:
        info = self._analyze_chain(level_name, chain_id, catalog_source=catalog_source)
        template_text = str(info.get("template_text", "{current}") or "{current}")
        mode = str(info.get("mode", "advanced") or "advanced")
        if mode == "advanced":
            return "自定义", template_text
        if mode == "current_only":
            return "仅当前", template_text

        start_idx = info.get("start_idx")
        try:
            start_idx = int(start_idx) if start_idx is not None else None
        except (TypeError, ValueError):
            start_idx = None
        if start_idx is None:
            return "自定义", template_text

        connector = str(info.get("connector", "") or "")
        return self._continuous_chain_summary_text(start_idx, connector), template_text

    def _set_chain_button_state(self, button: QPushButton, level_name: str, chain_id: str) -> None:
        summary, template_text = self._numbering_chain_summary(level_name, chain_id)
        button.setText(summary)
        button.setToolTip(f"当前组合：{template_text}\n点击编辑")
        button._num_chain_id = str(chain_id)
        button._num_chain_text = template_text

    def _edit_numbering_chain(self, row: int, level_name: str) -> None:
        button = self._num_cell_widget(row, "chain_summary")
        if not isinstance(button, QPushButton):
            return

        current_chain_id = str(getattr(button, "_num_chain_id", "current_only") or "current_only")
        info = self._analyze_chain(level_name, current_chain_id)
        current_level_idx = heading_level_index(level_name)

        dialog = QDialog(self)
        dialog.setWindowTitle(f"{self._LEVEL_DISPLAY.get(level_name, level_name)} - 前级组合")
        dialog.setMinimumWidth(480)
        layout = QVBoxLayout(dialog)

        note = QLabel(
            "这里单独定义前级引用逻辑。常规情况只选“仅当前”或“从某一级连续到当前”，"
            "只有特殊模板再进入高级模板。"
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        mode_combo = ScrollSafeComboBox()
        mode_options = [("current_only", "仅当前")]
        if current_level_idx > 1:
            mode_options.append(("continuous", "连续引用上级"))
        mode_options.append(("advanced", "高级模板"))
        for key, label in mode_options:
            mode_combo.addItem(label, key)
        initial_mode = str(info.get("mode", "advanced") or "advanced")
        if initial_mode == "continuous" and current_level_idx <= 1:
            initial_mode = "current_only"
        mode_idx = mode_combo.findData(initial_mode)
        mode_combo.setCurrentIndex(max(mode_idx, 0))
        mode_form = QFormLayout()
        mode_form.addRow("组合方式:", mode_combo)
        layout.addLayout(mode_form)

        structured_group = QGroupBox("结构化定义")
        structured_layout = QFormLayout(structured_group)

        start_combo = ScrollSafeComboBox()
        for idx in range(1, current_level_idx):
            start_combo.addItem(
                self._LEVEL_DISPLAY.get(f"heading{idx}", f"{idx}级").replace("标题", ""),
                idx,
            )
        if start_combo.count() == 0:
            start_combo.addItem("无可用上级", None)
        start_idx = info.get("start_idx")
        if start_idx is None and current_level_idx > 1:
            start_idx = 1
        start_pos = start_combo.findData(start_idx)
        if start_pos >= 0:
            start_combo.setCurrentIndex(start_pos)
        structured_layout.addRow("起始上级:", start_combo)

        connector_edit = WhitespaceVisibleLineEdit(str(info.get("connector", ".") or ""))
        connector_edit.setToolTip("例如 .、-、/ 或空。")
        structured_layout.addRow("连接符:", connector_edit)
        layout.addWidget(structured_group)

        advanced_group = QGroupBox("高级模板")
        advanced_layout = QFormLayout(advanced_group)
        advanced_edit = QLineEdit(str(info.get("template_text", "{current}") or "{current}"))
        advanced_edit.setToolTip(
            "支持 {current}、{level1}...{level8}，兼容 {parent} / {n} / {cn}。"
        )
        advanced_layout.addRow("组合模板:", advanced_edit)
        layout.addWidget(advanced_group)

        preview_label = QLabel("")
        preview_label.setWordWrap(True)
        layout.addWidget(preview_label)
        advanced_state = {"user_edited": initial_mode == "advanced"}

        def _sync_advanced_edit(template_text: str) -> None:
            if advanced_state["user_edited"]:
                return
            if advanced_edit.text() == template_text:
                return
            blocked = advanced_edit.blockSignals(True)
            advanced_edit.setText(template_text)
            advanced_edit.blockSignals(blocked)

        def _set_mode_enabled() -> None:
            mode_key = str(mode_combo.currentData() or "current_only")
            structured_group.setEnabled(mode_key != "advanced")
            start_combo.setEnabled(mode_key == "continuous" and start_combo.count() > 0)
            connector_edit.setEnabled(mode_key == "continuous")
            advanced_group.setEnabled(mode_key == "advanced")

        def _update_preview() -> None:
            mode_key = str(mode_combo.currentData() or "current_only")
            if mode_key == "advanced":
                preview_label.setText(f"摘要：自定义\n模板：{advanced_edit.text().strip() or '{current}'}")
                return
            if mode_key == "continuous" and start_combo.count() > 0:
                start_idx_local = int(start_combo.currentData() or 1)
                refs = list(range(start_idx_local, current_level_idx))
                template_text = self._simple_chain_template(refs, connector_edit.text())
                summary = self._continuous_chain_summary_text(start_idx_local, connector_edit.text())
                _sync_advanced_edit(template_text)
                preview_label.setText(f"摘要：{summary}\n模板：{template_text}")
                return
            template_text = "{current}"
            _sync_advanced_edit(template_text)
            preview_label.setText("摘要：仅当前\n模板：{current}")

        mode_combo.currentIndexChanged.connect(_set_mode_enabled)
        mode_combo.currentIndexChanged.connect(_update_preview)
        connector_edit.textChanged.connect(_update_preview)
        start_combo.currentIndexChanged.connect(_update_preview)

        def _on_advanced_text_changed() -> None:
            if str(mode_combo.currentData() or "current_only") == "advanced":
                advanced_state["user_edited"] = True
            _update_preview()

        advanced_edit.textChanged.connect(_on_advanced_text_changed)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        _set_mode_enabled()
        _update_preview()

        if dialog.exec() != QDialog.Accepted:
            return

        try:
            mode_key = str(mode_combo.currentData() or "current_only")
            if mode_key == "advanced":
                chain_text = advanced_edit.text().strip() or "{current}"
            elif mode_key == "continuous" and start_combo.count() > 0:
                start_idx = int(start_combo.currentData() or 1)
                refs = list(range(start_idx, current_level_idx))
                chain_text = self._simple_chain_template(refs, connector_edit.text())
            else:
                chain_text = "{current}"
            chain_id = self._resolve_or_create_chain_id_from_text(level_name, chain_text)
        except ValueError as exc:
            QMessageBox.warning(self, "前级组合", str(exc))
            return

        self._set_chain_button_state(button, level_name, chain_id)
        self._refresh_numbering_previews()

    def _numbering_binding_from_row(
        self,
        row: int,
        level_name: str,
        *,
        error_collector: list[str] | None = None,
        catalog_source=None,
    ) -> HeadingLevelBindingConfig:
        source_v2 = self._num_catalog_source if catalog_source is None else catalog_source
        base_bindings = getattr(source_v2, "level_bindings", {})
        binding = copy.deepcopy(base_bindings.get(level_name, HeadingLevelBindingConfig()))

        shell_combo = self._num_cell_widget(row, "display_shell")
        chain_button = self._num_cell_widget(row, "chain_summary")
        core_style_combo = self._num_cell_widget(row, "display_core_style")
        ref_combo = self._num_cell_widget(row, "reference_core_style")
        separator_edit = self._num_cell_widget(row, "title_separator")
        start_spin = self._num_cell_widget(row, "start_at")
        restart_combo = self._num_cell_widget(row, "restart_on")
        toc_toggle = self._num_cell_widget(row, "include_in_toc")
        enabled_toggle = self._num_cell_widget(row, "enabled")

        if isinstance(shell_combo, QComboBox):
            shell_text = shell_combo.currentText().strip()
            try:
                binding.display_shell = self._resolve_or_create_shell_id_from_text(
                    level_name,
                    shell_text,
                    catalog_source=source_v2,
                )
            except ValueError as exc:
                if error_collector is not None:
                    error_collector.append(
                        f"{self._LEVEL_DISPLAY.get(level_name, level_name)} / 外壳 = {shell_text} ({exc})"
                    )
        if isinstance(chain_button, QPushButton):
            binding.chain = str(getattr(chain_button, "_num_chain_id", binding.chain) or binding.chain)
        if isinstance(core_style_combo, QComboBox):
            binding.display_core_style = str(core_style_combo.currentData() or binding.display_core_style)
        if isinstance(ref_combo, QComboBox):
            ref_key = str(ref_combo.currentData() or self._NUM_REFERENCE_SAME_AS_DISPLAY)
            binding.reference_core_style = (
                binding.display_core_style
                if ref_key == self._NUM_REFERENCE_SAME_AS_DISPLAY
                else ref_key
            )
        previous_separator = str(getattr(binding, "title_separator", "") or "")
        binding.title_separator = self._whitespace_widget_text(separator_edit)
        if binding.title_separator == previous_separator:
            pass
        elif binding.title_separator == "\t":
            binding.ooxml_separator_mode = "suff"
            binding.ooxml_suff = "tab"
        elif binding.title_separator == "":
            binding.ooxml_separator_mode = "none"
            binding.ooxml_suff = None
        else:
            binding.ooxml_separator_mode = "inline"
            binding.ooxml_suff = "nothing"
        if isinstance(start_spin, QSpinBox):
            binding.start_at = max(1, int(start_spin.value()))
        if isinstance(restart_combo, QComboBox):
            restart_value = str(restart_combo.currentData() or "").strip()
            binding.restart_on = restart_value or None
        if isinstance(toc_toggle, OnOffToggleButton):
            binding.include_in_toc = toc_toggle.isChecked()
        if isinstance(enabled_toggle, OnOffToggleButton):
            binding.enabled = enabled_toggle.isChecked()
        return binding

    def _uses_decimal_parent_chain_preview(
        self,
        level_name: str,
        binding: HeadingLevelBindingConfig,
        *,
        bindings: dict[str, HeadingLevelBindingConfig] | None = None,
        catalog_source=None,
    ) -> bool:
        source_v2 = copy.deepcopy(self._num_catalog_source if catalog_source is None else catalog_source)
        if bindings is not None:
            source_v2.level_bindings = copy.deepcopy(bindings)
        return uses_decimal_parent_chain(level_name, binding, source_v2)

    def _numbering_bindings_from_ui(
        self,
        *,
        error_collector: list[str] | None = None,
        catalog_source=None,
    ) -> dict[str, HeadingLevelBindingConfig]:
        bindings: dict[str, HeadingLevelBindingConfig] = {}
        for row, level_name in enumerate(getattr(self, "_num_level_keys", [])):
            bindings[level_name] = self._numbering_binding_from_row(
                row,
                level_name,
                error_collector=error_collector,
                catalog_source=catalog_source,
            )
        return bindings

    def _numbering_preview_text(
        self,
        level_name: str,
        bindings: dict[str, HeadingLevelBindingConfig],
        *,
        catalog_source=None,
    ) -> str:
        source_v2 = self._num_catalog_source if catalog_source is None else catalog_source
        binding = bindings.get(level_name, HeadingLevelBindingConfig())
        shell = getattr(source_v2, "shell_catalog", {}).get(binding.display_shell)
        chain = getattr(source_v2, "chain_catalog", {}).get(binding.chain)
        preview_counters = self._preview_counters_for_level(level_name, bindings)

        prefix = str(getattr(shell, "prefix", "") or "")
        suffix = str(getattr(shell, "suffix", "") or "")
        segments = getattr(chain, "segments", None) or []
        force_decimal_chain = self._uses_decimal_parent_chain_preview(
            level_name,
            binding,
            bindings=bindings,
            catalog_source=source_v2,
        )

        chain_parts: list[str] = []
        for segment in segments:
            segment_type = str(getattr(segment, "type", "") or "")
            if segment_type == "literal":
                chain_parts.append(str(getattr(segment, "text", "") or ""))
                continue

            source = str(getattr(segment, "source", "") or "")
            if source == "current":
                style_id = binding.display_core_style
                level_value = int(preview_counters.get(level_name, 0) or 0)
            elif source.startswith("level"):
                if force_decimal_chain:
                    style_id = "arabic"
                else:
                    ref_level_name = f"heading{source[5:]}"
                    ref_binding = bindings.get(ref_level_name, HeadingLevelBindingConfig())
                    style_id = str(ref_binding.reference_core_style or ref_binding.display_core_style or "arabic")
                ref_level_name = f"heading{source[5:]}"
                level_value = int(preview_counters.get(ref_level_name, 0) or 0)
            else:
                style_id = "arabic"
                level_value = 1
            chain_parts.append(
                self._numbering_core_style_preview_value(
                    style_id,
                    level_value,
                    catalog_source=source_v2,
                )
            )

        if not chain_parts:
            chain_parts.append(
                self._numbering_core_style_preview_value(
                    binding.display_core_style,
                    int(preview_counters.get(level_name, 0) or 1),
                    catalog_source=source_v2,
                )
            )

        title = self._LEVEL_PREVIEW_TITLES.get(level_name, "示例标题")
        preview_separator = self._numbering_preview_separator_text(binding.title_separator)
        preview = f"{prefix}{''.join(chain_parts)}{suffix}{preview_separator}{title}"
        if not binding.enabled:
            return "[停用]"
        return preview

    @staticmethod
    def _numbering_preview_separator_text(separator: str | None) -> str:
        raw = str(separator or "")
        if "\t" not in raw:
            return raw
        return raw.replace("\t", "   ")

    def _refresh_numbering_previews(self) -> None:
        preview_catalog_source = copy.deepcopy(self._num_catalog_source)
        bindings = self._numbering_bindings_from_ui(
            error_collector=[],
            catalog_source=preview_catalog_source,
        )
        preview_col = self._num_col_index("preview")
        for row, level_name in enumerate(getattr(self, "_num_level_keys", [])):
            item = self._num_table.item(row, preview_col)
            if item is None:
                item = QTableWidgetItem()
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self._num_table.setItem(row, preview_col, item)
            item.setTextAlignment(Qt.AlignCenter)
            preview_text = self._numbering_preview_text(
                level_name,
                bindings,
                catalog_source=preview_catalog_source,
            )
            item.setText(preview_text)
            item.setToolTip(preview_text)

    def _sync_caption_separator_from_ui(self, cap_cfg) -> None:
        combo = getattr(self, "_cap_separator_mode_combo", None)
        custom_edit = getattr(self, "_cap_custom_separator_edit", None)
        if combo is None or custom_edit is None:
            return
        mode_key = str(combo.currentData() or "fullwidth_space")
        if mode_key == "custom":
            cap_cfg.separator = custom_edit.text()
            return
        cap_cfg.separator = self._NUM_SEPARATOR_PRESET_VALUES.get(mode_key, "\u3000")

    def _build_numbering_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        self._numbering_scheme_id = self._preferred_numbering_scheme_id()
        controls_layout = QHBoxLayout()
        controls_layout.addWidget(QLabel("配置方式:"))
        self._num_mode_combo = ScrollSafeComboBox()
        for preset_id, meta in self._NUM_SCHEME_PRESETS.items():
            self._num_mode_combo.addItem(str(meta.get("label", preset_id)), preset_id)
        self._num_mode_combo.addItem("自定义", "custom")
        controls_layout.addWidget(self._num_mode_combo, 1)
        controls_layout.addStretch(1)
        layout.addLayout(controls_layout)

        self._num_preset_desc_label = QLabel("")
        self._num_preset_desc_label.setWordWrap(True)
        layout.addWidget(self._num_preset_desc_label)

        self._num_detail_container = QWidget()
        detail_layout = QVBoxLayout(self._num_detail_container)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_layout.setSpacing(layout.spacing())

        note = QLabel(
            "此处支持“预设”和“自定义”两种配置方式。预设会批量套用整套标题编号方案，"
            "自定义则按结构化组件逐级编辑：预览 / 外壳 / 前级组合 / 本体样式 / 下级引用。"
            "外壳支持直接输入自定义文本，且必须包含一个 {}；前级组合在弹出编辑器中定义，主表只显示摘要。"
            "起始值、遇级重排规则、是否进入目录也会在这里逐级控制；保存时会同步回旧版编号模板以兼容当前引擎。"
        )
        note.setWordWrap(True)
        detail_layout.addWidget(note)

        self._num_table = QTableWidget()
        self._num_table.setObjectName("StyleConfigTable")
        control_inner_height = max(18, self._NUM_TABLE_CONTROL_HEIGHT - 2)
        self._num_table.setStyleSheet(
            f"""
            QTableWidget#StyleConfigTable QComboBox,
            QTableWidget#StyleConfigTable QSpinBox,
            QTableWidget#StyleConfigTable QPushButton {{
                min-height: {control_inner_height}px;
                max-height: {control_inner_height}px;
            }}
            QTableWidget#StyleConfigTable QComboBox {{
                padding: 0 14px 0 3px;
            }}
            QTableWidget#StyleConfigTable QComboBox::drop-down {{
                width: 14px;
            }}
            QTableWidget#StyleConfigTable QSpinBox {{
                padding: 0 14px 0 4px;
            }}
            QTableWidget#StyleConfigTable QSpinBox::up-button,
            QTableWidget#StyleConfigTable QSpinBox::down-button {{
                width: 14px;
            }}
            QTableWidget#StyleConfigTable QPushButton {{
                padding: 0 8px;
            }}
            """
        )
        cols = self._NUM_COLS
        self._num_table.setColumnCount(len(cols))
        self._num_table.setHorizontalHeaderLabels([c[1] for c in cols])
        for index, (attr, _label, _kind) in enumerate(cols):
            header_item = self._num_table.horizontalHeaderItem(index)
            if header_item is not None:
                header_item.setToolTip(self._NUM_COL_TOOLTIPS.get(attr, ""))
        header = self._num_table.horizontalHeader()
        vertical_header = self._num_table.verticalHeader()
        header.setSectionResizeMode(QHeaderView.Fixed)
        header.setStretchLastSection(False)
        vertical_header.setSectionResizeMode(QHeaderView.Fixed)
        vertical_header.setDefaultSectionSize(self._NUM_TABLE_ROW_HEIGHT)
        vertical_header.setMinimumSectionSize(self._NUM_TABLE_ROW_HEIGHT)
        self._num_table.setMinimumWidth(self._numbering_table_min_width())
        self._num_table.setColumnWidth(self._num_col_index("preview"), self._NUM_PREVIEW_BASE_WIDTH)
        for attr, width in self._NUM_TABLE_COLUMN_WIDTHS.items():
            self._num_table.setColumnWidth(self._num_col_index(attr), width)
        self._num_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self._num_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        detail_layout.addWidget(self._num_table)
        layout.addWidget(self._num_detail_container)
        self._load_numbering_rows_for_scheme(self._numbering_scheme_id)
        self._num_custom_ui_state = self._capture_numbering_ui_state()
        # Use explicitly stored preset mode first; fall back to signature detection
        stored_mode = str(
            getattr(self._config, "_numbering_preset_mode", "") or ""
        ).strip()
        if stored_mode and (
            stored_mode in self._NUM_SCHEME_PRESETS or stored_mode == "custom"
        ):
            initial_mode = stored_mode
        else:
            initial_mode = self._detect_matching_numbering_preset_id(
                copy.deepcopy(getattr(self._num_catalog_source, "level_bindings", {}))
            ) or "custom"
        self._set_combo_current_data(self._num_mode_combo, initial_mode)
        self._update_numbering_preset_description()
        self._sync_numbering_mode_controls(apply_selected_preset=False)
        self._num_mode_combo.currentIndexChanged.connect(
            lambda *_args: self._sync_numbering_mode_controls(
                apply_selected_preset=str(self._num_mode_combo.currentData() or "custom") != "custom"
            )
        )
        return w

    def _preferred_numbering_scheme_id(self) -> str:
        hn = self._config.heading_numbering
        schemes = getattr(hn, "schemes", {}) or {}
        if "2" in schemes:
            return "2"
        current_scheme = str(getattr(hn, "scheme", "") or "").strip()
        if current_scheme and current_scheme in schemes:
            return current_scheme
        if schemes:
            return str(next(iter(schemes.keys())))
        return "__levels__"

    def _resolve_numbering_levels(self, scheme_id: str, *, config: SceneConfig | None = None) -> dict:
        cfg = config or self._config
        hn = cfg.heading_numbering
        sid = str(scheme_id or "").strip()
        if sid == "__levels__":
            return hn.levels
        if sid in hn.schemes:
            return hn.schemes[sid]
        # 兜底：当 scheme 缺失时，退回活动 levels。
        return hn.levels

    # Canonical ordered list of all 8 heading level keys
    _ALL_HEADING_KEYS = [f"heading{i}" for i in range(1, 9)]

    def _load_numbering_rows_for_scheme(self, scheme_id: str):
        del scheme_id
        self._num_catalog_source = copy.deepcopy(self._numbering_v2_display_config())
        # Always show all 8 heading levels
        self._num_level_keys = list(self._ALL_HEADING_KEYS)
        self._num_table.setRowCount(len(self._num_level_keys))
        self._num_table.setVerticalHeaderLabels(
            [self._LEVEL_DISPLAY.get(k, k) for k in self._num_level_keys]
        )
        preview_col = self._num_col_index("preview")
        self._num_row_widgets = {}
        shell_options = self._numbering_shell_options()
        core_style_options = self._numbering_core_style_options()
        reference_options = self._numbering_reference_options()
        core_style_compact_labels = self._numbering_core_style_compact_labels()
        reference_compact_labels = self._numbering_reference_compact_labels()
        for row, key in enumerate(self._num_level_keys):
            binding = copy.deepcopy(
                getattr(self._num_catalog_source, "level_bindings", {}).get(key, HeadingLevelBindingConfig())
            )
            preview_item = QTableWidgetItem("")
            preview_item.setFlags(preview_item.flags() & ~Qt.ItemIsEditable)
            preview_item.setTextAlignment(Qt.AlignCenter)
            self._num_table.setItem(row, preview_col, preview_item)

            shell_combo = self._build_numbering_choice_combo(
                shell_options,
                binding.display_shell,
                editable=True,
                text_alignment=Qt.AlignCenter,
            )
            shell_combo.setToolTip("可直接输入自定义外壳，格式示例：第{}章、【{}】、附录{}。")
            shell_combo.currentIndexChanged.connect(lambda *_args: self._refresh_numbering_previews())
            if shell_combo.lineEdit() is not None:
                shell_combo.lineEdit().textChanged.connect(lambda *_args: self._refresh_numbering_previews())
            self._num_table.setCellWidget(
                row,
                self._num_col_index("display_shell"),
                self._wrap_num_cell_widget(shell_combo),
            )

            chain_button = QPushButton(self._numbering_chain_summary(key, binding.chain)[0])
            chain_button.setFixedHeight(self._NUM_TABLE_CONTROL_HEIGHT)
            chain_button.clicked.connect(lambda *_args, row=row, key=key: self._edit_numbering_chain(row, key))
            self._set_chain_button_state(chain_button, key, binding.chain)
            self._num_table.setCellWidget(
                row,
                self._num_col_index("chain_summary"),
                self._wrap_num_cell_widget(chain_button),
            )

            core_style_combo = self._build_numbering_choice_combo(
                core_style_options,
                binding.display_core_style,
                compact_current_labels=core_style_compact_labels,
            )
            core_style_combo.currentIndexChanged.connect(lambda *_args: self._refresh_numbering_previews())
            self._num_table.setCellWidget(
                row,
                self._num_col_index("display_core_style"),
                self._wrap_num_cell_widget(core_style_combo),
            )

            ref_key = (
                self._NUM_REFERENCE_SAME_AS_DISPLAY
                if binding.reference_core_style == binding.display_core_style
                else binding.reference_core_style
            )
            reference_combo = self._build_numbering_choice_combo(
                reference_options,
                ref_key,
                text_alignment=Qt.AlignCenter,
                compact_current_labels=reference_compact_labels,
            )
            reference_combo.currentIndexChanged.connect(lambda *_args: self._refresh_numbering_previews())
            self._num_table.setCellWidget(
                row,
                self._num_col_index("reference_core_style"),
                self._wrap_num_cell_widget(reference_combo),
            )

            separator_edit = WhitespacePresetWidget(
                self._NUM_SEPARATOR_OPTIONS,
                self._NUM_SEPARATOR_VALUE_TO_KEY,
                text=binding.title_separator,
                compact_labels=self._NUM_SEPARATOR_COMPACT_LABELS,
                control_height=self._NUM_TABLE_CONTROL_HEIGHT,
                compact_mode_width=self._NUM_TITLE_SEPARATOR_COMPACT_CONTROL_WIDTH,
            )
            separator_edit.setFixedHeight(self._NUM_TABLE_CONTROL_HEIGHT)
            separator_edit.textChanged.connect(lambda *_args: self._refresh_numbering_separator_column_width())
            separator_edit.textChanged.connect(lambda *_args: self._refresh_numbering_previews())
            self._num_table.setCellWidget(
                row,
                self._num_col_index("title_separator"),
                self._wrap_num_cell_widget(separator_edit),
            )

            start_spin = ScrollSafeSpinBox(self._num_table)
            start_spin.setRange(1, 999)
            start_spin.setValue(normalize_start_at(getattr(binding, "start_at", 1)))
            start_spin.setAlignment(Qt.AlignCenter)
            start_spin.setButtonSymbols(QAbstractSpinBox.UpDownArrows)
            start_spin.setFixedHeight(self._NUM_TABLE_CONTROL_HEIGHT)
            start_spin.setFixedWidth(self._NUM_START_AT_CONTROL_WIDTH)
            start_spin.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            start_spin.valueChanged.connect(lambda *_args: self._refresh_numbering_previews())
            self._num_table.setCellWidget(
                row,
                self._num_col_index("start_at"),
                self._wrap_num_cell_widget(start_spin, center_horizontally=True),
            )

            restart_combo = self._build_numbering_choice_combo(
                self._numbering_restart_options(key),
                str(getattr(binding, "restart_on", "") or ""),
                text_alignment=Qt.AlignCenter,
            )
            restart_combo.setFixedHeight(self._NUM_TABLE_CONTROL_HEIGHT)
            restart_combo.currentIndexChanged.connect(lambda *_args: self._refresh_numbering_previews())
            self._num_table.setCellWidget(
                row,
                self._num_col_index("restart_on"),
                self._wrap_num_cell_widget(restart_combo),
            )

            toc_toggle = self._build_bool_toggle(bool(getattr(binding, "include_in_toc", True)))
            toc_toggle.setFixedHeight(self._NUM_TABLE_CONTROL_HEIGHT)
            self._num_table.setCellWidget(
                row,
                self._num_col_index("include_in_toc"),
                self._wrap_num_cell_widget(toc_toggle, center_horizontally=True),
            )

            enabled_toggle = self._build_bool_toggle(binding.enabled)
            enabled_toggle.setFixedHeight(self._NUM_TABLE_CONTROL_HEIGHT)
            enabled_toggle.toggled.connect(lambda *_args: self._refresh_numbering_previews())
            self._num_table.setCellWidget(
                row,
                self._num_col_index("enabled"),
                self._wrap_num_cell_widget(enabled_toggle, center_horizontally=True),
            )
            self._num_table.setRowHeight(row, self._NUM_TABLE_ROW_HEIGHT)

            self._num_row_widgets[row] = {
                "display_shell": shell_combo,
                "chain_summary": chain_button,
                "display_core_style": core_style_combo,
                "reference_core_style": reference_combo,
                "title_separator": separator_edit,
                "start_at": start_spin,
                "restart_on": restart_combo,
                "include_in_toc": toc_toggle,
                "enabled": enabled_toggle,
            }
        self._refresh_numbering_separator_column_width()
        self._refresh_numbering_previews()

    def _merge_numbering_catalog_source_into_config(self, config: SceneConfig) -> None:
        target_v2 = config.heading_numbering_v2
        source_v2 = getattr(self, "_num_catalog_source", None)
        if source_v2 is None:
            return
        for attr in ("shell_catalog", "core_style_catalog", "chain_catalog", "preset_catalog"):
            target_catalog = getattr(target_v2, attr, None)
            source_catalog = getattr(source_v2, attr, None)
            if not isinstance(target_catalog, dict) or not isinstance(source_catalog, dict):
                continue
            for key, value in source_catalog.items():
                target_catalog[str(key)] = copy.deepcopy(value)

    def _legacy_format_from_binding(self, binding: HeadingLevelBindingConfig) -> str:
        return legacy_format_from_binding(binding)

    def _legacy_chain_template_from_binding(
        self,
        level_name: str,
        binding: HeadingLevelBindingConfig,
        bindings: dict[str, HeadingLevelBindingConfig],
    ) -> str:
        source_v2 = copy.deepcopy(getattr(self, "_num_catalog_source", HeadingNumberingV2Config()))
        source_v2.level_bindings = copy.deepcopy(bindings)
        if self._uses_decimal_parent_chain_preview(level_name, binding, bindings=bindings, catalog_source=source_v2):
            return "{parent}.{current}"
        return explicit_chain_template_from_binding(level_name, binding, source_v2)

    def _legacy_template_from_binding(
        self,
        level_name: str,
        binding: HeadingLevelBindingConfig,
        bindings: dict[str, HeadingLevelBindingConfig],
    ) -> str:
        source_v2 = copy.deepcopy(getattr(self, "_num_catalog_source", HeadingNumberingV2Config()))
        source_v2.level_bindings = copy.deepcopy(bindings)
        return legacy_template_from_binding(level_name, binding, source_v2)

    def _legacy_levels_from_numbering_bindings(
        self,
        bindings: dict[str, HeadingLevelBindingConfig],
    ) -> dict[str, HeadingLevelConfig]:
        source_v2 = copy.deepcopy(getattr(self, "_num_catalog_source", HeadingNumberingV2Config()))
        source_v2.level_bindings = copy.deepcopy(bindings)
        return legacy_levels_from_v2(source_v2)

    def _sync_legacy_heading_numbering_from_v2(
        self,
        config: SceneConfig,
        bindings: dict[str, HeadingLevelBindingConfig],
    ) -> list[str]:
        levels = self._legacy_levels_from_numbering_bindings(bindings)
        invalid_templates: list[str] = []
        for level_name, level_cfg in levels.items():
            template_errors = validate_heading_numbering_template(level_name, level_cfg.template)
            if template_errors:
                invalid_templates.append(
                    f"{self._LEVEL_DISPLAY.get(level_name, level_name)} / 模板 = {level_cfg.template} "
                    f"({'；'.join(template_errors)})"
                )

        hn = config.heading_numbering
        selected_scheme = str(
            getattr(self, "_numbering_scheme_id", "") or self._preferred_numbering_scheme_id()
        ).strip()
        if selected_scheme and selected_scheme != "__levels__":
            hn.scheme = selected_scheme
            schemes = getattr(hn, "schemes", None)
            if isinstance(schemes, dict):
                schemes[selected_scheme] = copy.deepcopy(levels)
        hn.levels = copy.deepcopy(levels)
        if selected_scheme and selected_scheme != "__levels__" and selected_scheme in getattr(hn, "schemes", {}):
            hn.apply_scheme(selected_scheme)
        return invalid_templates

    @staticmethod
    def _parse_bool_text(text: str) -> bool:
        val = (text or "").strip().lower()
        truthy = {"1", "true", "yes", "y", "on", "是", "√"}
        falsy = {"0", "false", "no", "n", "off", "否", "×"}
        if val in truthy:
            return True
        if val in falsy:
            return False
        raise ValueError(f"invalid bool literal: {text}")

    # ── 标题高级 ──

    def _build_heading_advanced_tab(self) -> QWidget:
        w = QWidget()
        layout = QFormLayout(w)
        hn = self._config.heading_numbering

        self._enforcement_checks = {}
        for attr, label in [
            ("ban_tab", "禁止Tab分隔"),
            ("ban_double_halfwidth_space", "禁止双半角空格"),
            ("ban_mixed_separator_per_level", "禁止同级混用分隔符"),
            ("auto_fix", "自动修复"),
        ]:
            cb = QCheckBox()
            cb.setChecked(bool(getattr(hn.enforcement, attr, False)))
            self._enforcement_checks[attr] = cb
            layout.addRow(label + ":", cb)

        layout.addRow(QLabel("标题识别风险防护"))
        self._risk_guard_controls = {}

        def _add_guard_bool(attr: str, label: str):
            cb = QCheckBox()
            cb.setChecked(bool(getattr(hn.risk_guard, attr, False)))
            self._risk_guard_controls[attr] = cb
            layout.addRow(label + ":", cb)

        def _add_guard_int(attr: str, label: str, minimum: int = 0, maximum: int = 10000):
            spin = ScrollSafeSpinBox()
            spin.setRange(minimum, maximum)
            spin.setValue(int(getattr(hn.risk_guard, attr, 0)))
            self._risk_guard_controls[attr] = spin
            layout.addRow(label + ":", spin)

        def _add_guard_float(attr: str, label: str, minimum: float = 0.0, maximum: float = 1.0):
            spin = ScrollSafeDoubleSpinBox()
            spin.setRange(minimum, maximum)
            spin.setDecimals(3)
            spin.setSingleStep(0.01)
            spin.setValue(float(getattr(hn.risk_guard, attr, 0.0)))
            self._risk_guard_controls[attr] = spin
            layout.addRow(label + ":", spin)

        _add_guard_bool("enabled", "启用防护")
        _add_guard_int("no_body_min_candidates", "无正文时最少候选数")
        _add_guard_int("no_body_min_chapters", "无正文时最少章数")
        _add_guard_int("no_chapter_min_outside_chapters", "无章时章外最少候选")
        _add_guard_int("tiny_body_max_abs_paras", "正文过短最大段数")
        _add_guard_float("tiny_body_max_ratio", "正文过短最大比例")
        _add_guard_int("tiny_body_min_candidates", "正文过短最少候选数")
        _add_guard_int("tiny_body_primary_margin", "正文过短主边界余量")
        _add_guard_bool("keep_after_first_chapter", "首章后保留")
        return w

    # ── 题注配置 ──

    def _build_caption_tab(self) -> QWidget:
        w = QWidget()
        layout = QFormLayout(w)
        cap = self._config.caption
        equation_table_cfg = getattr(self._config, "equation_table_format", None)
        self._cap_edits = {}
        for attr, label in [
            ("figure_prefix", "图前缀"),
            ("table_prefix", "表前缀"),
            ("placeholder", "缺失题注占位"),
        ]:
            edit = QLineEdit(str(getattr(cap, attr, "")))
            self._cap_edits[attr] = edit
            layout.addRow(label + ":", edit)

        # 题注分隔符：预设下拉 + 自定义输入（仅在“自定义”时启用）
        cap_separator = getattr(cap, "separator", "")
        if cap_separator is None:
            cap_separator = ""
        cap_separator = str(cap_separator)
        mode_key = self._NUM_SEPARATOR_VALUE_TO_KEY.get(cap_separator, "custom")
        custom_value = cap_separator if mode_key == "custom" else ""

        self._cap_separator_mode_combo = ScrollSafeComboBox()
        for key, label, _ in self._NUM_SEPARATOR_OPTIONS:
            self._cap_separator_mode_combo.addItem(label, key)
        self._cap_separator_mode_combo.setToolTip("全角空格显示为 □，半角空格显示为 ·。")
        mode_idx = self._cap_separator_mode_combo.findData(mode_key)
        if mode_idx < 0:
            mode_idx = self._cap_separator_mode_combo.findData("fullwidth_space")
        self._cap_separator_mode_combo.setCurrentIndex(max(mode_idx, 0))
        layout.addRow("分隔符:", self._cap_separator_mode_combo)

        self._cap_custom_separator_edit = WhitespaceVisibleLineEdit(custom_value)
        self._cap_custom_separator_edit.setToolTip("输入时：全角空格=□，半角空格=·，Tab=➡。")
        self._cap_custom_separator_edit.setEnabled(
            str(self._cap_separator_mode_combo.currentData()) == "custom"
        )
        self._cap_separator_mode_combo.currentIndexChanged.connect(
            lambda *_args: self._cap_custom_separator_edit.setEnabled(
                str(self._cap_separator_mode_combo.currentData()) == "custom"
            )
        )
        layout.addRow("自定义分隔符:", self._cap_custom_separator_edit)

        # 编号格式下拉框
        self._numbering_format_combo = ScrollSafeComboBox()
        _NUMBERING_OPTIONS = [
            ("seq",                  "图1、图2、图3 …（纯序号）"),
            ("chapter.seq",          "图1.1、图1.2 … 句点分隔"),
            ("chapter-seq",          "图1-1、图1-2 … 连字符分隔"),
            ("chapter:seq",          "图1:1、图1:2 … 冒号分隔"),
            ("chapter\u2014seq",     "图1\u20141、图1\u20142 … 长划线分隔"),
            ("chapter\u2013seq",     "图1\u20131、图1\u20132 … 短划线分隔"),
        ]
        current_fmt = (cap.numbering_format or "chapter.seq").strip().lower()
        selected_idx = 0
        for i, (value, label) in enumerate(_NUMBERING_OPTIONS):
            self._numbering_format_combo.addItem(label, value)
            if current_fmt == value:
                selected_idx = i
        self._numbering_format_combo.setCurrentIndex(selected_idx)
        layout.addRow("编号格式:", self._numbering_format_combo)

        self._equation_numbering_format_combo = ScrollSafeComboBox()
        equation_numbering_options = [
            ("seq", "(1)、(2) … 纯序号"),
            ("chapter-seq", "(1-1)、(1-2) … 连字符"),
            ("chapter.seq", "(1.1)、(1.2) … 句点"),
        ]
        current_equation_fmt = str(
            getattr(equation_table_cfg, "numbering_format", "chapter.seq") or "chapter.seq"
        ).strip().lower()
        equation_selected_idx = 2
        for i, (value, label) in enumerate(equation_numbering_options):
            self._equation_numbering_format_combo.addItem(label, value)
            if current_equation_fmt == value:
                equation_selected_idx = i
        self._equation_numbering_format_combo.setCurrentIndex(equation_selected_idx)
        layout.addRow("公式编号格式:", self._equation_numbering_format_combo)

        self._cap_enabled = QCheckBox("启用题注处理")
        self._cap_enabled.setChecked(cap.enabled)
        layout.addRow(self._cap_enabled)
        return w

    # ── 表格配置 ──

    def _build_table_tab(self) -> QWidget:
        w = QWidget()
        outer_layout = QVBoxLayout(w)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        outer_layout.addWidget(scroll)

        content = QWidget()
        scroll.setWidget(content)

        layout = QFormLayout(content)
        layout.setFormAlignment(Qt.AlignTop)
        layout.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        layout.setHorizontalSpacing(18)
        layout.setVerticalSpacing(10)
        cfg = self._config
        formula_style_cfg = getattr(cfg, "formula_style", None)
        formula_table_cfg = getattr(cfg, "formula_table", None)

        layout.addRow(self._build_section_divider("常规表格线宽"))

        note = QLabel(
            "线宽单位：磅；常用值：0.5 / 0.75 / 1.0 / 1.5"
        )
        note.setWordWrap(True)
        layout.addRow(note)
        layout.addRow(self._build_section_divider("线宽设置"))

        # 全框线表线宽
        self._grid_border_width_spin = ScrollSafeDoubleSpinBox()
        self._grid_border_width_spin.setRange(0.25, 3.0)
        self._grid_border_width_spin.setSingleStep(0.25)
        self._grid_border_width_spin.setDecimals(2)
        self._grid_border_width_spin.setValue(
            getattr(cfg, "table_border_width_pt", 0.5)
        )
        layout.addRow("全框线线宽:", self._grid_border_width_spin)

        # 三线表表头线宽
        self._three_line_header_spin = ScrollSafeDoubleSpinBox()
        self._three_line_header_spin.setRange(0.25, 3.0)
        self._three_line_header_spin.setSingleStep(0.25)
        self._three_line_header_spin.setDecimals(2)
        self._three_line_header_spin.setValue(
            getattr(cfg, "three_line_header_width_pt", 1.0)
        )
        layout.addRow("三线表表头线宽:", self._three_line_header_spin)

        # 三线表表尾线宽
        self._three_line_bottom_spin = ScrollSafeDoubleSpinBox()
        self._three_line_bottom_spin.setRange(0.25, 3.0)
        self._three_line_bottom_spin.setSingleStep(0.25)
        self._three_line_bottom_spin.setDecimals(2)
        self._three_line_bottom_spin.setValue(
            getattr(cfg, "three_line_bottom_width_pt", 0.5)
        )
        layout.addRow("三线表表尾线宽:", self._three_line_bottom_spin)

        layout.addRow(self._build_section_divider("公式表格配置"))

        formula_panel = QWidget()
        formula_layout = QVBoxLayout(formula_panel)
        formula_layout.setContentsMargins(0, 0, 0, 0)
        formula_layout.setSpacing(10)

        formula_note = QLabel(
            "说明：这里只配置样式参数；是否执行，仍由“实验室 - 公式表格管理”控制。"
        )
        formula_note.setWordWrap(True)
        formula_layout.addWidget(formula_note)

        formula_layout.addWidget(self._build_section_divider("统一处理"))
        self._formula_table_unify_font_check = QCheckBox("统一公式字体")
        self._formula_table_unify_font_check.setChecked(
            bool(getattr(formula_style_cfg, "unify_font", True))
        )
        self._formula_table_unify_size_check = QCheckBox("统一公式字号")
        self._formula_table_unify_size_check.setChecked(
            bool(getattr(formula_style_cfg, "unify_size", True))
        )
        self._formula_table_unify_spacing_check = QCheckBox("统一公式段距/行距")
        self._formula_table_unify_spacing_check.setChecked(
            bool(getattr(formula_style_cfg, "unify_spacing", True))
        )
        unify_row = QWidget()
        unify_row_layout = QHBoxLayout(unify_row)
        unify_row_layout.setContentsMargins(0, 0, 0, 0)
        unify_row_layout.setSpacing(12)
        unify_row_layout.addWidget(self._formula_table_unify_font_check)
        unify_row_layout.addWidget(self._formula_table_unify_size_check)
        unify_row_layout.addWidget(self._formula_table_unify_spacing_check)
        unify_row_layout.addStretch(1)
        formula_layout.addWidget(
            self._build_compact_option_row(("统一处理:", unify_row), add_stretch=False)
        )

        formula_layout.addWidget(self._build_section_divider("公式样式"))
        self._formula_table_font_combo = self._build_font_family_combo(
            getattr(formula_table_cfg, "formula_font_name", "Cambria Math")
        )
        self._formula_table_number_font_combo = self._build_font_family_combo(
            getattr(formula_table_cfg, "number_font_name", "Times New Roman")
        )
        formula_layout.addWidget(
            self._build_compact_option_row(
                ("公式字体:", self._formula_table_font_combo),
                ("编号字体:", self._formula_table_number_font_combo),
            )
        )

        self._formula_table_font_size_combo = self._build_font_size_combo(
            self._font_size_display_text(
                getattr(formula_table_cfg, "formula_font_size_pt", 12.0),
                getattr(formula_table_cfg, "formula_font_size_display", ""),
            )
        )
        self._formula_table_number_size_combo = self._build_font_size_combo(
            self._font_size_display_text(
                getattr(formula_table_cfg, "number_font_size_pt", 10.5),
                getattr(formula_table_cfg, "number_font_size_display", ""),
            )
        )
        formula_layout.addWidget(
            self._build_compact_option_row(
                ("公式字号:", self._formula_table_font_size_combo),
                ("编号字号:", self._formula_table_number_size_combo),
            )
        )

        formula_layout.addWidget(self._build_section_divider("间距"))
        self._formula_table_line_spacing_spin = ScrollSafeDoubleSpinBox()
        self._formula_table_line_spacing_spin.setRange(0.5, 3.0)
        self._formula_table_line_spacing_spin.setSingleStep(0.1)
        self._formula_table_line_spacing_spin.setDecimals(2)
        self._formula_table_line_spacing_spin.setFixedWidth(88)
        self._formula_table_line_spacing_spin.setValue(
            float(getattr(formula_table_cfg, "formula_line_spacing", 1.0))
        )

        self._formula_table_space_before_spin = ScrollSafeDoubleSpinBox()
        self._formula_table_space_before_spin.setRange(0.0, 24.0)
        self._formula_table_space_before_spin.setSingleStep(0.5)
        self._formula_table_space_before_spin.setDecimals(1)
        self._formula_table_space_before_spin.setFixedWidth(88)
        self._formula_table_space_before_spin.setValue(
            float(getattr(formula_table_cfg, "formula_space_before_pt", 0.0))
        )

        self._formula_table_space_after_spin = ScrollSafeDoubleSpinBox()
        self._formula_table_space_after_spin.setRange(0.0, 24.0)
        self._formula_table_space_after_spin.setSingleStep(0.5)
        self._formula_table_space_after_spin.setDecimals(1)
        self._formula_table_space_after_spin.setFixedWidth(88)
        self._formula_table_space_after_spin.setValue(
            float(getattr(formula_table_cfg, "formula_space_after_pt", 0.0))
        )
        formula_layout.addWidget(
            self._build_compact_option_row(
                ("公式行距(倍):", self._formula_table_line_spacing_spin),
                ("段前(磅):", self._formula_table_space_before_spin),
                ("段后(磅):", self._formula_table_space_after_spin),
            )
        )

        formula_layout.addWidget(self._build_section_divider("对齐"))
        self._formula_table_block_alignment_combo = self._build_alignment_combo(
            getattr(formula_table_cfg, "block_alignment", "center")
        )

        self._formula_table_table_alignment_combo = self._build_choice_combo(
            [
                ("left", "左对齐"),
                ("center", "居中对齐"),
                ("right", "右对齐"),
            ],
            getattr(formula_table_cfg, "table_alignment", "center"),
            tooltip="公式表格容器的整体对齐方式。",
        )

        self._formula_table_cell_alignment_combo = self._build_alignment_combo(
            getattr(formula_table_cfg, "formula_cell_alignment", "center")
        )

        self._formula_table_number_alignment_combo = self._build_alignment_combo(
            getattr(formula_table_cfg, "number_alignment", "right")
        )
        formula_layout.addWidget(
            self._build_compact_option_row(
                ("行间公式对齐:", self._formula_table_block_alignment_combo),
                ("公式表格对齐:", self._formula_table_table_alignment_combo),
            )
        )
        formula_layout.addWidget(
            self._build_compact_option_row(
                ("公式单元格对齐:", self._formula_table_cell_alignment_combo),
                ("编号单元格对齐:", self._formula_table_number_alignment_combo),
            )
        )

        formula_layout.addWidget(self._build_section_divider("编号列"))
        self._formula_table_auto_shrink_check = QCheckBox("自动压缩右侧编号列")
        self._formula_table_auto_shrink_check.setChecked(
            bool(getattr(formula_table_cfg, "auto_shrink_number_column", True))
        )
        self._formula_table_auto_shrink_check.setToolTip(
            "开启后会尽量缩窄公式表格右侧编号列，把更多宽度留给左侧公式。"
        )
        formula_layout.addWidget(
            self._build_compact_option_row(("编号列策略:", self._formula_table_auto_shrink_check))
        )

        layout.addRow(formula_panel)

        return w

    # ── 页面设置 ──

    def _build_page_tab(self) -> QWidget:
        w = QWidget()
        layout = QFormLayout(w)
        ps = self._config.page_setup
        self._paper_size_combo = ScrollSafeComboBox()
        for value, label in PAPER_SIZE_OPTIONS:
            self._paper_size_combo.addItem(label, value)
        paper_idx = self._paper_size_combo.findData(ps.paper_size)
        if paper_idx >= 0:
            self._paper_size_combo.setCurrentIndex(paper_idx)
        layout.addRow("纸张:", self._paper_size_combo)
        self._page_spins = {}
        for attr, label, val in [
            ("top_cm", "上边距(cm)", ps.margin.top_cm),
            ("bottom_cm", "下边距(cm)", ps.margin.bottom_cm),
            ("left_cm", "左边距(cm)", ps.margin.left_cm),
            ("right_cm", "右边距(cm)", ps.margin.right_cm),
            ("gutter_cm", "装订线(cm)", ps.gutter_cm),
            ("header_distance_cm", "页眉距离(cm)", ps.header_distance_cm),
            ("footer_distance_cm", "页脚距离(cm)", ps.footer_distance_cm),
        ]:
            spin = ScrollSafeDoubleSpinBox()
            spin.setRange(0, 20)
            spin.setDecimals(1)
            spin.setSingleStep(0.1)
            spin.setValue(val)
            self._page_spins[attr] = spin
            layout.addRow(label + ":", spin)
        return w

    # ── 输出选项 ──

    def _build_output_tab(self) -> QWidget:
        w = QWidget()
        layout = QFormLayout(w)
        out = self._config.output
        self._output_checks = {}
        for attr, label in [
            ("final_docx", "输出最终稿 DOCX"),
            ("compare_docx", "输出对比稿 DOCX"),
            ("compare_text", "对比稿包含文本变化"),
            ("compare_formatting", "对比稿包含格式变化"),
            ("report_json", "输出 JSON 报告"),
            ("report_markdown", "输出 Markdown 报告"),
        ]:
            cb = QCheckBox()
            cb.setChecked(bool(getattr(out, attr, False)))
            self._output_checks[attr] = cb
            layout.addRow(label + ":", cb)
        compare_doc_toggle = self._output_checks.get("compare_docx")
        if compare_doc_toggle is not None:
            compare_doc_toggle.toggled.connect(self._sync_compare_output_option_state)
        self._sync_compare_output_option_state(
            bool(self._output_checks.get("compare_docx") and self._output_checks["compare_docx"].isChecked())
        )
        return w

    def _sync_compare_output_option_state(self, enabled: bool) -> None:
        for attr in ("compare_text", "compare_formatting"):
            cb = getattr(self, "_output_checks", {}).get(attr)
            if cb is None:
                continue
            cb.setEnabled(bool(enabled))
            cb.setToolTip("" if enabled else "需先开启“输出对比稿 DOCX”后此选项才会生效。")

    # ── 执行流程 ──

    def _build_pipeline_tab(self) -> QWidget:
        w = QWidget()
        layout = QFormLayout(w)
        note = QLabel(
            "说明：md_cleanup 由主界面“实验室 - Markdown 文本修复”开关统一控制，"
            "whitespace_normalize 由主界面“实验室 - 空白清洗与中英文符号规范”开关统一控制，"
            "formula_convert / formula_to_table / equation_table_format / formula_style 由主界面“实验室 - 公式表格管理（总开关）”统一控制，"
            "citation_link 由主界面“实验室 - 正文参考文献域关联”开关统一控制，"
            "运行前会覆盖此页同名步骤状态。"
        )
        note.setWordWrap(True)
        layout.addRow(note)
        current_pipeline = list(self._config.pipeline or [])
        ordered_steps = []
        seen = set()
        for step in PIPELINE_DEFAULT_ORDER + current_pipeline:
            if step not in seen:
                seen.add(step)
                ordered_steps.append(step)
        self._pipeline_step_order = ordered_steps
        self._pipeline_checks = {}
        for step in ordered_steps:
            cb = QCheckBox()
            cb.setChecked(step in current_pipeline)
            if step == "md_cleanup":
                cb.setEnabled(False)
                cb.setToolTip("由主界面“实验室 - Markdown 文本修复”开关控制。")
            elif step == "whitespace_normalize":
                cb.setEnabled(False)
                cb.setToolTip("由主界面“实验室 - 空白清洗与中英文符号规范”开关控制。")
            elif step in {"formula_convert", "formula_to_table", "equation_table_format", "formula_style"}:
                cb.setEnabled(False)
                cb.setToolTip("由主界面“实验室 - 公式表格管理（总开关）”控制。")
            elif step == "citation_link":
                cb.setEnabled(False)
                cb.setToolTip("由主界面“实验室 - 正文参考文献域关联”开关控制。")
            self._pipeline_checks[step] = cb
            layout.addRow(f"{PIPELINE_STEP_LABELS.get(step, step)} ({step}):", cb)
        return w

    # ── 应用修改 ──

    def _apply_changes(self):
        draft = self._build_config_draft_from_ui()
        if draft is None:
            return
        self._commit_scene_config(self._config, draft)
        self.accept()

    def _sync_config_from_ui(self, *, target_config: SceneConfig | None = None) -> bool:
        config = target_config or self._config
        invalid_bool_cells = []
        invalid_font_size_cells = []
        invalid_indent_cells = []
        invalid_line_spacing_cells = []
        invalid_numbering_template_cells = []
        # 样式
        for row, key in enumerate(self._style_keys):
            sc = config.styles[key]
            for col, (attr, _, typ) in enumerate(_STYLE_COLUMNS):
                table_col = self._style_table_col_index(attr)
                if attr in {"font_cn", "font_en", "size_pt", "alignment"}:
                    widget = self._style_cell_widget(row, table_col)
                    if isinstance(widget, QComboBox):
                        if attr == "alignment":
                            text = str(widget.currentData() or "").strip()
                        else:
                            text = widget.currentText().strip()
                    else:
                        item = self._style_table.item(row, table_col)
                        if not item:
                            continue
                        text = item.text().strip()
                elif attr == "line_spacing_pt":
                    widget = self._style_cell_widget(row, table_col)
                    if isinstance(widget, LineSpacingValueWidget):
                        try:
                            line_spacing_type = widget.configSpacingType()
                            line_spacing_value = widget.resolvedValue()
                            sc.line_spacing_type = line_spacing_type
                            sc.line_spacing_pt = line_spacing_value
                        except (ValueError, TypeError):
                            invalid_line_spacing_cells.append(
                                f"{_STYLE_DISPLAY_NAMES.get(key, key)} / {_STYLE_COLUMNS[col][1]} = {widget.text()}"
                            )
                        continue
                    item = self._style_table.item(row, table_col)
                    if not item:
                        continue
                    text = item.text().strip()
                elif attr == "special_indent_value":
                    widget = self._style_cell_widget(row, table_col)
                    if isinstance(widget, SpecialIndentWidget):
                        try:
                            sc.special_indent_mode = widget.mode()
                            sc.special_indent_value = widget.configValue()
                            sc.special_indent_unit = widget.unit()
                        except (ValueError, TypeError):
                            invalid_indent_cells.append(
                                f"{_STYLE_DISPLAY_NAMES.get(key, key)} / {_STYLE_COLUMNS[col][1]} = {widget.text()}"
                            )
                        continue
                    item = self._style_table.item(row, table_col)
                    if not item:
                        continue
                    text = item.text().strip()
                elif attr in {"left_indent_chars", "right_indent_chars"}:
                    widget = self._style_cell_widget(row, table_col)
                    if isinstance(widget, IndentValueWidget):
                        try:
                            if attr == "left_indent_chars":
                                sc.left_indent_chars = widget.configValue()
                                sc.left_indent_unit = widget.unit()
                            else:
                                sc.right_indent_chars = widget.configValue()
                                sc.right_indent_unit = widget.unit()
                        except (ValueError, TypeError):
                            invalid_indent_cells.append(
                                f"{_STYLE_DISPLAY_NAMES.get(key, key)} / {_STYLE_COLUMNS[col][1]} = {widget.text()}"
                            )
                        continue
                    item = self._style_table.item(row, table_col)
                    if not item:
                        continue
                    text = item.text().strip()
                elif typ == "bool":
                    widget = self._style_cell_widget(row, table_col)
                    if isinstance(widget, OnOffToggleButton):
                        setattr(sc, attr, widget.isChecked())
                        continue
                    item = self._style_table.item(row, table_col)
                    if not item:
                        continue
                    text = item.text().strip()
                else:
                    item = self._style_table.item(row, table_col)
                    if not item:
                        continue
                    text = item.text().strip()
                if attr == "alignment":
                    text = normalize_alignment_value(text)
                try:
                    if attr == "size_pt":
                        setattr(sc, attr, parse_font_size_input(text))
                        sc.size_display = text
                    elif attr == "line_spacing_pt":
                        sc.line_spacing_pt = resolve_line_spacing_value(
                            getattr(sc, "line_spacing_type", "exact"),
                            text,
                        )
                    elif typ == "float":
                        setattr(sc, attr, float(text))
                    elif typ == "bool":
                        setattr(sc, attr, self._parse_bool_text(text))
                    else:
                        setattr(sc, attr, text)
                except (ValueError, TypeError):
                    if attr == "size_pt":
                        invalid_font_size_cells.append(
                            f"{_STYLE_DISPLAY_NAMES.get(key, key)} / {_STYLE_COLUMNS[col][1]} = {text}"
                        )
                    elif attr == "line_spacing_pt":
                        invalid_line_spacing_cells.append(
                            f"{_STYLE_DISPLAY_NAMES.get(key, key)} / {_STYLE_COLUMNS[col][1]} = {text}"
                        )
                    elif typ == "bool":
                        invalid_bool_cells.append(
                            f"{_STYLE_DISPLAY_NAMES.get(key, key)} / {_STYLE_COLUMNS[col][1]} = {text}"
                        )
            sync_style_config_indent_fields(sc)

        # 标题编号
        current_mode = str(
            getattr(self, "_num_active_mode", "custom") or "custom"
        ).strip()
        if current_mode in self._NUM_SCHEME_PRESETS:
            # Use canonical preset bindings for exact round-trip fidelity
            bindings = self._numbering_scheme_preset_bindings(current_mode)
        else:
            bindings = self._numbering_bindings_from_ui(
                error_collector=invalid_numbering_template_cells,
            )
        self._merge_numbering_catalog_source_into_config(config)
        config.heading_numbering_v2.enabled = bool(config.capabilities.get("heading_numbering", True))
        for level_name, binding in bindings.items():
            config.heading_numbering_v2.level_bindings[level_name] = copy.deepcopy(binding)
        config._heading_numbering_v2_source = "payload"  # type: ignore[attr-defined]
        config._numbering_preset_mode = current_mode  # type: ignore[attr-defined]
        invalid_numbering_template_cells.extend(
            self._sync_legacy_heading_numbering_from_v2(config, bindings)
        )

        # 标题高级
        hn = config.heading_numbering
        for attr, cb in getattr(self, "_enforcement_checks", {}).items():
            setattr(hn.enforcement, attr, cb.isChecked())
        for attr, ctrl in getattr(self, "_risk_guard_controls", {}).items():
            if isinstance(ctrl, QCheckBox):
                value = ctrl.isChecked()
            elif isinstance(ctrl, QSpinBox):
                value = ctrl.value()
            elif isinstance(ctrl, QDoubleSpinBox):
                value = ctrl.value()
            else:
                continue
            setattr(hn.risk_guard, attr, value)

        # 题注
        cap = config.caption
        cap.enabled = self._cap_enabled.isChecked()
        for attr, edit in self._cap_edits.items():
            setattr(cap, attr, edit.text())
        self._sync_caption_separator_from_ui(cap)
        # 编号格式从下拉框读取
        cap.numbering_format = str(
            self._numbering_format_combo.currentData() or "chapter.seq"
        )
        equation_table_cfg = getattr(config, "equation_table_format", None)
        if equation_table_cfg is not None:
            equation_table_cfg.numbering_format = str(
                self._equation_numbering_format_combo.currentData() or "chapter.seq"
            ).strip().lower()

        # 表格线宽
        config.table_border_width_pt = self._grid_border_width_spin.value()
        config.three_line_header_width_pt = self._three_line_header_spin.value()
        config.three_line_bottom_width_pt = self._three_line_bottom_spin.value()
        formula_style_cfg = getattr(config, "formula_style", None)
        if formula_style_cfg is not None:
            formula_style_cfg.unify_font = self._formula_table_unify_font_check.isChecked()
            formula_style_cfg.unify_size = self._formula_table_unify_size_check.isChecked()
            formula_style_cfg.unify_spacing = self._formula_table_unify_spacing_check.isChecked()

        formula_table_cfg = getattr(config, "formula_table", None)
        if formula_table_cfg is not None:
            formula_table_cfg.formula_font_name = (
                self._formula_table_font_combo.currentText().strip() or "Cambria Math"
            )
            formula_font_size_text = self._formula_table_font_size_combo.currentText().strip()
            try:
                formula_table_cfg.formula_font_size_pt = parse_font_size_input(formula_font_size_text)
                formula_table_cfg.formula_font_size_display = formula_font_size_text
            except (ValueError, TypeError):
                invalid_font_size_cells.append(f"公式表格 / 公式字号 = {formula_font_size_text}")
            formula_table_cfg.formula_line_spacing = self._formula_table_line_spacing_spin.value()
            formula_table_cfg.formula_space_before_pt = self._formula_table_space_before_spin.value()
            formula_table_cfg.formula_space_after_pt = self._formula_table_space_after_spin.value()
            formula_table_cfg.block_alignment = normalize_alignment_value(
                self._formula_table_block_alignment_combo.currentData() or "center"
            )
            formula_table_cfg.table_alignment = str(
                self._formula_table_table_alignment_combo.currentData() or "center"
            ).strip().lower() or "center"
            formula_table_cfg.formula_cell_alignment = normalize_alignment_value(
                self._formula_table_cell_alignment_combo.currentData() or "center"
            )
            formula_table_cfg.number_alignment = normalize_alignment_value(
                self._formula_table_number_alignment_combo.currentData() or "right"
            )
            formula_table_cfg.number_font_name = (
                self._formula_table_number_font_combo.currentText().strip() or "Times New Roman"
            )
            number_font_size_text = self._formula_table_number_size_combo.currentText().strip()
            try:
                formula_table_cfg.number_font_size_pt = parse_font_size_input(number_font_size_text)
                formula_table_cfg.number_font_size_display = number_font_size_text
            except (ValueError, TypeError):
                invalid_font_size_cells.append(f"公式表格 / 编号字号 = {number_font_size_text}")
            formula_table_cfg.auto_shrink_number_column = (
                self._formula_table_auto_shrink_check.isChecked()
            )

        # 页面
        ps = config.page_setup
        selected_paper = self._paper_size_combo.currentData()
        if selected_paper:
            ps.paper_size = str(selected_paper)
        margin_attrs = ("top_cm", "bottom_cm", "left_cm", "right_cm")
        for attr, spin in self._page_spins.items():
            if attr in margin_attrs:
                setattr(ps.margin, attr, spin.value())
            else:
                setattr(ps, attr, spin.value())

        # 输出
        out = config.output
        for attr, cb in getattr(self, "_output_checks", {}).items():
            setattr(out, attr, cb.isChecked())

        # 执行流程
        previous_pipeline = list(config.pipeline or [])
        selected_steps = [
            step for step in getattr(self, "_pipeline_step_order", [])
            if self._pipeline_checks.get(step) and self._pipeline_checks[step].isChecked()
        ]
        unknown_steps = [
            step for step in (config.pipeline or [])
            if step not in getattr(self, "_pipeline_checks", {})
        ]
        config.pipeline = selected_steps + unknown_steps
        if config.pipeline != previous_pipeline:
            _mark_pipeline_critical_rules_auto_managed(config)

        invalid_cells = (
            invalid_bool_cells
            + invalid_font_size_cells
            + invalid_indent_cells
            + invalid_line_spacing_cells
            + invalid_numbering_template_cells
        )
        if invalid_cells:
            preview = "\n".join(invalid_cells[:8])
            if len(invalid_cells) > 8:
                preview += f"\n... 共 {len(invalid_cells)} 处"
            tips = []
            if invalid_font_size_cells:
                tips.append("字号支持中文字号（如小四、五号）或磅值（如 12、10.5、12磅）。")
            if invalid_indent_cells:
                tips.append("缩进值必须是有效数字；单位可选字或磅。")
            if invalid_line_spacing_cells:
                tips.append("行距值必须是有效数字；固定值使用磅，多倍行距使用倍数。")
            if invalid_bool_cells:
                tips.append("布尔值可用值示例: true/false, 1/0, yes/no。")
            if invalid_numbering_template_cells:
                tips.append(
                    "标题编号外壳必须包含一个 {}；高级模板推荐占位符：{current}、{level1}...{level8}，兼容旧写法 {n}/{cn}/{parent}。"
                )
            tips_text = "\n".join(tips)
            QMessageBox.warning(
                self,
                "格式配置",
                f"以下输入格式无效，已保留原值：\n{preview}\n\n{tips_text}",
            )
            return False
        return True


class MainWindow(QMainWindow):
    """docx 一键排版 主窗口"""
    _INIT_WIDTH_RATIO = 0.35
    _INIT_HEIGHT_RATIO = 0.76
    _MAX_SCREEN_USAGE = 0.90

    def __init__(self):
        super().__init__()
        install_global_wheel_guard(QApplication.instance())
        self.setWindowTitle("Lark-Formatter V0.2")
        self.setMinimumSize(960, 760)
        self._apply_initial_window_size()
        self.setAcceptDrops(True)

        self._doc_path: str = ""
        self._config: SceneConfig | None = None
        self._worker: PipelineWorker | None = None
        self._scene_item_paths: dict[int, Path] = {}
        self._current_scene_path: str = ""
        self._current_scene_is_custom = False
        self._is_restoring_state = False
        self._suspend_scene_autosave = False
        self._clone_in_progress = False

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        self.setCentralWidget(scroll_area)

        central = QWidget()
        scroll_area.setWidget(central)
        
        self._main_layout = QVBoxLayout(central)
        self._main_layout.setSpacing(6)
        self._main_layout.setContentsMargins(16, 10, 16, 10)

        # Theme toggle header
        self._build_header_section()

        self._build_file_section()
        self._build_scene_section()
        self._build_mode_section()
        self._build_scope_section()
        self._build_lab_section()
        self._build_action_section()
        self._build_log_section()

        self._on_scene_changed(self._scene_combo.currentIndex())
        self._restore_ui_state()
        self._main_layout.addStretch(1)

    def _apply_initial_window_size(self) -> None:
        """Set initial size from available screen geometry."""
        screen = self.screen() or QApplication.primaryScreen()
        if screen is None:
            self.resize(960, 760)
            return

        avail = screen.availableGeometry()
        cap_w = int(avail.width() * self._MAX_SCREEN_USAGE)
        cap_h = int(avail.height() * self._MAX_SCREEN_USAGE)
        target_w = min(cap_w, max(960, int(avail.width() * self._INIT_WIDTH_RATIO)))
        target_h = min(cap_h, max(760, int(avail.height() * self._INIT_HEIGHT_RATIO)))
        self.resize(target_w, target_h)

    @staticmethod
    def _load_svg_icon(svg_path: Path, size: int = 20):
        """Render an SVG file into a QIcon via QSvgRenderer for reliable display."""
        from PySide6.QtGui import QIcon, QPixmap, QPainter
        from PySide6.QtSvg import QSvgRenderer
        from PySide6.QtCore import QSize
        if not svg_path.exists():
            return QIcon()
        renderer = QSvgRenderer(str(svg_path))
        pixmap = QPixmap(QSize(size, size))
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()
        return QIcon(pixmap)

    def _build_header_section(self):
        header_layout = QHBoxLayout()
        title_label = QLabel("Lark-Formatter 论文一键排版工具V0.2")
        title_label.setObjectName("SectionHeader")
        title_label.setStyleSheet("font-size: 18px;")
        header_layout.addWidget(title_label, stretch=1)

        base_dir = Path(__file__).parent

        self._about_btn = QPushButton(" 关于")
        info_icon = self._load_svg_icon(base_dir / "icons" / "info.svg")
        self._about_btn.setIcon(info_icon)
        self._about_btn.clicked.connect(self._open_about_dialog)
        header_layout.addWidget(self._about_btn)

        self._theme_toggle_btn = QPushButton(" 切换主题")
        self._update_theme_btn_icon(ThemeManager.get_current_theme())
        self._theme_toggle_btn.clicked.connect(self._toggle_theme)
        header_layout.addWidget(self._theme_toggle_btn)
        self._main_layout.addLayout(header_layout)

    def _open_about_dialog(self):
        from src.ui.about_dialog import AboutDialog
        dialog = AboutDialog(
            self,
            format_name=self._current_format_display_name(),
            format_signature=self._current_format_signature(),
            format_signable=self._current_format_signable(),
            format_signable_reason=self._current_format_signature_block_reason(),
            format_unsignable_text=self._current_format_signature_block_label(),
            on_update_format_signature=self._update_current_format_signature,
        )
        dialog.exec()

    def _current_format_display_name(self) -> str:
        if self._config is None:
            return ""

        name = str(getattr(self._config, "name", "") or "").strip()
        if name:
            return name
        if self._current_scene_path:
            return Path(self._current_scene_path).stem
        return "\u672a\u547d\u540d\u683c\u5f0f"

    def _current_format_signature(self) -> str:
        if self._config is None:
            return ""
        return str(getattr(self._config, "format_signature", "") or "").strip()

    def _current_format_signable(self) -> bool:
        return not bool(self._current_format_signature_block_reason())

    def _current_format_signature_block_label(self) -> str:
        if self._config is None:
            return "\u5f53\u524d\u683c\u5f0f\u4e0d\u53ef\u7f72\u540d"

        if not self._current_scene_path:
            return "\u672a\u4fdd\u5b58\u683c\u5f0f\u4e0d\u53ef\u7f72\u540d"

        try:
            if is_protected_scene_path(Path(self._current_scene_path)):
                return "\u9ed8\u8ba4\u683c\u5f0f\u4e0d\u53ef\u7f72\u540d"
        except Exception:
            pass

        return "\u5f53\u524d\u683c\u5f0f\u4e0d\u53ef\u7f72\u540d"

    def _current_format_signature_block_reason(self) -> str:
        if self._config is None:
            return "\u5f53\u524d\u6ca1\u6709\u53ef\u7f72\u540d\u7684\u683c\u5f0f\u914d\u7f6e\u3002"

        if not self._current_scene_path:
            return (
                "\u5f53\u524d\u683c\u5f0f\u672a\u5173\u8054\u5230\u914d\u7f6e\u6587\u4ef6\uff0c"
                "\u4e0d\u80fd\u5199\u5165\u6c38\u4e45\u683c\u5f0f\u7f72\u540d\u3002"
            )

        try:
            if is_protected_scene_path(Path(self._current_scene_path)):
                return (
                    "\u9ed8\u8ba4\u683c\u5f0f\u4e3a\u7cfb\u7edf\u57fa\u51c6\u6a21\u677f\uff0c"
                    "\u4e0d\u652f\u6301\u5199\u5165\u6c38\u4e45\u683c\u5f0f\u7f72\u540d\u3002"
                    "\u8bf7\u5148\u589e\u91cf\u4fdd\u5b58\u6216\u514b\u9686\u4e3a\u65b0\u683c\u5f0f\u540e\u518d\u7f72\u540d\u3002"
                )
        except Exception:
            pass

        return ""

    def _update_current_format_signature(self, signature: str) -> str:
        block_reason = self._current_format_signature_block_reason()
        if block_reason:
            raise RuntimeError(block_reason)

        normalized = str(signature or "").strip()
        if not normalized:
            raise ValueError("\u7f72\u540d\u4e0d\u80fd\u4e3a\u7a7a\u3002")
        if str(getattr(self._config, "format_signature", "") or "").strip():
            raise RuntimeError(
                "\u5f53\u524d\u683c\u5f0f\u5df2\u5b58\u5728\u6c38\u4e45\u7f72\u540d\uff0c\u4e0d\u80fd\u4fee\u6539\u3002"
            )

        self._config.format_signature = normalized
        self._persist_current_scene_config()
        return normalized

    def _update_theme_btn_icon(self, theme: str):
        base_dir = Path(__file__).parent
        if theme == "dark":
            icon_path = base_dir / "icons" / "sun.svg"
        else:
            icon_path = base_dir / "icons" / "moon.svg"
        self._theme_toggle_btn.setIcon(self._load_svg_icon(icon_path))

    def _toggle_theme(self):
        new_theme = ThemeManager.toggle_theme(QApplication.instance())
        self._update_theme_btn_icon(new_theme)

    def showEvent(self, event):
        super().showEvent(event)

    def _create_card(self, title: str) -> tuple[QFrame, QVBoxLayout]:
        card = QFrame()
        card.setObjectName("CardPanel")
        card.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(4)
        
        if title:
            title_lbl = QLabel(title)
            title_lbl.setObjectName("SectionHeader")
            layout.addWidget(title_lbl)
            
        return card, layout

    def _build_file_section(self):
        card, layout = self._create_card("输入文件")
        
        row_layout = QHBoxLayout()
        self._file_label = QLineEdit()
        self._file_label.setReadOnly(True)
        self._file_label.setPlaceholderText("请选择或拖入 .docx 文件...")
        row_layout.addWidget(self._file_label, stretch=1)
        btn = QPushButton("浏览...")
        btn.clicked.connect(self._browse_file)
        row_layout.addWidget(btn)
        
        layout.addLayout(row_layout)
        self._main_layout.addWidget(card)

    def _build_scene_section(self):
        card, layout = self._create_card("场景与模板")
        
        row_layout = QHBoxLayout()
        row_layout.addWidget(QLabel("预设场景:"))
        self._scene_combo = ScrollSafeComboBox()
        self._scene_combo.currentIndexChanged.connect(self._on_scene_changed)
        self._populate_scene_combo()
        row_layout.addWidget(self._scene_combo, stretch=1)
        load_btn = QPushButton("加载自定义...")
        load_btn.clicked.connect(self._load_custom_scene)
        row_layout.addWidget(load_btn)
        
        self._clone_btn = QPushButton("克隆格式...")
        self._clone_btn.setEnabled(False)
        self._clone_btn.clicked.connect(self._clone_word_template_style)
        row_layout.addWidget(self._clone_btn)
        
        self._fmt_btn = QPushButton("格式配置...")
        self._fmt_btn.setEnabled(False)
        self._fmt_btn.clicked.connect(self._open_format_config)
        row_layout.addWidget(self._fmt_btn)
        
        layout.addLayout(row_layout)
        self._main_layout.addWidget(card)

    def _populate_scene_combo(self):
        """按分类分组填充场景下拉框。"""
        self._scene_combo.blockSignals(True)
        self._scene_combo.clear()
        self._scene_item_paths.clear()

        presets = list_presets()
        if not presets:
            self._scene_combo.addItem("(无预设)")
            self._scene_combo.setItemData(0, {"kind": "empty"}, Qt.UserRole)
            self._scene_combo.blockSignals(False)
            return

        groups: dict[str, list[tuple[str, Path]]] = {}
        name_counts: dict[str, int] = {}
        for name, path, _cat, cat_label in presets:
            groups.setdefault(cat_label, []).append((name, path))
            name_counts[name] = name_counts.get(name, 0) + 1

        for cat_label in sorted(groups.keys()):
            for name, path in sorted(groups[cat_label], key=lambda x: x[0]):
                display_name = name
                if name_counts.get(name, 0) > 1:
                    display_name = f"{name} [{path.name}]"
                self._scene_combo.addItem(display_name)
                idx = self._scene_combo.count() - 1
                self._scene_combo.setItemData(idx, {"kind": "scene"}, Qt.UserRole)
                self._scene_item_paths[idx] = path

        first_scene_index = -1
        for i in range(self._scene_combo.count()):
            data = self._scene_combo.itemData(i, Qt.UserRole)
            if isinstance(data, dict) and data.get("kind") == "scene":
                first_scene_index = i
                break
        if first_scene_index >= 0:
            self._scene_combo.setCurrentIndex(first_scene_index)
        self._scene_combo.blockSignals(False)

    def _build_mode_section(self):
        card, layout = self._create_card("结构策略")
        self._mode_section_group = card
        
        row_layout = QHBoxLayout()
        row_layout.addWidget(QLabel("行为:"))
        self._mode_group = QButtonGroup(self)
        self._mode_group.setExclusive(True)
        self._radio_a = QPushButton("保留原编号，仅修复样式")
        self._radio_b = QPushButton("重建编号体系（推荐）")
        self._radio_a.setCheckable(True)
        self._radio_b.setCheckable(True)
        self._radio_a.setObjectName("SegmentedRadio")
        self._radio_b.setObjectName("SegmentedRadio")
        self._radio_b.setChecked(True)
        self._mode_group.addButton(self._radio_a, 0)
        self._mode_group.addButton(self._radio_b, 1)
        row_layout.addWidget(self._radio_a)
        row_layout.addWidget(self._radio_b)
        row_layout.addStretch(1)
        
        layout.addLayout(row_layout)
        self._main_layout.addWidget(card)

    def _build_scope_section(self):
        card, outer = self._create_card("版式与后处理")
        self._scope_group = card

        # 第一行：自动/手动模式
        self._scope_mode_row = QWidget()
        row1 = QHBoxLayout(self._scope_mode_row)
        row1.setContentsMargins(0, 0, 0, 0)
        row1.setSpacing(16)
        self._scope_mode_label = QLabel("识别模式:")
        row1.addWidget(self._scope_mode_label)
        self._scope_mode_group = QButtonGroup(self)
        self._scope_mode_group.setExclusive(True)
        self._radio_auto = QPushButton("自动识别分区")
        self._radio_manual = QPushButton("指定修正范围")
        self._radio_auto.setCheckable(True)
        self._radio_manual.setCheckable(True)
        self._radio_auto.setObjectName("SegmentedRadio")
        self._radio_manual.setObjectName("SegmentedRadio")
        self._radio_auto.setChecked(True)
        self._scope_mode_group.addButton(self._radio_auto, 0)
        self._scope_mode_group.addButton(self._radio_manual, 1)
        row1.addWidget(self._radio_auto)
        row1.addWidget(self._radio_manual)
        self._manual_inline = QWidget()
        manual_inline_layout = QHBoxLayout(self._manual_inline)
        manual_inline_layout.setContentsMargins(0, 0, 0, 0)
        manual_inline_layout.setSpacing(8)
        self._manual_page_label = QLabel("页码范围:")
        manual_inline_layout.addWidget(self._manual_page_label)
        self._page_ranges_edit = QLineEdit()
        self._page_ranges_edit.setPlaceholderText("例如：27-40,44-56")
        self._page_ranges_edit.setToolTip(
            "输入物理页码范围，支持多段，英文逗号或中文逗号分隔。示例：27-40,44-56。\n"
            "局部修正模式会跳过页面设置、全局样式、目录、页眉页脚等整篇规则。"
        )
        self._page_ranges_edit.setMinimumWidth(220)
        manual_inline_layout.addWidget(self._page_ranges_edit)
        self._manual_inline.setVisible(False)
        row1.addWidget(self._manual_inline)
        row1.addStretch(1)
        self._scope_mode_row.setMinimumHeight(40)
        outer.addWidget(self._scope_mode_row)
        
        self._scope_sep_mode = QFrame()
        self._scope_sep_mode.setFrameShape(QFrame.HLine)
        self._scope_sep_mode.setStyleSheet("color: #e1e4e8; background-color: #e1e4e8;")
        outer.addWidget(self._scope_sep_mode)

        # 第二行：分区勾选（作为统一 Panel）
        self._section_panel = QFrame()
        self._section_panel.setObjectName("SubGroupPanel")
        self._section_panel.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        section_vbox = QVBoxLayout(self._section_panel)
        section_vbox.setContentsMargins(0, 0, 12, 0)
        section_vbox.setSpacing(6)
        section_vbox.setSizeConstraint(QLayout.SetMinAndMaxSize)
        
        self._section_enable_row = QWidget()
        se_row = QHBoxLayout(self._section_enable_row)
        se_row.setContentsMargins(0, 0, 0, 0)
        self._section_enable_check = QCheckBox("排版分区")
        self._section_enable_check.setChecked(True)
        se_row.addWidget(self._section_enable_check)
        se_row.addStretch(1)
        self._section_enable_row.setFixedHeight(36)
        section_vbox.addWidget(self._section_enable_row)

        self._scope_checks = {}
        self._scope_syncing = False
        self._section_checks_layout = QHBoxLayout()
        # 增加左侧缩进对齐，并留少许上下边距
        self._section_checks_layout.setContentsMargins(24, 0, 0, 0)
        self._section_checks_layout.setSpacing(16)
        self._scope_all_check = QCheckBox("全部")
        self._section_checks_layout.addWidget(self._scope_all_check)
        
        section_options_row = QWidget()
        section_options_row.setLayout(self._section_checks_layout)
        section_vbox.addWidget(section_options_row)
        self._section_options_widget = section_options_row

        outer.addWidget(self._section_panel)

        self._rebuild_section_checkboxes([
            "body", "references", "errata", "acknowledgment", "appendix",
            "resume", "abstract_cn", "abstract_en", "toc",
        ])
        self._scope_all_check.toggled.connect(self._on_scope_all_toggled)
        
        self._section_enable_check.toggled.connect(self._on_section_enable_toggled)
        
        self._scope_sep1 = QFrame()
        self._scope_sep1.setFrameShape(QFrame.HLine)
        self._scope_sep1.setStyleSheet("color: #e1e4e8; background-color: #e1e4e8;")
        outer.addWidget(self._scope_sep1)

        # 表格管理 - 启用开关
        self._table_enable_row = QWidget()
        te_row = QHBoxLayout(self._table_enable_row)
        te_row.setContentsMargins(0, 0, 0, 0)
        self._table_enable_check = QCheckBox("表格管理")
        self._table_enable_check.setChecked(True)
        self._table_enable_check.setToolTip(
            "启用或禁用表格格式管理功能（table_format 步骤）"
        )
        te_row.addWidget(self._table_enable_check)
        te_row.addStretch(1)
        self._table_enable_row.setFixedHeight(36)

        # 表格版式（单行）
        self._table_mode_row = QWidget()
        row4 = QHBoxLayout(self._table_mode_row)
        row4.setContentsMargins(24, 0, 0, 0)
        row4.setSpacing(12)
        row4.addWidget(QLabel("宽度策略:"))
        self._table_layout_mode_combo = ScrollSafeComboBox()
        self._table_layout_mode_combo.addItem("智能总宽", "smart")
        self._table_layout_mode_combo.addItem("铺满页面", "full")
        self._table_layout_mode_combo.addItem("适宜压缩", "compact")
        self._table_layout_mode_combo.setCurrentIndex(0)
        row4.addWidget(self._table_layout_mode_combo)
        
        row4.addWidget(QLabel("智能档位:"))
        self._table_smart_levels_combo = ScrollSafeComboBox()
        self._table_smart_levels_combo.addItem("3档", 3)
        self._table_smart_levels_combo.addItem("4档", 4)
        self._table_smart_levels_combo.addItem("5档", 5)
        self._table_smart_levels_combo.addItem("6档", 6)
        self._table_smart_levels_combo.setCurrentIndex(1)
        row4.addWidget(self._table_smart_levels_combo)
        
        row4.addWidget(QLabel("边框样式:"))
        self._table_border_mode_combo = ScrollSafeComboBox()
        self._table_border_mode_combo.addItem("全框线", "full_grid")
        self._table_border_mode_combo.addItem("三线表", "three_line")
        self._table_border_mode_combo.addItem("不调整", "keep")
        self._table_border_mode_combo.setCurrentIndex(1)
        row4.addWidget(self._table_border_mode_combo)
        
        row4.addWidget(QLabel("表内行距:"))
        self._table_line_spacing_combo = ScrollSafeComboBox()
        self._table_line_spacing_combo.addItem("单倍", "single")
        self._table_line_spacing_combo.addItem("1.5倍", "one_half")
        self._table_line_spacing_combo.addItem("双倍", "double")
        self._table_line_spacing_combo.setCurrentIndex(0)
        row4.addWidget(self._table_line_spacing_combo)
        
        self._repeat_header_check = QCheckBox("跨页重复表头")
        self._repeat_header_check.setChecked(False)
        row4.addWidget(self._repeat_header_check)
        row4.addStretch(1)
        self._table_layout_mode_combo.currentIndexChanged.connect(
            self._on_table_layout_mode_changed
        )
        self._on_table_layout_mode_changed()
        
        self._table_panel = QFrame()
        self._table_panel.setObjectName("SubGroupPanel")
        self._table_panel.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)
        table_vbox = QVBoxLayout(self._table_panel)
        table_vbox.setContentsMargins(0, 0, 12, 0)
        table_vbox.setSpacing(6)
        table_vbox.setSizeConstraint(QLayout.SetMinAndMaxSize)
        table_vbox.addWidget(self._table_enable_row)
        table_vbox.addWidget(self._table_mode_row)
        outer.addWidget(self._table_panel)
        self._table_enable_check.toggled.connect(self._on_table_enable_toggled)
        self._on_table_enable_toggled(self._table_enable_check.isChecked())
        
        self._scope_sep2 = QFrame()
        self._scope_sep2.setFrameShape(QFrame.HLine)
        self._scope_sep2.setStyleSheet("color: #e1e4e8; background-color: #e1e4e8;")
        outer.addWidget(self._scope_sep2)

        self._ext_row = QWidget()
        row6 = QHBoxLayout(self._ext_row)
        row6.setContentsMargins(0, 0, 0, 0)
        row6.setSpacing(16)
        self._ext_label = QLabel("扩展选项:")
        row6.addWidget(self._ext_label)
        self._header_check = QCheckBox("自动更新页眉")
        self._header_check.setChecked(True)
        row6.addWidget(self._header_check)
        self._header_line_inline = QWidget()
        header_line_row = QHBoxLayout(self._header_line_inline)
        header_line_row.setContentsMargins(14, 0, 0, 0)
        header_line_row.setSpacing(0)
        self._header_line_check = QCheckBox("页眉横线")
        self._header_line_check.setChecked(True)
        header_line_row.addWidget(self._header_line_check)
        row6.addWidget(self._header_line_inline)
        self._pagenum_check = QCheckBox("自动更新页码")
        self._pagenum_check.setChecked(True)
        row6.addWidget(self._pagenum_check)
        self._toc_mode_label = QLabel("目录:")
        row6.addWidget(self._toc_mode_label)
        self._toc_mode_combo = ScrollSafeComboBox()
        self._toc_mode_combo.addItem("自动目录", "word_native")
        self._toc_mode_combo.addItem("普通目录", "plain")
        self._toc_mode_combo.setCurrentIndex(1)
        self._toc_mode_combo.setToolTip(
            "自动目录：插入 Word TOC 域，可直接“更新目录”；普通目录：按条目重建并保留旧版兼容行为。"
        )
        row6.addWidget(self._toc_mode_combo)
        self._auto_insert_caption_check = QCheckBox("更新表头图尾")
        self._auto_insert_caption_check.setChecked(True)
        row6.addWidget(self._auto_insert_caption_check)
        self._format_inserted_inline = QWidget()
        format_inserted_row = QHBoxLayout(self._format_inserted_inline)
        format_inserted_row.setContentsMargins(14, 0, 0, 0)
        format_inserted_row.setSpacing(0)
        self._format_inserted_check = QCheckBox("序号用域")
        self._format_inserted_check.setChecked(False)
        format_inserted_row.addWidget(self._format_inserted_check)
        row6.addWidget(self._format_inserted_inline)
        self._header_check.toggled.connect(self._on_header_update_toggled)
        self._auto_insert_caption_check.toggled.connect(self._on_auto_insert_caption_toggled)
        self._on_header_update_toggled(self._header_check.isChecked())
        self._on_auto_insert_caption_toggled(self._auto_insert_caption_check.isChecked())
        row6.addStretch(1)
        self._ext_row.setFixedHeight(36)
        outer.addWidget(self._ext_row)

        # 信号：切换自动/手动时显示/隐藏输入框
        self._scope_mode_group.idToggled.connect(self._on_scope_mode_changed)

        self._main_layout.addWidget(card)

    def _build_lab_section(self):
        card, outer = self._create_card("实验室")
        card.setObjectName("LabCardPanel")
        self._lab_group = card

        self._md_row = QWidget()
        row = QHBoxLayout(self._md_row)
        row.setContentsMargins(0, 0, 0, 0)
        self._md_cleanup_check = QCheckBox("Markdown 文本修复")
        self._md_cleanup_check.setChecked(False)
        self._md_cleanup_check.setToolTip(
            "该开关控制执行流程中的 md_cleanup 步骤；运行前会覆盖“格式配置-执行流程”中同名步骤的状态。"
        )
        row.addWidget(self._md_cleanup_check)
        row.addSpacing(8)
        self._md_inline_controls = QWidget()
        md_inline_row = QHBoxLayout(self._md_inline_controls)
        md_inline_row.setContentsMargins(0, 0, 0, 0)
        md_inline_row.setSpacing(8)
        self._md_preserve_lists_check = QCheckBox("保留列表")
        self._md_preserve_lists_check.setChecked(True)
        self._md_preserve_lists_check.setToolTip("保留原有列表标记。")
        md_inline_row.addWidget(self._md_preserve_lists_check)

        self._md_options_row = QWidget()
        md_opts_row = QHBoxLayout(self._md_options_row)
        md_opts_row.setContentsMargins(0, 0, 0, 0)
        md_opts_row.setSpacing(8)
        md_opts_row.addWidget(QLabel("间隔符："))
        self._md_list_separator_combo = ScrollSafeComboBox()
        self._md_list_separator_combo.addItem("制表符", "tab")
        self._md_list_separator_combo.addItem("半角空格", "half_space")
        self._md_list_separator_combo.addItem("全角空格", "full_space")
        self._md_list_separator_combo.setCurrentIndex(0)
        self._md_list_separator_combo.setToolTip("间隔符。")
        md_opts_row.addWidget(self._md_list_separator_combo)
        md_opts_row.addSpacing(12)
        md_opts_row.addWidget(QLabel("有序样式："))
        self._md_ordered_style_combo = ScrollSafeComboBox()
        for value, label in _MD_ORDERED_LIST_STYLE_OPTIONS:
            self._md_ordered_style_combo.addItem(label, value)
        ordered_idx = self._md_ordered_style_combo.findData("mixed")
        self._md_ordered_style_combo.setCurrentIndex(ordered_idx if ordered_idx >= 0 else 0)
        self._md_ordered_style_combo.setToolTip("有序列表样式。")
        md_opts_row.addWidget(self._md_ordered_style_combo)
        md_opts_row.addSpacing(12)
        md_opts_row.addWidget(QLabel("无序样式："))
        self._md_unordered_style_combo = ScrollSafeComboBox()
        for value, label in _MD_UNORDERED_LIST_STYLE_OPTIONS:
            self._md_unordered_style_combo.addItem(label, value)
        unordered_idx = self._md_unordered_style_combo.findData("word_default")
        self._md_unordered_style_combo.setCurrentIndex(unordered_idx if unordered_idx >= 0 else 0)
        self._md_unordered_style_combo.setToolTip("无序列表样式。")
        md_opts_row.addWidget(self._md_unordered_style_combo)
        md_inline_row.addSpacing(12)
        md_inline_row.addWidget(self._md_options_row)
        row.addWidget(self._md_inline_controls)
        row.addStretch(1)
        self._lab_help_md_btn = self._build_lab_help_button("md_cleanup")
        row.addWidget(self._lab_help_md_btn)
        self._md_row.setFixedHeight(36)
        outer.addWidget(self._md_row)

        self._md_cleanup_check.toggled.connect(self._on_md_cleanup_toggled)
        self._md_preserve_lists_check.toggled.connect(self._on_md_preserve_lists_toggled)
        self._on_md_cleanup_toggled(self._md_cleanup_check.isChecked())

        self._lab_sep0 = QFrame()
        self._lab_sep0.setFrameShape(QFrame.HLine)
        self._lab_sep0.setStyleSheet("color: #e1e4e8; background-color: #e1e4e8;")
        outer.addWidget(self._lab_sep0)

        self._whitespace_row = QWidget()
        ws_row = QHBoxLayout(self._whitespace_row)
        ws_row.setContentsMargins(0, 0, 0, 0)
        ws_row.setSpacing(0)
        self._whitespace_normalize_check = QCheckBox("文本空白清洗")
        self._whitespace_normalize_check.setChecked(False)
        self._whitespace_normalize_check.setToolTip(
            "清洗不可见异常字符（特殊空格/Tab/零宽符），并可按语境智能修正中英文标点。"
        )
        ws_row.addWidget(self._whitespace_normalize_check)

        # --- 内联子选项（跟随主开关显示/隐藏） ---
        self._ws_inline_controls = QWidget()
        ws_inline = QHBoxLayout(self._ws_inline_controls)
        ws_inline.setContentsMargins(24, 0, 0, 0)
        ws_inline.setSpacing(12)

        self._ws_cleanup_aggregate = QCheckBox("清洗异常字符")
        self._ws_cleanup_aggregate.setChecked(True)
        self._ws_cleanup_aggregate.setTristate(True)
        self._ws_cleanup_aggregate.setToolTip(
            "统一特殊空格、Tab→空格、移除零宽字符、压缩连续空格、清理段落首尾空格。"
        )
        ws_inline.addWidget(self._ws_cleanup_aggregate)

        self._ws_opt_smart_convert = QCheckBox("标点智能修正")
        self._ws_opt_smart_convert.setChecked(True)
        self._ws_opt_smart_convert.setToolTip(
            "按语境自动修正中英文标点（,→，  .→。）、括号（按内容语种）、全角英数→半角。"
        )
        ws_inline.addWidget(self._ws_opt_smart_convert)

        self._ws_opt_smart_quote = QCheckBox("引号修正")
        self._ws_opt_smart_quote.setChecked(False)
        self._ws_opt_smart_quote.setToolTip(
            "按语境修正引号（中文→""''，英文→\"\"''）。默认关闭，改动较大请谨慎启用。"
        )
        ws_inline.addWidget(self._ws_opt_smart_quote)

        self._ws_opt_protect_ref = QCheckBox("保护文献编号")
        self._ws_opt_protect_ref.setChecked(True)
        self._ws_opt_protect_ref.setToolTip(
            "避免修改 [1]、(2)、（3）等参考文献编号中的标点和括号。URL/邮箱/路径始终受保护。"
        )
        ws_inline.addWidget(self._ws_opt_protect_ref)

        self._ws_opt_context_conf_spin = ScrollSafeComboBox()
        self._ws_opt_context_conf_spin.addItem("积极", 1)
        self._ws_opt_context_conf_spin.addItem("适中", 2)
        self._ws_opt_context_conf_spin.addItem("保守", 3)
        self._ws_opt_context_conf_spin.setCurrentIndex(1)
        self._ws_opt_context_conf_spin.setToolTip(
            "标点修正的保守程度：积极=稍有语境倾向即修正，适中=推荐，保守=需明确信号才修正"
        )
        ws_inline.addWidget(self._ws_opt_context_conf_spin)

        ws_inline.addStretch(1)
        ws_row.addWidget(self._ws_inline_controls, 1)

        self._lab_help_whitespace_btn = self._build_lab_help_button("whitespace_normalize")
        ws_row.addWidget(self._lab_help_whitespace_btn)
        self._whitespace_row.setFixedHeight(36)

        # 5 个清洗子项（不暴露 UI，由聚合 checkbox 统一控制，config sync 时引用）
        self._ws_opt_space_variants = QCheckBox()
        self._ws_opt_convert_tabs = QCheckBox()
        self._ws_opt_zero_width = QCheckBox()
        self._ws_opt_collapse_spaces = QCheckBox()
        self._ws_opt_trim_edges = QCheckBox()
        self._ws_opt_space_variants.setChecked(True)
        self._ws_opt_convert_tabs.setChecked(True)
        self._ws_opt_zero_width.setChecked(True)
        self._ws_opt_collapse_spaces.setChecked(True)
        self._ws_opt_trim_edges.setChecked(True)

        # 3 个修正子项（不暴露 UI，由"标点智能修正"总开关统一控制）
        self._ws_opt_smart_punctuation = QCheckBox()
        self._ws_opt_smart_bracket = QCheckBox()
        self._ws_opt_smart_alnum = QCheckBox()
        self._ws_opt_smart_punctuation.setChecked(True)
        self._ws_opt_smart_bracket.setChecked(True)
        self._ws_opt_smart_alnum.setChecked(True)

        # 组装面板
        self._ws_panel = QFrame()
        self._ws_panel.setObjectName("SubGroupPanel")
        self._ws_panel.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
        ws_vbox = QVBoxLayout(self._ws_panel)
        ws_vbox.setContentsMargins(0, 0, 0, 0)
        ws_vbox.setSpacing(6)
        ws_vbox.addWidget(self._whitespace_row)
        outer.addWidget(self._ws_panel)

        self._lab_sep_eq = QFrame()
        self._lab_sep_eq.setFrameShape(QFrame.HLine)
        self._lab_sep_eq.setStyleSheet("color: #e1e4e8; background-color: #e1e4e8;")
        outer.addWidget(self._lab_sep_eq)

        self._formula_management_row = QWidget()
        formula_row = QHBoxLayout(self._formula_management_row)
        formula_row.setContentsMargins(0, 0, 0, 0)
        formula_row.setSpacing(0)
        self._formula_management_check = QCheckBox("公式表格管理（总开关）")
        self._formula_management_check.setChecked(False)
        self._formula_management_check.setToolTip(
            "该开关为公式相关子功能的父级开关；关闭时子项会隐藏且不执行。"
        )
        formula_row.addWidget(self._formula_management_check)

        self._formula_options_inline = QWidget()
        formula_inline_row = QHBoxLayout(self._formula_options_inline)
        formula_inline_row.setContentsMargins(24, 0, 0, 0)
        formula_inline_row.setSpacing(12)

        self._formula_convert_check = QCheckBox("公式编码统一")
        self._formula_convert_check.setChecked(False)
        self._formula_convert_check.setToolTip(
            "语义层：识别来源并统一为可编辑公式对象；不负责转表格、编号和样式。"
        )
        formula_inline_row.addWidget(self._formula_convert_check)
        self._formula_output_label = QLabel("输出格式:")
        formula_inline_row.addWidget(self._formula_output_label)
        self._formula_output_mode_combo = ScrollSafeComboBox()
        self._formula_output_mode_combo.addItem("OMML", "word_native")
        self._formula_output_mode_combo.addItem("LaTeX", "latex")
        self._formula_output_mode_combo.setCurrentIndex(0)
        self._formula_output_mode_combo.setToolTip(
            "公式编码统一输出目标：OMML（直接生成原生公式）或 LaTeX（按 LaTeX 语法转换后仍落地为可编辑原生公式）。"
        )
        formula_inline_row.addWidget(self._formula_output_mode_combo)

        self._formula_to_table_check = QCheckBox("公式转公式表格")
        self._formula_to_table_check.setChecked(False)
        self._formula_to_table_check.setToolTip(
            "结构层：仅将行间公式封装为两列表格骨架；不处理编号和样式细化。"
        )
        formula_inline_row.addWidget(self._formula_to_table_check)

        self._equation_table_check = QCheckBox("公式表格编号")
        self._equation_table_check.setChecked(False)
        self._equation_table_check.setToolTip(
            "编号层：仅做编号生成/纠正/续编；不负责公式内容转换和样式美化。"
        )
        formula_inline_row.addWidget(self._equation_table_check)

        self._formula_style_check = QCheckBox("公式格式统一")
        self._formula_style_check.setChecked(False)
        self._formula_style_check.setToolTip(
            "版式层：统一公式与公式表格视觉样式（字体、行距、列宽、对齐等）；不做语义转换和编号判定。"
        )
        formula_inline_row.addWidget(self._formula_style_check)
        formula_inline_row.addStretch(1)
        formula_row.addWidget(self._formula_options_inline, 1)

        self._lab_help_formula_management_btn = self._build_lab_help_button("formula_management")
        formula_row.addWidget(self._lab_help_formula_management_btn)
        self._formula_management_row.setFixedHeight(36)
        outer.addWidget(self._formula_management_row)

        self._lab_sep_cite = QFrame()
        self._lab_sep_cite.setFrameShape(QFrame.HLine)
        self._lab_sep_cite.setStyleSheet("color: #e1e4e8; background-color: #e1e4e8;")
        outer.addWidget(self._lab_sep_cite)

        self._citation_link_row = QWidget()
        cite_row = QHBoxLayout(self._citation_link_row)
        cite_row.setContentsMargins(0, 0, 0, 0)
        self._citation_link_check = QCheckBox("正文参考文献域关联")
        self._citation_link_check.setChecked(False)
        self._citation_link_check.setToolTip(
            "该开关控制执行流程中的 citation_link 步骤；运行前会覆盖“格式配置-执行流程”中同名步骤的状态。"
        )
        cite_row.addWidget(self._citation_link_check)
        cite_row.addSpacing(24)
        
        self._citation_enhancement_label = QLabel("增强选项:")
        cite_row.addWidget(self._citation_enhancement_label)
        self._citation_ref_auto_number_check = QCheckBox("参考文献自动编号并纠偏")
        self._citation_ref_auto_number_check.setChecked(True)
        self._citation_ref_auto_number_check.setToolTip(
            "将参考文献序号改为 SEQ 域自动编号，并让正文引用通过 REF 域同步显示，插入新文献后可整体更新纠偏。"
        )
        cite_row.addWidget(self._citation_ref_auto_number_check)
        self._citation_outer_page_sup_check = QCheckBox("方括号外页码跟随上标")
        self._citation_outer_page_sup_check.setChecked(False)
        self._citation_outer_page_sup_check.setToolTip(
            "仅在高置信场景下，将类似 [320]198 的括号外页码一并设为上标。默认关闭以避免误改普通数字。"
        )
        cite_row.addWidget(self._citation_outer_page_sup_check)
        
        cite_row.addStretch(1)
        self._lab_help_citation_btn = self._build_lab_help_button("citation_link")
        cite_row.addWidget(self._lab_help_citation_btn)
        self._citation_link_row.setFixedHeight(36)
        outer.addWidget(self._citation_link_row)

        self._lab_sep1 = QFrame()
        self._lab_sep1.setFrameShape(QFrame.HLine)
        self._lab_sep1.setStyleSheet("color: #e1e4e8; background-color: #e1e4e8;")
        outer.addWidget(self._lab_sep1)

        self._chem_row = QWidget()
        chem_row = QHBoxLayout(self._chem_row)
        chem_row.setContentsMargins(0, 0, 0, 0)
        chem_row.setSpacing(0)
        self._chem_restore_check = QCheckBox("自动恢复上下角标")
        self._chem_restore_check.setChecked(False)
        self._chem_restore_check.setToolTip("启用后，可按参考文献/正文/摘要/标题/题注/表格分别恢复化学式上下角标")
        chem_row.addWidget(self._chem_restore_check)
        self._chem_scope_inline = QWidget()
        chem_scope_inline_row = QHBoxLayout(self._chem_scope_inline)
        chem_scope_inline_row.setContentsMargins(24, 0, 0, 0)
        chem_scope_inline_row.setSpacing(12)
        self._chem_scope_label = QLabel("作用范围:")
        chem_scope_inline_row.addWidget(self._chem_scope_label)
        self._chem_scope_all_check = QCheckBox("全部")
        self._chem_scope_refs_check = QCheckBox("参考文献")
        self._chem_scope_body_check = QCheckBox("正文")
        self._chem_scope_abstract_check = QCheckBox("摘要/Abstract")
        self._chem_scope_headings_check = QCheckBox("标题")
        self._chem_scope_captions_check = QCheckBox("题注")
        self._chem_scope_tables_check = QCheckBox("表格")
        self._chem_scope_refs_check.setChecked(False)
        self._chem_scope_body_check.setChecked(False)
        self._chem_scope_abstract_check.setChecked(False)
        self._chem_scope_headings_check.setChecked(False)
        self._chem_scope_captions_check.setChecked(False)
        self._chem_scope_tables_check.setChecked(False)
        self._chem_scope_refs_check.setToolTip("对参考文献段落中的化学式尝试恢复上下角标")
        self._chem_scope_body_check.setToolTip("对正文普通段落中的化学式尝试恢复上下角标")
        self._chem_scope_abstract_check.setToolTip("对中文摘要与英文摘要中的化学式尝试恢复上下角标")
        self._chem_scope_headings_check.setToolTip("对标题中的化学式尝试恢复上下角标")
        self._chem_scope_captions_check.setToolTip("对图题、表题等题注中的化学式尝试恢复上下角标")
        self._chem_scope_tables_check.setToolTip("对表格表头/单元格中的化学式尝试恢复上下角标")
        chem_scope_inline_row.addWidget(self._chem_scope_all_check)
        chem_scope_inline_row.addWidget(self._chem_scope_refs_check)
        chem_scope_inline_row.addWidget(self._chem_scope_body_check)
        chem_scope_inline_row.addWidget(self._chem_scope_abstract_check)
        chem_scope_inline_row.addWidget(self._chem_scope_headings_check)
        chem_scope_inline_row.addWidget(self._chem_scope_captions_check)
        chem_scope_inline_row.addWidget(self._chem_scope_tables_check)
        chem_scope_inline_row.addStretch(1)
        chem_row.addWidget(self._chem_scope_inline, 1)
        self._lab_help_chem_btn = self._build_lab_help_button("chem_typography")
        chem_row.addWidget(self._lab_help_chem_btn)

        self._chem_row.setFixedHeight(36)

        self._chem_panel = QFrame()
        self._chem_panel.setObjectName("SubGroupPanel")
        self._chem_panel.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
        chem_vbox = QVBoxLayout(self._chem_panel)
        chem_vbox.setContentsMargins(0, 0, 0, 0)
        chem_vbox.setSpacing(6)
        chem_vbox.addWidget(self._chem_row)
        outer.addWidget(self._chem_panel)

        self._chem_scope_syncing = False
        self._whitespace_normalize_check.toggled.connect(self._on_whitespace_normalize_toggled)
        self._ws_opt_smart_convert.toggled.connect(self._on_whitespace_smart_convert_toggled)
        self._ws_cleanup_aggregate.clicked.connect(self._on_ws_cleanup_aggregate_clicked)
        self._ws_opt_space_variants.toggled.connect(self._sync_ws_cleanup_aggregate_state)
        self._ws_opt_convert_tabs.toggled.connect(self._sync_ws_cleanup_aggregate_state)
        self._ws_opt_zero_width.toggled.connect(self._sync_ws_cleanup_aggregate_state)
        self._ws_opt_collapse_spaces.toggled.connect(self._sync_ws_cleanup_aggregate_state)
        self._ws_opt_trim_edges.toggled.connect(self._sync_ws_cleanup_aggregate_state)

        self._formula_management_check.toggled.connect(self._on_formula_management_toggled)
        self._formula_convert_check.toggled.connect(self._on_formula_convert_toggled)
        self._citation_link_check.toggled.connect(self._on_citation_link_toggled)
        self._chem_restore_check.toggled.connect(self._on_chem_restore_toggled)
        self._chem_scope_all_check.toggled.connect(self._on_chem_scope_all_toggled)
        self._chem_scope_refs_check.toggled.connect(self._sync_chem_scope_all_state)
        self._chem_scope_body_check.toggled.connect(self._sync_chem_scope_all_state)
        self._chem_scope_abstract_check.toggled.connect(self._sync_chem_scope_all_state)
        self._chem_scope_headings_check.toggled.connect(self._sync_chem_scope_all_state)
        self._chem_scope_captions_check.toggled.connect(self._sync_chem_scope_all_state)
        self._chem_scope_tables_check.toggled.connect(self._sync_chem_scope_all_state)
        self._sync_chem_scope_all_state()
        self._on_whitespace_normalize_toggled(self._whitespace_normalize_check.isChecked())
        self._on_whitespace_smart_convert_toggled(self._ws_opt_smart_convert.isChecked())
        self._on_formula_management_toggled(self._formula_management_check.isChecked())
        self._on_formula_convert_toggled(self._formula_convert_check.isChecked())
        self._on_citation_link_toggled(self._citation_link_check.isChecked())
        self._on_chem_restore_toggled(self._chem_restore_check.isChecked())

        self._main_layout.addWidget(card)

    def _build_lab_help_button(self, topic: str) -> QPushButton:
        btn = QPushButton("!")
        btn.setObjectName("LabHelpButton")
        btn.setFixedSize(20, 20)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setStyleSheet(
            "QPushButton#LabHelpButton {"
            " border: 1.5px solid #e3a936;"
            " border-radius: 10px;"
            " background: transparent;"
            " color: #e3a936;"
            " font-size: 13px;"
            " font-family: 'Segoe UI', Arial, sans-serif;"
            " font-weight: bold;"
            " padding: 0px;"
            " margin: 0px;"
            "}"
            "QPushButton#LabHelpButton:hover {"
            " background: rgba(227, 169, 54, 0.1);"
            "}"
            "QPushButton#LabHelpButton:pressed {"
            " background: rgba(227, 169, 54, 0.2);"
            "}"
        )
        btn.setToolTip("查看风险、功能与使用方法")
        btn.clicked.connect(lambda _=False, t=topic: self._show_lab_help_dialog(t))
        return btn

    def _show_lab_help_dialog(self, topic: str):
        docs = {
            "md_cleanup": {
                "title": "Markdown 文本修复 使用须知",
                "status": "开发状态：已完成并通过回归测试",
                "feature": "实际功能：修复 Word 中残留的 Markdown 文本，并支持列表样式与分隔符规范化",
                "note": "注意事项：仅在文档存在 Markdown 残留时开启，避免无关改动",
            },
            "whitespace_normalize": {
                "title": "文本空白清洗 使用须知",
                "status": "开发状态：已完成并通过回归测试",
                "feature": "实际功能：清洗异常空白字符，并支持中英文标点与符号规范化",
                "note": "注意事项：不调整段落换行结构；启用标点修正后建议复核关键内容",
            },
            "formula_convert": {
                "title": "公式编码统一 使用须知",
                "status": "开发状态：MVP框架已接入，核心流程已打通，细节策略持续完善中",
                "feature": "实际功能：识别多来源公式并转换到统一中间表示；LaTeX 模式按 LaTeX 语法执行转换，但最终仍生成可编辑原生公式对象",
                "note": "注意事项：低置信公式默认跳过并标记，不自动替换",
            },
            "formula_to_table": {
                "title": "公式转公式表格 使用须知",
                "status": "开发状态：MVP框架已接入，规则持续完善中",
                "feature": "实际功能：将块级公式封装为公式表格骨架（容器化），便于后续编号与版式规则独立接力",
                "note": "注意事项：默认仅处理块级公式，行内公式不会自动转换",
            },
            "equation_table_format": {
                "title": "公式表格编号 使用须知",
                "status": "开发状态：开发完成、测试通过，场景受限",
                "feature": "实际功能：自动生成/纠正/续编公式编号，不承担公式表格样式统一",
                "note": "注意事项：当前主要识别两列表格场景，左侧公式右侧序号",
            },
            "formula_style": {
                "title": "公式格式统一 使用须知",
                "status": "开发状态：MVP框架已接入，细节策略持续完善中",
                "feature": "实际功能：统一公式与公式表格视觉样式（字体、字号、行距、列宽、对齐、边框）",
                "note": "注意事项：建议先在样本文档验证后再批量处理",
            },
            "formula_management": {
                "title": "公式表格管理（总开关） 使用须知",
                "status": "开发状态：父子开关已拆分，建议先小样本验证",
                "feature": "实际功能：统一启停公式编码统一、公式转公式表格、公式表格编号、公式格式统一",
                "note": "注意事项：关闭总开关时将隐藏子功能并禁用对应流程步骤",
            },
            "citation_link": {
                "title": "正文参考文献域关联 使用须知",
                "status": "开发状态：已完成并通过回归测试",
                "feature": "实际功能：将正文中的编号引用转换为可跳转域，并支持参考文献序号自动纠偏",
                "note": "注意事项：主要识别常见编号引用格式，已有域代码内容会自动跳过",
            },
            "chem_typography": {
                "title": "自动恢复上下角标 使用须知",
                "status": "开发状态：已完成并通过回归测试",
                "feature": "实际功能：自动恢复导入或转换后丢失的上下角标格式，并可分别作用于参考文献、正文、摘要、标题、题注、表格",
                "note": "注意事项：现支持按参考文献、正文、摘要、标题、题注、表格分别控制；复杂公式、缩写或特殊写法建议人工复核",
            },
        }
        payload = docs.get(topic) or {
            "title": "实验室功能 使用须知",
            "status": "开发状态：实验性能力，默认谨慎使用",
            "feature": "实际功能：用于增强自动化修复能力",
            "note": "注意事项：建议先小样本验证后再批量应用",
        }
        dialog = QDialog(self)
        dialog.setWindowTitle(payload["title"])
        dialog.setModal(True)
        dialog.setMinimumWidth(480)
        
        root = QVBoxLayout(dialog)
        root.setContentsMargins(28, 28, 28, 24)
        root.setSpacing(16)
        
        title_lbl = QLabel(payload["title"])
        title_lbl.setStyleSheet("font-size: 19px; font-weight: bold; color: #d0a24a; margin-bottom: 8px;")
        root.addWidget(title_lbl)
        
        def _add_block(title_text: str, content_text: str, color_hex: str, icon_char: str):
            prefix = f"{title_text}："
            tail = content_text[len(prefix):].strip() if content_text.startswith(prefix) else content_text.strip()
            
            frame = QFrame()
            r, g, b = int(color_hex[1:3], 16), int(color_hex[3:5], 16), int(color_hex[5:7], 16)
            frame.setStyleSheet(f"""
                QFrame {{
                    background-color: rgba({r}, {g}, {b}, 0.05);
                    border: 1px solid rgba({r}, {g}, {b}, 0.15);
                    border-left: 4px solid {color_hex};
                    border-radius: 8px;
                }}
            """)
            
            h_layout = QHBoxLayout(frame)
            h_layout.setContentsMargins(12, 14, 16, 14)
            h_layout.setSpacing(10)
            
            icon_lbl = QLabel(icon_char)
            icon_lbl.setStyleSheet(f"font-size: 24px; color: {color_hex}; background: transparent; border: none;")
            icon_lbl.setAlignment(Qt.AlignCenter)
            icon_lbl.setFixedWidth(32)
            h_layout.addWidget(icon_lbl, 0, Qt.AlignVCenter)
            
            v_layout = QVBoxLayout()
            v_layout.setSpacing(4)
            
            h_lbl = QLabel(title_text)
            h_lbl.setStyleSheet(f"font-size: 14px; font-weight: bold; color: {color_hex}; background: transparent; border: none;")
            v_layout.addWidget(h_lbl)
            
            c_lbl = QLabel(tail)
            c_lbl.setWordWrap(True)
            c_lbl.setStyleSheet("font-size: 13px; color: #555555; background: transparent; border: none; line-height: 1.4;")
            v_layout.addWidget(c_lbl)
            
            h_layout.addLayout(v_layout, 1)
            root.addWidget(frame)

        _add_block("开发状态", payload["status"], "#d0a24a", "●")
        _add_block("实际功能", payload["feature"], "#d0a24a", "✦")
        _add_block("注意事项", payload["note"], "#d0a24a", "！")
        
        root.addStretch(1)
        
        button_row = QHBoxLayout()
        button_row.addStretch(1)
        ok_btn = QPushButton("知道了")
        ok_btn.setMinimumWidth(80)
        ok_btn.setCursor(Qt.PointingHandCursor)
        ok_btn.setDefault(True)
        ok_btn.clicked.connect(dialog.accept)
        button_row.addWidget(ok_btn)
        root.addLayout(button_row)
        
        dialog.adjustSize()
        dialog.exec()

    def _sync_lab_controls_from_config(self):
        if not self._config:
            return

        pipeline = list(getattr(self._config, "pipeline", []) or [])
        md_cfg = getattr(self._config, "md_cleanup", None)
        ws_cfg = getattr(self._config, "whitespace_normalize", None)
        formula_convert_cfg = getattr(self._config, "formula_convert", None)
        formula_to_table_cfg = getattr(self._config, "formula_to_table", None)
        equation_table_cfg = getattr(self._config, "equation_table_format", None)
        formula_style_cfg = getattr(self._config, "formula_style", None)

        self._md_cleanup_check.setChecked(
            ("md_cleanup" in pipeline) or bool(getattr(md_cfg, "enabled", False))
        )
        self._md_preserve_lists_check.setChecked(
            bool(getattr(md_cfg, "preserve_existing_word_lists", True))
        )
        md_list_separator = str(getattr(md_cfg, "list_marker_separator", "tab")).strip().lower()
        if md_list_separator not in {"tab", "half_space", "full_space"}:
            md_list_separator = "tab"
        sep_idx = self._md_list_separator_combo.findData(md_list_separator)
        self._md_list_separator_combo.setCurrentIndex(sep_idx if sep_idx >= 0 else 0)
        ordered_styles = {v for v, _ in _MD_ORDERED_LIST_STYLE_OPTIONS}
        unordered_styles = {v for v, _ in _MD_UNORDERED_LIST_STYLE_OPTIONS}
        md_ordered_style = str(getattr(md_cfg, "ordered_list_style", "mixed")).strip().lower()
        if md_ordered_style not in ordered_styles:
            md_ordered_style = "mixed"
        md_ordered_idx = self._md_ordered_style_combo.findData(md_ordered_style)
        self._md_ordered_style_combo.setCurrentIndex(md_ordered_idx if md_ordered_idx >= 0 else 0)
        md_unordered_style = str(getattr(md_cfg, "unordered_list_style", "word_default")).strip().lower()
        if md_unordered_style not in unordered_styles:
            md_unordered_style = "word_default"
        md_unordered_idx = self._md_unordered_style_combo.findData(md_unordered_style)
        self._md_unordered_style_combo.setCurrentIndex(md_unordered_idx if md_unordered_idx >= 0 else 0)
        self._whitespace_normalize_check.setChecked(
            ("whitespace_normalize" in pipeline)
            or bool(getattr(ws_cfg, "enabled", False))
        )
        self._formula_convert_check.setChecked(
            ("formula_convert" in pipeline)
            or bool(getattr(formula_convert_cfg, "enabled", False))
        )
        self._formula_to_table_check.setChecked(
            ("formula_to_table" in pipeline)
            or bool(getattr(formula_to_table_cfg, "enabled", False))
        )
        self._equation_table_check.setChecked(
            ("equation_table_format" in pipeline)
            or bool(getattr(equation_table_cfg, "enabled", False))
        )
        self._formula_style_check.setChecked(
            ("formula_style" in pipeline)
            or bool(getattr(formula_style_cfg, "enabled", False))
        )
        self._formula_management_check.setChecked(
            self._formula_convert_check.isChecked()
            or self._formula_to_table_check.isChecked()
            or self._equation_table_check.isChecked()
            or self._formula_style_check.isChecked()
        )
        output_mode = str(
            getattr(formula_convert_cfg, "output_mode", "word_native")
        ).strip().lower()
        output_idx = self._formula_output_mode_combo.findData(
            output_mode if output_mode in {"word_native", "latex"} else "word_native"
        )
        self._formula_output_mode_combo.setCurrentIndex(output_idx if output_idx >= 0 else 0)
        self._on_formula_management_toggled(self._formula_management_check.isChecked())
        self._on_formula_convert_toggled(self._formula_convert_check.isChecked())
        self._citation_link_check.setChecked("citation_link" in pipeline)
        cite_cfg = getattr(self._config, "citation_link", None)
        if cite_cfg is not None:
            self._citation_ref_auto_number_check.setChecked(
                bool(getattr(cite_cfg, "auto_number_reference_entries", True))
            )
            self._citation_outer_page_sup_check.setChecked(
                bool(getattr(cite_cfg, "superscript_outer_page_numbers", False))
            )
        if ws_cfg is not None:
            self._ws_opt_space_variants.setChecked(
                bool(getattr(ws_cfg, "normalize_space_variants", True))
            )
            self._ws_opt_convert_tabs.setChecked(
                bool(getattr(ws_cfg, "convert_tabs", True))
            )
            self._ws_opt_zero_width.setChecked(
                bool(getattr(ws_cfg, "remove_zero_width", True))
            )
            self._ws_opt_collapse_spaces.setChecked(
                bool(getattr(ws_cfg, "collapse_multiple_spaces", True))
            )
            self._ws_opt_trim_edges.setChecked(
                bool(getattr(ws_cfg, "trim_paragraph_edges", True))
            )
            self._ws_opt_smart_convert.setChecked(
                bool(getattr(ws_cfg, "smart_full_half_convert", True))
            )
            self._ws_opt_smart_punctuation.setChecked(
                bool(getattr(ws_cfg, "punctuation_by_context", True))
            )
            self._ws_opt_smart_bracket.setChecked(
                bool(getattr(ws_cfg, "bracket_by_inner_language", True))
            )
            self._ws_opt_smart_alnum.setChecked(
                bool(getattr(ws_cfg, "fullwidth_alnum_to_halfwidth", True))
            )
            self._ws_opt_smart_quote.setChecked(
                bool(getattr(ws_cfg, "quote_by_context", False))
            )
            self._ws_opt_protect_ref.setChecked(
                bool(getattr(ws_cfg, "protect_reference_numbering", True))
            )
            try:
                conf = int(getattr(ws_cfg, "context_min_confidence", 2))
            except (TypeError, ValueError):
                conf = 2
            conf_idx = self._ws_opt_context_conf_spin.findData(max(1, min(conf, 3)))
            self._ws_opt_context_conf_spin.setCurrentIndex(conf_idx if conf_idx >= 0 else 1)
        self._on_whitespace_normalize_toggled(self._whitespace_normalize_check.isChecked())
        self._on_whitespace_smart_convert_toggled(self._ws_opt_smart_convert.isChecked())
        self._on_citation_link_toggled(self._citation_link_check.isChecked())
        self._on_md_cleanup_toggled(self._md_cleanup_check.isChecked())

        chem_cfg = getattr(self._config, "chem_typography", None)
        if chem_cfg is not None:
            scopes = getattr(chem_cfg, "scopes", {}) or {}
            self._chem_scope_syncing = True
            self._chem_scope_refs_check.setChecked(bool(scopes.get("references", True)))
            self._chem_scope_body_check.setChecked(bool(scopes.get("body", False)))
            self._chem_scope_abstract_check.setChecked(
                bool(scopes.get("abstract", False))
                or bool(scopes.get("abstract_cn", False))
                or bool(scopes.get("abstract_en", False))
            )
            self._chem_scope_headings_check.setChecked(bool(scopes.get("headings", False)))
            self._chem_scope_captions_check.setChecked(bool(scopes.get("captions", False)))
            self._chem_scope_tables_check.setChecked(bool(scopes.get("tables", False)))
            self._chem_scope_syncing = False
            self._sync_chem_scope_all_state()
            self._chem_restore_check.setChecked(bool(getattr(chem_cfg, "enabled", True)))
        self._on_chem_restore_toggled(self._chem_restore_check.isChecked())

    def _on_md_cleanup_toggled(self, checked: bool):
        visible = bool(checked)
        self._md_inline_controls.setVisible(visible)
        self._on_md_preserve_lists_toggled(self._md_preserve_lists_check.isChecked())

    def _on_md_preserve_lists_toggled(self, checked: bool):
        self._md_options_row.setVisible(
            self._md_cleanup_check.isChecked() and bool(checked)
        )

    def _on_whitespace_normalize_toggled(self, checked: bool):
        self._ws_inline_controls.setVisible(checked)

    def _on_table_enable_toggled(self, checked: bool):
        self._table_mode_row.setVisible(checked)
        margins = self._table_panel.layout().contentsMargins()
        margins.setBottom(8 if checked else 0)
        self._table_panel.layout().setContentsMargins(margins)

    def _on_header_update_toggled(self, checked: bool):
        # Use explicit hidden-state instead of isVisible():
        # during startup widgets may be not yet shown, but should still reveal children.
        visible = bool(checked) and (not self._header_check.isHidden())
        self._header_line_inline.setVisible(visible)
        self._header_line_check.setEnabled(visible)

    def _on_auto_insert_caption_toggled(self, checked: bool):
        # Same rationale as header child option visibility.
        visible = bool(checked) and (not self._auto_insert_caption_check.isHidden())
        self._format_inserted_inline.setVisible(visible)
        self._format_inserted_check.setEnabled(visible)


    def _on_whitespace_smart_convert_toggled(self, checked: bool):
        # Sync hidden sub-items to follow smart convert toggle
        self._ws_opt_smart_punctuation.setChecked(checked)
        self._ws_opt_smart_bracket.setChecked(checked)
        self._ws_opt_smart_alnum.setChecked(checked)
        # Enable/disable visible inline controls that depend on smart convert
        self._ws_opt_smart_quote.setEnabled(checked)
        self._ws_opt_protect_ref.setEnabled(checked)
        self._ws_opt_context_conf_spin.setEnabled(checked)

    def _on_ws_cleanup_aggregate_clicked(self):
        """User clicked the aggregate checkbox – set all sub-items to on or off."""
        # After a user click on a tristate checkbox, Qt cycles through
        # Unchecked → PartiallyChecked → Checked.  We only want two
        # user-visible states: Checked (all on) and Unchecked (all off).
        target = self._ws_cleanup_aggregate.checkState() != Qt.Unchecked
        self._ws_opt_space_variants.setChecked(target)
        self._ws_opt_convert_tabs.setChecked(target)
        self._ws_opt_zero_width.setChecked(target)
        self._ws_opt_collapse_spaces.setChecked(target)
        self._ws_opt_trim_edges.setChecked(target)
        # Force fully checked / unchecked (no partial after user click)
        self._ws_cleanup_aggregate.setCheckState(
            Qt.Checked if target else Qt.Unchecked
        )

    def _sync_ws_cleanup_aggregate_state(self, _=None):
        """Reflect sub-item states into the aggregate tri-state checkbox."""
        checks = [
            self._ws_opt_space_variants.isChecked(),
            self._ws_opt_convert_tabs.isChecked(),
            self._ws_opt_zero_width.isChecked(),
            self._ws_opt_collapse_spaces.isChecked(),
            self._ws_opt_trim_edges.isChecked(),
        ]
        if all(checks):
            self._ws_cleanup_aggregate.setCheckState(Qt.Checked)
        elif any(checks):
            self._ws_cleanup_aggregate.setCheckState(Qt.PartiallyChecked)
        else:
            self._ws_cleanup_aggregate.setCheckState(Qt.Unchecked)



    def _on_formula_management_toggled(self, checked: bool):
        self._formula_options_inline.setVisible(checked)
        self._formula_convert_check.setEnabled(checked)
        self._formula_to_table_check.setEnabled(checked)
        self._equation_table_check.setEnabled(checked)
        self._formula_style_check.setEnabled(checked)
        self._on_formula_convert_toggled(self._formula_convert_check.isChecked())

    def _on_formula_convert_toggled(self, checked: bool):
        enable_output = bool(checked) and self._formula_management_check.isChecked()
        self._formula_output_label.setEnabled(enable_output)
        self._formula_output_mode_combo.setEnabled(enable_output)

    def _on_citation_link_toggled(self, checked: bool):
        self._citation_ref_auto_number_check.setVisible(checked)
        self._citation_outer_page_sup_check.setVisible(checked)
        self._citation_enhancement_label.setVisible(checked)
        self._citation_ref_auto_number_check.setEnabled(checked)
        self._citation_outer_page_sup_check.setEnabled(checked)

    def _on_chem_restore_toggled(self, checked: bool):
        self._chem_scope_inline.setVisible(checked)
        self._chem_scope_label.setEnabled(checked)
        self._chem_scope_all_check.setEnabled(checked)
        self._chem_scope_refs_check.setEnabled(checked)
        self._chem_scope_body_check.setEnabled(checked)
        self._chem_scope_abstract_check.setEnabled(checked)
        self._chem_scope_headings_check.setEnabled(checked)
        self._chem_scope_captions_check.setEnabled(checked)
        self._chem_scope_tables_check.setEnabled(checked)

    def _on_section_enable_toggled(self, checked: bool):
        self._section_options_widget.setVisible(checked)
        margins = self._section_panel.layout().contentsMargins()
        margins.setBottom(8 if checked else 0)
        self._section_panel.layout().setContentsMargins(margins)

    def _on_scope_all_toggled(self, checked: bool):
        if self._scope_syncing:
            return
        self._scope_syncing = True
        for cb in self._scope_checks.values():
            cb.setChecked(checked)
        self._scope_syncing = False
        self._sync_scope_all_state()

    def _sync_scope_all_state(self):
        if self._scope_syncing:
            return
        all_checked = bool(self._scope_checks) and all(
            cb.isChecked() for cb in self._scope_checks.values()
        )
        self._scope_syncing = True
        self._scope_all_check.setChecked(all_checked)
        self._scope_syncing = False

    def _on_chem_scope_all_toggled(self, checked: bool):
        if self._chem_scope_syncing:
            return
        self._chem_scope_syncing = True
        self._chem_scope_refs_check.setChecked(checked)
        self._chem_scope_body_check.setChecked(checked)
        self._chem_scope_abstract_check.setChecked(checked)
        self._chem_scope_headings_check.setChecked(checked)
        self._chem_scope_captions_check.setChecked(checked)
        self._chem_scope_tables_check.setChecked(checked)
        self._chem_scope_syncing = False
        self._sync_chem_scope_all_state()

    def _sync_chem_scope_all_state(self):
        if self._chem_scope_syncing:
            return
        all_checked = (
            self._chem_scope_refs_check.isChecked()
            and self._chem_scope_body_check.isChecked()
            and self._chem_scope_abstract_check.isChecked()
            and self._chem_scope_headings_check.isChecked()
            and self._chem_scope_captions_check.isChecked()
            and self._chem_scope_tables_check.isChecked()
        )
        self._chem_scope_syncing = True
        self._chem_scope_all_check.setChecked(all_checked)
        self._chem_scope_syncing = False

    def _has_valid_doc_path(self) -> bool:
        if not self._doc_path:
            return False
        try:
            return Path(self._doc_path).exists()
        except OSError:
            return False

    def _collect_config_snapshot(self) -> dict | None:
        if not self._config:
            return None
        try:
            return asdict(self._config)
        except Exception as e:
            self._log(f"缓存配置快照失败: {e}")
            return None

    def _ui_state_path(self) -> Path:
        appdata = os.getenv("APPDATA")
        if appdata:
            root = Path(appdata) / "Lark-Formatter"
            legacy_root = Path(appdata) / "DOCXFormatter"
        else:
            root = Path.home() / ".lark_formatter"
            legacy_root = Path.home() / ".docx_formatter"
        root.mkdir(parents=True, exist_ok=True)
        state_path = root / "ui_state.json"
        legacy_state_path = legacy_root / "ui_state.json"
        # Migrate legacy state once to keep user preferences after renaming.
        if (not state_path.exists()) and legacy_state_path.exists():
            try:
                shutil.copy2(legacy_state_path, state_path)
            except Exception:
                pass
        return state_path

    def _collect_ui_state(self) -> dict:
        scene_path, scene_is_custom = self._normalized_scene_state_for_storage()
        return {
            "doc_path": self._doc_path,
            "scene_path": scene_path,
            "scene_is_custom": scene_is_custom,
            "config_snapshot": self._collect_config_snapshot(),
            "controls": {
                "mode": "A" if self._mode_group.checkedId() == 0 else "B",
                "scope_mode": "manual" if self._scope_mode_group.checkedId() == 1 else "auto",
                "page_ranges_text": self._page_ranges_edit.text().strip(),
                "section_enabled": self._section_enable_check.isChecked(),
                "sections": {k: cb.isChecked() for k, cb in self._scope_checks.items()},
                "table_enabled": self._table_enable_check.isChecked(),
                "table_layout_mode": self._table_layout_mode_combo.currentData(),
                "table_smart_levels": self._table_smart_levels_combo.currentData(),
                "table_border_mode": self._table_border_mode_combo.currentData(),
                "table_line_spacing_mode": self._table_line_spacing_combo.currentData(),
                "table_repeat_header": self._repeat_header_check.isChecked(),
                "md_cleanup": self._md_cleanup_check.isChecked(),
                "md_preserve_lists": self._md_preserve_lists_check.isChecked(),
                "md_list_marker_separator": self._md_list_separator_combo.currentData(),
                "md_ordered_list_style": self._md_ordered_style_combo.currentData(),
                "md_unordered_list_style": self._md_unordered_style_combo.currentData(),
                "whitespace_normalize": self._whitespace_normalize_check.isChecked(),
                "whitespace_options": {
                    "normalize_space_variants": self._ws_opt_space_variants.isChecked(),
                    "convert_tabs": self._ws_opt_convert_tabs.isChecked(),
                    "remove_zero_width": self._ws_opt_zero_width.isChecked(),
                    "collapse_multiple_spaces": self._ws_opt_collapse_spaces.isChecked(),
                    "trim_paragraph_edges": self._ws_opt_trim_edges.isChecked(),
                    "smart_full_half_convert": self._ws_opt_smart_convert.isChecked(),
                    "punctuation_by_context": self._ws_opt_smart_punctuation.isChecked(),
                    "bracket_by_inner_language": self._ws_opt_smart_bracket.isChecked(),
                    "fullwidth_alnum_to_halfwidth": self._ws_opt_smart_alnum.isChecked(),
                    "quote_by_context": self._ws_opt_smart_quote.isChecked(),
                    "protect_reference_numbering": self._ws_opt_protect_ref.isChecked(),
                    "context_min_confidence": self._ws_opt_context_conf_spin.currentData() or 2,
                },
                "formula_management": self._formula_management_check.isChecked(),
                "formula_convert": self._formula_convert_check.isChecked(),
                "formula_convert_output_mode": self._formula_output_mode_combo.currentData(),
                "formula_to_table": self._formula_to_table_check.isChecked(),
                "equation_table_format": self._equation_table_check.isChecked(),
                "formula_style": self._formula_style_check.isChecked(),
                "update_header": self._header_check.isChecked(),
                "update_page_number": self._pagenum_check.isChecked(),
                "toc_mode": self._toc_mode_combo.currentData(),
                "update_header_line": self._header_line_check.isChecked(),
                "auto_insert_caption": self._auto_insert_caption_check.isChecked(),
                "format_inserted_caption": self._format_inserted_check.isChecked(),
                "chem_restore": self._chem_restore_check.isChecked(),
                "chem_scopes": {
                    "references": self._chem_scope_refs_check.isChecked(),
                    "body": self._chem_scope_body_check.isChecked(),
                    "abstract": self._chem_scope_abstract_check.isChecked(),
                    "headings": self._chem_scope_headings_check.isChecked(),
                    "captions": self._chem_scope_captions_check.isChecked(),
                    "tables": self._chem_scope_tables_check.isChecked(),
                },
            },
        }

    def _find_preset_index(self, scene_path: Path) -> int | None:
        """Find preset combo index by absolute path first, then by filename."""
        if not scene_path:
            return None

        try:
            target_resolved = str(scene_path.resolve())
        except Exception:
            target_resolved = str(scene_path)
        target_name = scene_path.name.lower()

        for idx, preset_path in self._scene_item_paths.items():
            cur_path = Path(preset_path)
            try:
                cur_resolved = str(cur_path.resolve())
            except Exception:
                cur_resolved = str(cur_path)
            if cur_resolved == target_resolved:
                return idx

        if target_name:
            for idx, preset_path in self._scene_item_paths.items():
                if Path(preset_path).name.lower() == target_name:
                    return idx
        return None

    def _canonicalize_scene_path(self, scene_path: Path, *, migrate: bool) -> Path:
        """Normalize legacy source preset path to runtime PRESETS_DIR in frozen mode."""
        p = Path(scene_path)
        if not (getattr(sys, "frozen", False) and self._looks_like_legacy_source_preset_path(p)):
            return p

        target = PRESETS_DIR / p.name
        if migrate and p.exists():
            try:
                target.parent.mkdir(parents=True, exist_ok=True)
                if not target.exists():
                    shutil.copy2(p, target)
            except Exception as e:
                self._log(f"迁移历史源码模板失败: {e}")
        return target

    @staticmethod
    def _looks_like_legacy_source_preset_path(scene_path: Path) -> bool:
        parts = [p.lower() for p in scene_path.parts]
        return (
            scene_path.suffix.lower() == ".json"
            and "src" in parts
            and "scene" in parts
            and "presets" in parts
        )

    def _normalized_scene_state_for_storage(self) -> tuple[str, bool]:
        scene_path = str(self._current_scene_path or "")
        scene_is_custom = bool(self._current_scene_is_custom)
        if not scene_path:
            return scene_path, scene_is_custom

        p = Path(scene_path)
        preset_idx = self._find_preset_index(p)
        if preset_idx is None:
            if (
                scene_is_custom
                and getattr(sys, "frozen", False)
                and self._looks_like_legacy_source_preset_path(p)
            ):
                canonical = self._canonicalize_scene_path(p, migrate=True)
                return str(canonical), True
            return scene_path, scene_is_custom

        mapped = self._scene_item_paths.get(preset_idx)
        if mapped is None:
            return scene_path, scene_is_custom

        # Migrate stale state from source preset path to packaged templates path.
        if scene_is_custom and getattr(sys, "frozen", False):
            if self._looks_like_legacy_source_preset_path(p):
                return str(mapped), False

        if not scene_is_custom:
            return str(mapped), False
        return scene_path, scene_is_custom

    def _save_ui_state(self):
        try:
            path = self._ui_state_path()
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self._collect_ui_state(), f, ensure_ascii=False, indent=2)
        except Exception as e:
            self._log(f"保存操作缓存失败: {e}")

    def _load_ui_state(self) -> dict:
        path = self._ui_state_path()
        if not path.exists():
            return {}
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def _restore_config_snapshot(self, snapshot: dict) -> bool:
        if not isinstance(snapshot, dict):
            return False
        try:
            self._config = load_scene_from_data(snapshot)
        except Exception as e:
            self._log(f"恢复配置快照失败: {e}")
            return False

        self._refresh_ui_panels()
        self._apply_config_to_controls()
        self._refresh_ui_panels()

        self._clone_btn.setEnabled(True)
        self._fmt_btn.setEnabled(True)
        return True

    def _apply_ui_state_controls(self, controls: dict):
        if not isinstance(controls, dict):
            return
        mode = str(controls.get("mode", "")).upper()
        if mode == "A":
            self._radio_a.setChecked(True)
        elif mode == "B":
            self._radio_b.setChecked(True)

        scope_mode = str(controls.get("scope_mode", "")).lower()
        if scope_mode == "manual":
            self._radio_manual.setChecked(True)
        elif scope_mode == "auto":
            self._radio_auto.setChecked(True)

        page_ranges_text = controls.get("page_ranges_text")
        if page_ranges_text is None and controls.get("body_page") is not None:
            try:
                legacy_page = int(controls.get("body_page"))
                if legacy_page >= 1:
                    page_ranges_text = f"{legacy_page}-{legacy_page}"
            except (TypeError, ValueError):
                page_ranges_text = ""
        if page_ranges_text is not None:
            self._page_ranges_edit.setText(str(page_ranges_text))

        sections = controls.get("sections", {})
        if isinstance(sections, dict):
            for key, checked in sections.items():
                cb = self._scope_checks.get(key)
                if cb is not None:
                    cb.setChecked(bool(checked))
            self._sync_scope_all_state()

        self._section_enable_check.setChecked(bool(controls.get("section_enabled", self._section_enable_check.isChecked())))

        def _set_combo_data(combo: QComboBox, value):
            if value is None:
                return
            idx = combo.findData(value)
            if idx >= 0:
                combo.setCurrentIndex(idx)

        self._table_enable_check.setChecked(bool(controls.get("table_enabled", self._table_enable_check.isChecked())))
        _set_combo_data(self._table_layout_mode_combo, controls.get("table_layout_mode"))
        _set_combo_data(self._table_smart_levels_combo, controls.get("table_smart_levels"))
        _set_combo_data(self._table_border_mode_combo, controls.get("table_border_mode"))
        _set_combo_data(self._table_line_spacing_combo, controls.get("table_line_spacing_mode"))

        self._repeat_header_check.setChecked(bool(controls.get("table_repeat_header", self._repeat_header_check.isChecked())))
        self._md_cleanup_check.setChecked(bool(controls.get("md_cleanup", self._md_cleanup_check.isChecked())))
        self._md_preserve_lists_check.setChecked(
            bool(controls.get("md_preserve_lists", self._md_preserve_lists_check.isChecked()))
        )
        _set_combo_data(
            self._md_list_separator_combo,
            controls.get("md_list_marker_separator"),
        )
        _set_combo_data(
            self._md_ordered_style_combo,
            controls.get("md_ordered_list_style"),
        )
        _set_combo_data(
            self._md_unordered_style_combo,
            controls.get("md_unordered_list_style"),
        )
        self._whitespace_normalize_check.setChecked(
            bool(controls.get("whitespace_normalize", self._whitespace_normalize_check.isChecked()))
        )
        ws_options = controls.get("whitespace_options", {})
        if isinstance(ws_options, dict):
            self._ws_opt_space_variants.setChecked(
                bool(ws_options.get("normalize_space_variants", self._ws_opt_space_variants.isChecked()))
            )
            self._ws_opt_convert_tabs.setChecked(
                bool(ws_options.get("convert_tabs", self._ws_opt_convert_tabs.isChecked()))
            )
            self._ws_opt_zero_width.setChecked(
                bool(ws_options.get("remove_zero_width", self._ws_opt_zero_width.isChecked()))
            )
            self._ws_opt_collapse_spaces.setChecked(
                bool(ws_options.get("collapse_multiple_spaces", self._ws_opt_collapse_spaces.isChecked()))
            )
            self._ws_opt_trim_edges.setChecked(
                bool(ws_options.get("trim_paragraph_edges", self._ws_opt_trim_edges.isChecked()))
            )
            self._ws_opt_smart_convert.setChecked(
                bool(ws_options.get("smart_full_half_convert", self._ws_opt_smart_convert.isChecked()))
            )
            self._ws_opt_smart_punctuation.setChecked(
                bool(ws_options.get("punctuation_by_context", self._ws_opt_smart_punctuation.isChecked()))
            )
            self._ws_opt_smart_bracket.setChecked(
                bool(ws_options.get("bracket_by_inner_language", self._ws_opt_smart_bracket.isChecked()))
            )
            self._ws_opt_smart_alnum.setChecked(
                bool(ws_options.get("fullwidth_alnum_to_halfwidth", self._ws_opt_smart_alnum.isChecked()))
            )
            self._ws_opt_smart_quote.setChecked(
                bool(ws_options.get("quote_by_context", self._ws_opt_smart_quote.isChecked()))
            )
            self._ws_opt_protect_ref.setChecked(
                bool(ws_options.get("protect_reference_numbering", self._ws_opt_protect_ref.isChecked()))
            )
            try:
                conf = int(
                    ws_options.get("context_min_confidence", self._ws_opt_context_conf_spin.currentData() or 2)
                )
            except (TypeError, ValueError):
                conf = self._ws_opt_context_conf_spin.currentData() or 2
            conf_idx = self._ws_opt_context_conf_spin.findData(max(1, min(conf, 3)))
            self._ws_opt_context_conf_spin.setCurrentIndex(conf_idx if conf_idx >= 0 else 1)
        formula_management = controls.get("formula_management", None)
        if formula_management is None:
            formula_management = any(
                bool(controls.get(key, False))
                for key in ("formula_convert", "formula_to_table", "equation_table_format", "formula_style")
            )
        self._formula_management_check.setChecked(
            bool(formula_management)
        )
        self._formula_convert_check.setChecked(
            bool(controls.get("formula_convert", self._formula_convert_check.isChecked()))
        )
        _set_combo_data(self._formula_output_mode_combo, controls.get("formula_convert_output_mode"))
        self._formula_to_table_check.setChecked(
            bool(controls.get("formula_to_table", self._formula_to_table_check.isChecked()))
        )
        self._equation_table_check.setChecked(
            bool(controls.get("equation_table_format", self._equation_table_check.isChecked()))
        )
        self._formula_style_check.setChecked(
            bool(controls.get("formula_style", self._formula_style_check.isChecked()))
        )
        # Citation-link controls are sourced from the active scene/preset
        # instead of sticky UI-state so stale local preferences won't silently
        # disable the rule across sessions.
        self._header_check.setChecked(bool(controls.get("update_header", self._header_check.isChecked())))
        self._pagenum_check.setChecked(bool(controls.get("update_page_number", self._pagenum_check.isChecked())))
        _set_combo_data(self._toc_mode_combo, controls.get("toc_mode"))
        self._header_line_check.setChecked(bool(controls.get("update_header_line", self._header_line_check.isChecked())))
        self._auto_insert_caption_check.setChecked(bool(controls.get("auto_insert_caption", self._auto_insert_caption_check.isChecked())))
        self._format_inserted_check.setChecked(bool(controls.get("format_inserted_caption", self._format_inserted_check.isChecked())))
        self._chem_restore_check.setChecked(bool(controls.get("chem_restore", self._chem_restore_check.isChecked())))

        chem_scopes = controls.get("chem_scopes", {})
        if isinstance(chem_scopes, dict):
            self._chem_scope_refs_check.setChecked(bool(chem_scopes.get("references", self._chem_scope_refs_check.isChecked())))
            self._chem_scope_body_check.setChecked(bool(chem_scopes.get("body", self._chem_scope_body_check.isChecked())))
            has_abstract_key = any(
                k in chem_scopes for k in ("abstract", "abstract_cn", "abstract_en")
            )
            if has_abstract_key:
                abstract_checked = (
                    bool(chem_scopes.get("abstract", False))
                    or bool(chem_scopes.get("abstract_cn", False))
                    or bool(chem_scopes.get("abstract_en", False))
                )
            else:
                abstract_checked = self._chem_scope_abstract_check.isChecked()
            self._chem_scope_abstract_check.setChecked(
                abstract_checked
            )
            self._chem_scope_headings_check.setChecked(bool(chem_scopes.get("headings", self._chem_scope_headings_check.isChecked())))
            self._chem_scope_captions_check.setChecked(bool(chem_scopes.get("captions", self._chem_scope_captions_check.isChecked())))
            self._chem_scope_tables_check.setChecked(bool(chem_scopes.get("tables", self._chem_scope_tables_check.isChecked())))
            self._sync_chem_scope_all_state()

        self._on_table_layout_mode_changed()
        self._on_scope_mode_changed(self._scope_mode_group.checkedId(), True)
        self._on_md_cleanup_toggled(self._md_cleanup_check.isChecked())
        self._on_whitespace_normalize_toggled(self._whitespace_normalize_check.isChecked())
        self._on_whitespace_smart_convert_toggled(self._ws_opt_smart_convert.isChecked())
        self._on_formula_management_toggled(self._formula_management_check.isChecked())
        self._on_formula_convert_toggled(self._formula_convert_check.isChecked())
        self._on_citation_link_toggled(self._citation_link_check.isChecked())
        self._on_header_update_toggled(self._header_check.isChecked())
        self._on_auto_insert_caption_toggled(self._auto_insert_caption_check.isChecked())
        self._on_chem_restore_toggled(self._chem_restore_check.isChecked())

    def _restore_ui_state(self):
        state = self._load_ui_state()
        if not state:
            return
        self._is_restoring_state = True
        try:
            doc_path = state.get("doc_path")
            if isinstance(doc_path, str) and doc_path:
                if Path(doc_path).exists():
                    self._doc_path = doc_path
                    self._file_label.setText(doc_path)
                else:
                    self._doc_path = ""
                    self._file_label.clear()
                    self._log(f"已忽略不存在的上次文件: {doc_path}")

            scene_path = state.get("scene_path")
            scene_is_custom = bool(state.get("scene_is_custom", False))
            if isinstance(scene_path, str):
                self._current_scene_path = scene_path
            self._current_scene_is_custom = scene_is_custom
            scene_restored = False
            if isinstance(scene_path, str) and scene_path:
                scene_restored = self._restore_scene_from_state(scene_path, scene_is_custom)

            config_snapshot = state.get("config_snapshot")
            snapshot_restored = False
            stale_preset_state = (
                isinstance(scene_path, str)
                and bool(scene_path)
                and not scene_is_custom
                and not scene_restored
            )
            if stale_preset_state:
                self._log(f"已忽略失效的历史场景缓存: {scene_path}")
            elif isinstance(config_snapshot, dict) and config_snapshot:
                snapshot_restored = self._restore_config_snapshot(config_snapshot)

            if (not scene_restored and not snapshot_restored) and self._config is not None:
                self._apply_config_to_controls()

            if scene_restored or snapshot_restored:
                controls = state.get("controls", {})
                self._apply_ui_state_controls(controls)
            self._refresh_ui_panels()
            self._run_btn.setEnabled(bool(self._has_valid_doc_path() and self._config))
            if self._config and (scene_restored or snapshot_restored):
                self._log("已恢复上次操作设置")
        finally:
            self._is_restoring_state = False

    def _restore_scene_from_state(self, scene_path: str, is_custom: bool) -> bool:
        p = self._canonicalize_scene_path(Path(scene_path), migrate=True)
        preset_idx = self._find_preset_index(p)

        if preset_idx is not None:
            # When state path points to an existing preset (even after relocation),
            # restore via preset combo to keep scene source within PRESETS_DIR.
            self._scene_combo.setCurrentIndex(preset_idx)
            self._current_scene_is_custom = False
            self._current_scene_path = str(self._scene_item_paths.get(preset_idx, p))
            return True

        if is_custom:
            if p.exists():
                return self._load_scene_from_path(
                    p, custom=True, source_label="自定义场景", clear_on_fail=False
                )
            return False

        return False

    def _clear_scene(self):
        self._config = None
        self._current_scene_path = ""
        self._current_scene_is_custom = False
        self._run_btn.setEnabled(False)

        self._clone_btn.setEnabled(False)
        self._fmt_btn.setEnabled(False)
        self._refresh_ui_panels()

    def _apply_config_to_controls(self):
        if not self._config:
            return
        cfg = self._config

        mode = str(getattr(cfg.heading_numbering, "mode", "B")).upper()
        if mode == "A":
            self._radio_a.setChecked(True)
        else:
            self._radio_b.setChecked(True)

        scope = cfg.format_scope
        scope_mode = str(getattr(scope, "mode", "auto")).strip().lower()
        if scope_mode == "manual":
            self._radio_manual.setChecked(True)
        else:
            self._radio_auto.setChecked(True)
        page_ranges_text = str(getattr(scope, "page_ranges_text", "") or "").strip()
        if not page_ranges_text:
            page = getattr(scope, "body_start_page", None)
            if page is not None and isinstance(page, int) and page >= 1:
                page_ranges_text = f"{page}-{page}"
        self._page_ranges_edit.setText(page_ranges_text)
        for key, cb in self._scope_checks.items():
            cb.setChecked(bool(scope.sections.get(key, True)))
        self._sync_scope_all_state()
        self._on_scope_mode_changed(self._scope_mode_group.checkedId(), True)

        layout_mode = str(getattr(cfg, "normal_table_layout_mode", "smart")).strip().lower()
        layout_idx = self._table_layout_mode_combo.findData(layout_mode)
        self._table_layout_mode_combo.setCurrentIndex(layout_idx if layout_idx >= 0 else 0)
        try:
            smart_levels = int(getattr(cfg, "normal_table_smart_levels", 4))
        except (TypeError, ValueError):
            smart_levels = 4
        smart_idx = self._table_smart_levels_combo.findData(smart_levels)
        self._table_smart_levels_combo.setCurrentIndex(smart_idx if smart_idx >= 0 else 1)
        border_mode = str(getattr(cfg, "normal_table_border_mode", "three_line")).strip().lower()
        border_idx = self._table_border_mode_combo.findData(border_mode)
        self._table_border_mode_combo.setCurrentIndex(border_idx if border_idx >= 0 else 1)
        line_spacing_mode = str(
            getattr(cfg, "normal_table_line_spacing_mode", "single")
        ).strip().lower()
        line_idx = self._table_line_spacing_combo.findData(line_spacing_mode)
        self._table_line_spacing_combo.setCurrentIndex(line_idx if line_idx >= 0 else 0)
        self._repeat_header_check.setChecked(bool(getattr(cfg, "normal_table_repeat_header", False)))
        self._on_table_layout_mode_changed()

        pipeline = list(getattr(cfg, "pipeline", []) or [])
        self._md_cleanup_check.setChecked("md_cleanup" in pipeline)
        md_cfg = getattr(cfg, "md_cleanup", None)
        self._md_preserve_lists_check.setChecked(
            bool(getattr(md_cfg, "preserve_existing_word_lists", True))
        )
        md_list_separator = str(getattr(md_cfg, "list_marker_separator", "tab")).strip().lower()
        md_sep_idx = self._md_list_separator_combo.findData(md_list_separator)
        self._md_list_separator_combo.setCurrentIndex(md_sep_idx if md_sep_idx >= 0 else 0)
        md_ordered_style = str(getattr(md_cfg, "ordered_list_style", "mixed")).strip().lower()
        md_ordered_idx = self._md_ordered_style_combo.findData(md_ordered_style)
        self._md_ordered_style_combo.setCurrentIndex(md_ordered_idx if md_ordered_idx >= 0 else 0)
        md_unordered_style = str(getattr(md_cfg, "unordered_list_style", "word_default")).strip().lower()
        md_unordered_idx = self._md_unordered_style_combo.findData(md_unordered_style)
        self._md_unordered_style_combo.setCurrentIndex(md_unordered_idx if md_unordered_idx >= 0 else 0)
        ws_cfg = getattr(cfg, "whitespace_normalize", None)
        convert_cfg = getattr(cfg, "formula_convert", None)
        to_table_cfg = getattr(cfg, "formula_to_table", None)
        equation_table_cfg = getattr(cfg, "equation_table_format", None)
        style_cfg = getattr(cfg, "formula_style", None)
        self._whitespace_normalize_check.setChecked(
            ("whitespace_normalize" in pipeline) or bool(getattr(ws_cfg, "enabled", False))
        )
        self._formula_convert_check.setChecked(
            ("formula_convert" in pipeline) or bool(getattr(convert_cfg, "enabled", False))
        )
        self._formula_to_table_check.setChecked(
            ("formula_to_table" in pipeline) or bool(getattr(to_table_cfg, "enabled", False))
        )
        self._equation_table_check.setChecked(
            ("equation_table_format" in pipeline) or bool(getattr(equation_table_cfg, "enabled", False))
        )
        self._formula_style_check.setChecked(
            ("formula_style" in pipeline) or bool(getattr(style_cfg, "enabled", False))
        )
        self._formula_management_check.setChecked(
            self._formula_convert_check.isChecked()
            or self._formula_to_table_check.isChecked()
            or self._equation_table_check.isChecked()
            or self._formula_style_check.isChecked()
        )
        output_mode = str(
            getattr(convert_cfg, "output_mode", "word_native")
        ).strip().lower()
        output_idx = self._formula_output_mode_combo.findData(
            output_mode if output_mode in {"word_native", "latex"} else "word_native"
        )
        self._formula_output_mode_combo.setCurrentIndex(output_idx if output_idx >= 0 else 0)
        self._citation_link_check.setChecked("citation_link" in pipeline)
        cite_cfg = getattr(cfg, "citation_link", None)
        if cite_cfg is not None:
            self._citation_ref_auto_number_check.setChecked(
                bool(getattr(cite_cfg, "auto_number_reference_entries", True))
            )
            self._citation_outer_page_sup_check.setChecked(
                bool(getattr(cite_cfg, "superscript_outer_page_numbers", False))
            )

        self._header_check.setChecked(bool(getattr(cfg, "update_header", False)))
        self._pagenum_check.setChecked(bool(getattr(cfg, "update_page_number", False)))
        toc_cfg = getattr(cfg, "toc", None)
        toc_mode = str(getattr(toc_cfg, "mode", "word_native")).strip().lower()
        toc_mode_idx = self._toc_mode_combo.findData(
            toc_mode if toc_mode in {"word_native", "plain"} else "word_native"
        )
        self._toc_mode_combo.setCurrentIndex(toc_mode_idx if toc_mode_idx >= 0 else 0)
        self._header_line_check.setChecked(bool(getattr(cfg, "update_header_line", False)))
        self._auto_insert_caption_check.setChecked(bool(getattr(cfg.caption, "auto_insert", True)))
        self._format_inserted_check.setChecked(bool(getattr(cfg.caption, "format_inserted", False)))
        self._on_header_update_toggled(self._header_check.isChecked())
        self._on_auto_insert_caption_toggled(self._auto_insert_caption_check.isChecked())
        self._sync_lab_controls_from_config()

    def _load_scene_from_path(
        self,
        scene_path: Path,
        *,
        custom: bool,
        source_label: str,
        clear_on_fail: bool,
    ) -> bool:
        canonical_path = self._canonicalize_scene_path(scene_path, migrate=True)
        try:
            self._config = load_scene(canonical_path)
        except Exception as e:
            self._log(f"{source_label}加载失败: {e}")
            if clear_on_fail:
                self._clear_scene()
            else:
                self._refresh_ui_panels()
            return False

        upgrade_notes = get_scene_upgrade_notes(self._config)
        upgrade_save_error = ""
        upgrade_saved = False
        if upgrade_notes:
            try:
                save_scene(self._config, canonical_path)
                upgrade_saved = True
            except Exception as exc:
                upgrade_save_error = str(exc)

        self._current_scene_path = str(canonical_path)
        self._current_scene_is_custom = custom
        self._refresh_ui_panels()
        self._apply_config_to_controls()
        self._refresh_ui_panels()

        self._run_btn.setEnabled(bool(self._has_valid_doc_path()))

        self._clone_btn.setEnabled(True)
        self._fmt_btn.setEnabled(True)
        self._log(f"已加载{source_label}: {self._config.name}")
        if upgrade_notes:
            self._log("已静默升级旧版模板以适配当前格式定义。")
            for note in upgrade_notes:
                self._log(f"  - {note}")
            if upgrade_saved:
                self._log(f"已静默回写升级后的模板: {canonical_path.name}")
            elif upgrade_save_error:
                self._log(f"旧版模板已按当前版本兼容加载，但静默回写失败: {upgrade_save_error}")
        return True

    def _rebuild_section_checkboxes(self, sections: list[str]):
        """按场景可用分区动态重建分区勾选项。"""
        existing_state = {k: cb.isChecked() for k, cb in self._scope_checks.items()}
        keep_all = getattr(self, "_scope_all_check", None)
        for cb in self._scope_checks.values():
            self._section_checks_layout.removeWidget(cb)
            cb.deleteLater()
        self._scope_checks.clear()

        while self._section_checks_layout.count():
            item = self._section_checks_layout.takeAt(0)
            if item.widget():
                if keep_all is not None and item.widget() is keep_all:
                    continue
                item.widget().deleteLater()

        if keep_all is not None:
            self._section_checks_layout.addWidget(keep_all)

        for idx, key in enumerate(sections):
            label = SECTION_LABELS.get(key, key)
            cb = QCheckBox(label)
            # UI 默认全勾选；若本轮刷新前用户已改过，则保留现有勾选状态。
            default_checked = existing_state.get(key, True)
            cb.setChecked(default_checked)
            cb.toggled.connect(self._sync_scope_all_state)
            self._scope_checks[key] = cb
            self._section_checks_layout.addWidget(cb)
            
        self._section_checks_layout.addStretch(1)
        self._sync_scope_all_state()

    def _refresh_ui_panels(self):
        """依据场景 capabilities 控制各 UI 模块显示。"""
        if self._config is None:
            self._mode_section_group.setVisible(True)
            self._scope_mode_row.setVisible(True)
            self._scope_sep_mode.setVisible(True)
            self._manual_inline.setVisible(False)
            self._section_enable_row.setVisible(True)
            self._section_panel.setVisible(True)
            self._table_panel.setVisible(True)
            self._scope_sep1.setVisible(True)
            self._scope_sep2.setVisible(True)
            self._ext_row.setVisible(True)
            self._toc_mode_label.setVisible(True)
            self._toc_mode_combo.setVisible(True)
            self._scope_group.setVisible(True)
            self._md_row.setVisible(True)
            self._on_md_cleanup_toggled(self._md_cleanup_check.isChecked())
            self._ws_panel.setVisible(True)
            self._formula_management_row.setVisible(True)
            self._citation_link_row.setVisible(True)
            self._lab_sep_cite.setVisible(True)
            self._lab_sep1.setVisible(True)
            self._chem_panel.setVisible(True)
            self._lab_group.setVisible(True)
            self._on_table_enable_toggled(self._table_enable_check.isChecked())
            self._on_whitespace_normalize_toggled(self._whitespace_normalize_check.isChecked())
            self._on_citation_link_toggled(self._citation_link_check.isChecked())
            return

        caps = self._config.capabilities or {}
        has_heading = caps.get("heading_numbering", True)
        has_sections = caps.get("section_detection", True)
        has_caption = caps.get("caption", True)
        has_header = caps.get("header_footer", True)
        has_toc = caps.get("toc_rebuild", True)
        has_md = caps.get("md_cleanup", True)
        has_whitespace_normalize = caps.get("whitespace_normalize", True)
        has_equation_table = caps.get("equation_numbering", True)
        has_citation_link = caps.get("citation_link_restore", True)
        has_chem_restore = caps.get("chem_typography_restore", True)

        self._mode_section_group.setVisible(has_heading)
        self._scope_mode_row.setVisible(has_sections)
        self._scope_sep_mode.setVisible(has_sections)
        self._section_enable_row.setVisible(has_sections)
        self._section_panel.setVisible(has_sections)
        self._on_section_enable_toggled(self._section_enable_check.isChecked())
        self._manual_inline.setVisible(has_sections and self._scope_mode_group.checkedId() == 1)
        self._table_panel.setVisible(True)
        self._on_table_enable_toggled(self._table_enable_check.isChecked())
        self._scope_sep1.setVisible(has_sections)

        self._header_check.setVisible(has_header)
        self._pagenum_check.setVisible(has_header)
        self._toc_mode_label.setVisible(has_toc)
        self._toc_mode_combo.setVisible(has_toc)
        self._auto_insert_caption_check.setVisible(has_caption)
        self._on_header_update_toggled(self._header_check.isChecked())
        self._on_auto_insert_caption_toggled(self._auto_insert_caption_check.isChecked())
        self._ext_label.setVisible(has_header or has_caption or has_toc)
        self._ext_row.setVisible(has_header or has_caption or has_toc)
        self._scope_sep2.setVisible(has_header or has_caption or has_toc)

        self._scope_group.setVisible(
            has_sections or has_header or has_caption or has_toc
        )
        self._md_row.setVisible(has_md)
        if has_md:
            self._on_md_cleanup_toggled(self._md_cleanup_check.isChecked())
        else:
            self._md_inline_controls.setVisible(False)
            self._md_options_row.setVisible(False)
        self._lab_sep0.setVisible(
            has_md and (has_whitespace_normalize or has_equation_table or has_citation_link or has_chem_restore)
        )
        self._ws_panel.setVisible(has_whitespace_normalize)
        self._lab_sep_eq.setVisible(
            has_whitespace_normalize and (has_equation_table or has_citation_link or has_chem_restore)
        )
        self._formula_management_row.setVisible(has_equation_table)
        self._lab_sep_cite.setVisible(has_equation_table and has_citation_link)
        self._citation_link_row.setVisible(has_citation_link)
        self._lab_sep1.setVisible(
            (has_md or has_whitespace_normalize or has_equation_table or has_citation_link)
            and has_chem_restore
        )
        self._chem_panel.setVisible(has_chem_restore)
        self._lab_group.setVisible(
            has_md or has_whitespace_normalize or has_equation_table or has_citation_link or has_chem_restore
        )
        self._rebuild_section_checkboxes(self._config.available_sections)

    def _on_scope_mode_changed(self, id_: int, checked: bool):
        if checked:
            allow_manual = self._config is None or self._config.capabilities.get("section_detection", True)
            self._manual_inline.setVisible(allow_manual and id_ == 1)

    def _on_table_layout_mode_changed(self):
        mode = self._table_layout_mode_combo.currentData()
        is_smart = str(mode).strip().lower() == "smart"
        self._table_smart_levels_combo.setEnabled(is_smart)

    def _build_action_section(self):
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 8, 0, 8)
        self._run_btn = QPushButton("开始排版")
        self._run_btn.setObjectName("PrimaryButton")
        self._run_btn.setFixedHeight(44)
        self._run_btn.setEnabled(False)
        self._run_btn.clicked.connect(self._run_pipeline)
        layout.addWidget(self._run_btn, stretch=2)

        self._progress_bar = QProgressBar()
        self._progress_bar.setFixedHeight(44)
        self._progress_bar.setRange(0, 100)
        self._progress_bar.setValue(0)
        self._progress_bar.setFormat("等待开始...")
        self._progress_bar.setAlignment(Qt.AlignCenter)
        self._progress_bar.setObjectName("PipelineProgress")
        layout.addWidget(self._progress_bar, stretch=3)

        self._main_layout.addWidget(row)

    def _build_log_section(self):
        card, layout = self._create_card("执行日志")
        card.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        self._log_text = QTextEdit()
        self._log_text.setReadOnly(True)
        self._log_text.setStyleSheet("border: none; background-color: transparent;")
        layout.addWidget(self._log_text)
        # 基于行高计算高度，确保不会截断文本
        fm = self._log_text.fontMetrics()
        line_h = fm.lineSpacing()
        visible_lines = 12  # 约显示 12 行日志
        # 内边距补偿（card 标题 + 布局 margin + 文本控件 margin）
        padding = 48
        total_h = line_h * visible_lines + padding
        card.setMinimumHeight(total_h)
        card.setMaximumHeight(total_h + line_h * 4)  # 允许向上弹性扩展几行
        self._main_layout.addWidget(card)

    # ── 事件处理 ──

    def _log(self, msg: str):
        self._log_text.append(msg)

    def _set_doc_path(self, path: str, source: str = "已选择"):
        """统一设置文档路径并刷新 UI 状态。"""
        self._doc_path = path
        self._file_label.setText(path)
        self._run_btn.setEnabled(bool(self._config and self._has_valid_doc_path()))
        self._log(f"{source}文件: {path}")

    def _browse_file(self):
        init_dir = os.path.dirname(self._doc_path) if self._doc_path else os.path.expanduser("~")
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 DOCX 文件", init_dir,
            "Word 文档 (*.docx);;所有文件 (*)",
        )
        if path:
            self._set_doc_path(path, "已选择")

    def _sync_all_controls_to_config(self):
        """将当前 UI 所有控件的状态同步回 self._config（不运行流水线）。"""
        if not self._config:
            return
        caps = self._config.capabilities or {}

        # 表格设置
        selected_layout_mode = self._table_layout_mode_combo.currentData()
        self._config.normal_table_layout_mode = (
            str(selected_layout_mode).strip().lower()
            if selected_layout_mode is not None else "smart"
        )
        selected_smart_levels = self._table_smart_levels_combo.currentData()
        try:
            smart_levels = int(selected_smart_levels)
        except (TypeError, ValueError):
            smart_levels = 4
        self._config.normal_table_smart_levels = (
            smart_levels if smart_levels in {3, 4, 5, 6} else 4
        )
        selected_border_mode = self._table_border_mode_combo.currentData()
        self._config.normal_table_border_mode = (
            str(selected_border_mode).strip().lower()
            if selected_border_mode is not None else "three_line"
        )
        selected_line_spacing_mode = self._table_line_spacing_combo.currentData()
        self._config.normal_table_line_spacing_mode = (
            str(selected_line_spacing_mode).strip().lower()
            if selected_line_spacing_mode is not None else "single"
        )
        self._config.normal_table_repeat_header = self._repeat_header_check.isChecked()

        # 编号模式
        has_heading = caps.get("heading_numbering", True)
        if has_heading:
            self._config.heading_numbering.mode = (
                "A" if self._mode_group.checkedId() == 0 else "B"
            )

        # 排版范围
        scope = self._config.format_scope
        has_sections = caps.get("section_detection", True)
        if has_sections:
            if self._scope_mode_group.checkedId() == 1:
                scope.mode = "manual"
                manual_ranges = self._page_ranges_edit.text().strip()
                scope.page_ranges_text = manual_ranges
                scope.body_start_page = None
                scope.body_start_index = None
                scope.body_start_keyword = ""
            else:
                scope.mode = "auto"
                scope.page_ranges_text = ""
            if self._section_enable_check.isChecked():
                for key, cb in self._scope_checks.items():
                    scope.sections[key] = cb.isChecked()
            else:
                for key in self._scope_checks.keys():
                    scope.sections[key] = False

        # 实验室 pipeline
        original_pipeline = list(self._config.pipeline or [])
        self._sync_lab_pipeline_from_checkboxes()
        if list(self._config.pipeline or []) != original_pipeline:
            _mark_pipeline_critical_rules_auto_managed(self._config)

        # Markdown 文本修复
        md_cfg = getattr(self._config, "md_cleanup", None)
        if md_cfg is not None:
            has_md = caps.get("md_cleanup", True)
            md_cfg.enabled = self._md_cleanup_check.isChecked() if has_md else False
            md_cfg.preserve_existing_word_lists = self._md_preserve_lists_check.isChecked()
            md_sep = self._md_list_separator_combo.currentData()
            md_cfg.list_marker_separator = (
                str(md_sep).strip().lower()
                if md_sep in {"tab", "half_space", "full_space"}
                else "tab"
            )
            md_ordered_style = str(self._md_ordered_style_combo.currentData() or "").strip().lower()
            if md_ordered_style not in {v for v, _ in _MD_ORDERED_LIST_STYLE_OPTIONS}:
                md_ordered_style = "mixed"
            md_cfg.ordered_list_style = md_ordered_style
            md_unordered_style = str(self._md_unordered_style_combo.currentData() or "").strip().lower()
            if md_unordered_style not in {v for v, _ in _MD_UNORDERED_LIST_STYLE_OPTIONS}:
                md_unordered_style = "word_default"
            md_cfg.unordered_list_style = md_unordered_style

        # 空白与全半角规范
        ws_cfg = getattr(self._config, "whitespace_normalize", None)
        if ws_cfg is not None:
            has_ws = caps.get("whitespace_normalize", True)
            ws_cfg.enabled = self._whitespace_normalize_check.isChecked() if has_ws else False
            ws_cfg.normalize_space_variants = self._ws_opt_space_variants.isChecked()
            ws_cfg.convert_tabs = self._ws_opt_convert_tabs.isChecked()
            ws_cfg.remove_zero_width = self._ws_opt_zero_width.isChecked()
            ws_cfg.collapse_multiple_spaces = self._ws_opt_collapse_spaces.isChecked()
            ws_cfg.trim_paragraph_edges = self._ws_opt_trim_edges.isChecked()
            ws_cfg.smart_full_half_convert = self._ws_opt_smart_convert.isChecked()
            ws_cfg.punctuation_by_context = self._ws_opt_smart_punctuation.isChecked()
            ws_cfg.bracket_by_inner_language = self._ws_opt_smart_bracket.isChecked()
            ws_cfg.fullwidth_alnum_to_halfwidth = self._ws_opt_smart_alnum.isChecked()
            ws_cfg.quote_by_context = self._ws_opt_smart_quote.isChecked()
            ws_cfg.protect_reference_numbering = self._ws_opt_protect_ref.isChecked()
            ws_cfg.context_min_confidence = self._ws_opt_context_conf_spin.currentData() or 2


        # Formula features
        has_formula = caps.get("equation_numbering", True)
        formula_management_enabled = has_formula and self._formula_management_check.isChecked()

        formula_convert_cfg = getattr(self._config, "formula_convert", None)
        if formula_convert_cfg is not None:
            formula_convert_cfg.enabled = (
                self._formula_convert_check.isChecked() if formula_management_enabled else False
            )
            output_mode = str(self._formula_output_mode_combo.currentData() or "word_native").strip().lower()
            formula_convert_cfg.output_mode = (
                output_mode if output_mode in {"word_native", "latex"} else "word_native"
            )
            # Product constraint: low-confidence formulas must default to skip+mark.
            formula_convert_cfg.low_confidence_policy = "skip_and_mark"

        formula_to_table_cfg = getattr(self._config, "formula_to_table", None)
        if formula_to_table_cfg is not None:
            formula_to_table_cfg.enabled = (
                self._formula_to_table_check.isChecked() if formula_management_enabled else False
            )
            formula_to_table_cfg.block_only = bool(
                getattr(formula_to_table_cfg, "block_only", True)
            )

        equation_table_cfg = getattr(self._config, "equation_table_format", None)
        if equation_table_cfg is not None:
            equation_table_cfg.enabled = (
                self._equation_table_check.isChecked() if formula_management_enabled else False
            )

        formula_style_cfg = getattr(self._config, "formula_style", None)
        if formula_style_cfg is not None:
            formula_style_cfg.enabled = (
                self._formula_style_check.isChecked() if formula_management_enabled else False
            )
            formula_style_cfg.unify_font = bool(
                getattr(formula_style_cfg, "unify_font", True)
            )
            formula_style_cfg.unify_size = bool(
                getattr(formula_style_cfg, "unify_size", True)
            )
            formula_style_cfg.unify_spacing = bool(
                getattr(formula_style_cfg, "unify_spacing", True)
            )
        # Citation link
        citation_cfg = getattr(self._config, "citation_link", None)
        if citation_cfg is not None:
            has_cite = caps.get("citation_link_restore", True)
            citation_cfg.enabled = self._citation_link_check.isChecked() if has_cite else False
            citation_cfg.auto_number_reference_entries = (
                self._citation_ref_auto_number_check.isChecked()
            )
            citation_cfg.superscript_outer_page_numbers = self._citation_outer_page_sup_check.isChecked()

        # 扩展选项
        toc_cfg = getattr(self._config, "toc", None)
        if toc_cfg is not None:
            toc_mode = str(self._toc_mode_combo.currentData() or "word_native").strip().lower()
            toc_cfg.mode = "plain" if toc_mode == "plain" else "word_native"

        has_header = caps.get("header_footer", True)
        self._config.update_header = self._header_check.isChecked() if has_header else False
        self._config.update_page_number = self._pagenum_check.isChecked() if has_header else False
        self._config.update_header_line = self._header_line_check.isChecked() if has_header else False

        has_caption = caps.get("caption", True)
        if has_caption:
            self._config.caption.auto_insert = self._auto_insert_caption_check.isChecked()
            self._config.caption.format_inserted = self._format_inserted_check.isChecked()

        # 化学式上下角标恢复
        has_chem = caps.get("chem_typography_restore", True)
        chem_cfg = getattr(self._config, "chem_typography", None)
        if chem_cfg is not None:
            chem_cfg.enabled = self._chem_restore_check.isChecked() if has_chem else False
            chem_cfg.scopes = {
                "references": self._chem_scope_refs_check.isChecked(),
                "body": self._chem_scope_body_check.isChecked(),
                "abstract_cn": self._chem_scope_abstract_check.isChecked(),
                "abstract_en": self._chem_scope_abstract_check.isChecked(),
                "headings": self._chem_scope_headings_check.isChecked(),
                "captions": self._chem_scope_captions_check.isChecked(),
                "tables": self._chem_scope_tables_check.isChecked(),
            }

    def _persist_current_scene_config(self):
        if (
            self._config is None
            or not self._current_scene_path
            or getattr(self, "_is_restoring_state", False)
            or getattr(self, "_suspend_scene_autosave", False)
        ):
            return
        self._sync_all_controls_to_config()
        scene_path = self._canonicalize_scene_path(
            Path(self._current_scene_path), migrate=True
        )
        try:
            if is_protected_scene_path(scene_path):
                if scene_path.name.lower() == "default_format.json":
                    scene_path = PRESETS_DIR / scene_path.name
                else:
                    return
        except Exception:
            pass
        self._current_scene_path = str(scene_path)
        save_scene(self._config, scene_path)

    def _save_current_scene_config(self):
        """将当前场景的 UI 修改保存回场景 JSON 文件。"""
        try:
            self._persist_current_scene_config()
        except Exception as e:
            self._log(f"自动保存场景配置失败: {e}")

    def _on_scene_changed(self, index: int):
        self._save_current_scene_config()
        data = self._scene_combo.itemData(index, Qt.UserRole)
        if not isinstance(data, dict) or data.get("kind") != "scene":
            self._clear_scene()
            return

        scene_path = self._scene_item_paths.get(index)
        if scene_path is None:
            self._clear_scene()
            return

        self._load_scene_from_path(
            scene_path,
            custom=False,
            source_label="场景",
            clear_on_fail=True,
        )

    def _load_custom_scene(self):
        init_dir = os.path.expanduser("~")
        path, _ = QFileDialog.getOpenFileName(
            self, "选择场景配置文件", init_dir,
            "JSON 文件 (*.json);;所有文件 (*)",
        )
        if path:
            self._load_scene_from_path(
                Path(path),
                custom=True,
                source_label="自定义场景",
                clear_on_fail=False,
            )

    def _clone_word_template_style(self):
        if self._clone_in_progress:
            self._log("克隆模板格式请求正在执行，已忽略重复触发。")
            return

        self._clone_in_progress = True
        clone_btn_was_enabled = bool(getattr(self, "_clone_btn", None) and self._clone_btn.isEnabled())
        if getattr(self, "_clone_btn", None) is not None:
            self._clone_btn.setEnabled(False)

        try:
            init_dir = os.path.expanduser("~")
            path, _ = QFileDialog.getOpenFileName(
                self, "选择用于克隆格式的 Word 文档", init_dir,
                "Word 文档 (*.docx);;所有文件 (*)",
            )
            if not path:
                return

            try:
                cloned_config = load_default_scene()
            except Exception as e:
                self._log(f"加载默认场景失败: {e}")
                QMessageBox.critical(self, "克隆模板格式失败", str(e))
                return

            try:
                summary = clone_scene_style_from_docx(cloned_config, path)
            except Exception as e:
                self._log(f"克隆模板格式失败: {e}")
                QMessageBox.critical(self, "克隆模板格式失败", str(e))
                return

            updated = list(summary.get("styles_updated", []) or [])
            missing = list(summary.get("styles_missing", []) or [])
            added = list(summary.get("styles_added", []) or [])
            page_updated = bool(summary.get("page_setup_updated", False))
            numbering_updated = bool(summary.get("numbering_updated", False))
            numbering_levels = list(summary.get("numbering_levels_updated", []) or [])
            numbering_fallback = bool(summary.get("numbering_fallback_used", False))
            numbering_inferred = bool(summary.get("numbering_inferred", False))

            if not updated and not page_updated and not numbering_updated and not added:
                self._log(f"未从模板匹配到可克隆的格式: {path}")
                QMessageBox.information(
                    self, "克隆模板格式",
                    "未匹配到可克隆的样式，已保留默认场景配置。"
                )
                return

            updated_labels = [_STYLE_DISPLAY_NAMES.get(k, k) for k in updated]
            preview = "、".join(updated_labels[:8])
            if len(updated_labels) > 8:
                preview += f" ...（共 {len(updated_labels)} 项）"

            msg_parts = []
            if updated:
                msg_parts.append(f"样式已克隆 {len(updated)} 项")
            if added:
                msg_parts.append(f"补全样式槽位 {len(added)} 项")
            if page_updated:
                msg_parts.append("页面设置已克隆")
            if numbering_updated:
                if numbering_fallback:
                    msg_parts.append(f"标题编号已补全默认 {len(numbering_levels)} 级")
                elif numbering_inferred:
                    msg_parts.append(f"标题编号已克隆 {len(numbering_levels)} 级")
                else:
                    msg_parts.append(f"标题编号已更新 {len(numbering_levels)} 级")
            if missing:
                msg_parts.append(f"未匹配 {len(missing)} 项（已保留默认值）")
            msg = "，".join(msg_parts)

            # 弹窗让用户命名新格式
            from pathlib import Path as _Path
            src_name = _Path(path).stem
            default_name = f"克隆自 {src_name}"
            detail = msg
            if preview:
                detail += f"\n\n更新项: {preview}"
            detail += "\n\n请为新格式命名："

            new_name, ok = QInputDialog.getText(
                self, "克隆模板格式", detail, text=default_name)
            if not ok:
                return  # 用户取消

            new_name = new_name.strip() or default_name
            cloned_config.name = new_name
            # 克隆为新格式时不继承源模板的永久署名。
            cloned_config.format_signature = ""

            # 保存为新文件
            from src.scene.manager import _safe_filename
            safe_stem = _safe_filename(new_name)
            new_path = PRESETS_DIR / f"{safe_stem}.json"
            idx = 1
            while new_path.exists():
                new_path = PRESETS_DIR / f"{safe_stem}_v{idx}.json"
                idx += 1

            try:
                save_scene(cloned_config, new_path)
            except Exception as e:
                QMessageBox.critical(self, "保存失败", str(e))
                return

            self._log(f"已从模板克隆格式: {path}")
            self._log(msg)
            if preview:
                self._log(f"已更新样式: {preview}")
            self._log(f"新格式已保存为: {new_path.name}")

            # 切换到新场景
            self._suspend_scene_autosave = True
            try:
                self._config = cloned_config
                self._current_scene_path = str(new_path)
                self._apply_config_to_controls()
                self._refresh_ui_panels()
                self._populate_scene_combo()
                target_idx = self._find_preset_index(Path(new_path))
                if target_idx is not None:
                    combo = self._scene_combo
                    combo.blockSignals(True)
                    combo.setCurrentIndex(target_idx)
                    combo.blockSignals(False)
            finally:
                self._suspend_scene_autosave = False
            self._save_ui_state()

            QMessageBox.information(
                self, "克隆模板格式",
                f"{msg}\n\n已保存为新格式「{new_name}」")
        finally:
            self._clone_in_progress = False
            if getattr(self, "_clone_btn", None) is not None:
                self._clone_btn.setEnabled(clone_btn_was_enabled)

    def _open_format_config(self):
        if not self._config:
            return
        if self._worker and self._worker.isRunning():
            QMessageBox.information(self, "正在排版", "当前任务执行中，暂不可修改格式配置。")
            return
        if self._current_scene_path:
            self._current_scene_path = str(
                self._canonicalize_scene_path(
                    Path(self._current_scene_path), migrate=True
                )
            )
        self._sync_lab_pipeline_from_checkboxes()
        dlg = FormatConfigDialog(
            self._config, self,
            scene_path=self._current_scene_path or None,
            scene_items=self._scene_item_paths,
        )
        result = dlg.exec()

        dialog_scene_path = getattr(dlg, "_scene_path", None)
        if dialog_scene_path:
            self._current_scene_path = str(
                self._canonicalize_scene_path(Path(dialog_scene_path), migrate=True)
            )

        deleted_scene_path = getattr(dlg, "_deleted_scene_path", None)
        if deleted_scene_path:
            def _norm_path(raw: str) -> str:
                try:
                    return str(Path(raw).resolve())
                except Exception:
                    return str(Path(raw))

            if (
                self._current_scene_path
                and _norm_path(self._current_scene_path) == _norm_path(deleted_scene_path)
            ):
                self._current_scene_path = ""
                self._current_scene_is_custom = False

        if result == QDialog.Accepted:
            self._apply_config_to_controls()
            self._refresh_ui_panels()
            self._save_current_scene_config()

        self._populate_scene_combo()

        target_idx: int | None = None
        if self._current_scene_path:
            target_idx = self._find_preset_index(Path(self._current_scene_path))

        if target_idx is not None:
            self._suspend_scene_autosave = True
            try:
                self._scene_combo.blockSignals(True)
                self._scene_combo.setCurrentIndex(target_idx)
                self._scene_combo.blockSignals(False)
                self._on_scene_changed(target_idx)
            finally:
                self._suspend_scene_autosave = False
        elif (
            self._current_scene_is_custom
            and self._current_scene_path
            and Path(self._current_scene_path).exists()
        ):
            # Keep current custom scene in memory; combo only lists preset scenes.
            self._refresh_ui_panels()
        else:
            first_scene_index = None
            for i in range(self._scene_combo.count()):
                data = self._scene_combo.itemData(i, Qt.UserRole)
                if isinstance(data, dict) and data.get("kind") == "scene":
                    first_scene_index = i
                    break
            if first_scene_index is not None:
                self._suspend_scene_autosave = True
                try:
                    self._scene_combo.blockSignals(True)
                    self._scene_combo.setCurrentIndex(first_scene_index)
                    self._scene_combo.blockSignals(False)
                    self._on_scene_changed(first_scene_index)
                finally:
                    self._suspend_scene_autosave = False
            else:
                self._clear_scene()

        self._save_ui_state()

    def _sync_md_cleanup_pipeline_from_checkbox(self):
        if not self._config:
            return
        caps = self._config.capabilities or {}
        has_md_cleanup = caps.get("md_cleanup", True)
        pipeline = (
            list(self._config.pipeline)
            if self._config.pipeline is not None else []
        )
        if has_md_cleanup and self._md_cleanup_check.isChecked():
            if "md_cleanup" not in pipeline:
                idx = pipeline.index("style_manager") if "style_manager" in pipeline else 0
                pipeline.insert(idx, "md_cleanup")
        else:
            pipeline = [s for s in pipeline if s != "md_cleanup"]
        self._config.pipeline = pipeline

    def _sync_whitespace_normalize_pipeline_from_checkbox(self):
        if not self._config:
            return
        caps = self._config.capabilities or {}
        has_whitespace_normalize = caps.get("whitespace_normalize", True)
        pipeline = (
            list(self._config.pipeline)
            if self._config.pipeline is not None else []
        )
        if has_whitespace_normalize and self._whitespace_normalize_check.isChecked():
            if "whitespace_normalize" not in pipeline:
                if "md_cleanup" in pipeline:
                    idx = pipeline.index("md_cleanup") + 1
                elif "style_manager" in pipeline:
                    idx = pipeline.index("style_manager")
                else:
                    idx = 0
                pipeline.insert(idx, "whitespace_normalize")
        else:
            pipeline = [s for s in pipeline if s != "whitespace_normalize"]
        self._config.pipeline = pipeline

    def _sync_table_format_pipeline_from_checkbox(self):
        if not self._config:
            return
        pipeline = (
            list(self._config.pipeline)
            if self._config.pipeline is not None else []
        )
        if self._table_enable_check.isChecked():
            if "table_format" not in pipeline:
                rank = {name: idx for idx, name in enumerate(PIPELINE_DEFAULT_ORDER)}
                target_rank = rank.get("table_format", 10_000)
                insert_idx = len(pipeline)
                for idx, step in enumerate(pipeline):
                    step_rank = rank.get(step, 10_000)
                    if step_rank > target_rank:
                        insert_idx = idx
                        break
                pipeline.insert(insert_idx, "table_format")
        else:
            pipeline = [s for s in pipeline if s != "table_format"]
        self._config.pipeline = pipeline

    def _sync_formula_pipeline_from_checkboxes(self):
        if not self._config:
            return
        caps = self._config.capabilities or {}
        has_formula_features = caps.get("equation_numbering", True)
        pipeline = (
            list(self._config.pipeline)
            if self._config.pipeline is not None else []
        )

        formula_steps = [
            "formula_convert",
            "formula_to_table",
            "equation_table_format",
            "formula_style",
        ]
        pipeline = [s for s in pipeline if s not in formula_steps]

        if has_formula_features and self._formula_management_check.isChecked():
            enabled_steps = [
                ("formula_convert", self._formula_convert_check.isChecked()),
                ("formula_to_table", self._formula_to_table_check.isChecked()),
                ("equation_table_format", self._equation_table_check.isChecked()),
                ("formula_style", self._formula_style_check.isChecked()),
            ]
            ordered_to_insert = [step for step, enabled in enabled_steps if enabled]
            if ordered_to_insert:
                if "table_format" in pipeline:
                    idx = pipeline.index("table_format") + 1
                elif "section_format" in pipeline:
                    idx = pipeline.index("section_format")
                else:
                    idx = len(pipeline)
                for step in ordered_to_insert:
                    pipeline.insert(idx, step)
                    idx += 1

        self._config.pipeline = pipeline

    def _sync_citation_link_pipeline_from_checkbox(self):
        if not self._config:
            return

        caps = self._config.capabilities or {}
        has_citation_link = caps.get("citation_link_restore", True)
        pipeline = (
            list(self._config.pipeline)
            if self._config.pipeline is not None else []
        )
        if has_citation_link and self._citation_link_check.isChecked():
            if "citation_link" not in pipeline:
                idx = len(pipeline)
                for anchor in (
                    "section_format",
                    "formula_style",
                    "equation_table_format",
                    "formula_to_table",
                    "formula_convert",
                    "table_format",
                ):
                    if anchor in pipeline:
                        idx = pipeline.index(anchor) + 1
                        break
                pipeline.insert(idx, "citation_link")
        else:
            pipeline = [s for s in pipeline if s != "citation_link"]
        self._config.pipeline = pipeline

    def _sync_lab_pipeline_from_checkboxes(self):
        self._sync_md_cleanup_pipeline_from_checkbox()
        self._sync_whitespace_normalize_pipeline_from_checkbox()
        self._sync_table_format_pipeline_from_checkbox()
        self._sync_formula_pipeline_from_checkboxes()
        self._sync_citation_link_pipeline_from_checkbox()

    def _run_pipeline(self):
        if not self._doc_path or not self._config:
            return
        if not self._has_valid_doc_path():
            self._run_btn.setEnabled(False)
            QMessageBox.warning(self, "文件不存在", "当前选择的 DOCX 文件不存在，请重新选择。")
            return
        caps = self._config.capabilities or {}
        selected_layout_mode = self._table_layout_mode_combo.currentData()
        self._config.normal_table_layout_mode = (
            str(selected_layout_mode).strip().lower()
            if selected_layout_mode is not None else "smart"
        )
        selected_smart_levels = self._table_smart_levels_combo.currentData()
        try:
            smart_levels = int(selected_smart_levels)
        except (TypeError, ValueError):
            smart_levels = 4
        self._config.normal_table_smart_levels = (
            smart_levels if smart_levels in {3, 4, 5, 6} else 4
        )
        selected_border_mode = self._table_border_mode_combo.currentData()
        self._config.normal_table_border_mode = (
            str(selected_border_mode).strip().lower()
            if selected_border_mode is not None else "three_line"
        )
        selected_line_spacing_mode = self._table_line_spacing_combo.currentData()
        self._config.normal_table_line_spacing_mode = (
            str(selected_line_spacing_mode).strip().lower()
            if selected_line_spacing_mode is not None else "single"
        )
        self._config.normal_table_repeat_header = self._repeat_header_check.isChecked()

        # 编号设置
        has_heading = caps.get("heading_numbering", True)
        mode = self._config.heading_numbering.mode
        if has_heading:
            mode = "A" if self._mode_group.checkedId() == 0 else "B"
            self._config.heading_numbering.mode = mode
            self._config.heading_numbering.apply_scheme()

        # 排版范围设置
        scope = self._config.format_scope
        has_sections = caps.get("section_detection", True)
        if has_sections:
            if self._scope_mode_group.checkedId() == 1:
                scope.mode = "manual"
                manual_ranges = self._page_ranges_edit.text().strip()
                scope.page_ranges_text = manual_ranges
                scope.body_start_page = None
                scope.body_start_index = None
                scope.body_start_keyword = ""
            else:
                scope.mode = "auto"
                scope.page_ranges_text = ""
            
            if self._section_enable_check.isChecked():
                for key, cb in self._scope_checks.items():
                    scope.sections[key] = cb.isChecked()
            else:
                for key in self._scope_checks.keys():
                    scope.sections[key] = False
        else:
            scope.mode = "auto"
            scope.page_ranges_text = ""

        if has_sections and scope.mode == "manual":
            try:
                parsed_ranges = parse_page_ranges_text(scope.page_ranges_text)
            except ValueError as exc:
                QMessageBox.warning(
                    self,
                    "页码范围格式错误",
                    f"修正范围填写有误：{exc}\n示例：27-40,44-56",
                )
                return
            if not parsed_ranges:
                QMessageBox.warning(
                    self,
                    "缺少修正范围",
                    "请选择“指定修正范围”后填写页码范围，示例：27-40,44-56",
                )
                return
            scope.page_ranges_text = format_page_ranges_text(parsed_ranges)

        # Markdown 粘贴修复开关（主界面开关优先）
        self._sync_lab_pipeline_from_checkboxes()
        md_cfg = getattr(self._config, "md_cleanup", None)
        if md_cfg is not None:
            has_md = caps.get("md_cleanup", True)
            md_cfg.enabled = self._md_cleanup_check.isChecked() if has_md else False
            md_cfg.preserve_existing_word_lists = self._md_preserve_lists_check.isChecked()
            md_sep = self._md_list_separator_combo.currentData()
            md_cfg.list_marker_separator = (
                str(md_sep).strip().lower() if md_sep in {"tab", "half_space", "full_space"} else "tab"
            )
            md_ordered_style = str(self._md_ordered_style_combo.currentData() or "").strip().lower()
            if md_ordered_style not in {v for v, _ in _MD_ORDERED_LIST_STYLE_OPTIONS}:
                md_ordered_style = "mixed"
            md_cfg.ordered_list_style = md_ordered_style
            md_unordered_style = str(self._md_unordered_style_combo.currentData() or "").strip().lower()
            if md_unordered_style not in {v for v, _ in _MD_UNORDERED_LIST_STYLE_OPTIONS}:
                md_unordered_style = "word_default"
            md_cfg.unordered_list_style = md_unordered_style

        # 空白与全半角规范（实验室）
        has_ws_normalize = caps.get("whitespace_normalize", True)
        ws_cfg = getattr(self._config, "whitespace_normalize", None)
        if ws_cfg is not None:
            ws_cfg.enabled = (
                self._whitespace_normalize_check.isChecked() if has_ws_normalize else False
            )
            ws_cfg.normalize_space_variants = self._ws_opt_space_variants.isChecked()
            ws_cfg.convert_tabs = self._ws_opt_convert_tabs.isChecked()
            ws_cfg.remove_zero_width = self._ws_opt_zero_width.isChecked()
            ws_cfg.collapse_multiple_spaces = self._ws_opt_collapse_spaces.isChecked()
            ws_cfg.trim_paragraph_edges = self._ws_opt_trim_edges.isChecked()
            ws_cfg.smart_full_half_convert = self._ws_opt_smart_convert.isChecked()
            ws_cfg.punctuation_by_context = self._ws_opt_smart_punctuation.isChecked()
            ws_cfg.bracket_by_inner_language = self._ws_opt_smart_bracket.isChecked()
            ws_cfg.fullwidth_alnum_to_halfwidth = self._ws_opt_smart_alnum.isChecked()
            ws_cfg.quote_by_context = self._ws_opt_smart_quote.isChecked()
            ws_cfg.protect_reference_numbering = self._ws_opt_protect_ref.isChecked()
            ws_cfg.context_min_confidence = self._ws_opt_context_conf_spin.currentData() or 2


        # Formula features (lab)
        has_formula = caps.get("equation_numbering", True)
        formula_management_enabled = has_formula and self._formula_management_check.isChecked()
        formula_convert_cfg = getattr(self._config, "formula_convert", None)
        if formula_convert_cfg is not None:
            formula_convert_cfg.enabled = (
                self._formula_convert_check.isChecked() if formula_management_enabled else False
            )
            output_mode = str(self._formula_output_mode_combo.currentData() or "word_native").strip().lower()
            formula_convert_cfg.output_mode = (
                output_mode if output_mode in {"word_native", "latex"} else "word_native"
            )
            # Product constraint: low-confidence formulas must default to skip+mark.
            formula_convert_cfg.low_confidence_policy = "skip_and_mark"

        formula_to_table_cfg = getattr(self._config, "formula_to_table", None)
        if formula_to_table_cfg is not None:
            formula_to_table_cfg.enabled = (
                self._formula_to_table_check.isChecked() if formula_management_enabled else False
            )
            formula_to_table_cfg.block_only = bool(
                getattr(formula_to_table_cfg, "block_only", True)
            )

        equation_table_cfg = getattr(self._config, "equation_table_format", None)
        if equation_table_cfg is not None:
            equation_table_cfg.enabled = (
                self._equation_table_check.isChecked() if formula_management_enabled else False
            )

        formula_style_cfg = getattr(self._config, "formula_style", None)
        if formula_style_cfg is not None:
            formula_style_cfg.enabled = (
                self._formula_style_check.isChecked() if formula_management_enabled else False
            )
            formula_style_cfg.unify_font = bool(
                getattr(formula_style_cfg, "unify_font", True)
            )
            formula_style_cfg.unify_size = bool(
                getattr(formula_style_cfg, "unify_size", True)
            )
            formula_style_cfg.unify_spacing = bool(
                getattr(formula_style_cfg, "unify_spacing", True)
            )
        # Citation link (lab)
        has_citation_link = caps.get("citation_link_restore", True)
        citation_cfg = getattr(self._config, "citation_link", None)
        if citation_cfg is not None:
            citation_cfg.enabled = (
                self._citation_link_check.isChecked() if has_citation_link else False
            )
            citation_cfg.auto_number_reference_entries = (
                self._citation_ref_auto_number_check.isChecked()
            )
            citation_cfg.superscript_outer_page_numbers = (
                self._citation_outer_page_sup_check.isChecked()
            )

        # 扩展选项
        toc_cfg = getattr(self._config, "toc", None)
        if toc_cfg is not None:
            toc_mode = str(self._toc_mode_combo.currentData() or "word_native").strip().lower()
            toc_cfg.mode = "plain" if toc_mode == "plain" else "word_native"

        has_header = caps.get("header_footer", True)
        self._config.update_header = self._header_check.isChecked() if has_header else False
        self._config.update_page_number = self._pagenum_check.isChecked() if has_header else False
        self._config.update_header_line = self._header_line_check.isChecked() if has_header else False

        has_caption = caps.get("caption", True)
        self._config.caption.auto_insert = (
            self._auto_insert_caption_check.isChecked() if has_caption else False
        )
        self._config.caption.format_inserted = (
            self._format_inserted_check.isChecked() if has_caption else False
        )

        # 化学式上下角标恢复（实验室）
        has_chem_restore = caps.get("chem_typography_restore", True)
        chem_cfg = getattr(self._config, "chem_typography", None)
        if chem_cfg is not None:
            chem_cfg.enabled = (
                self._chem_restore_check.isChecked() if has_chem_restore else False
            )
            chem_cfg.scopes = {
                "references": self._chem_scope_refs_check.isChecked(),
                "body": self._chem_scope_body_check.isChecked(),
                "abstract_cn": self._chem_scope_abstract_check.isChecked(),
                "abstract_en": self._chem_scope_abstract_check.isChecked(),
                "headings": self._chem_scope_headings_check.isChecked(),
                "captions": self._chem_scope_captions_check.isChecked(),
                "tables": self._chem_scope_tables_check.isChecked(),
            }

        if has_sections:
            if scope.mode == "auto":
                scope_desc = "自动识别"
            else:
                ranges = str(getattr(scope, "page_ranges_text", "") or "").strip()
                scope_desc = f"指定页码({ranges or '未填写'})"
            enabled = [k for k, v in scope.sections.items() if v]
        else:
            scope_desc = "不启用"
            enabled = []
        self._log(f"编号: {'重建' if mode == 'B' else '保留'}, "
                  f"场景: {self._config.name}")
        self._log(f"范围: {scope_desc}, 分区: {', '.join(enabled)}")
        self._log(
            "流程控制: md_cleanup={}, whitespace_normalize={}, formula_convert={}, formula_to_table={}, equation_table_format={}, formula_style={}, citation_link={}（以实验室开关为准）".format(
                "开启" if "md_cleanup" in (self._config.pipeline or []) else "关闭",
                "开启" if "whitespace_normalize" in (self._config.pipeline or []) else "关闭",
                "开启" if "formula_convert" in (self._config.pipeline or []) else "关闭",
                "开启" if "formula_to_table" in (self._config.pipeline or []) else "关闭",
                "开启" if "equation_table_format" in (self._config.pipeline or []) else "关闭",
                "开启" if "formula_style" in (self._config.pipeline or []) else "关闭",
                "开启" if "citation_link" in (self._config.pipeline or []) else "关闭",
            )
        )
        if "whitespace_normalize" in (self._config.pipeline or []) and ws_cfg is not None:
            ws_options = []
            if ws_cfg.normalize_space_variants:
                ws_options.append("特殊空格统一为半角空格")
            if ws_cfg.convert_tabs:
                ws_options.append("Tab归一为空格")
            if ws_cfg.remove_zero_width:
                ws_options.append("移除零宽字符")
            if ws_cfg.collapse_multiple_spaces:
                ws_options.append("压缩连续半角空格")
            if ws_cfg.trim_paragraph_edges:
                ws_options.append("清理段落首尾空格")
            if ws_cfg.smart_full_half_convert:
                ws_options.append("语境符号规范")
                smart_opts = []
                if ws_cfg.punctuation_by_context:
                    smart_opts.append("标点自动修正（中英文）")
                if ws_cfg.bracket_by_inner_language:
                    smart_opts.append("括号按括号内语种修正")
                if ws_cfg.fullwidth_alnum_to_halfwidth:
                    smart_opts.append("全角字母数字转半角")
                if ws_cfg.quote_by_context:
                    smart_opts.append("引号按语境修正（中英文）")
                if ws_cfg.protect_reference_numbering:
                    smart_opts.append("额外保护参考文献编号")
                smart_opts.append(f"阈值={ws_cfg.context_min_confidence}")
                ws_options.append("[" + ", ".join(smart_opts) + "]")
            self._log("空白/符号子项: " + ("、".join(ws_options) if ws_options else "（无）"))
        if "citation_link" in (self._config.pipeline or []) and citation_cfg is not None:
            self._log(
                "引用域关联子项: 参考文献自动编号纠偏={}，方括号外页码跟随上标={}".format(
                    "开启" if citation_cfg.auto_number_reference_entries else "关闭",
                    "开启" if citation_cfg.superscript_outer_page_numbers else "关闭"
                )
            )
        self._run_btn.setEnabled(False)
        self._clone_btn.setEnabled(False)
        self._fmt_btn.setEnabled(False)
        self._log("开始排版...")
        self._progress_bar.setValue(0)
        self._progress_bar.setFormat("开始排版...")
        run_config = copy.deepcopy(self._config)
        self._worker = PipelineWorker(run_config, self._doc_path)
        self._worker.progress.connect(self._on_worker_progress)
        self._worker.finished.connect(self._on_worker_finished)
        self._worker.start()

    def _on_cancel_requested(self):
        if self._worker and self._worker.isRunning():
            self._worker.cancel()
            self._log("已请求取消，等待当前步骤结束...")

    def _on_worker_progress(self, current, total, message):
        if total > 0:
            pct = int(current / total * 100)
            self._progress_bar.setValue(pct)
            self._progress_bar.setFormat(f"{message} ({pct}%)")
        else:
            self._progress_bar.setFormat(message)
        self._log(message)

    def _on_worker_finished(self, result: PipelineResult):
        # 进度条完成
        self._progress_bar.setValue(100)
        self._progress_bar.setFormat("完成")

        self._run_btn.setEnabled(True)
        self._clone_btn.setEnabled(True)
        self._fmt_btn.setEnabled(True)
        self._worker = None

        if result is None:
            self._log("排版失败: 后台任务未返回结果")
            QMessageBox.critical(self, "排版失败", "后台任务未返回结果")
            return

        if result and result.cancelled:
            self._log("排版已取消。")
            return

        status = getattr(result, "status", "success")
        if not result.success and status not in {"partial_success", "failed"}:
            self._log(f"排版失败: {result.error}")
            QMessageBox.critical(self, "排版失败", result.error or "未知错误")
            return
        if status == "failed" and result.doc is None and not result.output_paths:
            self._log(f"排版失败: {result.error}")
            QMessageBox.critical(self, "排版失败", result.error or "未知错误")
            return

        if status == "partial_success":
            self._log("排版部分成功，正在生成输出文件...")
        elif status == "failed":
            self._log("排版完成但存在关键步骤失败（严格模式）。")
        else:
            self._log("排版完成，正在生成输出文件...")

        paths = result.output_paths

        # 对比稿由 Pipeline 统一生成，主线程仅展示结果路径
        compare_path = paths.get("compare")
        if compare_path:
            if Path(compare_path).exists():
                self._log(f"对比稿已生成: {compare_path}")
            else:
                self._log(f"对比稿生成失败: 未找到输出文件 ({compare_path})")

        # 生成报告
        validation_issues = result.tracker.records
        context_issues = []
        for rec in result.tracker.records:
            if rec.rule_name == "validation":
                from src.engine.rules.base import ValidationIssue
                context_issues.append(ValidationIssue(
                    level="warning" if rec.success else "error",
                    rule_name=rec.rule_name,
                    message=rec.before,
                    location=rec.target,
                ))

        report_data = collect_report(
            result.tracker,
            scene_name=self._config.name if self._config else "",
            input_file=self._doc_path,
            validation_issues=context_issues,
        )

        json_path = paths.get("report_json")
        if json_path:
            try:
                generate_json_report(report_data, json_path)
                self._log(f"JSON 报告已生成: {json_path}")
            except Exception as e:
                self._log(f"JSON 报告生成失败: {e}")

        md_path = paths.get("report_md")
        if md_path:
            try:
                generate_markdown_report(report_data, md_path)
                self._log(f"Markdown 报告已生成: {md_path}")
            except Exception as e:
                self._log(f"Markdown 报告生成失败: {e}")

        # 汇总
        summary = result.tracker.summary()
        final_path = paths.get("final", "")
        self._log("─" * 40)
        status_label = {
            "success": "成功",
            "partial_success": "部分成功",
            "failed": "失败",
            "cancelled": "已取消",
        }.get(status, status)
        self._log(
            f"排版{status_label}! 共 {summary.get('total', 0)} 项修改, "
            f"{summary.get('failures', 0)} 项失败"
        )
        # 显示失败详情
        failed_items = list(getattr(result, "failed_items", []) or [])
        if not failed_items:
            failed_items = [
                {
                    "rule_name": rec.rule_name,
                    "target": rec.target,
                    "change_type": rec.change_type,
                    "reason": rec.failure_reason,
                }
                for rec in result.tracker.get_failures()
            ]

        for item in failed_items:
            rule_name = item.get("rule_name", "")
            target = item.get("target", "")
            change_type = item.get("change_type", "")
            reason_raw = item.get("reason") or "未知原因"
            if (
                rule_name == "pipeline"
                and target == "final_doc_fields"
                and change_type == "refresh"
            ):
                reason = "请手动更新页码。"
            else:
                reason = reason_raw
            self._log(f"  ✗ [{rule_name}] {target}: {reason}")

        if status == "failed":
            self._log(f"严格模式命中关键步骤失败，未成功项: {len(failed_items)}")
            if result.error:
                self._log(f"失败摘要: {result.error}")
        if final_path:
            primary_final_path = paths.get("final_primary", "")
            if primary_final_path and primary_final_path != final_path:
                self._log(
                    f"注意: 目标文件被占用，最终稿改存为: {final_path} "
                    f"(原目标: {primary_final_path})"
                )
            self._log(f"最终稿: {final_path}")
        sub_dir = paths.get("sub_dir", "")
        if sub_dir:
            self._log(f"附件文件夹: {sub_dir}")


    def closeEvent(self, event):
        self._save_current_scene_config()
        self._save_ui_state()
        super().closeEvent(event)

    # ── Drag & Drop ─────────────────────────────────────────
    def dragEnterEvent(self, event):
        """Accept the drag only if it contains at least one .docx file."""
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith(".docx"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        """Handle the drop: pick the first .docx file and set it as input."""
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".docx"):
                self._set_doc_path(path, "已拖入")
                event.acceptProposedAction()
                return
        event.ignore()
