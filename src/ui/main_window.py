"""主窗口：docx 一键排版桌面应用"""

import os
import json
import copy
import html
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
    QHeaderView, QDialogButtonBox, QFormLayout, QFrame,
    QDoubleSpinBox, QSpinBox, QApplication, QSizePolicy,
    QProgressBar, QInputDialog, QLayout, QScrollArea, QStyle
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtCore import QTimer

from src.scene.manager import (
    PRESETS_DIR,
    list_presets,
    load_scene,
    load_scene_from_data,
    save_scene,
)
from src.scene.schema import SceneConfig
from src.engine.pipeline import Pipeline, PipelineResult
from src.report.collector import collect_report
from src.report.json_report import generate_json_report
from src.report.markdown_report import generate_markdown_report
from src.ui.progress_dialog import ProgressDialog
from src.converter.md_to_docx import convert_md_to_docx
from src.docx_io.style_clone import clone_scene_style_from_docx
from src.ui.theme_manager import ThemeManager
from src.utils.line_spacing import normalize_line_spacing


class ScrollSafeComboBox(QComboBox):
    """QComboBox that ignores mouse wheel events to prevent accidental value changes."""
    def wheelEvent(self, event):
        event.ignore()

class WhitespaceVisibleLineEdit(QLineEdit):
    """QLineEdit that shows whitespace as visible symbols.

    Display: full-width space → □, regular space → ·, tab → →
    The actual raw value is returned by text() transparently.
    """

    _TO_SYMBOL = {
        "\u3000": "\u25A1",   # □  full-width space
        " ": "\u2022",        # •  regular space
        "\t": "\u2192",       # →  tab
    }
    _FROM_SYMBOL = {
        "\u25A1": "\u3000",   # □ → full-width space
        "\u2022": " ",        # • → regular space
        "\u2192": "\t",       # → → tab
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
    ("size_pt", "字号(pt)", "float"),
    ("bold", "加粗", "bool"),
    ("italic", "斜体", "bool"),
    ("alignment", "对齐", "str"),
    ("first_line_indent_chars", "首行缩进(字)", "float"),
    ("left_indent_chars", "左缩进(字)", "float"),
    ("line_spacing_type", "行距类型", "str"),
    ("line_spacing_pt", "行距值(倍数/pt)", "float"),
    ("space_before_pt", "段前(pt)", "float"),
    ("space_after_pt", "段后(pt)", "float"),
]

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
    "cover": "封面",
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
    "whitespace_normalize": "空白字符统一（低风险）",
    "style_manager": "样式管理",
    "heading_detect": "标题识别",
    "heading_numbering": "标题编号",
    "toc_format": "目录格式",
    "caption_format": "题注格式",
    "table_format": "表格格式",
    "equation_table_format": "公式表格识别与调整",
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
    "equation_table_format",
    "section_format",
    "citation_link",
    "header_footer",
    "validation",
]


class FormatConfigDialog(QDialog):
    """格式配置查看/编辑对话框"""

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
        self.setWindowTitle("格式配置")
        self.setMinimumSize(960, 560)
        self._config = config
        self._normalize_heading_numbering_legacy_keys()
        self._scene_path = scene_path  # 当前场景文件路径
        self._scene_items = scene_items or {}  # combo index → path
        self._deleted_scene_path: str | None = None
        self._build_ui()

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

        # ── 格式管理工具栏 ──
        toolbar = QHBoxLayout()
        toolbar.addWidget(QLabel("当前格式:"))
        self._fmt_name_label = QLabel(
            getattr(self._config, "name", "") or "(未命名)")
        self._fmt_name_label.setStyleSheet("font-weight: bold;")
        toolbar.addWidget(self._fmt_name_label, stretch=1)

        self._rename_btn = QPushButton("重命名")
        self._rename_btn.setEnabled(bool(self._scene_path))
        self._rename_btn.clicked.connect(self._rename_format)
        toolbar.addWidget(self._rename_btn)

        self._delete_btn = QPushButton("删除")
        self._delete_btn.setEnabled(bool(self._scene_path))
        self._delete_btn.clicked.connect(self._delete_format)
        toolbar.addWidget(self._delete_btn)

        layout.addLayout(toolbar)

        # ── Tab 页 ──
        tabs = QTabWidget()
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
        self._save_btn.setEnabled(bool(self._scene_path))
        self._save_btn.clicked.connect(self._save_format)
        self._apply_btn = btn_box.addButton("应用修改", QDialogButtonBox.AcceptRole)
        self._apply_btn.clicked.connect(self._apply_changes)
        btn_box.addButton("关闭", QDialogButtonBox.RejectRole)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

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
        old_name = getattr(self._config, "name", "") or path.stem
        new_name, ok = QInputDialog.getText(
            self, "重命名格式", "新名称:", text=old_name)
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

    def _save_format(self):
        from src.scene.manager import save_scene, _safe_filename
        from pathlib import Path
        if not self._scene_path:
            QMessageBox.warning(self, "增量保存", "无关联的场景文件，无法保存。")
            return
        src_path = Path(self._scene_path)
        # 先应用当前界面修改到 config
        self._sync_config_from_ui()
        # 生成默认名称
        old_config_name = getattr(self._config, "name", "") or src_path.stem
        default_name = f"{old_config_name} v2"
        # 弹窗让用户确认或修改名称
        new_name, ok = QInputDialog.getText(
            self, "增量保存", "新格式名称:", text=default_name)
        if not ok:
            return  # 用户取消
        new_name = new_name.strip() or default_name
        self._config.name = new_name
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
            save_scene(self._config, new_path)
            self._scene_path = str(new_path)
            self._fmt_name_label.setText(self._config.name)
            QMessageBox.information(
                self, "增量保存成功",
                f"已保存为新文件：\n{new_path.name}")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "增量保存失败", str(e))

    # ── 样式表格 ──

    def _build_styles_tab(self) -> QWidget:
        from PySide6.QtGui import QColor, QBrush
        w = QWidget()
        layout = QVBoxLayout(w)
        self._style_table = QTableWidget()
        cols = _STYLE_COLUMNS
        style_keys = list(self._config.styles.keys())
        backfilled = getattr(self._config, "_backfilled_styles", set())

        self._style_table.setColumnCount(len(cols))
        self._style_table.setRowCount(len(style_keys))
        self._style_table.setHorizontalHeaderLabels([c[1] for c in cols])
        self._style_table.setVerticalHeaderLabels(
            [_STYLE_DISPLAY_NAMES.get(k, k) for k in style_keys])

        gray_brush = QBrush(QColor(180, 180, 180))
        for row, key in enumerate(style_keys):
            sc = self._config.styles[key]
            is_placeholder = key in backfilled
            for col, (attr, _, typ) in enumerate(cols):
                val = getattr(sc, attr, "")
                item = QTableWidgetItem(str(val))
                if is_placeholder:
                    item.setForeground(gray_brush)
                self._style_table.setItem(row, col, item)

        self._style_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeToContents)
        self._style_keys = style_keys
        layout.addWidget(self._style_table)
        return w

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
    _NUM_COLS = [
        ("format", "编号格式", "str"),
        ("template", "模板", "str"),
        ("separator", "分隔符", "str"),
        ("custom_separator", "自定义分隔符", "str"),
        ("alignment", "对齐", "str"),
        ("left_indent_chars", "左缩进(字)", "float"),
    ]
    _NUM_SEPARATOR_OPTIONS = [
        ("fullwidth_space", "全角空格(□ / U+3000)", "\u3000"),
        ("halfwidth_space", "半角空格(· / U+0020)", " "),
        ("underscore", "下划线( _ / U+005F )", "_"),
        ("custom", "自定义", None),
    ]
    _NUM_SEPARATOR_PRESET_VALUES = {
        key: value
        for key, _, value in _NUM_SEPARATOR_OPTIONS
        if value is not None
    }
    _NUM_SEPARATOR_VALUE_TO_KEY = {
        value: key for key, value in _NUM_SEPARATOR_PRESET_VALUES.items()
    }

    @classmethod
    def _num_col_index(cls, attr_name: str) -> int:
        for idx, (attr, _, _) in enumerate(cls._NUM_COLS):
            if attr == attr_name:
                return idx
        raise KeyError(f"unknown numbering column: {attr_name}")

    def _build_separator_widgets_for_level(self, level_cfg) -> tuple[QComboBox, WhitespaceVisibleLineEdit]:
        raw_separator = getattr(level_cfg, "separator", "")
        if raw_separator is None:
            raw_separator = ""
        raw_separator = str(raw_separator)
        custom_separator = getattr(level_cfg, "custom_separator", None)

        if custom_separator is not None:
            mode_key = "custom"
            custom_value = str(custom_separator)
        else:
            mode_key = self._NUM_SEPARATOR_VALUE_TO_KEY.get(raw_separator, "custom")
            custom_value = raw_separator if mode_key == "custom" else ""

        combo = ScrollSafeComboBox()
        for key, label, _ in self._NUM_SEPARATOR_OPTIONS:
            combo.addItem(label, key)
        combo.setToolTip("全角空格显示为 □，半角空格显示为 ·。")
        idx = combo.findData(mode_key)
        if idx < 0:
            idx = combo.findData("fullwidth_space")
        combo.setCurrentIndex(max(idx, 0))

        custom_edit = WhitespaceVisibleLineEdit(custom_value)
        custom_edit.setToolTip("输入时：全角空格=□，半角空格=·，Tab=→。")
        custom_edit.setEnabled(str(combo.currentData()) == "custom")
        combo.currentIndexChanged.connect(
            lambda *_args, c=combo, edit=custom_edit: edit.setEnabled(
                str(c.currentData()) == "custom"
            )
        )
        return combo, custom_edit

    def _sync_numbering_separator_from_row(self, level_cfg, row: int) -> None:
        widgets = getattr(self, "_num_separator_widgets", {})
        pair = widgets.get(row)
        if not pair:
            return
        combo, custom_edit = pair
        mode_key = str(combo.currentData() or "fullwidth_space")
        if mode_key == "custom":
            level_cfg.custom_separator = custom_edit.text()
            if not isinstance(level_cfg.separator, str):
                level_cfg.separator = "\u3000"
            return
        level_cfg.separator = self._NUM_SEPARATOR_PRESET_VALUES.get(mode_key, "\u3000")
        level_cfg.custom_separator = None

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
        hn = self._config.heading_numbering
        scheme_items: list[tuple[str, str]] = []
        if hn.schemes:
            scheme_items = [(str(sid), str(sid)) for sid in hn.schemes.keys()]
            if hn.scheme and hn.scheme not in hn.schemes and hn.levels:
                # 当前场景未声明该方案，但存在活动 levels，提供可编辑兜底入口。
                scheme_items.insert(0, (f"{hn.scheme} (活动 levels)", "__levels__"))
        else:
            scheme_items = [("活动 levels", "__levels__")]
        self._scheme_combo = ScrollSafeComboBox()
        for label, sid in scheme_items:
            self._scheme_combo.addItem(label, sid)
        idx = self._scheme_combo.findData(hn.scheme)
        if idx < 0:
            idx = self._scheme_combo.findData("__levels__")
        if idx >= 0:
            self._scheme_combo.setCurrentIndex(idx)
        row = QWidget()
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.addWidget(QLabel("方案:"))
        row_layout.addWidget(self._scheme_combo)
        row_layout.addWidget(QLabel("提示: 仅修改当前方案"))
        row_layout.addStretch(1)
        layout.addWidget(row)

        self._num_table = QTableWidget()
        cols = self._NUM_COLS
        self._num_table.setColumnCount(len(cols))
        self._num_table.setHorizontalHeaderLabels([c[1] for c in cols])

        self._num_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeToContents)
        layout.addWidget(self._num_table)
        self._scheme_combo.currentIndexChanged.connect(self._on_scheme_changed)
        self._load_numbering_rows_for_scheme(
            str(self._scheme_combo.currentData() or "__levels__")
        )
        return w

    def _resolve_numbering_levels(self, scheme_id: str) -> dict:
        hn = self._config.heading_numbering
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
        from PySide6.QtGui import QColor, QBrush
        levels = self._resolve_numbering_levels(scheme_id)
        # Always show all 8 heading levels
        self._num_level_keys = list(self._ALL_HEADING_KEYS)
        self._num_table.setRowCount(len(self._num_level_keys))
        self._num_table.setVerticalHeaderLabels(
            [self._LEVEL_DISPLAY.get(k, k) for k in self._num_level_keys]
        )
        sep_col = self._num_col_index("separator")
        custom_sep_col = self._num_col_index("custom_separator")
        self._num_separator_widgets = {}
        gray_brush = QBrush(QColor(180, 180, 180))
        for row, key in enumerate(self._num_level_keys):
            is_defined = key in levels
            if is_defined:
                lc = levels[key]
            else:
                from src.scene.schema import HeadingLevelConfig
                lc = HeadingLevelConfig()
            for col, (attr, _, _) in enumerate(self._NUM_COLS):
                if attr in {"separator", "custom_separator"}:
                    continue
                val = getattr(lc, attr, "")
                if val is None:
                    val = ""
                item = QTableWidgetItem(str(val) if is_defined else "")
                if not is_defined:
                    item.setForeground(gray_brush)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self._num_table.setItem(row, col, item)
            sep_combo, custom_edit = self._build_separator_widgets_for_level(lc)
            if not is_defined:
                sep_combo.setEnabled(False)
                custom_edit.setEnabled(False)
            self._num_table.setCellWidget(row, sep_col, sep_combo)
            self._num_table.setCellWidget(row, custom_sep_col, custom_edit)
            self._num_separator_widgets[row] = (sep_combo, custom_edit)

    def _on_scheme_changed(self, *_):
        if not hasattr(self, "_scheme_combo"):
            return
        scheme_id = str(self._scheme_combo.currentData() or "").strip()
        if not scheme_id:
            return
        self._load_numbering_rows_for_scheme(scheme_id)

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
            spin = QSpinBox()
            spin.setRange(minimum, maximum)
            spin.setValue(int(getattr(hn.risk_guard, attr, 0)))
            self._risk_guard_controls[attr] = spin
            layout.addRow(label + ":", spin)

        def _add_guard_float(attr: str, label: str, minimum: float = 0.0, maximum: float = 1.0):
            spin = QDoubleSpinBox()
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
        self._cap_custom_separator_edit.setToolTip("输入时：全角空格=□，半角空格=·，Tab=→。")
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

        self._cap_enabled = QCheckBox("启用题注处理")
        self._cap_enabled.setChecked(cap.enabled)
        layout.addRow(self._cap_enabled)
        return w

    # ── 表格配置 ──

    def _build_table_tab(self) -> QWidget:
        w = QWidget()
        layout = QFormLayout(w)
        cfg = self._config

        note = QLabel(
            "线宽单位为磅（pt），常用值：0.5磅/0.75磅/1磅/1.5磅"
        )
        note.setWordWrap(True)
        layout.addRow(note)

        # 全框线表线宽
        self._grid_border_width_spin = QDoubleSpinBox()
        self._grid_border_width_spin.setRange(0.25, 3.0)
        self._grid_border_width_spin.setSingleStep(0.25)
        self._grid_border_width_spin.setDecimals(2)
        self._grid_border_width_spin.setValue(
            getattr(cfg, "table_border_width_pt", 0.5)
        )
        layout.addRow("全框线 — 线宽(磅):", self._grid_border_width_spin)

        # 三线表表头线宽
        self._three_line_header_spin = QDoubleSpinBox()
        self._three_line_header_spin.setRange(0.25, 3.0)
        self._three_line_header_spin.setSingleStep(0.25)
        self._three_line_header_spin.setDecimals(2)
        self._three_line_header_spin.setValue(
            getattr(cfg, "three_line_header_width_pt", 1.0)
        )
        layout.addRow("三线表 — 表头线宽(磅):", self._three_line_header_spin)

        # 三线表表尾线宽
        self._three_line_bottom_spin = QDoubleSpinBox()
        self._three_line_bottom_spin.setRange(0.25, 3.0)
        self._three_line_bottom_spin.setSingleStep(0.25)
        self._three_line_bottom_spin.setDecimals(2)
        self._three_line_bottom_spin.setValue(
            getattr(cfg, "three_line_bottom_width_pt", 0.5)
        )
        layout.addRow("三线表 — 表尾线宽(磅):", self._three_line_bottom_spin)

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
            spin = QDoubleSpinBox()
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
        return w

    # ── 执行流程 ──

    def _build_pipeline_tab(self) -> QWidget:
        w = QWidget()
        layout = QFormLayout(w)
        note = QLabel(
            "说明：md_cleanup 由主界面“实验室 - Markdown 文本修复”开关统一控制，"
            "whitespace_normalize 由主界面“实验室 - 空白字符统一（低风险）”开关统一控制，"
            "equation_table_format 由主界面“实验室 - 公式表格识别与调整”开关统一控制，"
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
                cb.setToolTip("由主界面“实验室 - 空白字符统一（低风险）”开关控制。")
            elif step == "equation_table_format":
                cb.setEnabled(False)
                cb.setToolTip("由主界面“实验室 - 公式表格识别与调整”开关控制。")
            elif step == "citation_link":
                cb.setEnabled(False)
                cb.setToolTip("由主界面“实验室 - 正文参考文献域关联”开关控制。")
            self._pipeline_checks[step] = cb
            layout.addRow(f"{PIPELINE_STEP_LABELS.get(step, step)} ({step}):", cb)
        return w

    # ── 应用修改 ──

    def _apply_changes(self):
        self._sync_config_from_ui()
        self.accept()

    def _sync_config_from_ui(self):
        invalid_cells = []
        # 样式
        for row, key in enumerate(self._style_keys):
            sc = self._config.styles[key]
            for col, (attr, _, typ) in enumerate(_STYLE_COLUMNS):
                item = self._style_table.item(row, col)
                if not item:
                    continue
                text = item.text().strip()
                if attr in {"alignment", "line_spacing_type"}:
                    text = text.lower()
                try:
                    if typ == "float":
                        setattr(sc, attr, float(text))
                    elif typ == "bool":
                        setattr(sc, attr, self._parse_bool_text(text))
                    else:
                        setattr(sc, attr, text)
                except (ValueError, TypeError):
                    if typ == "bool":
                        invalid_cells.append(
                            f"{_STYLE_DISPLAY_NAMES.get(key, key)} / {_STYLE_COLUMNS[col][1]} = {text}"
                        )
            normalized = normalize_line_spacing(
                getattr(sc, "line_spacing_type", ""),
                getattr(sc, "line_spacing_pt", 0),
            )
            if normalized is not None:
                sc.line_spacing_type, sc.line_spacing_pt = normalized

        # 标题编号
        hn = self._config.heading_numbering
        selected_scheme = str(self._scheme_combo.currentData() or "").strip()
        if selected_scheme and selected_scheme != "__levels__":
            hn.scheme = selected_scheme
        levels = self._resolve_numbering_levels(selected_scheme or hn.scheme)
        for row, key in enumerate(self._num_level_keys):
            lc = levels.get(key)
            if not lc:
                continue
            for col, (attr, _, typ) in enumerate(self._NUM_COLS):
                if attr in {"separator", "custom_separator"}:
                    continue
                item = self._num_table.item(row, col)
                if not item:
                    continue
                text = item.text().strip()
                if attr == "alignment":
                    text = text.lower()
                try:
                    if typ == "float":
                        setattr(lc, attr, float(text))
                    else:
                        setattr(lc, attr, text)
                except (ValueError, TypeError):
                    pass
            self._sync_numbering_separator_from_row(lc, row)
        if selected_scheme and selected_scheme != "__levels__" and selected_scheme in hn.schemes:
            hn.apply_scheme(selected_scheme)

        # 标题高级
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
        cap = self._config.caption
        cap.enabled = self._cap_enabled.isChecked()
        for attr, edit in self._cap_edits.items():
            setattr(cap, attr, edit.text())
        self._sync_caption_separator_from_ui(cap)
        # 编号格式从下拉框读取
        cap.numbering_format = str(
            self._numbering_format_combo.currentData() or "chapter.seq"
        )

        # 表格线宽
        self._config.table_border_width_pt = self._grid_border_width_spin.value()
        self._config.three_line_header_width_pt = self._three_line_header_spin.value()
        self._config.three_line_bottom_width_pt = self._three_line_bottom_spin.value()

        # 页面
        ps = self._config.page_setup
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
        out = self._config.output
        for attr, cb in getattr(self, "_output_checks", {}).items():
            setattr(out, attr, cb.isChecked())

        # 执行流程
        selected_steps = [
            step for step in getattr(self, "_pipeline_step_order", [])
            if self._pipeline_checks.get(step) and self._pipeline_checks[step].isChecked()
        ]
        unknown_steps = [
            step for step in (self._config.pipeline or [])
            if step not in getattr(self, "_pipeline_checks", {})
        ]
        self._config.pipeline = selected_steps + unknown_steps

        if invalid_cells:
            preview = "\n".join(invalid_cells[:8])
            if len(invalid_cells) > 8:
                preview += f"\n... 共 {len(invalid_cells)} 处"
            QMessageBox.warning(
                self,
                "格式配置",
                f"以下布尔值格式无效，已保留原值：\n{preview}\n\n可用值示例: true/false, 1/0, yes/no",
            )


class MainWindow(QMainWindow):
    """docx 一键排版 主窗口"""
    _INIT_WIDTH_RATIO = 0.35
    _INIT_HEIGHT_RATIO = 0.76
    _MAX_SCREEN_USAGE = 0.90

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Lark-Formatter V0.1")
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
        title_label = QLabel("Lark-Formatter 论文一键排版工具V0.1")
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
        dialog = AboutDialog(self)
        dialog.exec()

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
        self._radio_manual = QPushButton("指定修正起点")
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
        self._manual_page_label = QLabel("起始页码:")
        manual_inline_layout.addWidget(self._manual_page_label)
        self._body_page_spin = QSpinBox()
        self._body_page_spin.setRange(1, 999)
        self._body_page_spin.setValue(1)
        self._body_page_spin.setToolTip(
            "输入 Word 文档中正文开始的实际页码（物理第几页，非页脚编号）")
        manual_inline_layout.addWidget(self._body_page_spin)
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
        self._auto_insert_caption_check = QCheckBox("自动插入缺失题注（表头图尾）")
        self._auto_insert_caption_check.setChecked(True)
        row6.addWidget(self._auto_insert_caption_check)
        self._format_inserted_inline = QWidget()
        format_inserted_row = QHBoxLayout(self._format_inserted_inline)
        format_inserted_row.setContentsMargins(14, 0, 0, 0)
        format_inserted_row.setSpacing(0)
        self._format_inserted_check = QCheckBox("插入题注使用域代码")
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
        card, outer = self._create_card("实验室（谨慎使用）")
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
        row.addStretch(1)
        self._lab_help_md_btn = self._build_lab_help_button("md_cleanup")
        row.addWidget(self._lab_help_md_btn)
        self._md_row.setFixedHeight(36)
        outer.addWidget(self._md_row)
        
        self._lab_sep0 = QFrame()
        self._lab_sep0.setFrameShape(QFrame.HLine)
        self._lab_sep0.setStyleSheet("color: #e1e4e8; background-color: #e1e4e8;")
        outer.addWidget(self._lab_sep0)
        self._whitespace_row = QWidget()
        ws_row = QHBoxLayout(self._whitespace_row)
        ws_row.setContentsMargins(0, 0, 0, 0)
        self._whitespace_normalize_check = QCheckBox("空白字符统一")
        self._whitespace_normalize_check.setChecked(False)
        self._whitespace_normalize_check.setToolTip(
            "该开关控制执行流程中的 whitespace_normalize 步骤；运行前会覆盖“格式配置-执行流程”中同名步骤的状态。"
        )
        ws_row.addWidget(self._whitespace_normalize_check)
        ws_row.addStretch(1)
        self._lab_help_whitespace_btn = self._build_lab_help_button("whitespace_normalize")
        ws_row.addWidget(self._lab_help_whitespace_btn)
        self._whitespace_row.setFixedHeight(36)

        self._whitespace_options_row = QWidget()
        ws_opts_row = QHBoxLayout(self._whitespace_options_row)
        ws_opts_row.setContentsMargins(24, 0, 0, 0)
        ws_opts_row.addWidget(QLabel("基础选项:"))
        
        self._ws_opt_space_variants = QCheckBox("统一全角/特殊空格")
        self._ws_opt_convert_tabs = QCheckBox("Tab 转空格")
        self._ws_opt_zero_width = QCheckBox("清理零宽字符")
        self._ws_opt_collapse_spaces = QCheckBox("折叠连续空格")
        self._ws_opt_trim_edges = QCheckBox("清理段首尾空格")
        self._ws_opt_space_variants.setChecked(True)
        self._ws_opt_convert_tabs.setChecked(True)
        self._ws_opt_zero_width.setChecked(True)
        self._ws_opt_collapse_spaces.setChecked(True)
        self._ws_opt_trim_edges.setChecked(True)
        
        ws_opts_row.addWidget(self._ws_opt_space_variants)
        ws_opts_row.addWidget(self._ws_opt_convert_tabs)
        ws_opts_row.addWidget(self._ws_opt_zero_width)
        ws_opts_row.addWidget(self._ws_opt_collapse_spaces)
        ws_opts_row.addWidget(self._ws_opt_trim_edges)
        ws_opts_row.addStretch(1)
        self._whitespace_options_row.setMinimumHeight(32)

        self._whitespace_smart_row = QWidget()
        ws_smart_row = QHBoxLayout(self._whitespace_smart_row)
        ws_smart_row.setContentsMargins(24, 0, 0, 0)
        self._whitespace_smart_label = QLabel("智能全半角:")
        ws_smart_row.addWidget(self._whitespace_smart_label)
        
        self._ws_opt_smart_convert = QCheckBox("启用智能转换")
        self._ws_opt_smart_punctuation = QCheckBox("标点按语境")
        self._ws_opt_smart_bracket = QCheckBox("括号按内文语种")
        self._ws_opt_smart_alnum = QCheckBox("全角英数字转半角")
        self._ws_opt_smart_convert.setChecked(True)
        self._ws_opt_smart_punctuation.setChecked(True)
        self._ws_opt_smart_bracket.setChecked(True)
        self._ws_opt_smart_alnum.setChecked(True)
        
        ws_smart_row.addWidget(self._ws_opt_smart_convert)
        ws_smart_row.addWidget(self._ws_opt_smart_punctuation)
        ws_smart_row.addWidget(self._ws_opt_smart_bracket)
        ws_smart_row.addWidget(self._ws_opt_smart_alnum)
        ws_smart_row.addStretch(1)
        self._whitespace_smart_row.setMinimumHeight(32)

        self._whitespace_smart_row2 = QWidget()
        ws_smart_row2 = QHBoxLayout(self._whitespace_smart_row2)
        ws_smart_row2.setContentsMargins(24, 0, 0, 0)
        self._whitespace_smart_label2 = QLabel("高级规则:")
        ws_smart_row2.addWidget(self._whitespace_smart_label2)
        
        self._ws_opt_smart_quote = QCheckBox("引号按语境")
        self._ws_opt_protect_ref = QCheckBox("保护文献编号")
        self._ws_opt_context_conf_spin = QSpinBox()
        self._ws_opt_context_conf_spin.setRange(1, 4)
        self._ws_opt_context_conf_spin.setValue(2)
        self._ws_opt_context_conf_spin.setToolTip("语境判定置信阈值，越大越保守。")
        self._ws_opt_smart_quote.setChecked(False)
        self._ws_opt_protect_ref.setChecked(True)
        
        ws_smart_row2.addWidget(self._ws_opt_smart_quote)
        ws_smart_row2.addWidget(self._ws_opt_protect_ref)
        ws_smart_row2.addWidget(QLabel("智能词法分析信任阈值:"))
        ws_smart_row2.addWidget(self._ws_opt_context_conf_spin)
        ws_smart_row2.addStretch(1)
        self._whitespace_smart_row2.setMinimumHeight(32)
        
        self._whitespace_smart_row.setMinimumHeight(32)
        
        self._ws_panel = QFrame()
        self._ws_panel.setObjectName("SubGroupPanel")
        self._ws_panel.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
        ws_vbox = QVBoxLayout(self._ws_panel)
        ws_vbox.setContentsMargins(0, 0, 0, 0)
        ws_vbox.setSpacing(6)
        ws_vbox.addWidget(self._whitespace_row)
        ws_vbox.addWidget(self._whitespace_options_row)
        ws_vbox.addWidget(self._whitespace_smart_row)
        ws_vbox.addWidget(self._whitespace_smart_row2)
        outer.addWidget(self._ws_panel)

        self._lab_sep_eq = QFrame()
        self._lab_sep_eq.setFrameShape(QFrame.HLine)
        self._lab_sep_eq.setStyleSheet("color: #e1e4e8; background-color: #e1e4e8;")
        outer.addWidget(self._lab_sep_eq)

        self._equation_table_row = QWidget()
        eq_row = QHBoxLayout(self._equation_table_row)
        eq_row.setContentsMargins(0, 0, 0, 0)
        self._equation_table_check = QCheckBox("公式表格识别与调整")
        self._equation_table_check.setChecked(False)
        self._equation_table_check.setToolTip(
            "该开关控制执行流程中的 equation_table_format 步骤；运行前会覆盖“格式配置-执行流程”中同名步骤的状态。"
        )
        eq_row.addWidget(self._equation_table_check)
        eq_row.addStretch(1)
        self._lab_help_equation_btn = self._build_lab_help_button("equation_table_format")
        eq_row.addWidget(self._lab_help_equation_btn)
        self._equation_table_row.setFixedHeight(36)
        outer.addWidget(self._equation_table_row)

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
        self._citation_ref_auto_number_check = QCheckBox("参考文献自动编号并纠偏（实验）")
        self._citation_ref_auto_number_check.setChecked(True)
        self._citation_ref_auto_number_check.setToolTip(
            "将参考文献序号改为 SEQ 域自动编号，并让正文引用通过 REF 域同步显示，插入新文献后可整体更新纠偏。"
        )
        cite_row.addWidget(self._citation_ref_auto_number_check)
        self._citation_outer_page_sup_check = QCheckBox("方括号外页码跟随上标（实验）")
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
        self._chem_restore_check.setChecked(True)
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
        self._chem_scope_refs_check.setChecked(True)
        self._chem_scope_body_check.setChecked(False)
        self._chem_scope_abstract_check.setChecked(False)
        self._chem_scope_headings_check.setChecked(False)
        chem_scope_inline_row.addWidget(self._chem_scope_all_check)
        chem_scope_inline_row.addWidget(self._chem_scope_refs_check)
        chem_scope_inline_row.addWidget(self._chem_scope_body_check)
        chem_scope_inline_row.addWidget(self._chem_scope_abstract_check)
        chem_scope_inline_row.addWidget(self._chem_scope_headings_check)
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
        self._citation_link_check.toggled.connect(self._on_citation_link_toggled)
        self._chem_restore_check.toggled.connect(self._on_chem_restore_toggled)
        self._chem_scope_all_check.toggled.connect(self._on_chem_scope_all_toggled)
        self._chem_scope_refs_check.toggled.connect(self._sync_chem_scope_all_state)
        self._chem_scope_body_check.toggled.connect(self._sync_chem_scope_all_state)
        self._chem_scope_abstract_check.toggled.connect(self._sync_chem_scope_all_state)
        self._chem_scope_headings_check.toggled.connect(self._sync_chem_scope_all_state)
        self._sync_chem_scope_all_state()
        self._on_whitespace_normalize_toggled(self._whitespace_normalize_check.isChecked())
        self._on_whitespace_smart_convert_toggled(self._ws_opt_smart_convert.isChecked())
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
                "status": "开发状态：经过小型测试回归，风险不明确，谨慎使用",
                "feature": "实际功能：修复word中残留的markdown格式文本",
                "note": "注意事项：未存在这种问题<b>请勿使用！！！</b>，避免误改",
            },
            "whitespace_normalize": {
                "title": "空白字符统一 使用须知",
                "status": "开发状态：完成开发未测试，风险较高，<b>请勿使用！！！</b>",
                "feature": "实际功能：自动统一半角、全角符号，以及回车、制表符不统一的问题",
                "note": "注意事项：功能面窄，功能不稳定，请谨慎校核",
            },
            "equation_table_format": {
                "title": "公式表格识别与调整 使用须知",
                "status": "开发状态：开发完成、测试通过，场景受限",
                "feature": "实际功能：自动分配公式与其序号的位置，调整公式格式",
                "note": "注意事项：只能识别两列表格形式的公式，左侧公式右侧序号",
            },
            "citation_link": {
                "title": "正文参考文献域关联 使用须知",
                "status": "开发状态：开发完成、测试通过，场景有限",
                "feature": "实际功能：将正文中的[1]类引用转换为可跳转域，并可启用参考文献序号自动编号纠偏",
                "note": "注意事项：仅识别常见编号格式，且会跳过已有域代码段落",
            },
            "chem_typography": {
                "title": "自动恢复上下角标 使用须知",
                "status": "开发状态：开发完成、测试通过，存在缺陷",
                "feature": "实际功能：自动修正如endnote导入后上下角标失效的问题",
                "note": "注意事项：黑名单与白名单可能不能覆盖全部潜在",
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
        self._whitespace_normalize_check.setChecked("whitespace_normalize" in pipeline)
        self._equation_table_check.setChecked("equation_table_format" in pipeline)
        self._citation_link_check.setChecked("citation_link" in pipeline)
        cite_cfg = getattr(self._config, "citation_link", None)
        if cite_cfg is not None:
            self._citation_ref_auto_number_check.setChecked(
                bool(getattr(cite_cfg, "auto_number_reference_entries", True))
            )
            self._citation_outer_page_sup_check.setChecked(
                bool(getattr(cite_cfg, "superscript_outer_page_numbers", False))
            )
        ws_cfg = getattr(self._config, "whitespace_normalize", None)
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
            self._ws_opt_context_conf_spin.setValue(max(1, min(conf, 4)))
        self._on_whitespace_normalize_toggled(self._whitespace_normalize_check.isChecked())
        self._on_whitespace_smart_convert_toggled(self._ws_opt_smart_convert.isChecked())
        self._on_citation_link_toggled(self._citation_link_check.isChecked())

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
            self._chem_scope_syncing = False
            self._sync_chem_scope_all_state()
            self._chem_restore_check.setChecked(bool(getattr(chem_cfg, "enabled", True)))
        self._on_chem_restore_toggled(self._chem_restore_check.isChecked())

    def _on_whitespace_normalize_toggled(self, checked: bool):
        self._whitespace_options_row.setVisible(checked)
        self._whitespace_smart_row.setVisible(checked)
        self._whitespace_smart_row2.setVisible(checked)
        margins = self._ws_panel.layout().contentsMargins()
        margins.setBottom(12 if checked else 0)
        self._ws_panel.layout().setContentsMargins(margins)

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
        self._ws_opt_smart_punctuation.setEnabled(checked)
        self._ws_opt_smart_bracket.setEnabled(checked)
        self._ws_opt_smart_alnum.setEnabled(checked)
        self._ws_opt_smart_quote.setEnabled(checked)
        self._ws_opt_protect_ref.setEnabled(checked)
        self._ws_opt_context_conf_spin.setEnabled(checked)

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
                "body_page": self._body_page_spin.value(),
                "section_enabled": self._section_enable_check.isChecked(),
                "sections": {k: cb.isChecked() for k, cb in self._scope_checks.items()},
                "table_enabled": self._table_enable_check.isChecked(),
                "table_layout_mode": self._table_layout_mode_combo.currentData(),
                "table_smart_levels": self._table_smart_levels_combo.currentData(),
                "table_border_mode": self._table_border_mode_combo.currentData(),
                "table_line_spacing_mode": self._table_line_spacing_combo.currentData(),
                "table_repeat_header": self._repeat_header_check.isChecked(),
                "md_cleanup": self._md_cleanup_check.isChecked(),
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
                    "context_min_confidence": self._ws_opt_context_conf_spin.value(),
                },
                "equation_table_format": self._equation_table_check.isChecked(),
                "citation_link_restore": self._citation_link_check.isChecked(),
                "citation_link_options": {
                    "auto_number_reference_entries": self._citation_ref_auto_number_check.isChecked(),
                    "superscript_outer_page_numbers": self._citation_outer_page_sup_check.isChecked(),
                },
                "update_header": self._header_check.isChecked(),
                "update_page_number": self._pagenum_check.isChecked(),
                "update_header_line": self._header_line_check.isChecked(),
                "auto_insert_caption": self._auto_insert_caption_check.isChecked(),
                "format_inserted_caption": self._format_inserted_check.isChecked(),
                "chem_restore": self._chem_restore_check.isChecked(),
                "chem_scopes": {
                    "references": self._chem_scope_refs_check.isChecked(),
                    "body": self._chem_scope_body_check.isChecked(),
                    "abstract": self._chem_scope_abstract_check.isChecked(),
                    "headings": self._chem_scope_headings_check.isChecked(),
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

        body_page = controls.get("body_page")
        if body_page is not None:
            try:
                self._body_page_spin.setValue(int(body_page))
            except (TypeError, ValueError):
                pass

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
                    ws_options.get("context_min_confidence", self._ws_opt_context_conf_spin.value())
                )
            except (TypeError, ValueError):
                conf = self._ws_opt_context_conf_spin.value()
            self._ws_opt_context_conf_spin.setValue(max(1, min(conf, 4)))
        self._equation_table_check.setChecked(
            bool(controls.get("equation_table_format", self._equation_table_check.isChecked()))
        )
        self._citation_link_check.setChecked(
            bool(controls.get("citation_link_restore", self._citation_link_check.isChecked()))
        )
        cite_options = controls.get("citation_link_options", {})
        if isinstance(cite_options, dict):
            self._citation_ref_auto_number_check.setChecked(
                bool(
                    cite_options.get(
                        "auto_number_reference_entries",
                        self._citation_ref_auto_number_check.isChecked(),
                    )
                )
            )
            self._citation_outer_page_sup_check.setChecked(
                bool(
                    cite_options.get(
                        "superscript_outer_page_numbers",
                        self._citation_outer_page_sup_check.isChecked(),
                    )
                )
            )
        self._header_check.setChecked(bool(controls.get("update_header", self._header_check.isChecked())))
        self._pagenum_check.setChecked(bool(controls.get("update_page_number", self._pagenum_check.isChecked())))
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
            self._sync_chem_scope_all_state()

        self._on_table_layout_mode_changed()
        self._on_scope_mode_changed(self._scope_mode_group.checkedId(), True)
        self._on_whitespace_normalize_toggled(self._whitespace_normalize_check.isChecked())
        self._on_whitespace_smart_convert_toggled(self._ws_opt_smart_convert.isChecked())
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
        page = getattr(scope, "body_start_page", None)
        if page is not None and isinstance(page, int) and page >= 1:
            self._body_page_spin.setValue(page)
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
        self._whitespace_normalize_check.setChecked("whitespace_normalize" in pipeline)
        self._equation_table_check.setChecked("equation_table_format" in pipeline)
        self._citation_link_check.setChecked("citation_link" in pipeline)

        self._header_check.setChecked(bool(getattr(cfg, "update_header", False)))
        self._pagenum_check.setChecked(bool(getattr(cfg, "update_page_number", False)))
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

        self._current_scene_path = str(canonical_path)
        self._current_scene_is_custom = custom
        self._refresh_ui_panels()
        self._apply_config_to_controls()
        self._refresh_ui_panels()

        self._run_btn.setEnabled(bool(self._has_valid_doc_path()))

        self._clone_btn.setEnabled(True)
        self._fmt_btn.setEnabled(True)
        self._log(f"已加载{source_label}: {self._config.name}")
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
            self._section_row.setVisible(True)
            self._table_panel.setVisible(True)
            self._scope_sep1.setVisible(True)
            self._scope_sep2.setVisible(True)
            self._ext_row.setVisible(True)
            self._scope_group.setVisible(True)
            self._md_row.setVisible(True)
            self._ws_panel.setVisible(True)
            self._equation_table_row.setVisible(True)
            self._citation_link_row.setVisible(True)
            self._citation_options_row.setVisible(self._citation_link_check.isChecked())
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
        self._auto_insert_caption_check.setVisible(has_caption)
        self._on_header_update_toggled(self._header_check.isChecked())
        self._on_auto_insert_caption_toggled(self._auto_insert_caption_check.isChecked())
        self._ext_label.setVisible(has_header or has_caption)
        self._ext_row.setVisible(has_header or has_caption)
        self._scope_sep2.setVisible(has_header or has_caption)

        self._scope_group.setVisible(
            has_sections or has_header or has_caption
        )
        self._md_row.setVisible(has_md)
        self._lab_sep0.setVisible(
            has_md and (has_whitespace_normalize or has_equation_table or has_citation_link or has_chem_restore)
        )
        self._ws_panel.setVisible(has_whitespace_normalize)
        self._lab_sep_eq.setVisible(
            has_whitespace_normalize and (has_equation_table or has_citation_link or has_chem_restore)
        )
        self._equation_table_row.setVisible(has_equation_table)
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
                page = self._body_page_spin.value()
                if page and page >= 1:
                    scope.body_start_page = page
                    scope.body_start_index = None
                    scope.body_start_keyword = ""
                else:
                    scope.mode = "auto"
            else:
                scope.mode = "auto"
            if self._section_enable_check.isChecked():
                for key, cb in self._scope_checks.items():
                    scope.sections[key] = cb.isChecked()
            else:
                for key in self._scope_checks.keys():
                    scope.sections[key] = False

        # 实验室 pipeline
        self._sync_lab_pipeline_from_checkboxes()

        # 空白字符统一
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
            ws_cfg.context_min_confidence = self._ws_opt_context_conf_spin.value()

        # 参考文献域关联
        citation_cfg = getattr(self._config, "citation_link", None)
        if citation_cfg is not None:
            has_cite = caps.get("citation_link_restore", True)
            citation_cfg.enabled = self._citation_link_check.isChecked() if has_cite else False
            citation_cfg.auto_number_reference_entries = (
                self._citation_ref_auto_number_check.isChecked()
            )
            citation_cfg.superscript_outer_page_numbers = self._citation_outer_page_sup_check.isChecked()

        # 扩展选项
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
            }

    def _save_current_scene_config(self):
        """将当前场景的 UI 修改保存回场景 JSON 文件。"""
        if (
            self._config is None
            or not self._current_scene_path
            or getattr(self, "_is_restoring_state", False)
            or getattr(self, "_suspend_scene_autosave", False)
        ):
            return
        try:
            self._sync_all_controls_to_config()
            scene_path = self._canonicalize_scene_path(
                Path(self._current_scene_path), migrate=True
            )
            self._current_scene_path = str(scene_path)
            save_scene(self._config, scene_path)
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
        if not self._config:
            QMessageBox.warning(
                self, "克隆模板格式", "请先选择一个场景或加载自定义场景。"
            )
            return

        init_dir = os.path.expanduser("~")
        path, _ = QFileDialog.getOpenFileName(
            self, "选择用于克隆格式的 Word 文档", init_dir,
            "Word 文档 (*.docx);;所有文件 (*)",
        )
        if not path:
            return

        # 在 config 的副本上做克隆，不影响当前配置
        import copy
        cloned_config = copy.deepcopy(self._config)

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
                "未匹配到可克隆的样式，已保留当前场景配置。"
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
            msg_parts.append(f"未匹配 {len(missing)} 项（已保留原值）")
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

        # 保存为新文件
        from src.scene.manager import save_scene, PRESETS_DIR
        import re
        safe_stem = re.sub(r'[\\/:*?"<>|]', '_', new_name)
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
        self._config = cloned_config
        self._current_scene_path = str(new_path)
        self._apply_config_to_controls()
        self._refresh_ui_panels()
        self._populate_scene_combo()
        # 选中新场景
        combo = self._scene_combo
        combo.blockSignals(True)
        for i in range(combo.count()):
            if combo.itemText(i) == new_name:
                combo.setCurrentIndex(i)
                break
        combo.blockSignals(False)
        self._save_ui_state()

        QMessageBox.information(
            self, "克隆模板格式",
            f"{msg}\n\n已保存为新格式「{new_name}」")

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

    def _sync_equation_table_pipeline_from_checkbox(self):
        if not self._config:
            return
        caps = self._config.capabilities or {}
        has_equation_table = caps.get("equation_numbering", True)
        pipeline = (
            list(self._config.pipeline)
            if self._config.pipeline is not None else []
        )
        if has_equation_table and self._equation_table_check.isChecked():
            if "equation_table_format" not in pipeline:
                if "table_format" in pipeline:
                    idx = pipeline.index("table_format") + 1
                elif "caption_format" in pipeline:
                    idx = pipeline.index("caption_format") + 1
                else:
                    idx = len(pipeline)
                pipeline.insert(idx, "equation_table_format")
        else:
            pipeline = [s for s in pipeline if s != "equation_table_format"]
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
                if "section_format" in pipeline:
                    idx = pipeline.index("section_format") + 1
                elif "equation_table_format" in pipeline:
                    idx = pipeline.index("equation_table_format") + 1
                elif "table_format" in pipeline:
                    idx = pipeline.index("table_format") + 1
                else:
                    idx = len(pipeline)
                pipeline.insert(idx, "citation_link")
        else:
            pipeline = [s for s in pipeline if s != "citation_link"]
        self._config.pipeline = pipeline

    def _sync_lab_pipeline_from_checkboxes(self):
        self._sync_md_cleanup_pipeline_from_checkbox()
        self._sync_whitespace_normalize_pipeline_from_checkbox()
        self._sync_table_format_pipeline_from_checkbox()
        self._sync_equation_table_pipeline_from_checkbox()
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
                page = self._body_page_spin.value()
                if page and page >= 1:
                    scope.body_start_page = page
                    scope.body_start_index = None
                    scope.body_start_keyword = ""
                else:
                    scope.mode = "auto"
            else:
                scope.mode = "auto"
            
            if self._section_enable_check.isChecked():
                for key, cb in self._scope_checks.items():
                    scope.sections[key] = cb.isChecked()
            else:
                for key in self._scope_checks.keys():
                    scope.sections[key] = False
        else:
            scope.mode = "auto"

        # Markdown 粘贴修复开关（主界面开关优先）
        self._sync_lab_pipeline_from_checkboxes()

        # 空白字符统一（实验室）
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
            ws_cfg.context_min_confidence = self._ws_opt_context_conf_spin.value()

        # 正文参考文献域关联（实验室）
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
            }

        if has_sections:
            scope_desc = "自动识别" if scope.mode == "auto" else f"手动(关键字: {scope.body_start_keyword})"
            enabled = [k for k, v in scope.sections.items() if v]
        else:
            scope_desc = "不启用"
            enabled = []
        self._log(f"编号: {'重建' if mode == 'B' else '保留'}, "
                  f"场景: {self._config.name}")
        self._log(f"范围: {scope_desc}, 分区: {', '.join(enabled)}")
        self._log(
            "流程控制: md_cleanup={}, whitespace_normalize={}, equation_table_format={}, citation_link={}（以实验室开关为准）".format(
                "开启" if "md_cleanup" in (self._config.pipeline or []) else "关闭",
                "开启" if "whitespace_normalize" in (self._config.pipeline or []) else "关闭",
                "开启" if "equation_table_format" in (self._config.pipeline or []) else "关闭",
                "开启" if "citation_link" in (self._config.pipeline or []) else "关闭",
            )
        )
        if "whitespace_normalize" in (self._config.pipeline or []) and ws_cfg is not None:
            ws_options = []
            if ws_cfg.normalize_space_variants:
                ws_options.append("统一特殊空格")
            if ws_cfg.convert_tabs:
                ws_options.append("Tab转空格")
            if ws_cfg.remove_zero_width:
                ws_options.append("清理零宽字符")
            if ws_cfg.collapse_multiple_spaces:
                ws_options.append("折叠连续空格")
            if ws_cfg.trim_paragraph_edges:
                ws_options.append("清理段首尾空格")
            if ws_cfg.smart_full_half_convert:
                ws_options.append("智能全半角")
                smart_opts = []
                if ws_cfg.punctuation_by_context:
                    smart_opts.append("标点按语境")
                if ws_cfg.bracket_by_inner_language:
                    smart_opts.append("括号按语种")
                if ws_cfg.fullwidth_alnum_to_halfwidth:
                    smart_opts.append("全角英数字半角化")
                if ws_cfg.quote_by_context:
                    smart_opts.append("引号按语境")
                if ws_cfg.protect_reference_numbering:
                    smart_opts.append("保护文献编号")
                smart_opts.append(f"阈值={ws_cfg.context_min_confidence}")
                ws_options.append("[" + ", ".join(smart_opts) + "]")
            self._log("空白统一子项: " + ("、".join(ws_options) if ws_options else "（无）"))
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
