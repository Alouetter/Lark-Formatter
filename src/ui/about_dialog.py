from collections.abc import Callable
from pathlib import Path

from PySide6.QtCore import QSize, Qt, QTimer, QUrl
from PySide6.QtGui import QDesktopServices, QIcon, QPainter, QPixmap
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from src.utils.app_meta import APP_NAME, APP_VERSION


_WINDOW_TITLE = "\u5173\u4e8e"
_APP_DESC = (
    "\u4e00\u6b3e\u805a\u7126\u6bd5\u4e1a\u8bba\u6587\u683c\u5f0f\u4fee\u8ba2\u7684\u4e13\u9879\u684c\u9762\u5de5\u5177\u3002"
    "0.20 LTS \u4ec5\u652f\u6301\u8be5\u573a\u666f\uff0c\u5176\u4ed6\u573a\u666f\u8ba1\u5212\u5728 1.00 \u7248\u672c\u5b8c\u6210\u5f00\u53d1\u4e0e\u53d1\u5e03\u3002"
)
_CREDITS_TEXT = (
    "Developed by \u738b\u4e91\u96c0Alouette. | "
    "Thanks to \u963f\u4e03."
)
_VISIT_GITHUB_TEXT = " \u8bbf\u95ee GitHub"
_VISIT_BILIBILI_TEXT = " \u8bbf\u95ee Bilibili"
_CONFIRM_TEXT = "\u786e\u5b9a"
_FORMAT_LOADING_TEXT = "\u672a\u52a0\u8f7d"
_FORMAT_UNSIGNED_TEXT = "\u672a\u7f72\u540d"
_FORMAT_LINE_TEMPLATE = (
    "\u683c\u5f0f\u540d\u79f0\uff1a{format_name} | "
    "\u683c\u5f0f\u914d\u7f6e\uff1a{signature}"
)
_ADD_SIGNATURE_TEXT = "\u589e\u52a0\u7f72\u540d"
_SIGNED_SIGNATURE_TEXT = "\u5df2\u6c38\u4e45\u7f72\u540d"
_UNSIGNABLE_SIGNATURE_TEXT = "\u5f53\u524d\u683c\u5f0f\u4e0d\u53ef\u7f72\u540d"
_SIGNATURE_DIALOG_TITLE = "\u683c\u5f0f\u7f72\u540d"
_SIGNATURE_DIALOG_LABEL = "\u8bf7\u8f93\u5165\u5f53\u524d\u683c\u5f0f\u914d\u7f6e\u7684\u7f72\u540d\uff1a"
_CANCEL_TEXT = "\u53d6\u6d88"
_SIGNATURE_EMPTY_TEXT = "\u7f72\u540d\u4e0d\u80fd\u4e3a\u7a7a\u3002"
_SIGNATURE_LOCKED_TEXT = "\u5f53\u524d\u683c\u5f0f\u5df2\u5b58\u5728\u6c38\u4e45\u7f72\u540d\uff0c\u4e0d\u80fd\u4fee\u6539\u3002"
_SIGNATURE_CONFIRM_TITLE = "\u786e\u8ba4\u5199\u5165\u683c\u5f0f\u7f72\u540d"
_SIGNATURE_CONFIRM_TEXT = (
    "\u683c\u5f0f\u7f72\u540d\u5199\u5165\u540e\u5c06\u6c38\u4e45\u4e0d\u53ef\u4fee\u6539\u3002\n\n"
    "\u786e\u8ba4\u5c06\u201c{signature}\u201d\u5199\u5165\u5f53\u524d\u683c\u5f0f\u5417\uff1f"
)


class AboutDialog(QDialog):
    def __init__(
        self,
        parent=None,
        *,
        format_name: str = "",
        format_signature: str = "",
        format_signable: bool = True,
        format_signable_reason: str = "",
        format_unsignable_text: str = "",
        on_update_format_signature: Callable[[str], str | None] | None = None,
    ):
        super().__init__(parent)
        self.setWindowTitle(_WINDOW_TITLE)
        self.setFixedSize(480, 360)

        self._format_name = str(format_name or "").strip()
        self._format_signature = str(format_signature or "").strip()
        self._format_signable = bool(format_signable)
        self._format_signable_reason = str(format_signable_reason or "").strip()
        self._format_unsignable_text = str(format_unsignable_text or "").strip()
        self._on_update_format_signature = on_update_format_signature

        self._build_ui()

    @staticmethod
    def _load_svg_icon(svg_path: Path, size: int = 18) -> QIcon:
        """Render an SVG file into a QIcon via QSvgRenderer."""
        if not svg_path.exists():
            return QIcon()

        renderer = QSvgRenderer(str(svg_path))
        pixmap = QPixmap(QSize(size, size))
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()
        return QIcon(pixmap)

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        header_layout = QHBoxLayout()
        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(4)

        app_name = QLabel(APP_NAME)
        app_name.setStyleSheet("font-size: 20px; font-weight: bold;")

        version = QLabel(f"Version {APP_VERSION}")
        version.setStyleSheet("color: #888888; font-size: 13px;")

        title_vbox.addWidget(app_name)
        title_vbox.addWidget(version)
        title_vbox.addStretch()

        header_layout.addLayout(title_vbox)
        header_layout.addStretch()
        layout.addLayout(header_layout)

        desc = QLabel(_APP_DESC)
        desc.setWordWrap(True)
        desc.setStyleSheet("font-size: 14px; margin-top: 8px;")
        layout.addWidget(desc)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(
            "color: #e1e4e8; background-color: #e1e4e8; "
            "margin-top: 10px; margin-bottom: 10px;"
        )
        layout.addWidget(sep)

        credits_layout = QVBoxLayout()
        credits_layout.setSpacing(6)

        format_row = QHBoxLayout()
        format_row.setSpacing(10)

        self._format_signature_label = QLabel()
        self._format_signature_label.setObjectName("AboutFormatSignatureLabel")
        self._format_signature_label.setStyleSheet("font-size: 13px; color: #888888;")

        self._format_signature_btn = QPushButton()
        self._format_signature_btn.setObjectName("AboutFormatSignatureButton")
        self._format_signature_btn.setCursor(Qt.PointingHandCursor)
        self._format_signature_btn.setFixedHeight(28)
        self._format_signature_btn.clicked.connect(self._edit_format_signature)

        format_row.addWidget(self._format_signature_label, 1)
        format_row.addWidget(self._format_signature_btn)

        credits_label = QLabel(_CREDITS_TEXT)
        credits_label.setObjectName("AboutCreditsLabel")
        credits_label.setStyleSheet("font-size: 13px; color: #888888;")

        license_label = QLabel("Released under the MIT License.")
        license_label.setStyleSheet("font-size: 13px; color: #888888;")

        tech_label = QLabel("Powered by Python 3, PySide6, and python-docx.")
        tech_label.setStyleSheet("font-size: 13px; color: #888888;")

        credits_layout.addLayout(format_row)
        credits_layout.addWidget(credits_label)
        credits_layout.addWidget(license_label)
        credits_layout.addWidget(tech_label)
        layout.addLayout(credits_layout)
        self._refresh_format_signature_ui()

        layout.addStretch()

        btn_layout = QHBoxLayout()
        base_dir = Path(__file__).parent

        self.github_btn = QPushButton(_VISIT_GITHUB_TEXT)
        self.github_btn.setCursor(Qt.PointingHandCursor)
        self.github_btn.setFixedHeight(32)
        self.github_btn.setIcon(self._load_svg_icon(base_dir / "icons" / "github.svg"))
        self.github_btn.clicked.connect(self._open_github)

        self.bilibili_btn = QPushButton(_VISIT_BILIBILI_TEXT)
        self.bilibili_btn.setCursor(Qt.PointingHandCursor)
        self.bilibili_btn.setFixedHeight(32)
        self.bilibili_btn.setIcon(self._load_svg_icon(base_dir / "icons" / "bilibili.svg"))
        self.bilibili_btn.clicked.connect(self._open_bilibili)

        self.close_btn = QPushButton(_CONFIRM_TEXT)
        self.close_btn.setFixedHeight(32)
        self.close_btn.clicked.connect(self.accept)
        self.close_btn.setDefault(True)

        btn_layout.addWidget(self.github_btn)
        btn_layout.addWidget(self.bilibili_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.close_btn)

        layout.addLayout(btn_layout)

    def _refresh_format_signature_ui(self) -> None:
        format_name = self._format_name or _FORMAT_LOADING_TEXT
        signature_text = self._format_signature or _FORMAT_UNSIGNED_TEXT
        self._format_signature_label.setText(
            _FORMAT_LINE_TEMPLATE.format(
                format_name=format_name,
                signature=signature_text,
            )
        )

        has_signature = bool(self._format_signature)
        if has_signature:
            btn_text = _SIGNED_SIGNATURE_TEXT
        elif not self._format_signable:
            btn_text = self._format_unsignable_text or _UNSIGNABLE_SIGNATURE_TEXT
        else:
            btn_text = _ADD_SIGNATURE_TEXT
        self._format_signature_btn.setText(btn_text)
        self._format_signature_btn.setEnabled(
            self._on_update_format_signature is not None
            and bool(self._format_name)
            and self._format_signable
            and not has_signature
        )

    def _confirm_format_signature(self, signature: str) -> bool:
        reply = QMessageBox.question(
            self,
            _SIGNATURE_CONFIRM_TITLE,
            _SIGNATURE_CONFIRM_TEXT.format(signature=signature),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        return reply == QMessageBox.Yes

    def _prompt_format_signature(self) -> tuple[str, bool]:
        dialog = QInputDialog(self)
        dialog.setInputMode(QInputDialog.TextInput)
        dialog.setWindowTitle(_SIGNATURE_DIALOG_TITLE)
        dialog.setLabelText(_SIGNATURE_DIALOG_LABEL)
        dialog.setTextValue(self._format_signature)
        dialog.setOkButtonText(_CONFIRM_TEXT)
        dialog.setCancelButtonText(_CANCEL_TEXT)
        dialog.setMinimumWidth(420)

        line_edit = dialog.findChild(QLineEdit)
        if line_edit is not None:
            QTimer.singleShot(0, line_edit.selectAll)

        accepted = dialog.exec() == QDialog.Accepted
        return dialog.textValue(), accepted

    def _edit_format_signature(self) -> None:
        if self._on_update_format_signature is None:
            return
        if not self._format_signable:
            if self._format_signable_reason:
                QMessageBox.information(
                    self,
                    _SIGNATURE_DIALOG_TITLE,
                    self._format_signable_reason,
                )
            return
        if self._format_signature:
            QMessageBox.information(self, _SIGNATURE_DIALOG_TITLE, _SIGNATURE_LOCKED_TEXT)
            return

        new_signature, ok = self._prompt_format_signature()
        if not ok:
            return

        normalized = str(new_signature or "").strip()
        if not normalized:
            QMessageBox.warning(self, _SIGNATURE_DIALOG_TITLE, _SIGNATURE_EMPTY_TEXT)
            return
        if not self._confirm_format_signature(normalized):
            return

        try:
            updated_signature = self._on_update_format_signature(normalized)
        except Exception as exc:
            QMessageBox.critical(self, _SIGNATURE_DIALOG_TITLE, str(exc))
            return

        self._format_signature = str(updated_signature or normalized).strip()
        self._refresh_format_signature_ui()

    def _open_github(self) -> None:
        QDesktopServices.openUrl(QUrl("https://github.com/Alouetter/Lark-Formatter"))

    def _open_bilibili(self) -> None:
        QDesktopServices.openUrl(QUrl("https://space.bilibili.com/10825084"))
