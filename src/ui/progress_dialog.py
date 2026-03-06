"""进度对话框：显示排版 Pipeline 执行进度"""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QProgressBar, QPushButton
)
from PySide6.QtCore import Qt, Signal


class ProgressDialog(QDialog):
    """模态进度对话框"""
    cancel_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("排版进行中...")
        self.setFixedSize(420, 150)
        self.setWindowFlags(
            self.windowFlags() & ~Qt.WindowContextHelpButtonHint
        )

        layout = QVBoxLayout(self)

        self.label = QLabel("准备中...")
        layout.addWidget(self.label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.reject)
        layout.addWidget(self.cancel_btn, alignment=Qt.AlignRight)

        self._cancelled = False

    def update_progress(self, current: int, total: int, message: str):
        """更新进度条和文本"""
        if total > 0:
            pct = int(current / total * 100)
            self.progress_bar.setValue(pct)
        self.label.setText(message)

    def reject(self):
        if self._cancelled:
            return
        self._cancelled = True
        self.label.setText("正在取消，请稍候...")
        self.cancel_btn.setEnabled(False)
        self.cancel_requested.emit()

    @property
    def cancelled(self) -> bool:
        return self._cancelled
