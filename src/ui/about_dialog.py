import os
from pathlib import Path
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
    QFrame
)
from PySide6.QtCore import Qt, QUrl, QSize
from PySide6.QtGui import QIcon, QDesktopServices, QPixmap, QPainter
from PySide6.QtSvg import QSvgRenderer

class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关于")
        self.setFixedSize(450, 340)
        
        self._build_ui()

    @staticmethod
    def _load_svg_icon(svg_path: Path, size: int = 18):
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
        
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)
        
        # Header Area
        header_layout = QHBoxLayout()
        
        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(4)
        
        app_name = QLabel("Lark-Formatter")
        app_name.setStyleSheet("font-size: 20px; font-weight: bold;")
        
        version = QLabel("Version 0.1.0")
        version.setStyleSheet("color: #888888; font-size: 13px;")
        
        title_vbox.addWidget(app_name)
        title_vbox.addWidget(version)
        title_vbox.addStretch()
        
        header_layout.addLayout(title_vbox)
        header_layout.addStretch()
        layout.addLayout(header_layout)
        
        # Description
        desc = QLabel("一款专注于学术论文、公文排版的自动化桌面利器。")
        desc.setWordWrap(True)
        desc.setStyleSheet("font-size: 14px; margin-top: 8px;")
        layout.addWidget(desc)
        
        # Separator
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #e1e4e8; background-color: #e1e4e8; margin-top: 10px; margin-bottom: 10px;")
        layout.addWidget(sep)
        
        # Credits & License
        credits_layout = QVBoxLayout()
        credits_layout.setSpacing(6)
        
        author = QLabel("Developed by 王云雀Alouette")
        author.setStyleSheet("font-size: 13px;")
        
        thanks_label = QLabel("Thanks to 阿七.")
        thanks_label.setStyleSheet("font-size: 13px;")

        license_label = QLabel("Released under the MIT License.")
        license_label.setStyleSheet("font-size: 13px; color: #888888;")
        
        tech_label = QLabel("Powered by Python 3, PySide6, and python-docx.")
        tech_label.setStyleSheet("font-size: 13px; color: #888888;")
        
        credits_layout.addWidget(author)
        credits_layout.addWidget(thanks_label)
        credits_layout.addWidget(license_label)
        credits_layout.addWidget(tech_label)
        layout.addLayout(credits_layout)
        
        layout.addStretch()
        
        # Action Buttons
        btn_layout = QHBoxLayout()
        
        base_dir = Path(__file__).parent

        self.github_btn = QPushButton(" 访问 GitHub")
        self.github_btn.setCursor(Qt.PointingHandCursor)
        self.github_btn.setFixedHeight(32)
        self.github_btn.setIcon(self._load_svg_icon(base_dir / "icons" / "github.svg"))
        self.github_btn.clicked.connect(self._open_github)

        self.bilibili_btn = QPushButton(" 访问 Bilibili")
        self.bilibili_btn.setCursor(Qt.PointingHandCursor)
        self.bilibili_btn.setFixedHeight(32)
        self.bilibili_btn.setIcon(self._load_svg_icon(base_dir / "icons" / "bilibili.svg"))
        self.bilibili_btn.clicked.connect(self._open_bilibili)
        
        self.close_btn = QPushButton("确定")
        self.close_btn.setFixedHeight(32)
        self.close_btn.clicked.connect(self.accept)
        self.close_btn.setDefault(True)
        
        btn_layout.addWidget(self.github_btn)
        btn_layout.addWidget(self.bilibili_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.close_btn)
        
        layout.addLayout(btn_layout)
        
    def _open_github(self):
        QDesktopServices.openUrl(QUrl("https://github.com/Alouetter/Lark-Formatter"))

    def _open_bilibili(self):
        QDesktopServices.openUrl(QUrl("https://space.bilibili.com/10825084"))
