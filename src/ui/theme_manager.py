import os
from pathlib import Path
from PySide6.QtWidgets import QApplication

class ThemeManager:
    _current_theme = "light"
    _themes_dir = Path(__file__).parent / "themes"

    @classmethod
    def load_theme(cls, theme_name: str) -> str:
        """Reads the QSS file for the given theme."""
        qss_path = cls._themes_dir / f"{theme_name}.qss"
        if not qss_path.exists():
            return ""
        try:
            with open(qss_path, "r", encoding="utf-8") as f:
                return f.read()
        except Exception as e:
            print(f"Error loading theme {theme_name}: {e}")
            return ""

    @classmethod
    def apply_theme(cls, app: QApplication, theme_name: str = "light"):
        """Applies the theme to the entire application."""
        qss = cls.load_theme(theme_name)
        if qss:
            app.setStyleSheet(qss)
            cls._current_theme = theme_name
            
    @classmethod
    def toggle_theme(cls, app: QApplication):
        """Toggles between light and dark themes."""
        new_theme = "dark" if cls._current_theme == "light" else "light"
        cls.apply_theme(app, new_theme)
        return new_theme

    @classmethod
    def get_current_theme(cls) -> str:
        return cls._current_theme
