"""Application entrypoint."""

from __future__ import annotations

import sys
import traceback
from pathlib import Path

# Ensure project root is importable regardless of launch location.
_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from src.utils.app_meta import APP_NAME, APP_USER_MODEL_ID, APP_VERSION


def main() -> None:
    from PySide6.QtWidgets import QApplication, QMessageBox
    from PySide6.QtGui import QIcon
    from src.ui.main_window import MainWindow
    from src.ui.theme_manager import ThemeManager
    from src.ui.wheel_guard import install_global_wheel_guard

    # Force Windows to treat this as a separate app from python.exe
    # This ensures the taskbar uses our custom icon instead of the Python logo
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass

    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    install_global_wheel_guard(app)
    
    icon_path = _ROOT / "src" / "ui" / "icons" / "app_icon.png"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    # Load initial theme
    ThemeManager.apply_theme(app, "light")

    try:
        window = MainWindow()
        window.show()
        sys.exit(app.exec())
    except Exception:
        QMessageBox.critical(None, "启动失败", traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        err = traceback.format_exc()
        log_path = _ROOT / "crash.log"
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(err)
        print("=" * 50)
        print("启动失败，错误信息已写入 crash.log")
        print("=" * 50)
        print(err)
        try:
            import os
            os.system("pause")
        except Exception:
            pass
