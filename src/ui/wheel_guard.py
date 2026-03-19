"""Application-wide wheel guard for value-editing widgets."""

from __future__ import annotations

from PySide6.QtCore import QEvent, QObject
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QAbstractSpinBox,
    QComboBox,
    QWidget,
)


class WheelValueGuard(QObject):
    """Block mouse-wheel value changes on combo boxes and spin boxes."""

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:
        if event.type() != QEvent.Wheel:
            return False

        if not isinstance(watched, QWidget):
            return False

        if self._should_ignore_wheel(watched):
            event.ignore()
            return True
        return False

    @staticmethod
    def _should_ignore_wheel(widget: QWidget) -> bool:
        current: QWidget | None = widget
        while current is not None:
            if isinstance(current, QAbstractItemView):
                return False
            if isinstance(current, (QComboBox, QAbstractSpinBox)):
                return True
            parent = current.parentWidget()
            current = parent if isinstance(parent, QWidget) else None
        return False


def install_global_wheel_guard(app: QApplication | None = None) -> None:
    """Install a single app-wide wheel guard; safe to call repeatedly."""

    app = app or QApplication.instance()
    if app is None:
        return
    if bool(app.property("_wheel_value_guard_installed")):
        return

    guard = WheelValueGuard(app)
    app.installEventFilter(guard)
    app.setProperty("_wheel_value_guard_installed", True)
    setattr(app, "_wheel_value_guard", guard)
