import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QLabel, QMessageBox, QPushButton

from src.ui.about_dialog import AboutDialog


def _app():
    return QApplication.instance() or QApplication([])


def test_about_dialog_shows_combined_credits_and_signature_row():
    _app()
    dlg = AboutDialog(
        format_name="\u9ed8\u8ba4\u683c\u5f0f",
        format_signature="",
        on_update_format_signature=lambda value: value,
    )

    credits_label = dlg.findChild(QLabel, "AboutCreditsLabel")
    format_label = dlg.findChild(QLabel, "AboutFormatSignatureLabel")
    format_button = dlg.findChild(QPushButton, "AboutFormatSignatureButton")

    assert credits_label is not None
    assert credits_label.text() == (
        "Developed by \u738b\u4e91\u96c0Alouette. | "
        "Thanks to \u963f\u4e03."
    )
    assert format_label is not None
    assert format_label.text() == (
        "\u683c\u5f0f\u914d\u7f6e\uff1a\u9ed8\u8ba4\u683c\u5f0f | "
        "\u683c\u5f0f\u7f72\u540d\uff1a\u672a\u7f72\u540d"
    )
    assert format_button is not None
    assert format_button.text() == "\u589e\u52a0\u7f72\u540d"
    assert format_button.isEnabled() is True


def test_about_dialog_edit_signature_requires_confirmation(monkeypatch):
    _app()
    updates = []
    dlg = AboutDialog(
        format_name="\u9ed8\u8ba4\u683c\u5f0f",
        format_signature="",
        on_update_format_signature=lambda value: updates.append(value) or value,
    )

    monkeypatch.setattr(
        dlg,
        "_prompt_format_signature",
        lambda: ("\u674e\u56db", True),
    )
    monkeypatch.setattr(
        QMessageBox,
        "question",
        staticmethod(lambda *args, **kwargs: QMessageBox.No),
    )

    dlg._edit_format_signature()

    format_label = dlg.findChild(QLabel, "AboutFormatSignatureLabel")
    format_button = dlg.findChild(QPushButton, "AboutFormatSignatureButton")

    assert updates == []
    assert format_label is not None
    assert format_label.text() == (
        "\u683c\u5f0f\u914d\u7f6e\uff1a\u9ed8\u8ba4\u683c\u5f0f | "
        "\u683c\u5f0f\u7f72\u540d\uff1a\u672a\u7f72\u540d"
    )
    assert format_button is not None
    assert format_button.text() == "\u589e\u52a0\u7f72\u540d"
    assert format_button.isEnabled() is True


def test_about_dialog_edit_signature_updates_text_and_locks_button(monkeypatch):
    _app()
    updates = []
    dlg = AboutDialog(
        format_name="\u9ed8\u8ba4\u683c\u5f0f",
        format_signature="",
        on_update_format_signature=lambda value: updates.append(value) or value,
    )

    monkeypatch.setattr(
        dlg,
        "_prompt_format_signature",
        lambda: ("\u674e\u56db", True),
    )
    monkeypatch.setattr(
        QMessageBox,
        "question",
        staticmethod(lambda *args, **kwargs: QMessageBox.Yes),
    )

    dlg._edit_format_signature()

    format_label = dlg.findChild(QLabel, "AboutFormatSignatureLabel")
    format_button = dlg.findChild(QPushButton, "AboutFormatSignatureButton")

    assert updates == ["\u674e\u56db"]
    assert format_label is not None
    assert format_label.text() == (
        "\u683c\u5f0f\u914d\u7f6e\uff1a\u9ed8\u8ba4\u683c\u5f0f | "
        "\u683c\u5f0f\u7f72\u540d\uff1a\u674e\u56db"
    )
    assert format_button is not None
    assert format_button.text() == "\u5df2\u6c38\u4e45\u7f72\u540d"
    assert format_button.isEnabled() is False


def test_about_dialog_with_existing_signature_is_locked():
    _app()
    dlg = AboutDialog(
        format_name="\u9ed8\u8ba4\u683c\u5f0f",
        format_signature="\u5f20\u4e09",
        on_update_format_signature=lambda value: value,
    )

    format_label = dlg.findChild(QLabel, "AboutFormatSignatureLabel")
    format_button = dlg.findChild(QPushButton, "AboutFormatSignatureButton")

    assert format_label is not None
    assert format_label.text() == (
        "\u683c\u5f0f\u914d\u7f6e\uff1a\u9ed8\u8ba4\u683c\u5f0f | "
        "\u683c\u5f0f\u7f72\u540d\uff1a\u5f20\u4e09"
    )
    assert format_button is not None
    assert format_button.text() == "\u5df2\u6c38\u4e45\u7f72\u540d"
    assert format_button.isEnabled() is False


def test_about_dialog_with_unsignable_format_disables_signature_button():
    _app()
    dlg = AboutDialog(
        format_name="\u9ed8\u8ba4\u683c\u5f0f",
        format_signature="",
        format_signable=False,
        format_signable_reason="\u9ed8\u8ba4\u683c\u5f0f\u4e0d\u53ef\u7f72\u540d",
        format_unsignable_text="\u9ed8\u8ba4\u683c\u5f0f\u4e0d\u53ef\u7f72\u540d",
        on_update_format_signature=lambda value: value,
    )

    format_button = dlg.findChild(QPushButton, "AboutFormatSignatureButton")

    assert format_button is not None
    assert format_button.text() == "\u9ed8\u8ba4\u683c\u5f0f\u4e0d\u53ef\u7f72\u540d"
    assert format_button.isEnabled() is False
