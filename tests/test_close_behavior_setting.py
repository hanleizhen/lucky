import os
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

import cnc_program_sheet.main as main_module
from cnc_program_sheet.main import MainWindow


class _CloseEvent:
    def __init__(self) -> None:
        self.accepted = False
        self.ignored = False

    def accept(self) -> None:
        self.accepted = True

    def ignore(self) -> None:
        self.ignored = True


def _window_for_close_test(settings: dict, monkeypatch) -> MainWindow:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.settings = settings
    monkeypatch.setattr(window, "isVisible", lambda: True)
    monkeypatch.setattr(window, "_hide_to_tray", lambda: setattr(window, "_test_hidden_to_tray", True))
    monkeypatch.setattr(
        main_module,
        "QApplication",
        SimpleNamespace(instance=lambda: SimpleNamespace(platformName=lambda: "windows")),
    )
    assert app is not None
    return window


def test_x_button_setting_exits_without_hiding_to_tray(monkeypatch) -> None:
    window = _window_for_close_test({"exit_on_window_close": True}, monkeypatch)
    event = _CloseEvent()

    window.closeEvent(event)

    assert event.accepted
    assert not event.ignored
    assert not getattr(window, "_test_hidden_to_tray", False)
    window.tray_icon.hide()
    window.hide()


def test_unchecked_x_button_setting_minimizes_to_tray(monkeypatch) -> None:
    window = _window_for_close_test({"exit_on_window_close": False}, monkeypatch)
    event = _CloseEvent()

    window.closeEvent(event)

    assert event.ignored
    assert not event.accepted
    assert getattr(window, "_test_hidden_to_tray", False)
    window.tray_icon.hide()
    window.hide()
