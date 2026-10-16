import os
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

import cnc_program_sheet.main as main_module


class _RecordedDialog:
    created: list[tuple[str, str, dict]] = []

    def __init__(self, title: str, message: str, **kwargs: object) -> None:
        self.created.append((title, message, dict(kwargs)))

    def exec(self) -> int:
        return 1


def test_release_notes_are_shown_once_per_installed_version(monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    settings: dict[str, object] = {}
    saved: list[dict[str, object]] = []
    _RecordedDialog.created.clear()
    monkeypatch.setattr(main_module, "load_settings", lambda: settings)
    monkeypatch.setattr(main_module, "save_settings", lambda current: saved.append(dict(current)))
    monkeypatch.setattr(
        main_module,
        "default_user_template",
        lambda: Path("assets/CNC程序单.xlsx").resolve(),
    )
    monkeypatch.setattr(main_module, "ApplicationDialog", _RecordedDialog)

    window = main_module.MainWindow()
    try:
        window.show_release_notes_if_needed()
        window.show_release_notes_if_needed()

        assert len(_RecordedDialog.created) == 1
        title, message, kwargs = _RecordedDialog.created[0]
        assert f"v{main_module.__version__}" in title
        assert "本次更新内容" in message
        assert "同一版本只显示一次" in message
        assert kwargs["parent"] is window
        assert settings["release_notes_seen_version"] == main_module.__version__
        assert saved
    finally:
        window.tray_icon.hide()
        window.close()
    assert app is not None


def test_seen_version_suppresses_the_startup_release_note(monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    settings: dict[str, object] = {"release_notes_seen_version": main_module.__version__}
    _RecordedDialog.created.clear()
    monkeypatch.setattr(main_module, "load_settings", lambda: settings)
    monkeypatch.setattr(main_module, "save_settings", lambda _current: None)
    monkeypatch.setattr(
        main_module,
        "default_user_template",
        lambda: Path("assets/CNC程序单.xlsx").resolve(),
    )
    monkeypatch.setattr(main_module, "ApplicationDialog", _RecordedDialog)

    window = main_module.MainWindow()
    try:
        window.show_release_notes_if_needed()
        assert _RecordedDialog.created == []
    finally:
        window.tray_icon.hide()
        window.close()
    assert app is not None


def test_an_older_seen_version_shows_the_new_version_note(monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    settings: dict[str, object] = {"release_notes_seen_version": "1.1.46"}
    _RecordedDialog.created.clear()
    monkeypatch.setattr(main_module, "load_settings", lambda: settings)
    monkeypatch.setattr(main_module, "save_settings", lambda _current: None)
    monkeypatch.setattr(
        main_module,
        "default_user_template",
        lambda: Path("assets/CNC程序单.xlsx").resolve(),
    )
    monkeypatch.setattr(main_module, "ApplicationDialog", _RecordedDialog)

    window = main_module.MainWindow()
    try:
        window.show_release_notes_if_needed()
        assert len(_RecordedDialog.created) == 1
        assert settings["release_notes_seen_version"] == main_module.__version__
    finally:
        window.tray_icon.hide()
        window.close()
    assert app is not None
