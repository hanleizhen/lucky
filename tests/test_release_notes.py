import os
from pathlib import Path
from types import SimpleNamespace

import pytest

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
        assert "只会在首次打开本版本时显示一次" in message
        assert "打印清晰度" in message
        assert "点击窗口 X" not in message
        assert "Mastercam NC 默认文件夹" not in message
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


def test_mastercam_button_installs_only_after_user_enables_the_setting(monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    settings: dict[str, object] = {"install_mastercam_button": True}
    installed: list[bool] = []

    class InfoDialog:
        created: list[bool] = []

        def __init__(self, *_args: object, **_kwargs: object) -> None:
            self.created.append(True)
        def exec(self) -> int:
            return 0

    monkeypatch.setattr(main_module, "load_settings", lambda: settings)
    monkeypatch.setattr(main_module, "save_settings", lambda _current: None)
    monkeypatch.setattr(
        main_module,
        "default_user_template",
        lambda: Path("assets/CNC程序单.xlsx").resolve(),
    )
    monkeypatch.setattr(
        main_module,
        "install_mastercam_addin",
        lambda: (installed.append(True), SimpleNamespace(installed_years=(2024,)))[1],
    )
    monkeypatch.setattr(
        main_module,
        "resource_path",
        lambda *_parts: Path("assets/mastercam_addin/GenerateProgramSheet.csx").resolve(),
    )
    monkeypatch.setattr(main_module, "ApplicationDialog", InfoDialog)

    window = main_module.MainWindow()
    try:
        InfoDialog.created.clear()
        window.install_mastercam_addin_if_enabled()
        window.install_mastercam_addin_if_enabled()

        assert len(installed) == 1
        assert InfoDialog.created == []
        assert settings["mastercam_addin_installed_version"] == main_module.__version__
    finally:
        window.tray_icon.hide()
        window.close()
    assert app is not None


def test_mastercam_install_success_popup_can_be_enabled_in_settings(monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    settings: dict[str, object] = {
        "install_mastercam_button": True,
        "show_mastercam_install_success": True,
    }
    dialogs: list[str] = []

    class InfoDialog:
        def __init__(self, title: str, *_args: object, **_kwargs: object) -> None:
            dialogs.append(title)

        def exec(self) -> int:
            return 0

    monkeypatch.setattr(main_module, "load_settings", lambda: settings)
    monkeypatch.setattr(main_module, "save_settings", lambda _current: None)
    monkeypatch.setattr(
        main_module,
        "default_user_template",
        lambda: Path("assets/CNC程序单.xlsx").resolve(),
    )
    monkeypatch.setattr(
        main_module,
        "install_mastercam_addin",
        lambda: SimpleNamespace(installed_years=(2024,)),
    )
    monkeypatch.setattr(
        main_module,
        "resource_path",
        lambda *_parts: Path("assets/mastercam_addin/GenerateProgramSheet.csx").resolve(),
    )
    monkeypatch.setattr(main_module, "ApplicationDialog", InfoDialog)

    window = main_module.MainWindow()
    try:
        window.install_mastercam_addin_if_enabled()
        assert dialogs == ["Mastercam 按钮安装完成"]
    finally:
        window.tray_icon.hide()
        window.close()
    assert app is not None
