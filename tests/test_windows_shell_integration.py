from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import cnc_program_sheet.main as app_main
from cnc_program_sheet.version import APP_USER_MODEL_ID


ROOT = Path(__file__).resolve().parents[1]


def test_windows_app_user_model_id_uses_the_stable_shell_identity(monkeypatch) -> None:
    calls: list[str] = []
    fake_shell32 = SimpleNamespace(SetCurrentProcessExplicitAppUserModelID=calls.append)

    monkeypatch.setattr(app_main.sys, "platform", "win32")
    monkeypatch.setattr(app_main.ctypes, "windll", SimpleNamespace(shell32=fake_shell32), raising=False)

    app_main._set_windows_app_user_model_id()

    assert calls == [APP_USER_MODEL_ID]


def test_installer_recreates_shortcuts_with_a_versioned_standalone_icon() -> None:
    installer = (ROOT / "installer.iss").read_text(encoding="utf-8")

    versioned_icon = "CNCProgramSheet_icon_{#MyAppVersion}.ico"
    assert f'Source: "assets\\cnc_smart_sheet_icon.ico"; DestDir: "{{app}}"; DestName: "{versioned_icon}"' in installer
    assert installer.count(f'IconFilename: "{{app}}\\{versioned_icon}"') == 2
    assert installer.count(f'AppUserModelID: "{APP_USER_MODEL_ID}"') == 2
    assert 'Type: files; Name: "{autodesktop}\\CNC 智能程序单.lnk"' in installer
    assert 'Type: files; Name: "{autoprograms}\\CNC 智能程序单.lnk"' in installer
