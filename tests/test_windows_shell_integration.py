from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import cnc_program_sheet.main as app_main
from PIL import Image
from cnc_program_sheet.nc_transform import M06, M304, TG, is_explorer_transform_request, run_explorer_transform
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


def test_application_icon_has_a_genuinely_transparent_background() -> None:
    """The desktop icon must not carry the old white rounded-square backdrop."""

    for filename in ("cnc_smart_sheet_icon.png", "cnc_smart_sheet_icon.ico"):
        with Image.open(ROOT / "assets" / filename) as image:
            alpha = image.convert("RGBA").getchannel("A")
            assert alpha.getextrema()[0] == 0
            assert alpha.getpixel((0, 0)) == 0


def test_explorer_context_menu_commands_are_registered_per_user_for_multi_select() -> None:
    installer = (ROOT / "installer.iss").read_text(encoding="utf-8")

    for mode, label in ((M06, "M06 修改"), (M304, "M304 修改"), (TG, "铜工修改")):
        assert "Root: HKCU; Subkey: \"Software\\Classes\\SystemFileAssociations\\.nc\\shell\\CNCProgramSheet." in installer
        assert f"CNC 智能程序单 - {label}" in installer
        assert f'--shell-transform {mode.value} %*"' in installer
    assert installer.count('ValueName: "MultiSelectModel"; ValueData: "Player"') == 3


def test_explorer_transform_runs_without_gui_and_applies_selected_files(tmp_path: Path) -> None:
    first = tmp_path / "one.NC"
    second = tmp_path / "two.nc"
    first.write_text("N10 M6\nN20 G0 X0\n", encoding="ascii")
    second.write_text("N30 M3\n", encoding="ascii")

    assert is_explorer_transform_request(["--shell-transform", "M06", str(first)])
    assert run_explorer_transform(["--shell-transform", "M06", str(first), str(second)]) == 0
    assert first.read_bytes() == b"N20 G0 X0\r\n"
    assert second.read_bytes() == b"N30 M3\r\n"


def test_explorer_transform_allows_a_dash_prefixed_nc_filename(tmp_path: Path) -> None:
    source = tmp_path / "-one.NC"
    source.write_text("N10 M6\n", encoding="ascii")

    assert run_explorer_transform(["--shell-transform=M06", str(source)]) == 0
    assert source.read_bytes() == b"\r\n"


def test_explorer_transform_rejects_non_nc_documents(tmp_path: Path, monkeypatch) -> None:
    document = tmp_path / "not_nc.txt"
    document.write_text("M6", encoding="ascii")
    logged: list[str] = []

    import cnc_program_sheet.app_paths as app_paths

    monkeypatch.setattr(app_paths, "log_exception", lambda message: logged.append(message))
    assert run_explorer_transform(["--shell-transform", "M06", str(document)]) == 1
    assert document.read_text(encoding="ascii") == "M6"
    assert "不是 .NC / .nc 文件" in logged[0]
