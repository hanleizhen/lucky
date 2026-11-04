from pathlib import Path
from types import SimpleNamespace

import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QDialog

from cnc_program_sheet import mastercam_install
import cnc_program_sheet.main as main_module


def _fake_mastercam_tree(root: Path, docs: Path, year: int) -> Path:
    install = root / f"Mastercam {year}"
    install.mkdir(parents=True)
    (install / "Mastercam.exe").write_bytes(b"mastercam")
    profile = docs / f"My Mastercam {year}" / "Mastercam"
    profile.mkdir(parents=True)
    return profile / "Add-Ins"


def test_discovery_finds_each_supported_year_and_ignores_others(tmp_path: Path) -> None:
    docs = tmp_path / "Documents"
    root = tmp_path / "Program Files"
    for year in (2020, 2021, 2024, 2027, 2028):
        _fake_mastercam_tree(root, docs, year)

    discovered = mastercam_install.discover_mastercam_addin_targets(docs, (root,))

    assert [year for year, _ in discovered] == [2021, 2024, 2027]
    assert all(path.name == "Add-Ins" for _, path in discovered)


def test_discovery_accepts_mastercam_executable_in_nested_product_folder(tmp_path: Path) -> None:
    docs = tmp_path / "Documents"
    root = tmp_path / "Program Files"
    installation = root / "Mastercam 2022" / "Mastercam"
    installation.mkdir(parents=True)
    (installation / "Mastercam.exe").write_bytes(b"mastercam")
    profile = docs / "My Mastercam 2022" / "Mastercam"
    profile.mkdir(parents=True)

    discovered = mastercam_install.discover_mastercam_addin_targets(docs, (root,))

    assert discovered == [(2022, profile / "Add-Ins")]


def test_fresh_mastercam_reinstall_can_create_missing_user_addin_folder(tmp_path: Path, monkeypatch) -> None:
    docs = tmp_path / "Documents"
    root = tmp_path / "Program Files"
    installed = root / "Mastercam 2026"
    installed.mkdir(parents=True)
    (installed / "Mastercam.exe").write_bytes(b"mastercam")
    expected = docs / "My Mastercam 2026" / "Mastercam" / "Add-Ins"

    assert mastercam_install.discover_mastercam_addin_targets(docs, (root,)) == [(2026, expected)]

    source_dir = tmp_path / "package" / "assets" / "mastercam_addin"
    source_dir.mkdir(parents=True)
    for name in mastercam_install.ADDIN_FILES:
        (source_dir / name).write_bytes(name.encode())
    app = tmp_path / "CNCProgramSheet.exe"
    app.write_bytes(b"exe")

    monkeypatch.setattr(mastercam_install, "resource_path", lambda *parts: tmp_path / "package" / Path(*parts))
    result = mastercam_install.install_mastercam_addin(
        documents_dir=docs, program_roots=(root,), application_exe=app,
        user_data=tmp_path / "user-data",
    )

    assert result.installed_years == (2026,)
    assert all((expected / name).is_file() for name in mastercam_install.ADDIN_FILES)


def test_install_copies_verified_files_and_backs_up_existing_addin(tmp_path: Path, monkeypatch) -> None:
    docs = tmp_path / "Documents"
    root = tmp_path / "Program Files"
    targets = {
        year: _fake_mastercam_tree(root, docs, year)
        for year in (2021, 2024, 2027)
    }
    source_dir = tmp_path / "package" / "assets" / "mastercam_addin"
    source_dir.mkdir(parents=True)
    source_files = {}
    for name in mastercam_install.ADDIN_FILES:
        path = source_dir / name
        path.write_bytes(f"packaged-{name}".encode())
        source_files[name] = path

    old_csx = targets[2024] / "GenerateProgramSheet.csx"
    old_csx.parent.mkdir(parents=True)
    old_csx.write_text("old user customization", encoding="utf-8")
    app_exe = tmp_path / "CNCProgramSheet.exe"
    app_exe.write_bytes(b"exe")
    data_dir = tmp_path / "LocalAppData" / "CNCProgramSheet"
    monkeypatch.setattr(
        mastercam_install,
        "resource_path",
        lambda *parts: tmp_path / "package" / Path(*parts),
    )

    result = mastercam_install.install_mastercam_addin(
        documents_dir=docs,
        program_roots=(root,),
        application_exe=app_exe,
        user_data=data_dir,
    )

    assert result.installed_years == (2021, 2024, 2027)
    for destination in targets.values():
        for name, source in source_files.items():
            assert (destination / name).read_bytes() == source.read_bytes()
    assert old_csx.with_name(old_csx.name + ".pre-cncprogramsheet.bak").read_text(
        encoding="utf-8"
    ) == "old user customization"
    assert (data_dir / "mastercam_bridge_path.txt").read_text(encoding="utf-8") == str(
        app_exe.resolve()
    )


def test_install_requires_at_least_one_detected_mastercam(tmp_path: Path) -> None:
    with pytest.raises(mastercam_install.MastercamInstallError, match="未检测到"):
        mastercam_install.install_mastercam_addin(
            documents_dir=tmp_path / "Documents",
            program_roots=(tmp_path / "Program Files",),
            application_exe=tmp_path / "app.exe",
            user_data=tmp_path / "settings",
        )


def test_settings_button_reinstalls_even_when_version_was_installed_before(monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    settings = {
        "install_mastercam_button": False,
        "mastercam_addin_installed_version": main_module.__version__,
    }
    installed: list[bool] = []
    dialogs: list[str] = []

    class Notice:
        def __init__(self, title: str, *_args: object, **_kwargs: object) -> None:
            dialogs.append(title)

        def exec(self) -> int:
            return 0

    def click_install_and_cancel(dialog: main_module.SettingsDialog) -> QDialog.DialogCode:
        dialog.install_mastercam_now.click()
        dialog.install_mastercam_now.click()
        return QDialog.DialogCode.Rejected

    monkeypatch.setattr(main_module, "load_settings", lambda: settings)
    monkeypatch.setattr(main_module, "save_settings", lambda _current: None)
    monkeypatch.setattr(main_module, "default_user_template", lambda: Path("assets/CNC程序单.xlsx").resolve())
    monkeypatch.setattr(main_module, "install_mastercam_addin", lambda: (
        installed.append(True), SimpleNamespace(installed_years=(2022,))
    )[1])
    monkeypatch.setattr(main_module, "ApplicationDialog", Notice)
    monkeypatch.setattr(main_module.SettingsDialog, "exec", click_install_and_cancel)

    window = main_module.MainWindow()
    try:
        window.edit_settings()
        assert installed == [True, True]
        assert dialogs == ["Mastercam 按钮安装完成", "Mastercam 按钮安装完成"]
        assert settings["mastercam_addin_installed_version"] == main_module.__version__
    finally:
        window.tray_icon.hide()
        window.close()
    assert app is not None
