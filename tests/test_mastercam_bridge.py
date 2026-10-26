from pathlib import Path

import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

import cnc_program_sheet.main as app_main
from cnc_program_sheet.main import (
    MainWindow,
    parse_conversion_launch_mode,
    parse_launch_arguments,
    parse_mastercam_launch,
    parse_apply_conversion,
)
from cnc_program_sheet.nc_transform import TransformMode


def test_mastercam_bridge_accepts_quoted_paths_with_spaces() -> None:
    paths = parse_launch_arguments(
        ["--import-nc", r"C:\Mastercam Files\Part 1.NC", r"D:\Posted NC\Part 2.NC"]
    )

    assert paths == [
        Path(r"C:\Mastercam Files\Part 1.NC"),
        Path(r"D:\Posted NC\Part 2.NC"),
    ]


def test_mastercam_bridge_accepts_repeated_import_options() -> None:
    paths = parse_launch_arguments(
        ["--import-nc", r"C:\NC\A.NC", "--import-nc", r"C:\NC\B.NC"]
    )

    assert paths == [Path(r"C:\NC\A.NC"), Path(r"C:\NC\B.NC")]


def test_regular_launch_has_no_import_paths() -> None:
    assert parse_launch_arguments([]) == []


def test_mastercam_bridge_selects_the_requested_conversion_mode() -> None:
    assert parse_conversion_launch_mode(["--conversion-mode", "M06"]) is TransformMode.M06
    assert parse_conversion_launch_mode(["--conversion-mode", "M304"]) is TransformMode.M304
    assert parse_conversion_launch_mode(["--conversion-mode", "COPPER"]) is TransformMode.COPPER


def test_regular_nc_import_does_not_open_a_conversion_mode() -> None:
    assert parse_conversion_launch_mode(["--import-nc", r"C:\NC\A.NC"]) is None


def test_mastercam_addin_launch_flag_is_detected() -> None:
    assert parse_mastercam_launch(["--from-mastercam", "--import-nc", r"C:\NC\A.NC"])
    assert not parse_mastercam_launch(["--import-nc", r"C:\NC\A.NC"])


def test_mastercam_can_request_immediate_conversion_but_regular_ui_stays_manual() -> None:
    assert parse_apply_conversion(
        ["--from-mastercam", "--conversion-mode", "M06", "--apply-conversion"]
    )
    assert not parse_apply_conversion(["--conversion-mode", "M06", "--import-nc", r"C:\NC\A.NC"])


def test_mastercam_chooser_no_longer_contains_g1_post_api() -> None:
    source = Path("mastercam_addin/GenerateProgramSheet.csx").read_text(encoding="utf-8")

    assert "PostSelectedOperations" not in source
    assert 'mode != "NORMAL"' in source
    assert "--apply-conversion" in source
    assert "PostSelectedOperationsAndImport" not in source


def test_mastercam_bridge_finds_year_in_parent_of_nested_executable_folder() -> None:
    source = Path("mastercam_addin/GenerateProgramSheet.csx").read_text(encoding="utf-8")
    packaged_source = Path("assets/mastercam_addin/GenerateProgramSheet.csx").read_text(encoding="utf-8")

    assert "while (installDirectory != null)" in source
    assert "installDirectory = installDirectory.Parent;" in source
    assert '"My Mastercam " + year' in source
    assert packaged_source == source


def test_conversion_launch_preloads_nc_into_the_matching_conversion_window(tmp_path, monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    source = tmp_path / "part.NC"
    imported: list[tuple[TransformMode, list[Path]]] = []

    class FakeConversionDialog:
        def __init__(self, mode, _parent):
            self.mode = mode
            self.converted_files = []

        def add_paths(self, paths):
            imported.append((self.mode, list(paths)))

        def exec(self):
            return app_main.QDialog.DialogCode.Accepted

    monkeypatch.setattr(app_main, "NcConversionDialog", FakeConversionDialog)
    window = MainWindow()
    window.open_conversion_dialog(TransformMode.M06, [source])

    assert imported == [(TransformMode.M06, [source])]
    window.close()
    assert app is not None


def test_mastercam_immediate_conversion_uses_existing_transform_and_import_flow(tmp_path, monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    source = tmp_path / "part.NC"
    source.write_bytes(b"%\r\nT1\r\nM6\r\nM30\r\n")
    calls: list[tuple[TransformMode, list[app_main.ConversionResult]]] = []
    window = MainWindow()
    monkeypatch.setattr(
        window,
        "_import_converted_files",
        lambda conversions, mode: calls.append((mode, conversions)),
    )

    window.convert_and_import_mastercam_files(TransformMode.M06, [source])

    assert source.read_bytes() == b"%\r\nT1\r\nM30\r\n"
    assert calls and calls[0][0] is TransformMode.M06
    assert calls[0][1][0].source_path == source
    assert calls[0][1][0].target_path == source.resolve()
    assert "M6" in (calls[0][1][0].parser_snapshot or "")
    window.tray_icon.hide()
    window.close()
    assert app is not None
