import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QPushButton

import cnc_program_sheet.main as app_main
from cnc_program_sheet.main import (
    CONVERSION_TAG_BACKGROUND,
    ConversionResult,
    MainWindow,
    NcConversionDialog,
)
from cnc_program_sheet.models import ProgramRecord
from cnc_program_sheet.nc_transform import TransformMode


def _record(path: Path) -> ProgramRecord:
    return ProgramRecord(
        source_path=path,
        program_name=path.name,
        tool_number="T1",
        diameter="D1",
        machining_data="",
        depth="Z-1",
    )


def test_conversion_toolbar_replaces_the_two_image_buttons() -> None:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()

    labels = [button.text() for button in window.findChildren(QPushButton)]

    assert {"M06", "M304", "铜工"}.issubset(labels)
    assert "插入图片" not in labels
    assert "删除图片" not in labels
    window.close()
    assert app is not None


def test_m304_dialog_keeps_an_editable_source_and_applies_the_legacy_transform(tmp_path: Path) -> None:
    app = QApplication.instance() or QApplication([])
    source = tmp_path / "O-2.NC"
    source.write_text("G43H1Z20.M8\nM5\nM30\n", encoding="ascii")
    dialog = NcConversionDialog(TransformMode.M304)

    dialog.add_paths([source])

    assert dialog.file_list.item(0).text() == "M304    O-2.NC"
    assert dialog.file_list.item(0).background().color() == CONVERSION_TAG_BACKGROUND
    assert dialog.editor.isEnabled()
    dialog.editor.setPlainText("G43H2Z20.M8\nM5\nM30\n")
    dialog.apply_conversion()

    assert dialog.converted_files == [ConversionResult(source.resolve(), source.resolve())]
    assert source.read_text(encoding="ascii") == (
        "G43H2Z20.M8\nM304\nG05P10000\nM5\nG05P0\nM300\nM30\n"
    )
    dialog.close()
    assert app is not None


def test_m06_dialog_writes_windows_crlf_after_the_qt_editor_normalises_text(tmp_path: Path) -> None:
    """Guard against QPlainTextEdit turning an imported CRLF file into LF."""

    app = QApplication.instance() or QApplication([])
    source = tmp_path / "O-1.NC"
    source.write_bytes(b"N10T1M6\r\nN20G0X0\r\nN30M3\r\n")
    dialog = NcConversionDialog(TransformMode.M06)

    dialog.add_paths([source])
    assert dialog.editor.toPlainText() == "N10T1M6\nN20G0X0\nN30M3\n"
    dialog.apply_conversion()

    assert source.read_bytes() == b"N20G0X0\r\nN30M3\r\n"
    dialog.close()
    assert app is not None


def test_converted_copper_source_is_replaced_by_suffixless_result_and_tagged(tmp_path: Path, monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    source = tmp_path / "O0788.NC"
    target = tmp_path / "O0788"
    source.write_text("M30\n", encoding="ascii")
    target.write_text("M99\n", encoding="ascii")
    monkeypatch.setattr(app_main, "parse_file_records", lambda path: [_record(Path(path))])
    window = MainWindow()

    window._import_converted_files([ConversionResult(source, target)], TransformMode.COPPER)

    assert window.records[0] is not None
    assert window.records[0].source_path == target
    item = window.file_list.item(0)
    assert item.text() == "TG    O0788"
    assert item.data(Qt.ItemDataRole.UserRole) == str(target.resolve())
    assert item.background().color() == CONVERSION_TAG_BACKGROUND
    window.close()
    assert app is not None


def test_conversion_replaces_existing_source_rows_without_shifting_later_programs(tmp_path: Path, monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    first = tmp_path / "O-1.NC"
    later = tmp_path / "O-2.NC"
    first.write_text("M30\n", encoding="ascii")
    later.write_text("M30\n", encoding="ascii")

    def parse(path: Path) -> list[ProgramRecord]:
        if Path(path).name == "O-1.NC":
            return [_record(Path(path))]
        return [_record(Path(path))]

    monkeypatch.setattr(app_main, "parse_file_records", parse)
    window = MainWindow()
    window.add_nc_files([first, later])
    original_later_row = next(
        index for index, record in enumerate(window.records) if record is not None and record.source_path == later
    )

    window._import_converted_files([ConversionResult(first, first)], TransformMode.M06)

    assert window.records[original_later_row] is not None
    assert window.records[original_later_row].source_path == later
    assert window.file_list.item(0).text() == "M06    O-1.NC"
    window.close()
    assert app is not None


def test_conversion_keeps_plain_program_names_in_natural_order_when_layout_is_unedited(tmp_path: Path, monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    first = tmp_path / "O-1.NC"
    later = tmp_path / "O-10.NC"
    first.write_text("M30\n", encoding="ascii")
    later.write_text("M30\n", encoding="ascii")
    monkeypatch.setattr(app_main, "parse_file_records", lambda path: [_record(Path(path))])
    window = MainWindow()
    window.add_nc_files([later])

    window._import_converted_files([ConversionResult(first, first)], TransformMode.M304)

    assert [record.program_name for record in window.records if record is not None] == ["O-1.NC", "O-10.NC"]
    window.close()
    assert app is not None


def test_copper_conversion_does_not_duplicate_an_already_imported_suffixless_target(tmp_path: Path, monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    source = tmp_path / "O0788.NC"
    target = tmp_path / "O0788"
    source.write_text("M30\n", encoding="ascii")
    target.write_text("M99\n", encoding="ascii")
    monkeypatch.setattr(app_main, "parse_file_records", lambda path: [_record(Path(path))])
    window = MainWindow()
    window.add_nc_files([source, target])

    window._import_converted_files([ConversionResult(source, target)], TransformMode.COPPER)

    records = [record for record in window.records if record is not None]
    assert len(records) == 1
    assert records[0].source_path == target
    assert window.file_list.count() == 1
    window.close()
    assert app is not None


def test_m06_preserves_multi_tool_program_sheet_sections_after_exact_m6_removal(tmp_path: Path) -> None:
    app = QApplication.instance() or QApplication([])
    source = tmp_path / "0505.NC"
    original_text = (
        "(T1|D1|H1)\n"
        "T1M6\n"
        "G0Z5\n"
        "G1Z-1\n"
        "(T2|D2|H2)\n"
        "T2M6\n"
        "G0Z5\n"
        "G1Z-2\n"
    )
    # This is the exact M06 disk result: the controller tool-change lines
    # are gone, but the program sheet must still retain both tool sections.
    source.write_text("(T1|D1|H1)\nG0Z5\nG1Z-1\n(T2|D2|H2)\nG0Z5\nG1Z-2\n", encoding="ascii")
    window = MainWindow()

    window._import_converted_files(
        [ConversionResult(source, source, parser_snapshot=original_text)],
        TransformMode.M06,
    )

    assert len([record for record in window.records if record is not None]) == 2
    assert [record.depth for record in window.records if record is not None] == ["Z-1", "Z-2"]
    window.reparse()
    assert len([record for record in window.records if record is not None]) == 2
    window.close()
    assert app is not None


def test_plain_text_paste_asks_for_a_real_nc_source_location(tmp_path: Path, monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    destination = tmp_path / "pasted.NC"
    monkeypatch.setattr(app_main.QFileDialog, "getSaveFileName", lambda *args, **kwargs: (str(destination), ""))
    dialog = NcConversionDialog(TransformMode.M06)

    dialog._save_pasted_source("O0001\nT1M6\nG1Z-1\n")

    assert destination.is_file()
    assert dialog.file_list.item(0).text() == "M06    pasted.NC"
    assert dialog._sources[0].source_path == destination.resolve()
    dialog.close()
    assert app is not None


def test_partial_conversion_success_is_not_lost_when_a_later_file_fails(tmp_path: Path, monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    first = tmp_path / "O-1.NC"
    second = tmp_path / "O-2.NC"
    first.write_text("T1M6\n", encoding="ascii")
    second.write_text("T1M6\n", encoding="ascii")
    dialog = NcConversionDialog(TransformMode.M06)
    dialog.add_paths([first, second])

    def transform(path: Path, mode: TransformMode, source_text: str | None = None) -> Path:
        if Path(path) == second.resolve():
            raise OSError("测试写入失败")
        return Path(path)

    monkeypatch.setattr(app_main, "transform_file", transform)
    monkeypatch.setattr(app_main.QMessageBox, "warning", lambda *args, **kwargs: None)
    dialog.apply_conversion()

    assert dialog.converted_files == [
        ConversionResult(first.resolve(), first.resolve(), parser_snapshot="T1M6\n")
    ]
    dialog.close()
    assert app is not None
