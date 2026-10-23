import os
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import load_workbook
from PySide6.QtWidgets import QApplication

from cnc_program_sheet.excel_service import LAYOUT, export_workbook
from cnc_program_sheet.main import MainWindow
from cnc_program_sheet.nc_parser import parse_records_text


def test_conflicting_same_tool_speeds_highlight_the_preview_and_export_row(tmp_path: Path) -> None:
    app = QApplication.instance() or QApplication([])
    template = Path("assets/CNC程序单.xlsx")
    source = tmp_path / "O-1.NC"
    source.write_text("test fixture", encoding="ascii")
    records = parse_records_text(
        """O0001
(T1|D4|H1)
T1M6
S3000M3
G1Z-1
S3500M3
M30
""",
        source.name,
        source,
    )
    window = MainWindow()
    window.load_template(template)
    window.records = records
    window._apply_records_to_view()

    preview_row = LAYOUT.first_program_row - 1
    for column in (0, 1, 10):
        item = window.preview.item(preview_row, column)
        assert item is not None and item.background().color().name() == "#ffc7ce"
    status_item = window.result_table.item(0, 7)
    assert status_item is not None and "S3000 / S3500" in status_item.text()

    target = tmp_path / "warning.xlsx"
    warning_cells = [
        f"{window.preview._column_name(column)}{LAYOUT.first_program_row}"
        for column in range(1, window.preview.columnCount() + 1)
    ]
    export_workbook(template, target, window.manual_cells, red_fill_cells=warning_cells)
    workbook = load_workbook(target, data_only=False)
    try:
        sheet = workbook[LAYOUT.sheet_name]
        for coordinate in ("A7", "B7", "K7"):
            assert sheet[coordinate].fill.fill_type == "solid"
            assert sheet[coordinate].fill.fgColor.rgb == "FFFFC7CE"
        assert sheet["A7"].value == "1"
    finally:
        workbook.close()
    window.close()
    assert app is not None


def test_consistent_speeds_do_not_leave_a_stale_red_fill() -> None:
    app = QApplication.instance() or QApplication([])
    template = Path("assets/CNC程序单.xlsx")
    source = Path("O-1.NC")
    window = MainWindow()
    window.load_template(template)
    window.records = parse_records_text(
        """O0001
(T1|D4|H1)
T1M6
S3000M3
G1Z-1
S3000M3
M30
""",
        source.name,
        source,
    )
    window._apply_records_to_view()

    item = window.preview.item(LAYOUT.first_program_row - 1, 1)
    assert item is not None and item.background().color().name() != "#ffc7ce"
    assert window.records[0].spindle_speed_warning == ""
    window.close()
    assert app is not None
