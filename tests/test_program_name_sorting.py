import os
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

import cnc_program_sheet.main as app_main
from cnc_program_sheet.main import MainWindow, natural_program_sort_key
from cnc_program_sheet.models import ProgramRecord


def test_natural_program_name_sort_orders_numeric_parts_as_numbers() -> None:
    names = ["O-10.NC", "O-2.NC", "O-1.NC", "DB-6C.NC", "DB-6A.NC", "DB-6B.NC"]

    assert sorted(names, key=natural_program_sort_key) == [
        "DB-6A.NC",
        "DB-6B.NC",
        "DB-6C.NC",
        "O-1.NC",
        "O-2.NC",
        "O-10.NC",
    ]


def test_imported_nc_files_fill_template_in_natural_name_order(tmp_path: Path, monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    paths = [tmp_path / name for name in ("O-10.NC", "O-2.NC", "O-1.NC")]
    for path in paths:
        path.write_text("M30\n", encoding="utf-8")

    def parse(path: Path) -> list[ProgramRecord]:
        return [
            ProgramRecord(
                source_path=path,
                program_name=path.name,
                tool_number="T1",
                diameter="D1",
                machining_data="",
                depth="Z-1",
            )
        ]

    monkeypatch.setattr(app_main, "parse_file_records", parse)
    window = MainWindow()
    window.load_template(Path("assets/CNC程序单.xlsx"))
    window.add_nc_files([paths[0], paths[1], paths[2]])

    assert [record.program_name for record in window.records if record is not None] == [
        "O-1.NC",
        "O-2.NC",
        "O-10.NC",
    ]
    assert [window.preview.item(row, 1).text() for row in range(6, 9)] == [
        "O-1.NC",
        "O-2.NC",
        "O-10.NC",
    ]
    assert window.preview.objectName() == "excelPreview"
    assert "QTableWidget#excelPreview { gridline-color: #000000" in window.styleSheet()
    window.close()
    assert app is not None
