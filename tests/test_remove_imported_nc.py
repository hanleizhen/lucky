import os
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QAbstractItemView

import cnc_program_sheet.main as app_main
from cnc_program_sheet.main import MainWindow
from cnc_program_sheet.models import ProgramRecord


class _ConfirmRemoval:
    def __init__(self, *args, **kwargs) -> None:
        self.choice = "remove"

    def exec(self) -> int:
        return 0


def test_remove_selected_imports_keeps_original_files_and_supports_multi_select(
    tmp_path: Path, monkeypatch
) -> None:
    app = QApplication.instance() or QApplication([])
    template = Path("assets/CNC程序单.xlsx")
    monkeypatch.setattr(app_main, "default_user_template", lambda: template)
    monkeypatch.setattr(app_main, "ApplicationDialog", _ConfirmRemoval)

    sources = [tmp_path / "O-1.NC", tmp_path / "O-2.NC"]
    for source in sources:
        source.write_text("O0001\nM30\n", encoding="ascii")

    window = MainWindow()
    window.load_template(template)
    window.records = [
        ProgramRecord(
            source_path=source,
            program_name=source.name,
            tool_number="T1",
            diameter="D1",
            machining_data="",
            depth="Z-1",
        )
        for source in sources
    ]
    window._apply_records_to_view()

    assert window.file_list.selectionMode() == QAbstractItemView.SelectionMode.ExtendedSelection
    window.file_list.item(0).setSelected(True)
    window.file_list.item(1).setSelected(True)
    window.remove_selected_nc_programs()

    assert window.records == []
    assert window.file_list.count() == 0
    assert window.preview.item(6, 1).text() == ""
    assert all(source.is_file() for source in sources)
    window.close()
    assert app is not None


def test_delete_key_removes_the_selected_imported_nc(tmp_path: Path, monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    template = Path("assets/CNC程序单.xlsx")
    monkeypatch.setattr(app_main, "default_user_template", lambda: template)
    monkeypatch.setattr(app_main, "ApplicationDialog", _ConfirmRemoval)
    source = tmp_path / "O-1.NC"
    source.write_text("O0001\nM30\n", encoding="ascii")

    window = MainWindow()
    window.load_template(template)
    window.records = [
        ProgramRecord(
            source_path=source,
            program_name=source.name,
            tool_number="T1",
            diameter="D1",
            machining_data="",
            depth="Z-1",
        )
    ]
    window._apply_records_to_view()
    window.show()
    window.file_list.setCurrentRow(0)
    app.processEvents()

    QTest.keyClick(window.file_list.viewport(), Qt.Key.Key_Delete)

    assert window.records == []
    assert window.file_list.count() == 0
    assert source.is_file()
    window.close()
    assert app is not None
