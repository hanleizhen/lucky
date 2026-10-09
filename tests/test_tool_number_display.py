import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from cnc_program_sheet.main import MainWindow


def test_nc_filename_hides_tool_number_but_suffixless_program_keeps_it(tmp_path) -> None:
    app = QApplication.instance() or QApplication([])
    with_suffix = tmp_path / "XB-6.NC"
    without_suffix = tmp_path / "XB-6"
    source = "O0006\n(T1|D4R0.2H1)\nT1M6\nG0Z5\nG1Z-1\nM30\n"
    with_suffix.write_text(source, encoding="utf-8")
    without_suffix.write_text(source, encoding="utf-8")
    window = MainWindow()

    window.add_nc_files([with_suffix, without_suffix])

    by_name = {record.program_name: record for record in window.records if record is not None}
    assert by_name["XB-6.NC"].tool_number == ""
    assert by_name["XB-6"].tool_number == "T1"
    window.close()
    assert app is not None
