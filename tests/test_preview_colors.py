import os
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QTableWidgetSelectionRange

from cnc_program_sheet.main import ExcelPreview


def test_preview_uses_template_fill_or_white_not_black() -> None:
    app = QApplication.instance() or QApplication([])
    preview = ExcelPreview()
    preview.load_template(Path("assets/CNC程序单.xlsx"))

    no_fill = preview.item(0, 0)
    blue_input = preview.item(6, 1)
    assert no_fill is not None and no_fill.background().color().name() == "#ffffff"
    assert blue_input is not None and blue_input.background().color().name() != "#000000"
    preview.close()
    assert app is not None


def test_preview_moves_program_values_without_changing_template_cell_style() -> None:
    app = QApplication.instance() or QApplication([])
    preview = ExcelPreview()
    preview.load_template(Path("assets/CNC程序单.xlsx"))
    original_color = preview.item(7, 1).background().color().name()
    preview.set_value("B7", "O-1.NC")
    preview.set_value("B8", "O-2.NC")

    # Table row 7 is Excel row 8: add a blank row below Excel row 7.
    preview.shift_program_row_values(7, 1)

    assert preview.item(6, 1).text() == "O-1.NC"
    assert preview.item(7, 1).text() == ""
    assert preview.item(8, 1).text() == "O-2.NC"
    assert preview.item(7, 1).background().color().name() == original_color
    preview.close()
    assert app is not None


def test_preview_pastes_tabular_text_at_the_selected_cell() -> None:
    app = QApplication.instance() or QApplication([])
    preview = ExcelPreview()
    preview.load_template(Path("assets/CNC程序单.xlsx"))
    preview.setCurrentCell(6, 1)  # B7

    assert preview.paste_text_grid("A\tB\nC\tD") == 4
    assert preview.item(6, 1).text() == "A"
    assert preview.item(6, 2).text() == "B"
    assert preview.item(7, 1).text() == "C"
    assert preview.item(7, 2).text() == "D"
    preview.close()
    assert app is not None


def test_preview_single_value_fills_the_selected_rectangle_like_excel() -> None:
    app = QApplication.instance() or QApplication([])
    preview = ExcelPreview()
    preview.load_template(Path("assets/CNC程序单.xlsx"))
    preview.setCurrentCell(6, 1)
    preview.setRangeSelected(QTableWidgetSelectionRange(6, 1, 7, 2), True)  # B7:C8

    assert preview.paste_text_grid("客户备注") == 4
    assert [preview.item(row, column).text() for row in (6, 7) for column in (1, 2)] == [
        "客户备注",
        "客户备注",
        "客户备注",
        "客户备注",
    ]
    assert len(preview.selectedIndexes()) == 4
    preview.close()
    assert app is not None
