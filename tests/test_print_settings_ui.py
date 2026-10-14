import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtGui import QPageLayout
from PySide6.QtPrintSupport import QPrinter
from PySide6.QtWidgets import QApplication

from cnc_program_sheet.main import PrinterSelectionDialog, ProgramSheetPrintPreviewDialog, SettingsDialog


def test_print_orientation_and_toolbar_visibility_preferences_are_available() -> None:
    app = QApplication.instance() or QApplication([])
    printer_dialog = PrinterSelectionDialog(QPageLayout.Orientation.Landscape)
    assert printer_dialog.orientation.currentData() == QPageLayout.Orientation.Landscape

    settings = SettingsDialog(
        "hanleizhen/lucky",
        True,
        True,
        True,
        False,
        False,
        {"print": False},
    )
    assert not settings.toolbar_buttons["print"].isChecked()

    preview = ProgramSheetPrintPreviewDialog(QPrinter(), lambda printer: None)
    preview.orientation.setCurrentIndex(0)
    assert preview._printer.pageLayout().orientation() == QPageLayout.Orientation.Portrait
    preview.close()
    settings.close()
    printer_dialog.close()
    assert app is not None
