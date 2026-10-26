import os
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtGui import QPageLayout
from PySide6.QtPrintSupport import QPrinter
from PySide6.QtWidgets import QApplication

import cnc_program_sheet.main as main_module
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
        "portrait",
        8,
        "#C6EFCE",
    )
    assert not settings.toolbar_buttons["print"].isChecked()
    assert not settings.install_mastercam_button.isChecked()
    assert not settings.exit_on_window_close.isChecked()
    assert not settings.mastercam_nc_folder.text()
    assert settings.mastercam_sequence_search_enabled.isChecked()
    assert settings.mastercam_sequence_search_count.value() == 11
    assert not settings.show_mastercam_install_success.isChecked()
    assert settings.print_quality.currentData() == "high"
    settings.install_mastercam_button.setChecked(True)
    assert settings.install_mastercam_button.isChecked()
    settings.exit_on_window_close.setChecked(True)
    assert settings.exit_on_window_close.isChecked()
    settings.mastercam_sequence_search_count.setValue(12)
    settings.print_quality.setCurrentIndex(settings.print_quality.findData("maximum"))
    assert settings.print_quality.currentData() == "maximum"
    settings.mastercam_nc_folder.setText(r"C:\Users\Administrator\Documents\My Mastercam 2022\Mastercam\Mill\NC")
    assert settings.mastercam_nc_folder.text().endswith(r"Mastercam\Mill\NC")
    assert settings.selected_selection_color() == "#c6efce"

    preview = ProgramSheetPrintPreviewDialog(QPrinter(), lambda printer: None, "A1:K25", 8)
    preview.orientation.setCurrentIndex(0)
    assert preview._printer.pageLayout().orientation() == QPageLayout.Orientation.Portrait
    selection_requested: list[bool] = []
    preview.select_print_area_requested.connect(lambda: selection_requested.append(True))
    preview._select_print_area()
    assert selection_requested == [True]
    preview.close()
    settings.close()
    printer_dialog.close()
    assert app is not None


def test_print_area_and_layout_preferences_are_persisted_per_template(monkeypatch) -> None:
    """A chosen print range and portrait/margin settings survive future prints."""

    app = QApplication.instance() or QApplication([])
    stored: list[dict] = []
    monkeypatch.setattr(main_module, "load_settings", lambda: {})
    monkeypatch.setattr(main_module, "save_settings", lambda settings: stored.append(dict(settings)))
    monkeypatch.setattr(
        main_module,
        "default_user_template",
        lambda: Path("assets/CNC程序单.xlsx").resolve(),
    )
    window = main_module.MainWindow()
    try:
        window.begin_print_area_selection()
        window._finish_print_area_selection(6, 0, 10, 4)
        assert window._saved_print_area() == (6, 0, 10, 4)
        assert window._print_area_label(window._saved_print_area()) == "A7:E11"

        window._save_print_preferences(QPageLayout.Orientation.Portrait, 12)
        assert window.settings["print_orientation"] == "portrait"
        assert window._configured_print_orientation() == QPageLayout.Orientation.Portrait
        assert window._configured_print_margin_mm() == 12
        window.settings["print_quality"] = "standard"
        assert window._configured_print_render_scale() == 3
        window.settings["print_quality"] = "maximum"
        assert window._configured_print_render_scale() == 7
        assert stored

        window.clear_print_area()
        assert window._saved_print_area() == window._full_print_area()
    finally:
        window.tray_icon.hide()
        window.close()
    assert app is not None
