import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QPushButton

from cnc_program_sheet.main import (
    APPLICATION_STYLE_SHEET,
    DEFAULT_DARK_WINDOW_COLOR,
    MainWindow,
    SettingsDialog,
    application_style_sheet,
    apply_application_theme,
)


def test_shared_theme_covers_main_window_dialogs_menus_and_sheet_controls() -> None:
    """Visual refactors must keep one shared design system for every window."""

    app = QApplication.instance() or QApplication([])
    original = app.styleSheet()
    try:
        apply_application_theme(app)
        assert app.styleSheet() == APPLICATION_STYLE_SHEET
        for selector in (
            "QMainWindow, QDialog, QMessageBox",
            "QPushButton#primaryButton",
            "QMenu::item:selected",
            "QTabBar::tab:selected",
            "QTableWidget#excelPreview",
            "#dropArea",
        ):
            assert selector in APPLICATION_STYLE_SHEET
    finally:
        app.setStyleSheet(original)


def test_visual_refresh_keeps_all_primary_toolbar_entrypoints() -> None:
    """A layout-only change must never silently remove a working function."""

    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    try:
        labels = {button.text() for button in window.findChildren(QPushButton)}
        assert {
            "完成并保存",
            "保存位置",
            "重新解析",
            "撤销",
            "重做",
            "查找替换",
            "清空",
            "打开模板",
            "打印",
            "M06",
            "M304",
            "铜工",
            "选择 NC 文件",
            "检查更新",
            "设置",
        }.issubset(labels)
    finally:
        window.tray_icon.hide()
        window.close()
    assert app is not None


def test_dark_mode_and_custom_window_colour_keep_the_excel_sheet_light() -> None:
    """The dark application shell must never make the printable sheet dark."""

    app = QApplication.instance() or QApplication([])
    original = app.styleSheet()
    try:
        apply_application_theme(app, "dark", "#263A55")
        assert app.styleSheet() == application_style_sheet("dark", "#263A55")
        assert "background: #263a55;" in app.styleSheet()
        assert "QTableWidget#excelPreview" in app.styleSheet()
        assert "background: #FFFFFF;" in app.styleSheet()

        settings = SettingsDialog(
            "hanleizhen/lucky",
            True,
            True,
            True,
            False,
            False,
            {},
            "portrait",
            8,
            "#B7D7FF",
            "dark",
            "#263A55",
        )
        assert settings.selected_appearance_mode() == "dark"
        assert settings.selected_window_color() == "#263a55"
        settings._restore_default_window_color()
        assert settings.selected_window_color() == DEFAULT_DARK_WINDOW_COLOR.lower()
        settings.close()
    finally:
        app.setStyleSheet(original)
