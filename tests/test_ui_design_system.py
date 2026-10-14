import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QPushButton

from cnc_program_sheet.main import APPLICATION_STYLE_SHEET, MainWindow, apply_application_theme


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
