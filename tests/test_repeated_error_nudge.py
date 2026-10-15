import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QFrame, QPushButton

from cnc_program_sheet.error_dialog import ApplicationDialog, ErrorDialog
from cnc_program_sheet.main import MainWindow, REPEATED_ERROR_WINDOW_SECONDS


def test_same_error_only_counts_as_consecutive_inside_the_short_window() -> None:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    try:
        assert window._record_error_occurrence("无法保存", now=100.0) == 1
        assert window._record_error_occurrence("无法保存", now=101.0) == 2
        assert window._record_error_occurrence("无法导入", now=102.0) == 1
        assert window._record_error_occurrence("无法导入", now=102.0 + REPEATED_ERROR_WINDOW_SECONDS + 0.1) == 1
    finally:
        window.tray_icon.hide()
        window.close()
    assert app is not None


def test_repeated_error_dialog_shows_the_fun_nudge_only_after_the_second_error(tmp_path) -> None:
    app = QApplication.instance() or QApplication([])
    first = ErrorDialog("无法保存", "请先导入 NC 文件。", tmp_path / "error.log", repeated_count=1, show_repeat_nudge=True)
    repeated = ErrorDialog("无法保存", "请先导入 NC 文件。", tmp_path / "error.log", repeated_count=2, show_repeat_nudge=True)
    try:
        assert first.findChild(QFrame, "repeatErrorNudge") is None
        assert repeated.findChild(QFrame, "repeatErrorNudge") is not None
    finally:
        first.close()
        repeated.close()
    assert app is not None


def test_application_dialog_keeps_details_collapsed_until_requested() -> None:
    app = QApplication.instance() or QApplication([])
    dialog = ApplicationDialog("测试", "简短提示", details="技术详情")
    try:
        assert dialog.details_view.isHidden()
        dialog._toggle_details()
        assert not dialog.details_view.isHidden()
    finally:
        dialog.close()
    assert app is not None


def test_error_dialog_can_offer_a_direct_retry_action(tmp_path) -> None:
    app = QApplication.instance() or QApplication([])
    dialog = ErrorDialog(
        "检查更新失败",
        "暂时无法检查更新。",
        tmp_path / "error.log",
        buttons=(("重试", "retry"), ("关闭", "close")),
    )
    try:
        retry = next(button for button in dialog.findChildren(QPushButton) if button.text() == "重试")
        retry.click()
        assert dialog.choice == "retry"
    finally:
        dialog.close()
    assert app is not None
