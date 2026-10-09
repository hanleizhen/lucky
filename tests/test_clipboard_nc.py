import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from pathlib import Path

from PySide6.QtCore import QEvent, QMimeData, Qt, QUrl
from PySide6.QtGui import QKeyEvent
from PySide6.QtWidgets import QApplication

from cnc_program_sheet import app_paths
import cnc_program_sheet.main as app_main
from cnc_program_sheet.main import ExcelPreview, MainWindow, clipboard_nc_filename, is_nc_program_text


def test_clipboard_nc_text_is_cached_in_user_owned_directory(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(app_paths, "pasted_nc_programs_dir", lambda: tmp_path)

    cached = app_paths.cache_pasted_nc_program("O0788\nT1M6\nG0Z5\nG1Z-1\n", "O0788.NC")

    assert cached.parent.parent == tmp_path
    assert cached.name == "O0788.NC"
    assert cached.read_text(encoding="utf-8") == "O0788\nT1M6\nG0Z5\nG1Z-1\n"


def test_nc_clipboard_text_accepts_compact_controller_programs() -> None:
    compact = "%\nO-12\n(T1|D4R0.2H1)\nT1M6\nG0G90G54X0Y0Z5\nG1Z-1F100\nM30\n"

    assert is_nc_program_text(compact)
    assert clipboard_nc_filename(compact) == "O-12.NC"
    assert not is_nc_program_text("这是一段普通说明文字")


def test_preview_ctrl_v_requests_nc_file_import(tmp_path) -> None:
    app = QApplication.instance() or QApplication([])
    source = tmp_path / "O-1.NC"
    source.write_text("O-1\nT1M6\nG0Z5\n", encoding="utf-8")
    mime_data = QMimeData()
    mime_data.setUrls([QUrl.fromLocalFile(str(source))])
    QApplication.clipboard().setMimeData(mime_data)
    preview = ExcelPreview()
    received: list[bool] = []
    preview.clipboard_nc_paste_requested.connect(lambda: received.append(True))

    event = QKeyEvent(QEvent.Type.KeyPress, Qt.Key.Key_V, Qt.KeyboardModifier.ControlModifier)
    preview.keyPressEvent(event)

    assert received == [True]
    assert event.isAccepted()
    QApplication.clipboard().clear()
    preview.close()
    assert app is not None


def test_main_window_imports_pasted_nc_text(tmp_path, monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    source = tmp_path / "O0788.NC"
    source.write_text("O0788\n(T1|D4R0.2H1)\nT1M6\nG0Z5\nG1Z-1\nM30\n", encoding="utf-8")
    monkeypatch.setattr(app_main, "cache_pasted_nc_program", lambda text, filename: source)
    window = MainWindow()
    QApplication.clipboard().setText(source.read_text(encoding="utf-8"))

    assert window.paste_nc_from_clipboard()
    assert len(window.records) == 1
    assert window.records[0] is not None
    assert window.records[0].program_name == "O0788.NC"

    QApplication.clipboard().clear()
    window.close()
    assert app is not None
