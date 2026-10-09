import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QEvent, Qt
from PySide6.QtGui import QImage, QKeyEvent
from PySide6.QtWidgets import QApplication

from cnc_program_sheet import app_paths
import cnc_program_sheet.main as app_main
from cnc_program_sheet.main import ExcelPreview, MainWindow


def test_clipboard_image_is_cached_in_user_owned_image_directory(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(app_paths, "user_images_dir", lambda: tmp_path)

    cached = app_paths.cache_user_image_bytes(b"clipboard image bytes")

    assert cached.parent == tmp_path
    assert cached.name.endswith("_clipboard.png")
    assert cached.read_bytes() == b"clipboard image bytes"
    assert app_paths.cache_user_image_bytes(b"clipboard image bytes") == cached


def test_preview_ctrl_v_requests_paste_for_clipboard_screenshot() -> None:
    app = QApplication.instance() or QApplication([])
    preview = ExcelPreview()
    received: list[bool] = []
    preview.clipboard_image_paste_requested.connect(lambda: received.append(True))
    image = QImage(12, 8, QImage.Format.Format_ARGB32)
    image.fill(Qt.GlobalColor.red)
    QApplication.clipboard().setImage(image)

    event = QKeyEvent(QEvent.Type.KeyPress, Qt.Key.Key_V, Qt.KeyboardModifier.ControlModifier)
    preview.keyPressEvent(event)

    assert received == [True]
    assert event.isAccepted()
    QApplication.clipboard().clear()
    preview.close()
    assert app is not None


def test_main_window_pastes_clipboard_screenshot_at_selected_excel_cell(tmp_path, monkeypatch) -> None:
    app = QApplication.instance() or QApplication([])
    cached_image = tmp_path / "clipboard.png"

    def cache_image(data: bytes, filename: str = "clipboard.png"):
        cached_image.write_bytes(data)
        return cached_image

    monkeypatch.setattr(app_main, "cache_user_image_bytes", cache_image)
    window = MainWindow()
    window.load_template(app_main.resource_path("assets", "CNC程序单.xlsx"))
    # B4 is a covered cell in the A4:B4 merge.  Selecting it must resolve to
    # the merge root instead of opening the "select a paste location" error.
    window.preview.setCurrentCell(3, 1)
    monkeypatch.setattr(window, "_choose_image_size", lambda *_: (120, 60))
    image = QImage(24, 12, QImage.Format.Format_ARGB32)
    image.fill(Qt.GlobalColor.blue)
    QApplication.clipboard().setImage(image)

    window.paste_image_from_clipboard()

    assert len(window.image_placements) == 1
    placement = window.image_placements[0]
    assert (placement.anchor, placement.width, placement.height) == ("A4", 120, 60)
    assert placement.source_path == cached_image and cached_image.is_file()
    assert not window.preview.item(3, 0).icon().isNull()
    QApplication.clipboard().clear()
    window.close()
    assert app is not None
