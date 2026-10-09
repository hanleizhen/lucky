import os
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PIL import Image
from PySide6.QtCore import QPoint, Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication

from cnc_program_sheet.main import MainWindow
from cnc_program_sheet.models import ImagePlacement


def test_preview_renders_a_large_floating_image_and_commits_dragged_geometry(tmp_path: Path) -> None:
    app = QApplication.instance() or QApplication([])
    source = tmp_path / "picture.png"
    Image.new("RGB", (180, 90), "#2f75b5").save(source)
    window = MainWindow()
    window.load_template(Path("assets/CNC程序单.xlsx"))
    placement = ImagePlacement(source_path=source, anchor="D7", width=180, height=90)
    window.image_placements = [placement]
    window.preview.refresh_floating_images(window.image_placements)

    overlay = window.preview._image_overlays[placement.placement_id]
    assert (overlay.width(), overlay.height()) == (180, 90)
    assert window.preview.item(6, 3).icon().isNull()

    window.preview._commit_overlay_geometry(placement.placement_id, 240, 190, 250, 110)

    moved = window.image_placements[0]
    assert (moved.width, moved.height) == (250, 110)
    assert moved.offset_x >= 0 and moved.offset_y >= 0
    assert moved.anchor
    window.close()
    assert app is not None


def test_removing_selected_floating_image_removes_only_that_picture(tmp_path: Path) -> None:
    app = QApplication.instance() or QApplication([])
    source = tmp_path / "picture.png"
    Image.new("RGB", (30, 20), "#2f75b5").save(source)
    window = MainWindow()
    window.load_template(Path("assets/CNC程序单.xlsx"))
    first = ImagePlacement(source_path=source, anchor="D7", width=120, height=80)
    second = ImagePlacement(source_path=source, anchor="D7", width=140, height=90)
    window.image_placements = [first, second]
    window.preview.refresh_floating_images(window.image_placements)
    window.preview.select_image(second.placement_id)

    window.remove_images_at_selected_cell()

    assert [placement.placement_id for placement in window.image_placements] == [first.placement_id]
    assert second.placement_id not in window.preview._image_overlays
    window.close()
    assert app is not None


def test_user_drag_and_corner_resize_update_the_excel_image_placement(tmp_path: Path) -> None:
    app = QApplication.instance() or QApplication([])
    source = tmp_path / "picture.png"
    Image.new("RGB", (160, 80), "#2f75b5").save(source)
    window = MainWindow()
    window.show()
    app.processEvents()
    placement = ImagePlacement(source_path=source, anchor="D7", width=160, height=80)
    window.image_placements = [placement]
    window.preview.refresh_floating_images(window.image_placements)
    app.processEvents()
    overlay = window.preview._image_overlays[placement.placement_id]

    QTest.mousePress(overlay, Qt.MouseButton.LeftButton, pos=QPoint(80, 40))
    QTest.mouseMove(overlay, QPoint(112, 58))
    QTest.mouseRelease(overlay, Qt.MouseButton.LeftButton, pos=QPoint(112, 58))
    app.processEvents()
    moved = window.image_placements[0]
    assert (moved.anchor, moved.offset_x, moved.offset_y) != (
        placement.anchor,
        placement.offset_x,
        placement.offset_y,
    )

    overlay = window.preview._image_overlays[placement.placement_id]
    QTest.mousePress(overlay, Qt.MouseButton.LeftButton, pos=QPoint(overlay.width() - 2, overlay.height() - 2))
    QTest.mouseMove(overlay, QPoint(overlay.width() + 30, overlay.height() + 18))
    QTest.mouseRelease(overlay, Qt.MouseButton.LeftButton, pos=QPoint(overlay.width() + 30, overlay.height() + 18))
    app.processEvents()
    resized = window.image_placements[0]
    assert resized.width > 160 and resized.height > 80
    window.close()
    assert app is not None
