from pathlib import Path

from openpyxl import load_workbook
from PIL import Image

from cnc_program_sheet.excel_service import export_workbook
from cnc_program_sheet.models import ImagePlacement


def test_export_anchors_user_image_at_selected_cell_with_requested_size(tmp_path: Path) -> None:
    source_template = Path("assets/CNC程序单.xlsx")
    original_template = source_template.read_bytes()
    source_image = tmp_path / "mark.png"
    Image.new("RGB", (20, 10), "#2f75b5").save(source_image)
    output = tmp_path / "with-image.xlsx"

    export_workbook(
        source_template,
        output,
        {"J3": "08 / 10 / 2026"},
        images=[ImagePlacement(source_path=source_image, anchor="B4", width=120, height=60)],
    )

    assert source_template.read_bytes() == original_template
    workbook = load_workbook(output)
    try:
        sheet = workbook["Sheet1"]
        assert len(sheet._images) == 1
        image = sheet._images[0]
        assert image.anchor._from.col == 1  # B column, zero based in DrawingML
        assert image.anchor._from.row == 3  # Excel row 4, zero based in DrawingML
        # DrawingML stores exported dimensions in EMUs (1 px = 9,525 EMUs).
        assert image.anchor.ext.width == 120 * 9525
        assert image.anchor.ext.height == 60 * 9525
    finally:
        workbook.close()
