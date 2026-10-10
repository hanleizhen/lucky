"""Create the multi-resolution Windows icon from the approved PNG artwork."""

from __future__ import annotations

from pathlib import Path

from PIL import Image


PROJECT = Path(__file__).resolve().parents[1]
SOURCE = PROJECT / "assets" / "cnc_smart_sheet_icon.png"
TARGET = PROJECT / "assets" / "cnc_smart_sheet_icon.ico"
SIZES = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]


def main() -> None:
    if not SOURCE.is_file():
        raise FileNotFoundError(f"Missing approved icon artwork: {SOURCE}")
    with Image.open(SOURCE) as image:
        image.convert("RGBA").save(TARGET, format="ICO", sizes=SIZES)
    print(TARGET)


if __name__ == "__main__":
    main()
