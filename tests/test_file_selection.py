from pathlib import Path

from cnc_program_sheet.main import is_supported_nc_file


def test_accepts_nc_and_extensionless_copper_programs(tmp_path: Path) -> None:
    conventional = tmp_path / "O-1.NC"
    copper = tmp_path / "O0788"
    unrelated = tmp_path / "notes.txt"
    for path in (conventional, copper, unrelated):
        path.write_text("%\nM30\n", encoding="ascii")

    assert is_supported_nc_file(conventional)
    assert is_supported_nc_file(copper)
    assert not is_supported_nc_file(unrelated)
