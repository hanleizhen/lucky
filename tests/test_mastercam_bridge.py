from pathlib import Path

from cnc_program_sheet.main import parse_launch_arguments


def test_mastercam_bridge_accepts_quoted_paths_with_spaces() -> None:
    paths = parse_launch_arguments(
        ["--import-nc", r"C:\Mastercam Files\Part 1.NC", r"D:\Posted NC\Part 2.NC"]
    )

    assert paths == [
        Path(r"C:\Mastercam Files\Part 1.NC"),
        Path(r"D:\Posted NC\Part 2.NC"),
    ]


def test_mastercam_bridge_accepts_repeated_import_options() -> None:
    paths = parse_launch_arguments(
        ["--import-nc", r"C:\NC\A.NC", "--import-nc", r"C:\NC\B.NC"]
    )

    assert paths == [Path(r"C:\NC\A.NC"), Path(r"C:\NC\B.NC")]


def test_regular_launch_has_no_import_paths() -> None:
    assert parse_launch_arguments([]) == []
