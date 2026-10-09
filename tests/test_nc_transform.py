from pathlib import Path

from cnc_program_sheet.nc_transform import COPPER, M06, M304, TransformMode, transform_file, transform_text


def test_m06_removes_every_line_containing_m6_like_the_batch_shortcut() -> None:
    source = "%\r\nN10 T1 M6\r\n; M60 is also removed by the original match\r\nN20 G0 X0\r\n"

    assert transform_text(source, M06) == "%\r\nN20 G0 X0\r\n"


def test_m06_preserves_the_original_lf_style_and_absence_of_final_newline() -> None:
    source = "N10 T1 M6\nN20 G0 X0"

    assert transform_text(source, "m6") == "N20 G0 X0"


def test_m304_inserts_blocks_after_the_first_matching_g43_and_m5() -> None:
    source = "N10 G43 H1 Z20 M8\nN20 M5\nM30"

    assert transform_text(source, M304) == (
        "N10 G43 H1 Z20 M8\r\n"
        "M304\r\n"
        "G05P10000\r\n"
        "N20 M5\r\n"
        "G05P0\r\n"
        "M300\r\n"
        "M30\r\n"
    )


def test_m304_is_idempotent_only_when_its_exact_marker_lines_already_exist() -> None:
    converted = transform_text("G43H1Z20.M8\nM5", TransformMode.M304)

    assert transform_text(converted, TransformMode.M304) == converted


def test_m304_keeps_the_shortcut_rule_that_m5_must_end_the_line() -> None:
    source = "G43H1Z20.M8\nM5 \nM30"

    assert transform_text(source, M304) == "G43H1Z20.M8\r\nM304\r\nG05P10000\r\nM5 \r\nM30\r\n"


def test_copper_transform_reproduces_the_existing_duplicate_safety_line_behaviour() -> None:
    # The original script looks only at the first line after %.  It therefore
    # inserts a second N110 block when one already appears later in the file.
    source = "%\nO0000(XB-1)\nN110G0G17G40G49G80G90\nM30"

    assert transform_text(source, COPPER) == (
        "%\r\n"
        "N110G0G17G40G49G80G90\r\n"
        "O0000(XB-1)\r\n"
        "N110G0G17G40G49G80G90\r\n"
        "M99\r\n"
    )


def test_copper_transform_replaces_only_the_last_m30_and_uses_extensionless_output() -> None:
    source = "O0000(TEST)\nM30\nM30"

    assert transform_text(source, "铜工") == (
        "N110G0G17G40G49G80G90\r\n"
        "O0000(TEST)\r\n"
        "M30\r\n"
        "M99\r\n"
    )


def test_transform_file_overwrites_for_m06_and_m304_but_copper_replaces_nc_with_no_suffix(tmp_path: Path) -> None:
    m06_path = tmp_path / "M06.NC"
    m06_path.write_bytes(b"T1 M6\nG0 X0\n")
    assert transform_file(m06_path, M06, encoding="ascii") == m06_path
    assert m06_path.read_bytes() == b"G0 X0\n"

    m304_path = tmp_path / "M304.NC"
    m304_path.write_bytes(b"G43H1Z20.M8\nM5\n")
    assert transform_file(m304_path, M304, encoding="ascii") == m304_path
    assert m304_path.read_bytes() == b"G43H1Z20.M8\r\nM304\r\nG05P10000\r\nM5\r\nG05P0\r\nM300\r\n"

    copper_path = tmp_path / "Copper.NC"
    copper_path.write_bytes(b"%\nO0000(COPPER)\nM30\n")
    target = transform_file(copper_path, COPPER, encoding="ascii")
    assert target == tmp_path / "Copper"
    assert not copper_path.exists()
    assert target.read_bytes() == b"%\r\nN110G0G17G40G49G80G90\r\nO0000(COPPER)\r\nM99\r\n"


def test_nc_only_shortcuts_leave_non_nc_files_untouched(tmp_path: Path) -> None:
    source = tmp_path / "program.txt"
    source.write_bytes(b"T1 M6\n")

    assert transform_file(source, M06, encoding="ascii") == source
    assert transform_file(source, COPPER, encoding="ascii") == source
    assert source.read_bytes() == b"T1 M6\n"
