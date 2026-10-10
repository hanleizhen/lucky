from pathlib import Path

from cnc_program_sheet.models import UNRECOGNIZED
from cnc_program_sheet.nc_parser import parse_records_text, parse_text


def test_drill_uses_header_diameter_not_cycle_r_plane() -> None:
    record = parse_text(
        """O0002
(D19.5DRILL)
T01 M06
G90 G80
G81 X0 Y0 Z-18.5 R1. F120
G80
M30
""",
        "O-2.NC",
        Path("O-2.NC"),
    )
    assert record.program_name == "O-2.NC"
    assert record.tool_number == "T01"
    assert record.tool_type == ""
    assert record.detected_tool_type == "钻头"
    assert record.diameter == "D19.5"
    assert record.radius == ""
    assert record.machining_data == "G81"
    assert record.depth == "Z-18.5"


def test_solidcam_tool_list_entry_supplies_diameter_without_a_type_keyword() -> None:
    """A formal ``T01 : D4`` list entry is geometry, not a D offset."""

    record = parse_text(
        """%
(TOOLS LIST)
(T01 : D4. L37.5 CL15. 1F)
G00 G91 G28 Z0.
N1 (D-DRILL2)
G98 G81 Z-1.5 R2. F50.
M30
%""",
        "3011-1.NC",
    )

    assert (record.tool_number, record.detected_tool_type, record.diameter, record.radius, record.machining_data, record.depth) == (
        "T01",
        "钻头",
        "D4",
        "",
        "G81",
        "Z-1.5",
    )


def test_solidcam_tool_list_entry_handles_corrupted_type_text() -> None:
    """The structured D field stays reliable even if a Chinese type label is garbled."""

    record = parse_text(
        """%
(TOOLS LIST)
(T02 : 资头 D6. L35. CL24. 1F)
G00 G91 G28 Z0.
N1 (D-DRILL1-1)
G98 G83 Z-23. R2. Q2. F50.
M30
%""",
        "3011-2.NC",
    )

    assert (record.tool_number, record.detected_tool_type, record.diameter, record.radius, record.machining_data, record.depth) == (
        "T02",
        "钻头",
        "D6",
        "",
        "G83",
        "Z-23",
    )


def test_solidcam_tool_list_rule_does_not_accept_executable_d_or_cycle_r() -> None:
    record = parse_text(
        """%
(TOOLS LIST)
(T01 : D-DRILL2 L37.5 CL15. 1F)
G41 D1
G81 X0 Y0 Z-18.5 R1.
M30
%""",
        "solidcam-offset.NC",
    )

    assert (record.diameter, record.radius, record.machining_data, record.depth) == (
        UNRECOGNIZED,
        "",
        "G81",
        "Z-18.5",
    )


def test_minimum_real_z_is_used() -> None:
    record = parse_text(
        """(D5 DRILL)
G81 X0 Y0 Z-2.46 R1.
G01 Z-4.14
G01 Z-5.08
G01 Z-7.50
""",
        "O-DEPTH.NC",
    )
    assert record.depth == "Z-7.50"


def test_explicit_d_and_r_pair_is_allowed_but_type_is_not_guessed() -> None:
    record = parse_text(
        """(T06 D1 R0.5)
G00 X0 Y0 Z3.
G01 Z-0.8 F100
""",
        "O-6.NC",
    )
    assert record.diameter == "D1"
    assert record.radius == "R0.5"
    assert record.tool_type == ""
    assert record.detected_tool_type == "圆鼻刀"
    assert record.tool_number == "T06"
    assert record.depth == "Z-0.8"


def test_compact_t_header_supplies_d_and_explicit_r() -> None:
    """Real controller header: the ')' after T1 is part of the format."""

    record = parse_text(
        """(T1)D4R0.2|H1)
G00 X0 Y0 Z4.03
G01 Z-0.8 F100
""",
        "CORE-1.NC",
    )
    assert record.diameter == "D4"
    assert record.radius == "R0.2"
    assert record.tool_type == ""
    assert record.detected_tool_type == "圆鼻刀"
    assert record.tool_number == "T1"


def test_compact_t_header_without_r_does_not_invent_a_radius() -> None:
    record = parse_text(
        """(T1)D4|H1)
G00 X0 Y0 Z3.953
""",
        "CORE-3.NC",
    )
    assert record.diameter == "D4"
    assert record.radius == ""


def test_compact_d_c_header_is_an_explicit_chamfer_tool() -> None:
    record = parse_text(
        """(T1)D1C40|H1)
G00 X0 Y0 Z5.
G01 Z-0.8
""",
        "CORE-CHAMFER.NC",
    )
    assert record.tool_type == ""
    assert record.detected_tool_type == "倒角刀"
    assert record.tool_number == "T1"
    assert record.diameter == "D1"
    assert record.radius == ""
    assert record.machining_data == "C40°"


def test_compensation_d_and_cycle_r_are_never_tool_dimensions() -> None:
    record = parse_text(
        """T01 M06
G43 H01
G41 D1
G81 X0 Y0 Z-18.5 R1.
""",
        "offset.NC",
    )
    assert record.diameter == UNRECOGNIZED
    assert record.radius == ""
    assert record.machining_data == "G81"


def test_chamfer_populates_the_single_machining_data_column() -> None:
    record = parse_text(
        """(D1 CHAMFER C40°)
G00 Z5.
G01 Z-0.8
""",
        "O-1.NC",
    )
    assert record.tool_type == ""
    assert record.detected_tool_type == "倒角刀"
    assert record.diameter == "D1"
    assert record.machining_data == "C40°"
    assert record.depth == "Z-0.8"


def test_mastercam_compact_blocks_match_the_seven_supplied_reference_programs() -> None:
    programs = {
        "O-1.NC": ("(T1|D1C40|H1)", "N170G1Z-.8F500.", "T1", "倒角刀", "D1", "", "C40°", "Z-0.8"),
        "O-2.NC": ("(T1|D19.5DRILL|H1)", "N160G98G81Z-18.5R1.F18.", "T1", "钻头", "D19.5", "", "G81", "Z-18.5"),
        "O-3.NC": ("(T1|D19.5DRILL|H1)", "N160G98G83Z-18.5R1.Q0.F18.", "T1", "钻头", "D19.5", "", "G83", "Z-18.5"),
        "O-4.NC": ("(T1|D19.5DRILL|H1)", "N160G98G85Z-18.5R1.F18.", "T1", "钻头", "D19.5", "", "G85", "Z-18.5"),
        "O-5.NC": ("(T1|D1|H1)", "N180G1X-125.7F300.\nN170G1Z-.8F500.", "T1", "平底刀", "D1", "", "", "Z-0.8"),
        "O-6.NC": ("(T1|D1R0.5|H1)", "N170G1Z-.8F500.", "T1", "圆鼻刀", "D1", "R0.5", "", "Z-0.8"),
        "O-7.NC": ("(T1|D1R0.2|H1)", "N170G1Z-.8F500.", "T1", "圆鼻刀", "D1", "R0.2", "", "Z-0.8"),
    }
    for filename, (header, cutting_code, tool_number, detected_type, diameter, radius, data, depth) in programs.items():
        record = parse_text(f"%\nO0000({filename[:-3]})\n{header}\nN110G0G17G40G49G80G90\n{cutting_code}\nM30\n%", filename)
        assert record.tool_type == ""
        assert record.tool_number == tool_number
        assert record.detected_tool_type == detected_type
        assert (record.diameter, record.radius, record.machining_data, record.depth) == (diameter, radius, data, depth)


def test_copper_program_header_and_compact_z_words_are_supported() -> None:
    record = parse_text(
        """%
O0000(O0782)
(T2|D4R0.2|H2|XY STOCK TO LEAVE - .08|Z STOCK TO LEAVE - 0.)
N160Z.2
N170G1Z-.143F100.
N1180G1X-20.0Y-10.0Z-2.05F800.
M30
%""",
        "O0782",
    )
    assert record.program_name == "O0782"
    assert record.tool_type == ""
    assert record.tool_number == "T2"
    assert record.detected_tool_type == "圆鼻刀"
    assert (record.diameter, record.radius, record.machining_data, record.depth) == ("D4", "R0.2", "", "Z-2.05")


def test_terminal_decimal_point_in_z_is_rendered_cleanly() -> None:
    record = parse_text("(T1|D0.6|H1)\nG1Z-1.F100.", "O0784")
    assert record.depth == "Z-1"


def test_positive_cutting_z_ignores_g28_reference_zero() -> None:
    record = parse_text(
        """%
O0000(G-2)
(T1|D1C40|H1)
N120G91G28Z0.
N150G43H1Z50.M8
N160Z20.2
N170G1Z19.2F500.
N240G0Z50.
N360G91G28Z0.M9
M30
%""",
        "G-2.NC",
    )
    assert record.depth == "Z19.2"


def test_number_drill_header_is_a_real_diameter_even_without_d_prefix() -> None:
    record = parse_text(
        """%
O0000(DB-2)
(T1|9 DRILL|H1)
N160G98G81Z-1.R1.F20.
M30
%""",
        "DB-2.NC",
    )
    assert (record.tool_number, record.detected_tool_type, record.diameter, record.machining_data, record.depth) == (
        "T1",
        "钻头",
        "D9",
        "G81",
        "Z-1",
    )


def test_modal_g1_xy_contour_marks_db8_style_program_as_flat_end_mill() -> None:
    record = parse_text(
        """%
O0000(DB-8)
(T1|D12|H1)
N150G43H1Z20.M8
N160Z.2
N170G1Z-3.667F500.
N180X164.F800.
N190Y-74.7
N220Z-7.333F500.
N230X164.F800.
N570Z-41.F500.
N580X164.F800.
N620G0Z20.
M30
%""",
        "DB-8.NC",
    )
    assert record.detected_tool_type == "平底刀"
    assert (record.diameter, record.machining_data, record.depth) == ("D12", "", "Z-41")


def test_t_shape_header_writes_t_size_in_d_column_and_keeps_radius() -> None:
    record = parse_text(
        """%
O0000(T-1)
(T1|T6R0.4|H1)
N150G43H1Z20.M8
N160Z1.2
N170G1Z1.F500.
N180X-129.2
M30
%""",
        "t-1.NC",
    )
    assert (record.tool_number, record.detected_tool_type, record.diameter, record.radius, record.machining_data, record.depth) == (
        "T1",
        "T型刀",
        "T6",
        "R0.4",
        "",
        "Z1",
    )


def test_decimal_t_style_header_writes_t_size_not_compensation_d() -> None:
    """3011-B-74 convention: the second T word is the T-cutter size."""

    record = parse_text(
        """%
O0000(3011-B-74)
( T1 | T5.9 | H1 | D1 | WEAR COMP | TOOL DIA. - 5.9 | XY STOCK TO LEAVE - .02 | Z STOCK TO LEAVE - 0. )
G21
G0 G17 G40 G49 G80 G90
G91 G28 Z0.
T1 M6
G0 G90 G54 X-98.263 Y-111.821 A0. S3500 M3
G43 H1 Z20. M8
Z-5.764
G1 Z-5.964 F500.
M30
%""",
        "3011-B-74.NC",
    )

    assert (record.tool_number, record.detected_tool_type, record.diameter, record.radius, record.depth) == (
        "T1",
        "T型刀",
        "T5.9",
        "",
        "Z-5.964",
    )


def test_multi_tool_program_creates_one_record_per_m6_with_scoped_depth() -> None:
    records = parse_records_text(
        """%
O0000(0505)
(T1|D3R0.2|H1)
(T3|D1|H3)
(T2|D1R0.5|H2)
N130T1M6
N170G1Z-0.704F500.
N580G0Z20.
N630T3M6
N670G1Z-6.554F500.
N800G0Z20.
N900T2M6
N910G98G85Z-1.804R1.F20.
N920G80
M30
%""",
        "0505.NC",
    )

    assert [(record.program_name, record.tool_number, record.diameter, record.radius, record.machining_data, record.depth) for record in records] == [
        ("0505.NC", "T1", "D3", "R0.2", "", "Z-0.704"),
        ("0505.NC", "T3", "D1", "", "", "Z-6.554"),
        ("0505.NC", "T2", "D1", "R0.5", "G85", "Z-1.804"),
    ]
    assert [record.source_tool_index for record in records] == [0, 1, 2]
