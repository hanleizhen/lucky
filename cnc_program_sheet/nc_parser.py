"""Conservative NC parsing.  Unknown is always safer than a plausible guess."""

from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation
from pathlib import Path

from .models import ProgramRecord, UNRECOGNIZED


TOOL_KEYWORDS: dict[str, tuple[str, ...]] = {
    "钻头": ("DRILL", "DRILLING", "钻头", "钻孔"),
    "倒角刀": ("CHAMFER", "CHAMFERING", "COUNTERSINK", "倒角", "锪孔"),
    "平底刀": ("ENDMILL", "END MILL", "MILLING", "FLAT ENDMILL", "FLAT END MILL", "平底刀", "铣刀"),
    "圆鼻刀": ("BULLNOSE", "BULL NOSE", "RADIUS MILL", "圆鼻刀"),
    "球刀": ("BALLNOSE", "BALL NOSE", "BALL MILL", "球刀"),
    "丝锥": ("TAP", "TAPPING", "丝锥"),
    "铰刀": ("REAMER", "REAMING", "铰刀"),
}
TOOL_TYPE_BY_KEYWORD = {
    keyword: tool_type for tool_type, keywords in TOOL_KEYWORDS.items() for keyword in keywords
}

NUMBER = r"([+-]?(?:\d+(?:\.\d*)?|\.\d+))"
NUMBER_VALUE = r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)"
DIAMETER_RE = re.compile(rf"(?<![A-Z0-9])(?:D|[ØΦφ])\s*[:=]?\s*{NUMBER}\s*(?:MM)?", re.I)
# Some CAM posts describe a drill as ``9 DRILL`` rather than ``D9DRILL``.
# This pattern is used only on header/comment evidence, so it cannot mistake a
# cycle parameter or an ordinary numeric coordinate for a tool diameter.
DRILL_DIAMETER_RE = re.compile(
    rf"(?<![A-Z0-9.])(?P<diameter>{NUMBER_VALUE})\s*(?:MM\s*)?DRILL(?:ING)?\b",
    re.I,
)
RADIUS_RE = re.compile(rf"(?<![A-Z0-9])R\s*[:=]?\s*{NUMBER}\s*(?:MM)?", re.I)
# Some controllers put a compact, but unambiguous, tool declaration at the
# beginning of a program, e.g. ``(T1)D4R0.2|H1)``.  The closing parenthesis
# after T1 prevents the usual word-boundary rule from seeing D/R as a pair.
# Restrict this pattern to the complete ``(T...)D...`` header form so ordinary
# D offsets and cycle R planes remain excluded.
TOOL_HEADER_PREFIX = r"(?:\(\s*)?T\s*\d+(?:\s*\)|\s*\|)?\s*"
TOOL_HEADER_SPEC_RE = re.compile(
    rf"{TOOL_HEADER_PREFIX}D\s*(?P<diameter>{NUMBER_VALUE})(?:\s*R\s*(?P<radius>{NUMBER_VALUE}))?",
    re.I,
)
# ``D1C40`` (often written as ``(T1)D1C40|H1)``) is an explicit chamfer-tool
# declaration in the supplied NC convention.  It is only considered while
# examining program-header/comment evidence, never arbitrary later G-code.
TOOL_HEADER_CHAMFER_RE = re.compile(
    rf"{TOOL_HEADER_PREFIX}D\s*(?P<diameter>{NUMBER_VALUE})\s*C\s*(?P<angle>{NUMBER_VALUE})(?![A-Z0-9])",
    re.I,
)
# T-type cutters in this shop are explicitly described in a header/comment as
# ``T6R0.4``. The leading T is the cutter-size convention, not the controller
# tool number (which is supplied separately as e.g. ``T1|...|H1``).
T_SHAPE_TOOL_RE = re.compile(
    rf"(?<![A-Z0-9])T\s*(?P<diameter>{NUMBER_VALUE})\s*R\s*(?P<radius>{NUMBER_VALUE})(?![A-Z0-9])",
    re.I,
)
# Some supplied programs identify a T-type cutter with a decimal T size in
# the tool-description header, e.g. ``( T1 | T5.9 | H1 | D1 | ... )``.  The
# first T word is the controller tool number; the second is the actual cutter
# designation that belongs in the template D column.  Requiring that explicit
# two-column header form prevents an executable ``T1 M6`` tool change from
# ever being treated as a cutter size.
T_STYLE_HEADER_RE = re.compile(
    rf"(?:\(\s*)?T\s*\d+\s*\|\s*T\s*(?P<diameter>{NUMBER_VALUE})(?=\s*(?:\||\)|H\s*\d|TOOL\b|$))",
    re.I,
)
COMPACT_CHAMFER_RE = re.compile(
    rf"(?<![A-Z0-9])D\s*(?P<diameter>{NUMBER_VALUE})\s*C\s*(?P<angle>{NUMBER_VALUE})(?![A-Z0-9])",
    re.I,
)
TOOL_NUMBER_RE = re.compile(r"T\s*(?P<tool_number>\d+)(?=\s*(?:M\d+|\)|\||D|H|\s|$))", re.I)
TOOL_CHANGE_RE = re.compile(r"T\s*(?P<tool_number>\d+)\s*M0?6(?!\d)", re.I)
TOOL_DESCRIPTION_RE = re.compile(r"^\s*\(\s*T\s*(?P<tool_number>\d+)\s*\|.*\)\s*$", re.I)
# Mastercam emits compact blocks such as ``N160G98G81Z-18.5R1.`` and
# ``G1Z-.8``. G/Z must therefore be recognised immediately after another
# numeric word, while the trailing guard still prevents partial G-code matches.
Z_RE = re.compile(rf"Z\s*{NUMBER}", re.I)
CYCLE_RE = re.compile(r"G(8[1-9])(?!\d)", re.I)
REFERENCE_RETURN_RE = re.compile(r"G(?:28|30)(?!\d)", re.I)
CHAMFER_RE = re.compile(rf"(?<![A-Z0-9])C\s*{NUMBER}\s*(?:°|DEG(?:REE)?S?|度)", re.I)
MOTION_RE = re.compile(r"(?:G0?[0-3](?!\d)|G8[1-9](?!\d)|[XYZ]\s*[+-]?(?:\d|\.))", re.I)
COORDINATE_RE = re.compile(r"\b[XYZF]\s*[+-]?(?:\d|\.)", re.I)
OFFSET_RE = re.compile(r"\bG4[12]\s*D\s*\d", re.I)


def _read_nc(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "gb18030", "shift_jis", "latin-1"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _normalise_number(value: str) -> str:
    try:
        number = Decimal(value)
    except InvalidOperation:
        return value
    if number == number.to_integral():
        return str(number.quantize(Decimal("1")))
    return format(number.normalize(), "f")


def _without_comments(line: str) -> str:
    return re.sub(r"\([^)]*\)|;.*$", "", line)


def _comment_text(line: str) -> str:
    parts = re.findall(r"\(([^)]*)\)", line)
    if ";" in line:
        parts.append(line.split(";", 1)[1])
    return " ".join(parts)


def _has_tool_keyword(line: str) -> bool:
    upper = line.upper()
    return any(keyword in upper for keyword in TOOL_TYPE_BY_KEYWORD)


def _tool_type(evidence_lines: list[str]) -> str | None:
    found: set[str] = set()
    for line in evidence_lines:
        upper = line.upper()
        for keyword, tool_type in TOOL_TYPE_BY_KEYWORD.items():
            if keyword in upper:
                found.add(tool_type)
    if len(found) == 1:
        return next(iter(found))
    return None


def _unique_values(values: list[str]) -> list[str]:
    return list(dict.fromkeys(values))


def _find_diameter(evidence_lines: list[str]) -> tuple[str | None, bool]:
    """Return an explicit tool diameter, never a standalone compensation D number."""

    values: list[str] = []
    for line in evidence_lines:
        header_match = TOOL_HEADER_SPEC_RE.search(line)
        if header_match:
            values.append(_normalise_number(header_match.group("diameter")))
            continue
        chamfer_match = TOOL_HEADER_CHAMFER_RE.search(line) or COMPACT_CHAMFER_RE.search(line)
        if chamfer_match:
            values.append(_normalise_number(chamfer_match.group("diameter")))
            continue
        drill_match = DRILL_DIAMETER_RE.search(line)
        if drill_match:
            values.append(_normalise_number(drill_match.group("diameter")))
            continue
        upper = line.upper()
        if OFFSET_RE.search(upper):
            continue
        match = DIAMETER_RE.search(line)
        if not match:
            continue
        # D numbers are trustworthy only in a tool description, an explicit comment,
        # or as a D/R pair. This excludes normal G41/G42 compensation words.
        has_pair = bool(RADIUS_RE.search(line))
        if not (_has_tool_keyword(line) or has_pair or "Ø" in line or "Φ" in line or "φ" in line):
            continue
        values.append(_normalise_number(match.group(1)))
    values = _unique_values(values)
    return (values[0], False) if len(values) == 1 else (None, len(values) > 1)


def _find_radius(evidence_lines: list[str]) -> tuple[str | None, bool]:
    """Accept only a radius described with the tool; never G81/G83 R-plane values."""

    values: list[str] = []
    for line in evidence_lines:
        header_match = TOOL_HEADER_SPEC_RE.search(line)
        if header_match:
            radius = header_match.group("radius")
            if radius:
                values.append(_normalise_number(radius))
            continue
        code = _without_comments(line)
        if CYCLE_RE.search(code) or COORDINATE_RE.search(code):
            continue
        match = RADIUS_RE.search(line)
        if not match:
            continue
        if not (_has_tool_keyword(line) or DIAMETER_RE.search(line) or _comment_text(line)):
            continue
        values.append(_normalise_number(match.group(1)))
    values = _unique_values(values)
    return (values[0], False) if len(values) == 1 else (None, len(values) > 1)


def _find_t_shape_tool(evidence_lines: list[str]) -> tuple[tuple[str, str] | None, bool]:
    """Find an explicit ``T<size>R<radius>`` T-cutter description.

    The match is intentionally limited to evidence/header lines. In executable
    code a ``T`` word is normally a controller tool number and must never be
    converted into a cutter size.
    """

    values = _unique_values(
        [
            f"{_normalise_number(match.group('diameter'))}|{_normalise_number(match.group('radius'))}"
            for line in evidence_lines
            for match in T_SHAPE_TOOL_RE.finditer(line)
        ]
    )
    if len(values) != 1:
        return None, len(values) > 1
    diameter, radius = values[0].split("|", 1)
    return (diameter, radius), False


def _find_t_style_tool(evidence_lines: list[str]) -> tuple[str | None, bool]:
    """Find a header-only decimal T-cutter size such as ``T5.9``.

    This convention has no tool-radius declaration.  It is deliberately
    separate from ``T6R0.4`` so R is never invented from the T value.
    """

    values = _unique_values(
        [
            _normalise_number(match.group("diameter"))
            for line in evidence_lines
            for match in T_STYLE_HEADER_RE.finditer(line)
        ]
    )
    return (values[0], False) if len(values) == 1 else (None, len(values) > 1)


def _find_chamfer(evidence_lines: list[str]) -> tuple[str | None, bool]:
    values: list[str] = []
    for line in evidence_lines:
        values.extend(_normalise_number(match.group(1)) for match in CHAMFER_RE.finditer(line))
        for pattern in (TOOL_HEADER_CHAMFER_RE, COMPACT_CHAMFER_RE):
            values.extend(_normalise_number(match.group("angle")) for match in pattern.finditer(line))
    values = _unique_values(values)
    return (values[0], False) if len(values) == 1 else (None, len(values) > 1)


def _find_tool_number(lines: list[str]) -> tuple[str | None, bool]:
    values = _unique_values([match.group("tool_number") for line in lines for match in TOOL_NUMBER_RE.finditer(line)])
    return (values[0], False) if len(values) == 1 else (None, len(values) > 1)


def _has_contour_motion(executable_lines: list[str]) -> bool:
    """Identify modal linear/circular contour moves without accepting cycles.

    A common post outputs ``G1 Z...`` on one block and the following ``X`` / ``Y``
    moves without repeating G1.  Treating each line independently leaves that
    valid flat-end-mill program without a recognised machining type.
    """

    motion_word = re.compile(r"G0?([0-3])(?!\d)", re.I)
    xy_word = re.compile(r"[XY]\s*[+-]?(?:\d|\.)", re.I)
    modal_motion: str | None = None
    for line in executable_lines:
        match = motion_word.search(line)
        if match:
            modal_motion = match.group(1)
        if modal_motion in {"1", "2", "3"} and xy_word.search(line):
            return True
    return False


def _z_display(value: str) -> str:
    """Keep program precision but render Mastercam's ``-.8`` as ``-0.8``."""

    result = value.lstrip("+")
    if result.startswith("-."):
        result = "-0" + result[1:]
    elif result.startswith("."):
        result = "0" + result
    return result[:-1] if result.endswith(".") else result


def _extract_evidence(lines: list[str]) -> tuple[list[str], list[str]]:
    """Use comments everywhere plus textual header lines before the first real move."""

    evidence: list[str] = []
    executable: list[str] = []
    in_header = True
    for line in lines:
        comment = _comment_text(line)
        code = _without_comments(line)
        # Preserve the original compact header as a single unit: stripping
        # ``(T1)`` would otherwise separate it from its following D/R values.
        if in_header and (TOOL_HEADER_SPEC_RE.search(line) or TOOL_HEADER_CHAMFER_RE.search(line)):
            evidence.append(line)
        if comment:
            evidence.append(comment)
        if code.strip():
            executable.append(code)
        if in_header:
            if MOTION_RE.search(code):
                in_header = False
            elif code.strip():
                evidence.append(code)
    return evidence, executable


def parse_text(text: str, filename: str, source_path: Path | None = None) -> ProgramRecord:
    """Parse one NC source string while retaining uncertainty for user review."""

    path = source_path or Path(filename)
    record = ProgramRecord(source_path=path, program_name=filename)
    lines = text.splitlines()
    evidence_lines, executable_lines = _extract_evidence(lines)

    tool_number, tool_number_ambiguous = _find_tool_number(evidence_lines + executable_lines)
    record.tool_number = f"T{tool_number}" if tool_number else UNRECOGNIZED
    if tool_number_ambiguous:
        record.warnings.append("发现多个刀具号，未自动填写 T")
    elif not tool_number:
        record.warnings.append("未找到明确刀具号")

    t_shape, t_shape_ambiguous = _find_t_shape_tool(evidence_lines)
    t_style, t_style_ambiguous = _find_t_style_tool(evidence_lines)
    if t_shape:
        # Per the supplied program-sheet convention, T-type cutter sizes are
        # displayed in the D column as T6, T8, etc., not rewritten as D6/D8.
        diameter, radius = t_shape
        diameter_ambiguous = radius_ambiguous = False
        record.diameter = f"T{diameter}"
        record.radius = f"R{radius}"
    elif t_style:
        # Header form: ``T1 | T5.9 | H1 | D1``.  T5.9 is the actual
        # cutter specification; D1 is only the controller compensation word.
        diameter = t_style
        radius = None
        diameter_ambiguous = radius_ambiguous = False
        record.diameter = f"T{diameter}"
        record.radius = ""
    else:
        diameter, diameter_ambiguous = _find_diameter(evidence_lines)
        # Keep the drawing/program convention requested for this template: D4,
        # D19.5, etc. Never convert the value to a phi symbol.
        record.diameter = f"D{diameter}" if diameter else UNRECOGNIZED
        radius, radius_ambiguous = _find_radius(evidence_lines)
        record.radius = f"R{radius}" if radius else ""
    if t_shape_ambiguous or t_style_ambiguous or diameter_ambiguous:
        record.warnings.append("发现多个刀具直径，未自动填写 D")
    elif not t_shape and not t_style and not diameter:
        record.warnings.append("未找到明确刀具直径")
    if t_shape_ambiguous or radius_ambiguous:
        record.warnings.append("发现多个刀具半径，未自动填写 R")

    chamfer, chamfer_ambiguous = _find_chamfer(evidence_lines)
    cycle_codes = _unique_values([f"G{match.group(1)}" for line in executable_lines for match in CYCLE_RE.finditer(line)])
    detected_type = _tool_type(evidence_lines)
    # Recognise the program's tool kind internally. It is deliberately not
    # copied to the template's blank “区分” column.
    if detected_type is None and (t_shape or t_style):
        detected_type = "T型刀"
    elif detected_type is None and chamfer:
        detected_type = "倒角刀"
    elif detected_type is None and cycle_codes:
        detected_type = "钻头"
    elif detected_type is None and radius:
        detected_type = "圆鼻刀"
    elif detected_type is None and diameter and _has_contour_motion(executable_lines):
        detected_type = "平底刀"
    record.detected_tool_type = detected_type or UNRECOGNIZED
    record.tool_type = ""

    if len(cycle_codes) > 1:
        record.warnings.append("发现多个固定循环 G 代码，加工 DATA 未自动填写")
    if chamfer_ambiguous:
        record.warnings.append("发现多个倒角角度，加工 DATA 未自动填写")
    if record.detected_tool_type == "倒角刀":
        record.machining_data = f"C{chamfer}°" if chamfer and not chamfer_ambiguous else UNRECOGNIZED
    elif len(cycle_codes) == 1:
        # A literal fixed-cycle code is safe to report even if the tool description is absent.
        record.machining_data = cycle_codes[0]
    else:
        # Per the program-sheet convention, 加工DATA is used only for a
        # confirmed fixed cycle (or a chamfer angle). No G81/G83/G85-style
        # cycle means this cell stays blank, even when the tool type itself is
        # not otherwise named in the NC comments.
        record.machining_data = ""
    if record.machining_data == UNRECOGNIZED:
        record.warnings.append("未找到可唯一确认的加工 DATA")

    z_values: list[tuple[Decimal, str]] = []
    for line in executable_lines:
        # G28/G30 is a return-to-reference command. In lines such as
        # ``G91G28Z0`` its Z0 is a home-return waypoint, not cutting depth.
        # Including it incorrectly turns an entirely positive-Z program into
        # Z0, so exclude the whole reference-return block before comparison.
        if REFERENCE_RETURN_RE.search(line):
            continue
        for match in Z_RE.finditer(line):
            try:
                z_values.append((Decimal(match.group(1)), match.group(1)))
            except InvalidOperation:
                continue
    if z_values:
        _, original = min(z_values, key=lambda item: item[0])
        record.depth = f"Z{_z_display(original)}"
    else:
        record.warnings.append("未找到实际 Z 坐标")

    return record


def parse_file(path: str | Path) -> ProgramRecord:
    nc_path = Path(path)
    return parse_file_records(nc_path)[0]


def parse_records_text(text: str, filename: str, source_path: Path | None = None) -> list[ProgramRecord]:
    """Parse one NC file into one row per actual M6 tool-change section.

    Single-tool files preserve the old one-record behaviour. Multi-tool files
    retain the source order and parse each Z depth only from that tool's block.
    """

    nc_path = source_path or Path(filename)
    lines = text.splitlines()
    changes = [(index, match.group("tool_number")) for index, line in enumerate(lines) if (match := TOOL_CHANGE_RE.search(_without_comments(line)))]
    if len(changes) <= 1:
        record = parse_text(text, filename, nc_path)
        record.source_tool_index = 0
        return [record]

    descriptions: dict[str, str] = {}
    for line in lines:
        match = TOOL_DESCRIPTION_RE.match(line)
        if match:
            descriptions.setdefault(str(int(match.group("tool_number"))), line)

    records: list[ProgramRecord] = []
    for tool_index, (start, tool_number) in enumerate(changes):
        end = changes[tool_index + 1][0] if tool_index + 1 < len(changes) else len(lines)
        normalised_number = str(int(tool_number))
        # Only the matching header is included, preventing D/R descriptions of
        # other tools from being mistaken for this tool's geometry.
        section_lines = [descriptions[normalised_number]] if normalised_number in descriptions else []
        section_lines.extend(lines[start:end])
        record = parse_text("\n".join(section_lines), filename, nc_path)
        record.source_tool_index = tool_index
        if record.tool_number == UNRECOGNIZED:
            record.tool_number = f"T{normalised_number}"
        records.append(record)
    return records


def parse_file_records(path: str | Path) -> list[ProgramRecord]:
    nc_path = Path(path)
    return parse_records_text(_read_nc(nc_path), nc_path.name, nc_path)
