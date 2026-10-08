from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


UNRECOGNIZED = "未识别"


@dataclass(slots=True)
class ProgramRecord:
    """A deliberately conservative interpretation of one NC file."""

    source_path: Path
    program_name: str
    # A single NC file may contain several physical tool-change sections.
    # This zero-based position keeps re-parse stable without appearing in the
    # Excel program sheet.
    source_tool_index: int = 0
    # The supplied template keeps the “区分” column blank.  Tool-kind
    # detection is retained separately for safe parsing decisions, but is not
    # written as Chinese text into that column.
    tool_type: str = ""
    detected_tool_type: str = UNRECOGNIZED
    tool_number: str = UNRECOGNIZED
    diameter: str = UNRECOGNIZED
    radius: str = ""
    machining_data: str = UNRECOGNIZED
    depth: str = UNRECOGNIZED
    warnings: list[str] = field(default_factory=list)

    @property
    def has_unrecognized_values(self) -> bool:
        return any(
            value == UNRECOGNIZED
            for value in (self.tool_number, self.diameter, self.machining_data, self.depth)
        )

    @property
    def status(self) -> str:
        return "；".join(self.warnings) if self.warnings else "已解析"


@dataclass(frozen=True, slots=True)
class ImagePlacement:
    """One operator-inserted image anchored to an Excel cell."""

    source_path: Path
    anchor: str
    width: int
    height: int
