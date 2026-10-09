"""Compatibility transforms for the shop's three existing NC shortcuts.

The application deliberately keeps these transforms separate from the NC
parser.  They are not "smart" post-processors: their purpose is to reproduce
the established batch/PowerShell shortcut behaviour exactly, including the
copper shortcut's historical duplicate safety-line insertion and its
extensionless output file.

All file transforms use the Windows ANSI/default encoding used by the original
scripts.  The text-level functions are intentionally public as well, so the
GUI can show and allow edits to a result before asking this module to write it
back to disk.
"""

from __future__ import annotations

import locale
import os
import re
from enum import Enum
from pathlib import Path


class TransformMode(str, Enum):
    """The three transformation shortcuts presented by the desktop UI."""

    M06 = "M06"
    M304 = "M304"
    COPPER = "TG"


# Short named constants are convenient for callers which do not need to
# import Enum members directly.  ``TG`` is the label used in the imported-file
# list for the copper-program conversion page.
M06 = TransformMode.M06
M304 = TransformMode.M304
TG = TransformMode.COPPER
COPPER = TransformMode.COPPER

_COPPER_INSERT_LINE = "N110G0G17G40G49G80G90"
_BOM_ENCODINGS: tuple[tuple[bytes, str], ...] = (
    (b"\xff\xfe\x00\x00", "utf-32"),
    (b"\x00\x00\xfe\xff", "utf-32"),
    (b"\xff\xfe", "utf-16"),
    (b"\xfe\xff", "utf-16"),
    (b"\xef\xbb\xbf", "utf-8-sig"),
)


def _coerce_mode(mode: TransformMode | str) -> TransformMode:
    if isinstance(mode, TransformMode):
        return mode

    value = str(mode).strip().casefold()
    aliases = {
        "m6": TransformMode.M06,
        "m06": TransformMode.M06,
        "m304": TransformMode.M304,
        "tg": TransformMode.COPPER,
        "copper": TransformMode.COPPER,
        "铜工": TransformMode.COPPER,
    }
    try:
        return aliases[value]
    except KeyError as exc:
        raise ValueError(f"不支持的 NC 转换方式：{mode}") from exc


def windows_default_encoding() -> str:
    """Return Python's closest equivalent to ``Encoding.Default``.

    ``mbcs`` maps to the Windows ANSI code page (the encoding selected by
    ``[System.Text.Encoding]::Default``).  It is Windows-only, so tests and
    source inspection on another platform use the active locale instead.
    """

    return "mbcs" if os.name == "nt" else locale.getpreferredencoding(False)


def decode_windows_default(data: bytes, *, encoding: str | None = None) -> str:
    """Decode NC bytes as the original PowerShell/.NET scripts do.

    ``ReadAllText(..., Encoding.Default)`` and ``Get-Content`` recognise a
    Unicode byte-order mark before falling back to the system ANSI code page.
    Keeping that behaviour avoids exposing a BOM as the first NC character.
    """

    for marker, bom_encoding in _BOM_ENCODINGS:
        if data.startswith(marker):
            return data.decode(bom_encoding)
    return data.decode(encoding or windows_default_encoding())


def encode_windows_default(text: str, *, encoding: str | None = None) -> bytes:
    """Encode output exactly as ``-Encoding Default`` / ``Encoding.Default``."""

    return text.encode(encoding or windows_default_encoding())


def _split_normalised_lines(text: str) -> tuple[list[str], bool]:
    """Match the explicit newline normalisation used by the M06/copper scripts."""

    normalised = text.replace("\r\n", "\n").replace("\r", "\n")
    had_final_newline = normalised.endswith("\n")
    lines = normalised.split("\n")
    if had_final_newline and lines and lines[-1] == "":
        lines = lines[:-1]
    return lines, had_final_newline


def _powershell_content_lines(text: str) -> list[str]:
    """Return the line values that ``Get-Content`` supplies to the M304 script."""

    lines, had_final_newline = _split_normalised_lines(text)
    # A zero-byte file yields no content objects from Get-Content.  For normal
    # NC files this special case is irrelevant, but it avoids inventing a
    # blank line during a no-op M304 conversion.
    if not text and not had_final_newline:
        return []
    return lines


def _m06_text(text: str) -> str:
    """Remove every whole line whose contents contain ``M6`` (case-insensitive)."""

    newline = "\r\n" if "\r\n" in text else "\n" if "\n" in text else "\r\n"
    lines, had_final_newline = _split_normalised_lines(text)
    kept = [line for line in lines if not re.search(r"M6", line, re.IGNORECASE)]
    result = newline.join(kept)
    return result + newline if had_final_newline else result


def _m304_text(text: str) -> str:
    """Reproduce ``convert_nc_DP-A-88_Only.ps1`` line for line."""

    lines = _powershell_content_lines(text)

    has_m304 = any(line.strip().upper() == "M304" for line in lines)
    if not has_m304:
        for index, original in enumerate(lines):
            compact_upper = original.replace(" ", "").upper()
            # The first pattern is intentionally retained even though the
            # second one also accepts it; this mirrors the supplied script.
            if re.search(r"G43H\d+.*M8", compact_upper) or re.search(r"G43.*M8", compact_upper):
                lines[index + 1 : index + 1] = ["M304", "G05P10000"]
                break

    has_g05p0 = any(line.strip().upper() == "G05P0" for line in lines)
    if not has_g05p0:
        for index, original in enumerate(lines):
            # No trim is intentional.  The source shortcut only sees M5 when
            # it is literally at the end of its line.
            if re.search(r"M5$", original.upper()):
                lines[index + 1 : index + 1] = ["G05P0", "M300"]
                break

    # Set-Content writes line records with Windows line endings and ends the
    # file with a newline.  It writes an empty file when no line records exist.
    return "\r\n".join(lines) + ("\r\n" if lines else "")


def _copper_text(text: str) -> str:
    """Reproduce the copper shortcut, including its first-line-only check."""

    lines, _ = _split_normalised_lines(text)

    # The original script only checks the first line after ``%``.  It does not
    # scan the rest of the program, so an existing safety block later in the
    # file is deliberately duplicated here for compatibility.
    if lines[0].strip() == "%":
        if len(lines) < 2 or lines[1].strip() != _COPPER_INSERT_LINE:
            lines = [lines[0], _COPPER_INSERT_LINE, *lines[1:]]
    elif lines[0].strip() != _COPPER_INSERT_LINE:
        lines = [_COPPER_INSERT_LINE, *lines]

    for index in range(len(lines) - 1, -1, -1):
        if re.search(r"M30\s*$", lines[index], re.IGNORECASE):
            lines[index] = re.sub(r"M30\s*$", "M99", lines[index], count=1, flags=re.IGNORECASE)
            break

    # The shortcut always emits CRLF and always adds one final line ending.
    return "\r\n".join(lines) + "\r\n"


def transform_text(text: str, mode: TransformMode | str) -> str:
    """Transform NC source text without reading or writing a file.

    ``M06`` removes every line containing the literal sequence ``M6``;
    ``M304`` inserts the M304/G05 blocks if their exact marker lines are
    absent; and ``TG`` applies the copper-program conversion.
    """

    selected = _coerce_mode(mode)
    if selected is TransformMode.M06:
        return _m06_text(text)
    if selected is TransformMode.M304:
        return _m304_text(text)
    return _copper_text(text)


def _is_nc_path(path: Path) -> bool:
    return path.suffix.casefold() == ".nc"


def transform_file(
    path: str | Path,
    mode: TransformMode | str,
    source_text: str | None = None,
    *,
    encoding: str | None = None,
) -> Path:
    """Apply one shortcut to a disk file and return its resulting path.

    M06 and M304 overwrite the original file.  The copper (``TG``) shortcut
    writes a same-name *extensionless* target, then deletes the original
    ``.NC`` file, exactly like the existing copper PowerShell shortcut.

    M06 and TG silently leave a missing or non-``.NC`` path unchanged, matching
    the batch scripts' ``Skipped`` behaviour.  M304 intentionally accepts any
    filename, just like the supplied PowerShell script.  Passing ``source_text``
    lets a UI save its hand-edited text without re-reading the file first.
    """

    source = Path(path)
    selected = _coerce_mode(mode)
    if selected in {TransformMode.M06, TransformMode.COPPER} and (not source.is_file() or not _is_nc_path(source)):
        return source
    if source_text is None:
        source_text = decode_windows_default(source.read_bytes(), encoding=encoding)

    output_text = transform_text(source_text, selected)
    if selected is TransformMode.COPPER:
        target = source.with_suffix("")
        target.write_bytes(encode_windows_default(output_text, encoding=encoding))
        if target != source and source.is_file():
            source.unlink()
        return target

    source.write_bytes(encode_windows_default(output_text, encoding=encoding))
    return source
