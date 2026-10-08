"""Frozen-app bootstrap.

Qt6Core.dll depends on ICU DLLs placed beside the extracted PyInstaller
payload. Add both payload locations before importing any PySide6 module.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path


_DLL_DIRECTORY_HANDLES: list[object] = []
if sys.platform == "win32" and hasattr(sys, "_MEIPASS"):
    bundle_root = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    for directory in (bundle_root, bundle_root / "PySide6"):
        if directory.is_dir():
            _DLL_DIRECTORY_HANDLES.append(os.add_dll_directory(str(directory)))

from cnc_program_sheet.main import _exception_hook, run, self_check


if __name__ == "__main__":
    sys.excepthook = _exception_hook
    raise SystemExit(self_check() if "--self-check" in sys.argv else run())
