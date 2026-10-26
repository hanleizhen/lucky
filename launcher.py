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
    bundle_dll_dirs = (bundle_root, bundle_root / "PySide6", bundle_root / "shiboken6")
    for directory in bundle_dll_dirs:
        if directory.is_dir():
            _DLL_DIRECTORY_HANDLES.append(os.add_dll_directory(str(directory)))

    # Mastercam and other CAD hosts prepend their own native DLL folders to
    # PATH. A child GUI process inherits that value; a same-named Qt/ICU DLL
    # from the host can then be selected while importing PySide6.QtCore.
    # Frozen builds are self-contained, so retain only their bundled DLL paths
    # and the Windows system locations before importing any Qt modules.
    system_root = Path(os.environ.get("SystemRoot", r"C:\Windows"))
    safe_dll_dirs = [str(path) for path in bundle_dll_dirs if path.is_dir()]
    safe_dll_dirs.extend((str(system_root / "System32"), str(system_root)))
    os.environ["PATH"] = os.pathsep.join(safe_dll_dirs)

if __name__ == "__main__":
    from cnc_program_sheet.main import _exception_hook, run, self_check

    sys.excepthook = _exception_hook
    raise SystemExit(self_check() if "--self-check" in sys.argv else run())
