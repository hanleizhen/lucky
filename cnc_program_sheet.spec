# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

# PyInstaller evaluates a spec from its build work directory on some Windows
# versions.  Releases intentionally invoke it from the repository root, so
# use the current project directory instead of SPECPATH.
project = Path.cwd().resolve()
datas = [
    (str(project / "assets" / "CNC程序单.xlsx"), "assets"),
    (str(project / "assets" / "update_source.json"), "assets"),
]

analysis = Analysis(
    [str(project / "launcher.py")],
    pathex=[str(project)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(analysis.pure)
exe = EXE(
    pyz,
    analysis.scripts,
    [],
    exclude_binaries=True,
    name="CNC程序单自动生成工具",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
)
coll = COLLECT(
    exe,
    analysis.binaries,
    analysis.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="CNC程序单自动生成工具",
)
