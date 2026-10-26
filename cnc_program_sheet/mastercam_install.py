"""Install the Mastercam NET-Hook button into supported user profiles."""

from __future__ import annotations

import os
import re
import shutil
import sys
from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path

from PySide6.QtCore import QStandardPaths

from .app_paths import resource_path, user_data_dir


SUPPORTED_MASTERCAM_YEARS = tuple(range(2021, 2028))
ADDIN_FILES = (
    "GenerateProgramSheet.dll",
    "GenerateProgramSheet.ft",
    "GenerateProgramSheet.csx",
)


@dataclass(frozen=True)
class MastercamInstallResult:
    installed_years: tuple[int, ...]
    destination_dirs: tuple[Path, ...]


class MastercamInstallError(RuntimeError):
    """Raised when no compatible Mastercam installation can be updated."""


def _standard_program_roots() -> tuple[Path, ...]:
    candidates = (
        os.environ.get("ProgramFiles"),
        os.environ.get("ProgramW6432"),
        os.environ.get("ProgramFiles(x86)"),
    )
    unique: list[Path] = []
    for value in candidates:
        if value:
            path = Path(value)
            if path not in unique:
                unique.append(path)
    return tuple(unique)


def discover_mastercam_addin_targets(
    documents_dir: Path | None = None,
    program_roots: tuple[Path, ...] | None = None,
) -> list[tuple[int, Path]]:
    """Find standard Mastercam 2021–2027 installs and their user Add-Ins dirs.

    Detection is deliberately based on both the installed executable and the
    matching ``My Mastercam YYYY`` profile. This avoids creating an Add-Ins
    tree for a stale profile left behind by an uninstalled Mastercam version.
    Both values are injectable so detection stays testable on non-Windows CI.
    """

    documents_location = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
    docs = documents_dir or (Path(documents_location) if documents_location else Path.home() / "Documents")
    roots = program_roots if program_roots is not None else _standard_program_roots()
    found: dict[int, Path] = {}
    for root in roots:
        if not root.is_dir():
            continue
        for installation in root.glob("Mastercam 20??"):
            match = re.fullmatch(r"Mastercam (20\d{2})", installation.name, re.IGNORECASE)
            if not match:
                continue
            year = int(match.group(1))
            if year not in SUPPORTED_MASTERCAM_YEARS:
                continue
            if not (installation / "Mastercam.exe").is_file():
                continue
            profile = docs / f"My Mastercam {year}" / "Mastercam"
            if not profile.is_dir():
                continue
            found[year] = profile / "Add-Ins"
    return sorted(found.items())


def _file_digest(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _running_application_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()
    raise MastercamInstallError(
        "Mastercam 自动安装只能从已打包的 CNC 智能程序单运行；当前是开发源码启动。"
    )


def install_mastercam_addin(
    *,
    documents_dir: Path | None = None,
    program_roots: tuple[Path, ...] | None = None,
    application_exe: Path | None = None,
    user_data: Path | None = None,
) -> MastercamInstallResult:
    """Install the versioned button files and remember this app's executable.

    Existing files with the same names are backed up once before replacement;
    unrelated Mastercam files are never touched.
    """

    targets = discover_mastercam_addin_targets(documents_dir, program_roots)
    if not targets:
        raise MastercamInstallError(
            "未检测到已安装的 Mastercam 2021–2027，或找不到对应的 My Mastercam 用户配置目录。"
        )

    source_dir = resource_path("assets", "mastercam_addin")
    sources = [source_dir / name for name in ADDIN_FILES]
    missing = [path.name for path in sources if not path.is_file()]
    if missing:
        raise MastercamInstallError("安装包缺少 Mastercam 加载项文件：" + ", ".join(missing))

    app_exe = (application_exe or _running_application_path()).resolve()
    if not app_exe.is_file():
        raise MastercamInstallError("找不到当前 CNC 智能程序单程序：" + str(app_exe))

    installed_years: list[int] = []
    destination_dirs: list[Path] = []
    for year, destination in targets:
        destination.mkdir(parents=True, exist_ok=True)
        for source in sources:
            target = destination / source.name
            if target.is_file() and _file_digest(target) == _file_digest(source):
                continue
            if target.is_file():
                backup = target.with_name(target.name + ".pre-cncprogramsheet.bak")
                if not backup.exists():
                    shutil.copy2(target, backup)
            shutil.copy2(source, target)
            if _file_digest(target) != _file_digest(source):
                raise MastercamInstallError(f"加载项文件校验失败：{target}")
        installed_years.append(year)
        destination_dirs.append(destination)

    data_dir = user_data or user_data_dir()
    data_dir.mkdir(parents=True, exist_ok=True)
    (data_dir / "mastercam_bridge_path.txt").write_text(str(app_exe), encoding="utf-8")
    return MastercamInstallResult(tuple(installed_years), tuple(destination_dirs))
