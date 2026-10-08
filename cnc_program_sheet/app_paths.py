"""Paths for code versus user-owned data.

The split is intentional: application upgrades only replace the installation
directory. Templates copied for the user, settings, logs and downloads stay in
LocalAppData and are never targets of an installer or updater.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
from hashlib import sha256
from pathlib import Path
from typing import Any

from .version import APP_NAME


def resource_dir() -> Path:
    """Directory of bundled read-only resources in source and PyInstaller builds."""

    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))


def resource_path(*parts: str) -> Path:
    return resource_dir().joinpath(*parts)


def user_data_dir() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    path = base / "CNCProgramSheet"
    path.mkdir(parents=True, exist_ok=True)
    return path


def logs_dir() -> Path:
    path = user_data_dir() / "logs"
    path.mkdir(exist_ok=True)
    return path


def updates_dir() -> Path:
    path = user_data_dir() / "updates"
    path.mkdir(exist_ok=True)
    return path


def user_images_dir() -> Path:
    """Return the user-owned image cache, preserved across app upgrades."""

    path = user_data_dir() / "images"
    path.mkdir(exist_ok=True)
    return path


def copy_user_image(source: str | Path) -> Path:
    """Keep an inserted image available even if its original file is moved.

    The cache is deliberately located in LocalAppData, not in the installed
    application directory, so an in-place upgrade cannot remove it.
    """

    source_path = Path(source)
    digest = sha256(source_path.read_bytes()).hexdigest()[:16]
    target = user_images_dir() / f"{digest}_{source_path.name}"
    if not target.exists():
        shutil.copy2(source_path, target)
    return target


def config_path() -> Path:
    return user_data_dir() / "settings.json"


def load_settings() -> dict[str, Any]:
    try:
        return json.loads(config_path().read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def save_settings(settings: dict[str, Any]) -> None:
    target = config_path()
    temporary = target.with_suffix(".json.tmp")
    content = json.dumps(settings, ensure_ascii=False, indent=2)
    temporary.write_text(content, encoding="utf-8")
    try:
        temporary.replace(target)
    except OSError:
        # Windows AppContainer redirection can reject os.replace even when both
        # paths display as LocalAppData. Fall back to a normal write there; in
        # a regular installed desktop app the atomic replace path is used.
        target.write_text(content, encoding="utf-8")
        temporary.unlink(missing_ok=True)


def default_user_template() -> Path:
    """Create the initial user template once, never replace it during upgrades."""

    target_dir = user_data_dir() / "templates"
    target_dir.mkdir(exist_ok=True)
    target = target_dir / "CNC程序单.xlsx"
    if not target.exists():
        source = resource_path("assets", "CNC程序单.xlsx")
        if not source.exists():
            raise FileNotFoundError(f"找不到随程序附带的模板：{source}")
        shutil.copy2(source, target)
    return target


def load_embedded_update_repository() -> str:
    """Release CI writes this file with the real owner/repository name."""

    try:
        content = json.loads(resource_path("assets", "update_source.json").read_text(encoding="utf-8"))
        return str(content.get("repository", "")).strip()
    except (OSError, json.JSONDecodeError):
        return ""


def log_exception(message: str) -> Path:
    import datetime as _datetime

    target = logs_dir() / "error.log"
    timestamp = _datetime.datetime.now().astimezone().isoformat(timespec="seconds")
    with target.open("a", encoding="utf-8") as handle:
        handle.write(f"\n[{timestamp}]\n{message}\n")
    return target
