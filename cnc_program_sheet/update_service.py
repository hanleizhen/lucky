"""GitHub Release update checks and a data-safe installer handoff."""

from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .app_paths import updates_dir
from .version import APP_NAME, __version__


# HTTP header values must be Latin-1/ASCII encodable.  APP_NAME is deliberately
# Chinese for the Windows UI, so it must never be sent verbatim as User-Agent.
HTTP_USER_AGENT = f"CNCProgramSheet/{__version__}"


class UpdateError(RuntimeError):
    pass


@dataclass(frozen=True, slots=True)
class UpdateInfo:
    version: str
    installer_name: str
    installer_url: str
    checksum_url: str
    release_url: str
    notes: str


def _version_tuple(value: str) -> tuple[int, ...]:
    value = value.strip().lstrip("vV")
    if not value or any(not piece.isdigit() for piece in value.split(".")):
        raise UpdateError(f"Release 版本号不符合 x.y.z 格式：{value!r}")
    return tuple(int(piece) for piece in value.split("."))


def is_newer(candidate: str, current: str = __version__) -> bool:
    newest, installed = _version_tuple(candidate), _version_tuple(current)
    length = max(len(newest), len(installed))
    return newest + (0,) * (length - len(newest)) > installed + (0,) * (length - len(installed))


def _json_request(url: str) -> dict[str, Any]:
    request = urllib.request.Request(
        url,
        headers={"Accept": "application/vnd.github+json", "User-Agent": HTTP_USER_AGENT},
    )
    try:
        with urllib.request.urlopen(request, timeout=12) as response:
            return json.loads(response.read().decode("utf-8"))
    except Exception as exc:
        raise UpdateError(f"无法连接 GitHub Release：{exc}") from exc


def check_latest(repository: str) -> UpdateInfo | None:
    repository = repository.strip().strip("/")
    if not repository or repository.count("/") != 1:
        raise UpdateError("请先在“设置更新源”中填写 GitHub 仓库，例如 owner/CNCProgramSheet。")
    release = _json_request(f"https://api.github.com/repos/{repository}/releases/latest")
    version = str(release.get("tag_name", "")).lstrip("vV")
    if not is_newer(version):
        return None
    assets = release.get("assets") or []
    installer = next(
        (asset for asset in assets if str(asset.get("name", "")).lower().endswith("_setup_x64.exe")),
        None,
    )
    checksum = next(
        (asset for asset in assets if str(asset.get("name", "")).upper() == "SHA256SUMS.TXT"),
        None,
    )
    if not installer or not checksum:
        raise UpdateError("最新 Release 缺少安装包或 SHA256SUMS.txt，已拒绝升级。")
    installer_name = str(installer.get("name", "")).strip()
    if not installer_name or Path(installer_name).name != installer_name:
        raise UpdateError("最新 Release 的安装包文件名无效，已拒绝升级。")
    return UpdateInfo(
        version=version,
        installer_name=installer_name,
        installer_url=str(installer["browser_download_url"]),
        checksum_url=str(checksum["browser_download_url"]),
        release_url=str(release.get("html_url", "")),
        notes=str(release.get("body", "")),
    )


def _download(url: str, target: Path) -> None:
    request = urllib.request.Request(url, headers={"User-Agent": HTTP_USER_AGENT})
    try:
        with urllib.request.urlopen(request, timeout=30) as response, target.open("wb") as handle:
            while chunk := response.read(1024 * 1024):
                handle.write(chunk)
    except Exception as exc:
        target.unlink(missing_ok=True)
        raise UpdateError(f"下载更新失败：{exc}") from exc


def download_verified_installer(update: UpdateInfo) -> Path:
    """Fetch installer and require an SHA-256 listed in the same published Release."""

    destination_dir = updates_dir() / f"v{update.version}"
    destination_dir.mkdir(parents=True, exist_ok=True)
    checksum_file = destination_dir / "SHA256SUMS.txt"
    # Preserve the filename published with the Release.  Inno Setup may
    # transliterate a Unicode display name, so constructing a Chinese name
    # locally can never match the filename recorded in SHA256SUMS.txt.
    installer = destination_dir / update.installer_name
    _download(update.checksum_url, checksum_file)
    _download(update.installer_url, installer)
    expected: str | None = None
    for line in checksum_file.read_text(encoding="utf-8", errors="replace").splitlines():
        parts = line.split()
        if len(parts) >= 2 and parts[-1].lstrip("*") == installer.name:
            expected = parts[0].lower()
            break
    actual = hashlib.sha256(installer.read_bytes()).hexdigest().lower()
    if not expected or actual != expected:
        installer.unlink(missing_ok=True)
        raise UpdateError("更新安装包的 SHA-256 校验失败，已取消升级。")
    return installer


def start_installer_after_exit(installer: Path) -> None:
    """Start Inno Setup after this GUI has released its executable file.

    The installer updates only its application directory. Templates, settings,
    logs and downloads are under LocalAppData/CNCProgramSheet and are outside
    the installer's [Files] scope.
    """

    script = updates_dir() / "run_update.cmd"
    script.write_text(
        "@echo off\r\n"
        "timeout /t 2 /nobreak >nul\r\n"
        f"start \"\" \"{installer}\" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /CLOSEAPPLICATIONS\r\n"
        "del \"%~f0\"\r\n",
        encoding="utf-8",
    )
    creationflags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0) | getattr(subprocess, "DETACHED_PROCESS", 0)
    subprocess.Popen(["cmd.exe", "/c", str(script)], creationflags=creationflags, close_fds=True)


def open_release_page(update: UpdateInfo) -> None:
    import webbrowser

    webbrowser.open(update.release_url)
