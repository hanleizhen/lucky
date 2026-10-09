"""GitHub Release update checks and a data-safe installer handoff."""

from __future__ import annotations

import hashlib
import json
import ssl
import subprocess
import sys
import urllib.request
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any
from urllib.parse import quote

from .app_paths import updates_dir
from .version import APP_NAME, __version__


# HTTP header values must be Latin-1/ASCII encodable.  APP_NAME is deliberately
# Chinese for the Windows UI, so it must never be sent verbatim as User-Agent.
HTTP_USER_AGENT = f"CNCProgramSheet/{__version__}"
TLS_SERVER_AUTH_OID = "1.3.6.1.5.5.7.3.1"
# A fixed-name asset can be fetched through GitHub's public Release download
# endpoint.  Unlike api.github.com it does not consume the shared anonymous
# API request quota that office networks can exhaust very quickly.
STABLE_UPDATE_MANIFEST = "CNCProgramSheet_update.json"


class UpdateError(RuntimeError):
    pass


@lru_cache(maxsize=1)
def _trusted_ssl_context() -> ssl.SSLContext:
    """Keep TLS verification while also trusting Windows-installed roots.

    A PyInstaller Python runtime can use an OpenSSL CA bundle that does not
    contain a root certificate installed by a corporate proxy or a security
    product's HTTPS scanner. On Windows, add the already trusted ROOT and CA
    store certificates to the normal secure context. This does *not* disable
    hostname or certificate validation; it only aligns the bundled runtime
    with the Windows trust decision on the user's computer.
    """

    context = ssl.create_default_context()
    if sys.platform != "win32" or not hasattr(ssl, "enum_certificates"):
        return context
    for store_name in ("ROOT", "CA"):
        try:
            certificates = ssl.enum_certificates(store_name)
        except OSError:
            continue
        for certificate, encoding, trust in certificates:
            if encoding != "x509_asn":
                continue
            if trust is not True and TLS_SERVER_AUTH_OID not in trust:
                continue
            try:
                # DER bytes are explicitly supported as ``cadata``. Preserve
                # the default context too, so normal public GitHub roots work.
                context.load_verify_locations(cadata=certificate)
            except ssl.SSLError:
                # One malformed local-store entry must not prevent the update
                # check from using the rest of the trusted certificate store.
                continue
    return context


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
        with urllib.request.urlopen(request, timeout=12, context=_trusted_ssl_context()) as response:
            return json.loads(response.read().decode("utf-8"))
    except Exception as exc:
        raise UpdateError(f"无法连接 GitHub Release：{exc}") from exc


def _manifest_update(repository: str, manifest: dict[str, Any]) -> UpdateInfo:
    """Build a safe update record from the fixed public Release manifest.

    The manifest carries only the version and installer name.  Every download
    URL is then constructed from the repository entered by the operator, so a
    malformed manifest cannot redirect the updater to an unrelated host.
    """

    version = str(manifest.get("version", "")).strip().lstrip("vV")
    _version_tuple(version)
    installer_name = str(manifest.get("installer_name", "")).strip()
    if not installer_name or Path(installer_name).name != installer_name or not installer_name.lower().endswith("_setup_x64.exe"):
        raise UpdateError("公开更新清单中的安装包文件名无效，已拒绝升级。")
    release_base = f"https://github.com/{repository}/releases/download/v{quote(version)}"
    return UpdateInfo(
        version=version,
        installer_name=installer_name,
        installer_url=f"{release_base}/{quote(installer_name)}",
        checksum_url=f"{release_base}/SHA256SUMS.txt",
        release_url=f"https://github.com/{repository}/releases/tag/v{quote(version)}",
        notes=str(manifest.get("notes", "")),
    )


def _api_latest(repository: str) -> UpdateInfo:
    """Return latest metadata through the API for old releases without a manifest."""

    release = _json_request(f"https://api.github.com/repos/{repository}/releases/latest")
    version = str(release.get("tag_name", "")).lstrip("vV")
    _version_tuple(version)
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


def check_latest(repository: str) -> UpdateInfo | None:
    repository = repository.strip().strip("/")
    if not repository or repository.count("/") != 1:
        raise UpdateError("请先在“设置更新源”中填写 GitHub 仓库，例如 owner/CNCProgramSheet。")
    manifest_url = f"https://github.com/{repository}/releases/latest/download/{STABLE_UPDATE_MANIFEST}"
    try:
        update = _manifest_update(repository, _json_request(manifest_url))
    except UpdateError as manifest_error:
        # v1.1.14 and earlier did not publish the fixed manifest.  Keep the
        # API fallback solely for those historical releases; all later normal
        # checks use the public asset path above and avoid rate limits.
        try:
            update = _api_latest(repository)
        except UpdateError as api_error:
            if "rate limit exceeded" in str(api_error).lower():
                raise UpdateError(
                    "GitHub 当前限制了此网络的匿名更新查询。请稍后再试，或直接从发布页下载安装包。"
                ) from api_error
            raise UpdateError(f"公开更新清单和 GitHub Release 均无法访问：{api_error}") from api_error
        # A successful fallback is expected while upgrading old release
        # history; do not expose a harmless missing-manifest error to users.
        _ = manifest_error
    return update if is_newer(update.version) else None


def _download(url: str, target: Path) -> None:
    request = urllib.request.Request(url, headers={"User-Agent": HTTP_USER_AGENT})
    try:
        with urllib.request.urlopen(request, timeout=30, context=_trusted_ssl_context()) as response, target.open("wb") as handle:
            while chunk := response.read(1024 * 1024):
                handle.write(chunk)
    except Exception as exc:
        target.unlink(missing_ok=True)
        raise UpdateError(f"下载更新失败：{exc}") from exc


def _installer_name_from_manifest(checksum_file: Path) -> str:
    """Return the single safe setup filename declared by SHA256SUMS.txt."""

    names: list[str] = []
    for line in checksum_file.read_text(encoding="utf-8", errors="replace").splitlines():
        parts = line.split()
        if len(parts) < 2:
            continue
        name = parts[-1].lstrip("*")
        if name.lower().endswith("_setup_x64.exe") and Path(name).name == name:
            names.append(name)
    if len(names) != 1:
        raise UpdateError("Release 的 SHA256SUMS.txt 缺少唯一的安装包文件名，已拒绝升级。")
    return names[0]


def download_verified_installer(update: UpdateInfo) -> Path:
    """Fetch installer and require an SHA-256 listed in the same published Release."""

    destination_dir = updates_dir() / f"v{update.version}"
    destination_dir.mkdir(parents=True, exist_ok=True)
    checksum_file = destination_dir / "SHA256SUMS.txt"
    _download(update.checksum_url, checksum_file)
    # The uploaded GitHub asset can have a transliterated filename while the
    # signed manifest uses the Windows installer name.  Follow the manifest:
    # it is the authority used for checksum verification.
    manifest_installer_name = _installer_name_from_manifest(checksum_file)
    # cmd.exe is used to launch the installer only after the GUI exits.  Keep
    # that handoff path ASCII-only: the legacy Windows console code page can
    # corrupt Chinese characters and then report a false “file not found”.
    installer = destination_dir / "installer.exe"
    _download(update.installer_url, installer)
    expected: str | None = None
    for line in checksum_file.read_text(encoding="utf-8", errors="replace").splitlines():
        parts = line.split()
        if len(parts) >= 2 and parts[-1].lstrip("*") == manifest_installer_name:
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
