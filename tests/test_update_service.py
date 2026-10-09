from urllib.request import Request

from cnc_program_sheet import update_service
from cnc_program_sheet.update_service import HTTP_USER_AGENT, UpdateInfo


def test_update_user_agent_is_ascii_safe() -> None:
    """GitHub requests must not include the Chinese display name in headers."""

    assert HTTP_USER_AGENT == "CNCProgramSheet/1.1.12"
    assert HTTP_USER_AGENT.isascii()
    assert Request("https://api.github.com", headers={"User-Agent": HTTP_USER_AGENT})


def test_download_uses_the_manifest_filename_for_checksum_matching(tmp_path, monkeypatch) -> None:
    """GitHub's uploaded name may differ from the name recorded in the manifest."""

    monkeypatch.setattr(update_service, "updates_dir", lambda: tmp_path)

    def fake_download(url: str, target):
        if url.endswith("SHA256SUMS.txt"):
            target.write_text("abc *CNCProgramSheet_1.1.6_Setup_x64.exe", encoding="utf-8")
        else:
            target.write_bytes(b"installer bytes")

    monkeypatch.setattr(update_service, "_download", fake_download)
    monkeypatch.setattr(update_service.hashlib, "sha256", lambda _: type("Digest", (), {"hexdigest": lambda self: "abc"})())

    update = UpdateInfo(
        version="1.1.6",
        installer_name="CNC._1.1.6_Setup_x64.exe",
        installer_url="https://example.test/CNC._1.1.6_Setup_x64.exe",
        checksum_url="https://example.test/SHA256SUMS.txt",
        release_url="https://example.test/release",
        notes="",
    )

    assert update_service.download_verified_installer(update).name == "installer.exe"
