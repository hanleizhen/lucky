from urllib.request import Request

from cnc_program_sheet import update_service
from cnc_program_sheet.update_service import HTTP_USER_AGENT, UpdateInfo


def test_update_user_agent_is_ascii_safe() -> None:
    """GitHub requests must not include the Chinese display name in headers."""

    assert HTTP_USER_AGENT == "CNCProgramSheet/1.1.14"
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


def test_update_context_adds_windows_trusted_roots_without_disabling_tls(monkeypatch) -> None:
    class FakeContext:
        def __init__(self) -> None:
            self.loaded: list[bytes] = []

        def load_verify_locations(self, *, cadata: bytes) -> None:
            self.loaded.append(cadata)

    context = FakeContext()
    monkeypatch.setattr(update_service.sys, "platform", "win32")
    monkeypatch.setattr(update_service.ssl, "create_default_context", lambda: context)
    monkeypatch.setattr(
        update_service.ssl,
        "enum_certificates",
        lambda store: [
            (f"{store}-root".encode(), "x509_asn", True),
            (b"server-only", "x509_asn", {update_service.TLS_SERVER_AUTH_OID}),
            (b"ignore-pkcs7", "pkcs_7_asn", True),
        ],
    )
    update_service._trusted_ssl_context.cache_clear()

    assert update_service._trusted_ssl_context() is context
    assert context.loaded == [b"ROOT-root", b"server-only", b"CA-root", b"server-only"]
    update_service._trusted_ssl_context.cache_clear()
