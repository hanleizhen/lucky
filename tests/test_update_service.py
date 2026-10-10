import ssl
from types import SimpleNamespace
from urllib.request import Request

from cnc_program_sheet import update_service
from cnc_program_sheet.update_service import HTTP_USER_AGENT, UpdateInfo


def test_update_user_agent_is_ascii_safe() -> None:
    """GitHub requests must not include the Chinese display name in headers."""

    assert HTTP_USER_AGENT == "CNCProgramSheet/1.1.18"
    assert HTTP_USER_AGENT.isascii()
    assert Request("https://api.github.com", headers={"User-Agent": HTTP_USER_AGENT})


def test_latest_prefers_public_release_manifest_without_calling_api(monkeypatch) -> None:
    calls: list[str] = []

    def fake_json_request(url: str):
        calls.append(url)
        return {"version": "1.1.19", "installer_name": "CNCProgramSheet_1.1.19_Setup_x64.exe"}

    monkeypatch.setattr(update_service, "_json_request", fake_json_request)

    update = update_service.check_latest("hanleizhen/lucky")

    assert update is not None
    assert update.version == "1.1.19"
    assert update.installer_url == "https://github.com/hanleizhen/lucky/releases/download/v1.1.19/CNCProgramSheet_1.1.19_Setup_x64.exe"
    assert update.checksum_url == "https://github.com/hanleizhen/lucky/releases/download/v1.1.19/SHA256SUMS.txt"
    assert calls == ["https://github.com/hanleizhen/lucky/releases/latest/download/CNCProgramSheet_update.json"]


def test_latest_falls_back_to_api_only_for_old_release_without_manifest(monkeypatch) -> None:
    calls: list[str] = []

    def fake_json_request(url: str):
        calls.append(url)
        if url.endswith(update_service.STABLE_UPDATE_MANIFEST):
            raise update_service.UpdateError("HTTP Error 404: Not Found")
        return {
            "tag_name": "v1.1.19",
            "assets": [
                {"name": "CNCProgramSheet_1.1.19_Setup_x64.exe", "browser_download_url": "https://example.test/setup"},
                {"name": "SHA256SUMS.txt", "browser_download_url": "https://example.test/sha"},
            ],
            "html_url": "https://example.test/release",
            "body": "notes",
        }

    monkeypatch.setattr(update_service, "_json_request", fake_json_request)

    update = update_service.check_latest("hanleizhen/lucky")

    assert update is not None
    assert update.installer_url == "https://example.test/setup"
    assert calls[1] == "https://api.github.com/repos/hanleizhen/lucky/releases/latest"


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


def test_update_context_prefers_strict_windows_native_truststore(monkeypatch) -> None:
    """The updater must prefer native Windows trust without weakening TLS."""

    class FakeContext:
        def __init__(self) -> None:
            self.verify_mode = ssl.CERT_NONE
            self.check_hostname = False

    calls: list[int] = []
    context = FakeContext()

    def fake_native_context(protocol: int) -> FakeContext:
        calls.append(protocol)
        return context

    monkeypatch.setattr(update_service.sys, "platform", "win32")
    monkeypatch.setattr(update_service, "truststore", SimpleNamespace(SSLContext=fake_native_context))
    monkeypatch.setattr(
        update_service.ssl,
        "create_default_context",
        lambda: (_ for _ in ()).throw(AssertionError("native truststore should be preferred")),
    )
    update_service._trusted_ssl_context.cache_clear()
    try:
        assert update_service._trusted_ssl_context() is context
        assert calls == [ssl.PROTOCOL_TLS_CLIENT]
        assert context.verify_mode == ssl.CERT_REQUIRED
        assert context.check_hostname is True
    finally:
        update_service._trusted_ssl_context.cache_clear()


def test_update_context_falls_back_to_windows_store_when_truststore_fails(monkeypatch) -> None:
    """A native truststore initialisation failure keeps the old secure fallback."""

    class FakeContext:
        def __init__(self) -> None:
            self.loaded: list[bytes] = []
            self.verify_mode = ssl.CERT_NONE
            self.check_hostname = False

        def load_verify_locations(self, *, cadata: bytes) -> None:
            self.loaded.append(cadata)

    context = FakeContext()

    def broken_native_context(_protocol: int) -> FakeContext:
        raise RuntimeError("native store unavailable")

    monkeypatch.setattr(update_service.sys, "platform", "win32")
    monkeypatch.setattr(update_service, "truststore", SimpleNamespace(SSLContext=broken_native_context))
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
    try:
        assert update_service._trusted_ssl_context() is context
        assert context.loaded == [b"ROOT-root", b"server-only", b"CA-root", b"server-only"]
        assert context.verify_mode == ssl.CERT_REQUIRED
        assert context.check_hostname is True
    finally:
        update_service._trusted_ssl_context.cache_clear()
