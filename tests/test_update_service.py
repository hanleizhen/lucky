from urllib.request import Request

from cnc_program_sheet.update_service import HTTP_USER_AGENT


def test_update_user_agent_is_ascii_safe() -> None:
    """GitHub requests must not include the Chinese display name in headers."""

    assert HTTP_USER_AGENT == "CNCProgramSheet/1.1.2"
    assert HTTP_USER_AGENT.isascii()
    assert Request("https://api.github.com", headers={"User-Agent": HTTP_USER_AGENT})
