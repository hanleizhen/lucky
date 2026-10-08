from cnc_program_sheet.main import _exception_hook, run, self_check

if __name__ == "__main__":
    import sys

    debug_marker = __import__("os").environ.get("CNC_PROGRAM_SHEET_SELF_CHECK_MARKER")
    if debug_marker:
        from pathlib import Path

        Path(debug_marker).write_text(f"argv={sys.argv!r}\n", encoding="utf-8")

    sys.excepthook = _exception_hook
    raise SystemExit(self_check() if "--self-check" in sys.argv else run())
