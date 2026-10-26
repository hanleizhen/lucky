"""Per-user single-instance lock and local request queue."""

from __future__ import annotations

import ctypes
import getpass
import hashlib
import json
import os
import sys
import tempfile
import time
import uuid
from pathlib import Path
from typing import Any

from PySide6.QtCore import QObject, QLockFile, QTimer, Signal


def _session_id() -> str:
    if os.name != "nt":
        return "default"
    value = ctypes.c_ulong()
    try:
        if ctypes.windll.kernel32.ProcessIdToSessionId(os.getpid(), ctypes.byref(value)):
            return str(value.value)
    except (AttributeError, OSError):
        pass
    return "default"


def default_server_name() -> str:
    """Stable lock name unique to user, session, and app installation path."""

    install_root = str(Path(sys.executable).resolve().parent).casefold()
    identity = f"{getpass.getuser().casefold()}:{_session_id()}:{install_root}"
    digest = hashlib.sha256(identity.encode("utf-8")).hexdigest()[:24]
    return f"cnc-program-sheet-{digest}"


def default_inbox_dir() -> Path:
    local_app_data = Path(
        os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local"
    )
    return local_app_data / "CNCProgramSheet" / "instance_requests"


class SingleInstanceCoordinator(QObject):
    """Own the app lock or enqueue a launch request for its existing process."""

    request_received = Signal(object)

    def __init__(
        self,
        server_name: str | None = None,
        parent: QObject | None = None,
        inbox_dir: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.server_name = server_name or default_server_name()
        self.inbox_dir = Path(inbox_dir) if inbox_dir is not None else default_inbox_dir()
        self._mutex_handle: int | None = None
        self._lock_file: QLockFile | None = None
        self._timer = QTimer(self)
        self._timer.setInterval(150)
        self._timer.timeout.connect(self._drain_inbox)

    def start_or_forward(self, request: dict[str, Any]) -> bool:
        """Return True for the primary app; queue and return False otherwise."""

        if self._acquire_primary_lock():
            self.inbox_dir.mkdir(parents=True, exist_ok=True)
            self._timer.start()
            return True
        self._enqueue(request)
        return False

    def _acquire_primary_lock(self) -> bool:
        if os.name == "nt":
            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            create_mutex = kernel32.CreateMutexW
            create_mutex.argtypes = (ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p)
            create_mutex.restype = ctypes.c_void_p
            kernel32.SetLastError(0)
            handle = create_mutex(None, False, "Local\\" + self.server_name)
            if not handle:
                raise ctypes.WinError(ctypes.get_last_error())
            if ctypes.get_last_error() == 183:  # ERROR_ALREADY_EXISTS
                kernel32.CloseHandle(handle)
                return False
            self._mutex_handle = int(handle)
            return True

        lock_path = Path(tempfile.gettempdir()) / (self.server_name + ".lock")
        self._lock_file = QLockFile(str(lock_path))
        if self._lock_file.tryLock(0):
            return True
        self._lock_file = None
        return False

    def _enqueue(self, request: dict[str, Any]) -> None:
        self.inbox_dir.mkdir(parents=True, exist_ok=True)
        envelope = {"created_at": time.time(), "request": request}
        filename = f"request_{os.getpid()}_{uuid.uuid4().hex}.json"
        target = self.inbox_dir / filename
        temporary = self.inbox_dir / (filename + ".tmp")
        temporary.write_text(json.dumps(envelope, ensure_ascii=False), encoding="utf-8")
        os.replace(temporary, target)

    def _drain_inbox(self) -> None:
        if not self.inbox_dir.is_dir():
            return
        now = time.time()
        for path in sorted(self.inbox_dir.glob("request_*.json"), key=lambda item: item.stat().st_mtime):
            try:
                envelope = json.loads(path.read_text(encoding="utf-8"))
                created_at = float(envelope.get("created_at", 0))
                request = envelope.get("request")
                if now - created_at <= 300 and isinstance(request, dict):
                    self.request_received.emit(request)
            except (OSError, ValueError, TypeError, json.JSONDecodeError):
                pass
            finally:
                try:
                    path.unlink(missing_ok=True)
                except OSError:
                    pass

    def close(self) -> None:
        self._timer.stop()
        if self._mutex_handle is not None:
            ctypes.WinDLL("kernel32", use_last_error=True).CloseHandle(self._mutex_handle)
            self._mutex_handle = None
        if self._lock_file is not None:
            self._lock_file.unlock()
            self._lock_file = None
