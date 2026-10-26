import os
import uuid

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtTest import QSignalSpy
from PySide6.QtWidgets import QApplication

from cnc_program_sheet.single_instance import SingleInstanceCoordinator


def test_second_instance_forwards_request_to_primary(tmp_path) -> None:
    app = QApplication.instance() or QApplication([])
    name = "cnc-program-sheet-test-" + uuid.uuid4().hex
    inbox = tmp_path / "inbox"
    primary = SingleInstanceCoordinator(name, inbox_dir=inbox)
    secondary = SingleInstanceCoordinator(name, inbox_dir=inbox)
    spy = QSignalSpy(primary.request_received)
    try:
        assert primary.start_or_forward({"import_paths": []})
        request = {"import_paths": [r"C:\NC\A.NC"]}

        assert not secondary.start_or_forward(request)
        assert spy.wait(1500)
        assert spy.count() == 1
        assert spy.at(0)[0] == request
        assert list(inbox.glob("request_*.json")) == []
    finally:
        secondary.close()
        primary.close()
    assert app is not None


def test_primary_restores_existing_app_and_queues_import(tmp_path, monkeypatch) -> None:
    from cnc_program_sheet.main import MainWindow

    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    restored: list[bool] = []
    monkeypatch.setattr(window, "_restore_from_tray", lambda: restored.append(True))
    monkeypatch.setattr(window, "_process_next_external_launch", lambda: None)
    request = {"import_paths": [str(tmp_path / "part.NC")]}

    window.queue_external_launch(request)

    assert restored == [True]
    assert window._external_launch_queue == [request]
    monkeypatch.undo()
    window.tray_icon.hide()
    window.close()
    assert app is not None
