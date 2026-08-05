from __future__ import annotations

import time
from pathlib import Path

import pytest

try:
    from PySide6.QtCore import QSettings

    from excel_git_viewer.ui import create_application
except ImportError:
    pytest.skip("Qt system libraries are unavailable", allow_module_level=True)


def test_background_result_reaches_the_ui_and_releases_its_worker(tmp_path: Path) -> None:
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path))
    application, window = create_application()
    results: list[object] = []

    window._start_task(
        window._commit_tasks,
        lambda _token: "loaded",
        results.append,
        "loading",
    )

    deadline = time.monotonic() + 3
    while not results and time.monotonic() < deadline:
        application.processEvents()
        time.sleep(0.01)

    assert results == ["loaded"]
    assert window._workers == {}
    window.close()
