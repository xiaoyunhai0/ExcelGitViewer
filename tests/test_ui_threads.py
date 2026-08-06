from __future__ import annotations

import time
from pathlib import Path

import pytest

try:
    from PySide6.QtCore import QSettings, Qt
    from PySide6.QtGui import QColor
    from PySide6.QtWidgets import QTableWidget

    from excel_git_viewer.models import CellContext
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


def test_workspace_panes_are_tables_with_two_axis_splitters(tmp_path: Path) -> None:
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path))
    _application, window = create_application()

    assert isinstance(window.commit_table, QTableWidget)
    assert isinstance(window.file_table, QTableWidget)
    assert window.main_splitter.orientation() == Qt.Orientation.Horizontal
    assert window.details_splitter.orientation() == Qt.Orientation.Vertical
    assert window.context_splitter.orientation() == Qt.Orientation.Horizontal
    assert window.main_splitter.count() == 3
    assert window.details_splitter.count() == 3
    assert window.context_splitter.count() == 2
    window.close()


def test_context_marks_the_changed_cell_and_its_headers(tmp_path: Path) -> None:
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path))
    _application, window = create_application()
    context = CellContext(
        start_row=3,
        start_column=2,
        values=((None, None, None), (None, "changed", None), (None, None, None)),
    )

    window._fill_context(
        window.old_context_table,
        context,
        "C4",
        background=QColor("#fecaca"),
        foreground=QColor("#991b1b"),
    )

    target = window.old_context_table.item(1, 1)
    assert target is not None
    assert target.background().color() == QColor("#fecaca")
    assert target.font().bold()
    assert target.toolTip() == "当前改动单元格 C4"
    assert window.old_context_table.horizontalHeaderItem(1).background().color() == QColor(
        "#fef3c7"
    )
    assert window.old_context_table.verticalHeaderItem(1).background().color() == QColor("#fef3c7")
    window.close()
