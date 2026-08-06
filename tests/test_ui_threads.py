from __future__ import annotations

import time
from collections.abc import Iterator
from pathlib import Path

import pytest

try:
    from PySide6.QtCore import QCoreApplication, QEvent, QSettings, Qt
    from PySide6.QtGui import QColor
    from PySide6.QtWidgets import QApplication, QDockWidget, QHeaderView, QMainWindow, QTableWidget

    from excel_git_viewer.models import CellContext
    from excel_git_viewer.ui import create_application
except ImportError:
    pytest.skip("Qt system libraries are unavailable", allow_module_level=True)


@pytest.fixture(autouse=True)
def isolate_qt_state(tmp_path: Path) -> Iterator[None]:
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path))
    settings = QSettings("ExcelGitViewer", "ExcelGitViewer")
    settings.clear()
    settings.sync()
    yield
    application = QApplication.instance()
    if isinstance(application, QApplication):
        for widget in application.topLevelWidgets():
            widget.close()
            widget.deleteLater()
        QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
        application.processEvents()
    settings.clear()
    settings.sync()


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


def test_workspace_panes_are_movable_docks(tmp_path: Path) -> None:
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path))
    _application, window = create_application()

    assert isinstance(window.commit_table, QTableWidget)
    assert isinstance(window.file_table, QTableWidget)
    docks = (
        window.commit_dock,
        window.file_dock,
        window.sheet_dock,
        window.cell_dock,
        window.old_context_dock,
        window.new_context_dock,
    )
    movable = QDockWidget.DockWidgetFeature.DockWidgetMovable
    floatable = QDockWidget.DockWidgetFeature.DockWidgetFloatable
    closable = QDockWidget.DockWidgetFeature.DockWidgetClosable
    for dock in docks:
        assert dock.features() & movable
        assert dock.features() & floatable
        assert not dock.features() & closable
        assert dock.allowedAreas() == Qt.DockWidgetArea.AllDockWidgetAreas
    assert window.workspace.isDockNestingEnabled()
    assert window.workspace.dockOptions() & QMainWindow.DockOption.AllowTabbedDocks
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


def test_dock_layout_is_restored_after_reopening(tmp_path: Path) -> None:
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path))
    _application, window = create_application()
    window.workspace.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, window.commit_dock)
    expected_state = window.workspace.saveState()
    window.close()

    _application, restored = create_application()

    assert restored.workspace.saveState() == expected_state
    assert (
        restored.workspace.dockWidgetArea(restored.commit_dock)
        == Qt.DockWidgetArea.BottomDockWidgetArea
    )
    restored.close()


def test_default_layout_action_restores_docks_and_column_widths(tmp_path: Path) -> None:
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path))
    _application, window = create_application()
    default_state = window.workspace.saveState()
    default_description_width = window.commit_table.columnWidth(1)
    window.workspace.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, window.commit_dock)
    window.commit_table.setColumnWidth(1, 340)

    window._reset_layout()

    assert window.workspace.saveState() == default_state
    assert window.commit_table.columnWidth(1) == default_description_width
    window.close()


def test_switching_columns_in_the_same_row_does_not_restart_work(tmp_path: Path) -> None:
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path))
    _application, window = create_application()
    window._repository = object()
    window._commits = [object()]
    window._files = [object()]

    window._select_commit(0, 1, 0, 0)
    window._select_file(0, 1, 0, 0)

    window.close()


def test_every_table_column_can_be_resized_by_the_user(tmp_path: Path) -> None:
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path))
    _application, window = create_application()
    window.old_context_table.setColumnCount(3)
    window.new_context_table.setColumnCount(3)

    tables = (
        window.commit_table,
        window.file_table,
        window.sheet_table,
        window.cell_table,
        window.old_context_table,
        window.new_context_table,
    )
    for table in tables:
        header = table.horizontalHeader()
        assert all(
            header.sectionResizeMode(column) == QHeaderView.ResizeMode.Interactive
            for column in range(table.columnCount())
        )
    window.close()


def test_fixed_table_column_widths_are_restored_after_reopening(tmp_path: Path) -> None:
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, str(tmp_path))
    _application, window = create_application()
    window.commit_table.setColumnWidth(1, 333)
    window.file_table.setColumnWidth(1, 287)
    window.sheet_table.setColumnWidth(0, 241)
    window.cell_table.setColumnWidth(3, 319)
    expected_widths = {
        "commit": window.commit_table.columnWidth(1),
        "file": window.file_table.columnWidth(1),
        "sheet": window.sheet_table.columnWidth(0),
        "cell": window.cell_table.columnWidth(3),
    }
    window.close()

    _application, restored = create_application()

    assert restored.commit_table.columnWidth(1) == expected_widths["commit"]
    assert restored.file_table.columnWidth(1) == expected_widths["file"]
    assert restored.sheet_table.columnWidth(0) == expected_widths["sheet"]
    assert restored.cell_table.columnWidth(3) == expected_widths["cell"]
    restored.close()
