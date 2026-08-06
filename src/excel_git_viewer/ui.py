from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import cast

from openpyxl.utils.cell import get_column_letter
from PySide6.QtCore import QByteArray, QObject, QRunnable, QSettings, Qt, QThreadPool, Signal, Slot
from PySide6.QtGui import QCloseEvent, QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDockWidget,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from excel_git_viewer.git_repository import GitRepository, GitRepositoryError
from excel_git_viewer.history_cache import HistoryCache
from excel_git_viewer.models import (
    CellChange,
    CellContext,
    ChangedFile,
    CommitHistory,
    CommitInfo,
    WorkbookDiff,
)
from excel_git_viewer.task_coordinator import TaskCoordinator, TaskHandle
from excel_git_viewer.view_models import context_target_index, summarize_sheets
from excel_git_viewer.workbook_differ import (
    CancellationToken,
    OperationCancelled,
    WorkbookDiffer,
)

_APP_STYLESHEET = """
QMainWindow, QWidget#appRoot {
    background: #f4f6f8;
    color: #172033;
    font-family: "Segoe UI", "Microsoft YaHei UI";
    font-size: 13px;
}
QWidget#topBar, QWidget#statusBar {
    background: #ffffff;
    border: 1px solid #d8dee8;
    border-radius: 4px;
}
QLabel#repositoryLabel {
    color: #334155;
    font-weight: 600;
}
QLabel#branchLabel {
    background: #e8f0fe;
    color: #174ea6;
    border: 1px solid #c5d7f2;
    border-radius: 4px;
    padding: 4px 8px;
}
QLabel#sectionTitle {
    color: #334155;
    font-size: 12px;
    font-weight: 700;
    padding: 2px 1px 5px 1px;
}
QLabel#summaryLabel {
    background: #ffffff;
    border: 1px solid #d8dee8;
    border-left: 3px solid #2563eb;
    border-radius: 3px;
    padding: 7px 10px;
    font-weight: 600;
}
QPushButton {
    background: #ffffff;
    border: 1px solid #b8c2d1;
    border-radius: 4px;
    padding: 6px 10px;
    min-height: 18px;
}
QPushButton:hover { background: #f2f6fc; border-color: #7b8ba3; }
QPushButton:pressed { background: #e5edf8; }
QPushButton:disabled { color: #94a3b8; background: #f8fafc; }
QTableWidget {
    background: #ffffff;
    alternate-background-color: #f8fafc;
    border: 1px solid #d8dee8;
    border-radius: 3px;
    gridline-color: #e5e9f0;
    selection-background-color: #dbeafe;
    selection-color: #172033;
}
QTableWidget::item { padding: 5px 7px; }
QHeaderView::section {
    border: 0;
    border-right: 1px solid #d8dee8;
    border-bottom: 1px solid #cbd5e1;
    padding: 6px 7px;
    font-weight: 700;
}
QTableWidget#dataTable QHeaderView::section {
    background: #eef2f6;
    color: #475569;
}
QDockWidget {
    color: #334155;
}
QDockWidget::title {
    background: #eef2f6;
    border: 1px solid #d8dee8;
    padding: 6px 8px;
    text-align: left;
    font-weight: 600;
}
QProgressBar {
    border: 1px solid #cbd5e1;
    border-radius: 3px;
    background: #eef2f6;
}
QProgressBar::chunk { background: #2563eb; }
"""


class _WorkerSignals(QObject):
    completed = Signal(int, object)
    failed = Signal(int, str)
    cancelled = Signal(int)


class _Worker(QRunnable):
    def __init__(
        self,
        handle: TaskHandle,
        operation: Callable[[CancellationToken], object],
    ) -> None:
        super().__init__()
        self.handle = handle
        self.operation = operation
        self.signals = _WorkerSignals()

    @Slot()
    def run(self) -> None:
        try:
            result = self.operation(self.handle.cancellation)
            self.handle.cancellation.raise_if_cancelled()
        except OperationCancelled:
            self.signals.cancelled.emit(self.handle.task_id)
        except Exception as error:
            self.signals.failed.emit(self.handle.task_id, str(error))
        else:
            self.signals.completed.emit(self.handle.task_id, result)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Excel Git Viewer")
        self.resize(1380, 820)
        self.setMinimumSize(980, 640)
        self.setStyleSheet(_APP_STYLESHEET)

        self._repository: GitRepository | None = None
        self._history_cache = HistoryCache.default()
        self._settings = QSettings("ExcelGitViewer", "ExcelGitViewer")
        self._history: CommitHistory | None = None
        self._commits: list[CommitInfo] = []
        self._files: list[ChangedFile] = []
        self._current_diff: WorkbookDiff | None = None
        self._thread_pool = QThreadPool.globalInstance()
        self._workers: dict[tuple[int, int], _Worker] = {}
        self._commit_tasks = TaskCoordinator()
        self._file_tasks = TaskCoordinator()
        self._diff_tasks = TaskCoordinator()
        self._active_coordinator: TaskCoordinator | None = None

        self._build_ui()
        self._set_busy(False)
        self._restore_layout_state()
        self._restore_last_repository()

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("appRoot")
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(12, 12, 12, 10)
        root_layout.setSpacing(10)

        top_bar = QWidget()
        top_bar.setObjectName("topBar")
        repository_bar = QHBoxLayout(top_bar)
        repository_bar.setContentsMargins(10, 7, 10, 7)
        repository_bar.setSpacing(8)
        self.select_repository_button = QPushButton("选择仓库")
        open_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon)
        self.select_repository_button.setIcon(open_icon)
        self.select_repository_button.clicked.connect(self._choose_repository)
        repository_bar.addWidget(self.select_repository_button)

        self.repository_label = QLabel("尚未选择仓库")
        self.repository_label.setObjectName("repositoryLabel")
        self.repository_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        repository_bar.addWidget(self.repository_label, 1)

        self.branch_label = QLabel("分支：-")
        self.branch_label.setObjectName("branchLabel")
        repository_bar.addWidget(self.branch_label)

        self.refresh_button = QPushButton("刷新")
        refresh_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload)
        self.refresh_button.setIcon(refresh_icon)
        self.refresh_button.clicked.connect(self._force_reload_commits)
        self.refresh_button.setEnabled(False)
        repository_bar.addWidget(self.refresh_button)
        self.reset_layout_button = QPushButton("恢复默认布局")
        reset_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_TitleBarNormalButton)
        self.reset_layout_button.setIcon(reset_icon)
        self.reset_layout_button.clicked.connect(self._reset_layout)
        repository_bar.addWidget(self.reset_layout_button)
        root_layout.addWidget(top_bar)

        self.workspace = QMainWindow()
        self.workspace.setObjectName("reviewWorkspace")
        self.workspace.setDockNestingEnabled(True)
        self.workspace.setDockOptions(
            QMainWindow.DockOption.AllowNestedDocks
            | QMainWindow.DockOption.AllowTabbedDocks
            | QMainWindow.DockOption.AnimatedDocks
        )

        self.commit_table = self._data_table(4)
        self.commit_table.setHorizontalHeaderLabels(["提交", "说明", "作者", "时间"])
        self.commit_table.setMinimumWidth(300)
        self.commit_table.verticalHeader().setDefaultSectionSize(42)
        self._set_initial_column_widths(self.commit_table, (82, 215, 88, 120))
        self.commit_table.currentCellChanged.connect(self._select_commit)
        self.commit_dock = self._dock_panel("提交记录", "commitDock", self.commit_table)

        self.file_table = self._data_table(2)
        self.file_table.setHorizontalHeaderLabels(["状态", "Excel 文件"])
        self.file_table.setMinimumWidth(250)
        self.file_table.verticalHeader().setDefaultSectionSize(34)
        self._set_initial_column_widths(self.file_table, (58, 420))
        self.file_table.currentCellChanged.connect(self._select_file)
        self.file_dock = self._dock_panel("Excel 文件", "fileDock", self.file_table)

        self.summary_label = QLabel("选择 Excel 文件查看差异")
        self.summary_label.setObjectName("summaryLabel")

        self.sheet_table = self._data_table(3, selectable=False)
        self.sheet_table.setHorizontalHeaderLabels(["工作表", "状态", "单元格变化"])
        self._set_initial_column_widths(self.sheet_table, (280, 90, 120))
        self.sheet_dock = self._dock_panel("工作表概览", "sheetDock", self.sheet_table)

        self.cell_table = self._data_table(6)
        self.cell_table.setHorizontalHeaderLabels(
            ["工作表", "位置", "类型", "旧值", "新值", "标记"]
        )
        self._set_initial_column_widths(self.cell_table, (120, 70, 70, 210, 210, 140))
        self.cell_table.currentCellChanged.connect(self._show_context_for_row)
        diff_content = QWidget()
        diff_layout = QVBoxLayout(diff_content)
        diff_layout.setContentsMargins(0, 0, 0, 0)
        diff_layout.setSpacing(6)
        diff_layout.addWidget(self.summary_label)
        diff_layout.addWidget(self.cell_table, 1)
        self.cell_dock = self._dock_panel("差异记录", "cellDock", diff_content)

        self.old_context_table = self._context_table()
        self.new_context_table = self._context_table()
        self.old_context_dock = self._dock_panel(
            "修改前上下文", "oldContextDock", self.old_context_table
        )
        self.new_context_dock = self._dock_panel(
            "修改后上下文", "newContextDock", self.new_context_table
        )

        left = Qt.DockWidgetArea.LeftDockWidgetArea
        right = Qt.DockWidgetArea.RightDockWidgetArea
        for dock in (self.commit_dock, self.file_dock):
            self.workspace.addDockWidget(left, dock)
        for dock in (
            self.sheet_dock,
            self.cell_dock,
            self.old_context_dock,
            self.new_context_dock,
        ):
            self.workspace.addDockWidget(right, dock)
        self.workspace.splitDockWidget(self.commit_dock, self.file_dock, Qt.Orientation.Vertical)
        self.workspace.splitDockWidget(self.sheet_dock, self.cell_dock, Qt.Orientation.Vertical)
        self.workspace.splitDockWidget(
            self.cell_dock, self.old_context_dock, Qt.Orientation.Vertical
        )
        self.workspace.splitDockWidget(
            self.old_context_dock, self.new_context_dock, Qt.Orientation.Horizontal
        )
        self.workspace.resizeDocks(
            [self.commit_dock, self.sheet_dock], [500, 860], Qt.Orientation.Horizontal
        )
        self.workspace.resizeDocks(
            [self.commit_dock, self.file_dock], [500, 260], Qt.Orientation.Vertical
        )
        self.workspace.resizeDocks(
            [self.sheet_dock, self.cell_dock, self.old_context_dock],
            [140, 360, 240],
            Qt.Orientation.Vertical,
        )
        self._default_workspace_state = self.workspace.saveState()
        self._default_header_states = {
            key: header.saveState() for key, header in self._persistent_table_headers()
        }
        root_layout.addWidget(self.workspace, 1)

        status_bar = QWidget()
        status_bar.setObjectName("statusBar")
        status_layout = QHBoxLayout(status_bar)
        status_layout.setContentsMargins(10, 5, 8, 5)
        status_layout.setSpacing(8)
        self.status_label = QLabel("就绪")
        status_layout.addWidget(self.status_label, 1)
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setFixedWidth(180)
        status_layout.addWidget(self.progress)
        self.cancel_button = QPushButton("取消")
        cancel_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_DialogCancelButton)
        self.cancel_button.setIcon(cancel_icon)
        self.cancel_button.clicked.connect(self._cancel_active_task)
        status_layout.addWidget(self.cancel_button)
        root_layout.addWidget(status_bar)

        self.setCentralWidget(root)

    def _dock_panel(self, title: str, object_name: str, content: QWidget) -> QDockWidget:
        dock = QDockWidget(title, self.workspace)
        dock.setObjectName(object_name)
        dock.setAllowedAreas(Qt.DockWidgetArea.AllDockWidgetAreas)
        dock.setFeatures(
            QDockWidget.DockWidgetFeature.DockWidgetMovable
            | QDockWidget.DockWidgetFeature.DockWidgetFloatable
        )
        dock.setWidget(content)
        return dock

    @staticmethod
    def _data_table(column_count: int, *, selectable: bool = True) -> QTableWidget:
        table = QTableWidget(0, column_count)
        table.setObjectName("dataTable")
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        mode = (
            QAbstractItemView.SelectionMode.SingleSelection
            if selectable
            else QAbstractItemView.SelectionMode.NoSelection
        )
        table.setSelectionMode(mode)
        table.setAlternatingRowColors(True)
        table.setShowGrid(False)
        table.setWordWrap(False)
        table.verticalHeader().setVisible(False)
        MainWindow._make_columns_resizable(table, default_width=120)
        return table

    @staticmethod
    def _context_table() -> QTableWidget:
        table = QTableWidget()
        table.setObjectName("contextTable")
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        table.setAlternatingRowColors(True)
        MainWindow._make_columns_resizable(table, default_width=96)
        return table

    @staticmethod
    def _make_columns_resizable(table: QTableWidget, *, default_width: int) -> None:
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setMinimumSectionSize(48)
        header.setDefaultSectionSize(default_width)
        header.sectionHandleDoubleClicked.connect(table.resizeColumnToContents)

    @staticmethod
    def _set_initial_column_widths(table: QTableWidget, widths: tuple[int, ...]) -> None:
        for column, width in enumerate(widths):
            table.setColumnWidth(column, width)

    @Slot()
    def _choose_repository(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "选择 Git 仓库")
        if not path:
            return
        self._open_repository(Path(path), show_error=True)

    def _open_repository(self, path: Path, *, show_error: bool) -> None:
        try:
            repository = GitRepository(path, history_cache=self._history_cache)
        except (GitRepositoryError, ValueError) as error:
            if show_error:
                QMessageBox.warning(self, "无法打开仓库", str(error))
            else:
                self.status_label.setText("上次使用的仓库已不可用")
                self._settings.remove("last_repository")
            return
        self._repository = repository
        self._settings.setValue("last_repository", str(repository.root))
        self.repository_label.setText(str(repository.root))
        branch = repository.current_branch or "detached HEAD"
        self.branch_label.setText(f"分支：{branch}")
        self.refresh_button.setEnabled(True)
        self._reload_commits()

    def _restore_last_repository(self) -> None:
        stored_path = self._settings.value("last_repository")
        if stored_path:
            self._open_repository(Path(str(stored_path)), show_error=False)

    def _persistent_table_headers(self) -> tuple[tuple[str, QHeaderView], ...]:
        return (
            ("layout/commit_columns", self.commit_table.horizontalHeader()),
            ("layout/file_columns", self.file_table.horizontalHeader()),
            ("layout/sheet_columns", self.sheet_table.horizontalHeader()),
            ("layout/cell_columns", self.cell_table.horizontalHeader()),
        )

    def _restore_layout_state(self) -> None:
        workspace_state = self._settings.value("layout/workspace")
        if isinstance(workspace_state, QByteArray):
            self.workspace.restoreState(workspace_state)
        for key, header in self._persistent_table_headers():
            state = self._settings.value(key)
            if isinstance(state, QByteArray):
                header.restoreState(state)

    @Slot()
    def _reset_layout(self) -> None:
        self.workspace.restoreState(self._default_workspace_state)
        for key, header in self._persistent_table_headers():
            header.restoreState(self._default_header_states[key])
        self.status_label.setText("已恢复默认布局")

    def closeEvent(self, event: QCloseEvent) -> None:
        self._settings.setValue("layout/workspace", self.workspace.saveState())
        for key, header in self._persistent_table_headers():
            self._settings.setValue(key, header.saveState())
        super().closeEvent(event)

    @Slot()
    def _force_reload_commits(self) -> None:
        self._reload_commits(force_refresh=True)

    def _reload_commits(self, *, force_refresh: bool = False) -> None:
        if self._repository is None:
            return
        repository = self._repository
        self._file_tasks.cancel()
        self._diff_tasks.cancel()
        self._clear_files_and_diff()
        self._history = None
        self._commits = []
        self.commit_table.setRowCount(0)

        def load(token: CancellationToken) -> object:
            token.raise_if_cancelled()
            result = repository.load_recent_history(force_refresh=force_refresh)
            token.raise_if_cancelled()
            return result

        self._start_task(
            self._commit_tasks,
            load,
            self._history_loaded,
            "正在读取提交记录……",
        )

    def _history_loaded(self, result: object) -> None:
        self._history = cast(CommitHistory, result)
        self._commits = list(self._history.all_commits)
        self.commit_table.setRowCount(len(self._commits))
        for row, commit in enumerate(self._commits):
            timestamp = commit.authored_at.astimezone().strftime("%Y-%m-%d %H:%M")
            values = [commit.commit_id[:8], commit.subject, commit.author_name, timestamp]
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setToolTip(value)
                self.commit_table.setItem(row, column, item)
        source = "本地缓存" if self._history.source == "cache" else "Git"
        self.status_label.setText(
            f"已从{source}加载最近 {self._history.scanned_commit_count} 条普通提交，"
            f"显示 {len(self._commits)} 条提交"
        )

    @Slot(int, int, int, int)
    def _select_commit(
        self,
        row: int,
        _column: int,
        previous_row: int,
        _previous_column: int,
    ) -> None:
        if row == previous_row:
            return
        if self._repository is None or row < 0 or row >= len(self._commits):
            return
        repository = self._repository
        commit = self._commits[row]
        self._diff_tasks.cancel()
        self._clear_files_and_diff()

        def load(token: CancellationToken) -> object:
            token.raise_if_cancelled()
            result = repository.list_changed_excel_files(commit.commit_id)
            token.raise_if_cancelled()
            return result

        self._start_task(self._file_tasks, load, self._files_loaded, "正在读取 Excel 文件列表……")

    def _files_loaded(self, result: object) -> None:
        self._files = cast(list[ChangedFile], result)
        self.file_table.setRowCount(len(self._files))
        labels = {"added": "A", "modified": "M", "deleted": "D", "renamed": "R"}
        for row, changed_file in enumerate(self._files):
            label = labels[changed_file.change_type]
            for column, value in enumerate((label, changed_file.display_path)):
                item = QTableWidgetItem(value)
                item.setToolTip(changed_file.display_path)
                self.file_table.setItem(row, column, item)
        self.status_label.setText(f"发现 {len(self._files)} 个 Excel 文件变化")

    @Slot(int, int, int, int)
    def _select_file(
        self,
        row: int,
        _column: int,
        previous_row: int,
        _previous_column: int,
    ) -> None:
        if row == previous_row:
            return
        if self._repository is None or row < 0 or row >= len(self._files):
            return
        repository = self._repository
        changed_file = self._files[row]
        self._clear_diff()

        def compare(token: CancellationToken) -> object:
            token.raise_if_cancelled()
            old_bytes, new_bytes = repository.read_versions(changed_file)
            token.raise_if_cancelled()
            return WorkbookDiffer().compare(old_bytes, new_bytes, cancellation=token)

        self._start_task(self._diff_tasks, compare, self._diff_loaded, "正在解析工作簿……")

    def _diff_loaded(self, result: object) -> None:
        workbook_diff = cast(WorkbookDiff, result)
        self._current_diff = workbook_diff
        self._fill_sheet_table(workbook_diff)
        self.cell_table.setRowCount(len(workbook_diff.cell_changes))
        type_labels = {"added": "新增", "modified": "修改", "deleted": "删除"}
        for row, change in enumerate(workbook_diff.cell_changes):
            flags = self._change_flags(change)
            values = [
                change.sheet_name,
                change.coordinate,
                type_labels[change.change_type],
                self._display_value(change.old_value, change.whitespace_warning),
                self._display_value(change.new_value, change.whitespace_warning),
                flags,
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column == 3:
                    item.setBackground(QColor("#fde8e8"))
                    item.setForeground(QColor("#9b1c1c"))
                elif column == 4:
                    item.setBackground(QColor("#dcfce7"))
                    item.setForeground(QColor("#166534"))
                elif column == 5 and flags:
                    item.setBackground(QColor("#fef3c7"))
                    item.setForeground(QColor("#92400e"))
                item.setToolTip(value)
                self.cell_table.setItem(row, column, item)

        sheet_count = len(workbook_diff.sheet_changes)
        cell_count = len(workbook_diff.cell_changes)
        if sheet_count == 0 and cell_count == 0:
            self.summary_label.setText(
                "Git 文件已变化，但未检测到当前支持范围内的内容差异；"
                "可能修改了格式、图表、批注或其他未支持对象。"
            )
        else:
            self.summary_label.setText(f"工作表变化 {sheet_count}，单元格变化 {cell_count}")
        if workbook_diff.cell_changes:
            self.cell_table.selectRow(0)
        self.status_label.setText("差异解析完成")

    def _fill_sheet_table(self, workbook_diff: WorkbookDiff) -> None:
        summaries = summarize_sheets(workbook_diff)
        self.sheet_table.setRowCount(len(summaries))
        status_labels = {"added": "新增", "deleted": "删除", "modified": "修改"}
        for row, summary in enumerate(summaries):
            values = [
                summary.sheet_name,
                status_labels[summary.status],
                str(summary.cell_change_count),
            ]
            for column, value in enumerate(values):
                self.sheet_table.setItem(row, column, QTableWidgetItem(value))

    @staticmethod
    def _change_flags(change: CellChange) -> str:
        flags: list[str] = []
        if change.hidden_sheet:
            flags.append("隐藏工作表")
        if change.hidden_row:
            flags.append("隐藏行")
        if change.hidden_column:
            flags.append("隐藏列")
        if change.whitespace_warning:
            flags.append("包含不可见空白")
        if change.old_data_type != change.new_data_type:
            old_type = change.old_data_type or "无"
            new_type = change.new_data_type or "无"
            flags.append(f"类型：{old_type} -> {new_type}")
        return "、".join(flags)

    @staticmethod
    def _display_value(value: str | None, reveal_whitespace: bool) -> str:
        if value is None:
            return ""
        if not reveal_whitespace:
            return value
        return (
            value.replace(" ", "<SP>")
            .replace("\t", "<TAB>")
            .replace("\r", "<CR>")
            .replace("\n", "<LF>\n")
        )

    @Slot(int, int, int, int)
    def _show_context_for_row(
        self,
        current_row: int,
        _current_column: int,
        _previous_row: int,
        _previous_column: int,
    ) -> None:
        if (
            self._current_diff is None
            or current_row < 0
            or current_row >= len(self._current_diff.cell_changes)
        ):
            return
        change = self._current_diff.cell_changes[current_row]
        self._fill_context(
            self.old_context_table,
            change.old_context,
            change.coordinate,
            background=QColor("#fecaca"),
            foreground=QColor("#991b1b"),
        )
        self._fill_context(
            self.new_context_table,
            change.new_context,
            change.coordinate,
            background=QColor("#bbf7d0"),
            foreground=QColor("#166534"),
        )

    @staticmethod
    def _fill_context(
        table: QTableWidget,
        context: CellContext | None,
        target_coordinate: str,
        *,
        background: QColor,
        foreground: QColor,
    ) -> None:
        table.clear()
        if context is None:
            table.setRowCount(0)
            table.setColumnCount(0)
            return
        row_count = len(context.values)
        column_count = len(context.values[0]) if context.values else 0
        table.setRowCount(row_count)
        table.setColumnCount(column_count)
        row_labels = [str(context.start_row + index) for index in range(row_count)]
        table.setVerticalHeaderLabels(row_labels)
        table.setHorizontalHeaderLabels(
            [get_column_letter(context.start_column + index) for index in range(column_count)]
        )
        header_background = QColor("#eef2f6")
        header_foreground = QColor("#475569")
        for column in range(column_count):
            item = table.horizontalHeaderItem(column)
            if item is not None:
                item.setBackground(header_background)
                item.setForeground(header_foreground)
        for row in range(row_count):
            item = table.verticalHeaderItem(row)
            if item is not None:
                item.setBackground(header_background)
                item.setForeground(header_foreground)
        target_index = context_target_index(context, target_coordinate)
        for row, values in enumerate(context.values):
            for column, value in enumerate(values):
                item = QTableWidgetItem(value or "")
                if target_index == (row, column):
                    item.setBackground(background)
                    item.setForeground(foreground)
                    font = item.font()
                    font.setBold(True)
                    item.setFont(font)
                    item.setToolTip(f"当前改动单元格 {target_coordinate}")
                table.setItem(row, column, item)
        if target_index is not None:
            target_row, target_column = target_index
            target_header_background = QColor("#fef3c7")
            target_header_foreground = QColor("#92400e")
            horizontal_header = table.horizontalHeaderItem(target_column)
            vertical_header = table.verticalHeaderItem(target_row)
            if horizontal_header is not None:
                horizontal_header.setText(f"> {horizontal_header.text()}")
                horizontal_header.setBackground(target_header_background)
                horizontal_header.setForeground(target_header_foreground)
            if vertical_header is not None:
                vertical_header.setText(f"> {vertical_header.text()}")
                vertical_header.setBackground(target_header_background)
                vertical_header.setForeground(target_header_foreground)
            target_item = table.item(target_row, target_column)
            if target_item is not None:
                table.scrollToItem(target_item)

    def _start_task(
        self,
        coordinator: TaskCoordinator,
        operation: Callable[[CancellationToken], object],
        on_success: Callable[[object], None],
        status: str,
    ) -> None:
        handle = coordinator.begin()
        self._active_coordinator = coordinator
        worker = _Worker(handle, operation)
        self._retain_worker(coordinator, handle.task_id, worker)
        worker.signals.completed.connect(
            lambda task_id, result: self._task_completed(coordinator, task_id, result, on_success)
        )
        worker.signals.failed.connect(
            lambda task_id, message: self._task_failed(coordinator, task_id, message)
        )
        worker.signals.cancelled.connect(lambda task_id: self._task_cancelled(coordinator, task_id))
        self.status_label.setText(status)
        self._set_busy(True)
        self._thread_pool.start(worker)

    def _retain_worker(
        self,
        coordinator: TaskCoordinator,
        task_id: int,
        worker: _Worker,
    ) -> None:
        self._workers[(id(coordinator), task_id)] = worker

    def _release_worker(self, coordinator: TaskCoordinator, task_id: int) -> None:
        self._workers.pop((id(coordinator), task_id), None)

    def _task_completed(
        self,
        coordinator: TaskCoordinator,
        task_id: int,
        result: object,
        on_success: Callable[[object], None],
    ) -> None:
        self._release_worker(coordinator, task_id)
        if not coordinator.is_current(task_id):
            return
        self._active_coordinator = None
        self._set_busy(False)
        on_success(result)

    def _task_failed(
        self,
        coordinator: TaskCoordinator,
        task_id: int,
        message: str,
    ) -> None:
        self._release_worker(coordinator, task_id)
        if not coordinator.is_current(task_id):
            return
        self._active_coordinator = None
        self._set_busy(False)
        self.status_label.setText("操作失败")
        QMessageBox.warning(self, "操作失败", message or "未知错误")

    def _task_cancelled(
        self,
        coordinator: TaskCoordinator,
        task_id: int,
    ) -> None:
        self._release_worker(coordinator, task_id)
        if not coordinator.is_current(task_id):
            return
        self._active_coordinator = None
        self._set_busy(False)
        self.status_label.setText("已取消")

    @Slot()
    def _cancel_active_task(self) -> None:
        if self._active_coordinator is not None:
            self._active_coordinator.cancel()
        self._active_coordinator = None
        self._set_busy(False)
        self.status_label.setText("已取消")

    def _set_busy(self, busy: bool) -> None:
        self.progress.setVisible(busy)
        self.cancel_button.setVisible(busy)

    def _clear_files_and_diff(self) -> None:
        self._files = []
        self.file_table.setRowCount(0)
        self._clear_diff()

    def _clear_diff(self) -> None:
        self._current_diff = None
        self.sheet_table.setRowCount(0)
        self.cell_table.setRowCount(0)
        self.old_context_table.clear()
        self.old_context_table.setRowCount(0)
        self.old_context_table.setColumnCount(0)
        self.new_context_table.clear()
        self.new_context_table.setRowCount(0)
        self.new_context_table.setColumnCount(0)
        self.summary_label.setText("选择 Excel 文件查看差异")


def create_application() -> tuple[QApplication, MainWindow]:
    existing = QApplication.instance()
    application = existing if isinstance(existing, QApplication) else QApplication([])
    window = MainWindow()
    return application, window
