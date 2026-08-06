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
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
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
    background: #eef2f6;
    color: #475569;
    border: 0;
    border-right: 1px solid #d8dee8;
    border-bottom: 1px solid #cbd5e1;
    padding: 6px 7px;
    font-weight: 700;
}
QSplitter::handle { background: #d8dee8; }
QSplitter::handle:hover { background: #94a3b8; }
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
        self._restore_splitter_state()
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
        root_layout.addWidget(top_bar)

        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.main_splitter.setChildrenCollapsible(False)

        self.commit_table = self._data_table(4)
        self.commit_table.setHorizontalHeaderLabels(["提交", "说明", "作者", "时间"])
        self.commit_table.setMinimumWidth(300)
        self.commit_table.verticalHeader().setDefaultSectionSize(42)
        commit_header = self.commit_table.horizontalHeader()
        commit_header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        commit_header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        commit_header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        commit_header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.commit_table.setColumnWidth(0, 82)
        self.commit_table.currentCellChanged.connect(self._select_commit)
        self.main_splitter.addWidget(self._panel("提交记录", self.commit_table))

        self.file_table = self._data_table(2)
        self.file_table.setHorizontalHeaderLabels(["状态", "Excel 文件"])
        self.file_table.setMinimumWidth(250)
        self.file_table.verticalHeader().setDefaultSectionSize(34)
        file_header = self.file_table.horizontalHeader()
        file_header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        file_header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.file_table.setColumnWidth(0, 52)
        self.file_table.currentCellChanged.connect(self._select_file)
        self.main_splitter.addWidget(self._panel("Excel 文件", self.file_table))

        details = QWidget()
        details_layout = QVBoxLayout(details)
        details_layout.setContentsMargins(0, 0, 0, 0)
        details_layout.setSpacing(8)
        self.summary_label = QLabel("选择 Excel 文件查看差异")
        self.summary_label.setObjectName("summaryLabel")
        details_layout.addWidget(self.summary_label)

        self.details_splitter = QSplitter(Qt.Orientation.Vertical)
        self.details_splitter.setChildrenCollapsible(False)

        self.sheet_table = self._data_table(3, selectable=False)
        self.sheet_table.setHorizontalHeaderLabels(["工作表", "状态", "单元格变化"])
        self.sheet_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.sheet_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.ResizeToContents
        )
        self.sheet_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeMode.ResizeToContents
        )
        self.details_splitter.addWidget(self._panel("工作表概览", self.sheet_table))

        self.cell_table = self._data_table(6)
        self.cell_table.setHorizontalHeaderLabels(
            ["工作表", "位置", "类型", "旧值", "新值", "标记"]
        )
        header = self.cell_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.cell_table.currentCellChanged.connect(self._show_context_for_row)
        self.details_splitter.addWidget(self._panel("差异记录", self.cell_table))

        self.context_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.context_splitter.setChildrenCollapsible(False)
        self.old_context_table = self._context_table()
        self.new_context_table = self._context_table()
        self.context_splitter.addWidget(self._panel("修改前上下文", self.old_context_table))
        self.context_splitter.addWidget(self._panel("修改后上下文", self.new_context_table))
        self.context_splitter.setSizes([1, 1])
        self.details_splitter.addWidget(self.context_splitter)
        self.details_splitter.setSizes([150, 350, 240])
        details_layout.addWidget(self.details_splitter, 1)

        self.main_splitter.addWidget(details)
        self.main_splitter.setSizes([360, 300, 720])
        self.main_splitter.setStretchFactor(0, 1)
        self.main_splitter.setStretchFactor(1, 1)
        self.main_splitter.setStretchFactor(2, 3)
        root_layout.addWidget(self.main_splitter, 1)

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

    @staticmethod
    def _panel(title: str, content: QWidget) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        heading = QLabel(title)
        heading.setObjectName("sectionTitle")
        layout.addWidget(heading)
        layout.addWidget(content, 1)
        return panel

    @staticmethod
    def _data_table(column_count: int, *, selectable: bool = True) -> QTableWidget:
        table = QTableWidget(0, column_count)
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
        return table

    @staticmethod
    def _context_table() -> QTableWidget:
        table = QTableWidget()
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        table.setAlternatingRowColors(True)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        return table

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

    def _restore_splitter_state(self) -> None:
        for key, splitter in (
            ("layout/main_splitter", self.main_splitter),
            ("layout/details_splitter", self.details_splitter),
            ("layout/context_splitter", self.context_splitter),
        ):
            state = self._settings.value(key)
            if isinstance(state, QByteArray):
                splitter.restoreState(state)

    def closeEvent(self, event: QCloseEvent) -> None:
        self._settings.setValue("layout/main_splitter", self.main_splitter.saveState())
        self._settings.setValue("layout/details_splitter", self.details_splitter.saveState())
        self._settings.setValue("layout/context_splitter", self.context_splitter.saveState())
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
        _previous_row: int,
        _previous_column: int,
    ) -> None:
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
        _previous_row: int,
        _previous_column: int,
    ) -> None:
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
            header_background = QColor("#fef3c7")
            header_foreground = QColor("#92400e")
            horizontal_header = table.horizontalHeaderItem(target_column)
            vertical_header = table.verticalHeaderItem(target_row)
            if horizontal_header is not None:
                horizontal_header.setBackground(header_background)
                horizontal_header.setForeground(header_foreground)
            if vertical_header is not None:
                vertical_header.setBackground(header_background)
                vertical_header.setForeground(header_foreground)
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
