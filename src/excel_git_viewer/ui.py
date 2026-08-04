from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import cast

from openpyxl.utils.cell import get_column_letter
from PySide6.QtCore import QObject, QRunnable, Qt, QThreadPool, Signal, Slot
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
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
from excel_git_viewer.models import CellChange, CellContext, ChangedFile, CommitInfo, WorkbookDiff
from excel_git_viewer.task_coordinator import TaskCoordinator, TaskHandle
from excel_git_viewer.view_models import summarize_sheets
from excel_git_viewer.workbook_differ import (
    CancellationToken,
    OperationCancelled,
    WorkbookDiffer,
)


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

        self._repository: GitRepository | None = None
        self._commits: list[CommitInfo] = []
        self._files: list[ChangedFile] = []
        self._current_diff: WorkbookDiff | None = None
        self._thread_pool = QThreadPool.globalInstance()
        self._commit_tasks = TaskCoordinator()
        self._file_tasks = TaskCoordinator()
        self._diff_tasks = TaskCoordinator()
        self._active_coordinator: TaskCoordinator | None = None

        self._build_ui()
        self._set_busy(False)

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(10, 10, 10, 10)
        root_layout.setSpacing(8)

        repository_bar = QHBoxLayout()
        self.select_repository_button = QPushButton("选择仓库")
        open_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon)
        self.select_repository_button.setIcon(open_icon)
        self.select_repository_button.clicked.connect(self._choose_repository)
        repository_bar.addWidget(self.select_repository_button)

        self.repository_label = QLabel("尚未选择仓库")
        self.repository_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        repository_bar.addWidget(self.repository_label, 1)

        self.branch_label = QLabel("分支：-")
        repository_bar.addWidget(self.branch_label)

        self.only_excel_checkbox = QCheckBox("仅 Excel 提交")
        self.only_excel_checkbox.setChecked(True)
        self.only_excel_checkbox.toggled.connect(self._reload_commits)
        repository_bar.addWidget(self.only_excel_checkbox)

        self.refresh_button = QPushButton("刷新")
        refresh_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload)
        self.refresh_button.setIcon(refresh_icon)
        self.refresh_button.clicked.connect(self._reload_commits)
        self.refresh_button.setEnabled(False)
        repository_bar.addWidget(self.refresh_button)
        root_layout.addLayout(repository_bar)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.commit_list = QListWidget()
        self.commit_list.setMinimumWidth(260)
        self.commit_list.currentRowChanged.connect(self._select_commit)
        splitter.addWidget(self._panel("提交记录", self.commit_list))

        self.file_list = QListWidget()
        self.file_list.setMinimumWidth(250)
        self.file_list.currentRowChanged.connect(self._select_file)
        splitter.addWidget(self._panel("Excel 文件", self.file_list))

        details = QWidget()
        details_layout = QVBoxLayout(details)
        details_layout.setContentsMargins(0, 0, 0, 0)
        self.summary_label = QLabel("选择 Excel 文件查看差异")
        details_layout.addWidget(self.summary_label)

        self.sheet_table = QTableWidget(0, 3)
        self.sheet_table.setHorizontalHeaderLabels(["工作表", "状态", "单元格变化"])
        self.sheet_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.sheet_table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.sheet_table.verticalHeader().setVisible(False)
        self.sheet_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.sheet_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.ResizeToContents
        )
        self.sheet_table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeMode.ResizeToContents
        )
        self.sheet_table.setMaximumHeight(150)
        details_layout.addWidget(self.sheet_table)

        self.cell_table = QTableWidget(0, 6)
        self.cell_table.setHorizontalHeaderLabels(
            ["工作表", "位置", "类型", "旧值", "新值", "标记"]
        )
        self.cell_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.cell_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.cell_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.cell_table.verticalHeader().setVisible(False)
        header = self.cell_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.cell_table.currentCellChanged.connect(self._show_context_for_row)
        details_layout.addWidget(self.cell_table, 3)

        context_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.old_context_table = self._context_table()
        self.new_context_table = self._context_table()
        context_splitter.addWidget(self._panel("修改前上下文", self.old_context_table))
        context_splitter.addWidget(self._panel("修改后上下文", self.new_context_table))
        context_splitter.setSizes([1, 1])
        details_layout.addWidget(context_splitter, 2)
        splitter.addWidget(details)
        splitter.setSizes([280, 280, 820])
        splitter.setStretchFactor(2, 1)
        root_layout.addWidget(splitter, 1)

        status_layout = QHBoxLayout()
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
        root_layout.addLayout(status_layout)

        self.setCentralWidget(root)

    @staticmethod
    def _panel(title: str, content: QWidget) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        heading = QLabel(title)
        heading.setStyleSheet("font-weight: 600; padding: 2px 0;")
        layout.addWidget(heading)
        layout.addWidget(content, 1)
        return panel

    @staticmethod
    def _context_table() -> QTableWidget:
        table = QTableWidget()
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        return table

    @Slot()
    def _choose_repository(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "选择 Git 仓库")
        if not path:
            return
        try:
            repository = GitRepository(Path(path))
        except (GitRepositoryError, ValueError) as error:
            QMessageBox.warning(self, "无法打开仓库", str(error))
            return
        self._repository = repository
        self.repository_label.setText(str(repository.root))
        branch = repository.current_branch or "detached HEAD"
        self.branch_label.setText(f"分支：{branch}")
        self.refresh_button.setEnabled(True)
        self._reload_commits()

    @Slot()
    def _reload_commits(self) -> None:
        if self._repository is None:
            return
        repository = self._repository
        only_excel = self.only_excel_checkbox.isChecked()
        self._file_tasks.cancel()
        self._diff_tasks.cancel()
        self._clear_files_and_diff()
        self.commit_list.clear()

        def load(token: CancellationToken) -> object:
            token.raise_if_cancelled()
            result = repository.list_commits(only_excel=only_excel)
            token.raise_if_cancelled()
            return result

        self._start_task(
            self._commit_tasks,
            load,
            self._commits_loaded,
            "正在读取提交记录……",
        )

    def _commits_loaded(self, result: object) -> None:
        self._commits = cast(list[CommitInfo], result)
        self.commit_list.clear()
        for commit in self._commits:
            timestamp = commit.authored_at.astimezone().strftime("%Y-%m-%d %H:%M")
            self.commit_list.addItem(
                f"{commit.commit_id[:8]}  {commit.subject}\n{commit.author_name}  {timestamp}"
            )
        self.status_label.setText(f"已加载 {len(self._commits)} 条提交")

    @Slot(int)
    def _select_commit(self, row: int) -> None:
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
        self.file_list.clear()
        labels = {"added": "A", "modified": "M", "deleted": "D", "renamed": "R"}
        for changed_file in self._files:
            label = labels[changed_file.change_type]
            self.file_list.addItem(f"{label}  {changed_file.display_path}")
        self.status_label.setText(f"发现 {len(self._files)} 个 Excel 文件变化")

    @Slot(int)
    def _select_file(self, row: int) -> None:
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
                if flags:
                    item.setBackground(QColor("#fff3cd"))
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
        self._fill_context(self.old_context_table, change.old_context)
        self._fill_context(self.new_context_table, change.new_context)

    @staticmethod
    def _fill_context(table: QTableWidget, context: CellContext | None) -> None:
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
        for row, values in enumerate(context.values):
            for column, value in enumerate(values):
                table.setItem(row, column, QTableWidgetItem(value or ""))

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

    def _task_completed(
        self,
        coordinator: TaskCoordinator,
        task_id: int,
        result: object,
        on_success: Callable[[object], None],
    ) -> None:
        if not coordinator.is_current(task_id):
            return
        self._active_coordinator = None
        self._set_busy(False)
        on_success(result)

    def _task_failed(self, coordinator: TaskCoordinator, task_id: int, message: str) -> None:
        if not coordinator.is_current(task_id):
            return
        self._active_coordinator = None
        self._set_busy(False)
        self.status_label.setText("操作失败")
        QMessageBox.warning(self, "操作失败", message or "未知错误")

    def _task_cancelled(self, coordinator: TaskCoordinator, task_id: int) -> None:
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
        self.file_list.clear()
        self._clear_diff()

    def _clear_diff(self) -> None:
        self._current_diff = None
        self.sheet_table.setRowCount(0)
        self.cell_table.setRowCount(0)
        self.old_context_table.clear()
        self.old_context_table.setRowCount(0)
        self.new_context_table.clear()
        self.new_context_table.setRowCount(0)
        self.summary_label.setText("选择 Excel 文件查看差异")


def create_application() -> tuple[QApplication, MainWindow]:
    existing = QApplication.instance()
    application = existing if isinstance(existing, QApplication) else QApplication([])
    window = MainWindow()
    return application, window
