from __future__ import annotations

from dataclasses import dataclass
from threading import Lock

from excel_git_viewer.workbook_differ import CancellationToken


@dataclass(frozen=True, slots=True)
class TaskHandle:
    task_id: int
    cancellation: CancellationToken


class TaskCoordinator:
    """Own the current background task and invalidate stale results."""

    def __init__(self) -> None:
        self._lock = Lock()
        self._current: TaskHandle | None = None
        self._next_id = 1

    def begin(self) -> TaskHandle:
        with self._lock:
            if self._current is not None:
                self._current.cancellation.cancel()
            handle = TaskHandle(self._next_id, CancellationToken())
            self._next_id += 1
            self._current = handle
            return handle

    def cancel(self) -> None:
        with self._lock:
            if self._current is not None:
                self._current.cancellation.cancel()
            self._current = None

    def is_current(self, task_id: int) -> bool:
        with self._lock:
            return self._current is not None and self._current.task_id == task_id
