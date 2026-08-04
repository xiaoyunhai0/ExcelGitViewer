from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Literal


@dataclass(frozen=True, slots=True)
class CommitInfo:
    commit_id: str
    parent_ids: tuple[str, ...]
    author_name: str
    author_email: str
    authored_at: datetime
    subject: str


@dataclass(frozen=True, slots=True)
class CommitHistory:
    all_commits: tuple[CommitInfo, ...]
    excel_commits: tuple[CommitInfo, ...]
    scanned_commit_count: int
    scan_limit: int
    source: Literal["git", "cache"] = "git"

    def visible_commits(self, *, only_excel: bool) -> tuple[CommitInfo, ...]:
        return self.excel_commits if only_excel else self.all_commits


FileChangeType = Literal["added", "modified", "deleted", "renamed"]
CellChangeType = Literal["added", "modified", "deleted"]
SheetChangeType = Literal["added", "deleted"]


@dataclass(frozen=True, slots=True)
class ChangedFile:
    commit_id: str
    parent_id: str | None
    change_type: FileChangeType
    old_path: str | None
    new_path: str | None

    @property
    def display_path(self) -> str:
        if self.change_type == "renamed":
            return f"{self.old_path} -> {self.new_path}"
        return self.new_path or self.old_path or ""


@dataclass(frozen=True, slots=True)
class CellContext:
    start_row: int
    start_column: int
    values: tuple[tuple[str | None, ...], ...]


@dataclass(frozen=True, slots=True)
class CellChange:
    sheet_name: str
    coordinate: str
    change_type: CellChangeType
    old_value: str | None
    new_value: str | None
    old_data_type: str | None = None
    new_data_type: str | None = None
    hidden_sheet: bool = False
    hidden_row: bool = False
    hidden_column: bool = False
    whitespace_warning: bool = False
    old_context: CellContext | None = None
    new_context: CellContext | None = None


@dataclass(frozen=True, slots=True)
class SheetChange:
    sheet_name: str
    change_type: SheetChangeType


@dataclass(frozen=True, slots=True)
class WorkbookDiff:
    cell_changes: tuple[CellChange, ...]
    sheet_changes: tuple[SheetChange, ...] = ()
