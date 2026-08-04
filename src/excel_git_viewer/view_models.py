from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from typing import Literal

from excel_git_viewer.models import WorkbookDiff

SheetStatus = Literal["added", "deleted", "modified"]


@dataclass(frozen=True, slots=True)
class SheetSummary:
    sheet_name: str
    status: SheetStatus
    cell_change_count: int


def summarize_sheets(workbook_diff: WorkbookDiff) -> tuple[SheetSummary, ...]:
    counts = Counter(change.sheet_name for change in workbook_diff.cell_changes)
    statuses: dict[str, SheetStatus] = {
        change.sheet_name: change.change_type for change in workbook_diff.sheet_changes
    }
    names = set(counts) | set(statuses)
    return tuple(
        SheetSummary(
            sheet_name=name,
            status=statuses.get(name, "modified"),
            cell_change_count=counts[name],
        )
        for name in sorted(names)
    )
