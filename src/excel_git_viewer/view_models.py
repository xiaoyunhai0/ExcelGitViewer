from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from typing import Literal

from openpyxl.utils.cell import coordinate_to_tuple

from excel_git_viewer.models import CellContext, WorkbookDiff

SheetStatus = Literal["added", "deleted", "modified"]


@dataclass(frozen=True, slots=True)
class SheetSummary:
    sheet_name: str
    status: SheetStatus
    cell_change_count: int


def context_target_index(context: CellContext, coordinate: str) -> tuple[int, int] | None:
    row, column = coordinate_to_tuple(coordinate)
    row_index = row - context.start_row
    column_index = column - context.start_column
    row_count = len(context.values)
    column_count = len(context.values[0]) if context.values else 0
    if 0 <= row_index < row_count and 0 <= column_index < column_count:
        return row_index, column_index
    return None


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
