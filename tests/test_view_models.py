from excel_git_viewer.models import CellChange, SheetChange, WorkbookDiff
from excel_git_viewer.view_models import summarize_sheets


def test_sheet_summary_names_added_deleted_and_changed_sheets() -> None:
    workbook_diff = WorkbookDiff(
        cell_changes=(
            CellChange("Items", "A1", "modified", "old", "new"),
            CellChange("Items", "A2", "added", None, "value"),
        ),
        sheet_changes=(
            SheetChange("New Sheet", "added"),
            SheetChange("Old Sheet", "deleted"),
        ),
    )

    summaries = summarize_sheets(workbook_diff)

    assert [
        (summary.sheet_name, summary.status, summary.cell_change_count) for summary in summaries
    ] == [
        ("Items", "modified", 2),
        ("New Sheet", "added", 0),
        ("Old Sheet", "deleted", 0),
    ]
