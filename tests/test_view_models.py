from excel_git_viewer.models import CellChange, CellContext, SheetChange, WorkbookDiff
from excel_git_viewer.view_models import context_target_index, summarize_sheets


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


def test_context_target_index_locates_the_changed_coordinate() -> None:
    context = CellContext(
        start_row=3,
        start_column=2,
        values=((None, None, None), (None, "changed", None), (None, None, None)),
    )

    assert context_target_index(context, "C4") == (1, 1)
    assert context_target_index(context, "A1") is None
