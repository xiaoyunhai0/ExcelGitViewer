from __future__ import annotations

import subprocess
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook

from excel_git_viewer.git_repository import GitRepository
from excel_git_viewer.workbook_differ import WorkbookDiffer


def _git(repo: Path, *args: str) -> str:
    return subprocess.run(
        ["git", "-C", str(repo), *args],
        check=True,
        capture_output=True,
        text=True,
    ).stdout.strip()


def _xlsx(value: int) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Rewards"
    sheet["A1"] = "Reward"
    sheet["B2"] = value
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def _commit_workbook(repo: Path, value: int, subject: str) -> None:
    path = repo / "Design" / "rewards.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(_xlsx(value))
    _git(repo, "add", "Design/rewards.xlsx")
    _git(repo, "commit", "-m", subject)


def test_user_can_review_a_cell_change_from_a_git_commit(tmp_path: Path) -> None:
    repo_path = tmp_path / "game"
    repo_path.mkdir()
    _git(repo_path, "init", "-b", "main")
    _git(repo_path, "config", "user.name", "Planner")
    _git(repo_path, "config", "user.email", "planner@example.com")
    _commit_workbook(repo_path, 100, "add rewards")
    _commit_workbook(repo_path, 150, "raise reward")

    repository = GitRepository(repo_path)
    [commit, _previous_commit] = repository.load_recent_history().excel_commits
    [changed_file] = repository.list_changed_excel_files(commit.commit_id)
    old_bytes, new_bytes = repository.read_versions(changed_file)
    [change] = WorkbookDiffer().compare(old_bytes, new_bytes).cell_changes

    assert (
        commit.subject,
        changed_file.display_path,
        change.sheet_name,
        change.coordinate,
        change.old_value,
        change.new_value,
    ) == ("raise reward", "Design/rewards.xlsx", "Rewards", "B2", "100", "150")
