from __future__ import annotations

import subprocess
from pathlib import Path

from excel_git_viewer.git_repository import GitRepository


def run_git(repo: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(repo), *args],
        check=True,
        capture_output=True,
        text=True,
    )
    return result.stdout.strip()


def commit_file(repo: Path, relative_path: str, content: bytes, subject: str) -> str:
    path = repo / relative_path
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)
    run_git(repo, "add", relative_path)
    run_git(repo, "commit", "-m", subject)
    return run_git(repo, "rev-parse", "HEAD")


def test_commit_list_defaults_to_commits_that_changed_xlsx(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    run_git(repo, "init", "-b", "main")
    run_git(repo, "config", "user.name", "Test User")
    run_git(repo, "config", "user.email", "test@example.com")

    commit_file(repo, "notes.txt", b"not an excel change", "update notes")
    excel_commit = commit_file(repo, "Data/Balancing.XLSX", b"xlsx bytes", "update balance")

    repository = GitRepository(repo)
    commits = repository.list_commits()

    assert repository.current_branch == "main"
    assert [commit.commit_id for commit in commits] == [excel_commit]
    assert commits[0].subject == "update balance"
    assert commits[0].author_name == "Test User"
    assert len(commits[0].parent_ids) == 1


def test_modified_xlsx_can_be_read_at_parent_and_current_commit(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    run_git(repo, "init", "-b", "main")
    run_git(repo, "config", "user.name", "Test User")
    run_git(repo, "config", "user.email", "test@example.com")
    commit_file(repo, "Data/items.xlsx", b"old workbook", "add workbook")
    current_commit = commit_file(repo, "Data/items.xlsx", b"new workbook", "edit workbook")

    changes = GitRepository(repo).list_changed_excel_files(current_commit)

    assert len(changes) == 1
    assert changes[0].change_type == "modified"
    assert changes[0].display_path == "Data/items.xlsx"
    assert GitRepository(repo).read_versions(changes[0]) == (b"old workbook", b"new workbook")


def test_empty_repository_has_a_branch_and_no_commits(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    run_git(repo, "init", "-b", "main")

    repository = GitRepository(repo)

    assert repository.current_branch == "main"
    assert repository.list_commits() == []


def test_renamed_xlsx_reads_old_and_new_paths(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    run_git(repo, "init", "-b", "main")
    run_git(repo, "config", "user.name", "Test User")
    run_git(repo, "config", "user.email", "test@example.com")
    commit_file(repo, "Data/old.xlsx", b"same workbook", "add workbook")
    run_git(repo, "mv", "Data/old.xlsx", "Data/new.xlsx")
    run_git(repo, "commit", "-m", "rename workbook")
    rename_commit = run_git(repo, "rev-parse", "HEAD")

    repository = GitRepository(repo)
    [change] = repository.list_changed_excel_files(rename_commit)

    assert (change.change_type, change.old_path, change.new_path) == (
        "renamed",
        "Data/old.xlsx",
        "Data/new.xlsx",
    )
    assert repository.read_versions(change) == (b"same workbook", b"same workbook")
