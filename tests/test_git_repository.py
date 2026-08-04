from __future__ import annotations

import json
import subprocess
from pathlib import Path

from excel_git_viewer.git_repository import GitRepository
from excel_git_viewer.history_cache import HistoryCache


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


def test_recent_history_identifies_commits_that_changed_xlsx(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    run_git(repo, "init", "-b", "main")
    run_git(repo, "config", "user.name", "Test User")
    run_git(repo, "config", "user.email", "test@example.com")

    commit_file(repo, "notes.txt", b"not an excel change", "update notes")
    excel_commit = commit_file(repo, "Data/Balancing.XLSX", b"xlsx bytes", "update balance")

    repository = GitRepository(repo)
    commits = repository.load_recent_history().excel_commits

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
    assert repository.load_recent_history().all_commits == ()


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


def test_recent_history_builds_all_and_excel_views_from_one_bounded_window(
    tmp_path: Path,
) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    run_git(repo, "init", "-b", "main")
    run_git(repo, "config", "user.name", "Test User")
    run_git(repo, "config", "user.email", "test@example.com")
    commit_file(repo, "Data/old.xlsx", b"old", "old excel change")
    commit_file(repo, "code.txt", b"one", "older code change")
    commit_file(repo, "Data/recent.xlsx", b"recent", "recent excel change")
    commit_file(repo, "code.txt", b"two", "latest code change")

    history = GitRepository(repo).load_recent_history(scan_limit=2, display_limit=2)

    assert history.scanned_commit_count == 2
    assert [commit.subject for commit in history.all_commits] == [
        "latest code change",
        "recent excel change",
    ]
    assert [commit.subject for commit in history.excel_commits] == ["recent excel change"]


def test_recent_history_allows_the_old_record_marker_in_a_subject(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    run_git(repo, "init", "-b", "main")
    run_git(repo, "config", "user.name", "Test User")
    run_git(repo, "config", "user.email", "test@example.com")
    commit_file(repo, "Data/items.xlsx", b"workbook", "EGV-COMMIT")

    history = GitRepository(repo).load_recent_history()

    assert [commit.subject for commit in history.excel_commits] == ["EGV-COMMIT"]


def test_recent_history_is_reused_by_a_new_repository_instance(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    run_git(repo, "init", "-b", "main")
    run_git(repo, "config", "user.name", "Test User")
    run_git(repo, "config", "user.email", "test@example.com")
    commit_file(repo, "Data/items.xlsx", b"workbook", "add workbook")
    cache = HistoryCache(tmp_path / "cache")

    first = GitRepository(repo, history_cache=cache).load_recent_history()
    second = GitRepository(repo, history_cache=cache).load_recent_history()

    assert first.source == "git"
    assert second.source == "cache"
    assert second.excel_commits == first.excel_commits


def test_recent_history_cache_is_invalidated_when_head_changes(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    run_git(repo, "init", "-b", "main")
    run_git(repo, "config", "user.name", "Test User")
    run_git(repo, "config", "user.email", "test@example.com")
    commit_file(repo, "Data/items.xlsx", b"workbook", "add workbook")
    cache = HistoryCache(tmp_path / "cache")
    GitRepository(repo, history_cache=cache).load_recent_history()
    commit_file(repo, "code.txt", b"new code", "new head")

    refreshed = GitRepository(repo, history_cache=cache).load_recent_history()

    assert refreshed.source == "git"
    assert refreshed.all_commits[0].subject == "new head"


def test_force_refresh_bypasses_recent_history_cache(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    run_git(repo, "init", "-b", "main")
    run_git(repo, "config", "user.name", "Test User")
    run_git(repo, "config", "user.email", "test@example.com")
    commit_file(repo, "Data/items.xlsx", b"workbook", "add workbook")
    cache = HistoryCache(tmp_path / "cache")
    GitRepository(repo, history_cache=cache).load_recent_history()

    refreshed = GitRepository(repo, history_cache=cache).load_recent_history(force_refresh=True)

    assert refreshed.source == "git"


def test_invalid_cached_commit_fields_fall_back_to_git(tmp_path: Path) -> None:
    repo = tmp_path / "repo"
    repo.mkdir()
    run_git(repo, "init", "-b", "main")
    run_git(repo, "config", "user.name", "Test User")
    run_git(repo, "config", "user.email", "test@example.com")
    commit_file(repo, "Data/items.xlsx", b"workbook", "add workbook")
    cache = HistoryCache(tmp_path / "cache")
    GitRepository(repo, history_cache=cache).load_recent_history()
    [cache_path] = cache.root.glob("*.json")
    payload = json.loads(cache_path.read_text(encoding="utf-8"))
    payload["all_commits"][0]["commit_id"] = None
    cache_path.write_text(json.dumps(payload), encoding="utf-8")

    history = GitRepository(repo, history_cache=cache).load_recent_history()

    assert history.source == "git"
