from __future__ import annotations

import subprocess
from datetime import datetime
from pathlib import Path
from string import hexdigits

from excel_git_viewer.history_cache import HistoryCache
from excel_git_viewer.models import ChangedFile, CommitHistory, CommitInfo, FileChangeType


class GitRepositoryError(RuntimeError):
    """Raised when a repository cannot satisfy a read request."""


class GitRepository:
    """Read-only access to the Git information needed by the viewer."""

    def __init__(self, path: Path, *, history_cache: HistoryCache | None = None) -> None:
        requested_path = path.expanduser().resolve()
        root = self._run_at(requested_path, "rev-parse", "--show-toplevel").decode().strip()
        self.root = Path(root)
        branch = self._try_run("symbolic-ref", "--quiet", "--short", "HEAD")
        self.current_branch: str | None = branch.decode().strip() if branch is not None else None
        self._history_cache = history_cache

    def load_recent_history(
        self,
        scan_limit: int = 1000,
        display_limit: int = 200,
        *,
        force_refresh: bool = False,
    ) -> CommitHistory:
        if scan_limit < 1:
            raise ValueError("scan_limit must be at least 1")
        if display_limit < 1:
            raise ValueError("display_limit must be at least 1")
        head_output = self._try_run("rev-parse", "--verify", "HEAD")
        if head_output is None:
            return CommitHistory((), (), 0, scan_limit)
        head_id = head_output.decode().strip()
        if self._history_cache is not None and not force_refresh:
            cached = self._history_cache.load(
                self.root,
                head_id,
                scan_limit,
                display_limit,
            )
            if cached is not None:
                return cached

        output = self._run(
            "log",
            "--no-merges",
            f"-n{scan_limit}",
            "--pretty=format:%x00%H%x00%P%x00%an%x00%ae%x00%aI%x00%s%x00",
            "--name-status",
            "-z",
        )
        records = self._parse_history(output)
        all_commits = tuple(record[0] for record in records[:display_limit])
        excel_commits = tuple(
            commit for commit, paths in records if any(self._is_excel_path(path) for path in paths)
        )[:display_limit]
        history = CommitHistory(
            all_commits=all_commits,
            excel_commits=excel_commits,
            scanned_commit_count=len(records),
            scan_limit=scan_limit,
        )
        if self._history_cache is not None:
            self._history_cache.save(self.root, head_id, display_limit, history)
        return history

    def list_changed_excel_files(self, commit_id: str) -> list[ChangedFile]:
        self._validate_object_id(commit_id)
        parent_fields = self._run("rev-list", "--parents", "-n", "1", commit_id).decode().split()
        if not parent_fields or parent_fields[0] != commit_id:
            raise GitRepositoryError("Commit could not be resolved")
        if len(parent_fields) > 2:
            raise GitRepositoryError("Merge commits are not supported")
        parent_id = parent_fields[1] if len(parent_fields) == 2 else None

        output = self._run(
            "diff-tree",
            "--root",
            "--no-commit-id",
            "--name-status",
            "-r",
            "-M",
            "-z",
            commit_id,
        )
        return self._parse_changed_files(output, commit_id, parent_id)

    def read_versions(self, change: ChangedFile) -> tuple[bytes | None, bytes | None]:
        old_bytes = self._read_file(change.parent_id, change.old_path)
        new_bytes = self._read_file(change.commit_id, change.new_path)
        return old_bytes, new_bytes

    def _run(self, *args: str) -> bytes:
        return self._run_at(self.root, *args)

    def _try_run(self, *args: str) -> bytes | None:
        try:
            return self._run(*args)
        except GitRepositoryError:
            return None

    @staticmethod
    def _run_at(path: Path, *args: str) -> bytes:
        try:
            result = subprocess.run(
                ["git", "-C", str(path), *args],
                check=True,
                capture_output=True,
            )
        except FileNotFoundError as error:
            raise GitRepositoryError("Git is not installed or is not available on PATH") from error
        except subprocess.CalledProcessError as error:
            message = error.stderr.decode("utf-8", errors="replace").strip()
            raise GitRepositoryError(message or "Git command failed") from error
        return result.stdout

    def _read_file(self, commit_id: str | None, repo_path: str | None) -> bytes | None:
        if commit_id is None or repo_path is None:
            return None
        self._validate_object_id(commit_id)
        return self._run("show", f"{commit_id}:{repo_path}")

    @staticmethod
    def _validate_object_id(object_id: str) -> None:
        if len(object_id) not in {40, 64} or any(char not in hexdigits for char in object_id):
            raise ValueError("object ID must be a full hexadecimal Git object ID")

    @staticmethod
    def _parse_history(output: bytes) -> list[tuple[CommitInfo, tuple[str, ...]]]:
        records: list[tuple[CommitInfo, tuple[str, ...]]] = []
        fields = output.split(b"\x00")
        index = 0
        while index < len(fields):
            while index < len(fields) and fields[index] == b"":
                index += 1
            if index >= len(fields):
                break
            if index + 5 >= len(fields):
                raise GitRepositoryError("Git returned an unexpected history record")
            commit_id, parents, name, email, authored_at, subject = (
                field.decode("utf-8", errors="replace") for field in fields[index : index + 6]
            )
            index += 6
            paths: list[str] = []
            while index < len(fields) and fields[index] != b"":
                raw_status = fields[index]
                if raw_status.startswith(b"\n"):
                    raw_status = raw_status[1:]
                status = raw_status.decode("ascii", errors="replace")
                if not status or status[0] not in {"A", "B", "C", "D", "M", "R", "T", "U", "X"}:
                    raise GitRepositoryError("Git returned an unexpected history change")
                path_count = 2 if status[0] in {"C", "R"} else 1
                if index + path_count >= len(fields):
                    raise GitRepositoryError("Git returned an incomplete history change")
                for offset in range(1, path_count + 1):
                    paths.append(fields[index + offset].decode("utf-8", errors="replace"))
                index += path_count + 1
            records.append(
                (
                    CommitInfo(
                        commit_id=commit_id,
                        parent_ids=tuple(parents.split()),
                        author_name=name,
                        author_email=email,
                        authored_at=datetime.fromisoformat(authored_at),
                        subject=subject,
                    ),
                    tuple(paths),
                )
            )
        return records

    @staticmethod
    def _parse_changed_files(
        output: bytes,
        commit_id: str,
        parent_id: str | None,
    ) -> list[ChangedFile]:
        fields = output.split(b"\x00")
        if fields and fields[-1] == b"":
            fields.pop()

        changes: list[ChangedFile] = []
        index = 0
        while index < len(fields):
            status = fields[index].decode("ascii", errors="replace")
            index += 1
            old_path: str | None
            new_path: str | None
            if status.startswith("R"):
                if index + 1 >= len(fields):
                    raise GitRepositoryError("Git returned an incomplete rename record")
                old_path = fields[index].decode("utf-8", errors="replace")
                new_path = fields[index + 1].decode("utf-8", errors="replace")
                index += 2
                change_type: FileChangeType = "renamed"
            else:
                if index >= len(fields):
                    raise GitRepositoryError("Git returned an incomplete file record")
                path = fields[index].decode("utf-8", errors="replace")
                index += 1
                type_by_status: dict[str, FileChangeType] = {
                    "A": "added",
                    "M": "modified",
                    "D": "deleted",
                }
                if status not in type_by_status:
                    continue
                change_type = type_by_status[status]
                old_path = None if change_type == "added" else path
                new_path = None if change_type == "deleted" else path

            if not GitRepository._is_excel_path(old_path) and not GitRepository._is_excel_path(
                new_path
            ):
                continue
            changes.append(
                ChangedFile(
                    commit_id=commit_id,
                    parent_id=parent_id,
                    change_type=change_type,
                    old_path=old_path,
                    new_path=new_path,
                )
            )
        return changes

    @staticmethod
    def _is_excel_path(path: str | None) -> bool:
        return path is not None and path.casefold().endswith(".xlsx")
