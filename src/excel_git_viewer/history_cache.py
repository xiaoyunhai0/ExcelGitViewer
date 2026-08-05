from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime
from pathlib import Path
from string import hexdigits
from tempfile import NamedTemporaryFile
from typing import Any

from excel_git_viewer.models import CommitHistory, CommitInfo


class HistoryCache:
    """Persist bounded commit history without storing workbook content."""

    FORMAT_VERSION = 3
    MAX_CACHE_BYTES = 5 * 1024 * 1024

    def __init__(self, root: Path) -> None:
        self.root = root

    @classmethod
    def default(cls) -> HistoryCache:
        if local_app_data := os.environ.get("LOCALAPPDATA"):
            base = Path(local_app_data)
        else:
            base = Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache"))
        return cls(base / "ExcelGitViewer" / "history")

    def load(
        self,
        repo_root: Path,
        head_id: str,
        scan_limit: int,
        display_limit: int,
    ) -> CommitHistory | None:
        path = self._path_for(repo_root)
        try:
            if path.stat().st_size > self.MAX_CACHE_BYTES:
                return None
            payload = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(payload, dict) or not self._matches(
                payload,
                repo_root,
                head_id,
                scan_limit,
                display_limit,
            ):
                return None
            scanned_commit_count = payload["scanned_commit_count"]
            if (
                type(scanned_commit_count) is not int
                or scanned_commit_count < 0
                or scanned_commit_count > scan_limit
            ):
                return None
            return CommitHistory(
                all_commits=self._read_commits(payload["all_commits"], display_limit),
                scanned_commit_count=scanned_commit_count,
                scan_limit=scan_limit,
                source="cache",
            )
        except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError):
            return None

    def save(
        self,
        repo_root: Path,
        head_id: str,
        display_limit: int,
        history: CommitHistory,
    ) -> None:
        self.root.mkdir(parents=True, exist_ok=True)
        destination = self._path_for(repo_root)
        payload = {
            "format_version": self.FORMAT_VERSION,
            "repo_root": self._normalize_repo_root(repo_root),
            "head_id": head_id,
            "scan_limit": history.scan_limit,
            "display_limit": display_limit,
            "scanned_commit_count": history.scanned_commit_count,
            "all_commits": [self._write_commit(commit) for commit in history.all_commits],
        }
        temp_path: Path | None = None
        try:
            with NamedTemporaryFile(
                "w",
                encoding="utf-8",
                dir=self.root,
                prefix="history-",
                suffix=".tmp",
                delete=False,
            ) as temp_file:
                temp_path = Path(temp_file.name)
                json.dump(payload, temp_file, ensure_ascii=False, separators=(",", ":"))
                temp_file.flush()
                os.fsync(temp_file.fileno())
            temp_path.replace(destination)
        except OSError:
            return
        finally:
            if temp_path is not None:
                temp_path.unlink(missing_ok=True)

    def _path_for(self, repo_root: Path) -> Path:
        normalized = self._normalize_repo_root(repo_root).encode("utf-8")
        digest = hashlib.sha256(normalized).hexdigest()
        return self.root / f"{digest}.json"

    @staticmethod
    def _normalize_repo_root(repo_root: Path) -> str:
        return os.path.normcase(str(repo_root.resolve()))

    def _matches(
        self,
        payload: dict[str, Any],
        repo_root: Path,
        head_id: str,
        scan_limit: int,
        display_limit: int,
    ) -> bool:
        return (
            type(payload.get("format_version")) is int
            and payload["format_version"] == self.FORMAT_VERSION
            and payload.get("repo_root") == self._normalize_repo_root(repo_root)
            and payload.get("head_id") == head_id
            and type(payload.get("scan_limit")) is int
            and payload["scan_limit"] == scan_limit
            and type(payload.get("display_limit")) is int
            and payload["display_limit"] == display_limit
        )

    @staticmethod
    def _write_commit(commit: CommitInfo) -> dict[str, object]:
        return {
            "commit_id": commit.commit_id,
            "parent_ids": list(commit.parent_ids),
            "author_name": commit.author_name,
            "author_email": commit.author_email,
            "authored_at": commit.authored_at.isoformat(),
            "subject": commit.subject,
        }

    @staticmethod
    def _read_commits(raw_commits: object, display_limit: int) -> tuple[CommitInfo, ...]:
        if not isinstance(raw_commits, list) or len(raw_commits) > display_limit:
            raise ValueError("cached commits must be a list")
        commits: list[CommitInfo] = []
        for raw_commit in raw_commits:
            if not isinstance(raw_commit, dict):
                raise ValueError("cached commit must be an object")
            commit_id = raw_commit["commit_id"]
            parent_ids = raw_commit["parent_ids"]
            author_name = raw_commit["author_name"]
            author_email = raw_commit["author_email"]
            authored_at = raw_commit["authored_at"]
            subject = raw_commit["subject"]
            if not HistoryCache._is_object_id(commit_id):
                raise ValueError("cached commit ID must be a full object ID")
            if (
                not isinstance(parent_ids, list)
                or len(parent_ids) > 1
                or not all(HistoryCache._is_object_id(parent_id) for parent_id in parent_ids)
            ):
                raise ValueError("cached parent IDs must be full object IDs")
            if not all(isinstance(value, str) for value in (author_name, author_email, subject)):
                raise ValueError("cached commit text fields must be strings")
            if not isinstance(authored_at, str):
                raise ValueError("cached authored time must be a string")
            parsed_authored_at = datetime.fromisoformat(authored_at)
            if parsed_authored_at.tzinfo is None:
                raise ValueError("cached authored time must include a timezone")
            commits.append(
                CommitInfo(
                    commit_id=commit_id,
                    parent_ids=tuple(parent_ids),
                    author_name=author_name,
                    author_email=author_email,
                    authored_at=parsed_authored_at,
                    subject=subject,
                )
            )
        return tuple(commits)

    @staticmethod
    def _is_object_id(value: object) -> bool:
        return (
            isinstance(value, str)
            and len(value) in {40, 64}
            and all(char in hexdigits for char in value)
        )
