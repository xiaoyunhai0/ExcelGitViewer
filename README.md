# Excel Git Viewer

<p align="center">
  <strong>English</strong> | <a href="README.zh-CN.md">简体中文</a>
</p>

[![CI](https://github.com/xiaoyunhai0/ExcelGitViewer/actions/workflows/ci.yml/badge.svg)](https://github.com/xiaoyunhai0/ExcelGitViewer/actions/workflows/ci.yml)
[![Release](https://img.shields.io/github/v/release/xiaoyunhai0/ExcelGitViewer)](https://github.com/xiaoyunhai0/ExcelGitViewer/releases/latest)
[![License](https://img.shields.io/github/license/xiaoyunhai0/ExcelGitViewer)](LICENSE)

Excel Git Viewer is a read-only Windows review tool for developers. It reads local Git history and
shows the cell-level changes made to `.xlsx` workbooks in non-merge commits.

The application never checks out files, switches branches, or modifies the repository. Git LFS is
not required. It fills the gap between seeing that an Excel file changed and knowing exactly what
changed inside it.

## Download and Use

Download the latest `ExcelGitViewer-*-windows-x64.zip` from
[GitHub Releases](https://github.com/xiaoyunhai0/ExcelGitViewer/releases/latest). Extract the entire
archive and run `ExcelGitViewer.exe`. Python is not required.

Requirements:

- Windows 10 or 11, x64;
- `git` available on the system `PATH`;
- a local repository cloned with Git, Sourcetree, or another Git client;
- the current desktop interface uses Simplified Chinese.

Basic workflow:

1. Select a local Git repository.
2. Select a non-merge commit from the commit table.
3. Select an Excel file changed by that commit.
4. Review its cell differences and the surrounding before/after context.

The application automatically reopens the last repository. When its HEAD is unchanged, commit
history is restored from the local cache. Use **Refresh** to force a new Git history scan.

## Review Features

- Displays commit hash, subject, author, time, and changed Excel files in compact tables.
- Detects added, modified, deleted, and renamed `.xlsx` files.
- Shows worksheet changes and added, deleted, or modified cell values.
- Compares formula text without executing formulas.
- Preserves and flags spaces, tabs, newlines, hidden worksheets, hidden rows, and hidden columns.
- Highlights the selected old cell in red, the new cell in green, and its row and column in amber.
- Provides six dockable panels that can be rearranged, tabbed, resized, or floated.
- Restores the panel layout after restart and provides a **Reset Layout** action.
- Allows every table column to be resized; double-clicking a divider fits it to its content.
- Persists column widths for the commit, file, worksheet, and cell-difference tables.

The current milestone covers Git commit browsing and cell differences. Embedded image additions,
deletions, replacements, movement, resizing, and previews are planned but not implemented.

## Performance

- Reads at most the latest 200 non-merge commits from the current branch.
- Does not scan changed paths while loading history; `.xlsx` paths are queried after commit selection.
- Uses an in-process LRU workbook snapshot cache with an approximate 384 MB limit.
- Parses workbooks in background tasks and discards stale results after selection changes.
- Persists commit metadata to disk and invalidates it automatically when the repository HEAD changes.

## Local Data and Reset

On Windows, the commit history cache is stored at:

```text
%LOCALAPPDATA%\ExcelGitViewer\history
```

The last repository, dock layout, and table widths are stored by Qt under:

```text
HKEY_CURRENT_USER\Software\ExcelGitViewer\ExcelGitViewer
```

Manual cleanup is normally unnecessary. Invalid or stale history caches are rebuilt automatically.
For a complete reset, close the application and remove the cache directory and registry key above.

## Supported Scope

The application handles `.xlsx` files and compares each non-merge commit with its single parent. Root
commits use an empty workbook as the old version. Merge commits, working-tree changes, arbitrary
commit pairs, and legacy `.xls` files are outside the current scope.

It compares cell values and formula text. Formatting, charts, shapes, comments, macros, and evaluated
external links are not compared. A Microsoft Excel fixture is covered by automated tests; a fixed WPS
fixture and manual WPS acceptance are still pending.

See the [development specification](docs/development-spec.md) for complete constraints and acceptance
criteria.

## Development

The project uses Python 3.12 and [uv](https://docs.astral.sh/uv/):

```bash
uv sync --locked
uv run ruff check .
uv run ruff format --check .
uv run mypy
uv run pytest
uv run excel-git-viewer
```

Window tests are skipped when the Linux host does not provide the required Qt system libraries. The
Windows CI job runs the real Qt window checks.

Workbook fixtures must contain synthetic data only and belong under `tests/fixtures/excel/`. Never
commit production workbooks, exported business data, or other sensitive files.

## Release Process

Pushes to `main` and pull requests run Ruff, mypy, and pytest on Linux and Windows. Windows also runs
a real Qt window smoke test.

Pushing a `v*` tag builds the directory-mode executable with PyInstaller, tests the packaged EXE, and
publishes a ZIP archive with its SHA-256 checksum to GitHub Releases.

## Security and License

Excel Git Viewer reads workbook bytes from Git objects without executing macros, formulas, or external
links. ZIP entry counts and extracted sizes are checked before a workbook is opened.

Licensed under the [MIT License](LICENSE).
