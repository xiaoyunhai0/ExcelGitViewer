<div align="center">
  <h1>Excel Git Viewer</h1>
  <p><strong>Review Excel changes in Git, down to the exact cell.</strong></p>
  <p><strong>English</strong> | <a href="README.zh-CN.md">简体中文</a></p>
  <p>
    <a href="https://github.com/xiaoyunhai0/ExcelGitViewer/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/xiaoyunhai0/ExcelGitViewer/actions/workflows/ci.yml/badge.svg"></a>
    <a href="https://github.com/xiaoyunhai0/ExcelGitViewer/releases/latest"><img alt="Release" src="https://img.shields.io/github/v/release/xiaoyunhai0/ExcelGitViewer"></a>
    <a href="LICENSE"><img alt="License" src="https://img.shields.io/github/license/xiaoyunhai0/ExcelGitViewer"></a>
  </p>
  <p>
    <a href="https://github.com/xiaoyunhai0/ExcelGitViewer/releases/latest"><strong>Download for Windows</strong></a>
    · <a href="docs/development-spec.md">Development specification</a>
  </p>
</div>

Excel Git Viewer is a read-only desktop tool for developers reviewing `.xlsx` and `.xls` changes
in local Git history. It shows what changed inside a workbook instead of stopping at “binary file
modified.”

> [!NOTE]
> The current desktop interface uses Simplified Chinese. The documentation is available in English
> and Chinese.

## At a Glance

| Git-aware review | Cell-level evidence | Flexible workspace |
|---|---|---|
| Browse up to 200 non-merge commits without checkout or branch changes. | Compare old/new values, formulas, coordinates, and nearby context. | Rearrange, tab, resize, or float six panels; restore them after restart. |

| Fast repeat reviews | Read-only by design | Your layout, remembered |
|---|---|---|
| Reuse commit metadata and an in-memory workbook snapshot cache. | Never modify Git or execute formulas, macros, or external links. | Resize columns, auto-fit on double-click, or reset the workspace. |

## Quick Start

**Requirements:** Windows 10/11 x64, `git` on `PATH`, and a local Git repository. Python and Git LFS
are not required for the packaged application.

1. Download and extract the latest `ExcelGitViewer-*-windows-x64.zip` from
   [GitHub Releases](https://github.com/xiaoyunhai0/ExcelGitViewer/releases/latest).
2. Run `ExcelGitViewer.exe` and select a local repository.
3. Select a commit, choose a changed Excel file, and inspect its differences.

```text
Repository -> Commit -> Changed workbook -> Cell difference -> Before/after context
```

The last repository opens automatically. If its HEAD is unchanged, commit history loads from the
local cache. Use **Refresh** to force a new Git scan.

## What It Reviews

| Area | Current behavior |
|---|---|
| Commits | Reads up to the latest 200 non-merge commits from the current branch. |
| Excel files | Detects added, modified, deleted, and renamed `.xlsx` and `.xls` files. |
| Worksheets | Reports added, deleted, and changed worksheets. |
| Cells | Shows value changes, `.xlsx` formula text, and `.xls` saved formula results. |
| Context | Highlights the old cell in red, the new cell in green, and matching headers in amber. |
| Warnings | Flags hidden sheets, rows, columns, and invisible whitespace. |

Embedded image differences are planned for a later milestone and are not implemented yet.
For legacy `.xls` workbooks, formulas are compared by their last saved result because the format
reader does not expose formula source text. `.xlsx` workbooks continue to compare formula text.

## Workspace

- Drag any panel title to dock it on the top, bottom, left, or right.
- Combine panels into tabs or drag them outside the application as floating windows.
- Resize panel boundaries and every table column; double-click a divider to fit its content.
- Restore the previous workspace on startup or use **Reset Layout** to return to the default.

## Performance

| Mechanism | Benefit |
|---|---|
| Bounded history | Loads at most 200 commit records and delays changed-path queries until selection. |
| Persistent history cache | Reuses commit metadata while the repository HEAD is unchanged. |
| 384 MB snapshot cache | Speeds up repeated and adjacent workbook comparisons within one session. |
| Background tasks | Keeps the interface responsive and discards stale results after selection changes. |

## Scope

| Supported now | Outside the current scope |
|---|---|
| `.xlsx` and `.xls` files from non-merge commits | Merge commits and arbitrary commit pairs |
| Cell values and formula text | Formatting, charts, shapes, and comments |
| Root commits compared with an empty workbook | Working-tree and uncommitted changes |
| Microsoft Excel `.xlsx`/`.xls` compatibility | Completed WPS acceptance |

See the [development specification](docs/development-spec.md) for complete constraints and acceptance
criteria.

<details>
<summary><strong>Local data and full reset</strong></summary>

Commit history cache:

```text
%LOCALAPPDATA%\ExcelGitViewer\history
```

Last repository, dock layout, and table widths:

```text
HKEY_CURRENT_USER\Software\ExcelGitViewer\ExcelGitViewer
```

Invalid history caches are rebuilt automatically. For a complete reset, close the application and
remove the directory and registry key above.

</details>

<details>
<summary><strong>Development and release</strong></summary>

The project uses Python 3.12 and [uv](https://docs.astral.sh/uv/):

```bash
uv sync --locked
uv run ruff check .
uv run ruff format --check .
uv run mypy
uv run pytest
uv run excel-git-viewer
```

Pushes and pull requests run checks on Linux and Windows. A `v*` tag builds and tests the packaged
Windows EXE, then publishes a ZIP archive and SHA-256 checksum.

Test workbooks must contain synthetic data only and belong under `tests/fixtures/excel/`.

</details>

## Security and License

Workbook bytes are read directly from Git objects. ZIP entry counts and extracted sizes are checked
before opening a workbook. The application never executes macros, formulas, or external links.

Excel Git Viewer is available under the [MIT License](LICENSE).
