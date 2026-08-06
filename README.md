# Excel Git Viewer

[![CI](https://github.com/xiaoyunhai0/ExcelGitViewer/actions/workflows/ci.yml/badge.svg)](https://github.com/xiaoyunhai0/ExcelGitViewer/actions/workflows/ci.yml)
[![Release](https://img.shields.io/github/v/release/xiaoyunhai0/ExcelGitViewer)](https://github.com/xiaoyunhai0/ExcelGitViewer/releases/latest)

Excel Git Viewer 是面向程序开发人员的 Windows 只读审查工具。它直接读取本地 Git
历史，展示策划在普通提交中对 `.xlsx` 工作簿造成的单元格变化。

工具不会切换分支、检出文件或修改仓库，也不要求项目使用 Git LFS。它用于补足 Git
客户端只能看到“Excel 文件已变化”，却看不到具体改动位置的问题。

## 下载与使用

从 [GitHub Releases](https://github.com/xiaoyunhai0/ExcelGitViewer/releases/latest)
下载最新的 `ExcelGitViewer-*-windows-x64.zip`。解压完整目录后运行
`ExcelGitViewer.exe`，无需安装 Python。

运行条件：

- Windows 10/11 x64；
- 系统 `PATH` 中可以直接运行 `git`；
- 目标仓库已由 Git、Sourcetree 或其他客户端拉取到本地。

基本流程：

1. 点击“选择仓库”，选择本地 Git 项目目录。
2. 在提交表中选择一次普通提交。
3. 选择该提交中发生变化的 Excel 文件。
4. 在差异表和新旧上下文中检查具体改动。

“刷新”会强制重新读取当前仓库的提交历史。正常启动会自动恢复上次仓库，并在 HEAD
未变化时直接使用本地历史缓存。

## 审查能力

- 表格显示提交哈希、说明、作者、时间以及发生变化的 Excel 文件。
- 识别 `.xlsx` 文件的新增、修改、删除和重命名。
- 显示工作表变化、单元格新增、删除、修改及新旧值。
- 比较公式文本但不执行公式，保留并提示空格、制表符和换行。
- 标记隐藏工作表、隐藏行和隐藏列中的变化。
- 在新旧上下文中定位当前改动格：旧值为红色，新值为绿色，行列标题为黄色。
- 支持左右和上下拖动各审查区域，并在下次启动时恢复布局比例。

当前版本完成了“Git 提交浏览与单元格差异”。图片新增、删除、替换、移动、缩放及
预览属于后续里程碑，尚未实现。

## 性能机制

- 历史读取限制在当前分支最近 200 条普通提交，界面最多显示 200 条。
- 加载提交历史时不扫描每条提交的文件路径；选中提交后才查询其 `.xlsx` 变化。
- 使用约 384 MB 上限的进程内 LRU 工作簿快照缓存，加速重复和相邻文件对比。
- 工作簿在后台解析；切换选择或取消任务后，过期结果不会写回界面。
- 提交历史缓存会持久化到磁盘，仓库 HEAD 变化时自动失效。

## 本地数据与重置

Windows 提交历史缓存位于：

```text
%LOCALAPPDATA%\ExcelGitViewer\history
```

上次仓库和界面布局由 Qt 保存在以下注册表位置：

```text
HKEY_CURRENT_USER\Software\ExcelGitViewer\ExcelGitViewer
```

通常不需要手动清理。历史缓存损坏或与当前 HEAD 不匹配时，程序会自动重新读取 Git。
需要完全重置时，可在退出程序后删除上述缓存目录和注册表项。

## 支持范围

工具只处理 `.xlsx`，只比较普通提交与其唯一父提交。根提交按空工作簿处理，合并提交、
未提交改动、任意两次提交对比和 `.xls` 均不在当前范围内。

当前比较单元格值和公式文本，不比较颜色、边框、字体、图表、形状、批注、宏或外部链接
的执行结果。Microsoft Excel 固定样本已进入自动测试；WPS 固定样本仍需补充人工验收。

完整约束和验收标准见
[开发文档](Excel_Git_Viewer_MVP_%E5%BC%80%E5%8F%91%E6%96%87%E6%A1%A3.md)。

## 本地开发

项目使用 Python 3.12 和 [uv](https://docs.astral.sh/uv/)：

```bash
uv sync --locked
uv run ruff check .
uv run ruff format --check .
uv run mypy
uv run pytest
uv run excel-git-viewer
```

Linux 缺少 Qt 系统动态库时，窗口测试会跳过；核心 Git 与 Excel 测试仍可运行。Windows
CI 会运行真实 Qt 窗口检查。

测试用工作簿必须只包含合成数据，统一放在 `tests/fixtures/excel/`。不要提交真实项目的
策划表、导出的敏感数据或其他业务文件。

## 发布

推送到 `main` 或创建 Pull Request 时，GitHub Actions 会在 Linux 和 Windows 上运行
Ruff、mypy、pytest，并在 Windows 上检查 Qt 窗口。

推送 `v*` 标签后，Windows Runner 使用 PyInstaller 构建目录模式 EXE，运行打包后烟雾
测试，生成 ZIP 和 SHA-256 文件，并创建 GitHub Release。

## 安全与许可

程序只通过 Git 对象读取历史文件，不执行宏、公式或外部链接。工作簿打开前会检查 ZIP
条目数量和解压尺寸。项目使用 [MIT License](LICENSE)。
