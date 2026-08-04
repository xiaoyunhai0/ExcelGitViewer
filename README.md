# Excel Git Viewer

Excel Git Viewer 是一个面向程序开发人员的 Windows 只读工具，用于查看策划在
Git 普通提交中对 `.xlsx` 工作簿造成的内容变化。

当前仓库正在实现 MVP 里程碑 1：Git 提交浏览与单元格差异。图片差异属于里程碑 2。

## 当前能力

- 选择本地 Git 仓库，不切换分支或检出文件。
- 默认显示当前分支中改动过 `.xlsx` 的普通提交。
- 识别 Excel 文件的新增、修改、删除和重命名。
- 显示单元格新增、删除、修改以及新旧值。
- 比较公式文本，不执行公式。
- 标记隐藏工作表、隐藏行列和不可见空白字符。
- 显示差异单元格周边的新旧只读上下文。
- 在后台解析工作簿，支持取消并丢弃过期任务结果。

## 运行条件

- Windows 10/11 x64。
- 系统 `PATH` 中可以直接运行 `git`。
- 使用 GitHub Release 压缩包时不需要安装 Python。

## 使用方法

1. 从 GitHub Releases 下载最新的 `ExcelGitViewer-*-windows-x64.zip`。
2. 解压完整目录，运行 `ExcelGitViewer.exe`。
3. 点击“选择仓库”，选择已经拉取到本地的 Git 项目。
4. 选择普通提交和 Excel 文件，查看单元格差异及上下文。

## 本地开发

项目使用 Python 3.12 和 [uv](https://docs.astral.sh/uv/)：

```bash
uv sync
uv run pytest
uv run ruff check .
uv run ruff format --check .
uv run mypy
uv run excel-git-viewer
```

当前 Linux 宿主需要 Qt 的系统动态库才能启动窗口。核心 Git/Excel 测试不依赖 Qt。

## 发布

- 推送到 `main` 或创建 Pull Request 时，GitHub Actions 在 Linux 和 Windows 上运行检查。
- 推送 `v*` 标签时，Windows Runner 使用 PyInstaller 生成目录模式压缩包，
  计算 SHA-256，然后创建 GitHub Release。

## 安全和范围

工具只通过 Git 对象读取历史文件，不执行 Excel、宏、公式或外部链接。工作簿在打开前会检查
ZIP 条目数量和解压尺寸。完整范围、限制与验收标准见
[`Excel_Git_Viewer_MVP_开发文档.md`](Excel_Git_Viewer_MVP_%E5%BC%80%E5%8F%91%E6%96%87%E6%A1%A3.md)。

## License

[MIT](LICENSE)
