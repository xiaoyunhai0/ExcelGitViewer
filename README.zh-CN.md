<div align="center">
  <h1>Excel Git Viewer</h1>
  <p><strong>在 Git 中审查 Excel 改动，直接定位到具体单元格。</strong></p>
  <p><a href="README.md">English</a> | <strong>简体中文</strong></p>
  <p>
    <a href="https://github.com/xiaoyunhai0/ExcelGitViewer/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/xiaoyunhai0/ExcelGitViewer/actions/workflows/ci.yml/badge.svg"></a>
    <a href="https://github.com/xiaoyunhai0/ExcelGitViewer/releases/latest"><img alt="Release" src="https://img.shields.io/github/v/release/xiaoyunhai0/ExcelGitViewer"></a>
    <a href="LICENSE"><img alt="License" src="https://img.shields.io/github/license/xiaoyunhai0/ExcelGitViewer"></a>
  </p>
  <p>
    <a href="https://github.com/xiaoyunhai0/ExcelGitViewer/releases/latest"><strong>下载 Windows 版本</strong></a>
    · <a href="docs/development-spec.md">开发规格</a>
  </p>
</div>

Excel Git Viewer 是面向程序开发人员的只读桌面工具，用于审查本地 Git 历史中的 `.xlsx`
改动。它不只告诉你“二进制文件发生变化”，还会显示工作簿内部具体改了什么。

## 核心能力

| Git 提交审查 | 单元格证据 | 自由工作区 |
|---|---|---|
| 浏览最近 200 条普通提交，不检出文件、不切换分支。 | 对比新旧值、公式、坐标和周边上下文。 | 六个面板可排列、组合页签、缩放或浮动，重启后恢复。 |

| 重复审查更快 | 全程只读 | 布局自动保存 |
|---|---|---|
| 复用持久化提交记录和进程内工作簿快照缓存。 | 不修改仓库，不执行公式、宏或外部链接。 | 所有表格列宽可调整、双击自动适配，也可一键恢复完整布局。 |

## 快速开始

**运行条件：** Windows 10/11 x64、系统 `PATH` 中可以运行 `git`，并且仓库已拉取到本地。
使用打包版本不需要安装 Python，也不要求项目使用 Git LFS。

1. 从 [GitHub Releases](https://github.com/xiaoyunhai0/ExcelGitViewer/releases/latest)
   下载并解压最新的 `ExcelGitViewer-*-windows-x64.zip`。
2. 运行 `ExcelGitViewer.exe`，选择本地 Git 仓库。
3. 选择提交和发生变化的 Excel 文件，检查具体差异。

```text
仓库 -> 提交 -> 变化的工作簿 -> 单元格差异 -> 修改前后上下文
```

程序会自动恢复上次仓库。仓库 HEAD 未变化时，提交记录直接使用本地缓存；点击“刷新”
可以强制重新读取 Git 历史。

## 审查内容

| 范围 | 当前行为 |
|---|---|
| 提交 | 读取当前分支最近 200 条普通提交。 |
| Excel 文件 | 识别 `.xlsx` 文件的新增、修改、删除和重命名。 |
| 工作表 | 报告工作表新增、删除和修改。 |
| 单元格 | 显示值和公式文本的新增、删除与修改。 |
| 上下文 | 旧值标红、新值标绿，对应行列标题标黄。 |
| 风险提示 | 标记隐藏工作表、隐藏行列和不可见空白。 |

内嵌图片差异属于后续里程碑，目前尚未实现。

## 工作区

- 拖动任意面板标题，可停靠到工作区上、下、左、右。
- 面板可以组合成页签，也可以拖出程序成为浮动窗口。
- 面板边界和所有表格列宽均可调整，双击列边界自动适配内容。
- 启动时恢复上次工作区，也可通过“恢复默认布局”回到推荐排列。

## 性能机制

| 机制 | 作用 |
|---|---|
| 有界提交记录 | 最多加载 200 条，选中提交后才查询变化文件路径。 |
| 持久化历史缓存 | 仓库 HEAD 未变化时直接复用提交记录。 |
| 384 MB 快照缓存 | 加速同一次运行中的重复和相邻工作簿比较。 |
| 后台任务 | 避免界面冻结，并在切换选择后丢弃过期结果。 |

## 支持范围

| 当前支持 | 暂不支持 |
|---|---|
| 普通提交中的 `.xlsx` 文件 | 合并提交和任意两次提交对比 |
| 单元格值和公式文本 | 格式、图表、形状和批注 |
| 根提交与空工作簿比较 | 工作区和未提交改动 |
| Microsoft Excel 自动化样本 | `.xls` 和已完成的 WPS 验收 |

完整约束和验收标准见 [开发规格](docs/development-spec.md)。

<details>
<summary><strong>本地数据与完全重置</strong></summary>

提交历史缓存：

```text
%LOCALAPPDATA%\ExcelGitViewer\history
```

上次仓库、停靠布局和表格列宽：

```text
HKEY_CURRENT_USER\Software\ExcelGitViewer\ExcelGitViewer
```

无效的历史缓存会自动重建。需要完全重置时，请先退出程序，再删除上述目录和注册表项。

</details>

<details>
<summary><strong>本地开发与发布</strong></summary>

项目使用 Python 3.12 和 [uv](https://docs.astral.sh/uv/)：

```bash
uv sync --locked
uv run ruff check .
uv run ruff format --check .
uv run mypy
uv run pytest
uv run excel-git-viewer
```

推送和 Pull Request 会在 Linux、Windows 上运行检查。推送 `v*` 标签后，工作流会构建并
测试 Windows EXE，然后发布 ZIP 和 SHA-256 文件。

测试工作簿必须只包含合成数据，并统一放在 `tests/fixtures/excel/`。

</details>

## 安全与许可

程序直接读取 Git 对象中的工作簿字节，并在打开前检查 ZIP 条目数量和解压尺寸。程序
不会执行宏、公式或外部链接。

Excel Git Viewer 使用 [MIT License](LICENSE)。
