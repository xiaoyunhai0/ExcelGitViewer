# Excel Git Viewer 开发规格

> 面向需要审查策划表改动的程序开发人员，提供只读 Git 历史中的 Excel 单元格与内嵌图片差异。

- 文档版本：0.8
- 产品阶段：MVP
- 目标平台：Windows 10/11
- 文件范围：`.xlsx`
- 核心原则：只查看，不修改 Excel，不执行宏，不改变 Git 仓库状态
- 产品定位：提交审查辅助工具，不承诺理解 Excel 中的业务语义，不等同于完整视觉差异
- 实施状态：里程碑 1 已进入可运行实现；WPS 人工样本验收与里程碑 2 图片差异尚未完成

---

## 1. 项目背景

游戏策划经常使用 Excel 编写任务、活动、关卡、奖励以及图文说明。Excel 文件在 Git 中属于二进制文件，Sourcetree 可以显示“文件发生了变化”，但无法展示具体哪些单元格或图片被修改。

本工具只解决一个问题：

> 导入本地 Git 仓库，选择任意一次提交，查看该提交相对父提交造成的 Excel 内容变化。

工具不替代 Sourcetree。Sourcetree继续负责提交、拉取、推送、分支和合并，本工具只负责读取 Git 历史并显示 Excel 差异。

---

## 2. MVP 范围

### 2.1 必须实现

1. 导入本地 Git 仓库。
2. 显示最近的提交记录。
3. 有界读取当前分支最近 200 个普通提交的元数据，不按文件类型筛选提交。
4. 点击提交后，与该普通提交的唯一父提交比较。
5. 列出该提交中新增、修改、删除或重命名的 `.xlsx` 文件。
6. 查看工作表变化。
7. 查看单元格值和公式文本变化，保留空格、制表符和换行。
8. 显示差异单元格周边的只读上下文，并标记隐藏工作表、行和列。
9. 查看 Excel 内嵌图片变化：
   - 新增图片；
   - 删除图片；
   - 替换图片；
   - 移动图片；
   - 缩放图片。
10. 显示图片修改前和修改后的预览。
11. 同时支持由 Microsoft Excel 和 WPS 生成的标准 `.xlsx`。
12. 全程只读，不切换分支，不检出文件，不修改工作区。

### 2.2 明确不做

为保证短时间落地，MVP 不实现以下能力：

- Excel 编辑和保存；
- Git 提交、推送、拉取、合并；
- 冲突解决；
- 两次任意提交的手动对比；
- 合并提交差异；
- 其他本地分支或远程分支浏览；
- 扫描超过最近 200 个普通提交的分支历史；
- 工作区未提交改动对比；
- `.xls` 老格式；
- VBA 宏差异；
- 图表、形状、SmartArt、批注和音视频对象差异；
- 单元格颜色、边框、字体等格式差异；
- 图片内容的 AI 语义识别；
- Excel 内由公式或外部链接动态生成的图片；
- 完整替代 Sourcetree。

---

## 3. 用户操作流程

```text
启动程序
  ↓
选择本地 Git 仓库目录
  ↓
读取并显示提交列表
  ↓
点击某次提交
  ↓
显示该提交中发生变化的 Excel 文件
  ↓
点击 Excel 文件
  ↓
解析父提交版本和当前提交版本
  ↓
显示单元格改动与图片改动
```

### 3.1 主界面布局

```text
┌────────────────────────────────────────────────────────────────────┐
│ 仓库：D:\GameProject\DesignData        分支：develop       刷新   │
├───────────────────────────┬────────────────────────────────────────┤
│ 提交记录（表格）          │ 工作表概览                             │
│ 哈希 / 说明 / 作者 / 时间 ├────────────────────────────────────────┤
│ a82… / 修改活动 / 张三…  │ 差异记录：位置 / 旧值 / 新值          │
├───────────────────────────┼────────────────────┬───────────────────┤
│ Excel 文件（表格）        │ 修改前上下文       │ 修改后上下文      │
│ M / 活动策划.xlsx         │                    │                   │
└───────────────────────────┴────────────────────┴───────────────────┘
```

主工作区采用高密度开发审查风格：中性浅色背景、蓝色选中态、旧值红色语义、
新值绿色语义、当前坐标黄色定位。提交、文件和差异信息优先使用可扫描的表格，
不使用大卡片。

提交记录、Excel 文件、工作表概览、差异记录、修改前上下文和修改后上下文是六个独立
停靠面板。每个面板可拖到工作区上、下、左、右重新组合，可与其他面板形成页签，也可
拖出成为浮动窗口。面板不可被误关闭，边界可拖动缩放；退出时保存完整停靠状态，下次
启动自动恢复。“恢复默认布局”将所有面板和固定结构表格列宽恢复为推荐排列。同一记录
内切换列不应重复读取文件或解析工作簿。

所有表格的列宽必须允许用户拖动调整，双击列边界自动适配当前内容。提交说明和新旧值
默认获得较大宽度，但不得使用 `Fixed`、`Stretch` 或 `ResizeToContents` 锁定用户操作。
提交、Excel 文件、工作表和差异表的列宽在退出后保存，下次启动自动恢复。

### 3.2 单元格差异页面

显示字段：

| 字段 | 说明 |
|---|---|
| 工作表 | 单元格所在工作表 |
| 位置 | 如 `B12` |
| 类型 | 新增、删除、修改 |
| 旧值 | 父提交中的值或公式 |
| 新值 | 当前提交中的值或公式 |

### 3.3 图片差异页面

显示字段：

| 字段 | 说明 |
|---|---|
| 工作表 | 图片所在工作表 |
| 类型 | 新增、删除、替换、移动、缩放 |
| 旧位置 | 修改前的锚点位置 |
| 新位置 | 修改后的锚点位置 |
| 旧尺寸 | 修改前宽高 |
| 新尺寸 | 修改后宽高 |
| 旧图预览 | 父提交中的图片缩略图 |
| 新图预览 | 当前提交中的图片缩略图 |

图片预览建议使用左右并排方式：

```text
类型：替换
工作表：活动说明
位置：B5 附近

修改前                    修改后
┌─────────────────┐      ┌─────────────────┐
│     旧图片       │      │     新图片       │
└─────────────────┘      └─────────────────┘
```

---

## 4. 技术选型

### 4.1 桌面技术

| 模块 | 技术 |
|---|---|
| 编程语言 | Python 3.12 或团队统一版本 |
| 桌面界面 | PySide6 |
| Git 读取 | 系统 Git CLI，通过 `subprocess` 调用 |
| 单元格解析 | openpyxl |
| 图片和 OOXML 解析 | Python `zipfile` + `lxml` |
| 图片读取与缩略图 | Pillow |
| EXE 打包 | PyInstaller |
| 哈希 | Python `hashlib.sha256` |

### 4.2 为什么图片不只依赖 openpyxl

openpyxl 适合读取单元格值和公式，也提供图片插入能力；但本项目需要稳定地获取图片二进制、DrawingML 锚点、尺寸和关系信息。MVP 应直接解析 `.xlsx` 内部 OOXML 包，而不是依赖 openpyxl 的私有属性。

`.xlsx` 本质上是 ZIP 包。工作表通过关系指向 Drawing 部件，Drawing 部件包含图片对象及 `oneCellAnchor`、`twoCellAnchor` 或 `absoluteAnchor`，图片对象再通过关系指向 `xl/media` 下的媒体文件。

---

## 5. 系统架构

```text
┌────────────────────────────────────┐
│ UI 层：PySide6                     │
│ 仓库选择 / 提交列表 / 文件列表    │
│ 单元格差异 / 图片差异 / 图片预览  │
└────────────────┬───────────────────┘
                 │
┌────────────────▼───────────────────┐
│ Application Service               │
│ 加载提交 / 加载文件 / 触发比较    │
└───────────┬────────────┬───────────┘
            │            │
┌───────────▼──────┐ ┌───▼────────────────┐
│ Git Service      │ │ Excel Diff Service │
│ 日志、父提交、   │ │ 工作表、单元格、   │
│ 文件状态、Blob   │ │ 图片解析与比较     │
└──────────────────┘ └────────────────────┘
```

### 5.1 建议工程目录

```text
excel-git-viewer/
├── app.py
├── requirements.txt
├── README.md
├── README.zh-CN.md
├── src/
│   ├── ui/
│   │   ├── main_window.py
│   │   ├── commit_panel.py
│   │   ├── file_panel.py
│   │   ├── cell_diff_panel.py
│   │   └── image_diff_panel.py
│   ├── git/
│   │   ├── repository.py
│   │   ├── history.py
│   │   └── blob_reader.py
│   ├── excel/
│   │   ├── workbook_reader.py
│   │   ├── cell_reader.py
│   │   ├── image_reader.py
│   │   ├── cell_comparer.py
│   │   └── image_comparer.py
│   ├── models/
│   │   ├── commit.py
│   │   ├── changed_file.py
│   │   ├── cell_change.py
│   │   └── image_change.py
│   └── common/
│       ├── errors.py
│       ├── paths.py
│       └── temp_files.py
└── tests/
    ├── fixtures/
    ├── test_git_repository.py
    ├── test_cell_diff.py
    └── test_image_diff.py
```

---

## 6. Git 数据读取设计

### 6.1 仓库校验

选择目录后执行：

```bash
git -C <repo> rev-parse --show-toplevel
```

成功则保存仓库根目录；失败则提示“所选目录不是有效 Git 仓库”。

### 6.2 获取提交列表

一次只读取当前 `HEAD` 可达的最近 200 个普通提交元数据。使用空字符和固定数量的提交字段解析，避免提交信息、中文和空格影响结果；历史加载不读取任何提交的变更文件路径：

```bash
git -C <repo> log -n 200 --no-merges \
  --pretty=format:%x00%H%x00%P%x00%an%x00%ae%x00%aI%x00%s%x00 \
  -z
```

模块从这一份有界结果显示最近 200 条普通提交。界面不提供“仅 Excel 提交”筛选，用户选择提交后才读取该提交中的 `.xlsx` 文件，避免历史加载和筛选状态影响操作响应。

提交数据模型：

```python
@dataclass(frozen=True)
class CommitInfo:
    commit_id: str
    parent_ids: tuple[str, ...]
    author_name: str
    author_email: str
    authored_at: datetime
    subject: str
```

### 6.3 获取本次提交的文件变化

```bash
git -C <repo> diff-tree \
  --root --no-commit-id --name-status -r -M -z <commit> \
  -- ':(icase,glob)**/*.xlsx'
```

只保留扩展名为 `.xlsx` 的文件。

需要识别：

- `A`：新增；
- `M`：修改；
- `D`：删除；
- `Rxxx`：重命名。

### 6.4 获取父提交

- 普通提交：使用唯一父提交；
- 合并提交：MVP 不在提交列表中显示；
- 根提交：旧版本视为空工作簿。

### 6.5 读取指定版本中的 Excel

推荐从 Git 对象直接读取二进制，不检出到工作区：

```bash
git -C <repo> show <commit>:<path>
```

程序必须使用参数数组调用 `subprocess`，禁止拼接 shell 字符串：

```python
subprocess.run(
    ["git", "-C", repo, "show", f"{commit}:{path}"],
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    check=True,
)
```

对于重命名文件，旧版本使用旧路径，新版本使用新路径。

---

## 7. Excel 单元格解析与比较

### 7.1 读取规则

使用 openpyxl：

```python
load_workbook(
    filename=stream,
    read_only=False,
    data_only=False,
    keep_links=False,
)
```

`data_only=False` 用于读取公式本身，而不是上次计算缓存值。

### 7.2 标准化单元格

每个非空单元格转换为：

```python
@dataclass(frozen=True)
class CellValue:
    sheet_name: str
    coordinate: str
    data_type: str
    normalized_value: str
```

标准化规则：

- `None` 统一为空；
- 日期和时间转为 ISO 8601；
- 布尔值转为 `true` / `false`；
- 数字使用稳定字符串表示；
- 公式保留开头的 `=`；
- 错误值保留原始错误字符串；
- 字符串不执行 `strip` 或其他空白归一化；
- 不比较样式。

公式单元格只比较公式文本。MVP 不执行公式，也不比较可能过期或缺失的缓存计算结果。

### 7.3 比较规则

以 `(sheet_name, coordinate)` 为键比较旧、新版本：

- 旧无、新有：新增；
- 旧有、新无：删除；
- 旧值与新值不同：修改；
- 完全相同：不显示。

数据模型：

```python
@dataclass(frozen=True)
class CellChange:
    sheet_name: str
    coordinate: str
    change_type: Literal["added", "deleted", "modified"]
    old_value: str | None
    new_value: str | None
```

### 7.4 已知限制

MVP 按单元格坐标比较。如果在表格中间插入整行，后续内容位置发生移动，可能出现大量差异。短期版本接受该行为，不实现按“任务 ID”等业务主键进行语义匹配。

---

## 8. Excel 图片解析设计

### 8.1 支持对象

MVP 仅支持普通工作表 DrawingML 中的内嵌位图，例如：

- PNG；
- JPEG；
- GIF；
- BMP；
- TIFF（能否预览取决于 Pillow 支持情况）。

不支持：

- 图表；
- 形状；
- SmartArt；
- VML 旧式对象；
- 单元格内的现代 IMAGE 数据类型或网络图片；
- 外部链接图片；
- OLE 对象。

### 8.2 OOXML 读取链路

```text
xl/workbook.xml
  ↓ 工作表 r:id
xl/_rels/workbook.xml.rels
  ↓ 工作表关系
xl/worksheets/sheetN.xml
  ↓ <drawing r:id="...">
xl/worksheets/_rels/sheetN.xml.rels
  ↓ Target="../drawings/drawingN.xml"
xl/drawings/drawingN.xml
  ↓ xdr:pic + 锚点
xl/drawings/_rels/drawingN.xml.rels
  ↓ Target="../media/imageN.png"
xl/media/imageN.png
```

### 8.3 需要解析的锚点

Drawing 部件可能包含：

- `xdr:oneCellAnchor`；
- `xdr:twoCellAnchor`；
- `xdr:absoluteAnchor`。

MVP 统一转换为标准位置对象：

```python
@dataclass(frozen=True)
class ImageAnchor:
    anchor_type: Literal["one_cell", "two_cell", "absolute"]
    from_col: int | None
    from_row: int | None
    from_col_offset: int | None
    from_row_offset: int | None
    to_col: int | None
    to_row: int | None
    to_col_offset: int | None
    to_row_offset: int | None
    x: int | None
    y: int | None
    width: int
    height: int
```

坐标和尺寸保留 OOXML 使用的 EMU 整数，界面显示时再转换为像素或“B5 附近”。为避免 Excel 或 WPS 重新保存造成噪声，位置以起始单元格变化为主；宽或高变化超过 `2 px` 或 `1%` 时才报告缩放，取两个阈值中先达到者。

### 8.4 图片标准数据模型

```python
@dataclass(frozen=True)
class EmbeddedImage:
    sheet_name: str
    object_name: str | None
    description: str | None
    media_path: str
    content_type: str
    content_hash: str
    pixel_width: int | None
    pixel_height: int | None
    anchor: ImageAnchor
    raw_bytes: bytes
```

其中：

```python
content_hash = sha256(raw_bytes).hexdigest()
```

哈希用于判断图片内容是否一致。

### 8.5 图片差异类型

```python
ImageChangeType = Literal[
    "added",
    "deleted",
    "replaced",
    "moved",
    "resized",
    "moved_and_resized",
]
```

数据模型：

```python
@dataclass(frozen=True)
class ImageChange:
    sheet_name: str
    change_type: ImageChangeType
    old_image: EmbeddedImage | None
    new_image: EmbeddedImage | None
```

---

## 9. 图片匹配和差异算法

图片对象的内部 ID、文件名和关系 ID 在 Excel 保存后可能变化，不能单独依赖 `image1.png`、`rId2` 或对象 ID 判断图片身份。

MVP 使用以下顺序匹配。

### 9.1 第一步：内容和位置都相同

条件：

- 工作表相同；
- 图片 SHA-256 相同；
- 锚点和尺寸相同。

结果：图片未变化，不显示。

### 9.2 第二步：内容相同，位置或尺寸不同

条件：

- 工作表相同；
- 图片 SHA-256 相同；
- 起始单元格不同，或尺寸变化超过容差。

结果：

- 仅位置不同：`moved`；
- 仅尺寸不同：`resized`；
- 两者都不同：`moved_and_resized`。

如果同一工作表有多张内容完全相同的图片，按锚点距离最近原则一一配对。

### 9.3 第三步：位置相同，内容不同

条件：

- 工作表相同；
- 主锚点位置相同或足够接近；
- 图片 SHA-256 不同。

结果：`replaced`。

“足够接近”在 MVP 中定义为：起始单元格行列相同。偏移量变化不影响替换判断。

### 9.4 第四步：无法配对

- 仅旧版本存在：`deleted`；
- 仅新版本存在：`added`。

### 9.5 匹配伪代码

```python
def compare_images(old_images, new_images):
    changes = []

    # 1. 完全相同
    match_exact_content_and_anchor()

    # 2. 内容相同，位置或尺寸变化
    match_same_hash_by_nearest_anchor()

    # 3. 锚点相同，内容变化
    match_same_start_cell_as_replacement()

    # 4. 剩余对象为新增或删除
    emit_deleted_for_unmatched_old()
    emit_added_for_unmatched_new()

    return changes
```

### 9.6 预览生成

- 使用 Pillow 从 `raw_bytes` 读取图片；
- 生成最大边 320 像素的缩略图；
- 保持原始宽高比；
- 不覆盖原始媒体；
- 解码失败时显示文件类型、字节大小和“无法预览”，差异记录仍然保留。

---

## 10. 工作表级变化

工作表需要显示以下状态：

- 新增工作表；
- 删除工作表；
- 已存在工作表中的单元格变化数量；
- 已存在工作表中的图片变化数量。

MVP 不自动识别工作表重命名。重命名会表现为一个工作表删除和另一个工作表新增。

示例：

```text
任务配置       单元格 12    图片 0
活动说明       单元格 3     图片 2
旧活动说明     已删除
新活动说明     已新增
```

---

## 11. UI 详细要求

### 11.1 仓库栏

- “选择仓库”按钮；
- 仓库路径；
- 当前分支；
- “刷新”按钮；
- 错误状态提示。

### 11.2 提交列表

提交记录使用四列表格，每行显示：

- 短哈希；
- 提交说明；
- 作者；
- 时间；
- 普通提交状态。

默认只读取当前分支最近 200 条普通提交的元数据。提交列表不按 Excel 变更筛选；点击提交后再读取其中变化的 `.xlsx` 文件。

### 11.3 Excel 文件列表

Excel 文件使用两列表格显示状态和路径，状态包括：

- `A` 新增；
- `M` 修改；
- `D` 删除；
- `R` 重命名。

非 Excel 文件不显示。

### 11.4 差异区域

里程碑 1 使用工作表概览、单元格差异记录和新旧上下文三个可上下缩放的区域。
里程碑 2 加入图片能力后可使用两个页签：

```text
[单元格改动 15] [图片改动 3]
```

单元格页签使用表格显示。

选中一条单元格变化时，在下方显示该坐标周边少量行列的只读上下文。修改前目标格
使用红色强调，修改后目标格使用绿色强调；对应行号与列号使用黄色强调，空值目标格也
必须可见。位于隐藏工作表、隐藏行或隐藏列的变化必须显示标记。字符串变化涉及空格、
制表符或换行时，界面提供可见提示。

Git 显示文件已修改、但没有检测到受支持的单元格或图片变化时，显示：“Git 文件已变化，但未检测到当前支持范围内的内容差异；可能修改了格式、图表、批注或其他未支持对象。”

图片页签使用列表或卡片显示，每项包含：

- 修改类型；
- 工作表；
- 位置变化；
- 尺寸变化；
- 旧图和新图预览。

### 11.5 加载状态

解析大文件必须在后台线程执行，界面显示：

```text
正在读取 Git 文件……
正在解析工作簿……
正在比较 12 张图片……
```

主窗口不得冻结。
用户必须能取消长时间解析；快速切换提交或文件时，过期任务结果不得覆盖当前选择。

取消采用协作式语义：界面立即恢复可操作并使当前任务结果失效；解析器在工作表、行和比较边界检查取消令牌。openpyxl 正在执行的单次包加载无法在线程内安全硬中断，但完成后的结果必须丢弃，不得回写 UI。

---

## 12. 核心模块接口

核心模块应将 Git 命令解析、Excel 读取和标准化细节隐藏在小接口后。UI 和测试使用同一公开接口，不依赖 openpyxl 或 Git 输出格式。

### 12.1 Git 服务

```python
class GitRepository:
    def __init__(self, path: Path) -> None: ...

    root: Path
    current_branch: str | None

    def load_recent_history(
        self,
        scan_limit: int = 200,
        display_limit: int = 200,
    ) -> CommitHistory: ...

    def list_changed_excel_files(
        self,
        commit_id: str,
    ) -> list[ChangedFile]: ...

    def read_versions(
        self,
        change: ChangedFile,
    ) -> tuple[bytes | None, bytes | None]: ...
```

接口只接受完整 Git 对象 ID。新增文件的旧版本和删除文件的新版本以 `None` 表示；重命名的旧、新路径选择由模块内部处理。

### 12.2 工作簿差异模块

```python
class WorkbookDiffer:
    def compare(
        self,
        old_bytes: bytes | None,
        new_bytes: bytes | None,
        *,
        cancellation: CancellationToken | None = None,
    ) -> WorkbookDiff: ...
```

`WorkbookDiff` 统一返回工作表、单元格以及里程碑 2 的图片差异。模块内部在解析前执行 ZIP 安全预检，并在工作表和行边界检查取消令牌。

---

## 13. 性能与缓存

### 13.1 性能目标

在 Windows 10/11 x64 普通办公电脑上，验收目标为：

- 200 条提交列表在 `2 秒`内首次显示；
- 普通工作簿差异在 `5 秒`内显示；
- 20 MB、1 万个非空单元格和 100 张内嵌图片的上限样本在 `15 秒`内完成；
- 解析期间 UI 保持可操作。
- 任务超过 `2 秒`时显示进度状态并可取消。

性能目标是工程验收基线，由固定样本和测试机配置记录复现，不是强实时保证。

### 13.2 提交历史缓存

程序在用户缓存目录中持久化最近一次有界扫描得到的提交元数据，使 EXE 退出后再次启动时
不必重复读取同一份 Git 历史。缓存文件使用 JSON 格式和原子替换写入，单文件读取上限为
`5 MB`。

```text
cache key = normalized_repository_path
cache validity = repository_path + HEAD + scan_limit + display_limit + format_version
```

缓存只保存提交 ID、父提交 ID、作者、时间和标题，不保存变更路径或 Excel 工作簿内容。
以下情况必须重新读取 Git 并覆盖缓存：

- 仓库 HEAD 变化；
- 扫描上限或显示上限变化；
- 缓存格式版本变化；
- 缓存缺失、损坏、过大或字段不合法；
- 用户点击“刷新”，此时强制绕过缓存。

程序记住最后一次成功打开的仓库路径并在下次启动时自动恢复。路径已失效时清除该设置，
不阻止用户重新选择仓库。

### 13.3 工作簿快照缓存

解析后的工作簿单元格快照按内容 SHA-256 放入进程内 LRU 缓存，近似内存预算为 `384 MB`。
相同工作簿在相邻提交或重复选择中再次出现时，直接复用快照，不重复执行 openpyxl 解压和
解析。缓存只存在于当前 EXE 进程，退出后由操作系统释放；不写入磁盘。超过预算时优先淘汰
最久未使用的快照。

---

## 14. 安全要求

Excel 和 Git 仓库都应视为不可信输入。

1. 不启动 Excel，不使用 COM 自动化。
2. 不执行宏、公式、外部链接或嵌入对象。
3. Git 命令通过参数数组调用，不使用 `shell=True`。
4. 解压前检查 ZIP 条目数量、单条目大小和总解压大小，防止 ZIP Bomb。
5. 禁止将 ZIP 条目直接按原路径写入磁盘，避免路径穿越。
6. 图片解码设置像素数量上限，防止超大图片耗尽内存。
7. 临时文件放入专用临时目录，使用后清理。
8. 仓库读取失败、Excel 损坏或图片解码失败时，只显示错误，不影响仓库内容。

---

## 15. 异常处理

| 场景 | 行为 |
|---|---|
| 目录不是 Git 仓库 | 阻止导入并显示提示 |
| 未安装 Git | 提示安装 Git 或配置 Git 路径 |
| 提交中的 Excel 损坏 | 文件列表显示“解析失败” |
| Excel 有密码保护或加密 | 显示“不支持加密工作簿” |
| 图片格式无法预览 | 显示元数据，保留差异记录 |
| 根提交 | 旧版本视为空 |
| 合并提交 | 不在 MVP 提交列表中显示 |
| 文件被删除 | 新版本视为空 |
| 文件被新增 | 旧版本视为空 |
| 文件被重命名 | 使用旧路径读取父提交，使用新路径读取当前提交 |
| Git LFS 指针文件 | 检测到非 ZIP 内容时提示“可能需要先拉取 LFS 对象” |

---

## 16. 开发顺序

### 阶段 A：Git 历史闭环

- 选择仓库；
- 加载提交列表；
- 加载提交中的 Excel 文件；
- 从两个提交中读取 Excel 二进制。

完成标准：能从任意普通提交取到旧、新两个 `.xlsx` 字节流。

### 阶段 B：单元格差异

- 解析工作表；
- 标准化单元格；
- 生成新增、删除和修改列表；
- 在界面展示。

完成标准：修改文本、数字和公式后能准确显示旧值和新值。

### 里程碑 1 交付

- 交付可运行的 Git 提交浏览和单元格差异版本；
- 先在真实工作流中验证只读性、差异可读性和 Excel/WPS 兼容性。

### 阶段 C：图片差异

- 解析 DrawingML 和媒体关系；
- 提取图片二进制和锚点；
- 计算哈希；
- 判断新增、删除、替换、移动和缩放；
- 显示左右预览。

完成标准：五种图片变化均能被正确识别。

### 里程碑 2 交付

- 接入图片解析、匹配和预览；
- 通过 Excel 与 WPS 人工保存样本的回归测试后，完成 MVP 验收。

### 阶段 D：稳定性

- 后台线程；
- 错误提示；
- 缓存；
- 打包 EXE；
- 测试真实策划仓库。

---

## 17. 测试设计

### 17.1 Git 测试

1. 普通提交；
2. 根提交；
3. 合并提交被过滤；
4. 中文文件名；
5. 路径包含空格；
6. Excel 新增；
7. Excel 删除；
8. Excel 重命名；
9. Git LFS 指针文件；
10. 相同仓库和 HEAD 在新进程中命中持久化提交历史缓存；
11. 仓库 HEAD 变化后自动重新读取 Git 历史；
12. 强制刷新绕过已有缓存并重新读取 Git 历史。

### 17.2 单元格测试

1. 字符串修改；
2. 数字修改；
3. 日期修改；
4. 公式修改；
5. 空单元格变为有值；
6. 有值变为空；
7. 新增工作表；
8. 删除工作表；
9. 中文和换行文本；
10. 错误公式值。
11. 公式文本变化但不计算结果；
12. 前后空格、制表符和换行变化；
13. 隐藏工作表、隐藏行和隐藏列中的变化；
14. Microsoft Excel 与 WPS 人工保存样本。

### 17.3 图片测试

1. 新增一张 PNG；
2. 删除一张图片；
3. 同一位置替换为另一张图片；
4. 图片从 `B5` 移动到 `F10`；
5. 图片宽高改变；
6. 图片同时移动和缩放；
7. 同一工作表存在两张相同图片；
8. 不同工作表存在相同图片；
9. 图片内部文件名变化但内容未变化；
10. 图片关系 ID 变化但内容未变化；
11. 图片无法被 Pillow 解码；
12. 超大图片触发安全限制。

### 17.4 UI 测试

1. 快速切换提交不会崩溃；
2. 快速切换 Excel 文件不会显示上一任务结果；
3. 大文件解析期间界面不冻结；
4. 旧图或新图不存在时布局正常；
5. 图片预览保持宽高比；
6. 错误信息可读且不遮挡主界面；
7. 超过 2 秒的任务可取消；
8. 快速切换时过期任务不会覆盖当前结果；
9. 提交列表没有“仅 Excel 提交”筛选，加载历史时不读取变更文件路径；
10. 启动时自动恢复上次成功打开的仓库；
11. 点击“刷新”强制重新读取 Git 历史；
12. 点击提交时只按 `.xlsx` 路径读取该提交的文件变化。
13. 选中差异记录时，新旧上下文明确标出目标单元格以及对应行号、列号。
14. 六个审查面板均可拖动停靠、组合为页签、浮动和缩放，退出后恢复布局。
15. 所有表格列宽可拖动，双击边界可自动适配；固定结构表格在重启后恢复列宽。
16. “恢复默认布局”能撤销用户的面板排列和固定结构表格列宽设置。

---

## 18. MVP 验收标准

必须全部通过：

1. 能导入一个有效的本地 Git 仓库。
2. 能读取并显示当前分支最近 200 条普通提交的元数据，历史加载不读取变更文件路径。
3. 点击提交后能列出该提交变化的 `.xlsx` 文件。
4. 能比较当前普通提交和它的唯一父提交。
5. 能显示单元格新增、删除和修改。
6. 能显示公式旧值和新值。
7. 能显示图片新增和删除。
8. 能显示图片替换，并提供旧图与新图预览。
9. 能显示图片移动、缩放以及同时移动和缩放。
10. 中文文件名、中文工作表名和中文内容显示正常。
11. 不修改工作区、暂存区、分支或提交历史。
12. 损坏 Excel 和无法解码的图片不会导致程序崩溃。
13. 解析过程中桌面界面不会长时间无响应。
14. 可打包为 Windows EXE，并在未安装 Python 的测试机器运行。
15. Microsoft Excel 和 WPS 保存的固定样本均通过回归测试。
16. 自动恢复上次仓库；HEAD 未变化时从本地缓存恢复提交列表，HEAD 变化或点击“刷新”时重新读取 Git。
17. 提交和 Excel 文件以表格显示；当前差异坐标在新旧上下文中有明确定位高亮。
18. 六个审查面板支持自由停靠、页签组合、浮动和缩放，并在重启后恢复布局。
19. 所有表格允许手动调整列宽，提交、文件、工作表和差异表在重启后恢复列宽。
20. 用户可一键恢复推荐面板布局和默认列宽。

---

## 19. 最终交付物

MVP 最终应包含：

```text
ExcelGitViewer-v<version>-windows-x64.zip
ExcelGitViewer-v<version>-windows-x64.zip.sha256
README.md
README.zh-CN.md
LICENSE (MIT)
sample-repository/ 或测试数据说明
```

README 至少包含：

- 运行环境；
- 如何选择仓库；
- 如何选择提交并查看 Excel 差异；
- 当前支持和不支持的内容；
- 常见错误处理；
- Git LFS 注意事项。

### 19.1 GitHub 发布

- 公开仓库不得包含真实项目数据、仓库路径、提交信息或未清理的 Office 作者元数据；
- 每次推送到 `main` 时由 GitHub Actions 运行测试；
- 推送 `v*` 标签时，在 Windows Runner 上使用 Python 3.12 和 PyInstaller 目录模式构建；
- 构建产物压缩后发布到 GitHub Releases，并附带 SHA-256 校验文件；
- MVP 只支持 Windows 10/11 x64。

---

## 20. 技术风险

### 20.1 Excel 保存会重写内部关系

即使图片未变化，Excel 也可能重新编号媒体文件和关系 ID。因此比较必须基于图片内容哈希和锚点，不依赖内部文件名或关系 ID。

### 20.2 重复图片匹配

同一工作表中多次插入相同图片时，哈希相同。MVP 使用“最近锚点”配对，极端情况下可能产生不符合用户直觉的匹配结果，但不会丢失新增或删除数量。

### 20.3 图片与单元格不是同一种对象

图片通常浮动在单元格网格上，并由 DrawingML 锚点定位。界面中的“B5 附近”只是可读表示，不能将图片简单当成 `B5` 单元格的值。

### 20.4 插入行造成大量单元格变化

MVP 按坐标比较，插入整行可能产生大量单元格差异。本阶段不实现基于业务 ID 的行匹配。

### 20.5 协作式取消不等于终止线程

openpyxl 的工作簿包加载是同步调用，无法在进程内安全强制终止。里程碑 1 保证取消后 UI 立即恢复、过期结果绝不回写，并在可检查边界尽快停止。如真实上限样本证明包加载本身造成长时间 CPU 或内存占用，后续版本需将解析移入可终止的独立子进程。

---

## 21. 参考资料

- Git `git-cat-file` 官方文档：<https://git-scm.com/docs/git-cat-file>
- Git `git-diff-tree` 官方文档：<https://git-scm.com/docs/git-diff-tree>
- Git `git-show` 官方文档：<https://git-scm.com/docs/git-show>
- openpyxl 图片文档：<https://openpyxl.readthedocs.io/en/3.1/images.html>
- Microsoft Open XML Spreadsheet Drawing：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.worksheetdrawing>
- Microsoft Open XML Spreadsheet Picture：<https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.picture>
