---
name: ppt-analyzer
description: 通过 WPS 手动导出幻灯片图片，再由 AI 读取图片进行内容分析，将 PPT 结构、要点和关键信息整理为 Markdown 文档。触发条件：用户要求"分析 PPT""学习 PPT""总结 PPT""提取 PPT 架构""根据 PPT 做..."等任何涉及 PPT 内容理解、信息提取、结构梳理的场景。不触发本技能的情况：仅要求将 PPT 导出为其他格式、或仅要求把分析结果写入 Wiki（Wiki 写入由 wiki 相关技能负责）。
---

# PPT 分析技能

本技能采用「人工导出 + AI 分析」的两段式流程：

1. **导出**：由用户在 WPS 中把 PPT/PDF 按页输出为 PNG 图片，存放到以文件名称命名的文件夹下的 `assets/` 目录中。
2. **建索引**：编写一个 Markdown 文件，按页码列出所有幻灯片图片的相对路径、演讲者备注和 AI 视觉识别概要。
3. **分析**：AI 读取 `assets/` 中的幻灯片图片，逐页（或分批）进行视觉分析，整理出结构化的 Markdown 分析报告。

这种流程避免了对 Office COM 组件或第三方库的依赖，兼容各种 Windows 环境，且图片质量由 WPS 直接保证。

> **支持格式**：本技能同时支持 `.pptx` 和 `.pdf` 格式的幻灯片分析。PDF 幻灯片通常不含演讲者备注，第七步会自动检测并生成补充讲稿。

### 参考文档

- **`references/sed-coordinate-positioning-multiline-replace-insert-cheatsheet.md`** — sed 命令在索引文件中的精确定位、替换和插入技巧。当 sed 可用时优先使用；不可用时退化到 Python 脚本。

---

## 前置条件

- Windows 操作系统
- 已安装 WPS Office（演示组件 + PDF 组件可用）
- 待分析的 `.pptx` 或 `.pdf` 文件可本地打开

---

## 第一步：创建分析目录

当用户说要分析某个 PPT 或 PDF 时，先以该文件的文件名（去掉 `.pptx` 或 `.pdf` 扩展名）在仓库根目录建立一个同名文件夹，并在其中创建 `assets/` 子目录。

### 目录结构示例

```
.
├── 综合感知服务（2022-2024年）-整体验收PPT20260429/
│   ├── assets/                  <-- 存放 WPS 导出的幻灯片图片
│   ├── assets/crops/            <-- 截取的子图（第六步产出）
│   ├── README.md                <-- 本 PPT 的分析说明（可选）
│   ├── 幻灯片索引.md             <-- 图片路径 + 备注 + 识别概要
│   ├── 备注提取.md               <-- 原始备注全文备份
│   └── 分析结果.md               <-- 整体分析结论（第七步产出）
```

---

## 第二步：使用 WPS 导出幻灯片图片

### 2.1 PPT 文件（.pptx）导出流程

1. 在资源管理器中找到目标 `.pptx` 文件，双击使用 **WPS 演示** 打开。
2. 点击顶部菜单栏：**文件 → 输出为图片**（部分版本为「文件 → 另存为 → 输出为图片」）。
3. 在弹出的对话框中设置以下选项：
   - **输出方式**：选择「逐页输出」（推荐，方便 AI 按页分析）。
   - **输出格式**：选择 `PNG`。
   - **输出尺寸**：建议选择「标清（默认尺寸 2 倍）」或更高，确保文字清晰可读。
   - **输出颜色**：彩色。
   - **输出目录**：选择「自定义位置」，然后点击浏览按钮，定位到第一步创建的 `assets/` 目录。
4. 点击「开始输出」，等待 WPS 将每一页幻灯片导出为一张独立的 PNG 图片。

> 提示：导出完成后，WPS 可能会弹出「已完成」提示，点击关闭即可。

### 2.2 PDF 文件（.pdf）导出流程

当用户提供的文件是 `.pdf` 格式时（例如从其他系统导出的幻灯片 PDF），WPS 演示无法直接打开，需要使用 **WPS PDF** 组件进行转换：

1. 在资源管理器中找到目标 `.pdf` 文件，双击使用 **WPS PDF** 打开。
2. 点击顶部菜单栏：**文件 → 导出 PDF 为图片**（部分版本为「转换 → PDF 转图片」）。
3. 在弹出的对话框中设置以下选项：
   - **输出格式**：选择 `PNG`。
   - **输出品质**：选择「高清」或「超清」，确保文字清晰可读。
   - **输出方式**：选择「逐页输出」。
   - **输出目录**：选择「自定义位置」，定位到第一步创建的 `assets/` 目录。
4. 点击「开始输出」或「转换」，等待 WPS 将每一页导出为独立的 PNG 图片。

> **PDF 与 PPT 的关键区别**：
> - PDF 文件**不含演讲者备注**，因此第四步提取备注时结果将全部为空（100% 空置率）。
> - 这属于正常情况，第七步会自动检测空置率并为所有页面生成补充讲稿备注。
> - PDF 导出后，直接跳过第四步（提取备注），从第五步（编写索引）继续执行。

### WPS 默认命名规则

WPS 通常按以下方式命名导出的图片：

- 若文件名较长，可能命名为：`原文件名_01.png`、`原文件名_02.png` …
- 若文件名较短，可能直接命名为：`幻灯片1.PNG`、`幻灯片2.PNG` …

无论哪种命名方式，文件名中通常都包含顺序编号，AI 在读取时应按编号排序处理。

---

## 第三步：检查导出结果并处理文件名

用户告知导出完成后，AI 应首先检查 `assets/` 目录下是否存在图片文件。

### 3.1 未找到文件时的处理

如果 `assets/` 目录为空，AI **必须主动询问用户**：

> 「`assets/` 目录下暂未发现图片文件，请确认：
> 1. 导出是否已完成？
> 2. 导出目录是否选择正确？当前期望路径为：`…/assets/`
> 3. 如果文件导出到了其他位置，请告诉我实际路径，我可以帮您移动过来。」

若用户提供了其他路径，AI 使用 Shell 命令将图片移动到 `assets/` 目录下。

### 3.2 文件名过长时的处理

如果 WPS 生成的文件名过长（例如 `原文件名_01.png` 超过 30 个字符），不便于阅读和索引，应统一重命名为 `幻灯片_XX.png` 格式。

**使用以下脚本自动重命名**（在 `assets/` 目录所在位置执行）：

```bash
cd assets
# 按文件名排序，用序号重新命名
ls -1 *.png *.PNG 2>/dev/null | sort -V | awk 'BEGIN{n=1} {printf "mv \"%s\" 幻灯片_%02d.png\n", $0, n++}' | bash
cd ..
```

重命名后，目录下应为：

```
assets/
├── 幻灯片_01.png
├── 幻灯片_02.png
├── 幻灯片_03.png
└── ...
```

> 提示：如果原文件名中已包含 `幻灯片` 字样且长度适中，可跳过重命名步骤。

---

## 第四步：提取演讲者备注

在生成索引之前，应先提取 PPT 中每页的**演讲者备注（Speaker Notes）**。备注中往往包含讲解要点、补充说明或汇报时的口述内容，对后续分析非常有价值。

> **PDF 文件跳过本步骤**：如果源文件是 `.pdf` 格式，直接跳过第四步，创建一个空的 `备注提取.md` 文件（仅包含页码标题和「备注：（无）」占位），然后从第五步继续。PDF 不含备注，第七步会自动检测并生成补充讲稿。

### 使用 python-pptx 提取备注

技能包内置了通用提取脚本 `scripts/extract_notes.py`，支持提取备注、文本内容、表格结构，并自动输出为 Markdown 格式。

**脚本位置**：`{skill-dir}/scripts/extract_notes.py`

**用法**：

```bash
python {skill-dir}/scripts/extract_notes.py <pptx路径> [输出md路径]
```

- 若省略输出路径，则打印到 stdout
- 输出格式兼容 `备注提取.md` 的规范

**脚本源码**（可直接复制使用）：

```python
#!/usr/bin/env python3
"""
通用 PPT 备注与文本提取脚本。
提取每页的演讲者备注和可见文本内容，输出为 Markdown 格式。
依赖：pip install python-pptx
"""
import sys
import os
from pptx import Presentation


def extract_notes_and_text(pptx_path, output_md_path=None):
    prs = Presentation(pptx_path)
    total = len(prs.slides)
    lines = []

    for i, slide in enumerate(prs.slides, start=1):
        # --- 提取备注 ---
        note_text = ""
        if slide.has_notes_slide:
            note_slide = slide.notes_slide
            text_frame = note_slide.notes_text_frame
            note_text = text_frame.text.strip()

        # --- 提取可见文本（去重、去空） ---
        visible_texts = []
        seen = set()
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = para.text.strip()
                    if t and t not in seen:
                        visible_texts.append(t)
                        seen.add(t)
            elif shape.has_table:
                for row in shape.table.rows:
                    row_texts = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                    if row_texts:
                        t = " | ".join(row_texts)
                        if t not in seen:
                            visible_texts.append(t)
                            seen.add(t)

        # --- 写入 Markdown ---
        lines.append(f"## 第 {i} 页")
        lines.append("")
        lines.append(f"**备注**：{note_text if note_text else '（无）'}")
        lines.append("")
        if visible_texts:
            lines.append("**页面文本**：")
            lines.append("")
            for t in visible_texts:
                lines.append(f"- {t}")
            lines.append("")
        lines.append("")

    md_content = "\n".join(lines)

    if output_md_path:
        os.makedirs(os.path.dirname(output_md_path) if os.path.dirname(output_md_path) else ".", exist_ok=True)
        with open(output_md_path, "w", encoding="utf-8") as f:
            f.write(md_content)
        print(f"已保存到: {output_md_path}  (共 {total} 页)")
    else:
        print(md_content)
        print(f"\n总共 {total} 页")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python extract_notes.py <pptx路径> [输出md路径]")
        sys.exit(1)
    pptx_path = sys.argv[1]
    output_md = sys.argv[2] if len(sys.argv) > 2 else None
    extract_notes_and_text(pptx_path, output_md)
```

**提取结果保存为 `备注提取.md`**，供后续写入索引时使用。

> 注意：若 PPT 由 WPS 创建，备注信息存储方式可能与标准 Office 略有差异。如果 `python-pptx` 读取不到备注，可尝试用 WPS 自带的「另存为」功能先保存为 `.pptx`（标准格式）后再提取。

### PDF 文件创建空备注文件

当源文件为 `.pdf` 时，无需运行 `python-pptx` 脚本，直接创建一个空的 `备注提取.md`：

```python
import os

def create_empty_notes(total_pages, output_path):
    """为 PDF 文件创建空的备注提取文件"""
    lines = []
    for i in range(1, total_pages + 1):
        lines.append(f"## 第 {i} 页")
        lines.append("")
        lines.append("**备注**：（无）")
        lines.append("")
        lines.append("")
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f"已创建空备注文件: {output_path} (共 {total_pages} 页)")

# 用法：根据实际页数创建
create_empty_notes(27, "备注提取.md")
```

---

## 第五步：编写幻灯片索引 Markdown

### ⚠️ 重要：写入顺序规划

**核心原则：避免"先写占位符，后回填"的模式。**

先写索引框架再回头替换占位符，容易因特殊字符（中文引号、换行等）导致 JSON 解析失败或替换错位，耗费大量工具调用。

**推荐做法（按优先级）：**

**方案 A：顺序追加生成（推荐）**
- 不提前写索引框架
- 收集完所有数据（备注 + 识别概要）后，按页码顺序在内存中拼接完整内容
- 一次性写入 `幻灯片索引.md`
- 工具选择：Python 脚本（`open(..., 'w')`）或 sed（如可用）

**方案 B：先写框架，但不用占位符替换**
- 若必须提前生成框架，每页只写图片引用和备注
- 识别概要栏完全留空（不写 `*待分析后填入*` 占位符）
- 分析完成后，按页码顺序在空行位置插入识别概要，而非替换文本
- 工具选择：sed 地址范围定位（`/## 第 N 页/,/^---$/`）或 Python 正则

**工具选择优先级：**
1. **sed 可用时**：使用地址范围精确定位页码，配合临时文件插入多行内容。详见 `references/sed-coordinate-positioning-multiline-replace-insert-cheatsheet.md`
2. **sed 不可用时**：退化到 Python 脚本，在内存中构建完整内容后一次性写入

### 索引文件格式

```markdown
# 幻灯片索引

## 第 1 页

![第1页](assets/幻灯片_01.png)

**备注**：（备注内容，若无则写「无」）

> **识别概要**
>
> 1. **标题区**（位置）：内容描述
> 2. **元素类型**（位置）：内容描述
> ...

---

## 第 2 页

![第2页](assets/幻灯片_02.png)

**备注**：（备注内容）

> **识别概要**
>
> ...

---
```

### 索引文件编写规范

- 每页一个二级标题 `## 第 N 页`
- 图片使用相对路径引用：`assets/文件名`
- **备注栏必须填入实际提取的演讲者备注**，而非留空；若某页无备注，标注「无」
- 识别概要使用 `> **识别概要**` 格式，按元素分条列出
- 如果页数较多（>30 页），可在文件顶部加一个「目录速览」小节，按章节分组列出页码范围
- 若备注内容较长（超过 300 字），可适当精简后填入，或注明「详见备注提取原始文件」

---

## 第六步：AI 视觉分析

这是本技能的核心分析环节。AI 读取每张幻灯片图片，进行视觉内容识别，并将识别结果以结构化格式**直接回填到 `幻灯片索引.md`** 中对应页面的 `> **识别概要**` 位置。

### 6.1 分析原则

1. **阅读顺序优先**：按人类正常的阅读顺序（通常从上到下、从左到右）依次描述页面内容，而非随机罗列。
2. **元素分类识别**：明确区分页面上的不同元素类型：
   - **标题文字块**：页面主标题、章节标题、小标题
   - **正文文字块**：段落文字、列表项、说明文字
   - **图片/图示块**：照片、插图、装饰图
   - **图表块**：数据图表、统计图、趋势图
   - **表格块**：结构化数据表格
   - **架构/流程图块**：系统架构图、业务流程图、网络拓扑图
   - **UI 截图块**：系统界面截图、APP 截图
   - **其他图形元素**：箭头、边框、色块、图标等
3. **位置描述**：对每类元素给出大致的页面位置（如「页面顶部居中」「左侧 1/3 区域」「右下角」等），必要时给出相对坐标百分比。
4. **表格特殊处理**：遇到表格时，不单纯描述为「一个表格」，而应尝试将表格内容转录为 Markdown 表格格式；若表格文字模糊或结构复杂无法恢复，再退化为列表式描述。

### 6.2 识别概要输出格式

识别概要统一使用 **Markdown 块引用（`>`）** 格式，插入在索引文件的每页图片下方、备注上方或下方。格式如下：

```markdown
> **识别概要**
>
> 1. **标题区**（页面顶部居中）：「综合感知服务系统升级改造与运营（2025）项目软件开发服务 项目验收汇报」
> 2. **背景图**（页面下半部）：广州市城市天际线全景照片，可见广州塔等地标建筑
> 3. **时间标注**（右下角）：「2025年12月」
> 4. **页面类型**：标题页 / 封面页
```

如果页面元素较多，可按区域分条列出；如果页面内容简单，可合并为一段描述。

**表格页面的识别概要示例**：

```markdown
> **识别概要**
>
> 1. **标题区**（顶部居中）：「2024年度监测成果统计」
> 2. **统计表格**（页面中上部，约占50%）：
>    | 监测类型 | 完成图斑数 | 核查通过率 | 完成时间 |
>    |---|---|---|---|
>    | 耕地保护监测 | 12,345 | 98.5% | 2024年12月 |
>    | 违建监测 | 3,456 | 95.2% | 2024年11月 |
>    | 生态修复监测 | 2,109 | 99.1% | 2024年10月 |
> 3. **说明文字**（表格下方）：数据来源及统计口径说明
> 4. **页面类型**：内容页（数据统计）
```

### 6.3 子图截取（可选但推荐）

对于页面上的**关键架构图、流程图、数据图表、UI 截图**等重要视觉元素，如果 AI 能够识别出其大致边界，应使用 Python PIL 截取出子图，保存到 `assets/crops/` 目录下，并在识别概要中引用。

> **注意**：表格通常**不需要截取为子图**，优先直接转录为 Markdown 表格。只有在表格包含复杂配图、颜色编码或合并单元格导致 Markdown 无法表达时，才考虑截取子图辅助说明。

**子图截取脚本示例**：

```python
from PIL import Image
import os

def crop_subimage(src_path, crop_path, bbox_ratio):
    """
    bbox_ratio: (left, top, right, bottom) 均为 0-1 的相对坐标
    """
    img = Image.open(src_path)
    w, h = img.size
    left = int(bbox_ratio[0] * w)
    top = int(bbox_ratio[1] * h)
    right = int(bbox_ratio[2] * w)
    bottom = int(bbox_ratio[3] * h)
    cropped = img.crop((left, top, right, bottom))
    os.makedirs(os.path.dirname(crop_path), exist_ok=True)
    cropped.save(crop_path)
    return crop_path

# 示例：截取第 5 页的架构图（假设位于页面中央 20%-80% 区域）
crop_subimage(
    "assets/幻灯片_05.png",
    "assets/crops/幻灯片_05_架构图.png",
    (0.1, 0.3, 0.9, 0.8)
)
```

截取后的子图在识别概要中引用：

```markdown
> **识别概要**
>
> 1. **标题区**：「1.1 项目背景：政策要求」
> 2. **左侧图片块**：领导人考察照片（[查看子图](assets/crops/幻灯片_04_左侧照片.png)）
> 3. **右侧文字块**：三段政策说明文字
> 4. **底部文件展示区**：三个政策文件封面缩略图横向排列
```

### 6.4 分批与并行处理策略

**单线程分批（页数 ≤ 30）**

| 幻灯片类型 | 批次大小 | 说明 |
|---|---|---|
| 标题页 + 目录页 + 过渡页 | 10 页 | 视觉密度低，可快速识别 |
| 纯文字内容页 | 7–10 页 | 以文字块识别为主 |
| 架构图 / 流程图 / 表格页 | 3–5 页 | 元素复杂，需仔细识别位置和关系 |
| UI 截图 / 数据图表页 | 3–5 页 | 细节多，可能需要子图截取 |
| 混合内容页 | 5–7 页 | 平衡处理速度和分析精度 |

**多 Agent 并行（页数 > 30）**

当 PPT 页数较多时，主 Agent 应启动多个后台子 Agent 并行处理，以显著缩短总耗时：

1. **任务拆分**：按页码将剩余页面划分为 4–5 个连续区间，每区间 10–15 页。
2. **子 Agent 职责**：读取分配到的图片 → 生成识别概要文本 → 返回给主 Agent。子 Agent **不直接修改**索引文件，避免并发写冲突。
3. **主 Agent 协调**：收集所有子 Agent 结果后，统一通过 `StrReplaceFile` 或 Python 脚本批量回填到 `幻灯片索引.md`。
4. **容错重试**：若某子 Agent 中断（如后台任务丢失），主 Agent 应重新启动该批次任务，或降级为单线程处理。
5. **回填脚本示例**：

```python
import re

# 从子 Agent 输出日志中提取识别概要
summaries = {}
pattern = r'## 第 (\d+) 页\s*\n\s*\n(> \*\*识别概要\*\*[\s\S]*?)(?=\n---|\n## 第|\Z)'
for match in re.finditer(pattern, agent_output):
    summaries[int(match.group(1))] = match.group(2).strip()

# 批量插入索引文件
with open('幻灯片索引.md', 'r', encoding='utf-8') as f:
    content = f.read()

for page_num in sorted(summaries.keys()):
    old = rf'(## 第 {page_num} 页\n\n!\[第{page_num}页\]\(assets/幻灯片_{page_num:02d}\.png\)\n\n)(\*\*备注\*\*：)'
    new = rf'\1{summaries[page_num]}\n\n\2'
    content = re.sub(old, new, content)

with open('幻灯片索引.md', 'w', encoding='utf-8') as f:
    f.write(content)
```

### 6.5 写入索引文件

**关键原则：按页码顺序一次性生成，避免逐页替换。**

分析完成后，主 Agent 应汇总所有识别概要，按页码顺序生成完整的索引文件内容，一次性写入。不要采用"先写占位符再替换"的模式。

**工具选择（按优先级）：**

#### 方案 1：sed 命令（推荐，如果可用）

```bash
# 1. 确认 sed 可用
sed --version

# 2. 为每页准备识别概要临时文件（如 summaries/page_05.md）
# 内容格式：
# > **识别概要**
# >
# > 1. **标题区**（位置）：内容...

# 3. 批量插入所有页的识别概要
for i in $(seq -w 1 41); do
    page_num=$(echo $i | sed 's/^0//')
    summary_file="summaries/page_$i.md"
    if [ -f "$summary_file" ]; then
        # 在该页范围内删除占位符，然后插入内容
        sed -i "/## 第 $page_num 页/,/^---$/{s/待分析后填入//}" 幻灯片索引.md
        sed -i "/## 第 $page_num 页/r $summary_file" 幻灯片索引.md
    fi
done
```

**sed 优势：**
- 使用页码标题 `## 第 N 页` 作为锚点，配合 `---` 分隔符精确定位
- 地址范围 `/## 第 N 页/,/^---$/` 确保只影响目标页
- 多行内容通过临时文件 + `r` 命令插入，避免特殊字符问题

详见 `references/Sed坐标定位+多行替换_插入速查文档.md`

#### 方案 2：Python 脚本（sed 不可用时退化）

```python
import re

# 1. 读取备注
notes = {}
with open('备注提取.md', 'r', encoding='utf-8') as f:
    content = f.read()
for m in re.finditer(r'## 第 (\d+) 页\n\*\*备注\*\*：(.*?)\n(?=\n## 第 |\Z)', content, re.DOTALL):
    notes[int(m.group(1))] = m.group(2).strip()

# 2. 收集识别概要（从子 Agent 输出或本 Agent 分析结果）
summaries = {}  # page_num -> summary_text

# 3. 按页码顺序生成完整内容
lines = ['# 幻灯片索引', '', '## 目录速览', ...]  # 头部

for i in range(1, total_pages + 1):
    note = notes.get(i, '')
    summary = summaries.get(i, '> **识别概要**：待补充')
    lines.append(f'## 第 {i} 页')
    lines.append('')
    lines.append(f'![第{i}页](assets/幻灯片_{i:02d}.png)')
    lines.append('')
    lines.append(f'**备注**：{note if note else "（无）"}')
    lines.append('')
    lines.append(summary)
    lines.append('')
    lines.append('---')
    lines.append('')

# 4. 一次性写入
with open('幻灯片索引.md', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
```

**为什么这样做更好：**
- 避免 `StrReplaceFile` 处理特殊字符时的 JSON 解析失败
- 避免逐页替换导致的错位问题
- 减少工具调用次数（从 N 次替换变为 1 次写入）
- 代码逻辑清晰，易于调试

**如果必须用 StrReplaceFile：**
仅在单页微调、已知文本不含特殊字符时使用。批量场景优先用 sed 或 Python 一次性生成。

### 6.6 特殊元素增强识别规则

除常规元素外，以下三类视觉元素需要触发**增强识别**逻辑：

#### 1. 证书 / 奖状 / 知识产权类

**触发条件**：页面标题或正文出现「证书」「奖状」「专利」「软著」「著作权」「获奖」「知识产权」等关键词；或页面中包含多张带边框/花纹的正式证件图片。

**增强动作**：
- 对**每张证书/奖状独立进行视觉精读**，尽可能提取：
  - 证书名称（如「2025地理信息产业优秀工程金奖」）
  - 颁发/授予机构（如「中国地理信息产业协会」）
  - 获奖/授权对象（如「广东省国土资源测绘院」）
  - 证书编号（如「软著登字第11841619号」）
  - 颁发时间（如「2025年9月」「2023年10月18日」）
- 如果证书缩略图文字过小无法完整辨认，至少记录：
  - 证书类型（发明专利 / 实用新型 / 软件著作权 / 获奖证书 等）
  - 可见数量（如「左侧竖向排列3张专利证书」）
  - 中央重点证书的标题信息

#### 2. 流程图 / 架构图 / 数据流

**触发条件**：页面中出现多个方框、菱形、圆角矩形等形状，并通过箭头、连线或折线相互连接；或页面标题/正文中含「流程」「架构」「体系」「拓扑」「数据流」等词。

**增强动作**：
- **自动裁剪**：识别出图表的整体外接边界后，使用 PIL 将整个流程图/架构图裁剪为子图，保存到 `assets/crops/幻灯片_XX_流程图.png`（或 `_架构图.png`）。
- **流向描述**：在识别概要中用文字描述：
  - 主要节点名称及顺序
  - 箭头方向（自上而下、从左到右、双向回流等）
  - 关键分支判断条件
- **引用子图**：在识别概要中附加子图引用链接，如：
  ```markdown
  > 3. **业务流程架构图**（页面中下部，约占60%）：[查看子图](assets/crops/幻灯片_60_业务流程架构图.png)
  >    - 顶部：日常变更部级核查系统 → 国土调查国家核查进度查询系统
  >    - 中部核心：广东省自然资源综合感知服务系统，包含……
  ```

#### 3. 表格

**触发条件**：页面中出现多行多列的结构化数据，带有表头行或单元格边框；或识别概要内容中出现「表格」字样。

**恢复判定标准**：

| 适合恢复为 Markdown 表格 | 建议降级为文字描述 |
|---|---|
| 结构规整，**无合并单元格**；或合并单元格仅表现为"左侧分类名对应右侧多行子项"（可用编号①②③或分号将子项合并到同一单元格） | 存在**复杂的跨行/跨列合并**（多个列同时交叉合并、层级嵌套），Markdown 表格无法表达 |
| 行列数适中（**行≤20，列≤8**） | 行列数过多（行>20 或 列>8），恢复后可读性差 |
| 单元格内容为**短文本、数字、日期、状态**（如名称、金额、完成率） | 某列内容为**大段描述文字**，恢复后行宽严重失控 |
| 数据具有**结构化查询价值**（工作任务清单、绩效目标表、合同清单、金额汇总、指标对比） | 仅为 UI 截图中的装饰性数据展示，信息量低 |
| 文字清晰，行列边界明确 | 表格文字模糊、行列边界不清 |

> **例外原则**：即使表格行列数略多（如 16 行×7 列的合同支付汇总表），只要其承载的是**验收关键数据**（工作任务清单、绩效目标、金额、合同编号、时间节点、责任主体等），仍应优先恢复为 Markdown 表格。Markdown 表格便于后续全文检索、数据引用和跨文档比对，其价值远超阅读时的排版成本。

**增强动作**：
- **优先恢复 Markdown 表格**：尝试将表格内容转录为 Markdown 表格格式（`| 表头1 | 表头2 | ... |`）。
- **降级策略**：如果出现以下情况，退化为列表式文字描述：
  - 表格文字模糊、行列难以对齐
  - **单元格合并复杂且无法简化**：左侧分类列虽有合并，但若合并仅表现为"一个分类名对应右侧若干子项"，应优先用编号①②③或分号将子项合并到同一单元格，**不应直接降级**。仅当多个列同时存在交叉合并、层级嵌套等复杂结构，导致 Markdown 表格无法表达时才降级
  - 某列内容极长（如大段描述文字），恢复为表格后导致行宽失控、影响阅读
  - 恢复后格式严重错位
- **数据准确性**：对于关键数字（金额、数量、百分比、日期），必须确保转录准确；若不确定，在概要用「约」「疑似」等词标注。

### 6.7 大规模 PPT 分析实战经验

本节总结在实际处理 80+ 页 PPT 过程中踩过的坑和验证有效的做法。

#### 1. 并行 Agent 管理

- **上限**：系统通常最多允许 **4 个后台 Agent** 同时运行。超出会排队或失败。
- **分波策略**：先启动 4 个子 Agent 处理前 4 个区间，待全部返回后再启动剩余批次。
- **输出方式**：让子 Agent **不直接修改**索引文件，而是将识别概要文本输出到日志或临时文件，由主 Agent 统一收集后批量回填。这是避免并发写冲突的唯一可靠方式。
- **容错重试**：Agent 可能因后台任务丢失（`task.lost`）而中断。主 Agent 应在收集结果时检查每个区间的返回状态，缺失的页码区间需重新派发任务。
- **回填验证**：回填完成后，用 `grep -c "识别概要"` 统计出现次数，应等于总页数。若不等于，定位缺失页码补分析。

#### 2. 证书/奖状类页面的深度处理

当页面包含证书、专利、软著等知识产权图片时，仅靠视觉识别缩略图往往无法读出证书编号、时间等关键字段。**正确流程**：

1. 用 `python-pptx` 定位到该页（注意 Windows 环境用 `python` 而非 `python3`）。
2. 遍历该页所有 `shape.shape_type == 13`（PICTURE）的形状，提取 `shape.image.blob` 保存为独立图片。
3. 按图片尺寸去重：PPT 中常同时嵌入**缩略图**（宽 400–700px）和**原图**（宽 2000+px），需通过内容比对或尺寸阈值去重。
4. 对去重后的每张证书图片进行视觉精读，提取：证书名称、编号/专利号/登记号、颁发机构、时间、权利人等。
5. 将结构化结果回填到索引文件，原缩略图描述替换为详细清单。

**证书提取脚本示例**：

```python
from pptx import Presentation
from PIL import Image
import io, os

pptx_path = "汇报.pptx"
prs = Presentation(pptx_path)
slide = prs.slides[12]  # 第13页

os.makedirs("assets/crops/slide13_certs", exist_ok=True)
seen_hashes = set()
for i, shape in enumerate(slide.shapes):
    if shape.shape_type == 13:  # PICTURE
        img_bytes = shape.image.blob
        img = Image.open(io.BytesIO(img_bytes))
        # 按尺寸过滤去重：跳过与已保存图片尺寸相同的缩略图
        size_key = (img.width, img.height)
        if size_key in seen_hashes:
            continue
        seen_hashes.add(size_key)
        ext = shape.image.ext if shape.image.ext != "jpeg" else "jpg"
        img.save(f"assets/crops/slide13_certs/cert_{i:02d}.{ext}")
```

#### 3. Windows 环境文件 IO 注意事项

- **exit code 49**：在 Windows bash 环境下，Python 直接执行带中文路径的文件写操作偶发失败（exit code 49）。
- **根本原因**：往往不是 Python 本身的问题，而是将含特殊字符（中文引号、反引号、多行文本等）的内容直接嵌入 Shell 命令的 JSON 参数中，导致解析失败。
- **解决方案**：
  - **首选**：用 Python 脚本在内存中构建完整内容，一次性写入文件。将 Python 代码先写入临时文件再执行：`cat > /tmp/script.py` 然后 `python /tmp/script.py`。
  - **避免**：将大段含特殊字符的文本直接作为 Shell 命令参数传递。
  - **StrReplaceFile 适用场景**：仅用于简单的、已知不含特殊字符的单行替换。批量写入不要用此工具逐行拼接。

#### 4. 子图截取坐标校准

AI 估算的相对坐标（如 `(0.1, 0.3, 0.9, 0.8)`）为近似值。截取后应打开子图检查，若边界不准：
- 微调坐标后重新裁剪；
- 或读取 PPT 中 shape 的绝对坐标（`shape.left`, `shape.top`, `shape.width`, `shape.height`）精确计算边界。

---

## 第七步：备注空置分析与补充备注生成

### 7.1 分析备注空置率

在第六步 AI 视觉分析完成后，应统计 `备注提取.md` 中的备注空置率：

```python
import re

def analyze_note_vacancy(notes_md_path):
    with open(notes_md_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 使用 [\s\S]*? 匹配跨页内容，\n\n 作为页分隔符，确保能正确匹配到每页末尾
    pages = re.findall(r'## 第 (\d+) 页[\s\S]*?\*\*备注\*\*：(.*?)\n\n', content)
    total = len(pages)
    empty = sum(1 for _, note in pages if not note.strip() or note.strip() == '（无）')
    vacancy_rate = empty / total if total > 0 else 0
    
    print(f"总页数: {total}, 空置页数: {empty}, 空置率: {vacancy_rate:.1%}")
    return vacancy_rate, [page for page, note in pages if not note.strip() or note.strip() == '（无）']
```

### 7.2 空置率判定与处理策略

| 空置率 | 判定 | 处理方式 |
|---|---|---|
| 0% | 备注完整 | 跳过本步骤，直接使用提取的原始备注 |
| < 30% | 少量缺失 | 仅对空置页根据识别概要和页面文本生成补充备注 |
| ≥ 30% | 大量缺失 | 判定为"PDF 做的幻灯片"或"演讲者未写讲稿"，**全部页面**根据识别概要和页面文本重新生成备注 |
| 100% | 完全空置 | 全部重新生成；原始 `备注提取.md` 仅保留作为页面文本参考 |

### 7.3 生成补充备注的原则

补充备注应作为**演讲者讲稿**，而非简单的页面内容复述：

1. **口语化表达**：使用适合口头汇报的语言，避免照念屏幕文字
2. **补充解释**：对页面上的关键词、数据、图表进行解读和延伸说明
3. **过渡衔接**：章节过渡页应有承上启下的引导语
4. **时间控制**：每页备注字数建议 50–200 字，根据页面重要性调整
5. **基于识别概要**：充分利用第六步生成的识别概要中的结构化信息
6. **结合页面文本**：参考 `备注提取.md` 中的页面文本，确保讲稿与页面内容一致

### 7.4 回填补充备注

生成的补充备注需要回填到两个位置：

1. **`备注提取.md`**：替换原有的「（无）」标注，保留页面文本部分不变
2. **`幻灯片索引.md`**：同步更新每页的 `**备注**：` 栏

**回填脚本示例**：

```python
import re
import os
import json

def fill_generated_notes(index_path, notes_md_path, generated_notes):
    """
    generated_notes: dict, page_num -> note_text
    """
    # 1. 更新 幻灯片索引.md
    with open(index_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    for page_num, note_text in generated_notes.items():
        # 替换该页的备注栏
        pattern = rf'(## 第 {page_num} 页\s*\n\s*!\[.*?\]\(.*?\)\s*\n\s*)\*\*备注\*\*：.*?\n'
        replacement = rf'\1**备注**：{note_text}\n'
        content = re.sub(pattern, replacement, content, flags=re.DOTALL)
    
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    # 2. 更新 备注提取.md
    with open(notes_md_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    for page_num, note_text in generated_notes.items():
        pattern = rf'(## 第 {page_num} 页\s*\n\s*)\*\*备注\*\*：.*?\n'
        replacement = rf'\1**备注**：{note_text}\n'
        content = re.sub(pattern, replacement, content, flags=re.DOTALL)
    
    with open(notes_md_path, 'w', encoding='utf-8') as f:
        f.write(content)

# 3. 验证回填结果
def verify_notes(index_path, notes_md_path, total_pages):
    for path in [index_path, notes_md_path]:
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()
        vacancy_count = content.count('**备注**：（无）')
        print(f"{os.path.basename(path)} 中剩余空置: {vacancy_count}/{total_pages}")
```

### 7.5 将备注写回 PPT 文件（可选，谨慎操作）

如果用户**明确要求**将 AI 生成的补充讲稿保存到 PPT 文件中，可以使用 `write_notes.py` 脚本批量写入。

> **⚠️ 强烈建议**：在副本文件上操作，保留原始 PPT 文件不变。脚本默认输出到 `<原文件名>_已更新备注.pptx`，不会覆盖原文件。

**脚本位置**：`{skill-dir}/scripts/write_notes.py`

**前置依赖**：
```bash
pip install python-pptx
```

**用法一：从 JSON 文件写入**

```bash
python {skill-dir}/scripts/write_notes.py <pptx路径> <json路径> [输出路径]
```

JSON 文件格式：
```json
{
    "pages": [
        {"page": 1, "note": "第一页补充讲稿内容"},
        {"page": 3, "note": "第三页补充讲稿内容"}
    ]
}
```

**用法二：从 Markdown 文件写入（推荐）**

```bash
python {skill-dir}/scripts/write_notes.py --from-md <pptx路径> <备注提取.md路径> [输出路径]
```

此命令会自动解析 `备注提取.md` 中的每页备注内容，过滤掉「（无）」占位符，将有效备注写回 PPT。

**用法三：批量范围写入**

```json
{
    "pages": [
        {"page": "1-5", "notes": ["备注1", "备注2", "备注3", "备注4", "备注5"]}
    ]
}
```

**脚本特性**：
- 自动为没有备注幻灯片的页面创建备注布局
- 写入前清空原有备注，避免追加重复
- 支持单页写入、批量范围写入、从 Markdown 解析写入三种模式
- **默认输出到 `<原文件名>_已更新备注.pptx`，不覆盖原文件**

> **注意**：
> 1. 回填时应保留 `备注提取.md` 中的「页面文本」部分，仅替换「备注」栏。
> 2. 若通过子 Agent 生成备注并以 JSON 格式传递，**必须使用 `json.dump()` 序列化**，禁止手动拼接 JSON 字符串（中文引号 `"` 未转义会导致解析失败）。
> 3. Windows 环境下，Python 的 `open('/tmp/...')` 可能无法访问 bash 的 `/tmp` 目录。应使用 `os.environ['TEMP']` 获取系统临时目录，或直接使用相对路径/绝对路径。

---

## 第八步（可选）：汇总分析与生成报告

所有页面的识别概要和备注回填完成后，**询问用户是否需要**进行整体汇总分析。若用户需要，则生成 `分析结果.md`：

1. **整体结构**：按章节划分，列出每章对应的页码范围和核心内容
2. **核心结论**：项目的主要建设成果、关键指标、验收结论
3. **关键概念与术语清单**：系统模块名、政策术语、技术名词
4. **可抽取的实体**：项目名、单位名、系统名、平台名、厂商名
5. **关键图表索引**：列出所有被截取子图的架构图、流程图、数据图表，注明所在页码和业务意义

---

## 分析提示词模板（供 AI 内部使用）

> 你正在审阅一份项目汇报 PPT 的幻灯片图片。请按以下要求输出识别结果：
>
> 1. **阅读顺序**：严格按照人类阅读顺序（从上到下、从左到右）描述页面内容。
> 2. **元素分类**：明确区分标题、正文、图片、图表、表格、架构图、UI 截图等不同类型的元素。
> 3. **位置描述**：给出每个元素的大致页面位置（如顶部居中、左栏、右下角等）。
> 4. **文字内容**：准确转录可见的关键文字，特别是系统名、模块名、数字、日期、政策引用。
> 5. **子图建议**：如果页面上有独立的架构图、流程图、数据图表或 UI 截图，指出其大致边界坐标（相对百分比），建议截取为子图。
> 6. **证书精读**：如遇证书、奖状、专利、软著等知识产权图片，逐张提取证书名称、颁发机构、获奖对象、时间、编号。
> 7. **流程图裁剪**：如遇流程图、架构图、数据流（含方框+箭头/连线），裁剪整个图表为子图，并描述节点流向与关键分支。
> 8. **表格恢复**：如遇表格，优先恢复为 Markdown 表格；若格式复杂或文字不清，退化为列表式描述，但确保关键数字准确。
> 9. **业务意义**：简要说明该页在项目叙事中的作用。

---

## 常见问题

| 问题 | 解决方法 |
|---|---|
| WPS 找不到「输出为图片」菜单 | 尝试「文件 → 另存为 → 输出为图片」或在搜索框中输入「输出为图片」 |
| WPS PDF 找不到「导出 PDF 为图片」 | 尝试「转换 → PDF 转图片」或在搜索框中输入「PDF 转图片」 |
| 导出图片顺序错乱 | WPS 通常按页码顺序导出，若出现乱序，请按文件名中的数字排序后读取 |
| 图片文字模糊 | 重新导出，选择更高的输出尺寸或 DPI |
| 幻灯片页数很多（>100 页） | 分批导出或一次性导出后分批次让 AI 分析 |
| 只想分析部分页面 | 可在对话框中选择「页码选择」，仅导出需要的页面 |
| python-pptx 读取不到备注 | 先用 WPS「另存为」保存为标准 .pptx 格式后再提取 |
| PDF 文件没有备注怎么办 | 属于正常情况，第七步会自动检测空置率并生成补充讲稿 |
| 子图截取坐标不准 | AI 给出的相对坐标为估算值，截取后可人工微调或重新裁剪 |
| Agent 并行数受限 | 系统通常最多同时运行 4 个后台 Agent，超出需分波处理 |
| Windows Python 写文件失败（exit code 49） | 避免将含特殊字符的文本直接嵌入 Shell 命令；将 Python 脚本先写入临时文件再执行，在内存中构建完整内容后一次性写入 |
| PPT 中证书图片重复 | 同一证书常同时嵌入缩略图和原图，按图片尺寸或内容哈希去重后再识别 |
| 子 Agent 结果丢失 | 收集结果时检查每页是否都有识别概要，缺失页码重新派发 |
| 如何将补充讲稿写回 PPT | 仅在用户明确要求时使用 `write_notes.py`；脚本默认输出到 `<原文件名>_已更新备注.pptx`，不覆盖原文件，建议在副本上操作 |

---

## 快速开始示例

### PPT 文件示例

以文件 `综合感知服务（2022-2024年）-整体验收PPT20260429.pptx` 为例：

1. 创建目录：`综合感知服务（2022-2024年）-整体验收PPT20260429/assets/`
2. 用 WPS 演示打开该 PPT，菜单选择「文件 → 输出为图片」。
3. 输出格式选 PNG，输出目录选上面创建的 `assets/` 文件夹。
4. 若文件名过长，运行重命名脚本统一为 `幻灯片_XX.png`。
5. 用 `python-pptx` 提取每页演讲者备注，保存为 `备注提取.md`。
6. AI 分批读取图片进行视觉分析：
   - 页数 ≤ 30：单线程逐批分析，每批 5–10 页。
   - 页数 > 30：启动 4 个子 Agent 并行处理，每 Agent 负责 10–15 页，子 Agent **不直接修改**索引文件，仅输出识别概要到日志。
   - 遇到证书/奖状页：回到 PPT 用 `python-pptx` 提取高清原图，逐张精读并结构化记录。
   - 遇到流程图/架构图：裁剪整图为子图，描述节点流向。
   - 遇到表格：优先恢复为 Markdown 表格。
7. 主 Agent 汇总所有结果，按页码顺序生成完整索引内容，**一次性写入** `幻灯片索引.md`。
   - 不要先写框架再回填！直接在内存中拼接「图片引用 + 备注 + 识别概要」，一次性写入。
8. 写入完成后验证：统计 `识别概要` 出现次数应等于总页数，抽样检查格式。
9. **分析备注空置率**：统计 `备注提取.md` 中备注为空的页数比例。
    - 若空置率 ≥ 30% 或 100%：判定为 PDF 幻灯片或未写讲稿，根据识别概要和页面文本为所有空置页生成补充备注（讲稿），回填到 `备注提取.md` 和 `幻灯片索引.md`。
    - 若空置率 < 30%：仅对个别空置页生成补充备注。
10. （可选，仅在用户明确要求时）使用 `write_notes.py` 将补充备注写回 PPT 副本：`python write_notes.py --from-md <pptx路径> <备注提取.md路径>`。脚本默认生成 `<原文件名>_已更新备注.pptx`，不会覆盖原文件。
11. 汇总分析生成 `分析结果.md`。

### PDF 文件示例

以文件 `【遥感中心 白云分院】AI智鉴，数护农险AI赋能农业保险全流程数智革新.pdf` 为例：

1. 创建目录：`【遥感中心 白云分院】AI智鉴，数护农险AI赋能农业保险全流程数智革新/assets/`
2. 用 **WPS PDF** 打开该 PDF，菜单选择「文件 → 导出 PDF 为图片」（或「转换 → PDF 转图片」）。
3. 输出格式选 PNG，输出品质选「高清」或「超清」，输出目录选上面创建的 `assets/` 文件夹。
4. 若文件名过长，运行重命名脚本统一为 `幻灯片_XX.png`。
5. **跳过提取备注步骤**：PDF 不含演讲者备注，直接创建一个空的 `备注提取.md`（每页标注「备注：（无）」）。
6. AI 分批读取图片进行视觉分析（同 PPT 流程）。
7. 主 Agent 汇总结果，一次性写入 `幻灯片索引.md`。
8. 写入完成后验证识别概要数量。
9. **分析备注空置率**：PDF 文件空置率通常为 100%，判定为完全空置。
   - 启动子 Agent 根据识别概要和页面文本为全部 27 页生成补充讲稿备注。
   - 回填到 `备注提取.md` 和 `幻灯片索引.md`。
10. 汇总分析生成 `分析结果.md`。
