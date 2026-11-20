<a id="readme-top"></a>

<div align="center">

# ppt-analyzer

### 用视觉 AI 把 PPTX 和 PDF 幻灯片整理为结构化 Markdown 的 AI 原生技能

WPS 逐页导出 PNG，AI 读取图片重建页面结构、表格与演讲者备注。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![GitHub stars](https://img.shields.io/github/stars/kuwork/ppt-analyzer?logo=github)](https://github.com/kuwork/ppt-analyzer/stargazers)
[![GitHub issues](https://img.shields.io/github/issues/kuwork/ppt-analyzer?logo=github)](https://github.com/kuwork/ppt-analyzer/issues)
[![Platform: Windows](https://img.shields.io/badge/Platform-Windows-0078D4?logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![Python: optional scripts](https://img.shields.io/badge/Python-optional%20scripts-3776AB?logo=python&logoColor=white)](https://www.python.org/)

[English](README.md) ｜ 简体中文

[功能亮点](#功能亮点) ·
[工作流程](#工作流程) ·
[快速开始](#快速开始) ·
[其他安装方式](#其他安装方式) ·
[常见问题](#常见问题) ·
[已知限制](#已知限制) ·
[项目结构](#项目结构)

</div>

**ppt-analyzer** 是一个在 Windows 上分析 `.pptx` 和 `.pdf` 幻灯片的 AI 原生技能。它不通过 COM 组件驱动 Office，而是采用两段式流程：先由你在 WPS Office 中把幻灯片逐页手动导出为 PNG，再由具备视觉能力的 AI 读取图片，产出结构化 Markdown——包含逐页识别概要的幻灯片索引、提取出的演讲者备注、Markdown 表格转录，以及可选的整体分析报告。

> [!NOTE]
> 仓库：[github.com/kuwork/ppt-analyzer](https://github.com/kuwork/ppt-analyzer) ｜ 问题反馈：[github.com/kuwork/ppt-analyzer/issues](https://github.com/kuwork/ppt-analyzer/issues) ｜ 技能入口：`skills/ppt-analyzer/SKILL.md` ｜ 支持格式：`.pptx` · `.pdf`

## 功能亮点

- 🧠 **AI 原生技能，无 COM 依赖**：技能本身是一套由 AI 宿主直接加载的指令集（`SKILL.md`），核心导出与分析链路不需要 Office COM 组件或第三方解析库。
- 🖼️ **基于 WPS 的图片流水线**：幻灯片由 WPS Office 逐页导出为 PNG，图片质量由 WPS 直接保证，可在各种 Windows 环境中运行。
- 📄 **一套流程同时支持 PPTX 与 PDF**：`.pptx` 和 `.pdf` 走同一套处理步骤；PDF 不含演讲者备注的情况会被自动识别并处理。
- 🗣️ **演讲者备注全生命周期管理**：`extract_notes.py` 基于 python-pptx 提取备注，统计空置率（达到 30% 即触发全量重生成），AI 撰写每页 50–200 字的补充讲稿，再由 `write_notes.py` 写回 PPT 副本。
- ✂️ **关键图示智能子图截取**：架构图、流程图、数据图表通过 PIL 裁剪到 `assets/crops/`，并在对应页面的识别概要中引用。
- 📊 **表格优先恢复为 Markdown**：结构规整的表格（不超过 20 行 × 8 列）转录为 Markdown 表格，便于检索与比对；只有复杂合并单元格表格才退化为文字描述。
- 🏅 **证书专利深度精读**：对证书、专利、软著等知识产权图片逐张提取名称、颁发机构、权利人、时间与登记编号。
- ⚡ **大型 PPT 并行分析**：超过 30 页的幻灯片拆分为每组 10–15 页的区间，由最多 4 个子 Agent 并行处理；子 Agent 只返回文本，由主 Agent 统一写入，避免并发写冲突。

## 工作流程

本技能采用「人工导出 + AI 分析」两段式流程。避开 Office COM 组件和第三方解析库，使其兼容各种 Windows 环境，同时导出图片的质量由 WPS 直接保证。

| 步骤 | 阶段 | 具体动作 |
|---|---|---|
| 1 | 建目录 | 在仓库根目录以幻灯片文件名建立同名文件夹，并创建 `assets/` 子目录 |
| 2 | 导出图片 | 在 WPS 演示（PPTX）或 WPS PDF 中将每一页导出为 PNG，输出到 `assets/` |
| 3 | 整理文件名 | 若 WPS 生成的文件名过长，统一重命名为 `幻灯片_XX.png` |
| 4 | 提取备注（仅 PPTX） | 运行 `extract_notes.py`，将演讲者备注与页面文本保存为 `备注提取.md`；PDF 则创建空占位文件 |
| 5 | 编写索引 | 生成 `幻灯片索引.md`，汇总图片路径、备注与逐页识别概要 |
| 6 | AI 视觉分析 | AI 按阅读顺序读取每张图片，分类识别元素、转录表格、裁剪图示、精读证书 |
| 7 | 备注空置处理 | 统计空置率，为空置页生成 AI 补充讲稿，并可选写回 PPTX |
| 8 | 汇总报告（可选） | 按需生成 `分析结果.md`，输出章节结构、核心结论与关键图表索引 |

分析某个幻灯片时，会在仓库根目录生成如下工作文件夹：

```text
我的汇报PPT/
├── assets/                  # WPS 导出的幻灯片图片
│   └── crops/               # 截取的子图（架构图、流程图、数据图表）
├── 幻灯片索引.md             # 图片路径 + 演讲者备注 + AI 识别概要
├── 备注提取.md               # 演讲者备注与页面文本的完整备份
└── 分析结果.md               # 整体分析结论（可选）
```

## 快速开始

### 前置条件

- Windows 操作系统
- 已安装 WPS Office，且演示组件与 PDF 组件可用
- 待分析的 `.pptx` 或 `.pdf` 文件可在本地打开

### 安装技能（推荐：用户级技能目录）

Kimi Code CLI 会自动从 `~/.config/agents/skills/` 加载技能；`~/.kimi/skills/` 与 `~/.claude/skills/` 下的同名技能目录同样会被识别。

```bash
# 1. 进入用户级技能目录
cd ~/.config/agents/skills/   # 或 ~/.kimi/skills/ 或 ~/.claude/skills/

# 2. 克隆仓库到临时位置
git clone https://github.com/kuwork/ppt-analyzer.git /tmp/ppt-analyzer

# 3. 将技能文件夹复制到当前目录
cp -r /tmp/ppt-analyzer/skills/ppt-analyzer ./ppt-analyzer

# 4. 清理临时文件
rm -rf /tmp/ppt-analyzer
```

安装后的目录结构如下，AI 宿主会自动识别并加载该技能：

```text
~/.config/agents/skills/
└── ppt-analyzer/
    ├── SKILL.md
    ├── scripts/
    │   ├── extract_notes.py
    │   ├── extract_ppt.py
    │   └── write_notes.py
    └── references/
```

> [!TIP]
> 验证安装：确认 `~/.config/agents/skills/ppt-analyzer/SKILL.md` 存在，然后向 AI 宿主提出「分析 PPT」，技能应自动触发。项目级安装、npm 安装与临时读取方式见 [其他安装方式](#其他安装方式)。

### 执行分析

1. 创建以幻灯片命名的文件夹及 `assets/` 子目录。
2. 导出 PNG 图片：PPTX 用 WPS 演示的「文件 → 输出为图片」，PDF 用 WPS PDF 的「文件 → 导出 PDF 为图片」，格式选 PNG、方式选逐页输出。
3. 仅 PPTX：运行 `python scripts/extract_notes.py <pptx路径>` 提取演讲者备注（首次使用需先执行 `pip install python-pptx`）。
4. 让 AI 读取 `assets/` 中的图片，逐页生成识别概要。
5. AI 汇总生成 `幻灯片索引.md`，为空置页补充讲稿，并按需生成 `分析结果.md`。

完整操作步骤、脚本模板与最佳实践见 [`skills/ppt-analyzer/SKILL.md`](skills/ppt-analyzer/SKILL.md)。

## 其他安装方式

<details>
<summary><b>安装到项目目录（仅当前项目生效）</b></summary>

```bash
# 在项目根目录执行
mkdir -p .agents/skills

git clone https://github.com/kuwork/ppt-analyzer.git /tmp/ppt-analyzer
cp -r /tmp/ppt-analyzer/skills/ppt-analyzer .agents/skills/
rm -rf /tmp/ppt-analyzer
```

</details>

<details>
<summary><b>通过 npm 从 GitHub 安装</b></summary>

本包未发布到 npm registry，但 npm 支持直接从 GitHub 安装：

```bash
npm install kuwork/ppt-analyzer
```

技能文件位于 `node_modules/ppt-analyzer/skills/ppt-analyzer/`，之后可手动复制到技能目录：

```bash
cp -r node_modules/ppt-analyzer/skills/ppt-analyzer ~/.config/agents/skills/
```

</details>

<details>
<summary><b>临时使用：直接让 AI 读取技能</b></summary>

```text
读取 https://github.com/kuwork/ppt-analyzer/blob/main/skills/ppt-analyzer/SKILL.md
```

不会产生任何安装，AI 仅在当前对话中按该指令集工作。

</details>

## 使用场景

- 项目验收汇报 PPT 的结构化归档
- 培训课件、技术分享材料的内容提取与整理
- 商务方案、产品路演材料的要点梳理
- 将会议演示材料沉淀为可检索的知识库资产

## 常见问题

### WPS 里找不到「输出为图片」菜单怎么办？

尝试「文件 → 另存为 → 输出为图片」，或在 WPS 搜索框中输入「输出为图片」。WPS PDF 中的对应入口是「转换 → PDF 转图片」。

### 核心流程需要安装 Python 或 Node.js 吗？

不需要。导出与视觉分析流程只依赖 WPS 和 AI 宿主。Python 仅用于可选辅助脚本：备注提取与回填需要 `pip install python-pptx`；一体化脚本 `extract_ppt.py` 还需要额外安装 Aspose.Slides。

### 为什么 PDF 幻灯片的演讲者备注全部为空？

PDF 本身不承载演讲者备注，100% 空置率属于正常情况。技能会检测到这一点，并依据识别概要与页面可见文本为每一页生成补充讲稿。

### python-pptx 读不到 WPS 保存的 PPTX 备注怎么办？

先在 WPS 中通过「另存为」保存为标准 `.pptx` 格式，再重新运行 `extract_notes.py`。WPS 的备注存储方式偶尔与标准 Office 存在细微差异。

### 导出的幻灯片图片文字模糊怎么办？

以更高的输出尺寸或 DPI 重新导出：WPS 演示至少选择默认尺寸 2 倍，WPS PDF 选择「高清」或「超清」品质。文字清晰是表格恢复与证书精读准确的前提。

### 超过 30 页的大型 PPT 如何处理？

拆分为 4–5 个连续区间，每区间 10–15 页，由最多 4 个子 Agent 并行分析且只返回文本；主 Agent 随后按页码顺序一次性写入 `幻灯片索引.md`，避免并发写冲突。

### 把生成的备注写回 PPTX 安全吗？

`write_notes.py` 默认不会覆盖原文件，而是输出到名为 `<原文件名>_已更新备注.pptx` 的新文件。仍建议始终在副本上操作，并且仅在用户明确要求时才执行备注回填。

### 可以只分析部分页面吗？

可以。WPS 导出对话框支持页码选择，只导出需要的页面，再按相同流程处理即可。

## 已知限制

- **依赖 Windows 与 WPS Office**：文档中的导出流程基于 WPS 演示与 WPS PDF，未覆盖 macOS/Linux 路径。
- **导出为手动步骤**：WPS 导出由人工完成，这是有意设计；主流程不提供无头全自动截图能力。
- **Python 依赖为可选项**：备注脚本需要 python-pptx；一体化脚本 `extract_ppt.py` 还需要 Aspose.Slides。
- **识别准确度取决于导出分辨率**：密集表格与证书上的小字需要高分辨率 PNG 才能读准。
- **备注回填仅支持 PPTX**：PDF 没有演讲者备注容器，生成的讲稿只保留在 Markdown 中。

## 安全说明

> [!CAUTION]
> 所有辅助脚本仅在本地读写文件，不发起任何网络请求，也不处理凭据。`write_notes.py` 会保存为新的 `<原文件名>_已更新备注.pptx` 文件而非覆盖源文件。运行前可审计 `skills/ppt-analyzer/scripts/` 下的三个脚本，敏感幻灯片请保留在本机。

## 项目结构

<details>
<summary>查看仓库布局</summary>

```text
ppt-analyzer/
├── skills/
│   └── ppt-analyzer/
│       ├── SKILL.md                # 面向 AI 宿主的完整指令集
│       ├── scripts/
│       │   ├── extract_notes.py    # 基于 python-pptx 提取备注、文本与表格
│       │   ├── extract_ppt.py      # 一体化提取：Aspose.Slides 出图与备注 + python-pptx 文本
│       │   └── write_notes.py      # 将生成的备注批量写入 PPTX 副本
│       └── references/
│           └── sed-coordinate-positioning-multiline-replace-insert-cheatsheet.md
├── .gitignore
├── LICENSE
├── package.json
└── README.md
```

</details>

## 许可证

基于 [MIT License](LICENSE) 开源。

<p align="right"><a href="#readme-top">↑ 回到顶部</a></p>
