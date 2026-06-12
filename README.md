# PPT Analyzer

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 🎯 项目简介

**PPT Analyzer** 是一个 AI 原生技能，采用「人工导出 + AI 分析」的两段式流程，将 PPT/PDF 幻灯片内容整理为结构化的 Markdown 文档。

- **人工导出**：在 WPS 中将 PPT/PDF 按页输出为 PNG 图片
- **AI 分析**：AI 读取图片进行视觉识别，提取结构、要点和关键信息

这种流程避免了对 Office COM 组件或第三方库的依赖，兼容各种 Windows 环境，且图片质量由 WPS 直接保证。

> **支持格式**：`.pptx` 和 `.pdf` 格式的幻灯片分析。

---

## 📂 目录架构

```text
{project-root}/
├── skills/
│   └── ppt-analyzer/
│       ├── SKILL.md              # 技能完整指令集
│       ├── scripts/
│       │   └── extract_notes.py  # PPT 备注提取脚本
│       └── references/
│           └── sed-coordinate-positioning-multiline-replace-insert-cheatsheet.md
├── .gitignore
├── LICENSE
└── README.md
```

分析某个 PPT 时，会在仓库根目录创建以文件名命名的文件夹：

```text
./
├── 我的汇报PPT/
│   ├── assets/                  # WPS 导出的幻灯片图片
│   ├── assets/crops/            # 截取的子图（架构图、流程图等）
│   ├── 幻灯片索引.md             # 图片路径 + 备注 + AI 识别概要
│   ├── 备注提取.md               # 演讲者备注全文备份
│   └── 分析结果.md               # 整体分析结论（可选）
```

---

## 🚀 快速开始

### 前置条件

- Windows 操作系统
- 已安装 WPS Office（演示组件 + PDF 组件可用）
- 待分析的 `.pptx` 或 `.pdf` 文件可本地打开

### 分析流程

1. **创建分析目录**：以 PPT 文件名在仓库根目录建立同名文件夹，并创建 `assets/` 子目录
2. **WPS 导出图片**：
   - PPT 文件：WPS 演示 → 文件 → 输出为图片 → PNG 格式 → 逐页输出
   - PDF 文件：WPS PDF → 文件 → 导出 PDF 为图片 → PNG 格式 → 逐页输出
3. **提取演讲者备注**（PPT 专属）：运行 `python scripts/extract_notes.py <pptx路径>` 提取备注
4. **AI 视觉分析**：AI 读取 `assets/` 中的图片，逐页生成识别概要
5. **生成索引文件**：汇总所有结果，生成 `幻灯片索引.md`
6. **备注空置处理**：统计备注空置率，为空置页生成补充讲稿
7. **汇总分析**（可选）：生成 `分析结果.md`，输出整体结构、核心结论、关键图表索引等

详细操作步骤、脚本模板和最佳实践，请参阅 [`skills/ppt-analyzer/SKILL.md`](skills/ppt-analyzer/SKILL.md)。

---

## 🛠️ 核心特性

| 特性 | 说明 |
|---|---|
| **AI 原生实现** | 无需安装 Node/Python 环境，直接注入 AI 指令集 |
| **零依赖分析** | 不依赖 Office COM 组件，WPS 导出 + AI 视觉识别 |
| **PPT/PDF 双支持** | 同时支持 `.pptx` 和 `.pdf` 格式的幻灯片 |
| **演讲者备注提取** | 自动提取 PPT 备注，并支持空置率检测与补充讲稿生成 |
| **智能子图截取** | 对架构图、流程图、数据图表等关键元素自动裁剪子图 |
| **表格恢复** | 优先将表格转录为 Markdown 表格，便于后续检索和比对 |
| **证书精读** | 对证书、专利、软著等知识产权图片进行深度信息提取 |
| **并行处理** | 大规模 PPT（>30 页）支持多 Agent 并行分析，显著提速 |

---

## 📖 使用场景

- 项目验收汇报 PPT 的结构化归档
- 培训课件、技术分享的内容提取与整理
- 商务方案、产品路演的要点梳理
- 会议演示材料的知识库沉淀

---

## 📄 License

MIT
