<a id="readme-top"></a>

<div align="center">

# ppt-analyzer

### AI-native skill that turns PPTX and PDF slides into structured Markdown with vision AI

WPS exports every page as PNG; the AI reads the images and rebuilds structure, tables and speaker notes.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![GitHub stars](https://img.shields.io/github/stars/kuwork/ppt-analyzer?logo=github)](https://github.com/kuwork/ppt-analyzer/stargazers)
[![GitHub issues](https://img.shields.io/github/issues/kuwork/ppt-analyzer?logo=github)](https://github.com/kuwork/ppt-analyzer/issues)
[![Platform: Windows](https://img.shields.io/badge/Platform-Windows-0078D4?logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![Python: optional scripts](https://img.shields.io/badge/Python-optional%20scripts-3776AB?logo=python&logoColor=white)](https://www.python.org/)

English ｜ [简体中文](README.zh-CN.md)

[Features](#features) ·
[How It Works](#how-it-works) ·
[Quick Start](#quick-start) ·
[Installation Options](#installation-options) ·
[FAQ](#faq) ·
[Known Limitations](#known-limitations) ·
[Project Structure](#project-structure)

</div>

**ppt-analyzer** is an AI-native skill for analyzing `.pptx` and `.pdf` slide decks on Windows. Instead of driving Office through COM components, it follows a two-stage workflow: you manually export every slide to PNG with WPS Office, then a vision-capable AI reads the images and produces structured Markdown — a slide index with per-page summaries, extracted speaker notes, Markdown table transcripts and an optional overall report.

> [!NOTE]
> Repository: [github.com/kuwork/ppt-analyzer](https://github.com/kuwork/ppt-analyzer) ｜ Issues: [github.com/kuwork/ppt-analyzer/issues](https://github.com/kuwork/ppt-analyzer/issues) ｜ Skill entry: `skills/ppt-analyzer/SKILL.md` ｜ Supported formats: `.pptx` · `.pdf`

## Features

- 🧠 **AI-native skill, no COM dependency** — the skill is an instruction set (`SKILL.md`) loaded directly by an AI host; the core export-and-analysis path needs no Office COM component or third-party parsing library.
- 🖼️ **WPS-powered image pipeline** — slides are exported page by page as PNG with WPS Office, so image quality is controlled by WPS and the flow runs across varied Windows environments.
- 📄 **PPTX and PDF in one flow** — the same process covers both `.pptx` and `.pdf` decks; PDFs, which carry no speaker notes, are detected and handled automatically.
- 🗣️ **Full speaker-notes lifecycle** — `extract_notes.py` pulls notes via python-pptx, the vacancy rate is measured (a 30% vacancy triggers full regeneration), the AI drafts 50–200-character talking points, and `write_notes.py` writes them back into a copy of the deck.
- ✂️ **Sub-image cropping for diagrams** — architecture diagrams, flowcharts and data charts are cropped with PIL into `assets/crops/` and linked from the corresponding page summary.
- 📊 **Tables reconstructed as Markdown** — regular tables (up to 20 rows × 8 columns) are transcribed into Markdown tables for later search and comparison; only complex merged-cell tables fall back to prose.
- 🏅 **Close reading of certificates** — certificates, patents and software copyrights are examined one by one for title, issuing body, owner, date and registration number.
- ⚡ **Parallel analysis for large decks** — decks over 30 pages are split into ranges of 10–15 pages and processed by up to four parallel sub-agents, which return summaries to the main agent instead of writing files concurrently.

## How It Works

The skill uses a two-stage "manual export + AI analysis" flow. Avoiding Office COM components and third-party parsing libraries keeps it compatible with varied Windows environments, while WPS directly guarantees the quality of the exported images.

| Step | Stage | What happens |
|---|---|---|
| 1 | Setup | Create a folder named after the deck in the repository root, with an `assets/` subfolder |
| 2 | Export | In WPS Presentation (PPTX) or WPS PDF, export every page as PNG into `assets/` |
| 3 | Rename | If WPS generates overlong filenames, normalize them to `幻灯片_XX.png` |
| 4 | Notes, PPTX only | Run `extract_notes.py` to save speaker notes and page text to `备注提取.md`; PDFs get an empty placeholder |
| 5 | Index | Build `幻灯片索引.md`, listing image paths, notes and per-page recognition summaries |
| 6 | Vision analysis | The AI reads every image in reading order, classifies elements, transcribes tables, crops diagrams and close-reads certificates |
| 7 | Notes gap-fill | Compute the vacancy rate; empty pages receive AI-generated talking points, optionally written back into the PPTX |
| 8 | Report, optional | On request, produce `分析结果.md` with chapter structure, key conclusions and a chart index |

Each analyzed deck produces this working folder in the repository root:

```text
我的汇报PPT/
├── assets/                  # Slide PNGs exported from WPS
│   └── crops/               # Cropped sub-images (diagrams, flowcharts, charts)
├── 幻灯片索引.md             # Image paths + speaker notes + AI recognition summaries
├── 备注提取.md               # Full backup of speaker notes and page text
└── 分析结果.md               # Overall analysis report (optional)
```

## Quick Start

### Prerequisites

- Windows
- WPS Office installed, with the Presentation and PDF components available
- The target `.pptx` or `.pdf` file openable locally

### Install the skill (recommended: user-level skills directory)

Kimi Code CLI auto-loads skills from `~/.config/agents/skills/`; the same skill folder is also recognized under `~/.kimi/skills/` and `~/.claude/skills/`.

```bash
# 1. Enter your user-level skills directory
cd ~/.config/agents/skills/   # or ~/.kimi/skills/ or ~/.claude/skills/

# 2. Clone the repository to a temporary location
git clone https://github.com/kuwork/ppt-analyzer.git /tmp/ppt-analyzer

# 3. Copy the skill folder into place
cp -r /tmp/ppt-analyzer/skills/ppt-analyzer ./ppt-analyzer

# 4. Clean up the temporary clone
rm -rf /tmp/ppt-analyzer
```

The resulting layout is auto-detected and loaded by the AI host:

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
> Verify the install by confirming that `~/.config/agents/skills/ppt-analyzer/SKILL.md` exists, then ask the host to "analyze a PPT" — the skill triggers automatically. Project-level, npm and ad-hoc installs are described in [Installation Options](#installation-options).

### Run the analysis

1. Create the deck-named folder and its `assets/` subfolder.
2. Export slides as PNG: WPS Presentation → File → Export as Images for PPTX, or WPS PDF → File → Export PDF as Images for PDF. Choose PNG and page-by-page output.
3. PPTX only: run `python scripts/extract_notes.py <pptx-path>` to extract speaker notes (`pip install python-pptx` is required once beforehand).
4. Let the AI read the images in `assets/` and produce per-page recognition summaries.
5. The AI assembles `幻灯片索引.md`, fills empty notes with generated talking points, and optionally writes `分析结果.md`.

The complete step-by-step playbook, script templates and best practices live in [`skills/ppt-analyzer/SKILL.md`](skills/ppt-analyzer/SKILL.md).

## Installation Options

<details>
<summary><b>Project-level install (current project only)</b></summary>

```bash
# Run in the project root
mkdir -p .agents/skills

git clone https://github.com/kuwork/ppt-analyzer.git /tmp/ppt-analyzer
cp -r /tmp/ppt-analyzer/skills/ppt-analyzer .agents/skills/
rm -rf /tmp/ppt-analyzer
```

</details>

<details>
<summary><b>Install from GitHub with npm</b></summary>

The package is not published to the npm registry, but npm can install it directly from GitHub:

```bash
npm install kuwork/ppt-analyzer
```

The skill lands in `node_modules/ppt-analyzer/skills/ppt-analyzer/`. Copy it into a skills directory afterward:

```bash
cp -r node_modules/ppt-analyzer/skills/ppt-analyzer ~/.config/agents/skills/
```

</details>

<details>
<summary><b>Ad-hoc use: let the AI read the skill directly</b></summary>

```text
读取 https://github.com/kuwork/ppt-analyzer/blob/main/skills/ppt-analyzer/SKILL.md
```

Nothing is installed; the AI follows the instructions for that conversation only.

</details>

## Use Cases

- Structured archiving of project acceptance and review presentations
- Content extraction from training courseware and technical sharing decks
- Distilling key points from business proposals and product roadshow materials
- Turning meeting presentations into searchable knowledge-base assets

## FAQ

### I cannot find the "Export as Images" menu in WPS. What do I do?

Use File → Save As → Export as Images, or search for "输出为图片" in the WPS search box. In WPS PDF, the equivalent entry is Convert → PDF to Image.

### Do I need Python or Node.js for the core workflow?

No. The export-and-vision workflow only needs WPS plus the AI host. Python is required solely by the optional helper scripts: run `pip install python-pptx` for note extraction and write-back; the all-in-one `extract_ppt.py` additionally requires Aspose.Slides.

### Why are all speaker notes empty in a PDF deck?

PDF slides do not carry speaker notes, so a 100% vacancy rate is expected. The skill detects this and generates talking points for every page based on the recognition summaries and visible page text.

### python-pptx cannot read notes from my WPS-saved PPTX. How do I fix it?

Re-save the deck in WPS with Save As into a standard `.pptx` file, then run `extract_notes.py` again. WPS occasionally stores notes in a way that differs slightly from standard Office.

### The text in exported slide images is blurry.

Re-export at a higher output size or DPI: choose at least twice the default size in WPS Presentation, or HD / Ultra-HD quality in WPS PDF. Crisp text is the prerequisite for accurate table recovery and certificate reading.

### How are decks with more than 30 pages handled?

They are split into four or five consecutive ranges of 10–15 pages. Up to four sub-agents analyze the ranges in parallel and return text only; the main agent then writes `幻灯片索引.md` once, in page order, avoiding concurrent-write conflicts.

### Is writing generated notes back into the PPTX safe?

`write_notes.py` never overwrites the source by default — output goes to a new file named `<deck>_已更新备注.pptx`. Running it on a copy is still the recommended workflow, and it should be used only when the user explicitly asks for note write-back.

### Can I analyze only a subset of pages?

Yes. The WPS export dialog accepts a page-range selection, so export just the pages you need and run the same flow on them.

## Known Limitations

- **Windows and WPS Office required** — the documented export flow relies on WPS Presentation and WPS PDF; a macOS/Linux path is not covered.
- **Export is a manual step** — by design, a human runs the WPS export; the main workflow has no headless, fully automatic capture.
- **Optional Python dependencies** — the note scripts need python-pptx; the all-in-one `extract_ppt.py` additionally needs Aspose.Slides.
- **Vision accuracy depends on export resolution** — dense tables and small certificate text require high-resolution PNGs.
- **Note write-back is PPTX-only** — PDFs have no speaker-note container, so generated talking points stay in Markdown.

## Security

> [!CAUTION]
> Every helper script works with local files only — none makes network requests or handles credentials. `write_notes.py` saves to a new `<deck>_已更新备注.pptx` file instead of overwriting the source. Audit the three scripts under `skills/ppt-analyzer/scripts/` before running them, and keep sensitive decks on your local machine.

## Project Structure

<details>
<summary>Show repository layout</summary>

```text
ppt-analyzer/
├── skills/
│   └── ppt-analyzer/
│       ├── SKILL.md                # Full instruction set for the AI host
│       ├── scripts/
│       │   ├── extract_notes.py    # Extract notes, text and tables via python-pptx
│       │   ├── extract_ppt.py      # All-in-one export: Aspose.Slides images/notes + python-pptx text
│       │   └── write_notes.py      # Batch-write generated notes into a PPTX copy
│       └── references/
│           └── sed-coordinate-positioning-multiline-replace-insert-cheatsheet.md
├── .gitignore
├── LICENSE
├── package.json
└── README.md
```

</details>

## License

Released under the [MIT License](LICENSE).

<p align="right"><a href="#readme-top">↑ Back to top</a></p>
