# ug-thesis-writing

<p align="center">
  <b>中文理工科本科毕业论文（设计）全流程辅助写作 Claude Code Skill</b>
</p>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License: MIT"></a>
  <img src="https://img.shields.io/badge/Python-3.10+-yellow.svg" alt="Python 3.10+">
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg" alt="Platform">
  <img src="https://img.shields.io/badge/Word-.docx-2B579A.svg" alt="Word .docx">
</p>

---

一个 Claude Code 自定义技能，帮助中国高校理工科本科生完成毕业论文（设计）的全流程写作。覆盖从开题到终稿提交的 **9 个阶段**，专门针对 Word (.docx) 格式进行了深度优化。

> 适用于需要做实验、跑模型、画图表、写代码的理工科毕业论文（设计）。

## 目录

- [解决的问题](#解决的问题)
- [核心能力](#核心能力)
- [九阶段工作流](#九阶段工作流)
- [安装](#安装)
- [使用方法](#使用方法)
- [格式配置](#格式配置)
- [目录结构](#目录结构)
- [与其他技能的集成](#与其他技能的集成)
- [常见问题](#常见问题)
- [许可证](#许可证)

---

## 解决的问题

写本科毕业论文（设计）的时候，你是不是遇到过这些问题：

| 痛点 | 本 Skill 的解决方案 |
|:-----|:-------------------|
| 格式要求复杂，字号/行距/缩进一项项检查太麻烦 | 自动格式检查 + 批量修正 |
| 导师批注一大堆，逐条修改还要保证不破坏引用 | 结构化批注处理 + Zotero 字段保护 |
| 用 AI 帮忙写了一部分，AIGC 检测率飙升 | 确定性替换降重，非 AI 重写 |
| 答辩 PPT 不知道怎么做，讲稿不知道怎么写 | PPT 模板 + 讲稿 + 25 个 Q&A |
| 格式改了好几版，学院要求和学校模板还冲突 | 可配置格式标准，一键适配 |

---

## 核心能力

| 能力 | 说明 |
|:-----|:-----|
| **九阶段工作流** | 项目初始化 → 文献综述 → 章节写作 → 图表管理 → 格式修正 → AIGC 降重 → 导师修改 → 答辩准备 → 终稿提交 |
| **格式自动检查** | 检测正文字号、行距、首行缩进、对齐方式、标题层级等格式问题 |
| **格式批量修正** | 一键修正所有格式问题，支持自定义格式配置文件 |
| **引用字段保护** | 编辑前后自动计数 Zotero/EndNote 引用字段，防止引用丢失 |
| **AIGC 降重指导** | 基于确定性替换策略，避免 AI 全文重写导致检测率反升 |
| **导师批注处理** | 结构化提取批注、分类处理、生成修改对照表 |
| **答辩准备** | PPT 结构模板、讲稿撰写指南、25+ 常见 Q&A 问题 |

---

## 九阶段工作流

```
┌─────────┐   ┌─────────┐   ┌─────────┐   ┌─────────┐   ┌─────────┐
│ 1. 初始化 │──▶│ 2. 综述  │──▶│ 3. 写作  │──▶│ 4. 图表  │──▶│ 5. 格式  │
└─────────┘   └─────────┘   └─────────┘   └─────────┘   └─────────┘
                                                                │
┌─────────┐   ┌─────────┐   ┌─────────┐   ┌─────────┐         │
│ 9. 终稿  │◀──│ 8. 答辩  │◀──│ 7. 修改  │◀──│ 6. 降重  │◀────────┘
└─────────┘   └─────────┘   └─────────┘   └─────────┘
```

| 阶段 | 名称 | 主要工作 |
|:----:|:-----|:---------|
| 1 | 项目初始化 | 读取任务书，提取硬约束，建立目录结构 |
| 2 | 文献综述 | 检索、阅读、按主题组织综述 |
| 3 | 章节写作 | 六章标准结构，推荐顺序：方法→结果→讨论→绪论→结论 |
| 4 | 图表管理 | 命名规范（图X-Y / 表X-Y），每个图表配分析段落 |
| 5 | 格式检查与修正 | 自动检测 + 批量修正字号、行距、缩进、对齐 |
| 6 | AIGC 降重 | 确定性替换，修改率 >40%，绝不用 AI 全文重写 |
| 7 | 导师批注处理 | 提取批注 → 分类处理 → 生成修改对照表 |
| 8 | 答辩准备 | PPT（20-25 页）+ 讲稿（8-10 分钟）+ Q&A |
| 9 | 终稿提交 | 最终检查清单，确保格式、引用、页码无误 |

---

## 安装

```bash
# 克隆到 Claude Code 技能目录
git clone https://github.com/kawh1111/ug-thesis-writing.git ~/.claude/skills/ug-thesis-writing
```

安装后在 Claude Code 中会自动加载，无需额外配置。

### 依赖

- Claude Code CLI
- Python 3.10+
- `python-docx`（`pip install python-docx`）
- （可选）`win32com` —— 仅 Windows，用于目录刷新等 Word 高级操作

---

## 使用方法

### 在 Claude Code 中使用

```
/ug-thesis-writing init                  # 初始化新论文项目
/ug-thesis-writing write 第3章           # 写作指定章节
/ug-thesis-writing format 论文.docx      # 检查论文格式
/ug-thesis-writing format 论文.docx --fix # 修正格式问题
/ug-thesis-writing revise 论文_导师批注版.docx  # 处理导师批注
/ug-thesis-writing reduce-aigc 论文.txt  # AIGC 降重
/ug-thesis-writing defense               # 答辩准备
```

### 独立脚本使用

即使不通过 Claude Code，也可以单独使用 Python 脚本：

```bash
# 提取 DOCX 文本（带段落编号，方便定位）
python scripts/extract_docx_text.py 论文.docx --output text.txt

# 检查格式（生成 Markdown 报告）
python scripts/check_docx_format.py 论文.docx --output report.md

# 预览需要修正的格式（不实际修改）
python scripts/apply_format_fixes.py 论文.docx --preview

# 执行格式修正（生成新文件，不覆盖原文件）
python scripts/apply_format_fixes.py 论文.docx --config format.json
```

---

## 格式配置

创建 `format.json` 文件来自定义格式标准。以下为常见理工科论文格式示例：

```json
{
  "body": {
    "font_cn": "宋体",
    "font_en": "Times New Roman",
    "size_pt": 12,
    "line_spacing": 1.5,
    "first_indent_chars": 2,
    "alignment": "both"
  },
  "heading1": {
    "font_cn": "黑体",
    "size_pt": 16,
    "bold": true,
    "alignment": "center"
  },
  "heading2": {
    "font_cn": "黑体",
    "size_pt": 14,
    "bold": true,
    "alignment": "left"
  },
  "heading3": {
    "font_cn": "宋体",
    "size_pt": 12,
    "bold": false,
    "alignment": "left"
  },
  "caption": {
    "font_cn": "宋体",
    "size_pt": 10.5,
    "alignment": "center"
  },
  "page": {
    "paper": "A4",
    "margin_top_cm": 2.54,
    "margin_bottom_cm": 2.54,
    "margin_left_cm": 3.17,
    "margin_right_cm": 3.17
  }
}
```

> 不同学校的格式要求可能不同，修改配置文件即可适配。

---

## 目录结构

```
ug-thesis-writing/
├── SKILL.md                          # 主技能定义（九阶段工作流）
├── README.md                         # 本文件
├── LICENSE                           # MIT 许可证
├── references/
│   ├── writing-style-guide.md        # 学术写作风格指南
│   ├── format-checklist.md           # 格式检查清单
│   ├── docx-editing-patterns.md      # DOCX 安全编辑模式
│   ├── advisor-revision-workflow.md  # 导师批注处理工作流
│   ├── aigc-reduction-brief.md       # AIGC 降重快速参考
│   ├── defense-prep-guide.md         # 答辩准备指南
│   └── case-study.md                 # 写作经验总结
└── scripts/
    ├── extract_docx_text.py          # DOCX 文本提取（带段落编号）
    ├── check_docx_format.py          # DOCX 格式检查（生成报告）
    └── apply_format_fixes.py         # DOCX 格式批量修正
```

---

## 与其他技能的集成

| 关系 | 技能 | 用途 |
|:-----|:-----|:-----|
| 上游 | ar-literature-search | 文献检索 |
| 上游 | ar-literature-review | 文献综述 |
| 并行 | aigc-reduce | AIGC 降重 |
| 并行 | humanizer-zh | 去 AI 味 |
| 下游 | ar-slide-generation | 答辩 PPT |
| 辅助 | ar-figure-generation | 图表生成 |

---

## 常见问题

<details>
<summary><b>这个 skill 和 aigc-reduce 有什么区别？</b></summary>

aigc-reduce 专门做 AIGC 降重，ug-thesis-writing 覆盖毕业论文写作的全流程，降重只是其中一个阶段。ug-thesis-writing 会调用 aigc-reduce 来完成降重工作。
</details>

<details>
<summary><b>支持 LaTeX 论文吗？</b></summary>

本 skill 专注于 Word (.docx) 格式，这是国内理工科本科毕业论文（设计）最常用的格式。如果你用 LaTeX 写论文，建议使用 ar-paper-writing-section 和 ar-latex-formatting 等技能。
</details>

<details>
<summary><b>支持哪些学校？</b></summary>

理论上支持所有中国高校的理工科本科毕业论文（设计）。格式检查通过配置文件适配，你可以根据学校的格式要求修改 `format.json`。
</details>

<details>
<summary><b>需要联网吗？</b></summary>

脚本不需要联网。Claude Code 的文献检索功能需要联网。
</details>

---

## 许可证

[MIT License](LICENSE) - 欢迎提交 Issue 和 Pull Request。
