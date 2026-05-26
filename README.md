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

写毕业论文的时候，格式改了八百遍、导师批注一大堆、AIGC 检测率飙红、答辩 PPT 不知道从哪下手——这些事我都经历过。这个 skill 就是把踩过的坑和解决方案整理成了标准化流程，让 Claude Code 能帮你把论文从开题到终稿跑完。

适合理工科毕业论文（设计），尤其是需要做实验、跑模型、画图表的那种。基于 Word (.docx) 格式。

## Quick Start

**安装**

```bash
git clone https://github.com/kawh1111/ug-thesis-writing.git ~/.claude/skills/ug-thesis-writing
```

**在 Claude Code 里用**

```
/ug-thesis-writing init              # 初始化项目
/ug-thesis-writing write 第3章        # 写作章节
/ug-thesis-writing format 论文.docx   # 检查格式
```

## 目录

- [它能干什么](#它能干什么)
- [工作流](#工作流)
- [安装](#安装)
- [使用方法](#使用方法)
- [格式配置](#格式配置)
- [项目结构](#项目结构)
- [相关技能](#相关技能)
- [常见问题](#常见问题)

## 它能干什么

**格式检查和修正** — 自动检测字号、行距、缩进、对齐这些问题，一键批量修正。不用再一行一行对着格式手册检查了。支持自定义格式配置，学院要求和学校模板不一样的话改下 JSON 就行。

**Zotero 引用保护** — 用 python-docx 改论文的时候，引用字段很容易被破坏。这个 skill 在编辑前后会自动计数 Zotero 字段，丢了就报错。

**AIGC 降重** — 核心原则：不用 AI 全文重写（重写反而会让检测率更高）。用确定性替换 + 句式重构，改够 40% 以上。

**导师批注处理** — 从 Word 文档里提取批注，按类型分类（文字修改 / 结构调整 / 内容补充 / 格式修正），生成修改对照表。

**答辩准备** — PPT 结构模板、逐页讲稿、25 个以上常见 Q&A 问题。

## 工作流

9 个阶段，不用全部走完，缺哪个用哪个。

```mermaid
graph LR
    A[1. 初始化] --> B[2. 综述]
    B --> C[3. 写作]
    C --> D[4. 图表]
    D --> E[5. 格式]
    E --> F[6. 降重]
    F --> G[7. 导师修改]
    G --> H[8. 答辩]
    H --> I[9. 终稿]
```

| 阶段 | 干什么 |
|:----:|:------|
| 1 | 读任务书，提取硬约束，建目录结构 |
| 2 | 文献检索和综述写作 |
| 3 | 按六章结构写，建议顺序：方法→结果→讨论→绪论→结论 |
| 4 | 图表命名规范，每个图/表下面写分析段落 |
| 5 | 自动检查 + 批量修格式 |
| 6 | 确定性替换降重，不走 AI 重写 |
| 7 | 提取批注，逐条处理，出修改对照表 |
| 8 | PPT + 讲稿 + Q&A |
| 9 | 最终检查，确认格式、引用、页码都没问题 |

## 安装

```bash
git clone https://github.com/kawh1111/ug-thesis-writing.git ~/.claude/skills/ug-thesis-writing
```

装完自动加载，不用额外配置。

依赖：
- Claude Code CLI
- Python 3.10+
- `python-docx`（`pip install python-docx`）
- 可选：`win32com`（仅 Windows，用来刷新目录等 Word 高级操作）

## 使用方法

### Claude Code 里直接用

```
/ug-thesis-writing init                        # 初始化
/ug-thesis-writing write 第3章                  # 写作
/ug-thesis-writing format 论文.docx             # 检查格式
/ug-thesis-writing format 论文.docx --fix       # 修正格式
/ug-thesis-writing revise 论文_导师批注版.docx   # 处理批注
/ug-thesis-writing reduce-aigc 论文.txt         # 降重
/ug-thesis-writing defense                      # 答辩准备
```

### 脚本单独用

不装 Claude Code 也能跑：

```bash
# 提取文本（带段落编号，方便定位问题在哪）
python scripts/extract_docx_text.py 论文.docx --output text.txt

# 检查格式，生成 Markdown 报告
python scripts/check_docx_format.py 论文.docx --output report.md

# 先预览要改哪些，不实际改
python scripts/apply_format_fixes.py 论文.docx --preview

# 正式修正（生成新文件，原文件不动）
python scripts/apply_format_fixes.py 论文.docx --config format.json
```

## 格式配置

学校格式不一样？改 `format.json` 就行。下面是一个常见的理工科论文配置，直接抄然后按你们学校的要求改：

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

## 项目结构

```
ug-thesis-writing/
├── SKILL.md                          # 主技能定义
├── README.md                         # 你在看的这个
├── LICENSE                           # MIT
├── references/
│   ├── writing-style-guide.md        # 写作风格指南（推荐句式、禁用词）
│   ├── format-checklist.md           # 格式检查清单
│   ├── docx-editing-patterns.md      # DOCX 安全编辑（python-docx + Word COM）
│   ├── advisor-revision-workflow.md  # 导师批注处理流程
│   ├── aigc-reduction-brief.md       # AIGC 降重参考
│   ├── defense-prep-guide.md         # 答辩准备指南
│   └── case-study.md                 # 踩坑经验总结
└── scripts/
    ├── extract_docx_text.py          # 提取文本
    ├── check_docx_format.py          # 检查格式
    └── apply_format_fixes.py         # 修正格式
```

## 相关技能

这个 skill 可以和其他 Claude Code 技能配合用：

- [ar-literature-search](https://github.com/anthropics/claude-code-skills/tree/main/ar-literature-search) — 文献检索
- [ar-literature-review](https://github.com/anthropics/claude-code-skills/tree/main/ar-literature-review) — 文献综述
- [aigc-reduce](https://github.com/anthropics/claude-code-skills/tree/main/aigc-reduce) — AIGC 降重（本 skill 的降重阶段会调用它）
- [ar-slide-generation](https://github.com/anthropics/claude-code-skills/tree/main/ar-slide-generation) — 答辩 PPT
- [ar-figure-generation](https://github.com/anthropics/claude-code-skills/tree/main/ar-figure-generation) — 图表生成

## 常见问题

**和 aigc-reduce 有什么区别？**

aigc-reduce 只管降重。这个 skill 覆盖论文写作全流程，降重是其中一个环节，会调用 aigc-reduce 来做。

**支持 LaTeX 吗？**

不支持，只做 Word (.docx)。LaTeX 论文看 ar-paper-writing-section 和 ar-latex-formatting。

**哪些学校能用？**

理论上都能用。格式检查靠配置文件适配，把你们学校的要求填进 `format.json` 就行。

**要联网吗？**

脚本不用。文献检索功能要联网。

---

<p align="center">
  如果这个项目对你有帮助，给个 Star 支持一下
</p>
