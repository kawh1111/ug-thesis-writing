# ug-thesis-writing

**中文理工科本科毕业论文（设计）全流程辅助写作 Claude Code Skill**

一个 Claude Code 自定义技能，帮助中国高校理工科本科生完成毕业论文（设计）的全流程写作。覆盖从开题到终稿提交的 9 个阶段，专门针对 Word (.docx) 格式的论文进行了深度优化。适用于需要做实验、跑模型、画图表、写代码的理工科毕业论文。

## 这是什么？

写本科毕业论文（设计）的时候，你是不是遇到过这些问题：

- 论文格式要求复杂，字号/行距/缩进/字体一项项检查太麻烦
- 导师批注一大堆，逐条修改还要保证不破坏引用和公式
- 用 AI 帮忙写了一部分，结果 AIGC 检测率飙升
- 答辩 PPT 不知道怎么做，讲稿不知道怎么写
- 格式改了好几版，学院要求和学校模板还冲突

这个 skill 就是为了解决这些问题而生的。它把写毕业论文的整个流程标准化为 9 个阶段，每个阶段都有明确的操作指南和自动化工具。

## 核心能力

| 能力 | 说明 |
|------|------|
| 九阶段工作流 | 项目初始化 → 文献综述 → 章节写作 → 图表管理 → 格式修正 → AIGC 降重 → 导师修改 → 答辩准备 → 终稿提交 |
| 格式自动检查 | 检测正文字号、行距、首行缩进、对齐方式、标题层级等格式问题 |
| 格式批量修正 | 一键修正所有格式问题，支持自定义格式配置文件 |
| 引用字段保护 | 编辑前后自动计数 Zotero/EndNote 引用字段，防止引用丢失 |
| AIGC 降重指导 | 基于确定性替换策略，避免 AI 全文重写导致检测率反升 |
| 导师批注处理 | 结构化提取批注、分类处理、生成修改对照表 |
| 答辩准备 | PPT 结构模板、讲稿撰写指南、25+ 常见 Q&A 问题 |

## 安装

```bash
# 克隆到 Claude Code 技能目录
git clone https://github.com/kawh1111/ug-thesis-writing.git ~/.claude/skills/ug-thesis-writing
```

安装后在 Claude Code 中会自动加载，无需额外配置。

## 依赖

- Claude Code CLI
- Python 3.10+
- python-docx（`pip install python-docx`）
- （可选）win32com —— 仅 Windows，用于目录刷新等 Word 高级操作

## 使用方法

### 在 Claude Code 中使用

```
# 初始化新论文项目（读取任务书，建立目录结构）
/ug-thesis-writing init

# 写作指定章节
/ug-thesis-writing write 第3章

# 检查论文格式
/ug-thesis-writing format 论文.docx

# 修正格式问题
/ug-thesis-writing format 论文.docx --fix

# 处理导师批注
/ug-thesis-writing revise 论文_导师批注版.docx

# AIGC 降重
/ug-thesis-writing reduce-aigc 论文.txt

# 答辩准备（PPT + 讲稿 + Q&A）
/ug-thesis-writing defense
```

### 独立脚本使用

即使不通过 Claude Code，你也可以单独使用这三个 Python 脚本：

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

## 九阶段工作流详解

### 阶段一：项目初始化
读取任务书/开题报告，提取硬约束（数据范围、模型、评价指标），建立标准目录结构，创建项目级 CLAUDE.md。

### 阶段二：文献综述
集成 ar-literature-search 和 ar-literature-review 技能，按主题组织综述，指出研究空白。

### 阶段三：章节写作
六章标准结构（绪论/研究区/方法/结果/讨论/结论），每章有独立的写作规则和风格约束。

### 阶段四：图表管理
图表命名规范（图X-Y / 表X-Y），图表内不放标题（由 Word 题注处理），每个图表下方必须有分析段落。

### 阶段五：格式检查与修正
自动检测字号、行距、缩进、对齐等格式问题，按严重程度分级（严重/中等/轻微），支持批量修正。

### 阶段六：AIGC 检测率降低
集成 aigc-reduce 技能，核心原则：绝不用 AI 全文重写，修改率超过 40%，使用确定性替换。

### 阶段七：导师批注处理
提取导师批注，分类处理（文字修改/结构调整/内容补充/格式修正），生成修改对照表。

### 阶段八：答辩准备
PPT 结构模板（20-25页），讲稿撰写（8-10分钟 + 3分钟压缩版），25+ 常见 Q&A 问题。

### 阶段九：终稿提交
提交清单：中英文摘要、目录更新、格式一致性、参考文献、致谢、图表编号、页眉页脚、页码。

## 目录结构

```
ug-thesis-writing/
├── SKILL.md                          # 主技能定义（九阶段工作流）
├── README.md                         # 本文件
├── LICENSE                           # MIT 许可证
├── references/
│   ├── writing-style-guide.md        # 学术写作风格指南（推荐句式、禁止用语、术语标准化）
│   ├── format-checklist.md           # 格式检查清单（字号对照表、格式配置模板）
│   ├── docx-editing-patterns.md      # DOCX 安全编辑模式（python-docx + Word COM）
│   ├── advisor-revision-workflow.md  # 导师批注处理工作流（提取、分类、对照表）
│   ├── aigc-reduction-brief.md       # AIGC 降重快速参考（替换表、句式重构技巧）
│   ├── defense-prep-guide.md         # 答辩准备指南（PPT、讲稿、Q&A）
│   └── case-study.md                 # 写作经验总结（常见问题与解决方案）
└── scripts/
    ├── extract_docx_text.py          # DOCX 文本提取（带段落编号）
    ├── check_docx_format.py          # DOCX 格式检查（生成报告）
    └── apply_format_fixes.py         # DOCX 格式批量修正
```

## 与其他技能的集成

| 关系 | 技能 | 用途 |
|------|------|------|
| 上游 | [ar-literature-search](https://github.com/anthropics/claude-code-skills/tree/main/ar-literature-search) | 文献检索 |
| 上游 | [ar-literature-review](https://github.com/anthropics/claude-code-skills/tree/main/ar-literature-review) | 文献综述 |
| 并行 | [aigc-reduce](https://github.com/anthropics/claude-code-skills/tree/main/aigc-reduce) | AIGC 降重 |
| 并行 | humanizer-zh | 去 AI 味 |
| 下游 | ar-slide-generation | 答辩 PPT |
| 辅助 | ar-figure-generation | 图表生成 |

## 格式配置

创建 `format.json` 文件来自定义格式标准（以常见理工科论文格式为例）：

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

不同学校的格式要求可能不同，修改配置文件即可适配。

## 常见问题

### Q: 这个 skill 和 aigc-reduce 有什么区别？
aigc-reduce 专门做 AIGC 降重，ug-thesis-writing 覆盖毕业论文写作的全流程，降重只是其中一个阶段。ug-thesis-writing 会调用 aigc-reduce 来完成降重工作。

### Q: 支持 LaTeX 论文吗？
本 skill 专注于 Word (.docx) 格式，这是国内理工科本科毕业论文（设计）最常用的格式。如果你用 LaTeX 写论文，建议使用 ar-paper-writing-section 和 ar-latex-formatting 等技能。

### Q: 支持哪些学校？
理论上支持所有中国高校的理工科本科毕业论文（设计）。格式检查通过配置文件适配，你可以根据学校的格式要求修改 `format.json`。

### Q: 需要联网吗？
脚本不需要联网。Claude Code 的文献检索功能需要联网。

## 许可证

MIT License - 详见 [LICENSE](LICENSE)

## 贡献

欢迎提交 Issue 和 Pull Request。
