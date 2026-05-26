# ug-thesis-writing

中文本科毕设全流程辅助写作 Claude Code Skill

A Claude Code skill that guides Chinese undergraduate students through the entire thesis writing process — from project setup to final submission, with a focus on Word (.docx) document handling.

## 这是什么？

一个 Claude Code 自定义技能，帮助中国高校本科生完成毕业论文的全流程写作。覆盖从开题到终稿提交的 9 个阶段，专门针对 Word (.docx) 格式的论文进行了深度优化。

**核心能力**:
- 九阶段标准化写作工作流
- DOCX 格式自动检查与批量修正（字号/行距/缩进/字体）
- Zotero 引用字段安全保护（编辑前后自动验证）
- AIGC 检测率降低策略（确定性替换，非 AI 重写）
- 导师批注结构化处理与修改对照表生成
- 答辩 PPT 结构、讲稿撰写与 Q&A 准备

## Features

- **9-stage workflow**: Project setup → Literature review → Chapter writing → Figure/table management → Format checking → AIGC reduction → Advisor revision → Defense preparation → Final submission
- **DOCX format checking**: Automated detection of font size, line spacing, indentation, and alignment issues
- **DOCX format fixing**: Batch format correction with Zotero field protection
- **AIGC reduction guidance**: Deterministic replacement strategies to reduce AI detection rates
- **Advisor revision workflow**: Structured approach to handling advisor comments
- **Defense preparation**: PPT structure, speech writing, and Q&A preparation

## Installation

```bash
# Clone into your Claude Code skills directory
git clone https://github.com/YOUR_USERNAME/ug-thesis-writing.git ~/.claude/skills/ug-thesis-writing

# Or copy the skill directory
cp -r ug-thesis-writing ~/.claude/skills/
```

## Requirements

- Claude Code CLI
- Python 3.10+
- python-docx (`pip install python-docx`)
- (Optional) win32com for Word COM operations on Windows

## Usage

### In Claude Code

```
# Initialize a new thesis project
/ug-thesis-writing init

# Write a specific chapter
/ug-thesis-writing write 第3章

# Check format
/ug-thesis-writing format 论文.docx

# Fix format issues
/ug-thesis-writing format 论文.docx --fix

# Handle advisor comments
/ug-thesis-writing revise 论文_导师批注版.docx

# AIGC reduction
/ug-thesis-writing reduce-aigc 论文.txt

# Defense preparation
/ug-thesis-writing defense
```

### Standalone Scripts

```bash
# Extract text with paragraph numbers
python scripts/extract_docx_text.py 论文.docx --output text.txt

# Check format
python scripts/check_docx_format.py 论文.docx --output report.md

# Fix format (with preview)
python scripts/apply_format_fixes.py 论文.docx --preview

# Fix format (apply changes)
python scripts/apply_format_fixes.py 论文.docx --config format.json
```

## File Structure

```
ug-thesis-writing/
├── SKILL.md                          # Main skill definition
├── README.md                         # This file
├── LICENSE                           # MIT License
├── references/
│   ├── writing-style-guide.md        # Academic writing style guide
│   ├── format-checklist.md           # Format checking checklist
│   ├── docx-editing-patterns.md      # DOCX safe editing patterns
│   ├── advisor-revision-workflow.md  # Advisor comment handling
│   ├── aigc-reduction-brief.md       # AIGC reduction quick reference
│   ├── defense-prep-guide.md         # Defense preparation guide
│   └── case-study.md                 # Case study from real projects
└── scripts/
    ├── extract_docx_text.py          # DOCX text extraction
    ├── check_docx_format.py          # DOCX format checking
    └── apply_format_fixes.py         # DOCX format fixing
```

## Integration with Other Skills

| Relationship | Skill | Purpose |
|--------------|-------|---------|
| Upstream | ar-literature-search | Literature search |
| Upstream | ar-literature-review | Literature review |
| Parallel | aigc-reduce | AIGC detection rate reduction |
| Parallel | humanizer-zh | Remove AI writing patterns |
| Downstream | ar-slide-generation | Defense PPT |
| Auxiliary | ar-figure-generation | Figure generation |

## Format Configuration

Create a `format.json` file to customize format standards:

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
  }
}
```

## License

MIT License - see [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.

## Acknowledgments

This skill was developed based on real-world experience of using Claude Code to assist in undergraduate thesis writing, including challenges with DOCX formatting, Zotero integration, AIGC detection, and advisor revision workflows.
