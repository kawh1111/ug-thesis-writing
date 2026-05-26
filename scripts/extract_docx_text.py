#!/usr/bin/env python3
"""
从 Word (.docx) 文件中提取带段落编号的纯文本。

用法:
    python extract_docx_text.py <docx_path> [--output output.txt] [--skip-toc] [--skip-headers]

输出格式:
    [P0] 段落文本
    [P1] 段落文本
    ...
    [样式: Heading 1] [P5] 第一章 绪论
"""

import argparse
import sys
from pathlib import Path

try:
    from docx import Document
except ImportError:
    print("错误: 需要安装 python-docx。运行: pip install python-docx", file=sys.stderr)
    sys.exit(1)


def extract_text(docx_path: str, skip_toc: bool = False, skip_headers: bool = False) -> list[dict]:
    """提取 DOCX 文本，返回段落列表。

    Args:
        docx_path: DOCX 文件路径
        skip_toc: 是否跳过目录段落
        skip_headers: 是否跳过页眉页脚

    Returns:
        段落列表，每个元素包含 index, text, style 信息
    """
    doc = Document(docx_path)
    paragraphs = []

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        # 跳过目录
        if skip_toc and para.style and "TOC" in para.style.name:
            continue

        # 跳过页眉页脚样式
        if skip_headers and para.style and any(
            keyword in para.style.name.lower()
            for keyword in ["header", "footer", "页眉", "页脚"]
        ):
            continue

        style_name = para.style.name if para.style else "Normal"

        paragraphs.append({
            "index": i,
            "text": text,
            "style": style_name,
            "is_heading": "Heading" in style_name or "标题" in style_name,
        })

    return paragraphs


def format_output(paragraphs: list[dict], show_style: bool = False) -> str:
    """格式化输出。"""
    lines = []
    for p in paragraphs:
        prefix = f"[样式: {p['style']}] " if show_style or p["is_heading"] else ""
        lines.append(f"[P{p['index']}] {prefix}{p['text']}")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(
        description="从 Word (.docx) 文件中提取带段落编号的纯文本"
    )
    parser.add_argument("docx_path", help="DOCX 文件路径")
    parser.add_argument("--output", "-o", help="输出文件路径（默认输出到终端）")
    parser.add_argument("--skip-toc", action="store_true", help="跳过目录段落")
    parser.add_argument("--skip-headers", action="store_true", help="跳过页眉页脚")
    parser.add_argument("--show-style", action="store_true", help="显示所有段落的样式名")
    parser.add_argument("--stats", action="store_true", help="显示统计信息")

    args = parser.parse_args()

    docx_path = Path(args.docx_path)
    if not docx_path.exists():
        print(f"错误: 文件不存在: {docx_path}", file=sys.stderr)
        sys.exit(1)

    paragraphs = extract_text(
        str(docx_path),
        skip_toc=args.skip_toc,
        skip_headers=args.skip_headers,
    )

    output = format_output(paragraphs, show_style=args.show_style)

    if args.output:
        Path(args.output).write_text(output, encoding="utf-8")
        print(f"已提取 {len(paragraphs)} 个段落到 {args.output}")
    else:
        print(output)

    if args.stats:
        headings = [p for p in paragraphs if p["is_heading"]]
        styles = {}
        for p in paragraphs:
            style = p["style"]
            styles[style] = styles.get(style, 0) + 1

        print(f"\n--- 统计 ---")
        print(f"总段落数: {len(paragraphs)}")
        print(f"标题段落数: {len(headings)}")
        print(f"样式分布:")
        for style, count in sorted(styles.items(), key=lambda x: -x[1]):
            print(f"  {style}: {count}")


if __name__ == "__main__":
    main()
