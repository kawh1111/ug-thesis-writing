# DOCX 安全编辑模式

## 一、双引擎策略

编辑 Word 文档时，根据操作类型选择不同引擎：

| 引擎 | 适用场景 | 优势 | 风险 |
|------|----------|------|------|
| python-docx | 段落级文本编辑、格式修改 | 安全保留字段、跨平台 | 无法刷新目录 |
| Word COM (win32com) | 目录刷新、图表替换、复杂渲染 | 使用 Word 自身引擎 | 仅 Windows、需安装 Word |

## 二、python-docx 安全编辑

### 基本操作流程

```python
from docx import Document
from datetime import datetime
import shutil

# 1. 备份
src = "论文.docx"
backup = f"论文_备份_{datetime.now():%Y%m%d_%H%M%S}.docx"
shutil.copy2(src, backup)

# 2. 打开
doc = Document(src)

# 3. 编辑段落文本（保留格式）
for para in doc.paragraphs:
    if "要替换的文本" in para.text:
        for run in para.runs:
            run.text = run.text.replace("旧文本", "新文本")

# 4. 保存为新文件
output = f"论文_修改_{datetime.now():%Y%m%d_%H%M%S}.docx"
doc.save(output)
```

### Zotero 字段保护

编辑含 Zotero 引用的段落时：

```python
# 编辑前计数
def count_zotero_fields(doc):
    count = 0
    for para in doc.paragraphs:
        if "ADDIN ZOTERO" in para.text:
            count += 1
    for section in doc.sections:
        for header in [section.header, section.footer]:
            if header:
                for para in header.paragraphs:
                    if "ADDIN ZOTERO" in para.text:
                        count += 1
    return count

before_count = count_zotero_fields(doc)
# ... 执行编辑 ...
after_count = count_zotero_fields(doc)

assert before_count == after_count, f"Zotero 字段丢失！编辑前{before_count}个，编辑后{after_count}个"
```

### 安全编辑规则

1. **只修改 run.text，不重建 run**: 保留字体、大小、颜色等格式
2. **不删除含 fldChar 的段落**: 这些是引用字段的组成部分
3. **不修改含图片的段落**: 图片对象与段落绑定
4. **不修改含公式的段落**: 公式对象使用 OLE 嵌入

## 三、Word COM 操作

### 基本用法

```python
import win32com.client
import os

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

doc_path = os.path.abspath("论文.docx")
doc = word.Documents.Open(doc_path)

# 刷新目录
for toc in doc.TablesOfContents:
    toc.Update()

# 保存
doc.Save()
doc.Close()
word.Quit()
```

### 适用场景

- **目录刷新**: python-docx 无法更新页码，必须用 COM
- **图表替换**: 通过 COM 的 InlineShapes 操作
- **样式检查**: COM 可以读取 Word 渲染后的实际样式
- **域代码更新**: 所有域代码（TOC、页码等）需要 COM 刷新

### 注意事项

- 操作前关闭所有 Word 窗口
- 使用 `word.Visible = False` 避免弹窗
- 操作后必须 `doc.Close()` 和 `word.Quit()`
- 异常时也要确保 `word.Quit()`，否则 Word 进程残留

## 四、备份策略

### 命名规范

```
论文_备份_YYYYMMDD_HHMMSS.docx     # 时间戳备份
论文_格式修正前.docx               # 阶段备份
论文_导师修改前.docx               # 里程碑备份
```

### 版本管理

- 每次批量操作前创建备份
- 里程碑节点（提交、修改前）单独备份
- 历史版本归档到 `99_历史阶段版本/` 目录
- 备份文件不覆盖，使用时间戳区分

## 五、格式修正脚本模式

### 批量修正段落格式

```python
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

def fix_paragraph_format(para, config):
    """修正单个段落的格式"""
    # 字号
    for run in para.runs:
        run.font.size = Pt(config["size_pt"])
        run.font.name = config["font_en"]
        # 中文字体需要设置 eastAsia
        run.font.element.rPr.rFonts.set(qn("w:eastAsia"), config["font_cn"])

    # 行距
    para.paragraph_format.line_spacing = config["line_spacing"]

    # 首行缩进
    if config.get("first_indent_chars"):
        para.paragraph_format.first_line_indent = Cm(config["first_indent_chars"] * 0.37)

    # 对齐
    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    if config.get("alignment"):
        para.paragraph_format.alignment = align_map[config["alignment"]]
```

### 识别段落类型

```python
def identify_paragraph_type(para):
    """识别段落类型，返回对应的格式配置"""
    style_name = para.style.name if para.style else ""

    if "Heading 1" in style_name or "标题 1" in style_name:
        return "heading1"
    elif "Heading 2" in style_name or "标题 2" in style_name:
        return "heading2"
    elif "Heading 3" in style_name or "标题 3" in style_name:
        return "heading3"
    elif "TOC" in style_name:
        return "toc"
    elif "Caption" in style_name or "题注" in style_name:
        return "caption"
    else:
        return "body"
```
