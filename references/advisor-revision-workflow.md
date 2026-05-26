# 导师批注处理工作流

## 一、批注提取

### 方法一：python-docx

```python
from docx import Document

doc = Document("论文.docx")
comments = doc.comments  # python-docx 1.1+

for comment in comments:
    print(f"批注ID: {comment.id}")
    print(f"作者: {comment.author}")
    print(f"日期: {comment.date}")
    print(f"内容: {comment.text}")
    print(f"引用文本: {comment.quote if hasattr(comment, 'quote') else 'N/A'}")
    print("---")
```

### 方法二：Word COM

```python
import win32com.client

word = win32com.client.Dispatch("Word.Application")
doc = word.Documents.Open("论文.docx")

for comment in doc.Comments:
    print(f"作者: {comment.Author}")
    print(f"内容: {comment.Range.Text}")
    print(f"引用文本: {comment.Scope.Text[:50]}...")
    print("---")
```

### 方法三：OOXML 直接解析

```python
import zipfile
from lxml import etree

with zipfile.ZipFile("论文.docx") as z:
    with z.open("word/comments.xml") as f:
        tree = etree.parse(f)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        for comment in tree.findall(".//w:comment", ns):
            author = comment.get(f"{{{ns['w']}}}author")
            texts = comment.findall(".//w:t", ns)
            content = "".join(t.text for t in texts if t.text)
            print(f"作者: {author}, 内容: {content}")
```

## 二、批注分类框架

| 类别 | 特征 | 处理方式 |
|------|------|----------|
| 文字修改类 | "改为""删除""这里不对" | 安全文本替换 |
| 结构调整类 | "这章要重组""合并这两节" | 章节重组 |
| 内容补充类 | "需要补充""这里太简单" | 标记"待人工补充" |
| 格式修正类 | "字号不对""行距不对" | 批量格式修复 |
| 质疑类 | "这个结论有依据吗？" | 查找支撑数据 |
| 确认类 | "这个图对吗？" | 验证后标记已确认 |

## 三、处理流程

### 步骤1：导出批注清单

生成 Markdown 格式的批注清单：

```markdown
# 导师批注处理清单

| # | 批注内容 | 引用文本 | 分类 | 处理状态 | 修改说明 |
|---|----------|----------|------|----------|----------|
| 1 | "改为XXX" | 原文摘要 | 文字修改 | ✅ 已处理 | 按要求修改 |
| 2 | "需要补充数据来源" | 第3章第1节 | 内容补充 | ⏳ 待人工 | 需要查找原始数据 |
| 3 | "字号不对" | 全文正文 | 格式修正 | ✅ 已处理 | 批量修正为12pt |
```

### 步骤2：逐条处理

- **文字修改**: 使用 python-docx 的 run.text 替换
- **结构调整**: 使用 OOXML 操作段落移动
- **内容补充**: 在对应位置插入 `[待补充：XXX]` 标记
- **格式修正**: 调用格式修正脚本批量处理
- **质疑类**: 查找项目文件中的支撑数据，无法支撑则标注

### 步骤3：生成修改对照表

```markdown
# 修改对照表

## 批注 #1
- **原文**: "研究表明该方法有效"
- **修改后**: "实验结果表明，该模型在测试集上的R²达到0.92"
- **修改原因**: 导师要求具体化，补充数据支撑
- **处理方式**: 文本替换
```

## 四、安全约束

1. **不破坏 Zotero 字段**: 编辑前后计数字段标记
2. **不破坏公式对象**: 不修改含 OLE 公式的段落
3. **不破坏图片**: 不修改含 InlineShape 的段落
4. **保留修改痕迹**: 生成修改对照表，便于导师复核
5. **分批处理**: 每次处理10-20条批注，避免批量错误

## 五、常见批注模板回复

| 批注类型 | 模板回复 |
|----------|----------|
| "改为XXX" | "已按要求修改为：XXX" |
| "补充XXX" | "已补充：[具体内容]" |
| "删除" | "已删除该段落/句子" |
| "这里不对" | "已核实并修正，原因为：XXX" |
| "需要数据支撑" | "已补充数据来源：[出处]" |
| "格式不对" | "已批量修正格式" |

## 六、与导师沟通的注意事项

- 修改完成后，发送修改对照表给导师
- 对于无法处理的批注，明确说明原因
- 对于有疑问的批注，先确认再修改
- 保持修改记录，便于后续追溯
