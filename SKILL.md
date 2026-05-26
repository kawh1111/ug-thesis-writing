---
name: ug-thesis-writing
description: >
  辅助中国高校本科生完成毕业论文全流程写作。
  覆盖：项目初始化 → 文献综述 → 章节写作 → 图表管理 → 格式修正 → AIGC 降重 → 导师批注修改 → 答辩准备 → 终稿提交。
  适用于 Word (.docx) 格式论文。
  Use when 用户提到写论文、毕业论文、毕设、本科论文、论文修改、格式修正、答辩准备、论文降重、AIGC 降重等。
argument-hint: "[action] [file]"
---

# 本科毕业论文全流程辅助写作 (UG Thesis Writing)

## Input

- `$0` — 动作（write / format / revise / reduce-aigc / defense / check）
- `$1` — 文件路径或章节名
- `$ARGUMENTS` — 所有参数的原始字符串

---

## 铁律 (Iron Rules)

1. **绝不伪造**: 数据、实验结果、图表、文献、DOI、模型性能、结论——全部必须有真实文件支撑。缺失内容标注"待补充"。
2. **真实优先**: 项目文件 > 实验日志 > 任务书/开题报告 > 学校要求 > 用户指令。
3. **备份先行**: 任何 DOCX 编辑操作前，必须先创建带时间戳的备份副本。
4. **字段保护**: 编辑含 Zotero/EndNote 引用字段的段落时，前后计数字段标记，验证未丢失。
5. **确定性编辑**: 优先使用确定性字符串替换，避免 AI 全文重写（特别是 AIGC 降重场景）。

---

## 九阶段工作流 (9-Stage Workflow)

### 阶段一：项目初始化 (Project Setup)

**触发**: 用户开始新论文项目

**步骤**:
1. 读取任务书/开题报告，提取硬约束：
   - 数据时间范围
   - 研究方法/模型
   - 评价指标
   - 论文章节要求
2. 建议目录结构：
   ```
   项目根目录/
   ├── 01_论文主线/
   │   ├── 00_当前写作版本/
   │   ├── 01_写作依据与要求/
   │   └── 99_历史阶段版本/
   ├── 02_项目工程/
   ├── 03_原始数据/
   ├── 04_参考资料归档/
   └── outputs/
   ```
3. 创建项目级 CLAUDE.md，记录：
   - 论文标题、作者、学校、专业、导师
   - 硬约束（数据范围、模型、指标）
   - source-of-truth 优先级
   - 术语标准化表

---

### 阶段二：文献综述 (Literature Review)

**技能集成**:
- 文献检索: 调用 [ar-literature-search](../ar-literature-search/)
- 文献综述: 调用 [ar-literature-review](../ar-literature-review/)
- 引用管理: 调用 [ar-citation-management](../ar-citation-management/)

**中文综述写作规范**:
- 按主题组织，不按时间罗列
- 每段以研究者工作为主体（"张三(2020)研究了..."而非"在2020年，有研究..."）
- 结尾必须指出研究空白
- 禁止"众所周知""不言而喻"等无引用断言

---

### 阶段三：章节写作 (Chapter Writing)

**六章标准结构**:

| 章 | 标题 | 写作要点 |
|----|------|----------|
| 1 | 绪论 | 研究背景→国内外现状→研究内容→技术路线 |
| 2 | 研究区概况 | 纯事实描述，使用项目生成的地图和数据 |
| 3 | 研究方法与数据 | 数据来源→预处理→特征工程→模型→评价指标 |
| 4 | 结果分析 | 只报告已完成实验，使用真实指标文件 |
| 5 | 讨论 | 最优模型+预测方案+预测结果+局限性 |
| 6 | 结论与展望 | 结论对应实际结果，展望要现实 |

**写作风格**:
- 使用: "本研究""结果表明""进一步分析发现""表X-Y列出了..."
- 禁止: "首次提出""颠覆性""降维打击""吊打""黑科技"
- 禁止: 博客体、产品推销体、AI 演示体
- 术语统一: 全文使用同一套术语（如选定"研究区"后不要混用"研究区域"）

**写作顺序建议**:
1. 第3章（方法）→ 第4章（结果）→ 第5章（讨论）
2. 第2章（研究区）→ 第1章（绪论）
3. 第6章（结论）→ 摘要

---

### 阶段四：图表管理 (Figure & Table Management)

**规范**:
- 命名: 图X-Y / 表X-Y（X=章号, Y=序号）
- 图内不放标题，由 Word 题注处理
- 分辨率: ≥300 DPI（出版质量）
- 格式: PNG（位图）或 SVG（矢量图）

**分析段落结构**（每个图表下方必须有分析段落）:
1. 说明内容: "图X-Y展示了..."
2. 解释模式: "从图中可以看出..."
3. 讨论含义: "这一结果表明..."

**禁止**: 逐字重复图表标题作为段落内容

---

### 阶段五：格式检查与修正 (Format Check & Fix)

**调用脚本**:
```bash
python ~/.claude/skills/ug-thesis-writing/scripts/check_docx_format.py <docx_path> [--config format.json]
python ~/.claude/skills/ug-thesis-writing/scripts/apply_format_fixes.py <docx_path> [--config format.json]
```

**常见格式问题**（按严重程度）:

| 级别 | 问题 | 典型表现 |
|------|------|----------|
| 严重 | 正文字号错误 | 9pt（小五号）→ 应为12pt（小四号） |
| 严重 | 行距不一致 | 1.0/1.2倍 → 应为1.5倍 |
| 中等 | 首行缩进缺失 | 大量段落无缩进 |
| 中等 | 标题层级冲突 | 学院要求与学校模板不一致 |
| 轻微 | 页边距不一致 | 不同节的页边距不同 |
| 轻微 | 对齐方式缺失 | 部分段落未设置两端对齐 |

**安全编辑策略**:
- python-docx: 段落级安全编辑，保留 Zotero 字段
- Word COM (win32com): 目录刷新、图表替换等需要 Word 渲染的操作
- 详见 [references/docx-editing-patterns.md](references/docx-editing-patterns.md)

---

### 阶段六：AIGC 检测率降低 (AIGC Reduction)

**技能集成**: 调用 [aigc-reduce](../aigc-reduce/) 获取完整降重工作流

**三条铁律**:
1. 绝不用 AI 全文重写（重写反而增加 AI 痕迹）
2. 修改率必须超过 40%
3. 确定性替换，非 LLM 重写

**操作顺序**: 扫描定位 → 词汇替换(10-15%) → 句式重构(15-20%) → 段落调整(10-15%) → 验证修改率

**也可调用**: [humanizer-zh](../humanizer-zh/) 进行更细粒度的去 AI 味处理

**检测平台注意事项**:
- PaperPure: 极敏感，全AI重写=100%检测率
- 知网3.0: 98.6%准确率
- 维普: 近期检测更严格
- PaperYY: 社区报告故意抬高AI率
- 腾讯朱雀: 免费，适合快速验证

---

### 阶段七：导师批注处理 (Advisor Revision)

**步骤**:
1. 提取导师批注（Word COM 或 python-docx）
2. 分类处理：
   - 文字修改类 → 安全文本替换
   - 结构调整类 → 章节重组
   - 内容补充类 → 标记"待人工补充"
   - 格式修正类 → 批量格式修复
3. 生成修改对照表（原文 → 修改后 → 修改原因）
4. 逐条标记处理状态

**关键约束**: 不破坏 Zotero 引用字段和公式对象

**详见**: [references/advisor-revision-workflow.md](references/advisor-revision-workflow.md)

---

### 阶段八：答辩准备 (Defense Preparation)

**PPT 生成**:
- 使用学校模板
- 16:9 比例，20-25 页
- 结构: 封面→目录→研究区→背景→目标→技术路线→数据→方法→结果→结论→致谢

**讲稿撰写**:
- 逐页讲稿（8-10分钟）
- 3分钟压缩版（备用）
- 关键口径: 不要讲成"精确预测"，要讲成"建立流程"

**Q&A 准备**:
- 25+ 个常见问题
- 每个问题2-3句回答
- 分类: 方法类/数据类/结果类/局限性类

**详见**: [references/defense-prep-guide.md](references/defense-prep-guide.md)

---

### 阶段九：终稿提交 (Final Submission)

**提交清单**:
- [ ] 中英文摘要完整
- [ ] 目录已自动更新（Word COM）
- [ ] 正文格式一致性（调用 check_docx_format.py）
- [ ] 参考文献格式统一
- [ ] 致谢已撰写
- [ ] 图表编号连续无遗漏
- [ ] 页眉页脚正确
- [ ] 页码连续
- [ ] AIGC 检测率通过（调用 aigc-reduce）
- [ ] 导师已确认终稿

---

## 动作路由 (Action Routing)

根据 `$0` 参数路由到对应阶段：

| 动作 | 阶段 | 描述 |
|------|------|------|
| `init` | 1 | 初始化项目，提取硬约束 |
| `write` | 3 | 写作指定章节 |
| `format` | 5 | 检查/修正格式 |
| `revise` | 7 | 处理导师批注 |
| `reduce-aigc` | 6 | AIGC 降重 |
| `defense` | 8 | 答辩准备 |
| `check` | 9 | 终稿检查 |

未指定动作时，根据用户描述自动判断。

---

## 参考文档 (References)

- 写作风格指南: [references/writing-style-guide.md](references/writing-style-guide.md)
- 格式检查清单: [references/format-checklist.md](references/format-checklist.md)
- DOCX 安全编辑: [references/docx-editing-patterns.md](references/docx-editing-patterns.md)
- 导师批注处理: [references/advisor-revision-workflow.md](references/advisor-revision-workflow.md)
- AIGC 降重参考: [references/aigc-reduction-brief.md](references/aigc-reduction-brief.md)
- 答辩准备指南: [references/defense-prep-guide.md](references/defense-prep-guide.md)
- 案例研究: [references/case-study.md](references/case-study.md)

---

## 脚本 (Scripts)

```bash
# 提取 DOCX 文本（带段落编号）
python ~/.claude/skills/ug-thesis-writing/scripts/extract_docx_text.py <docx_path>

# 检查格式
python ~/.claude/skills/ug-thesis-writing/scripts/check_docx_format.py <docx_path> [--config format.json]

# 批量修正格式
python ~/.claude/skills/ug-thesis-writing/scripts/apply_format_fixes.py <docx_path> [--config format.json]
```

---

## 与其他技能的集成 (Integration)

| 关系 | 技能 | 用途 |
|------|------|------|
| 上游 | ar-literature-search | 文献检索 |
| 上游 | ar-literature-review | 文献综述 |
| 上游 | ar-data-analysis | 实验数据分析 |
| 并行 | aigc-reduce | AIGC 降重 |
| 并行 | humanizer-zh | 去 AI 味 |
| 下游 | ar-slide-generation | 答辩 PPT |
| 辅助 | ar-figure-generation | 图表生成 |
| 辅助 | ar-table-generation | 表格生成 |

---

## 案例参考 (Case Study)

本 skill 基于真实本科毕设项目经验提炼，涵盖完整的从零到终稿提交过程。
- 工具链: python-docx + Word COM + Zotero + python-pptx
- 详细记录: [references/case-study.md](references/case-study.md)
