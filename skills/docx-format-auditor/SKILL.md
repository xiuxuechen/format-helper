---
name: docx-format-auditor
description: Internal DOCX format audit skill. Use only when format-helper needs to snapshot a Word document, audit headings, body text, TOC, tables, pages, and style governance against a confirmed rule profile, or perform second-round structured review.
---

# DOCX Format Auditor

## 定位

内部能力。只读审计 `.docx` 和结构快照，输出专项审计或复核 JSON。不得修改 Word 文件，不得生成最终报告。

## 阶段

- `audit`：第一轮专项审计，只发现问题并给出建议动作。
- `review`：第二轮复核，对修复前后快照进行结构化比对。

## 输入

- `PLAN.yaml`
- 修复前或修复后结构快照
- 已确认规则包
- 当前任务范围，例如 `T02` 目录、`T03` 标题、`T04` 正文

## 输出

- `audit_results/{task_id}.audit.json`
- `review_results/{task_id}.review.json`

## 工作流

1. 确认输入文本编码，读取 `PLAN.yaml`、规则 YAML 和快照 JSON。
2. 若缺少快照，使用 `scripts/snapshot_docx.py` 创建只读快照。
3. 根据任务 `element_types` 过滤元素。
4. 使用 `scripts/audit_docx.py` 输出结构化问题。
5. 每个问题必须绑定 `element_id`，声明当前格式、期望格式、建议动作、置信度和风险标记。
6. 不输出“已修改”；只输出问题和建议。

## 人工确认默认项

- 静态目录中存在非正文标题项。
- 原目录项与正文标题不一致。
- 疑似标题但编号不连续。
- `1.` 或 `1.1` 编号含义不确定。
- 合并单元格、跨页表格、横向页表格。
- 页眉页脚、脚注、尾注。
- 段内加粗含义不确定。

## 按需读取

- `references/audit-schema.md`：审计 JSON。
- `references/element-taxonomy.md`：元素分类。
- `references/issue-severity.md`：严重程度。
