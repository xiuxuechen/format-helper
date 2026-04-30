---
name: docx-format-auditor
description: 内部 DOCX 格式审计技能。仅当 format-helper 需要基于 document_snapshot.json、semantic_role_map.before.json 和已确认规则包执行真实格式审计，或基于 before/after snapshot 做二轮结构化复核时使用；只输出审计或复核 JSON，不修改 Word 文件。
---

# DOCX Format Auditor

## 定位

内部能力。消费事实快照、语义角色映射和规则包，执行真实格式数值比对；也用于修复后的二轮复核。

## 输入

- `PLAN.yaml`
- `document_snapshot.before.json` 或 `document_snapshot.after.json`
- `semantic_role_map.before.json`
- 已确认规则包
- 任务范围，例如标题、正文、目录、表格、页面或样式治理

## 输出

- `audit_results/{task_id}.audit.json`
- `review_results/{task_id}.review.json`
- 可复用脚本：`scripts/build_second_round_review.py`

## 强制边界

- 不修改 `.docx`。
- 不以 Word 内部样式名作为唯一通过依据。
- 不输出最终验收结论；最终汇总由 `format-helper` 和 `docx-format-reporter` 完成。
- 每个问题必须包含 `element_id`、当前事实、期望规则、建议动作、置信度和风险等级。

## 工作流

1. 读取快照、语义角色映射和规则包。
2. 按任务范围筛选元素。
3. 比对真实格式、大纲级别、目录字段、表格和页面设置。
4. 输出结构化问题或复核项。
5. 对低置信度、高风险或复杂结构标记人工确认。

## 首阶段落地脚本

```powershell
python .codex/skills/docx-format-auditor/scripts/build_second_round_review.py --run-dir format_runs/code005-fixture
```

脚本生成 T01-T06 二轮复核 JSON：OOXML 完整性、快照完整性、执行日志、自动动作追溯、人工确认留痕和渲染证据。
