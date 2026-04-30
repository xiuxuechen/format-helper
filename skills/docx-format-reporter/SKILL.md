---
name: docx-format-reporter
description: Internal DOCX format reporting skill. Use only when format-helper needs to generate human-readable Chinese reports, repair logs, review reports, manual confirmation lists, diff summaries, and final acceptance reports from DOCX governance artifacts.
---

# DOCX Format Reporter

## 定位

内部能力。基于审计、修复、复核和验收追溯文件生成中文人类可读报告，不修改 `.docx`。

## 输入

- `PLAN.yaml`
- `repair_plan.yaml`
- `audit_results/*.audit.json`
- `review_results/*.review.json`
- `logs/repair_log.yaml`
- `rules/selected_rule/RULE_SUMMARY.md`

## 输出

- `reports/AUDIT_REPORT.md`
- `reports/REVIEW_REPORT.md`
- `reports/MANUAL_CONFIRMATION.md`
- `reports/DIFF_SUMMARY.md`
- `reports/REPAIR_LOG.md`
- `reports/FINAL_ACCEPTANCE_REPORT.md`

## 报告要求

- 报告正文使用中文。
- 人类可读报告必须站在用户视角描述“输入文档哪里不符合要求”和“输出文档实际改变了什么”。
- 不得在给用户看的 Markdown 正文中暴露底层样式名、动作名、任务 ID、JSON 文件名、字段名或内部状态值。
- 必须把底层样式和动作翻译为中文业务表达，例如“普通正文格式”“规则正文格式”“统一正文段落格式”“自动目录仍未生成”。
- 只有最终 Word 文件名、报告文件名、用户原始文件名、必要路径和通用技术术语可以保留英文或原始名称。
- 所有报告必须引用同一个 `run_id`。
- 最终验收报告必须记录规则推荐、用户选择、规则版本、确认时间和说明。
- 默认最终回复只展示人类可读交付物和修复后 Word。
- 机器可读追溯文件保留内部标识，但不得直接整段塞进人类报告。

## 工作流

1. 确认输入文本编码。
2. 汇总审计结果，生成审计摘要、自动修复建议和人工确认项。
3. 汇总修复计划与实际修复日志，生成修复日志。
4. 汇总复核 JSON，生成复核报告。
5. 对比修复前后快照，生成差异摘要。
6. 根据验收 Gate 生成最终验收报告。

## 输出口径

- `AUDIT_REPORT.md` 必须列出问题分组、影响数量、代表性位置、原文摘录、现状、规则要求和建议处理。
- `DIFF_SUMMARY.md` 必须列出输出文档相对输入文档的真实变化，不允许保留“待补充”占位。
- `REPAIR_LOG.md` 必须说明成功处理、跳过和失败的数量及原因，不能只贴机器日志。
- `REVIEW_REPORT.md` 必须说明哪些问题已确认解决，哪些仍未解决。
- `MANUAL_CONFIRMATION.md` 必须列出仍需用户确认或专项处理的项目；若没有，明确写“无”。
- `FINAL_ACCEPTANCE_REPORT.md` 必须把通过/未通过原因讲清楚，并指向用户下一步动作。
- 六类报告必须使用 `references/report_templates/` 下的固定模板文件生成，不得在脚本里临时拼接另一套结构。

## 按需读取

- `references/report-templates.md`：报告内容模板。
