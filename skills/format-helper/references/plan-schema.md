# PLAN.yaml 契约

`PLAN.yaml` 必须在用户确认规则版本后生成。

必备顶层字段：

- `schema_version`
- `plan_id`
- `created_at`
- `document`
- `rule_profile`
- `global_objectives`
- `global_constraints`
- `artifacts`
- `tasks`
- `acceptance`

任务必须包含：

- `task_id`
- `owner`
- `scope`
- `element_types`
- `audit_output_path`
- `review_output_path`
- `status`

默认任务：

| task_id | owner | scope |
| --- | --- | --- |
| T01 | title-agent | 总题目、封面、副标题 |
| T02 | toc-agent | 静态目录、自动目录、目录层级 |
| T03 | heading-agent | 标题层级、编号连续性、大纲级别 |
| T04 | body-agent | 正文、缩进、行距、字体字号 |
| T05 | table-agent | 普通表格、专栏、附表、横向表格 |
| T06 | page-agent | 页面节、横向页、页眉页脚、脚注 |

约束：

- 子线程只能写自己的 `audit_output_path` 和 `review_output_path`。
- 主控线程负责更新 `status`、生成 `repair_plan.yaml` 和写回 `.docx`。
- 第二轮复核必须由原任务范围对应专项线程执行。
- 第二轮复核必须覆盖任务范围内全部元素，不以抽样代替结构化复核。
- T03 必须覆盖 `ambiguous-numbered-item`，避免阿拉伯数字编号候选被遗漏。
- 验收条件使用真实格式和大纲级别，不使用内部样式名作为通过依据。
- 必需标题验收条件为 `required_heading_outline_levels_set`，正文和表格验收条件为真实格式字段达标。
