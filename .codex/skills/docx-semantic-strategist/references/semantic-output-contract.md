# DOCX 语义输出契约

## 通用规则

- `schema_version` 第一阶段固定为 `1.0.0`。
- `generated_by` 固定为 `codex`。
- 时间字段使用 ISO 8601。
- `confidence` 必须在 `0` 到 `1` 之间。
- `evidence` 至少 1 条，且不得为空字符串。
- `confidence < 0.85` 必须设置 `requires_user_confirmation: true`。
- `risk_level: high` 不得进入自动修复。
- 复杂表格、页眉页脚、脚注尾注、静态目录替换默认进入人工确认。

## `semantic_rule_draft.json`

用途：从标准文档快照归纳可确认的规则草案。

必填根字段：

- `schema_version`
- `run_id`
- `created_at`
- `rule_id`
- `source_snapshot`
- `source`
- `document_type`
- `roles`
- `manual_confirmation`
- `validation`

角色项必填字段：

- `role`
- `description`
- `evidence`
- `confidence`
- `format`
- `write_strategy`
- `requires_user_confirmation`

允许的 `write_strategy`：

- `style-definition`
- `direct-format`
- `audit-only`

低置信度角色必须补充 `manual_confirmation_reason`。

## `semantic_role_map.before.json`

用途：描述待处理文档中元素到语义角色的映射。

元素项必填字段：

- `element_id`
- `semantic_role`
- `confidence`
- `evidence`
- `risk_level`
- `requires_user_confirmation`

约束：

- `element_id` 必须来自对应 `document_snapshot.json`。
- `confidence < 0.85` 不得标记为 `risk_level: low`。
- 语义不确定时优先设置 `requires_user_confirmation: true`，不得猜测为可自动修复。

## `semantic_audit.json`

用途：描述语义层发现的问题和建议动作。

问题项必填字段：

- `issue_id`
- `element_id`
- `semantic_role`
- `current_problem`
- `expected_role`
- `confidence`
- `evidence`
- `recommended_action`
- `risk_level`

`recommended_action` 必须包含：

- `action_type`
- `auto_fix_policy`

允许的 `auto_fix_policy`：

- `auto-fix`
- `manual-review`
- `audit-only`

约束：

- `risk_level: high` 或 `confidence < 0.85` 时，`auto_fix_policy` 不得为 `auto-fix`。
- `action_type` 必须可映射到后续修复计划白名单或人工确认项。
- 语义审计只描述问题和建议，不输出“已修改”或最终验收结论。
