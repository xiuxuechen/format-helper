# repair_plan.yaml 最小契约

每个自动动作必须包含：

- `action_id`
- `source_issue_ids`
- `action_type`
- `target.element_id`
- `confidence`
- `semantic_evidence`
- `auto_fix_policy`
- `risk_level`
- `status`

硬规则：

- `confidence < 0.85` 不得自动修复。
- `risk_level: high` 不得自动修复。
- 缺少 `semantic_evidence` 或 `source_issue_ids` 的动作不得执行。
- 未在白名单内的 `action_type` 必须进入人工确认。
