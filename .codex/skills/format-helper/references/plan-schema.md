# PLAN.yaml 最小契约

```yaml
schema_version: 1.0.0
run_id: 20260430-153700-official-report
intent: repair
inputs:
  source_docx: input/original.docx
  working_docx: input/working.docx
rule:
  rule_id: official-report-v1
  selected_rule_path: rules/selected_rule/profile.yaml
stages:
  - id: snapshot_before
    status: done
gates:
  rule_selection:
    status: passed
```

`status` 使用 `pending`、`done`、`blocked`、`waiting_user` 或 `skipped`。
