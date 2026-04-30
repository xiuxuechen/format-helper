# v3 运行产物路径

## 运行目录

```text
format_runs/{run_id}/
├── input/
├── rules/selected_rule/
├── snapshots/
├── semantic/
├── plans/
├── audit_results/
├── review_results/
├── reports/
├── logs/
└── output/
```

## 恢复必查

- `logs/state.yaml`
- `plans/PLAN.yaml`
- `snapshots/*.json`
- `semantic/*.json`
- 修复模式下必查 `plans/repair_plan.yaml` 和 `output/*.docx`

缺失关键产物时，恢复任务必须列出阻塞项，不得跳过 Gate。
