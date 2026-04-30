# 运行目录与恢复

## 输出根目录

按顺序确定：

1. 用户显式指定的输出目录。
2. 当前工作区下的 `format_runs/`。
3. 当前工作区不可写时，要求用户指定可写输出目录。

不默认写入原始文档所在目录。

## 目录结构

```text
format_runs/{run_id}/
  input/
  rules/selected_rule/
  snapshots/
  plans/
  audit_results/
  review_results/
  reports/
  logs/
  output/
```

`run_id` 使用 `{yyyyMMdd-HHmmss}-{source_name_slug}`，同一分钟重复运行时追加短随机串或递增序号。

## 人类可读交付物

- `output/{原文件名}{yyyyMMddHHmm}.docx`
- `reports/AUDIT_REPORT.md`
- `reports/REVIEW_REPORT.md`
- `reports/MANUAL_CONFIRMATION.md`
- `reports/DIFF_SUMMARY.md`
- `reports/REPAIR_LOG.md`
- `reports/FINAL_ACCEPTANCE_REPORT.md`
- `rules/selected_rule/RULE_SUMMARY.md`

## 机器可读追溯文件

- `logs/run_log.yaml`
- `logs/repair_log.yaml`
- `plans/PLAN.yaml`
- `plans/repair_plan.yaml`
- `snapshots/document_snapshot.before.json`
- `snapshots/document_snapshot.after.json`
- `audit_results/*.audit.json`
- `review_results/*.review.json`
- `rules/overrides.yaml`

## 恢复机制

支持：

```text
format-helper resume
继续上次 Word 格式处理
恢复 run_id: 20260429-103000-central-budget-application
```

用户指定 `run_id` 时，读取对应 `logs/run_log.yaml`，展示恢复摘要并询问：

```text
是否继续这个任务？
```

用户未指定 `run_id` 时，列出状态为 `in_progress`、`waiting_user`、`blocked`、`failed` 的任务，让用户选择序号或任务 ID，再展示恢复摘要并二次确认。
