---
name: docx-format-reporter
description: 内部 DOCX 报告生成技能。仅当 format-helper 需要从 format_runs/{run_id}/ 的计划、审计、修复、复核、人工确认和最终验收产物生成中文可读报告时使用；不得把内部 JSON 原样作为报告正文。
---

# DOCX Format Reporter

## 定位

内部能力。把机器可读追溯产物转换为用户可读中文报告，说明问题、规则、动作、风险、未修项和验收结论。

## 输入

- `plans/PLAN.yaml`
- `plans/repair_plan.yaml`
- `audit_results/*.audit.json`
- `review_results/*.review.json`
- `logs/state.yaml`
- 规则摘要和人工确认记录

## 输出

- `reports/AUDIT_REPORT.md`
- `reports/REVIEW_REPORT.md`
- `reports/MANUAL_CONFIRMATION.md`
- `reports/DIFF_SUMMARY.md`
- `reports/REPAIR_LOG.md`
- `reports/FINAL_ACCEPTANCE_REPORT.md`
- `logs/final_acceptance.json`
- `logs/state.yaml`
- 可复用脚本：`scripts/render_final_reports.py`

## 强制边界

- 不修改 `.docx`。
- 不把内部 JSON/YAML 原样堆到报告正文。
- 未完成二轮复核时，不输出最终通过结论。
- 报告必须解释剩余风险、人工确认项和阻塞原因。

## 工作流

1. 读取运行计划、规则、审计、修复、复核和状态产物。
2. 生成审计报告和人工确认清单。
3. 修复后生成修复日志、差异摘要和复核报告。
4. 根据验收证据生成最终验收报告；若存在 blocker，明确标记未通过。
5. 将报告路径返回给 `format-helper` 汇总展示。

## 首阶段落地脚本

```powershell
python .codex/skills/docx-format-reporter/scripts/render_final_reports.py --run-dir format_runs/code005-fixture
```

脚本读取 `repair_plan.yaml`、执行日志、before/after 快照和 T01-T06 复核结果，生成全部 Markdown 报告、`final_acceptance.json` 和 `state.yaml`。
