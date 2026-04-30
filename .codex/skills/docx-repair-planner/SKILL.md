---
name: docx-repair-planner
description: 内部 DOCX 修复计划技能。仅当 format-helper 已获得 semantic_audit.json、format audit 结果和 risk-policy.yaml，并需要生成可追溯 repair_plan.yaml 与 manual_review_items 时使用；负责风险、置信度和白名单校验，不执行 Word 写回。
---

# DOCX Repair Planner

## 定位

内部能力。把语义审计、真实格式审计和风险策略合成为可校验、可恢复、可人工确认的 `repair_plan.yaml`。

## 输入

- `semantic_audit.json`
- `audit_results/*.audit.json`
- `risk-policy.yaml`
- `PLAN.yaml`

## 输出

- `plans/repair_plan.yaml`
- `plans/manual_review_items.yaml`
- 可复用脚本：`scripts/build_repair_plan.py`

## 强制边界

- 不写 Word，不修改工作副本。
- 不放行缺少 `confidence`、`semantic_evidence` 或 `source_issue_ids` 的动作。
- `confidence < 0.85` 或 `risk_level: high` 的项目不得设为 `auto_fix_policy: auto-fix`。
- `action_type` 必须进入白名单或转为人工确认。

## 工作流

1. 读取语义审计、格式审计和风险策略。
2. 合并重复问题，保留源问题 ID。
3. 为每个候选动作写入目标元素、前后值、置信度、证据、风险等级和执行顺序。
4. 输出自动修复动作和人工确认项。
5. 把无法自动修复的项目交给 `format-helper` 的修复确认 Gate。

## 白名单方向

首阶段白名单包括标题原生样式映射、正文格式、表格单元格格式、表格边框、目录内容审计、自动目录插入或替换。具体动作集合以 schema 和风险策略为准。

## 首阶段落地脚本

```powershell
python .codex/skills/docx-repair-planner/scripts/build_repair_plan.py --semantic-audit semantic/semantic_audit.json --snapshot snapshots/document_snapshot.before.json --rule-id official-report-v1 --source-docx input/original.docx --working-docx input/working.docx --output-docx output --output plans/repair_plan.yaml
```

脚本必须先校验 `semantic_audit.json`，再按置信度、风险等级和动作白名单决定 `auto-fix` 或 `manual-review`。
