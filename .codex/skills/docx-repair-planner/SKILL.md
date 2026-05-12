---
name: docx-repair-planner
description: 使用时机：内部 DOCX 修复计划技能。仅当 format-helper 已获得 semantic_audit.json、format audit 结果和 risk-policy.yaml，并需要生成可追溯 repair_plan.yaml 与 manual_review_items 时使用；负责风险、置信度和白名单校验，不执行 Word 写回。
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

## 固定执行步骤（参考 40-§6.14）

1. 读取审计结果
2. 读取风险策略
3. 生成 draft 修复计划和人工确认 proposals
4. 读取由 format-helper 主控生成的 manual review 决策快照（用户决策后）
5. 在用户决策后重算 finalized 修复计划
6. 写入双通道输出：
   - 业务产物：`plans/repair_plan.draft.yaml` 和 `plans/repair_plan.finalized.r{plan_revision}.yaml`
   - 状态信封：`logs/skill_results/{seq}_docx-repair-planner.result.json`

## draft/finalized 分离（参考 40-§5.3, 41-§11.9）

**重要**：repair-plan 分为两个阶段，draft 不可执行，finalized 才能写回：

| plan_state | 标准路径 | 可执行 | manual_review_items_ref | decision_snapshot |
|-----------|---------|--------|------------------------|-------------------|
| `draft` | `plans/repair_plan.draft.yaml` | ❌ 否 | `absent` / `draft` | 必须为 null |
| `finalized` | `plans/repair_plan.finalized.r{plan_revision}.yaml` | ✅ 是 | `finalized` | 必须非 null |

**关键约束**：
- finalized 必须使用 revisioned canonical 路径（不得使用 `plans/repair_plan.finalized.yaml` 或 `plans/repair_plan.yaml`）
- finalized 必须绑定决策快照、manual review hash、风险策略 hash 和白名单复算证据
- `plan_revision` 必须由输入 hash 集合确定性派生（禁止随机数/时间戳）

## 双通道输出协议（参考 40-§6.4）

每次执行必须同时输出：

1. **业务产物**（机器权威）：
   - `plans/repair_plan.draft.yaml`（draft 阶段）
   - `plans/repair_plan.finalized.r{plan_revision}.yaml`（finalized 阶段）
   - `manual_review_proposals[]`（在 draft 中输出，不直接写权威清单）

2. **状态信封**（机器权威）：
   - 路径：`logs/skill_results/{seq}_docx-repair-planner.result.json`
   - Schema：`skill-result`（参考 41-§5）
   - 必须包含：`result_id`、`status`、`schema_valid`、`gate_passed`、`artifacts`、`next_action`

3. **人工确认候选**（只输出 proposals，不直接写权威清单）：
   - 在 draft plan 中输出 `manual_review_proposals[]`
   - 不得直接写 `plans/manual_review_items.json`（由 format-helper 主控唯一写入）

## 成功输出模板（参考 40-§6.14）

```text
任务清单
1. ✅ 已读取审计结果：{audit_result_path}
2. ✅ 已读取风险策略：{risk_policy_path}
3. ✅ 已生成 draft 修复计划和人工确认 proposals
4. ✅ 已读取 manual review 决策快照：{decision_snapshot_path}
5. ✅ 已在用户决策后重算 finalized 修复计划：{是|否}

当前阶段
{repair_planning|manual_review}

执行结果
✅ 成功
- 可执行修复动作：{executable_action_count}
- 人工确认动作：{manual_action_count}
- 拒绝/禁止候选项：{rejected_or_blocked_candidate_count}
- 目录相关动作：{toc_action_count}
- 表格结构动作：{table_structure_action_count}

交付物
- draft 修复计划：plans/repair_plan.draft.yaml
- 人工确认机器清单引用：plans/manual_review_items.json（仅 format-helper 主控写入）
- 人工确认清单：reports/MANUAL_CONFIRMATION.md
- finalized 修复计划：plans/repair_plan.finalized.r{plan_revision}.yaml（仅用户决策后生成）
- 状态信封：logs/skill_results/{seq}_docx-repair-planner.result.json

阻塞/人工确认
{无；或列出高风险、低置信度、目录刷新、表格结构动作}

下一步
- 若仍有 pending 阻塞项，等待用户确认
- 若已生成 finalized 计划，进入 docx-format-repairer 执行安全写回

验收自检
- [x] draft repair_plan schema 通过
- [x] finalized repair_plan schema 通过或尚未进入 finalize
- [x] finalized 计划绑定 decision_snapshot、manual_review_items hash/size 和 risk-policy hash
- [x] 每个动作包含 source_issue_ids、source_semantic_finding_ids 或 source_format_issue_ids
- [x] 每个动作包含 target.attribute、before_value、after_value 或 desired_value
- [x] 每个动作包含 confidence 和 evidence_refs
- [x] 每个动作包含 risk_level
- [x] 未执行 Word 写回
- [x] 只输出 manual_review_proposals[]，未直接写 manual_review_items.json
- [x] finalized 使用 revisioned canonical 路径
```

## 失败输出模板（参考 40-§6.14）

```text
任务清单
1. ✅ 已检查审计结果：{audit_result_path}
2. ❌ 修复计划未生成或不可用

当前阶段
repair_planning

执行结果
❌ 失败 / ⏸️ 阻塞

交付物
- 部分修复计划（如有）：{partial_plan_path}
- 状态信封：logs/skill_results/{seq}_docx-repair-planner.result.json

阻塞/人工确认
阻塞原因：
- 错误码：{error_code}（如 RPL-AUDIT-MISSING、RPL-RISK-MISSING、RPL-SCHEMA-FAILED、RPL-NON-WHITELIST）
- 错误消息：{error_message}
- 恢复建议：{recovery_suggestion}

下一步
- 如果审计结果缺失：请先完成 docx-format-auditor 审计
- 如果风险策略缺失：请检查规则包是否包含 risk-policy.yaml
- 如果人工确认决策缺失：请等待用户完成决策快照
- 如果包含未白名单动作：请检查动作是否在 action_whitelist[] 中
- 如果 schema 不通过：请检查 repair_plan 格式

验收自检
- [x] 未执行 Word 写回
- [ ] 可进入安全写回
- [x] 已记录失败原因和错误码
- [x] 已生成状态信封（status=blocked 或 synthetic_failure）
```

## 首阶段落地脚本

```powershell
python .codex/skills/docx-repair-planner/scripts/build_repair_plan.py --semantic-audit semantic/semantic_audit.json --snapshot snapshots/document_snapshot.before.json --rule-id official-report-v1 --source-docx input/original.docx --working-docx input/working.docx --output-docx output --output plans/repair_plan.yaml
```

脚本必须先校验 `semantic_audit.json`，再按置信度、风险等级和动作白名单决定 `auto-fix` 或 `manual-review`。
