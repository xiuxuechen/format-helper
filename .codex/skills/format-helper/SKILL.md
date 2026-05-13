---
name: format-helper
description: 使用时机：当用户请求 Word/DOCX 格式治理、标准文档建规、规则选择、待处理文档审计、自动修复、二轮复核、报告生成、run_id 恢复，或明确提到 format-helper、$format-helper 时使用。本技能是 format-helper 唯一外部入口，负责按 Gate 编排内部 docx-* 技能；普通用户不应直接调用内部 docx-* 技能。
---

# Format Helper

## 定位

唯一外部入口。接收用户的 `.docx` 格式治理请求，创建 `format_runs/{run_id}/`，编排内部 `docx-*` 能力完成建规、审计、修复、复核、报告和恢复。

## 强制边界

- 不直接覆盖原始 `.docx`；只复制输入到运行目录并处理工作副本。
- 不直接写 Word；最终写回只能由 `docx-format-repairer` 执行已校验白名单动作。
- 不绕过规则确认、规则选择、参数覆盖、修复确认和验收 Gate。
- 建规阶段若 `role_format_slot_facts.json` 存在 unresolved 或 blocking conflict，必须生成 `logs/rule_confirmation_gate.json` 摘要，并进入 `waiting_user_on_unresolved_slots`；真实用户决策只能写入 `plans/manual_review_items.json`。
- 不把机器可读 JSON/YAML 原样堆给用户；最终解释必须使用中文报告或摘要。如果用户未显式要求，面向用户输出的"交付物"分块只能列出用户可直接使用或阅读的最终产物（修复后 .docx、中文报告、RULE_SUMMARY.md），不得列出 .yaml / .json / Schema / manifest 等机器内部追溯文件。内部追溯产物仅记录在 logs/skill_results/ 信封和 logs/state.yaml 中，不在用户输出中展示。
- 根目录 `skills/` 仅作为历史参考，入口位于 `.codex/skills/`。

## 内部能力

- `docx-fact-extractor`：生成事实快照。
- `docx-semantic-strategist`：生成语义规则、角色映射和语义审计。
- `docx-rule-packager`：把语义规则草案打包为规则包和 `RULE_SUMMARY.md`。
- `docx-format-auditor`：执行真实格式审计和二轮复核。
- `docx-repair-planner`：生成可校验的 `repair_plan.yaml`。
- `docx-format-repairer`：只对工作副本执行白名单修复。
- `docx-format-reporter`：生成中文报告和最终验收说明。

## 规则确认 Gate（CODE-016）

- `rule_confirmation_gate` 是 `rule_packaging` 阶段的 Gate 摘要与人类视图，标准路径为 `logs/rule_confirmation_gate.json`。
- `rule_confirmation_gate.json` 不承载决策权威，不写 `decision`；真实决策权威只有 `plans/manual_review_items.json`。
- 当 `role_format_slot_facts.json.gate_status=blocked` 时，format-helper 主控必须把 `gate_blockers` 提升为 `plans/manual_review_items.json.items[]`，分配 `review_id`，初始化 `decision.status=pending`。
- `rule_confirmation_gate.manual_review_item_refs[]` 只能引用 `plans/manual_review_items.json.items[].review_id`，用于展示和 Gate 索引。
- 用户确认后，format-helper 主控读取 `plans/manual_review_items.json` 的 `decision`，生成新的 `role_format_slot_facts` revision，把对应样式元素回填为 `status=user_confirmed`、`primary_source=user_confirmed`、`confidence=1.0`，再将 Gate 摘要状态更新为 `cleared`。
- 若仍存在 pending blocking 决策，运行状态必须保持 `stage=rule_packaging`、`status=waiting_user`、`next_action.kind=wait_user`，并标记原因 `waiting_user_on_unresolved_slots`。

## 工作流

1. 确认当前环境和文本文件编码；`.docx` 始终按 OOXML/ZIP 处理。
2. 解析用户意图：`extract-rule`、`audit-only`、`repair` 或 `resume`。
3. 初始化或读取 `format_runs/{run_id}/`。
4. 按运行模式调用内部能力并记录 Gate 状态。
5. 对低置信度、高风险、缺失产物、未解析样式元素或冲突样式元素停止推进，输出人工确认项；建规确认必须同时写 `logs/rule_confirmation_gate.json` 与 `plans/manual_review_items.json`。
6. 汇总用户可读交付物和机器可读追溯路径。

## 固定输出

每轮触发本技能时，最终回复必须按顺序包含：

1. 任务清单
2. 当前阶段
3. 执行结果
4. 交付物
5. 阻塞/人工确认
6. 下一步
7. 验收自检

不适用的分块必须写“无”或说明原因。

## 目录预创建流程

在确定 `run_id` 后立即执行（参考 `40-§5.1`）：

```text
ensure_dir("format-rules")
ensure_dir("format_runs/{run_id}/input")
ensure_dir("format_runs/{run_id}/snapshots")
ensure_dir("format_runs/{run_id}/semantic")
ensure_dir("format_runs/{run_id}/plans")
ensure_dir("format_runs/{run_id}/output")
ensure_dir("format_runs/{run_id}/output/_internal")
ensure_dir("format_runs/{run_id}/reports")
ensure_dir("format_runs/{run_id}/logs")
ensure_dir("format_runs/{run_id}/logs/skill_results")
ensure_dir("format_runs/{run_id}/review_results")
```

目录创建必须幂等，重复执行不得删除或覆盖已有产物。

## 失败模板

当流程阻塞或失败时，输出必须包含：

1. **任务清单**：列出已完成和未完成的步骤
2. **当前阶段**：明确停在哪个阶段（如 `fact_extraction`、`format_audit`、`repair_execution`）
3. **执行结果**：❌ 失败 / ⏸️ 阻塞
4. **交付物**：已生成的中间产物路径
5. **阻塞/人工确认**：
   - 阻塞原因（错误码、错误消息）
   - 人工确认项清单（如有）
   - 恢复建议
6. **下一步**：
   - 如果可恢复：提供恢复命令或操作步骤
   - 如果不可恢复：说明需要用户做什么
7. **验收自检**：列出已通过和未通过的检查项

## 最终自检清单

每次运行结束前，必须检查：

- [ ] 原始 `.docx` 未被覆盖
- [ ] 最终交付物命名符合 `{原文件名}{yyyyMMddHHmm}(_r[0-9]{2})?.docx`
- [ ] 内部临时文件（`_internal/`）未出现在最终交付清单
- [ ] `logs/final_acceptance.json` 已生成且 `status` 明确
- [ ] 报告已生成或明确说明为何未生成
- [ ] 所有人工确认项已处理或明确标记为待处理
- [ ] `logs/state.yaml` 已更新且可用于恢复

## 按需读取

- `references/workflow.md`：运行模式和 Gate 顺序。
- `references/artifact-routing.md`：运行目录和恢复产物。
- `references/plan-schema.md`：`PLAN.yaml` 最小契约。
- `references/repair-plan-schema.md`：`repair_plan.yaml` 最小契约。
- `docs/v4/40_DESIGN_FINAL.md`：v4 设计文档。
- `docs/v4/41_SCHEMA_CONTRACTS.md`：v4 schema 契约。
- `docs/v4/50_DEV_PLAN.md`：Phase 5 开发计划。
