---
name: docx-format-repairer
description: 内部 DOCX 安全写回技能。仅当 format-helper 已获得通过校验的 repair_plan.yaml，并需要对工作副本执行白名单自动修复、生成输出 docx 和 after snapshot 前置产物时使用；不得覆盖原始 Word，不得执行未确认高风险动作。
---

# DOCX Format Repairer

## 定位

内部能力。只对 `input/working.docx` 执行已校验的白名单动作，生成 `output/*.docx`。

## 输入

- `plans/repair_plan.yaml`
- `input/working.docx`
- 输出路径，例如 `output/{source}{yyyyMMddHHmm}.docx`

## 输出

- 修复后 `.docx`
- 写回执行日志
- 拒绝执行项清单
- 可复用脚本：`scripts/apply_repair_plan.py`

## 强制边界

- 不覆盖原始 `.docx`。
- 只执行 `allowed_by_policy=true` 且 `policy_match_ref.source_kind=action_whitelist` 的动作（参考 41-§3.13.1、41-§11.9）。
- 不自动创建缺失的 Word 原生样式。
- 不执行高风险、低置信度或人工确认未通过的动作。
- 写回后必须允许后续生成 `document_snapshot.after.json`。

## 工作流

1. 校验 `repair_plan.yaml` 的 schema、白名单、风险和证据。
2. 复制或打开工作副本，执行允许的修复动作。
3. 记录已执行、跳过、拒绝和失败动作。
4. 保存输出副本并验证 OOXML 可打开。
5. 将结果交给 `format-helper` 触发 after snapshot 和二轮复核。

## 固定执行步骤（参考 40-§6.15）

1. 读取并校验 finalized 修复计划（必须是 `plans/repair_plan.finalized.r{plan_revision}.yaml`）
2. 对工作副本执行白名单动作
3. 刷新 Word 自动目录（如适用）
4. 生成规范命名最终交付物
5. 写入修复日志和 TOC 验收结果
6. 写入双通道输出：
   - 业务产物：`output/{原文件名}{yyyyMMddHHmm}(_r[0-9]{2})?.docx` + `logs/repair_execution_log.json` + `logs/toc_acceptance.json`
   - 状态信封：`logs/skill_results/{seq}_docx-format-repairer.result.json`

## 关键约束（参考 40-§5.5, 41-§11.9, 41-§11.10）

**最终输出命名**：
- 必须匹配正则 `output/{原文件名}{yyyyMMddHHmm}(_r[0-9]{2})?.docx`
- 无冲突时不得加 `_rNN` 后缀
- 禁止 `.formatted.docx`、`_with_toc.docx` 等内部状态后缀作为最终交付
- 内部临时文件只能放在 `output/_internal/`，不进入最终交付清单

**拒绝消费 draft plan**：
- 只能消费 `plans/repair_plan.finalized.r{plan_revision}.yaml`（finalized，revisioned canonical）
- 拒绝 `plans/repair_plan.draft.yaml`（draft 不可执行）
- 拒绝 `plans/repair_plan.yaml`、`plans/repair_plan.finalized.yaml`（历史兼容别名）
- 指向非 finalized plan 时必须 blocked

**只执行白名单动作**（参考 41-§3.13.1, 41-§11.9）：
- 只执行 `allowed_by_policy=true` 且 `policy_match_ref.source_kind=action_whitelist` 的动作
- 其他 5 个分支（allowed_operations、blocked_operations、risk_override、default_policy、no_policy_match）都不得写回
- `execution_status=requires_manual_review` 的动作必须停在 waiting_user，不得执行

**原始 Word 未覆盖证明**（参考 41-§3.8）：
- 必须生成 `original_docx_proof`，包含 initial/current sha256/size_bytes
- `initial_sha256=current_sha256` 且 `initial_size_bytes=current_size_bytes` 才能通过
- 任一不一致必须阻断最终验收

**TOC 验收**（参考 40-§5.6）：
- 必须生成 `logs/toc_acceptance.json`
- `toc_mode` 只能为 `native_toc`、`equivalent_visible_toc`、`not_required`
- Word/Office 不可用但需要 TOC 刷新时必须 blocked，不得伪造 accepted

## 双通道输出协议（参考 40-§6.4）

每次执行必须同时输出：

1. **业务产物**（机器权威）：
   - `output/{原文件名}{yyyyMMddHHmm}(_r[0-9]{2})?.docx`（最终 Word）
   - `output/_internal/`（内部临时产物，不进入最终交付）
   - `logs/repair_execution_log.json`（包含 original_docx_proof）
   - `logs/toc_acceptance.json`（TOC 验收结果，仅 repair_execution 阶段生成）

2. **状态信封**（机器权威）：
   - 路径：`logs/skill_results/{seq}_docx-format-repairer.result.json`
   - Schema：`skill-result`（参考 41-§5）
   - 必须包含：`result_id`、`status`、`schema_valid`、`gate_passed`、`artifacts`、`next_action`

## 成功输出模板（参考 40-§6.15）

```text
任务清单
1. ✅ 已读取并校验 finalized 修复计划：{repair_plan_path}
2. ✅ 已对工作副本执行白名单动作
3. ✅ 已刷新 Word 自动目录：{是|否|不适用}
4. ✅ 已生成规范命名最终交付物
5. ✅ 已写入修复日志

当前阶段
repair_execution

执行结果
✅ 成功
- executed：{executed_count}
- skipped：{skipped_count}
- rejected：{rejected_count}
- 自动目录可见：{是|否|不适用}
- 最终文件名：{原文件名}{yyyyMMddHHmm}(_r[0-9]{2})?.docx
- 原始文件证明：initial/current hash、size、checked_at 已记录

交付物
- 最终 Word：output/{原文件名}{yyyyMMddHHmm}(_r[0-9]{2})?.docx
- 修复日志：logs/repair_execution_log.json（包含 original_docx_proof）
- TOC 验收：logs/toc_acceptance.json
- 内部临时产物：output/_internal/
- 状态信封：logs/skill_results/{seq}_docx-format-repairer.result.json

阻塞/人工确认
无

下一步
- 生成 after snapshot
- 执行二轮复核

验收自检
- [x] 原始 Word 未覆盖
- [x] original_docx_proof hash 与 size 均匹配
- [x] 输出命名符合规范 {原文件名}{yyyyMMddHHmm}(_r[0-9]{2})?.docx
- [x] 自动目录已可见或明确不适用
- [x] 内部临时文件未作为最终交付（仅存放于 output/_internal/）
- [x] 未执行未确认高风险动作
- [x] 只消费 finalized plan（revisioned canonical 路径）
- [x] 只执行 policy_match_ref.source_kind=action_whitelist 的动作
```

## 失败输出模板（参考 40-§6.15）

```text
任务清单
1. ✅ 已检查修复计划：{repair_plan_path}
2. ❌ 安全写回未完成

当前阶段
repair_execution

执行结果
❌ 失败 / ⏸️ 阻塞

交付物
- 部分输出（如有）：{partial_output_path}
- 状态信封：logs/skill_results/{seq}_docx-format-repairer.result.json

阻塞/人工确认
阻塞原因：
- 错误码：{error_code}（如 FR-PLAN-INVALID、FR-DRAFT-REJECTED、FR-TOC-FAILED、FR-NAMING-INVALID、FR-OFFICE-UNAVAILABLE、FR-ORIGINAL-MODIFIED）
- 错误消息：{error_message}
- 恢复建议：{recovery_suggestion}

下一步
- 如果修复计划未通过校验：请先完成 docx-repair-planner 生成 finalized plan
- 如果消费 draft plan：请等待用户决策后重新生成 finalized plan
- 如果存在未确认高风险动作：请先完成人工确认
- 如果自动目录刷新失败：请检查 Word/Office 是否可用
- 如果最终命名不合规：请检查命名规则
- 如果 Word COM 不可用：必须 blocked，不得伪造 accepted
- 如果原始 Word 被修改：请从备份恢复原始文件

验收自检
- [x] 原始 Word 未覆盖（或明确说明）
- [ ] 生成 accepted 最终 Word
- [x] 目录占位是否残留已明确记录
- [x] 已记录失败原因和错误码
- [x] 已生成状态信封（status=blocked 或 synthetic_failure）
```

**失败模板关键规则**（参考 40-§6.15）：
- 若"原始 Word 未覆盖"为"否"或"未知"，必须把 `original_docx_proof` 缺失、hash/size 不一致或无法复算列入阻塞原因
- 最终验收不得进入 `accepted` 或 `accepted_with_warnings`

## 后续实现

旧 `skills/docx-format-repairer` 只作为参考。写回入口必须重建安全校验，不保留兼容 fallback。

## 首阶段落地脚本

```powershell
python .codex/skills/docx-format-repairer/scripts/apply_repair_plan.py --repair-plan plans/repair_plan.finalized.r{plan_revision}.yaml --log logs/repair_execution_log.json
```

脚本只执行 `allowed_by_policy=true` 且 `policy_match_ref.source_kind=action_whitelist` 的动作（参考 41-§3.13.1）；当前最小实现覆盖 `map_heading_native_style` 和 `apply_body_direct_format`，其他白名单动作先安全跳过并记录日志。

**路径约束**（参考 40-§5.5, 41-§11.9）：
- 脚本输入必须是 revisioned canonical 路径 `plans/repair_plan.finalized.r{plan_revision}.yaml`
- 不接受 `plans/repair_plan.yaml`、`plans/repair_plan.finalized.yaml` 或 `plans/repair_plan.draft.yaml`
- 脚本输出日志必须写入 `logs/repair_execution_log.json`（不得使用历史名 `repair_execution.json`）
