---
name: docx-rule-packager
description: 使用时机：内部 DOCX 规则打包技能。仅当 format-helper 已获得 semantic_rule_draft.json、role_format_slot_facts.json、role_slot_contract.yaml，并需要生成 format-rules/{rule_id}/ 规则包、style-map.yaml、risk-policy.yaml 或 RULE_SUMMARY.md 时使用；只消费语义规则草案和槽位事实，不自行推断新规则。
---

# DOCX Rule Packager

## 定位

内部能力。把已生成并通过校验的 `semantic_rule_draft.json`、`role_format_slot_facts.json` 与 `role_slot_contract.yaml` 落成可确认的规则包和中文规则摘要。`RULE_SUMMARY.md` 只是人类阅读视图，不作为机器权威；规则值权威来自槽位事实、`slot_facts_ref` 与结构化规则包。

## 输入

- `semantic_rule_draft.json`
- `role_format_slot_facts.json`（必选，来自 `semantic/role_format_slot_facts.json`）
- `role_slot_contract.yaml`（必选，来自 `docs/v4/schemas/role_slot_contract.yaml`）
- 规则输出目录，必须为 `format-rules/{rule_id}/`（参考 40-§5.2，禁止 `format_runs/*/rules` 或 `format_rules/`）
- 用户确认过的规则名称、说明和适用范围

## 输出

所有输出必须位于 `format-rules/{rule_id}/`（参考 41-§11.2, 41-§11.3）：

- `profile.yaml`
- `style-map.yaml`
- `element-rules.yaml`
- `toc-rules.yaml`
- `table-rules.yaml`
- `page-rules.yaml`
- `risk-policy.yaml`
- `RULE_SUMMARY.md`
- `rule-package-manifest.json`（规则包清单）
- `semantic_rule_draft.json`（归档副本）
- `role_format_slot_facts.json`（归档副本或 `slot_facts_ref` 指向的可复算引用）

## 强制边界

- 不自行创造规则；所有规则必须来自 `semantic_rule_draft.json`、`role_format_slot_facts.json` 或用户确认。
- `element-rules.yaml`、`style-map.yaml`、`page-rules.yaml`、`table-rules.yaml`、`toc-rules.yaml` 的 `expected_format.{slot}` 必须来自 `role_format_slot_facts.json` 的 `slot_summary.mode_value` 或已追溯的用户确认值。
- `RULE_SUMMARY.md` 必须从 `role_format_slot_facts.json` 机械渲染，仅作人类阅读，不得被主控、Gate、审计、规划或修复链路消费。
- 不把内部样式 ID 作为规则摘要的唯一说明。
- 低置信度、`audit-only` 和人工确认项必须保留到规则包与 `RULE_SUMMARY.md`。
- 新规则默认 `status: draft`，不得自动升级为 `active`。

## 工作流

1. 读取并校验 `semantic_rule_draft.json`。
2. 读取并校验 `role_format_slot_facts.json`，确认 `slot_facts_ref`、hash、role/fact 绑定和 Gate 状态可追溯。
3. 读取并校验 `role_slot_contract.yaml`，确认 required_slots / optional_slots 与槽位事实一致。
4. 将语义角色和槽位事实映射为规则包文件；所有 `expected_format` 只允许使用 `slot_summary.mode_value` 或用户确认值。
5. 从槽位事实生成用户可读 `RULE_SUMMARY.md`，固定输出来源范围、已确定规则、未确定属性、冲突属性、人工确认清单和规则证据 6 节。
6. 输出人工确认项和 metrics，等待 `format-helper` 展示和记录用户决定。

## 脚本

- `scripts/render_rule_summary.py`：优先从 `role_format_slot_facts.json` + `role_slot_contract.yaml` 生成用户可读 `RULE_SUMMARY.md`，并返回 `resolved_slot_count`、`unresolved_slot_count`、`conflict_slot_count`、`user_confirmed_slot_count`、`resolved_rule_row_count`、`blocking_conflict_count`、`warning_conflict_count`、`gate_blocker_count`；保留 `semantic_rule_draft.json` 旧入口用于历史测试兼容。

## 固定执行步骤（参考 40-§6.12）

1. 读取并校验 semantic_rule_draft.json
2. 读取并校验 role_format_slot_facts.json 与 role_slot_contract.yaml
3. 确认规则 ID 和目标目录
4. 校验目标目录（必须是 format-rules/{rule_id}/，禁止 format_runs/*/rules）
5. 生成规则包文件和用户可读摘要；RULE_SUMMARY.md 必须从 slot facts 渲染，固定 6 节
6. 写入双通道输出：
   - 业务产物：`format-rules/{rule_id}/*.yaml` + `RULE_SUMMARY.md`
   - 状态信封：`logs/skill_results/{seq}_docx-rule-packager.result.json`

## 双通道输出协议（参考 40-§6.4）

每次执行必须同时输出：

1. **业务产物**（机器权威）：
   - `format-rules/{rule_id}/rule-package-manifest.json`
   - `format-rules/{rule_id}/profile.yaml`
   - `format-rules/{rule_id}/style-map.yaml`
   - `format-rules/{rule_id}/risk-policy.yaml`
   - `format-rules/{rule_id}/element-rules.yaml`
   - `format-rules/{rule_id}/toc-rules.yaml`
   - `format-rules/{rule_id}/table-rules.yaml`
   - `format-rules/{rule_id}/page-rules.yaml`
   - `format-rules/{rule_id}/semantic_rule_draft.json`（归档副本）
   - `format-rules/{rule_id}/role_format_slot_facts.json`（归档副本或 `slot_facts_ref`）
   - `format-rules/{rule_id}/RULE_SUMMARY.md`（中文摘要）
   - `logs/rule_ref.json`（规则引用）

2. **状态信封**（机器权威）：
   - 路径：`logs/skill_results/{seq}_docx-rule-packager.result.json`
   - Schema：`skill-result`（参考 41-§5）
   - 必须包含：`result_id`、`status`、`schema_valid`、`gate_passed`、`artifacts`、`next_action`
   - `metrics` 必须包含：`resolved_slot_count`、`unresolved_slot_count`、`conflict_slot_count`、`user_confirmed_slot_count`、`resolved_rule_row_count`、`blocking_conflict_count`、`warning_conflict_count`、`gate_blocker_count`

## 成功输出模板（参考 40-§6.12）

```text
任务清单
1. ✅ 已读取语义规则草案：{semantic_rule_draft}
2. ✅ 已读取槽位事实：{role_format_slot_facts}
3. ✅ 已读取槽位契约：{role_slot_contract}
4. ✅ 已确认规则 ID：{rule_id}
5. ✅ 已校验目标目录：format-rules/{rule_id}/
6. ✅ 已生成规则包和用户可读摘要

当前阶段
rule_packaging

执行结果
✅ 成功
- 规则包状态：{draft|pending_confirmation}
- 规则版本：{rule_version}
- 待确认规则：{manual_review_count}
- 已确定槽位：{resolved_slot_count}
- 未确定槽位：{unresolved_slot_count}
- 冲突槽位：{conflict_slot_count}
- 用户确认槽位：{user_confirmed_slot_count}
- 风险策略：{risk_policy_summary}

交付物
- 规则摘要：format-rules/{rule_id}/RULE_SUMMARY.md
- 规则包清单：format-rules/{rule_id}/rule-package-manifest.json
- 规则画像：format-rules/{rule_id}/profile.yaml
- 样式映射：format-rules/{rule_id}/style-map.yaml
- 风险策略：format-rules/{rule_id}/risk-policy.yaml
- 元素规则：format-rules/{rule_id}/element-rules.yaml
- 目录规则：format-rules/{rule_id}/toc-rules.yaml
- 表格规则：format-rules/{rule_id}/table-rules.yaml
- 页面规则：format-rules/{rule_id}/page-rules.yaml
- 语义草案归档：format-rules/{rule_id}/semantic_rule_draft.json
- 槽位事实引用：{slot_facts_ref}
- 规则引用：logs/rule_ref.json
- 状态信封：logs/skill_results/{seq}_docx-rule-packager.result.json

阻塞/人工确认
{无；或列出低置信度规则、覆盖已有规则风险}

下一步
- 选择该规则包执行审计，或先确认待确认规则

验收自检
- [x] 未写入 format_runs/*/rules
- [x] 规则路径符合 format-rules/{rule_id}/
- [x] rule-package-manifest 记录文件 hash、规则状态和激活决策
- [x] RULE_SUMMARY 从 role_format_slot_facts.json 渲染，固定 6 节，且不作为机器权威
- [x] metrics 含 resolved_slot_count、unresolved_slot_count、conflict_slot_count、user_confirmed_slot_count
- [x] 新规则默认 status: draft（不自动升级为 active）
```

## 失败输出模板（参考 40-§6.12）

```text
任务清单
1. ✅ 已检查语义规则草案：{semantic_rule_draft}
2. ❌ 规则包生成失败

当前阶段
rule_packaging

执行结果
❌ 失败 / ⏸️ 阻塞

交付物
- 部分规则包（如有）：{partial_output_path}
- 状态信封：logs/skill_results/{seq}_docx-rule-packager.result.json

阻塞/人工确认
阻塞原因：
- 错误码：{error_code}（如 RP-PATH-INVALID、RP-RULE-EXISTS、RP-DRAFT-INVALID）
- 错误消息：{error_message}
- 恢复建议：{recovery_suggestion}

下一步
- 如果目标路径位于 format_runs/：请修正为 format-rules/{rule_id}/
- 如果规则 ID 已存在且未确认覆盖：请确认是否覆盖或使用新 rule_id
- 如果草案缺少 evidence/confidence：请检查语义规则草案是否完整
- 如果风险策略缺失：请补充风险策略或使用默认策略

验收自检
- [x] 未写入 format_runs/*/rules
- [ ] 规则路径符合 format-rules/{rule_id}/
- [x] 已记录失败原因和错误码
- [x] 已生成状态信封（status=blocked 或 synthetic_failure）
```

## 后续实现

当前脚本已支持 CODE-015 槽位事实驱动的 RULE_SUMMARY 渲染，并保留 CODE-004 旧草案入口兼容；完整规则包 YAML 写入仍由后续链路继续补齐。
