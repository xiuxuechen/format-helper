---
name: docx-semantic-strategist
description: 使用时机：内部 DOCX 语义策略技能。仅当 format-helper 需要基于 document_snapshot.json 和已确认规则生成 semantic_rule_draft.json、semantic_role_map.before.json 或 semantic_audit.json 时使用；用于语义角色判断、规则归纳、证据、置信度、风险等级和人工确认建议，不得修改 Word 文件或生成可执行代码。
---

# DOCX Semantic Strategist

## 定位

内部能力。根据 `.docx` 事实快照和规则上下文生成结构化语义 JSON，供规则打包、格式审计和修复计划使用。

## 强制边界

- 只输出 JSON/YAML 语义产物，不直接修改 `.docx`。
- 不生成、拼接或执行 Python、PowerShell、Shell 等可执行代码。
- 不把 Word 内部样式名作为唯一语义依据；必须结合文本模式、位置、上下文、真实格式和结构证据。
- `confidence < 0.85`、`risk_level: high`、复杂表格、页眉页脚、脚注尾注必须进入人工确认。
- `requires_user_confirmation` 只是语义层建议，最终 Gate 以 schema、风险策略和白名单校验层为准。

## 输入

- `mode`：`rule-draft`、`role-map` 或 `audit`。
- `snapshot`：`document_snapshot.json` 或标准文档快照。
- `rule_profile`：审计和修复场景必需，指向已确认规则包。
- `output_path`：本次语义产物写入路径，位于 `format_runs/{run_id}/semantic/`。

## 输出

- `mode=rule-draft`：`semantic/semantic_rule_draft.json`。
- `mode=rule-draft` 必选附加产物：`semantic/role_format_slot_facts.json`，用于记录每个 role × slot 的抽取状态、证据、冲突和 Gate blocker。
- `mode=role-map`：`semantic/semantic_role_map.before.json`。
- `mode=audit`：`semantic/semantic_audit.json`。

详细字段、阈值和禁止项按需读取 `references/semantic-output-contract.md`。

## 工作流

1. 确认输入快照和规则文件编码；`.docx` 不在本技能内读取或写回。
2. 根据 `mode` 选择目标产物和 schema。
3. 从快照中提取可解释证据：文本模式、段落位置、相邻结构、真实格式、编号、目录字段和表格结构。
4. `mode=rule-draft` 时读取 `docs/v4/schemas/role_slot_contract.yaml`，生成 `semantic/role_format_slot_facts.json`，并复算 contract/snapshot hash。
5. 为每个语义判断输出 `evidence`、`confidence`、`risk_level` 和人工确认建议。
6. 对低置信度、高风险或证据不足项，显式设置人工确认原因。
7. 输出结构化 JSON；不要在报告正文中直接堆叠内部 JSON。

## 模式要求

### rule-draft

从标准文档快照归纳规则草案。每个角色必须包含 `role`、`description`、`format`、`evidence`、`confidence`、`write_strategy` 和 `requires_user_confirmation`。

### role-map

为待处理文档元素映射语义角色。每个元素必须引用快照中存在的 `element_id`，并给出角色证据和风险等级。

### audit

基于语义角色、规则包和快照生成语义审计问题。建议动作必须能映射到后续白名单动作或人工确认，不得直接声明已修复。

## 固定执行步骤（参考 40-§6.11）

1. 读取事实快照
2. 读取规则上下文（如适用）
3. 生成语义产物（rule-draft/role-map/audit）
4. 执行语义产物 schema 校验
5. 若 `mode=rule-draft`，执行样式元素事实校验：`role_slot_contract` 覆盖、`role_format_slot_facts` schema、`FH-SLOT-FACTS-UNRESOLVED`/`FH-SLOT-FACTS-CONFLICT` blocker
6. 写入双通道输出：
   - 业务产物：`semantic/*.json`
   - 状态信封：`logs/skill_results/{seq}_docx-semantic-strategist.result.json`

## 双通道输出协议（参考 40-§6.4）

每次执行必须同时输出：

1. **业务产物**（机器权威）：
   - `semantic/semantic_rule_draft.json`（mode=rule-draft）
   - `semantic/role_format_slot_facts.json`（mode=rule-draft 必选；schema_id=`role-format-slot-facts`）
   - `semantic/semantic_role_map.before.json`（mode=role-map）
   - `semantic/semantic_audit.json`（mode=audit）

2. **状态信封**（机器权威）：
   - 路径：`logs/skill_results/{seq}_docx-semantic-strategist.result.json`
   - Schema：`skill-result`（参考 41-§5）
   - 必须包含：`result_id`、`status`、`schema_valid`、`gate_passed`、`artifacts`、`next_action`

3. **人工确认候选**（只输出 proposals，不直接写权威清单）：
   - 在业务产物中输出 `manual_review_proposals[]`
   - 不得直接写 `plans/manual_review_items.json`（由 format-helper 主控唯一写入）

## 成功输出模板（参考 40-§6.11）

```text
任务清单
1. ✅ 已读取事实快照：{snapshot_path}
2. ✅ 已读取规则上下文：{rule_context}
3. ✅ 已生成语义产物：{semantic_output_type}
4. ✅ 已执行语义产物 schema 校验

当前阶段
semantic_strategy

执行结果
✅ 成功
- 已识别语义角色：{role_count} 个
- 规则项：{rule_count} 个
- 审计项：{audit_count} 个
- 低置信度项：{low_confidence_count}
- 高风险项：{high_risk_count}

交付物
- 语义产物：{semantic_output_path}
- 语义产物 hash：{semantic_output_sha256}
- 人工确认候选：{manual_review_proposals_count} 项
  （来源见 `{semantic_output_path}.manual_review_proposals[]`）
- 状态信封：logs/skill_results/{seq}_docx-semantic-strategist.result.json

阻塞/人工确认
{无；或列出低置信度、证据不足、复杂结构项}

下一步
- 若为建规流程，进入 docx-rule-packager
- 若为审计流程，进入 docx-format-auditor 或 docx-repair-planner

验收自检
- [x] 输出 schema 通过
- [x] 每个关键判断包含 evidence
- [x] 每个关键判断包含 confidence
- [x] 未修改 Word 文件
- [x] 只输出 manual_review_proposals[]，未直接写 manual_review_items.json
```

## 失败输出模板（参考 40-§6.11）

```text
任务清单
1. ✅ 已检查事实快照：{snapshot_path}
2. ❌ 生成语义产物失败

当前阶段
semantic_strategy

执行结果
❌ 失败 / ⏸️ 阻塞

交付物
- 部分语义产物（如有）：{partial_output_path}
- 状态信封：logs/skill_results/{seq}_docx-semantic-strategist.result.json

阻塞/人工确认
阻塞原因：
- 错误码：{error_code}（如 SS-SNAPSHOT-INVALID、SS-RULE-MISSING）
- 错误消息：{error_message}
- 恢复建议：{recovery_suggestion}

下一步
- 如果快照无效：请检查 document_snapshot.json 是否完整
- 如果规则缺失：请先完成规则选择或建规流程
- 如果证据不足：请检查快照是否包含必要的段落、样式和结构信息

验收自检
- [ ] 输出 schema 通过
- [x] 已记录失败原因和错误码
- [x] 已生成状态信封（status=blocked 或 synthetic_failure）
- [x] 未修改 Word 文件
```

## 按需读取

- `references/semantic-output-contract.md`：三种语义产物的字段契约、阈值规则和安全边界。
- `../../../schemas/semantic_rule_draft.schema.json`：`semantic_rule_draft.json` 机器校验契约。
- `docs/v4/40_DESIGN_FINAL.md`：v4 设计文档。
- `docs/v4/41_SCHEMA_CONTRACTS.md`：v4 schema 契约。
