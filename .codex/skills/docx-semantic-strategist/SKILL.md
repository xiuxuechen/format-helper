---
name: docx-semantic-strategist
description: 内部 DOCX 语义策略技能。仅当 format-helper 需要基于 document_snapshot.json 和已确认规则生成 semantic_rule_draft.json、semantic_role_map.before.json 或 semantic_audit.json 时使用；用于语义角色判断、规则归纳、证据、置信度、风险等级和人工确认建议，不得修改 Word 文件或生成可执行代码。
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
- `mode=role-map`：`semantic/semantic_role_map.before.json`。
- `mode=audit`：`semantic/semantic_audit.json`。

详细字段、阈值和禁止项按需读取 `references/semantic-output-contract.md`。

## 工作流

1. 确认输入快照和规则文件编码；`.docx` 不在本技能内读取或写回。
2. 根据 `mode` 选择目标产物和 schema。
3. 从快照中提取可解释证据：文本模式、段落位置、相邻结构、真实格式、编号、目录字段和表格结构。
4. 为每个语义判断输出 `evidence`、`confidence`、`risk_level` 和人工确认建议。
5. 对低置信度、高风险或证据不足项，显式设置人工确认原因。
6. 输出结构化 JSON；不要在报告正文中直接堆叠内部 JSON。

## 模式要求

### rule-draft

从标准文档快照归纳规则草案。每个角色必须包含 `role`、`description`、`format`、`evidence`、`confidence`、`write_strategy` 和 `requires_user_confirmation`。

### role-map

为待处理文档元素映射语义角色。每个元素必须引用快照中存在的 `element_id`，并给出角色证据和风险等级。

### audit

基于语义角色、规则包和快照生成语义审计问题。建议动作必须能映射到后续白名单动作或人工确认，不得直接声明已修复。

## 按需读取

- `references/semantic-output-contract.md`：三种语义产物的字段契约、阈值规则和安全边界。
- `../../../schemas/semantic_rule_draft.schema.json`：`semantic_rule_draft.json` 机器校验契约。
