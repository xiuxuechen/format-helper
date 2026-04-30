---
name: docx-rule-packager
description: 内部 DOCX 规则打包技能。仅当 format-helper 已获得 semantic_rule_draft.json 并需要生成 format_rules/{rule_id}/ 规则包、style-map.yaml、risk-policy.yaml 或 RULE_SUMMARY.md 时使用；只消费语义规则草案，不自行推断新规则。
---

# DOCX Rule Packager

## 定位

内部能力。把已生成并通过校验的 `semantic_rule_draft.json` 落成可确认的规则包和中文规则摘要。

## 输入

- `semantic_rule_draft.json`
- 规则输出目录，例如 `format_rules/{rule_id}/`
- 用户确认过的规则名称、说明和适用范围

## 输出

- `profile.yaml`
- `style-map.yaml`
- `element-rules.yaml`
- `toc-rules.yaml`
- `table-rules.yaml`
- `page-rules.yaml`
- `risk-policy.yaml`
- `RULE_SUMMARY.md`

## 强制边界

- 不自行创造规则；所有规则必须来自 `semantic_rule_draft.json` 或用户确认。
- 不把内部样式 ID 作为规则摘要的唯一说明。
- 低置信度、`audit-only` 和人工确认项必须保留到规则包与 `RULE_SUMMARY.md`。
- 新规则默认 `status: draft`，不得自动升级为 `active`。

## 工作流

1. 读取并校验 `semantic_rule_draft.json`。
2. 将语义角色映射为规则包文件。
3. 生成用户可读 `RULE_SUMMARY.md`，说明字体、字号、加粗、缩进、行距、目录、表格和风险边界。
4. 输出人工确认项，等待 `format-helper` 展示和记录用户决定。

## 脚本

- `scripts/render_rule_summary.py`：从 `semantic_rule_draft.json` 生成用户可读 `RULE_SUMMARY.md`。

## 后续实现

首版脚本只覆盖 CODE-004 最小建规闭环；完整规则包 YAML 渲染将在后续链路任务中补齐。
