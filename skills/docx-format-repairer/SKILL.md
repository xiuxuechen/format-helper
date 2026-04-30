---
name: docx-format-repairer
description: Internal DOCX format repair skill. Use only when format-helper has generated an approved repair_plan.yaml and needs to apply safe DOCX style, TOC, body, heading, table, or page repairs to a working copy while preserving the original document.
---

# DOCX Format Repairer

## 定位

内部能力。只能执行 `format-helper` 生成并确认的 `repair_plan.yaml`，不得自行扩大修改范围。

## 输入

- `plans/repair_plan.yaml`
- 工作副本 `.docx`
- 选中规则包

## 输出

- `output/{原文件名}{yyyyMMddHHmm}.docx`
- `snapshots/document_snapshot.after.json`
- `logs/repair_log.yaml`

## 强制约束

- 不读取或修改原始文档本体，只处理工作副本并输出新文件。
- 只执行 `auto_fix_policy: auto-fix` 的动作。
- 人工确认项不得自动修复。
- 不创建或写回 `Official*` 自定义样式。
- 如目标 Word 原生样式不存在，不得自动创建新样式；将动作标记为阻塞或人工确认。
- 标题和正文默认通过修改原生样式定义写回；直接格式覆盖必须由 `repair_plan.yaml` 逐条标记 `format_write_strategy: direct-format-override`。
- 修改标题原生样式定义前，必须确认已通过样式使用范围审计。
- 修改 `Normal` 或正文基础样式定义前，必须确认已通过继承影响范围审计。
- 复杂脚注、复杂页眉页脚、复杂表格重排第一版只审计不修复。
- 修复后必须生成新快照供第二轮复核。

## 工作流

1. 确认 `repair_plan.yaml` 编码和字段完整。
2. 复制工作副本到输出路径。
3. 按 `execution_order` 执行动作。
4. 定位 Word 原生样式；缺失样式进入阻塞或人工确认。
5. 按已确认策略修改原生样式定义或执行逐段直接格式覆盖。
6. 写入表格单元格真实格式；行高、边框仅在规则显式 `auto-fix` 且已确认时处理。
7. 在 `toc-content-audit` 已通过或已确认后，替换或插入自动目录字段。
8. 保存修复副本并生成修复日志。
9. 调用快照脚本生成 `document_snapshot.after.json`。

## 按需读取

- `references/repair-plan-schema.md`：修复计划契约。
- `references/safe-repair-policy.md`：安全修复边界。
