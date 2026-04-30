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
- 只执行 `auto_fix_policy: auto-fix` 且通过白名单、置信度、风险和语义证据校验的动作。
- 不自动创建缺失的 Word 原生样式。
- 不执行高风险、低置信度或人工确认未通过的动作。
- 写回后必须允许后续生成 `document_snapshot.after.json`。

## 工作流

1. 校验 `repair_plan.yaml` 的 schema、白名单、风险和证据。
2. 复制或打开工作副本，执行允许的修复动作。
3. 记录已执行、跳过、拒绝和失败动作。
4. 保存输出副本并验证 OOXML 可打开。
5. 将结果交给 `format-helper` 触发 after snapshot 和二轮复核。

## 后续实现

旧 `skills/docx-format-repairer` 只作为参考。v3 写回入口必须重建安全校验，不保留兼容 fallback。

## 首阶段落地脚本

```powershell
python .codex/skills/docx-format-repairer/scripts/apply_repair_plan.py --repair-plan plans/repair_plan.yaml --log logs/repair_execution.json
```

脚本只执行 `auto_fix_policy: auto-fix` 且通过 schema、白名单、置信度、风险等级和证据校验的动作；当前最小实现覆盖 `map_heading_native_style` 和 `apply_body_direct_format`，其他白名单动作先安全跳过并记录日志。
