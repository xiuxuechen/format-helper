---
name: format-helper
description: Orchestrate formal document format governance. Use when Codex needs to process Word .docx formatting with rule selection, standard style governance, automatic TOC migration, audit, repair planning, review, reporting, or resume support.
---

# Format Helper

## 定位

作为格式治理总入口，接收用户的文档格式处理请求，判断文件类型，选择或创建规则版本，生成运行目录与 `PLAN.yaml`，编排内部 `docx-*` skill 完成审计、修复计划、统一修复、复核和最终验收。

当前只实现 `.docx`。后续 `.xlsx`、`.pptx` 能力必须继续复用本入口。

## 强制边界

- 不直接覆盖原始文档；始终创建工作副本和修复副本。
- 用户默认只调用 `format-helper`，内部 `docx-*` 由本 skill 编排。
- 内部审计与复核只能写各自 JSON；最终 `.docx` 只能由主控线程按 `repair_plan.yaml` 写回。
- 未确认规则版本前，不生成正式 `PLAN.yaml`。
- 未完成第二轮结构化复核前，不输出最终通过结论。
- 读取、编辑、写回文本文件前先确认编码；新建文本文件默认 UTF-8 无 BOM。
- `.docx` 必须按 ZIP/OOXML 处理，不按文本转码。

## 按需读取

- `references/workflow.md`：完整执行流程、规则选择 Gate、覆盖 Gate、验收 Gate。
- `references/artifact-routing.md`：运行目录、交付物命名、恢复机制。
- `references/plan-schema.md`：`PLAN.yaml` 契约。
- `references/repair-plan-schema.md`：`repair_plan.yaml` 契约。

## 规则库

共享规则库位于仓库根目录：

```text
format_rules/docx/rule_profiles/{rule-id}/
```

每个规则版本必须包含：

- `profile.yaml`
- `style-map.yaml`
- `element-rules.yaml`
- `toc-rules.yaml`
- `table-rules.yaml`
- `page-rules.yaml`
- `risk-policy.yaml`
- `RULE_SUMMARY.md`

规则状态使用 `draft`、`active`、`deprecated`。新建规则默认为 `draft`；真实文档处理并最终验收通过后，再由用户确认是否升级为 `active`。

## 工作流

1. 确认当前环境与相关文本文件编码；对 `.docx` 使用 OOXML 方式读取。
2. 解析用户输入：原始文档、标准文档、输出目录、是否恢复、是否只审计。
3. 如果是恢复请求，按 `references/artifact-routing.md` 查找 `format_runs/{run_id}/logs/run_log.yaml`，展示恢复摘要并询问是否继续。
4. 如果是首次建规，调用 `docx-rule-extractor` 从标准 `.docx` 生成规则草案和 `RULE_SUMMARY.md`，展示规则名称、说明和摘要，等待用户确认。
5. 如果是处理原始 `.docx`，先生成文档画像，列出可用规则版本，给出推荐并等待用户确认。
6. 如果用户选择本次参数覆盖，使用覆盖 Gate 表格展示差异、风险、推荐处理和用户选择。
7. 创建 `format_runs/{run_id}/` 运行目录，复制输入文件到 `input/`，复制选中规则到 `rules/selected_rule/`。
8. 生成正式 `plans/PLAN.yaml`。
9. 调用 `docx-format-auditor` 按任务输出 `audit_results/*.audit.json`。
10. 合并审计结果，完成去重和冲突仲裁，生成 `plans/repair_plan.yaml`。
11. 调用 `docx-format-repairer` 按修复计划统一写回修复副本。
12. 生成修复后快照，按原 `PLAN.yaml` 调用 `docx-format-auditor` 或专项线程输出 `review_results/*.review.json`。
13. 调用 `docx-format-reporter` 生成审计报告、修复日志、复核报告、人工确认项、差异摘要和最终验收报告。
14. 最终回复只展示修复后 Word 与人类可读报告；机器可读追溯文件按需提供。

## 常用脚本

- `scripts/list_rule_profiles.py`：列出 `format_rules/docx/rule_profiles/` 下可用规则。
- `scripts/build_task_plan.py`：根据文档画像、规则版本和运行目录生成 `PLAN.yaml`。
- `scripts/merge_audit_results.py`：合并 `audit_results/*.audit.json`。
- `scripts/build_repair_plan.py`：从合并审计结果生成 `repair_plan.yaml` 初稿。

脚本只能辅助生成可追溯文件；涉及用户确认的 Gate 仍必须由主控线程展示并记录决定。

## 每轮交互固定输出

凡触发 `format-helper` 的每轮回复，无论处于建规、待处理文档确认、审计、修复、复核、验收、恢复、排查或整改报告阶段，都必须使用以下固定结构，并按顺序输出。除非某项确实不适用，否则不得省略；不适用时明确写“无”或说明原因。

1. **任务清单**
   - 列出本轮需要完成或已完成的关键事项。
   - 对跨轮任务标注状态：`已完成`、`进行中`、`待确认`、`阻塞`。
   - 建规场景必须包含规则抽取、规则摘要、规则确认状态。
   - 待处理文档确认场景必须包含待处理文档、候选规则、推荐规则、确认事项。
   - 审计修复完成场景必须包含审计、修复计划、修复执行、复核、报告生成。
   - 排查或整改报告逻辑场景必须包含问题定位、影响范围、整改动作、验证结论。

2. **当前阶段**
   - 明确当前所处阶段，例如：`规则建规`、`规则确认 Gate`、`文档处理确认 Gate`、`审计中`、`修复计划生成`、`修复执行`、`复核验收`、`恢复待确认`、`排查整改`。
   - 如阶段不能推进，说明停留原因。

3. **执行结果**
   - 概述本轮已经真实执行的动作和结论。
   - 对只做了检查、读取或计划而未改动文件的情况，也必须明确说明。
   - 不得把内部执行步骤隐去；凡工作流要求展示、说明、列出、输出或提示给用户的内容，都必须在本节或对应固定分块中可见。

4. **交付物**
   - 列出本轮新增、更新或确认可用的文件路径。
   - 若尚未生成交付物，写“无”，并说明预计在哪个阶段生成。
   - 最终交付时优先展示修复后 Word 和人类可读报告；机器可读追溯文件按需列出。

5. **阻塞/人工确认**
   - 列出需要用户确认、审批或补充的信息。
   - 规则选择 Gate、参数覆盖 Gate、恢复继续 Gate、规则升级为 `active` 等必须出现在本节。
   - 若没有阻塞或人工确认项，写“无”。

6. **下一步**
   - 给出下一步具体动作，必要时按序号列出。
   - 如果需要用户先确认，下一步必须明确等待的确认项，不能假设已获批准。
   - 如果本轮已经完成全部工作，说明后续可选动作或验收入口。

7. **验收自检**
   - 固定列出自检结果，至少覆盖：
     - 是否已覆盖所有面向用户可见的执行步骤。
     - 是否保留本节要求的固定分块和顺序。
     - 缺失或不适用内容是否已标注为“无”或说明原因。
     - 若涉及修复或验收，是否已完成第二轮结构化复核后才输出最终通过结论。
   - 自检未通过时，不得给出最终完成或验收通过结论。

该固定输出结构是强约束，不得改写成自由总结，不得只完成内部动作而不按结构输出给用户。

## 输出要求

最终回答默认包含：

- 修复后 `.docx`
- `AUDIT_REPORT.md`
- `REVIEW_REPORT.md`
- `MANUAL_CONFIRMATION.md`
- `DIFF_SUMMARY.md`
- `REPAIR_LOG.md`
- `FINAL_ACCEPTANCE_REPORT.md`
- `RULE_SUMMARY.md`

若无法完成修复或验收，明确说明当前阶段、阻塞项、已生成文件和建议下一步。
