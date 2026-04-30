---
name: format-helper
description: 使用时机：当用户请求 Word/DOCX 格式治理、标准文档建规、规则选择、待处理文档审计、自动修复、二轮复核、报告生成、run_id 恢复，或明确提到 format-helper、$format-helper 时使用。本技能是 v3 唯一外部入口，负责按 Gate 编排内部 docx-* 技能；普通用户不应直接调用内部 docx-* 技能。
---

# Format Helper

## 定位

v3 唯一外部入口。接收用户的 `.docx` 格式治理请求，创建 `format_runs/{run_id}/`，编排内部 `docx-*` 能力完成建规、审计、修复、复核、报告和恢复。

## 强制边界

- 不直接覆盖原始 `.docx`；只复制输入到运行目录并处理工作副本。
- 不直接写 Word；最终写回只能由 `docx-format-repairer` 执行已校验白名单动作。
- 不绕过规则确认、规则选择、参数覆盖、修复确认和验收 Gate。
- 不把机器可读 JSON/YAML 原样堆给用户；最终解释必须使用中文报告或摘要。
- 根目录 `skills/` 仅作为历史参考，v3 新入口位于 `.codex/skills/`。

## 内部能力

- `docx-fact-extractor`：生成事实快照。
- `docx-semantic-strategist`：生成语义规则、角色映射和语义审计。
- `docx-rule-packager`：把语义规则草案打包为规则包和 `RULE_SUMMARY.md`。
- `docx-format-auditor`：执行真实格式审计和二轮复核。
- `docx-repair-planner`：生成可校验的 `repair_plan.yaml`。
- `docx-format-repairer`：只对工作副本执行白名单修复。
- `docx-format-reporter`：生成中文报告和最终验收说明。

## 工作流

1. 确认当前环境和文本文件编码；`.docx` 始终按 OOXML/ZIP 处理。
2. 解析用户意图：`extract-rule`、`audit-only`、`repair` 或 `resume`。
3. 初始化或读取 `format_runs/{run_id}/`。
4. 按运行模式调用内部能力并记录 Gate 状态。
5. 对低置信度、高风险或缺失产物停止推进，输出人工确认项。
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

## 按需读取

- `references/workflow.md`：运行模式和 Gate 顺序。
- `references/artifact-routing.md`：运行目录和恢复产物。
- `references/plan-schema.md`：`PLAN.yaml` 最小契约。
- `references/repair-plan-schema.md`：`repair_plan.yaml` 最小契约。
