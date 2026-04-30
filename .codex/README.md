# format-helper Codex v3

本目录承载 v3 Codex-only 能力入口。普通用户只调用 `.codex/skills/format-helper/`，内部 `docx-*` 技能由 `format-helper` 编排。

## 能力入口

- `format-helper`：唯一外部入口，负责运行编排、Gate、恢复和最终输出。
- `docx-fact-extractor`：只读抽取 `.docx` 客观事实快照。
- `docx-semantic-strategist`：生成语义规则、角色映射和语义审计 JSON。
- `docx-rule-packager`：打包规则包和 `RULE_SUMMARY.md`。
- `docx-format-auditor`：执行真实格式审计和二轮复核。
- `docx-repair-planner`：生成 `repair_plan.yaml` 和人工确认项。
- `docx-format-repairer`：只对工作副本执行白名单安全写回。
- `docx-format-reporter`：生成中文报告和最终验收说明。

## 运行产物

- 规则资产：`format_rules/{rule_id}/`
- 运行目录：`format_runs/{run_id}/`
- 契约文件：`schemas/*.schema.json`

## 安全边界

- 原始 `.docx` 不覆盖。
- AI 只输出结构化 JSON/YAML，不直接写 Word。
- Python 只负责事实抽取、schema 校验和安全写回。
- 高风险、低置信度、复杂结构默认进入人工确认。
- 根目录 `skills/` 是历史参考，不作为 v3 新能力入口。
