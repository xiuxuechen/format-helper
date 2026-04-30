# format-helper v3 工作流

## 运行模式

- `extract-rule`：标准 Word -> 事实快照 -> 语义规则草案 -> 规则包 -> 规则确认 Gate。
- `audit-only`：待处理 Word -> 事实快照 -> 语义角色 -> 格式审计 -> 审计报告。
- `repair`：审计 -> 修复计划 -> 修复确认 Gate -> 安全写回 -> after snapshot -> 二轮复核 -> 报告。
- `resume`：读取 `format_runs/{run_id}/logs/state.yaml`，从最近安全阶段恢复。

## Gate 顺序

1. 规则确认 Gate。
2. 规则选择 Gate。
3. 参数覆盖 Gate。
4. 修复确认 Gate。
5. 验收 Gate。

任何 Gate 未通过时，停止推进并输出人工确认项。
