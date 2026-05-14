# format-helper

`format-helper` 是一个面向 Word/DOCX 格式治理的 Codex-only 框架，用来从标准文档建规、对待处理文档做审计与修复、生成二轮复核和中文报告，并把全过程产物统一落到 `format_runs/{run_id}/`。

更通俗一点说：

- 你给它一份**格式没有问题的 Word 标准文档**
- 它会自动抽取这份文档里的标题、正文、列表、表格、目录、分页等格式特征
- 然后生成一套可以复用的“规则包”
- 你确认这套规则没问题之后，再给它一份**需要按这套标准修复的 Word 文档**
- 它就会根据规则对待处理文档做审计、修复、复核，并生成最终报告

它不是简单地“批量改字体字号”的小脚本，而是一套带规则确认、风险控制、修复追溯和最终验收的 Word 格式治理流程。

## 它能做什么

- 帮你把一份“格式正确的标准 Word”变成可复用的规则
- 帮你检查另一份 Word 到底哪些地方没有按标准来
- 在可控范围内自动修复这些格式问题
- 修完以后再做一轮复核，避免“修坏了”或者“修得不完整”
- 生成给人看的中文报告，而不是只丢一堆 JSON/YAML
- 运行中断后，可以按 `run_id` 接着上次的进度继续

## 一个典型流程

最常见的使用方式是这样的：

1. 先拿一份格式没有问题的标准文档建规
2. 确认 `RULE_SUMMARY.md` 里的规则和你的真实预期一致
3. 再提供需要修复的待处理文档
4. 让 `format-helper` 按已确认规则执行审计和修复
5. 查看修复后的 Word、最终报告和验收结果

如果中间有低置信度、冲突样式或高风险动作，它不会闷头继续跑，而是会停下来让你确认。

## Quickstart

### 1. 准备文档

把标准 `.docx` 或待处理 `.docx` 放到工作区里。

建议按下面的节奏使用：

- 第一次使用：先准备一份**标准文档**
- 规则确认完成后：再准备一份**待修复文档**

### 2. 让 Codex 处理

直接用中文告诉 Codex 你要什么：

- `请使用 format-helper 为 xxx.docx 建规`
- `请使用 format-helper 审计 xxx.docx`
- `请使用 format-helper 修复 xxx.docx`
- `请恢复 run_id=...`

如果你是第一次使用，最推荐这样开头：

```text
请使用 format-helper 为这份标准 Word 建规，生成规则摘要给我确认。
```

确认规则没问题之后，再继续：

```text
请使用刚才确认过的规则，修复这份待处理 Word。
```

### 3. 看结果

常见交付物：

- 规则摘要：`format-rules/{rule_id}/RULE_SUMMARY.md`
- 最终报告：`format_runs/{run_id}/reports/FINAL_ACCEPTANCE_REPORT.md`
- 最终状态：`format_runs/{run_id}/logs/final_acceptance.json`
- 运行状态：`format_runs/{run_id}/logs/state.yaml`

## 直接跑脚本

如果你想直接跑底层脚本，可以用这些入口：

```powershell
python scripts/ooxml/extract_docx_snapshot.py <docx> --snapshot-kind before --output format_runs/<run_id>/snapshots/document_snapshot.before.json
python .codex/skills/docx-rule-packager/scripts/render_rule_summary.py --slot-facts <slot_facts.json> --contract docs/v4/schemas/role_slot_contract.yaml --output format-rules/<rule_id>/RULE_SUMMARY.md
python .codex/skills/docx-format-reporter/scripts/render_final_reports.py --run-dir format_runs/<run_id>
```

## 工作模式

- `extract-rule`：标准文档 -> 事实快照 -> 语义规则 -> 规则包 -> `RULE_SUMMARY.md`
- `audit-only`：待处理文档 -> 事实快照 -> 格式审计 -> 审计报告
- `repair`：审计 -> 修复计划 -> 安全写回 -> 复核 -> 最终报告
- `resume`：读取既有 `format_runs/{run_id}/logs/state.yaml` 继续执行

## 目录概览

| 路径 | 作用 |
| --- | --- |
| `.codex/skills/` | 框架的核心技能入口 |
| `scripts/` | 通用脚本、校验器和报告工具 |
| `schemas/` | JSON/YAML 契约与示例 |
| `docs/v4/` | 设计说明、开发规范和接口契约 |
| `format-rules/` | 已生成的规则包 |
| `format_runs/` | 每次运行的产物目录 |
| `tests/` | 回归测试与 fixture |

## 约定

- 不覆盖原始 `.docx`
- `format-rules/` 和 `format_runs/` 都是框架运行产物
- 用户可读报告优先使用中文，不直接暴露内部 JSON/YAML
- 真正给用户看的主要是修复后 `.docx`、`RULE_SUMMARY.md` 和最终中文报告

## 参考文档

- `docs/v4/10_CONTEXT.md`
- `docs/v4/20_API_SPEC.md`
- `docs/v4/41_SCHEMA_CONTRACTS.md`
- `docs/v4/50_DEV_PLAN.md`
