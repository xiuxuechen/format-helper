---
name: docx-format-reporter
description: 使用时机：内部 DOCX 报告生成技能。仅当 format-helper 需要从 format_runs/{run_id}/ 的计划、审计、修复、复核、人工确认和最终验收产物生成中文可读报告时使用；不得把内部 JSON 原样作为报告正文。
---

# DOCX Format Reporter

## 定位

内部能力。把机器可读追溯产物转换为用户可读中文报告，说明问题、规则、动作、风险、未修项和验收结论。

## 输入

- `plans/PLAN.yaml`
- `plans/repair_plan.yaml`
- `audit_results/*.audit.json`
- `review_results/*.review.json`
- `logs/state.yaml`
- 规则摘要和人工确认记录

## 输出

- `reports/AUDIT_REPORT.md`
- `reports/REVIEW_REPORT.md`
- `reports/MANUAL_CONFIRMATION.md`
- `reports/DIFF_SUMMARY.md`
- `reports/REPAIR_LOG.md`
- `reports/FINAL_ACCEPTANCE_REPORT.md`
- `logs/reporting_result.json`（报告阶段后置引用）
- `logs/evidence_manifest.reporting.json`（报告阶段证据索引）
- `logs/final_acceptance.json`（**只读引用**，不得修改）
- 可复用脚本：`scripts/render_final_reports.py`

## 强制边界

- 不修改 `.docx`。
- 不把内部 JSON/YAML 原样堆到报告正文。
- 未完成二轮复核时，不输出最终通过结论。
- 报告必须解释剩余风险、人工确认项和阻塞原因。
- **不得反向驱动或改写最终验收状态**（参考 40-§6.16, 41-§11.8）。
- **只读 `logs/final_acceptance.json`，不得修改其内容或追加 `report_refs`**。
- **报告引用只写入 `logs/reporting_result.json`**（参考 41-§11.8.1）。
- 报告失败不得改变 `final_acceptance.status`（即使报告失败，最终验收状态仍保持）。

## 固定执行步骤（参考 40-§6.16）

1. 读取最终验收 JSON（`logs/final_acceptance.json`）
2. 按 acceptance_type/workflow_mode 选择报告分支（`final_delivery`/`audit_only_terminal`/`build_rules_terminal`/`blocked_terminal`）
3. 生成中文最终验收报告
4. 校验本分支 required 产物；非本分支产物不作为失败条件
5. 写入 reporting result 后置引用
6. 写入双通道输出：
   - 业务产物：`reports/*.md` + `logs/reporting_result.json`
   - 状态信封：`logs/skill_results/{seq}_docx-format-reporter.result.json`

## final_acceptance 不可变边界（参考 40-§6.16, 41-§11.8, 41-§11.8.1）

**关键约束**：
- `logs/final_acceptance.json` 一经生成即不可变
- 报告引用只能写入 `logs/reporting_result.json`（reporting-result schema）
- `reporting_result.status` 只能使用 `done` 或 `blocked`
- 报告阶段不得输出 `accepted`、`accepted_with_warnings` 或改写 `final-acceptance.status`
- `final_acceptance_sha256` 必须等于 `final_acceptance.json` 文件的复算 hash
- 如果报告生成失败：
  - reporting_result.status = blocked
  - 但 final_acceptance 状态保持不变
  - 用户输出必须区分"最终 Word 已验收"和"报告阶段失败"

## acceptance_type 分支规则（参考 40-§6.16, 41-§11.8）

| acceptance_type | 必需输入 | 不适用产物 |
|-----------------|---------|-----------|
| `final_delivery` | 最终 Word、复核报告、修复摘要 | - |
| `audit_only_terminal` | 审计报告、审计证据、规则引用 | 不要求最终 Word |
| `build_rules_terminal` | 规则包摘要、规则引用、规则包 manifest | 不要求最终 Word |
| `blocked_terminal` | 阻断报告、可恢复入口和 blockers | 不要求补齐非适用分支产物 |

## 双通道输出协议（参考 40-§6.4）

每次执行必须同时输出：

1. **业务产物**（机器权威）：
   - `reports/FINAL_ACCEPTANCE_REPORT.md`（主报告）
   - `reports/AUDIT_REPORT.md`（审计报告）
   - `reports/REVIEW_REPORT.md`（复核报告）
   - `reports/MANUAL_CONFIRMATION.md`（人工确认）
   - `reports/DIFF_SUMMARY.md`（差异摘要）
   - `reports/REPAIR_LOG.md`（修复日志）
   - `logs/reporting_result.json`（报告阶段后置引用）
   - `logs/evidence_manifest.reporting.json`（报告阶段证据索引）

2. **状态信封**（机器权威）：
   - 路径：`logs/skill_results/{seq}_docx-format-reporter.result.json`
   - Schema：`skill-result`（参考 41-§5）
   - 必须包含：`result_id`、`status`、`schema_valid`、`gate_passed`、`artifacts`、`next_action`
   - **status 只能为 `done` 或 `blocked`**（不得使用 accepted 类状态）

## 成功输出模板（参考 40-§6.16）

```text
任务清单
1. ✅ 已读取最终验收 JSON，并按 acceptance_type/workflow_mode 选择报告分支
2. ✅ 已生成中文最终验收报告
3. ✅ 已校验本分支 required 产物；非本分支产物未作为失败条件
4. ✅ 已写入 reporting result 后置引用

当前阶段
reporting

执行结果
✅ 成功
- reporter 局部状态：done
- 最终验收状态（logs/final_acceptance.json.status）：{accepted|accepted_with_warnings|blocked}
- 验收分支：{final_delivery|audit_only_terminal|build_rules_terminal|blocked_terminal}
- 最终验收 hash：{final_acceptance_sha256}
- 修复摘要：{final_delivery 时填写；其他分支写"不适用"}
- 复核摘要：{final_delivery 时填写；其他分支写"不适用"}
- 审计摘要：{audit_only_terminal 时填写；其他分支按需摘要}
- 规则包摘要：{build_rules_terminal 时填写；其他分支按需摘要}
- 阻断摘要：{blocked_terminal 时填写；其他分支写"无"}
- 警示项：{warning_summary}

交付物
- 最终验收报告：reports/FINAL_ACCEPTANCE_REPORT.md
- final_delivery：最终 Word output/{原文件名}{yyyyMMddHHmm}(_r[0-9]{2})?.docx、复核报告、修复摘要
- audit_only_terminal：审计报告、审计证据、规则引用；不要求最终 Word
- build_rules_terminal：规则包摘要、规则引用、规则包 manifest；不要求最终 Word
- blocked_terminal：阻断报告、可恢复入口和 blockers；不要求补齐非适用分支产物
- 报告结果：logs/reporting_result.json
- 状态信封：logs/skill_results/{seq}_docx-format-reporter.result.json

内部追踪
- 引用的最终验收 JSON：logs/final_acceptance.json（只读）
- 报告结果：logs/reporting_result.json
- 后置报告证据 generation：logs/evidence_manifest.reporting.json 或无

阻塞/人工确认
{无；或列出本 acceptance_type 分支内仍影响 reporting 的报告生成问题}

下一步
- 交付用户，或按警示项继续人工处理

验收自检
- [x] 当前 acceptance_type 分支选择正确
- [x] 非适用分支产物未被当作失败条件
- [x] 报告未粘贴原始 JSON
- [x] 警示项说明清楚
- [x] final_acceptance 与实际产物一致，且报告未反向改写状态或文件内容
- [x] reporting_result 引用的 final_acceptance hash 匹配
- [x] final_acceptance.json 文件未被修改（hash 保持一致）
- [x] 报告引用只写入 logs/reporting_result.json
```

## 失败输出模板（参考 40-§6.16）

```text
任务清单
1. ✅ 已检查最终验收 JSON、reporting-result 写入条件和当前 acceptance_type 分支 required 产物
2. ❌ 最终报告未完成

当前阶段
reporting

执行结果
❌ 失败 / ⏸️ 阻塞
- reporter 局部状态：blocked
- 最终验收状态（logs/final_acceptance.json.status）：{保持不变，不受报告失败影响}

交付物
- 部分报告（如有）：{partial_report_path}
- 报告结果：logs/reporting_result.json（status=blocked）
- 状态信封：logs/skill_results/{seq}_docx-format-reporter.result.json

阻塞/人工确认
阻塞原因：
- 错误码：{error_code}（如 RPT-FA-MISSING、RPT-HASH-MISMATCH、RPT-BRANCH-INVALID、RPT-WRITE-FAILED）
- 错误消息：{error_message}
- 恢复建议：{recovery_suggestion}

下一步
- 如果缺少本分支 required 报告输入：请补齐 audit、repair、review 等产物
- 如果 reporting_result 无法写入：请检查磁盘空间和权限
- 如果 final_acceptance hash 不匹配：请检查 final_acceptance.json 是否被修改
- 如果报告 artifact 登记失败：请检查 evidence_manifest.reporting.json

验收自检
- [ ] 未生成 accepted 结论报告（reporter 局部 blocked）
- [x] 未暴露内部 JSON 作为正文
- [x] final_acceptance.json 未被修改（保持不变）
- [x] 已记录失败原因和错误码
- [x] 已生成状态信封（status=blocked）
- [x] 报告失败不影响 final_acceptance 最终状态
```

## 首阶段落地脚本

```powershell
python .codex/skills/docx-format-reporter/scripts/render_final_reports.py --run-dir format_runs/code005-fixture
```

脚本读取 `logs/final_acceptance.json`、`logs/repair_execution_log.json`、before/after 快照和 `review_results/*.review.json`，生成全部 Markdown 报告和 `logs/reporting_result.json`。

**路径约束**（参考 41-§11.8, 41-§11.8.1）：
- 脚本**只读** `logs/final_acceptance.json`，不得修改
- 脚本**写入** `logs/reporting_result.json` 作为后置引用
- 报告 artifact 登记在 `logs/evidence_manifest.reporting.json`
