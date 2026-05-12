---
name: docx-format-auditor
description: 使用时机：内部 DOCX 格式审计技能。仅当 format-helper 需要基于 document_snapshot.json、semantic_role_map.before.json 和已确认规则包执行真实格式审计，或基于 before/after snapshot 做二轮结构化复核时使用；只输出审计或复核 JSON，不修改 Word 文件。
---

# DOCX Format Auditor

## 定位

内部能力。消费事实快照、语义角色映射和规则包，执行真实格式数值比对；也用于修复后的二轮复核。

## 输入

- `PLAN.yaml`
- `document_snapshot.before.json` 或 `document_snapshot.after.json`
- `semantic_role_map.before.json`
- 已确认规则包
- 任务范围，例如标题、正文、目录、表格、页面或样式治理

## 输出

- `audit_results/{task_id}.audit.json`
- `review_results/{task_id}.review.json`
- 可复用脚本：`scripts/build_second_round_review.py`

## 强制边界

- 不修改 `.docx`。
- 不以 Word 内部样式名作为唯一通过依据。
- 不输出最终验收结论；最终汇总由 `format-helper` 和 `docx-format-reporter` 完成。
- 每个问题必须包含 `element_id`、当前事实、期望规则、建议动作、置信度和风险等级。

## 工作流

1. 读取快照、语义角色映射和规则包。
2. 按任务范围筛选元素。
3. 比对真实格式、大纲级别、目录字段、表格和页面设置。
4. 输出结构化问题或复核项。
5. 对低置信度、高风险或复杂结构标记人工确认。

## 固定执行步骤（参考 40-§6.13）

1. 读取规则引用
2. 读取文档快照
3. 读取 semantic-role-map 并校验 source hash
4. 执行模式（format_audit 或 review）
5. 生成结构化结果
6. 写入双通道输出：
   - 业务产物：`audit_results/*.audit.json` 或 `review_results/*.review.json`
   - 状态信封：`logs/skill_results/{seq}_docx-format-auditor.result.json`

## 双通道输出协议（参考 40-§6.4）

每次执行必须同时输出：

1. **业务产物**（机器权威）：
   - `audit_results/{task_id}.audit.json`（mode=format_audit）
   - `review_results/{task_id}.review.json`（mode=review）
   - 渲染证据（如需要）

2. **状态信封**（机器权威）：
   - 路径：`logs/skill_results/{seq}_docx-format-auditor.result.json`
   - Schema：`skill-result`（参考 41-§5）
   - 必须包含：`result_id`、`status`、`schema_valid`、`gate_passed`、`artifacts`、`next_action`

## 成功输出模板（参考 40-§6.13）

```text
任务清单
1. ✅ 已读取规则引用：{rule_ref}
2. ✅ 已读取文档快照：{snapshot_path}
3. ✅ 已读取 semantic-role-map 并校验 source hash：{semantic_role_map_path}
4. ✅ 已执行模式：{audit_mode=format_audit|review}
5. ✅ 已按模式生成结构化结果

当前阶段
{format_audit|review}

执行结果
✅ 成功
- 问题总数：{issue_count}
- 可自动修复：{auto_fixable_count}
- 需人工确认：{manual_review_count}
- 复核状态：{review_summary|不适用}

交付物
- 审计结果：{audit_result_path|不适用}
- 复核结果：{review_result_path|不适用}
- 渲染证据：{render_dir}
- 状态信封：logs/skill_results/{seq}_docx-format-auditor.result.json

阻塞/人工确认
{无；或列出高风险结构、渲染疑似空白页、目录未刷新、表格跨页问题}

下一步
- 若 audit_mode=format_audit 且存在可修复问题，进入 docx-repair-planner
- 若 audit_mode=review 且已完成二轮复核，进入 toc_acceptance 或 final_acceptance

验收自检
- [x] 未修改 Word
- [x] 真实格式比对完成
- [x] audit_mode=format_audit 时未强制生成 review-result
- [x] audit_mode=review 时 checks[] 使用 ReviewCheck Object
- [x] semantic-role-map/source snapshot/rule manifest hash 已记录
- [x] 目录字段检查完成
- [x] 表格检查完成
```

## 失败输出模板（参考 40-§6.13）

```text
任务清单
1. ✅ 已检查规则引用：{rule_ref}
2. ❌ 审计/复核未完成

当前阶段
{format_audit|review}

执行结果
❌ 失败 / ⏸️ 阻塞

交付物
- 部分审计结果（如有）：{partial_output_path}
- 状态信封：logs/skill_results/{seq}_docx-format-auditor.result.json

阻塞/人工确认
阻塞原因：
- 错误码：{error_code}（如 FA-RULE-MISSING、FA-SNAPSHOT-INVALID、FA-SCHEMA-FAILED）
- 错误消息：{error_message}
- 恢复建议：{recovery_suggestion}

下一步
- 如果规则包缺失：请先完成规则选择或建规流程
- 如果快照缺失：请先完成事实抽取
- 如果 schema 不通过：请检查快照或规则包格式
- 如果渲染证据不足：请检查 Word 文件是否完整

验收自检
- [x] 未修改 Word
- [ ] 可进入修复计划或报告阶段
- [x] 已记录失败原因和错误码
- [x] 已生成状态信封（status=blocked 或 synthetic_failure）
```

## 首阶段落地脚本

```powershell
python .codex/skills/docx-format-auditor/scripts/build_second_round_review.py --run-dir format_runs/code005-fixture
```

脚本生成 T01-T06 二轮复核 JSON：OOXML 完整性、快照完整性、执行日志、自动动作追溯、人工确认留痕和渲染证据。
