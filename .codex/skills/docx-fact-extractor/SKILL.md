---
name: docx-fact-extractor
description: 使用时机：内部 DOCX 事实抽取技能。仅当 format-helper 需要从标准或待处理 .docx 生成 document_snapshot.json、standard_snapshot.json、before/after snapshot 时使用；只读取 OOXML 事实，不判断标题、正文、目录等语义角色，不修改 Word 文件。
---

# DOCX Fact Extractor

## 定位

内部能力。读取 `.docx` 的 OOXML 结构，输出客观事实快照，作为语义策略、审计、复核和验收的输入。

## 输入

- `docx_path`：标准文档、待处理文档、工作副本或修复后副本。
- `output_snapshot`：快照输出路径，例如 `format_runs/{run_id}/snapshots/document_snapshot.before.json`。
- `snapshot_kind`：`standard`、`before` 或 `after`。
- `--with-source`（默认开启）：生成带格式来源标注的 `resolved_paragraph_format` / `resolved_run_format`，每个格式字段均为 `{value, source, confidence}` 对象。source 枚举为 `direct`、`style_inherit`、`doc_defaults`、`theme`、`word_ui_default`、`unresolved`、`legacy`。关闭 `--without-source` 时回退为旧式基础值快照。

## 输出

- `snapshots/standard_snapshot.json`
- `snapshots/document_snapshot.before.json`
- `snapshots/document_snapshot.after.json`
- 可复用脚本：`scripts/ooxml/extract_docx_snapshot.py`

## 强制边界

- 不判断语义角色，不输出最终 `semantic_role`。
- 不修改 `.docx`，不创建样式，不刷新目录。
- 不以终端显示乱码判断文件内容异常；`.docx` 只按 ZIP/OOXML 读取。
- 快照字段必须可被后续 schema 校验和恢复流程引用。

## 工作流

1. 确认输入 `.docx` 存在且可作为 ZIP 打开。
2. 读取段落、run、样式 ID、直接格式、表格、节、编号、目录字段和页眉页脚事实。
3. 生成稳定 `element_id`，例如 `p-00012`、`table-0003`。
4. 写入 JSON 快照，并保留输入文件哈希和生成时间。
5. 若文档损坏或关键 OOXML 缺失，返回阻塞项给 `format-helper`。

## 固定执行步骤（参考 40-§6.10）

1. 确认运行环境和文本文件编码
2. 按 OOXML/ZIP 读取输入 Word
3. 生成 snapshot（standard/before/after）
4. 执行 snapshot schema 校验
5. 写入双通道输出：
   - 业务产物：`snapshots/*.json`
   - 状态信封：`logs/skill_results/{seq}_docx-fact-extractor.result.json`

## 双通道输出协议（参考 40-§6.4）

每次执行必须同时输出：

1. **业务产物**（机器权威）：
   - `snapshots/standard_snapshot.json`
   - `snapshots/document_snapshot.before.json`
   - `snapshots/document_snapshot.after.json`

2. **状态信封**（机器权威）：
   - 路径：`logs/skill_results/{seq}_docx-fact-extractor.result.json`
   - Schema：`skill-result`（参考 41-§5）
   - 必须包含：`result_id`、`status`、`schema_valid`、`gate_passed`、`artifacts`、`next_action`

## 成功输出模板（参考 40-§6.10）

```text
任务清单
1. ✅ 已确认运行环境和文本文件编码
2. ✅ 已按 OOXML/ZIP 读取输入 Word：{input_docx}
3. ✅ 已生成 {snapshot_kind} 快照
4. ✅ 已执行 snapshot schema 校验

当前阶段
fact_extraction

执行结果
✅ 成功

交付物
- 快照：{snapshot_path}
- 快照 hash：{snapshot_sha256}
- 快照大小：{snapshot_size_bytes} bytes
- 段落数：{paragraph_count}
- 表格数：{table_count}
- 状态信封：logs/skill_results/{seq}_docx-fact-extractor.result.json

阻塞/人工确认
无

下一步
- 快照已就绪，可进入语义策略或格式审计阶段

验收自检
- [x] 快照 schema 校验通过
- [x] 快照包含段落、run、样式 ID、直接格式、表格、节、编号、目录字段和页眉页脚事实
- [x] 生成稳定 element_id（如 p-00012、table-0003）
- [x] 保留输入文件 hash 和生成时间
- [x] 未判断语义角色
- [x] 未修改 Word 文件
```

## 失败输出模板（参考 40-§6.10）

```text
任务清单
1. ✅ 已确认运行环境和文本文件编码
2. ❌ 读取输入 Word 失败

当前阶段
fact_extraction

执行结果
❌ 失败 / ⏸️ 阻塞

交付物
- 部分快照（如有）：{partial_snapshot_path}
- 状态信封：logs/skill_results/{seq}_docx-fact-extractor.result.json

阻塞/人工确认
阻塞原因：
- 错误码：{error_code}（如 FE-DOCX-CORRUPT、FE-OOXML-MISSING）
- 错误消息：{error_message}
- 恢复建议：{recovery_suggestion}

下一步
- 如果文档损坏：请检查输入 Word 文件是否完整
- 如果关键 OOXML 缺失：请使用 Word 打开并另存为新文件
- 如果路径问题：请检查文件路径是否正确

验收自检
- [ ] 快照 schema 校验通过
- [x] 已记录失败原因和错误码
- [x] 已生成状态信封（status=blocked 或 synthetic_failure）
```

## 首阶段落地脚本

```powershell
python scripts/ooxml/extract_docx_snapshot.py input/working.docx --snapshot-kind before --output snapshots/document_snapshot.before.json
```

脚本只读取 OOXML 事实，输出段落、表格、节、样式 ID、直接格式和文件哈希，不判断语义角色。
