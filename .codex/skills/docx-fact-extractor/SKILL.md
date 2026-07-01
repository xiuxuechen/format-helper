---
name: docx-fact-extractor
description: 使用时机：内部 DOCX 事实抽取技能。仅当 format-helper 需要通过 OfficeCLI 生成 officecli-document-snapshot v2 的 standard/before/after/post_toc 快照时使用；不判断语义角色，不修改 Word 文件。
---

# DOCX Fact Extractor

## 定位

内部能力。通过固定版本 OfficeCLI 读取 `.docx`，输出 `officecli-document-snapshot` v2，作为语义、审计、复核和验收的唯一事实输入。

## 输入

- `docx_path`：标准文档、待处理文档、工作副本或修复后副本。
- `output_snapshot`：`snapshots/officecli-document-snapshot.{standard|before|after|post_toc}.json`。
- `snapshot_kind`：`standard`、`before`、`after` 或 `post_toc`。
- `officecli_executable`：由锁文件解析并通过 hash/version 校验的可执行文件。
- `capability_manifest`：固定版本能力清单。

## 输出

- `snapshots/officecli-document-snapshot.standard.json`
- `snapshots/officecli-document-snapshot.before.json`
- `snapshots/officecli-document-snapshot.after.json`
- `snapshots/officecli-document-snapshot.post_toc.json`
- 生产脚本：`scripts/officecli/snapshot_adapter.py`

## 强制边界

- 不判断语义角色，不输出最终 `semantic_role`。
- 不修改 `.docx`，不创建样式，不刷新目录。
- Python 禁止直接使用 ZIP/XML DOM 读取 DOCX；事实只能来自 OfficeCLI JSON/raw 证据。
- 快照字段必须可被后续 schema 校验和恢复流程引用。

## 工作流

1. 解析锁定 OfficeCLI 并完成版本/能力自检。
2. 使用 BFS get、query completeness、dump warning 和 required raw parts 采集事实。
3. 生成 path、logical identity、native identity 和类型索引。
4. 写入 snapshot v2，并保留输入文件 hash、raw ArtifactRef 和 GateCheck。
5. 若遍历不完整、路径冲突或 blocking warning，返回阻塞项给 `format-helper`。

## 固定执行步骤（参考 40-§6.10）

1. 确认运行环境和文本文件编码
2. 通过 OfficeCLI 读取工作副本
3. 生成 snapshot（standard/before/after）
4. 执行 snapshot schema 校验
5. 写入双通道输出：
   - 业务产物：`snapshots/*.json`
   - 状态信封：`logs/skill_results/{seq}_docx-fact-extractor.result.json`

## 双通道输出协议（参考 40-§6.4）

每次执行必须同时输出：

1. **业务产物**（机器权威）：
   - `snapshots/officecli-document-snapshot.standard.json`
   - `snapshots/officecli-document-snapshot.before.json`
   - `snapshots/officecli-document-snapshot.after.json`

2. **状态信封**（机器权威）：
   - 路径：`logs/skill_results/{seq}_docx-fact-extractor.result.json`
   - Schema：`skill-result`（参考 41-§5）
   - 必须包含：`result_id`、`status`、`schema_valid`、`gate_passed`、`artifacts`、`next_action`

## 成功输出模板（参考 40-§6.10）

```text
任务清单
1. ✅ 已确认运行环境和文本文件编码
2. ✅ 已通过锁定 OfficeCLI 读取输入 Word：{input_docx}
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
- 节点数：{node_count}
- 部件数：{part_count}
- 状态信封：logs/skill_results/{seq}_docx-fact-extractor.result.json

阻塞/人工确认
无

下一步
- 快照已就绪，可进入语义策略或格式审计阶段

验收自检
- [x] 快照 schema 校验通过
- [x] BFS/query/dump/raw completeness Gate 通过
- [x] 生成稳定 OfficeCLI path 与 logical identity
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
- 错误码：{error_code}（如 FH-OFFICECLI-SNAPSHOT-COMPLETENESS、FH-OFFICECLI-SNAPSHOT-LIMIT）
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
python scripts/officecli/snapshot_adapter.py --help
```

旧 `scripts/ooxml/extract_docx_snapshot.py` 仅作为 v4 迁移输入保留，不得由 v5 生产流程调用。
