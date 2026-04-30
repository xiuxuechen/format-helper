---
name: docx-fact-extractor
description: 内部 DOCX 事实抽取技能。仅当 format-helper 需要从标准或待处理 .docx 生成 document_snapshot.json、standard_snapshot.json、before/after snapshot 时使用；只读取 OOXML 事实，不判断标题、正文、目录等语义角色，不修改 Word 文件。
---

# DOCX Fact Extractor

## 定位

内部能力。读取 `.docx` 的 OOXML 结构，输出客观事实快照，作为语义策略、审计、复核和验收的输入。

## 输入

- `docx_path`：标准文档、待处理文档、工作副本或修复后副本。
- `output_snapshot`：快照输出路径，例如 `format_runs/{run_id}/snapshots/document_snapshot.before.json`。
- `snapshot_kind`：`standard`、`before` 或 `after`。

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

## 首阶段落地脚本

```powershell
python scripts/ooxml/extract_docx_snapshot.py input/working.docx --snapshot-kind before --output snapshots/document_snapshot.before.json
```

脚本只读取 OOXML 事实，输出段落、表格、节、样式 ID、直接格式和文件哈希，不判断语义角色。
