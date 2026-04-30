# 修复计划约束

读取 `format-helper/references/repair-plan-schema.md` 作为主契约。本 skill 只补充执行边界：

- `source_docx` 只用于追溯，不直接写。
- `working_docx` 是唯一输入副本。
- `output_docx` 是唯一写出文件。
- `manual_review_items` 只记录，不执行。
- 未识别动作写入修复日志，状态为 `skipped`。
