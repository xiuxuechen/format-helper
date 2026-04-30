---
name: docx-rule-extractor
description: Internal DOCX rule extraction skill. Use only when format-helper needs to inspect a standard Word .docx file, infer a reusable format rule profile, generate RULE_SUMMARY.md, and create draft YAML rule files for user confirmation.
---

# DOCX Rule Extractor

## 定位

内部能力。只负责从用户提供的标准 `.docx` 中抽取规则草案，不处理待修复文档，不写最终文档。

## 输入

- 标准 `.docx`
- 输出规则目录，例如 `format_rules/docx/rule_profiles/{rule-id}/`
- 用户确认后的英文规则 `name` 和中文 `description`

## 输出

- `profile.yaml`
- `role-map.yaml`
- `style-map.yaml`
- `element-rules.yaml`
- `toc-rules.yaml`
- `table-rules.yaml`
- `page-rules.yaml`
- `risk-policy.yaml`
- `RULE_SUMMARY.md`

其中 `RULE_SUMMARY.md` 是给业务用户确认的规则说明书，不是内部规则清单。正文必须使用用户能理解的格式语言，明确写出字体、字号、加粗、对齐、缩进、行距、页边距、页眉页脚、目录、表格等规则。不得用内部样式名代替规则内容。

## 工作流

1. 确认当前环境和输出目录；`.docx` 按 ZIP/OOXML 读取。
2. 使用 `scripts/inspect_docx_profile.py` 生成标准文档结构画像。
3. 识别页面节、标题样式、正文样式、表格、静态目录和自动 TOC 字段。
4. 使用 `scripts/extract_rule_profile.py` 生成规则草案。
5. 将规则状态设为 `draft`。
6. 交由 `format-helper` 展示 `RULE_SUMMARY.md` 并等待用户确认。

## 规则抽取原则

- 不内置具体业务规则；规则必须来自标准文件和用户确认。
- 标题识别不能只看样式名，要结合编号模式、字体、字号、加粗、大纲级别。
- `RULE_SUMMARY.md` 必须展示明确规则值；无法稳定提取的项目必须标为“需人工确认”，不能展示内部占位。
- 自动目录规则必须声明纳入范围；默认以规则版本配置为准。
- 不得生成 `Official*` 自定义样式映射；内部角色必须映射到 Word 原生样式 ID 或真实格式策略。
- 标题和正文默认采用 `style-definition` 写回策略；直接格式覆盖只能作为局部例外并显式标记。
- 表格行高和边框必须生成显式 `mode`；无法稳定抽取或未获用户确认时使用 `audit-only` 或 `skip`。
- 表格复杂修复边界只生成规则建议，不直接承诺自动修复。

## 按需读取

- `references/rule-profile-schema.md`：规则包字段。
- `references/extraction-checklist.md`：抽取核对清单。
- `references/RULE_SUMMARY_TEMPLATE.md`：生成 `RULE_SUMMARY.md` 的固定模板，按用户打开文档后的可视优先级排序。
- `references/rule-summary-template.md`：模板使用原则说明。
- `references/word-format-model.md`：基于 WordprocessingML 和 python-docx 的格式规则模型。
