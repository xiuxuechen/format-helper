# 审计输出 Schema

```json
{
  "schema_version": "2.0.0",
  "task_id": "T03",
  "agent": "heading-agent",
  "phase": "audit",
  "source_snapshot": "document_snapshot.before.json",
  "summary": {
    "checked_items": 42,
    "issues_found": 8,
    "auto_fixable": 6,
    "manual_review": 2,
    "blocked": 0
  },
  "issues": []
}
```

每个 issue 必须包含：

- `issue_id`
- `issue_type`
- `severity`
- `confidence`
- `element_ref`
- `detected_role`
- `expected_role`
- `problem`
- `format_source`
- `current_format`
- `expected_format`
- `recommended_action`
- `risk_flags`

## issue_type

允许值：

- `role_classification`：元素角色识别不确定或错误。
- `style_mapping`：角色到 Word 原生样式 / 大纲级别映射不符合规则。
- `format_mismatch`：真实字体、字号、行距、缩进、底纹、边框等格式不符合规则。
- `toc_content_mismatch`：静态目录条目与将生成的自动目录条目不一致。
- `high_risk_structure`：合并单元格、横向页表格、页眉页脚、脚注尾注等高风险结构。

## 标题置信度

- `confidence >= 0.85`：高置信度，且目标原生样式存在、影响范围审计可通过时才允许自动映射。
- `0.60 <= confidence < 0.85`：中置信度，进入人工确认或只审计。
- `confidence < 0.60`：低置信度，不自动改为标题，也不按正文强制处理。

`ambiguous-numbered-item` 必须作为 `role_classification` 问题输出，不得直接生成 `auto-fix` 修复动作。

## format_source

允许值：

- `style-definition`：格式来自样式定义。
- `direct-format`：格式来自段落或 run 直接格式。
- `inherited`：格式来自继承链。
- `mixed`：同一元素混合了样式定义和直接格式。

复核阶段的 `current_format` / `expected_format` 必须记录真实字段，例如：

- `font_east_asia`
- `font_ascii`
- `font_size_pt`
- `bold`
- `line_spacing_multiple`
- `space_before_pt`
- `space_after_pt`
- `first_line_indent_cm`
- `outline_level`
- `table_cell_margin`
- `border`
- `shading_fill`

不得只用样式名作为通过或失败依据。
