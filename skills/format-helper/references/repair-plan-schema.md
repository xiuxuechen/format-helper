# repair_plan.yaml 契约

`repair_plan.yaml` 由主控线程在合并第一轮审计结果并完成冲突仲裁后生成。

必备顶层字段：

- `schema_version`
- `repair_plan_id`
- `created_at`
- `based_on_plan_id`
- `based_on_snapshot`
- `rule_profile`
- `source_docx`
- `working_docx`
- `output_docx`
- `conflict_resolution`
- `actions`
- `manual_review_items`
- `execution_order`
- `post_repair`

每个 `actions` 项必须包含：

- `action_id`
- `source_issue_ids`
- `action_type`
- `format_write_strategy`
- `target`
- `before`
- `after`
- `auto_fix_policy`
- `risk_level`
- `status`

修复脚本只能执行 `auto_fix_policy: auto-fix` 的动作。人工确认项必须进入 `manual_review_items`，不得混入自动动作。

允许的 `action_type`：

- `map_heading_native_style`
- `apply_body_style_definition`
- `apply_body_direct_format`
- `apply_table_cell_format`
- `apply_table_border`
- `toc_content_audit`
- `insert_or_replace_toc_field`

允许的 `format_write_strategy`：

- `style-definition`
- `direct-format-override`
- `preserve`

直接格式覆盖必须显式使用 `direct-format-override`；缺失目标原生样式时不得自动创建新样式。

默认执行顺序：

1. `normalize_styles`
2. `apply_page_section_rules`
3. `apply_heading_styles`
4. `apply_body_styles`
5. `apply_table_safe_fixes`
6. `toc_content_audit`
7. `replace_or_insert_auto_toc`
8. `refresh_fields_or_mark_for_update`
9. `save_repaired_docx`
