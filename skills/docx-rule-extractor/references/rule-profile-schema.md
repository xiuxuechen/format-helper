# DOCX 规则包 Schema

规则包 v2 使用“内部角色 + Word 原生样式 / 真实格式”模型。规则包由标准 `.docx` 运行时抽取生成，默认状态为 `draft`，用户确认后才可作为后续文档处理规则。

## profile.yaml

```yaml
id: project-application-v1
rule_schema_version: 2.0.0
name: project-application-v1
description: 中文规则说明
version: 1.0.0
status: draft
based_on:
  - 标准文件.docx
document_types:
  - 正式申报材料
features:
  auto_toc: true
  style_driven: true
  landscape_sections: true
  appendix_tables: true
toc:
  levels: 3
  include_patterns:
    - 一级标题
risk_level: medium
change_summary: 初始版本
last_updated: 2026-04-29
```

## 必备规则文件

- `role-map.yaml`：内部角色定义，例如 `cover-title`、`heading-level-1`、`body-paragraph`、`table-header`。
- `style-map.yaml`：内部角色到 Word 原生样式或格式策略的映射，不得生成 `Official*` 样式。
- `element-rules.yaml`：元素分类、识别线索和置信度策略。
- `toc-rules.yaml`：是否使用自动目录、是否替换静态目录、纳入目录的标题层级、目录内容校验策略。
- `table-rules.yaml`：普通表格、专栏表格、附表、横向页表格的审计与修复边界。
- `page-rules.yaml`：页边距、纸张、方向、节和页眉页脚策略。
- `risk-policy.yaml`：自动修复、人工确认、阻塞项策略。

## style-map.yaml

```yaml
styles:
  heading-level-1:
    word_style_id: Heading1
    word_style_name: Heading 1
    localized_style_candidates: [标题 1, Heading 1]
    outline_level: 1
    write_strategy: style-definition
    format:
      font_east_asia: 仿宋_GB2312
      font_size_pt: 22
      bold: true

  body-paragraph:
    word_style_id: Normal
    word_style_name: Normal
    localized_style_candidates: [正文, Normal]
    write_strategy: style-definition
    format:
      font_east_asia: 仿宋_GB2312
      font_size_pt: 10.5
      line_spacing_multiple: 1.5
```

写回时优先使用 `word_style_id`；`localized_style_candidates` 只用于审计和兼容匹配。若目标原生样式不存在，修复器不得自动创建新样式，应将动作标记为阻塞或人工确认。

允许的 `write_strategy`：

- `style-definition`：修改原生样式定义，默认策略。
- `direct-format-override`：写入段落或 run 直接格式，只能用于局部例外，并必须在修复计划和报告中说明。
- `preserve`：保留原格式，仅审计或人工确认。

## table-rules.yaml

表格规则必须显式声明行高和边框策略：

```yaml
tables:
  header:
    font_east_asia: 仿宋_GB2312
    font_size_pt: 10.5
    bold: true
    line_spacing_multiple: 1.15
    alignment: center
    shading_fill: 323E4F

  body:
    font_east_asia: 仿宋_GB2312
    font_size_pt: 10.5
    bold: false
    line_spacing_multiple: 1.15

  row_height:
    mode: audit-only  # skip | audit-only | auto-fix
    value: null
    rule: null
    confirmed_by: null
    confirmed_at: null

  border:
    mode: audit-only  # skip | audit-only | auto-fix
    top: { style: single, width_pt: 0.5, color: "000000" }
    bottom: { style: single, width_pt: 0.5, color: "000000" }
```

行高只有在 `hRule` 一致、均为明确数值、标准文档稳定抽取且用户确认后才能设置为 `auto-fix`。标准文档无表格时，行高使用 `mode: skip`。

边框示例值不是默认值。`auto-fix` 的边框值必须来自标准文档稳定抽取或用户显式确认；未经确认的边框规则必须保持 `audit-only` 或 `skip`。
