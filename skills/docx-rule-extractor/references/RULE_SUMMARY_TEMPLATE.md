# {{RULE_NAME}}

## 1. 规则来源与适用范围

- 来源标准文件：{{STANDARD_DOCX}}
- 适用文档类型：{{DOCUMENT_TYPE}}
- 规则状态：{{RULE_STATUS}}

## 2. 文档结构总览

| 项目 | 规则内容 |
| --- | --- |
| 页面节 | {{SECTION_SUMMARY}} |
| 段落 | {{PARAGRAPH_SUMMARY}} |
| 表格 | {{TABLE_SUMMARY}} |
| 自动目录 | {{TOC_SUMMARY}} |
| 页眉页脚 | {{HEADER_FOOTER_SUMMARY}} |
| 图片与媒体 | {{MEDIA_SUMMARY}} |
| 编号体系 | {{NUMBERING_SUMMARY}} |
| 脚注、批注、修订 | {{ANNOTATION_SUMMARY}} |

### 覆盖率说明

{{COVERAGE_STATEMENT}}

## 3. 页面与节规则

| 检查项 | 标准规则 | 说明 |
| --- | --- | --- |
| 纸张大小 | {{PAPER_RULE}} | 对应 Word 页面设置 |
| 页面方向 | {{ORIENTATION_RULE}} | 支持同一文档内纵向、横向混排 |
| 页边距 | {{MARGIN_RULE}} | 上、下、左、右边距需逐节核对 |
| 页眉距离 | {{HEADER_DISTANCE_RULE}} | 页眉与页面顶端距离 |
| 页脚距离 | {{FOOTER_DISTANCE_RULE}} | 页脚与页面底端距离 |
| 装订线 | {{GUTTER_RULE}} | 如无装订要求，可为无 |
| 分节策略 | {{SECTION_BREAK_RULE}} | 横向页、附件、附表等不得破坏原分节 |
| 分栏 | {{COLUMN_RULE}} | 多栏正文、附表等需要保留 |
| 文档网格 | {{DOC_GRID_RULE}} | 公文常见固定行距与网格 |
| 文字方向 | {{TEXT_DIRECTION_RULE}} | 横排、竖排或特殊页面 |
| 页面垂直对齐 | {{PAGE_VERTICAL_ALIGN_RULE}} | 顶端、居中、两端等 |

## 4. 页眉页脚规则

| 检查项 | 标准规则 | 说明 |
| --- | --- | --- |
| 首页页眉页脚 | {{FIRST_PAGE_HEADER_FOOTER_RULE}} | 首页是否单独设置 |
| 奇偶页页眉页脚 | {{ODD_EVEN_HEADER_FOOTER_RULE}} | 是否区分奇偶页 |
| 普通页页眉 | {{HEADER_RULE}} | 包含文字、页码、图片时需记录 |
| 普通页页脚 | {{FOOTER_RULE}} | 页码位置、单位名称、密级等需记录 |
| 继承关系 | {{HEADER_FOOTER_LINK_RULE}} | 分节后是否继承上一节 |
| 页码字段 | {{PAGE_NUMBER_FIELD_RULE}} | 页码必须可更新 |

## 5. 封面与题名规则

| 检查项 | 标准规则 |
| --- | --- |
| 总题名 | {{COVER_TITLE_RULE}} |
| 副题名 | {{SUBTITLE_RULE}} |
| 编制单位 | {{ORG_NAME_RULE}} |
| 日期 | {{DATE_RULE}} |
| 封面空行与位置 | {{COVER_LAYOUT_RULE}} |

## 6. 目录规则

| 检查项 | 标准规则 |
| --- | --- |
| 目录形态 | {{TOC_RULE}} |
| 纳入范围 | {{TOC_SCOPE}} |
| 层级显示 | {{TOC_LEVEL_RULE}} |
| 页码显示 | {{TOC_PAGE_NUMBER_RULE}} |
| 超链接 | {{TOC_HYPERLINK_RULE}} |
| 手工目录处理 | {{STATIC_TOC_RULE}} |

## 7. 章节标题规则

| 标题层级 | 字体 | 字号 | 字符效果 | 行间距与段落间距 | 缩进与对齐 | 分页控制 | 是否进入自动目录 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 一级标题 | {{H1_FONT}} | {{H1_SIZE}} | {{H1_CHAR_EFFECTS}} | {{H1_SPACING}} | {{H1_PARAGRAPH_LAYOUT}} | {{H1_PAGE_CONTROL}} | 是 |
| 二级标题 | {{H2_FONT}} | {{H2_SIZE}} | {{H2_CHAR_EFFECTS}} | {{H2_SPACING}} | {{H2_PARAGRAPH_LAYOUT}} | {{H2_PAGE_CONTROL}} | 是 |
| 三级标题 | {{H3_FONT}} | {{H3_SIZE}} | {{H3_CHAR_EFFECTS}} | {{H3_SPACING}} | {{H3_PARAGRAPH_LAYOUT}} | {{H3_PAGE_CONTROL}} | 是 |

### 标题结构要求

- {{HEADING_NUMBERING_RULE}}
- {{HEADING_KEEP_RULE}}
- {{HEADING_OUTLINE_RULE}}

## 8. 正文段落规则

| 项目 | 规则内容 |
| --- | --- |
| 字体 | {{BODY_FONT}} |
| 字号 | {{BODY_SIZE}} |
| 字符效果 | {{BODY_CHAR_EFFECTS}} |
| 缩进 | {{BODY_INDENT}} |
| 行间距与段落间距 | {{BODY_SPACING}} |
| 对齐方式 | {{BODY_ALIGNMENT}} |
| 制表位 | {{BODY_TAB_RULE}} |
| 分页控制 | {{BODY_PAGE_CONTROL}} |

### 正文特殊情况

- 段内加粗、下划线、超链接、脚注引用等保留原意，不按普通正文强行抹平。
- {{BODY_SPECIAL_RULE}}

## 9. 编号与列表规则

| 检查项 | 标准规则 |
| --- | --- |
| 章节编号 | {{HEADING_NUMBER_RULE}} |
| 正文编号 | {{BODY_NUMBER_RULE}} |
| 项目符号 | {{BULLET_RULE}} |
| 多级列表 | {{MULTILEVEL_LIST_RULE}} |
| 编号连续性 | {{NUMBER_CONTINUITY_RULE}} |

## 10. 表格规则

| 检查项 | 标准规则 |
| --- | --- |
| 表格数量参考 | {{TABLE_COUNT}} |
| 表格宽度与对齐 | {{TABLE_WIDTH_RULE}} |
| 表格边框 | {{TABLE_BORDER_RULE}} |
| 表头 | {{TABLE_HEADER_RULE}} |
| 表格正文 | {{TABLE_BODY_RULE}} |
| 行与列 | {{TABLE_ROW_COLUMN_RULE}} |
| 单元格格式 | {{TABLE_CELL_RULE}} |
| 单元格合并 | {{MERGED_CELL_RULE}} |
| 跨页表格 | {{CROSS_PAGE_TABLE_RULE}} |
| 横向页面表格 | {{LANDSCAPE_TABLE_RULE}} |
| 嵌套表格 | {{NESTED_TABLE_RULE}} |

## 11. 图片、脚注、批注与修订规则

| 检查项 | 标准规则 |
| --- | --- |
| 图片 | {{IMAGE_RULE}} |
| 图题表题 | {{CAPTION_RULE}} |
| 脚注尾注 | {{FOOTNOTE_RULE}} |
| 批注 | {{COMMENT_RULE}} |
| 修订记录 | {{REVISION_RULE}} |
| 超链接 | {{HYPERLINK_RULE}} |
| 字段 | {{FIELD_RULE}} |

## 12. 样式与直接格式规则

| 检查项 | 标准规则 |
| --- | --- |
| 样式驱动 | {{STYLE_DRIVEN_RULE}} |
| 直接格式覆盖 | {{DIRECT_FORMAT_RULE}} |
| 样式继承 | {{STYLE_INHERITANCE_RULE}} |
| 默认格式 | {{DOC_DEFAULT_RULE}} |
| 颜色与高亮 | {{COLOR_HIGHLIGHT_RULE}} |

## 13. 自动处理边界

### 可以自动处理

{{AUTO_FIX_SCOPE}}

### 需要用户确认或专项处理

{{MANUAL_SCOPE}}

## 14. 需用户确认的规则缺口

{{RULE_GAPS}}
