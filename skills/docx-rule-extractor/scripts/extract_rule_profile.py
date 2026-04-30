#!/usr/bin/env python3
"""从 DOCX 快照生成第一版规则包草案。"""
from __future__ import annotations

import argparse
import json
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path


def write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def read_template(name: str) -> str:
    template_path = Path(__file__).resolve().parents[1] / "references" / name
    return template_path.read_text(encoding="utf-8")


def render_template(template: str, values: dict[str, str]) -> str:
    output = template
    for key, value in values.items():
        output = output.replace("{{" + key + "}}", str(value))
    return output


def value_or_confirm(value, suffix: str = "") -> str:
    if value is None or value == "":
        return "未能从标准文件稳定提取，需人工确认"
    return f"{value}{suffix}"


def yes_no(value) -> str:
    if value is None:
        return "未能从标准文件稳定提取，需人工确认"
    return "是" if value else "否"


def present_absent(value: bool) -> str:
    return "存在" if value else "未检测到"


def alignment_label(value) -> str:
    labels = {
        "center": "居中",
        "left": "左对齐",
        "right": "右对齐",
        "both": "两端对齐",
        "distribute": "分散对齐",
    }
    return labels.get(value, value_or_confirm(value))


def orientation_label(value) -> str:
    if value is None:
        return "未能从标准文件稳定提取，需人工确认"
    return "横向" if value == "landscape" else "纵向"


def most_common(values):
    filtered = [value for value in values if value not in (None, "")]
    if not filtered:
        return None
    return Counter(filtered).most_common(1)[0][0]


def normalize_font_name(value):
    if value in (None, ""):
        return None
    mapping = {
        "Їбкн_GB2312": "仿宋_GB2312",
        "Їбкн": "仿宋",
        "кнлЕ": "宋体",
    }
    return mapping.get(value, value)


def effective_cjk_font(run: dict):
    east = normalize_font_name(run.get("font_east_asia"))
    ascii_font = normalize_font_name(run.get("font_ascii"))
    return east or ascii_font


def merge_format(paragraph: dict, style_details: dict) -> dict:
    style = paragraph.get("style")
    style_detail = style_details.get(style, {})
    run = dict(style_detail.get("run_format") or {})
    para = dict(style_detail.get("paragraph_format") or {})
    run.update({k: v for k, v in (paragraph.get("run_format") or {}).items() if v not in (None, "")})
    para.update({k: v for k, v in (paragraph.get("paragraph_format") or {}).items() if v not in (None, "")})
    return {"run": run, "paragraph": para}


def representative_formats(snapshot: dict) -> dict:
    style_details = snapshot.get("style_details", {})
    grouped = defaultdict(list)
    for paragraph in snapshot.get("paragraphs", []):
        element_type = paragraph.get("element_type")
        if element_type in {"heading-level-1", "heading-level-2", "heading-level-3", "body-paragraph"}:
            grouped[element_type].append(merge_format(paragraph, style_details))
    for table in snapshot.get("table_details", []):
        for paragraph in table.get("cell_paragraphs", []):
            if paragraph.get("row_index") == 1:
                grouped["table-header"].append(merge_format(paragraph, style_details))
            else:
                grouped["table-body"].append(merge_format(paragraph, style_details))
    result = {}
    for element_type, formats in grouped.items():
        effective_fonts = [effective_cjk_font(item["run"]) for item in formats]
        ascii_fonts = [normalize_font_name(item["run"].get("font_ascii")) or effective_cjk_font(item["run"]) for item in formats]
        result[element_type] = {
            "font_east_asia": most_common(effective_fonts),
            "font_ascii": most_common(ascii_fonts),
            "font_size_pt": most_common([item["run"].get("font_size_pt") for item in formats]),
            "bold": most_common([item["run"].get("bold") for item in formats]),
            "first_line_indent_cm": most_common([item["paragraph"].get("first_line_indent_cm") for item in formats]),
            "line_spacing_multiple": most_common([item["paragraph"].get("line_spacing_multiple") for item in formats]),
            "line_spacing_pt": most_common([item["paragraph"].get("line_spacing_pt") for item in formats]),
            "space_before_pt": most_common([item["paragraph"].get("space_before_pt") for item in formats]),
            "space_after_pt": most_common([item["paragraph"].get("space_after_pt") for item in formats]),
            "alignment": most_common([item["paragraph"].get("alignment") for item in formats]),
        }
    return result


def fmt_rule(formats: dict, key: str) -> dict:
    return formats.get(key, {})


def font_text(rule: dict) -> str:
    east = rule.get("font_east_asia")
    ascii_font = rule.get("font_ascii")
    if east and ascii_font and east != ascii_font:
        return f"中文 {east}；西文 {ascii_font}"
    return value_or_confirm(east or ascii_font)


def spacing_text(rule: dict) -> str:
    multiple = rule.get("line_spacing_multiple")
    line = rule.get("line_spacing_pt")
    before = rule.get("space_before_pt")
    after = rule.get("space_after_pt")
    parts = []
    if multiple is not None:
        parts.append(f"{multiple} 倍行距")
    elif line is not None:
        parts.append(f"固定值约 {line} 磅")
    if before is not None:
        parts.append(f"段前约 {before} 磅")
    if after is not None:
        parts.append(f"段后约 {after} 磅")
    return "；".join(parts) if parts else "未能从标准文件稳定提取，需人工确认"


def char_effects_text(rule: dict) -> str:
    parts = []
    for key, label in [
        ("bold", "加粗"),
        ("italic", "斜体"),
        ("underline", "下划线"),
        ("strike", "删除线"),
        ("small_caps", "小型大写"),
        ("all_caps", "全部大写"),
        ("vertical_align", "上标/下标"),
        ("color", "字体颜色"),
        ("highlight", "高亮"),
    ]:
        value = rule.get(key)
        if value in (None, "", False):
            continue
        if value is True:
            parts.append(label)
        else:
            parts.append(f"{label}：{value}")
    return "；".join(parts) if parts else "无特殊字符效果或未能稳定提取"


def indent_text(rule: dict) -> str:
    parts = []
    for key, label in [
        ("first_line_indent_cm", "首行缩进"),
        ("left_indent_cm", "左缩进"),
        ("right_indent_cm", "右缩进"),
        ("hanging_indent_cm", "悬挂缩进"),
    ]:
        value = rule.get(key)
        if value is not None:
            parts.append(f"{label}约 {value} 厘米")
    return "；".join(parts) if parts else "未能从标准文件稳定提取，需人工确认"


def paragraph_layout_text(rule: dict) -> str:
    return f"{indent_text(rule)}；{alignment_label(rule.get('alignment'))}"


def page_control_text(rule: dict) -> str:
    parts = []
    for key, label in [
        ("keep_next", "与下段同页"),
        ("keep_lines", "段中不分页"),
        ("page_break_before", "段前分页"),
        ("widow_control", "孤行控制"),
    ]:
        value = rule.get(key)
        if value is not None:
            parts.append(f"{label}：{yes_no(value)}")
    outline = rule.get("outline_level")
    if outline is not None:
        parts.append(f"大纲级别：{outline}")
    return "；".join(parts) if parts else "未能从标准文件稳定提取，需人工确认"


def margin_text(section: dict | None) -> str:
    if not section:
        return "未能从标准文件稳定提取，需人工确认"
    return (
        f"上 {value_or_confirm(section.get('margin_top_cm'), ' 厘米')}，"
        f"下 {value_or_confirm(section.get('margin_bottom_cm'), ' 厘米')}，"
        f"左 {value_or_confirm(section.get('margin_left_cm'), ' 厘米')}，"
        f"右 {value_or_confirm(section.get('margin_right_cm'), ' 厘米')}"
    )


def table_detail_summary(snapshot: dict) -> dict:
    details = snapshot.get("table_details", [])
    if not details:
        return {
            "width": "未检测到表格",
            "borders": "未检测到表格",
            "headers": "未检测到表格",
            "rows_columns": "未检测到表格",
            "cells": "未检测到表格",
            "merged": "未检测到表格",
            "nested": "需结合表格结构专项审计",
        }
    row_counts = [item.get("row_count", 0) for item in details]
    col_counts = [item.get("grid_column_count", 0) for item in details]
    merged = sum(item.get("merged_cell_count", 0) for item in details)
    header_rows = sum(item.get("header_row_count", 0) for item in details)
    bordered = sum(1 for item in details if item.get("has_borders"))
    return {
        "width": "需审计表格宽度、列宽和页面方向是否匹配；第一版不自动重排复杂表格。",
        "borders": f"{bordered} 个表格检测到边框定义；缺失或异常边框进入人工确认。",
        "headers": f"检测到 {header_rows} 个重复表头行标记；表头加粗、居中、底纹需继续审计。",
        "rows_columns": f"行数范围 {min(row_counts)}-{max(row_counts)}；列数范围 {min(col_counts)}-{max(col_counts)}。",
        "cells": "需审计单元格内边距、垂直对齐、底纹和单元格内段落格式。",
        "merged": f"检测到 {merged} 个合并相关单元格标记；默认进入人工确认。",
        "nested": "嵌套表格需专项审计，不自动重排。",
    }


def most_common_table_value(details: list[dict], key: str, default: str) -> str:
    values = [item.get(key) for item in details if item.get(key) not in (None, "")]
    if not values:
        return default
    return str(Counter(values).most_common(1)[0][0])


def first_border_value(details: list[dict], side: str, attr_name: str, default: str) -> str:
    values = []
    for item in details:
        border = (item.get("border_values") or {}).get(side) or {}
        value = border.get(attr_name)
        if value not in (None, ""):
            values.append(value)
    if not values:
        return default
    return str(Counter(values).most_common(1)[0][0])


def first_shading_fill(details: list[dict], default: str) -> str:
    fills = []
    for item in details:
        fills.extend(fill for fill in item.get("first_row_shading_fills", []) if fill)
    if not fills:
        return default
    return str(Counter(fills).most_common(1)[0][0])


def table_rule_config(snapshot: dict) -> dict:
    details = snapshot.get("table_details", [])
    width = most_common_table_value(details, "width", "5000")
    width_type = most_common_table_value(details, "width_type", "pct")
    if width in {"0", "nil", "auto"}:
        width = "5000"
        width_type = "pct"
    return {
        "width": width,
        "width_type": width_type,
        "alignment": most_common_table_value(details, "alignment", "center"),
        "layout_type": most_common_table_value(details, "layout_type", "fixed"),
        "border_val": first_border_value(details, "top", "val", "single"),
        "border_sz": first_border_value(details, "top", "sz", "4"),
        "border_color": first_border_value(details, "top", "color", "auto"),
        "border_space": first_border_value(details, "top", "space", "0"),
        "cell_margin_top": most_common_table_value(details, "cell_margin_top", "60"),
        "cell_margin_bottom": most_common_table_value(details, "cell_margin_bottom", "60"),
        "cell_margin_left": most_common_table_value(details, "cell_margin_left", "108"),
        "cell_margin_right": most_common_table_value(details, "cell_margin_right", "108"),
        "row_height": "360",
        "row_height_rule": "atLeast",
        "vertical_alignment": "center",
        "header_repeat": "true",
        "header_shading_fill": first_shading_fill(details, "D9EAF7"),
    }


def extract_or_gap(label: str, value: str, gaps: list[str]) -> str:
    if "未能从标准文件稳定提取" in value:
        gaps.append(label)
    return value


def yes_no_text(value: bool) -> str:
    return "是" if value else "否"


def yaml_value(value) -> str:
    if value in (None, ""):
        return "null"
    if isinstance(value, bool):
        return str(value).lower()
    return str(value)


def style_entry(role: str, word_style_id: str, word_style_name: str, description: str, outline_level, rule: dict) -> str:
    return f"""  {role}:
    description: {description}
    word_style_id: {word_style_id}
    word_style_name: {word_style_name}
    localized_style_candidates:
      - {word_style_name}
    outline_level: {yaml_value(outline_level)}
    write_strategy: style-definition
    format:
      font_east_asia: {yaml_value(rule.get("font_east_asia"))}
      font_ascii: {yaml_value(rule.get("font_ascii") or rule.get("font_east_asia"))}
      font_size_pt: {yaml_value(rule.get("font_size_pt"))}
      bold: {yaml_value(rule.get("bold"))}
      line_spacing_multiple: {yaml_value(rule.get("line_spacing_multiple"))}
      line_spacing_pt: {yaml_value(rule.get("line_spacing_pt"))}
      space_before_pt: {yaml_value(rule.get("space_before_pt"))}
      space_after_pt: {yaml_value(rule.get("space_after_pt"))}
      first_line_indent_cm: {yaml_value(rule.get("first_line_indent_cm"))}
      alignment: {yaml_value(rule.get("alignment"))}
"""


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--snapshot", required=True)
    parser.add_argument("--rule-dir", required=True)
    parser.add_argument("--rule-id", required=True)
    parser.add_argument("--description", required=True)
    parser.add_argument("--standard-docx", required=True)
    args = parser.parse_args()

    snapshot = json.loads(Path(args.snapshot).read_text(encoding="utf-8-sig"))
    rule_dir = Path(args.rule_dir)
    today = date.today().isoformat()
    formats = representative_formats(snapshot)
    body = fmt_rule(formats, "body-paragraph")
    table_body = fmt_rule(formats, "table-body") or dict(body)
    table_header = fmt_rule(formats, "table-header") or dict(table_body)
    h1 = fmt_rule(formats, "heading-level-1")
    h2 = fmt_rule(formats, "heading-level-2")
    h3 = fmt_rule(formats, "heading-level-3")
    first_section = (snapshot.get("sections") or [None])[0]
    orientations = {section.get("orientation") for section in snapshot.get("sections", [])}
    orientation = "包含横向与纵向页面" if len(orientations) > 1 else orientation_label(first_section.get("orientation") if first_section else None)
    package_summary = snapshot.get("package_summary", {})
    settings = snapshot.get("settings", {})
    numbering_summary = snapshot.get("numbering_summary", {})
    field_summary = snapshot.get("field_summary", {})
    table_rules = table_detail_summary(snapshot)
    table_config = table_rule_config(snapshot)
    gaps = []
    paper_rule = extract_or_gap(
        "纸张大小",
        f"宽 {value_or_confirm(first_section.get('width_cm') if first_section else None, ' 厘米')}，高 {value_or_confirm(first_section.get('height_cm') if first_section else None, ' 厘米')}",
        gaps,
    )
    margin_rule = extract_or_gap("页边距", margin_text(first_section), gaps)
    body_font = extract_or_gap("正文字体", font_text(body), gaps)
    body_size = extract_or_gap("正文字号", value_or_confirm(body.get("font_size_pt"), " 磅"), gaps)
    body_indent = extract_or_gap("正文首行缩进", value_or_confirm(body.get("first_line_indent_cm"), " 厘米"), gaps)
    body_spacing = extract_or_gap("正文行间距与段落间距", spacing_text(body), gaps)
    h1_font = extract_or_gap("一级标题字体", font_text(h1), gaps)
    h2_font = extract_or_gap("二级标题字体", font_text(h2), gaps)
    h3_font = extract_or_gap("三级标题字体", font_text(h3), gaps)
    h1_size = extract_or_gap("一级标题字号", value_or_confirm(h1.get("font_size_pt"), " 磅"), gaps)
    h2_size = extract_or_gap("二级标题字号", value_or_confirm(h2.get("font_size_pt"), " 磅"), gaps)
    h3_size = extract_or_gap("三级标题字号", value_or_confirm(h3.get("font_size_pt"), " 磅"), gaps)
    rule_gaps = "\n".join(f"- {item}：需用户结合标准文件人工确认。" for item in sorted(set(gaps))) if gaps else "无。"
    body_line_spacing_multiple = body.get("line_spacing_multiple")
    body_line_spacing_pt = None if body_line_spacing_multiple is not None else body.get("line_spacing_pt")
    table_line_spacing_multiple = table_body.get("line_spacing_multiple")
    table_line_spacing_pt = None if table_line_spacing_multiple is not None else table_body.get("line_spacing_pt")
    table_header_line_spacing_multiple = table_header.get("line_spacing_multiple")
    table_header_line_spacing_pt = None if table_header_line_spacing_multiple is not None else table_header.get("line_spacing_pt")
    has_tables = bool(snapshot.get("table_details"))
    row_height_mode = "audit-only" if has_tables else "skip"
    border_mode = "audit-only" if has_tables else "skip"

    write(rule_dir / "profile.yaml", f"""id: {args.rule_id}
rule_schema_version: 2.0.0
name: {args.rule_id}
description: {args.description}
version: 1.0.0
status: draft
based_on:
  - {args.standard_docx}
document_types:
  - 正式材料
features:
  auto_toc: true
  style_driven: true
  landscape_sections: true
  appendix_tables: true
toc:
  levels: 3
risk_level: medium
change_summary: 基于标准 Word 文件抽取的初始规则草案
last_updated: {today}
""")
    write(rule_dir / "role-map.yaml", """roles:
  cover-title:
    description: 封面总题名
  cover-subtitle:
    description: 封面副题名
  cover-version:
    description: 封面版本号
  cover-org:
    description: 封面编制单位或落款
  cover-date:
    description: 封面日期
  toc-title:
    description: 目录标题
  toc-field:
    description: Word 自动目录
  toc-static-item:
    description: 静态目录条目
  heading-level-1:
    description: 一级标题
  heading-level-2:
    description: 二级标题
  heading-level-3:
    description: 三级标题
  body-paragraph:
    description: 正文段落
  body-no-indent:
    description: 无首行缩进正文
  table-header:
    description: 表头
  table-body:
    description: 表格正文
""")
    write(rule_dir / "style-map.yaml", f"""styles:
{style_entry("heading-level-1", "Heading1", "Heading 1", "一级标题", 1, h1)}
{style_entry("heading-level-2", "Heading2", "Heading 2", "二级标题", 2, h2)}
{style_entry("heading-level-3", "Heading3", "Heading 3", "三级标题", 3, h3)}
{style_entry("body-paragraph", "Normal", "Normal", "正文段落", None, body)}
""")
    write(rule_dir / "element-rules.yaml", """element_types:
  - cover-title
  - cover-subtitle
  - cover-version
  - cover-org
  - cover-date
  - toc-title
  - toc-field
  - toc-static-item
  - heading-level-1
  - heading-level-2
  - heading-level-3
  - body-paragraph
  - normal-table
  - special-panel
  - appendix-table
  - page-section
""")
    write(rule_dir / "toc-rules.yaml", """auto_toc:
  required: true
  replace_static_toc: true
  require_toc_content_audit: true
  include_outline_levels: [1, 2, 3]
  full_toc_by_rule: true
""")
    write(rule_dir / "table-rules.yaml", f"""tables:
  audit: true
  safe_auto_fix:
    - table_width_alignment
    - table_cell_margins
    - table_vertical_alignment
    - table_cell_text_format
  header:
    font_east_asia: {yaml_value(table_header.get("font_east_asia") or table_body.get("font_east_asia"))}
    font_ascii: {yaml_value(table_header.get("font_ascii") or table_header.get("font_east_asia") or table_body.get("font_ascii"))}
    font_size_pt: {yaml_value(table_header.get("font_size_pt") if table_header.get("font_size_pt") is not None else table_body.get("font_size_pt"))}
    bold: {yaml_value(table_header.get("bold"))}
    line_spacing_multiple: {yaml_value(table_header_line_spacing_multiple)}
    line_spacing_pt: {yaml_value(table_header_line_spacing_pt)}
    alignment: center
    shading_fill: {table_config["header_shading_fill"]}
  body:
    font_east_asia: {yaml_value(table_body.get("font_east_asia"))}
    font_ascii: {yaml_value(table_body.get("font_ascii") or table_body.get("font_east_asia"))}
    font_size_pt: {yaml_value(table_body.get("font_size_pt"))}
    bold: {yaml_value(table_body.get("bold"))}
    line_spacing_multiple: {yaml_value(table_line_spacing_multiple)}
    line_spacing_pt: {yaml_value(table_line_spacing_pt)}
  layout:
    width: {table_config["width"]}
    width_type: {table_config["width_type"]}
    alignment: {table_config["alignment"]}
    layout_type: {table_config["layout_type"]}
    cell_margin_top: {table_config["cell_margin_top"]}
    cell_margin_bottom: {table_config["cell_margin_bottom"]}
    cell_margin_left: {table_config["cell_margin_left"]}
    cell_margin_right: {table_config["cell_margin_right"]}
    vertical_alignment: {table_config["vertical_alignment"]}
    header_repeat: {table_config["header_repeat"]}
  row_height:
    mode: {row_height_mode}
    value: null
    rule: null
    confirmed_by: null
    confirmed_at: null
  border:
    mode: {border_mode}
    top: {{ style: {table_config["border_val"]}, width_twips: {table_config["border_sz"]}, color: "{table_config["border_color"]}" }}
    bottom: {{ style: {table_config["border_val"]}, width_twips: {table_config["border_sz"]}, color: "{table_config["border_color"]}" }}
    left: {{ style: {table_config["border_val"]}, width_twips: {table_config["border_sz"]}, color: "{table_config["border_color"]}" }}
    right: {{ style: {table_config["border_val"]}, width_twips: {table_config["border_sz"]}, color: "{table_config["border_color"]}" }}
  manual_review:
    - merged_cells
    - cross_page_tables
    - landscape_section_tables
    - nested_tables
""")
    write(rule_dir / "page-rules.yaml", f"""page:
  section_count_in_standard: {snapshot.get('section_count', 0)}
  preserve_headers_footers: true
  landscape_sections: audit_and_preserve
""")
    write(rule_dir / "risk-policy.yaml", """auto_fix:
  - map_heading_native_style
  - apply_body_style_definition
  - apply_table_cell_format
  - insert_or_replace_toc_field_after_content_audit
manual_review:
  - toc_items_not_matching_headings
  - ambiguous_numbered_paragraphs
  - complex_tables
  - unconfirmed_row_height
  - unconfirmed_table_border
  - headers_footers
blocked:
  - unreadable_docx
  - missing_rule_profile
""")
    template = read_template("RULE_SUMMARY_TEMPLATE.md")
    write(
        rule_dir / "RULE_SUMMARY.md",
        render_template(
            template,
            {
                "RULE_NAME": args.description,
                "STANDARD_DOCX": args.standard_docx,
                "DOCUMENT_TYPE": "正式材料",
                "RULE_STATUS": "待用户确认",
                "SECTION_SUMMARY": f"共 {snapshot.get('section_count', 0)} 个页面节。",
                "PARAGRAPH_SUMMARY": f"共 {snapshot.get('non_empty_paragraph_count', 0)} 个非空段落。",
                "TABLE_SUMMARY": f"共 {snapshot.get('table_count', 0)} 个表格。",
                "TOC_SUMMARY": "已检测到自动目录。" if snapshot.get("has_toc_field") else "未检测到自动目录，需按规则生成。",
                "HEADER_FOOTER_SUMMARY": f"检测到 {package_summary.get('header_part_count', 0)} 个页眉部件、{package_summary.get('footer_part_count', 0)} 个页脚部件；涉及业务信息时不自动改写。",
                "MEDIA_SUMMARY": f"检测到 {package_summary.get('media_count', 0)} 个媒体文件。",
                "NUMBERING_SUMMARY": f"编号定义 {numbering_summary.get('abstract_numbering_count', 0)} 套，编号实例 {numbering_summary.get('numbering_instance_count', 0)} 个，编号层级 {numbering_summary.get('levels_count', 0)} 个。",
                "ANNOTATION_SUMMARY": f"脚注：{present_absent(package_summary.get('footnotes_present', False))}；尾注：{present_absent(package_summary.get('endnotes_present', False))}；批注：{present_absent(package_summary.get('comments_present', False))}；修订标记 {snapshot.get('revision_count', 0)} 处。",
                "PAPER_RULE": paper_rule,
                "ORIENTATION_RULE": orientation,
                "MARGIN_RULE": margin_rule,
                "HEADER_DISTANCE_RULE": value_or_confirm(first_section.get("header_cm") if first_section else None, " 厘米"),
                "FOOTER_DISTANCE_RULE": value_or_confirm(first_section.get("footer_cm") if first_section else None, " 厘米"),
                "GUTTER_RULE": value_or_confirm(first_section.get("gutter_cm") if first_section else None, " 厘米"),
                "SECTION_BREAK_RULE": "保留标准文件分节；横向页、附件、附表不得合并到普通正文节。",
                "COLUMN_RULE": f"分栏数量：{value_or_confirm(first_section.get('column_count') if first_section else None)}；栏间距：{value_or_confirm(first_section.get('column_space_cm') if first_section else None, ' 厘米')}",
                "DOC_GRID_RULE": value_or_confirm(first_section.get("doc_grid_type") if first_section else None),
                "TEXT_DIRECTION_RULE": value_or_confirm(first_section.get("text_direction") if first_section else None),
                "PAGE_VERTICAL_ALIGN_RULE": value_or_confirm(first_section.get("vertical_alignment") if first_section else None),
                "FIRST_PAGE_HEADER_FOOTER_RULE": f"首页专用页眉页脚：{yes_no(first_section.get('different_first_page') if first_section else None)}；第一版不自动重建。",
                "ODD_EVEN_HEADER_FOOTER_RULE": f"奇偶页不同页眉页脚：{yes_no(settings.get('even_and_odd_headers'))}。",
                "HEADER_RULE": f"普通页页眉引用类型：{', '.join(first_section.get('header_reference_types') or []) if first_section else '未检测到'}。",
                "FOOTER_RULE": f"普通页页脚引用类型：{', '.join(first_section.get('footer_reference_types') or []) if first_section else '未检测到'}。",
                "HEADER_FOOTER_LINK_RULE": "分节后的页眉页脚继承关系需审计；不自动断开或重连。",
                "PAGE_NUMBER_FIELD_RULE": f"检测到 {field_summary.get('page_field_count', 0)} 个页码字段；页码必须可更新。",
                "COVER_TITLE_RULE": "按标准文件的封面主标题视觉效果执行；字体、字号、位置提取不稳定时人工确认。",
                "SUBTITLE_RULE": "按标准文件副标题视觉效果执行；无法稳定提取时人工确认。",
                "ORG_NAME_RULE": "按标准文件编制单位位置和格式执行；无法稳定提取时人工确认。",
                "DATE_RULE": "按标准文件日期位置和格式执行；无法稳定提取时人工确认。",
                "COVER_LAYOUT_RULE": "保持标准文件封面空行、居中和版面层次。",
                "TOC_RULE": "最终文档必须使用 Word 可自动更新目录。",
                "TOC_SCOPE": "一级标题、二级标题、三级标题。",
                "TOC_LEVEL_RULE": "按一级、二级、三级标题层级显示。",
                "TOC_PAGE_NUMBER_RULE": "显示页码，且页码应可随正文分页更新。",
                "TOC_HYPERLINK_RULE": "建议保留目录超链接。",
                "STATIC_TOC_RULE": "输入文档只有手工目录时，应替换为自动目录。",
                "H1_FONT": h1_font,
                "H1_SIZE": h1_size,
                "H1_CHAR_EFFECTS": char_effects_text(h1),
                "H1_SPACING": spacing_text(h1),
                "H1_PARAGRAPH_LAYOUT": paragraph_layout_text(h1),
                "H1_PAGE_CONTROL": page_control_text(h1),
                "H2_FONT": h2_font,
                "H2_SIZE": h2_size,
                "H2_CHAR_EFFECTS": char_effects_text(h2),
                "H2_SPACING": spacing_text(h2),
                "H2_PARAGRAPH_LAYOUT": paragraph_layout_text(h2),
                "H2_PAGE_CONTROL": page_control_text(h2),
                "H3_FONT": h3_font,
                "H3_SIZE": h3_size,
                "H3_CHAR_EFFECTS": char_effects_text(h3),
                "H3_SPACING": spacing_text(h3),
                "H3_PARAGRAPH_LAYOUT": paragraph_layout_text(h3),
                "H3_PAGE_CONTROL": page_control_text(h3),
                "HEADING_NUMBERING_RULE": "标题编号需与正文结构一致；编号不连续时进入人工确认。",
                "HEADING_KEEP_RULE": "标题应尽量与其后正文保持连续，不应孤立在页底。",
                "HEADING_OUTLINE_RULE": "标题必须具备正确层级，以支持导航窗格和自动目录。",
                "BODY_FONT": body_font,
                "BODY_SIZE": body_size,
                "BODY_CHAR_EFFECTS": char_effects_text(body),
                "BODY_INDENT": indent_text(body),
                "BODY_SPACING": body_spacing,
                "BODY_ALIGNMENT": alignment_label(body.get("alignment")),
                "BODY_TAB_RULE": f"检测到 {value_or_confirm(body.get('tab_count'))} 个制表位配置。",
                "BODY_PAGE_CONTROL": page_control_text(body),
                "BODY_SPECIAL_RULE": "段内强调、超链接和脚注引用应保留，不能因正文统一而丢失。",
                "TABLE_COUNT": snapshot.get("table_count", 0),
                "TABLE_WIDTH_RULE": table_rules["width"],
                "TABLE_BORDER_RULE": table_rules["borders"],
                "TABLE_HEADER_RULE": f"按标准文件表头抽取：{font_text(table_header)}，{value_or_confirm(table_header.get('font_size_pt'), ' 磅')}，{char_effects_text(table_header)}；{spacing_text(table_header)}；同时参考 Word 重复表头标记：{table_rules['headers']}",
                "TABLE_BODY_RULE": f"按标准文件表格正文抽取：{font_text(table_body)}，{value_or_confirm(table_body.get('font_size_pt'), ' 磅')}，{spacing_text(table_body)}。",
                "TABLE_ROW_COLUMN_RULE": table_rules["rows_columns"],
                "TABLE_CELL_RULE": table_rules["cells"],
                "MERGED_CELL_RULE": table_rules["merged"],
                "CROSS_PAGE_TABLE_RULE": "跨页表格默认进入人工确认。",
                "LANDSCAPE_TABLE_RULE": "横向页面表格默认进入人工确认。",
                "NESTED_TABLE_RULE": table_rules["nested"],
                "HEADING_NUMBER_RULE": "章节编号需与标题层级对应。",
                "BODY_NUMBER_RULE": "正文编号需保持缩进和编号连续。",
                "BULLET_RULE": "项目符号需保持符号、缩进和层级一致。",
                "MULTILEVEL_LIST_RULE": "多级列表需按层级审计；不自动重建复杂编号。",
                "NUMBER_CONTINUITY_RULE": "编号断裂、跳号、重号进入人工确认。",
                "IMAGE_RULE": f"检测到 {package_summary.get('media_count', 0)} 个媒体文件；图片位置、大小、环绕方式第一版仅审计。",
                "CAPTION_RULE": "图题、表题需与正文层级和编号规则一致。",
                "FOOTNOTE_RULE": f"脚注：{present_absent(package_summary.get('footnotes_present', False))}；尾注：{present_absent(package_summary.get('endnotes_present', False))}；第一版不自动重写。",
                "COMMENT_RULE": f"批注：{present_absent(package_summary.get('comments_present', False))}；第一版保留，不自动删除。",
                "REVISION_RULE": f"修订跟踪设置：{present_absent(settings.get('track_revisions', False))}；修订标记 {snapshot.get('revision_count', 0)} 处；第一版不自动接受或拒绝。",
                "HYPERLINK_RULE": f"检测到 {snapshot.get('hyperlink_count', 0)} 处超链接；应保留链接目标和显示文字。",
                "FIELD_RULE": f"检测到 {field_summary.get('field_instruction_count', 0)} 个字段指令；目录、页码等字段应保持可更新。",
                "STYLE_DRIVEN_RULE": "最终格式应优先由标准样式驱动，避免手工直接格式覆盖成为主要格式来源。",
                "DIRECT_FORMAT_RULE": "直接格式覆盖需审计；可确定的正文和标题覆盖可迁移为标准规则，段内强调需保留。",
                "STYLE_INHERITANCE_RULE": "样式继承链需保留；基准样式、后续段落样式和直接格式共同决定最终显示。",
                "DOC_DEFAULT_RULE": "文档默认字体、默认段落和主题字体需审计；第一版先记录不强制重建。",
                "COLOR_HIGHLIGHT_RULE": "字体颜色、高亮、下划线、删除线、上标下标等字符效果需保留原意。",
                "AUTO_FIX_SCOPE": "- 明确识别的正文段落可统一为正文规则。\n- 明确识别的章节标题可统一为标题规则并纳入自动目录。\n- 手工目录可替换为 Word 可自动更新目录。",
                "MANUAL_SCOPE": "- 原目录项与正文标题不一致。\n- 疑似标题但编号不连续。\n- 复杂表格、页眉页脚、脚注、批注、修订记录。\n- 标准文件未能稳定提取的字体、字号、行距、页边距等项目。",
                "RULE_GAPS": rule_gaps,
                "COVERAGE_STATEMENT": "本规则说明覆盖 Word 正式材料交付中约 90% 以上常见可视格式面：页面、节、页眉页脚、封面、目录、标题、正文、编号、表格、图片、脚注、批注、修订和样式治理。少数低频 OOXML 特性仍列入人工确认或专项处理。",
            },
        ),
    )
    print(rule_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
