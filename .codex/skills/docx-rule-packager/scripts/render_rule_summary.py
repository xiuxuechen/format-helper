"""从语义规则草案或槽位事实生成 RULE_SUMMARY.md。"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[4]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.reporting.human_readable import (
    human_alignment,
    human_font_size,
    human_indent,
    human_line_spacing,
    markdown_list,
    markdown_table,
    render_template,
    safe_markdown_text,
    status_marker,
)
from scripts.utils.simple_yaml import load_yaml
from scripts.validation.human_readable_report import (
    RULE_SUMMARY_REQUIRED_SECTIONS,
    assert_human_readable_report,
)


RESOLVED_STATUSES = {"resolved", "resolved_with_conflicts", "not_applicable", "user_confirmed"}
TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "templates" / "RULE_SUMMARY.template.md"

ROLE_LABELS = {
    "cover-title": "封面标题",
    "cover-subtitle": "封面副标题",
    "cover-meta": "封面信息",
    "revision-table-title": "修订表标题",
    "toc-title": "目录标题",
    "toc-level-1": "一级目录项",
    "toc-level-2": "二级目录项",
    "toc-level-3": "三级目录项",
    "heading-level-1": "一级标题",
    "heading-level-2": "二级标题",
    "heading-level-3": "三级标题",
    "heading-level-4": "四级标题",
    "body-paragraph": "正文段落",
    "list-paragraph": "列表段落",
    "table-content": "表格内容",
    "header-footer": "页眉页脚",
    "section-page-setup": "分节页面设置",
}

SLOT_LABELS = {
    "font_east_asia": "中文字体",
    "font_ascii": "西文字体",
    "font_size_pt": "字号",
    "bold": "加粗",
    "italic": "斜体",
    "underline": "下划线",
    "alignment": "对齐方式",
    "vertical_alignment": "垂直对齐",
    "outline_level": "大纲级别",
    "first_line_indent_cm": "首行缩进",
    "left_indent_cm": "左缩进",
    "right_indent_cm": "右缩进",
    "space_before_pt": "段前间距",
    "space_after_pt": "段后间距",
    "line_spacing_multiple": "行距倍数",
    "style_id": "样式 ID",
    "toc_level": "目录级别",
    "numbering_ref": "编号引用",
    "page_orientation": "页面方向",
    "page_width_twips": "页面宽度",
    "page_height_twips": "页面高度",
    "margin_top_cm": "上页边距",
    "margin_bottom_cm": "下页边距",
    "margin_left_cm": "左页边距",
    "margin_right_cm": "右页边距",
    "header_distance_cm": "页眉距边界",
    "footer_distance_cm": "页脚距边界",
    "page_number_format": "页码格式",
}

CONFIRMED_RULE_TABLE_HEADERS = [
    "文档元素",
    "中文字体",
    "西文字体",
    "字号",
    "加粗",
    "对齐",
    "行距",
    "补充说明",
]


def format_bool(value: Any) -> str:
    """将布尔值渲染为中文。"""
    if value is True:
        return "是"
    if value is False:
        return "否"
    return "未指定"


def alignment_label(value: Any) -> str:
    """渲染对齐方式中文标签。"""
    labels = {
        "left": "左对齐",
        "center": "居中",
        "right": "右对齐",
        "justify": "两端对齐",
        "distributed": "分散对齐",
        "both": "两端对齐",
        "distribute": "分散对齐",
    }
    return labels.get(value, format_value(value))


def strategy_label(value: str) -> str:
    """渲染写入策略中文标签。"""
    labels = {
        "style-definition": "样式定义",
        "direct-format": "直接格式",
        "audit-only": "仅审计",
    }
    return labels.get(value, value)


def format_value(value: Any, suffix: str = "") -> str:
    """渲染格式值。"""
    if value is None or value == "":
        return "未指定"
    return f"{value}{suffix}"


def role_label(role_kind: Any, contract: dict[str, Any] | None = None) -> str:
    """渲染用户可读角色名称。"""
    del contract
    return ROLE_LABELS.get(str(role_kind), "未知角色")


def _trim_slot_description(label: str) -> str:
    """清理槽位说明中的冗余后缀。"""
    trimmed = label.strip()
    if "，单位" in trimmed:
        trimmed = trimmed.split("，单位", 1)[0]
    return trimmed


def slot_label(slot_name: Any, contract: dict[str, Any] | None = None) -> str:
    """渲染用户可读属性名称。"""
    slot_key = str(slot_name)
    if contract:
        registry = contract.get("slot_type_registry", {})
        slot_type = registry.get(slot_key)
        if isinstance(slot_type, dict):
            short_label = slot_type.get("short_label")
            if isinstance(short_label, str) and short_label.strip():
                return _trim_slot_description(short_label)
            description = slot_type.get("description")
            if isinstance(description, str) and description.strip():
                return _trim_slot_description(description)
    return SLOT_LABELS.get(slot_key, "未知属性")


def humanize_text(value: Any, contract: dict[str, Any] | None = None) -> str:
    """清洗用户可见文本中的内部键名。"""
    text = str(value)
    replacements = {
        "required_slots": "必需格式属性",
        "optional_slots": "可选格式属性",
        "slot_name": "属性名称",
        "role_kind": "角色类型",
    }
    for key in sorted(ROLE_LABELS, key=len, reverse=True):
        text = text.replace(key, role_label(key, contract))
    slot_keys = set(SLOT_LABELS)
    if contract:
        slot_keys.update((contract.get("slot_type_registry") or {}).keys())
    for key in sorted(slot_keys, key=len, reverse=True):
        text = text.replace(key, slot_label(key, contract))
    for key, label in replacements.items():
        text = text.replace(key, label)
    return text


def render_format(format_rule: dict[str, Any]) -> str:
    """渲染用户可读格式说明。"""
    parts = []
    if "font_east_asia" in format_rule:
        parts.append(f"字体：{format_value(format_rule.get('font_east_asia'))}")
    if "font_size_pt" in format_rule:
        parts.append(f"字号：{format_value(format_rule.get('font_size_pt'), 'pt')}")
    if "bold" in format_rule:
        parts.append(f"加粗：{format_bool(format_rule.get('bold'))}")
    if "first_line_indent_cm" in format_rule:
        parts.append(f"首行缩进：{format_value(format_rule.get('first_line_indent_cm'), 'cm')}")
    if "line_spacing_multiple" in format_rule:
        parts.append(f"行距：{format_value(format_rule.get('line_spacing_multiple'), ' 倍')}")
    if "outline_level" in format_rule:
        parts.append(f"大纲级别：{format_value(format_rule.get('outline_level'))}")
    if "alignment" in format_rule:
        parts.append(f"对齐：{alignment_label(format_rule.get('alignment'))}")
    return "；".join(parts) if parts else "未指定"


def unwrap_slot_value(value: Any) -> Any:
    """兼容带来源对象和值本身两种槽位值形态。"""
    if isinstance(value, dict) and set(value.keys()) & {"value", "source", "confidence"}:
        return value.get("value")
    return value


def format_slot_value(slot_name: str, value: Any, unit: Any = None) -> str:
    """按槽位类型渲染用户可读值。"""
    raw_value = unwrap_slot_value(value)
    if raw_value is None or raw_value == "":
        return "未指定"
    if slot_name in {"alignment"}:
        return human_alignment(raw_value)
    if slot_name == "vertical_alignment":
        labels = {"top": "顶端", "center": "中部", "bottom": "底端"}
        return labels.get(raw_value, format_value(raw_value))
    if slot_name == "font_size_pt":
        return human_font_size(raw_value)
    if slot_name == "line_spacing_multiple":
        return human_line_spacing(raw_value, unit=unit or "multiple")
    if slot_name in {"first_line_indent_cm", "left_indent_cm", "right_indent_cm"}:
        return human_indent(raw_value, kind=slot_name)
    if slot_name in {"space_before_pt", "space_after_pt"}:
        return f"约 {raw_value} 磅"
    if slot_name == "outline_level":
        return f"{raw_value} 级"
    if slot_name == "toc_level":
        return f"{raw_value} 级"
    if slot_name == "page_orientation":
        return {"portrait": "纵向", "landscape": "横向"}.get(str(raw_value), str(raw_value))
    if slot_name in {"page_width_twips", "page_height_twips"}:
        try:
            cm_value = round(float(raw_value) / 1440 * 2.54, 2)
            return f"约 {cm_value} 厘米"
        except (TypeError, ValueError):
            return str(raw_value)
    if slot_name in {"margin_top_cm", "margin_bottom_cm", "margin_left_cm", "margin_right_cm", "header_distance_cm", "footer_distance_cm"}:
        return f"约 {raw_value} 厘米"
    return str(raw_value)


def load_contract(path: Path | None) -> dict[str, Any] | None:
    """读取 JSON/YAML 契约文件。"""
    if path is None:
        return None
    if path.suffix.lower() == ".json":
        return json.loads(path.read_text(encoding="utf-8"))
    return load_yaml(path)


def role_required_slots(role: dict[str, Any], contract: dict[str, Any] | None) -> list[str]:
    """从契约获取必需槽位。"""
    role_kind = role.get("role_kind", "")
    contracts = (contract or {}).get("role_slot_contracts", {})
    if role_kind in contracts:
        return list(contracts[role_kind].get("required_slots", []))
    return []


def role_optional_slots(role: dict[str, Any], contract: dict[str, Any] | None) -> list[str]:
    """从契约获取可选槽位。"""
    role_kind = role.get("role_kind", "")
    contracts = (contract or {}).get("role_slot_contracts", {})
    if role_kind in contracts:
        return list(contracts[role_kind].get("optional_slots", []))
    return []


def slot_confidence(summary: dict[str, Any]) -> float:
    """推导槽位置信度。"""
    if isinstance(summary.get("confidence"), (int, float)):
        return float(summary["confidence"])
    coverage = summary.get("mode_coverage")
    if isinstance(coverage, (int, float)):
        return float(coverage)
    return 0.0


def format_required_slot_cells(role: dict[str, Any], required_slots: list[str], contract: dict[str, Any] | None) -> str:
    """渲染一个角色的必需属性摘要。"""
    slot_summary = role.get("slot_summary", {})
    cells = []
    for slot_name in required_slots:
        summary = slot_summary.get(slot_name, {})
        cells.append(
            "{label}：{value}".format(
                label=slot_label(slot_name, contract),
                value=format_slot_value(slot_name, summary.get("mode_value"), summary.get("unit")),
            )
        )
    return "；".join(cells) if cells else "无必需属性"


def get_slot_summary(role: dict[str, Any], slot_name: str) -> dict[str, Any]:
    """读取角色的单个槽位摘要。"""
    slot_summary = role.get("slot_summary", {})
    value = slot_summary.get(slot_name)
    return value if isinstance(value, dict) else {}


def get_slot_display_value(role: dict[str, Any], slot_name: str, contract: dict[str, Any] | None) -> str:
    """读取槽位的用户可读值。"""
    del contract
    summary = get_slot_summary(role, slot_name)
    if not summary:
        return ""
    return format_slot_value(slot_name, summary.get("mode_value"), summary.get("unit"))


def get_slot_bool_text(role: dict[str, Any], slot_name: str) -> str:
    """读取布尔槽位的中文值。"""
    summary = get_slot_summary(role, slot_name)
    if not summary:
        return ""
    value = unwrap_slot_value(summary.get("mode_value"))
    if value in (None, ""):
        return ""
    return format_bool(value)


def extra_slot_note(role: dict[str, Any], slot_name: str, contract: dict[str, Any] | None) -> str:
    """将非主展示槽位转换为补充说明。"""
    value = get_slot_display_value(role, slot_name, contract)
    if not value or value == "未指定":
        return ""
    return f"{slot_label(slot_name, contract)}：{value}"


def build_confirmed_rule_row(role: dict[str, Any], contract: dict[str, Any] | None) -> list[str]:
    """构建“已确定的格式规则”总表的一行。"""
    role_name = role_label(role.get("role_kind"), contract)
    if role_name == "页眉页脚":
        return [role_name, "不纳入规则", "不纳入规则", "不纳入规则", "不纳入规则", "不纳入规则", "不纳入规则", "不纳入规则"]
    notes: list[str] = []
    for slot_name in (
        "vertical_alignment",
        "first_line_indent_cm",
        "left_indent_cm",
        "right_indent_cm",
        "space_before_pt",
        "space_after_pt",
        "outline_level",
        "toc_level",
        "numbering_ref",
        "page_orientation",
        "page_width_twips",
        "page_height_twips",
        "margin_top_cm",
        "margin_bottom_cm",
        "margin_left_cm",
        "margin_right_cm",
        "header_distance_cm",
        "footer_distance_cm",
        "page_number_format",
    ):
        note = extra_slot_note(role, slot_name, contract)
        if note:
            notes.append(note)

    alignment = get_slot_display_value(role, "alignment", contract)
    vertical_alignment = get_slot_display_value(role, "vertical_alignment", contract)
    if alignment and vertical_alignment:
        alignment = f"{alignment} / 垂直{vertical_alignment}"

    note_text = "；".join(notes) if notes else "无"
    return [
        role_name,
        get_slot_display_value(role, "font_east_asia", contract) or "未指定",
        get_slot_display_value(role, "font_ascii", contract) or "未指定",
        get_slot_display_value(role, "font_size_pt", contract) or "未指定",
        get_slot_bool_text(role, "bold") or "未指定",
        alignment or "未指定",
        get_slot_display_value(role, "line_spacing_multiple", contract) or "未指定",
        note_text,
    ]


def format_evidence_heading(role: dict[str, Any], sample: dict[str, Any], contract: dict[str, Any] | None) -> str:
    """渲染证据标题。"""
    fact_id = sample.get("fact_id", "unknown")
    return f"### {role_label(role.get('role_kind'), contract)}（事实 `{fact_id}`）"


def format_histogram(histogram: Any, slot_name: str, unit: Any = None) -> str:
    """渲染冲突分布。"""
    if not histogram:
        return "无 histogram。"
    parts = []
    for item in histogram:
        if isinstance(item, dict):
            value = item.get("value")
            count = item.get("count", item.get("sample_count", item.get("frequency", 0)))
            ratio = item.get("ratio", item.get("coverage"))
            suffix = f"，占比 {float(ratio):.2f}" if isinstance(ratio, (int, float)) else ""
            parts.append(f"{format_slot_value(slot_name, value, unit)}：{count}{suffix}")
        else:
            parts.append(str(item))
    return "；".join(parts)


def format_histogram_lines(histogram: Any, slot_name: str, unit: Any = None) -> list[str]:
    """将候选值分布渲染为多行列表。"""
    if not histogram:
        return ["- 暂无候选值分布"]
    lines: list[str] = []
    for item in histogram:
        if not isinstance(item, dict):
            lines.append(f"- {safe_markdown_text(item, max_length=120)}")
            continue
        value = format_slot_value(slot_name, item.get("value"), unit)
        count = item.get("count", item.get("sample_count", item.get("frequency", 0)))
        ratio = item.get("ratio", item.get("coverage"))
        if isinstance(ratio, (int, float)):
            percent = round(float(ratio) * 100)
            lines.append(f"- {value}：{count} 个样本，占比 {percent}%")
        else:
            lines.append(f"- {value}：{count} 个样本")
    return lines


def summarize_role_result(role: dict[str, Any], contract: dict[str, Any] | None) -> str:
    """将角色核心格式摘要为一条短句。"""
    parts: list[str] = []
    font = get_slot_display_value(role, "font_east_asia", contract)
    size = get_slot_display_value(role, "font_size_pt", contract)
    spacing = get_slot_display_value(role, "line_spacing_multiple", contract)
    alignment = get_slot_display_value(role, "alignment", contract)
    if font and font != "未指定":
        parts.append(font)
    if size and size != "未指定":
        parts.append(size)
    if get_slot_bool_text(role, "bold") == "是":
        parts.append("加粗")
    if spacing and spacing != "未指定":
        parts.append(spacing)
    if alignment and alignment != "未指定":
        parts.append(alignment)
    vertical_alignment = get_slot_display_value(role, "vertical_alignment", contract)
    if vertical_alignment and vertical_alignment != "未指定":
        parts.append(f"垂直{vertical_alignment}")
    return "，".join(parts) if parts else "未提取到核心格式属性"


def locator_summary(sample: dict[str, Any]) -> str:
    """将样本定位信息转换为用户可读位置。"""
    locator = sample.get("locator") or {}
    parts: list[str] = []
    if "page" in locator:
        parts.append(f"第 {locator['page']} 页")
    if "paragraph_index" in locator:
        parts.append(f"第 {locator['paragraph_index']} 段")
    if "table_index" in locator:
        parts.append(f"第 {locator['table_index']} 个表格")
    if "cell_index" in locator:
        parts.append(f"第 {locator['cell_index']} 个单元格")
    if not parts and sample.get("fact_kind") == "table_cell":
        parts.append("表格单元格")
    if not parts:
        parts.append("样本位置未标注")
    return "，".join(parts)


def build_blocking_items_section(
    conflict_rows: list[dict[str, Any]],
    gate_blockers: list[dict[str, Any]],
    contract: dict[str, Any] | None,
) -> str:
    """构建阻断项说明。"""
    sections: list[str] = []
    for index, row in enumerate([item for item in conflict_rows if item["severity"] == "blocking"], start=1):
        role = row["role"]
        summary = row["summary"]
        heading = f"### {index}. {role_label(role.get('role_kind'), contract)} · {slot_label(row['slot_name'], contract)}"
        lines = [
            heading,
            "",
            f"- 当前问题：{safe_markdown_text(summary.get('confirmation_prompt') or '候选值存在冲突，无法直接确定标准值。', max_length=120)}",
            f"- 影响范围：{role_label(role.get('role_kind'), contract)}规则。",
            "- 建议选择：",
            *[f"  {line}" for line in format_histogram_lines(summary.get('value_histogram'), row['slot_name'], summary.get('unit'))],
        ]
        samples = role.get("samples", [])[:2]
        if samples:
            lines.extend(["", "证据："])
            for sample_index, sample in enumerate(samples, start=1):
                lines.append(
                    f"- 样本 {sample_index}：{locator_summary(sample)}，{summarize_role_result(role, contract)}"
                )
        sections.append("\n".join(lines))

    base_index = len(sections)
    for offset, blocker in enumerate(gate_blockers, start=1):
        role_name = role_label(blocker.get("role_kind"), contract)
        slot_name = slot_label(blocker.get("slot_name"), contract)
        heading = f"### {base_index + offset}. {role_name} · {slot_name}"
        options = blocker.get("suggested_options", [])
        lines = [
            heading,
            "",
            f"- 当前问题：{safe_markdown_text(humanize_text(blocker.get('message', '存在待处理阻断项。'), contract), max_length=120)}",
            f"- 影响范围：{role_name}规则。",
            "- 建议选择：",
        ]
        if options:
            for option in options:
                option_text = humanize_text(option.get("label", option) if isinstance(option, dict) else option, contract)
                lines.append(f"  - {safe_markdown_text(option_text, max_length=120)}")
        else:
            lines.append("  - 暂无建议选项")
        sections.append("\n".join(lines))

    if not sections:
        return "无阻断项"
    return "\n\n".join(sections)


def build_manual_review_items_section(unresolved_rows: list[dict[str, Any]], contract: dict[str, Any] | None) -> str:
    """构建人工确认项说明。"""
    sections: list[str] = []
    for index, row in enumerate(unresolved_rows, start=1):
        role = row["role"]
        summary = row["summary"]
        confidence = slot_confidence(summary)
        inferred = format_slot_value(row["slot_name"], summary.get("mode_value"), summary.get("unit"))
        if inferred == "未指定":
            inferred = "尚未确定"
        lines = [
            f"### {index}. {role_label(role.get('role_kind'), contract)} · {slot_label(row['slot_name'], contract)}",
            "",
            f"- 当前推断：{inferred}",
            f"- 置信度：{confidence:.2f}",
            f"- 原因：{safe_markdown_text(humanize_text(summary.get('confirmation_prompt') or '样本不足，仍需人工确认。', contract), max_length=120)}",
            f"- 建议：请确认{role_label(role.get('role_kind'), contract)}的{slot_label(row['slot_name'], contract)}。",
        ]
        sections.append("\n".join(lines))
    if not sections:
        return "无待确认项"
    return "\n\n".join(sections)


def build_conflict_section(conflict_rows: list[dict[str, Any]], contract: dict[str, Any] | None) -> str:
    """构建冲突与异常说明。"""
    sections: list[str] = []
    for row in conflict_rows:
        role = row["role"]
        summary = row["summary"]
        mode_value = format_slot_value(row["slot_name"], summary.get("mode_value"), summary.get("unit"))
        lines = [
            f"### {role_label(role.get('role_kind'), contract)} · {slot_label(row['slot_name'], contract)}",
            "",
            "候选值分布：",
            "",
            *format_histogram_lines(summary.get("value_histogram"), row["slot_name"], summary.get("unit")),
            "",
            "处理建议：",
            f"优先采用主流值 {mode_value}；若少量偏差属于特殊场景，请人工确认后再定稿。",
        ]
        sections.append("\n".join(lines))
    if not sections:
        return "无冲突或异常"
    return "\n\n".join(sections)


def build_evidence_section(roles: list[dict[str, Any]], contract: dict[str, Any] | None) -> str:
    """构建证据样本说明。"""
    sections: list[str] = []
    for role in roles:
        samples = role.get("samples", [])
        if not samples:
            continue
        sample = samples[0]
        lines = [
            f"### {role_label(role.get('role_kind'), contract)}",
            "",
            f"- 样本位置：{locator_summary(sample)}",
            f"- 文本预览：{safe_markdown_text(humanize_text(sample.get('text_preview', ''), contract), max_length=120)}",
            f"- 检测结果：{summarize_role_result(role, contract)}",
        ]
        sections.append("\n".join(lines))
    if not sections:
        return "暂无证据样本"
    return "\n\n".join(sections)


def extract_slot_rows(
    slot_facts: dict[str, Any],
    contract: dict[str, Any] | None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    """抽取已确定、未确定和冲突三类展示行。"""
    resolved_rows: list[dict[str, Any]] = []
    unresolved_rows: list[dict[str, Any]] = []
    conflict_rows: list[dict[str, Any]] = []

    for role in slot_facts.get("roles", []):
        required_slots = role_required_slots(role, contract)
        optional_slots = role_optional_slots(role, contract)
        slot_summary = role.get("slot_summary", {})
        required_statuses = [slot_summary.get(slot, {}).get("status", "unresolved") for slot in required_slots]
        if required_slots and all(status in RESOLVED_STATUSES for status in required_statuses):
            resolved_rows.append({"role": role, "required_slots": required_slots})
        for slot_name, summary in slot_summary.items():
            status = summary.get("status")
            is_required = slot_name in required_slots
            is_optional = slot_name in optional_slots or not is_required
            if status == "unresolved":
                unresolved_rows.append({"role": role, "slot_name": slot_name, "summary": summary})
            if status == "conflict" and is_required:
                conflict_rows.append({"role": role, "slot_name": slot_name, "summary": summary, "severity": "blocking"})
            elif status == "resolved_with_conflicts" and is_required:
                conflict_rows.append({"role": role, "slot_name": slot_name, "summary": summary, "severity": "warning"})
            elif status == "conflict" and is_optional:
                conflict_rows.append({"role": role, "slot_name": slot_name, "summary": summary, "severity": "warning"})
    return resolved_rows, unresolved_rows, conflict_rows


def render_slot_facts_summary(
    slot_facts: dict[str, Any] | None,
    contract: dict[str, Any] | None = None,
) -> tuple[str, dict[str, Any]]:
    """从 role_format_slot_facts.json 渲染 RULE_SUMMARY.md 和 metrics。

    若 slot_facts 为 None 或格式错误，返回错误占位文本并记录 metrics 为全 0。
    """
    if not isinstance(slot_facts, dict):
        return (
            "# 规则摘要\n\n## 错误\n\n规则提取产物 `role_format_slot_facts.json` 缺失或格式错误，无法渲染规则摘要。",
            {"sections_rendered": 0, "error": "slot_facts_missing"},
        )
    resolved_rows, unresolved_rows, conflict_rows = extract_slot_rows(slot_facts, contract)
    gate_blockers = slot_facts.get("gate_blockers", [])
    roles = slot_facts.get("roles", [])

    blocking_conflict_count = sum(1 for row in conflict_rows if row["severity"] == "blocking")
    warning_conflict_count = sum(1 for row in conflict_rows if row["severity"] == "warning")
    metrics = {
        "sections_rendered": 6,
        "resolved_rule_row_count": len(resolved_rows),
        "unresolved_slot_count": len(unresolved_rows),
        "blocking_conflict_count": blocking_conflict_count,
        "warning_conflict_count": warning_conflict_count,
        "gate_blocker_count": len(gate_blockers),
        "resolved_slot_count": sum(
            1
            for role in roles
            for summary in role.get("slot_summary", {}).values()
            if summary.get("status") in {"resolved", "resolved_with_conflicts", "not_applicable"}
        ),
        "conflict_slot_count": blocking_conflict_count + warning_conflict_count,
        "user_confirmed_slot_count": sum(
            1
            for role in roles
            for summary in role.get("slot_summary", {}).values()
            if summary.get("status") == "user_confirmed"
        ),
    }
    view_model = build_rule_summary_view_model(slot_facts, contract, use_icons=True)
    content = render_template(TEMPLATE_PATH.read_text(encoding="utf-8"), view_model)
    return content, metrics


def build_rule_summary_view_model(
    slot_facts: dict[str, Any],
    contract: dict[str, Any] | None,
    *,
    use_icons: bool = True,
) -> dict[str, object]:
    """从槽位事实构建 RULE_SUMMARY view model。"""
    resolved_rows, unresolved_rows, conflict_rows = extract_slot_rows(slot_facts, contract)
    gate_blockers = slot_facts.get("gate_blockers", [])
    roles = slot_facts.get("roles", [])

    conflict_count = len(conflict_rows)
    blocking_count = sum(1 for row in conflict_rows if row["severity"] == "blocking") + len(gate_blockers)
    manual_review_count = len(unresolved_rows)
    gate_status = str(slot_facts.get("gate_status") or "")

    if blocking_count > 0 or gate_status == "blocked":
        status = "blocked"
        status_label = "已阻断"
        stage_conclusion = "当前不能继续，必须先处理阻断项或完成规则确认。"
        next_step = "请先处理阻断项或确认待决规则。"
    elif manual_review_count > 0:
        status = "waiting_user"
        status_label = "需要人工确认"
        stage_conclusion = "当前不能继续，需要先确认待决样式元素。"
        next_step = "请先确认待人工确认的规则项。"
    else:
        status = "accepted"
        status_label = "已通过"
        stage_conclusion = "当前可以继续，规则摘要已达到可打包状态。"
        next_step = "无需进一步操作"

    confirmed_rules_rows = []
    for row in resolved_rows:
        role = row["role"]
        confirmed_rules_rows.append(build_confirmed_rule_row(role, contract))

    contract_ref = slot_facts.get("contract_ref", {})
    technical_appendix_items = [
        f"运行 ID：{safe_markdown_text(slot_facts.get('run_id', '未指定'), max_length=None)}",
        f"来源快照：{safe_markdown_text(slot_facts.get('source_snapshot_path', '未指定'), max_length=None)}",
        f"槽位契约：{safe_markdown_text(contract_ref.get('contract_path', '未指定'), max_length=None)}",
    ]

    return {
        "status_marker": status_marker(status, use_icons=use_icons),
        "status_label": status_label,
        "stage_conclusion": stage_conclusion,
        "confirmed_count": len(resolved_rows),
        "manual_review_count": manual_review_count,
        "blocking_count": blocking_count,
        "conflict_count": conflict_count,
        "next_step": next_step,
        "blocking_items_section": build_blocking_items_section(conflict_rows, gate_blockers, contract),
        "manual_review_items_section": build_manual_review_items_section(unresolved_rows, contract),
        "confirmed_rules_section": markdown_table(CONFIRMED_RULE_TABLE_HEADERS, confirmed_rules_rows, empty_text="暂无已确定规则"),
        "conflict_section": build_conflict_section(conflict_rows, contract),
        "evidence_section": build_evidence_section(roles, contract),
        "technical_appendix_section": markdown_list(technical_appendix_items, empty_text="无技术附录"),
    }


def render_rule_summary(
    slot_facts: dict[str, Any],
    rule_package: dict[str, Any] | None,
    contract: dict[str, Any] | None,
    output_path: Path,
) -> dict[str, Any]:
    """从槽位事实渲染 RULE_SUMMARY.md，并返回 skill-result.metrics 所需字段。"""
    del rule_package
    content, metrics = render_slot_facts_summary(slot_facts, contract)
    assert_human_readable_report(
        content,
        report_kind="rule_summary",
        required_sections=RULE_SUMMARY_REQUIRED_SECTIONS,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    temp_path.write_text(content, encoding="utf-8")
    os.replace(temp_path, output_path)
    data = output_path.read_bytes()
    metrics.update(
        {
            "output_path": str(output_path),
            "sha256": hashlib.sha256(data).hexdigest(),
            "size_bytes": len(data),
        }
    )
    return metrics


def scan_rule_summary_text(content: str) -> list[str]:
    """兼容旧测试接口，复用新的用户可读报告校验器。"""
    try:
        assert_human_readable_report(
            content,
            report_kind="rule_summary",
            required_sections=RULE_SUMMARY_REQUIRED_SECTIONS,
        )
    except ValueError as exc:
        return [str(exc)]
    return []


def render_summary(draft: dict[str, Any]) -> str:
    """拒绝旧 draft 摘要通道，避免绕过 slot facts 约束。"""
    del draft
    raise ValueError("RULE_SUMMARY.md 必须从 role_format_slot_facts.json 渲染，不再支持 semantic_rule_draft.json 直出。")


def main() -> int:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="从槽位事实生成 RULE_SUMMARY.md")
    parser.add_argument("--draft", type=Path)
    parser.add_argument("--slot-facts", type=Path)
    parser.add_argument("--contract", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()

    if args.slot_facts:
        if args.contract is None:
            parser.error("使用 --slot-facts 时必须同时提供 --contract")
        slot_facts = json.loads(args.slot_facts.read_text(encoding="utf-8"))
        metrics = render_rule_summary(slot_facts, {}, load_contract(args.contract), args.output)
        print(json.dumps(metrics, ensure_ascii=False, indent=2))
        return 0
    if args.draft:
        parser.error("RULE_SUMMARY.md 必须使用 --slot-facts + --contract 生成，--draft 已禁用")
    parser.error("必须提供 --slot-facts")
    return 0


def main_from_test(draft: Path, output: Path) -> int:
    """测试入口：按 CLI 等价逻辑生成摘要。"""
    del draft, output
    raise ValueError("RULE_SUMMARY.md 必须从 role_format_slot_facts.json 渲染，不再支持 semantic_rule_draft.json 直出。")


def main_from_slot_facts_test(slot_facts: Path, output: Path, contract: Path | None = None) -> dict[str, Any]:
    """测试入口：按槽位事实生成摘要并返回 metrics。"""
    data = json.loads(slot_facts.read_text(encoding="utf-8"))
    return render_rule_summary(data, {}, load_contract(contract), output)


if __name__ == "__main__":
    raise SystemExit(main())
