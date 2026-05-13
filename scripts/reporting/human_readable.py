"""用户可读报告渲染公共工具。"""

from __future__ import annotations

import re
from typing import Any


STATUS_ICON_MAP = {
    "blocked": "🔴",
    "waiting_user": "🟡",
    "passed": "🟢",
    "accepted": "🟢",
    "accepted_with_warnings": "⚠️",
    "done": "✅",
}

STATUS_TEXT_FALLBACK_MAP = {
    "blocked": "[阻断]",
    "waiting_user": "[待确认]",
    "passed": "[可继续]",
    "accepted": "[已通过]",
    "accepted_with_warnings": "[风险]",
    "done": "[已完成]",
}

TEMPLATE_PLACEHOLDER_RE = re.compile(r"\{([a-z][a-z0-9_]*(?:_section|_table|_list)?)\}")
CONTROL_CHAR_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]")
WHITESPACE_RE = re.compile(r"\s+")

FONT_SIZE_LABELS = {
    9.0: "小五",
    10.5: "五号",
    12.0: "小四",
    14.0: "四号",
    15.0: "小三",
    16.0: "三号",
    18.0: "小二",
    22.0: "二号",
}

ALIGNMENT_LABELS = {
    "left": "左对齐",
    "center": "居中",
    "right": "右对齐",
    "both": "两端对齐",
    "justify": "两端对齐",
    "distributed": "分散对齐",
    "distribute": "分散对齐",
}


def status_marker(status: str, *, use_icons: bool = True) -> str:
    """返回状态图标或纯文本降级标签。"""
    if use_icons:
        return STATUS_ICON_MAP.get(status, "[状态未知]")
    return STATUS_TEXT_FALLBACK_MAP.get(status, "[状态未知]")


def _to_float(value: object) -> float | None:
    try:
        if value is None or value == "":
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def _round_key(value: float) -> float:
    return round(value, 2)


def human_font_size(value_pt: object) -> str:
    """将磅值转换为用户熟悉的中文字号。"""
    value = _to_float(value_pt)
    if value is None:
        return "未指定"
    rounded = _round_key(value)
    if rounded in FONT_SIZE_LABELS:
        return FONT_SIZE_LABELS[rounded]
    best_key = min(FONT_SIZE_LABELS, key=lambda item: abs(item - rounded))
    if abs(best_key - rounded) <= 0.25:
        return f"约{FONT_SIZE_LABELS[best_key]}"
    if float(rounded).is_integer():
        number = str(int(rounded))
    else:
        number = str(rounded)
    return f"非标准字号（约 {number} 磅）"


def human_line_spacing(value: object, *, unit: str | None = None) -> str:
    """将行距转换为办公软件用户熟悉的表达。"""
    numeric = _to_float(value)
    if numeric is None:
        return "未指定"
    if unit == "pt":
        if float(numeric).is_integer():
            number = str(int(numeric))
        else:
            number = str(round(numeric, 2))
        return f"固定值 {number} 磅"
    if numeric == 1 or numeric == 1.0:
        return "单倍行距"
    if numeric == 1.5:
        return "1.5 倍行距"
    if numeric == 2 or numeric == 2.0:
        return "2 倍行距"
    number = str(int(numeric)) if float(numeric).is_integer() else str(round(numeric, 2))
    return f"{number} 倍行距"


def human_indent(value_cm: object, *, kind: str) -> str:
    """将缩进转换为用户可读表达。"""
    del kind
    numeric = _to_float(value_cm)
    if numeric is None:
        return "未指定"
    if numeric == 0:
        return "无缩进"
    if 0.70 <= numeric <= 0.75:
        return "约 2 字符"
    number = str(int(numeric)) if float(numeric).is_integer() else str(round(numeric, 2))
    return f"约 {number} 厘米"


def human_alignment(value: object) -> str:
    """将对齐值转换为中文。"""
    if value is None or value == "":
        return "未指定"
    return ALIGNMENT_LABELS.get(str(value), str(value))


def safe_markdown_text(value: object, *, max_length: int | None = 120, table_cell: bool = False) -> str:
    """清理用户报告中的动态文本。"""
    text = "" if value is None else str(value)
    text = CONTROL_CHAR_RE.sub("", text)
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    text = WHITESPACE_RE.sub(" ", text).strip()
    if table_cell:
        text = text.replace("|", r"\|")
    if max_length is not None and len(text) > max_length:
        text = text[: max_length - 1].rstrip() + "…"
    return text


def render_template(template_text: str, values: dict[str, object]) -> str:
    """用 view model 渲染模板。"""

    def replace(match: re.Match[str]) -> str:
        key = match.group(1)
        if key not in values:
            return match.group(0)
        value = values[key]
        return "" if value is None else str(value)

    return TEMPLATE_PLACEHOLDER_RE.sub(replace, template_text)


def markdown_list(items: list[str], *, empty_text: str) -> str:
    """渲染 Markdown 列表。"""
    cleaned = [safe_markdown_text(item, max_length=120, table_cell=False) for item in items if safe_markdown_text(item, max_length=120, table_cell=False)]
    if not cleaned:
        return empty_text
    return "\n".join(f"- {item}" for item in cleaned)


def markdown_table(headers: list[str], rows: list[list[str]], *, empty_text: str) -> str:
    """渲染 Markdown 表格。"""
    if not headers or not rows:
        return empty_text
    cleaned_headers = [safe_markdown_text(item, max_length=None, table_cell=True) for item in headers]
    cleaned_rows = []
    for row in rows:
        cleaned_rows.append([safe_markdown_text(item, max_length=120, table_cell=True) for item in row])
    header_line = "| " + " | ".join(cleaned_headers) + " |"
    separator_line = "| " + " | ".join("---" for _ in cleaned_headers) + " |"
    row_lines = ["| " + " | ".join(row) + " |" for row in cleaned_rows]
    return "\n".join([header_line, separator_line, *row_lines])
