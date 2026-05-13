"""用户可读 Markdown 报告校验器。"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


FORBIDDEN_PLACEHOLDER_WORDS = (
    "DIFF_SUMMARY",
    "TODO",
    "PLACEHOLDER",
)

FORBIDDEN_MACHINE_KEYS = (
    "body-paragraph",
    "heading-level-1",
    "heading-level-2",
    "heading-level-3",
    "heading-level-4",
    "toc-level-1",
    "toc-level-2",
    "toc-level-3",
    "list-paragraph",
    "table-content",
    "header-footer",
    "section-page-setup",
    "font_size_pt",
    "font_east_asia",
    "font_ascii",
    "line_spacing_multiple",
    "required_slots",
    "optional_slots",
    "slot_name",
    "role_kind",
    "mode_value",
    "vertical_alignment",
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
)

FORBIDDEN_TECH_UNITS = (
    "twip",
    "twips",
    "half_pt",
)

TEMPLATE_PLACEHOLDER_PATTERN = r"\{[a-z][a-z0-9_]*(?:_section|_table|_list)?\}"
PT_PATTERN = re.compile(r"\b\d+(?:\.\d+)?\s*pt\b")

DEFAULT_ALLOWED_EMPTY_SECTION_TEXTS = {
    "必须先处理的问题": ["无阻断项"],
    "需要人工确认的项目": ["无待确认项"],
    "冲突与异常说明": ["无冲突或异常"],
    "审计发现摘要": ["未发现格式问题"],
    "修复执行摘要": ["本次未执行自动修复"],
    "修复前后对比": ["无可展示的修复前后对比项"],
    "未修复项与原因": ["无未修复项"],
    "风险和限制": ["无已知剩余风险"],
    "下一步": ["无需进一步操作"],
}

RULE_SUMMARY_REQUIRED_SECTIONS = [
    "当前结论",
    "必须先处理的问题",
    "需要人工确认的项目",
    "已确定的格式规则",
    "冲突与异常说明",
    "证据样本",
    "技术附录",
]

FINAL_REPORT_REQUIRED_SECTIONS = [
    "当前结论",
    "本次输入与规则来源",
    "审计发现摘要",
    "修复执行摘要",
    "修复前后对比",
    "未修复项与原因",
    "风险和限制",
    "验收证据",
    "下一步",
]


@dataclass
class HumanReadableReportValidationResult:
    valid: bool
    score: int
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _section_map(content: str) -> dict[str, str]:
    pattern = re.compile(r"^##\s+(.+?)\s*$", re.MULTILINE)
    matches = list(pattern.finditer(content))
    sections: dict[str, str] = {}
    for index, match in enumerate(matches):
        title = match.group(1).strip()
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(content)
        sections[title] = content[start:end].strip()
    return sections


def _is_allowed_empty(section_name: str, body: str, allowed_empty_section_texts: dict[str, list[str]]) -> bool:
    allowed = allowed_empty_section_texts.get(section_name, [])
    return any(text in body for text in allowed)


def _validate_status_markers(content: str) -> list[str]:
    errors: list[str] = []
    checks = (
        ("🔴", ("阻断", "必须先处理")),
        ("🟡", ("待确认", "人工确认")),
        ("🟢", ("可继续", "已通过", "已完成")),
        ("✅", ("已完成", "已修复")),
        ("⚠️", ("风险", "冲突", "警告")),
    )
    for marker, keywords in checks:
        if marker in content and not any(keyword in content for keyword in keywords):
            errors.append(f"状态标记 {marker} 与状态文本不一致")
    return errors


def _score_report(report_kind: str, sections: dict[str, str], errors: list[str]) -> int:
    score = 0
    current = sections.get("当前结论", "")
    if current and any(token in current for token in ("结论", "通过", "阻断", "可继续", "已完成", "已通过", "不能继续", "需要先确认", "已阻断")):
        score += 15
    if any(name in sections for name in ("必须先处理的问题", "需要人工确认的项目", "未修复项与原因", "风险和限制")):
        score += 20
    if ("已确定的格式规则" in sections and sections.get("已确定的格式规则")) or (
        "修复执行摘要" in sections and sections.get("修复执行摘要")
    ):
        score += 20
    if report_kind == "final_report":
        if sections.get("修复前后对比") and sections.get("验收证据"):
            score += 15
    else:
        score += 15
    if not errors:
        score += 15
    if report_kind == "final_report" and sections.get("下一步"):
        score += 10
    if report_kind == "rule_summary" and any(token in current for token in ("建议下一步", "下一步", "无需进一步操作")):
        score += 10
    summary_sections = (
        "必须先处理的问题",
        "需要人工确认的项目",
        "冲突与异常说明",
        "未修复项与原因",
        "风险和限制",
    )
    if any(name in sections for name in summary_sections):
        score += 5
    full_content = "\n".join(sections.values())
    if any(marker in current or marker in full_content for marker in ("🔴", "🟡", "🟢", "✅", "⚠️", "[阻断]", "[待确认]", "[可继续]", "[已通过]", "[已完成]", "[风险]")):
        score += 5
    return score


def validate_human_readable_report(
    content: str,
    *,
    report_kind: str,
    required_sections: list[str],
    allowed_empty_section_texts: dict[str, list[str]] | None = None,
) -> HumanReadableReportValidationResult:
    """校验用户可读报告正文。"""
    errors: list[str] = []
    warnings: list[str] = []
    if report_kind not in {"rule_summary", "final_report"}:
        return HumanReadableReportValidationResult(False, 0, [f"report_kind 不支持：{report_kind}"], warnings)
    allowed_empty_section_texts = allowed_empty_section_texts or DEFAULT_ALLOWED_EMPTY_SECTION_TEXTS

    for word in FORBIDDEN_PLACEHOLDER_WORDS:
        if word in content:
            errors.append(f"报告包含禁止占位符或内部字段：{word}")
    if re.search(TEMPLATE_PLACEHOLDER_PATTERN, content):
        errors.append("报告存在未替换模板占位符")
    for key in FORBIDDEN_MACHINE_KEYS:
        if key in content:
            errors.append(f"报告包含禁止机器字段：{key}")
    lowered = content.lower()
    for unit in FORBIDDEN_TECH_UNITS:
        if unit in lowered:
            errors.append(f"报告包含禁止技术单位：{unit}")
    if PT_PATTERN.search(content):
        errors.append("报告包含禁止技术单位：pt")

    sections = _section_map(content)
    for section in required_sections:
        if section not in sections:
            errors.append(f"缺少必备章节：{section}")
            continue
        body = sections[section].strip()
        if not body:
            errors.append(f"章节为空：{section}")
            continue
        if body in {"-", "无", "暂无"} and not _is_allowed_empty(section, body, allowed_empty_section_texts):
            errors.append(f"章节内容过空：{section}")
            continue
        if _is_allowed_empty(section, body, allowed_empty_section_texts):
            continue
        normalized = body.replace("-", "").strip()
        if not normalized:
            errors.append(f"章节内容过空：{section}")

    errors.extend(_validate_status_markers(content))
    score = _score_report(report_kind, sections, errors)
    valid = not errors and score >= 95
    return HumanReadableReportValidationResult(valid, score, errors, warnings)


def assert_human_readable_report(
    content: str,
    *,
    report_kind: str,
    required_sections: list[str],
    allowed_empty_section_texts: dict[str, list[str]] | None = None,
) -> None:
    """校验失败时抛出异常。"""
    result = validate_human_readable_report(
        content,
        report_kind=report_kind,
        required_sections=required_sections,
        allowed_empty_section_texts=allowed_empty_section_texts,
    )
    if not result.valid:
        details = "；".join(result.errors) or f"score={result.score}"
        raise ValueError(f"报告渲染不满足用户可读性 Gate：{details}")
