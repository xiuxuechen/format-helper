#!/usr/bin/env python3
"""按 V2 角色模型执行第一版格式审计。"""
from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path


TASK_OWNER = {
    "T01": "title-agent",
    "T02": "toc-agent",
    "T03": "heading-agent",
    "T04": "body-agent",
    "T05": "table-agent",
    "T06": "page-agent",
}


TASK_TYPES = {
    "T02": {"toc-title", "toc-static-item", "toc-field", "heading-level-1", "heading-level-2", "heading-level-3"},
    "T03": {"heading-level-1", "heading-level-2", "heading-level-3", "heading-level-4", "ambiguous-numbered-item"},
    "T04": {"body-paragraph", "body-no-indent"},
}


NATIVE_HEADING_STYLE = {
    "heading-level-1": ("Heading1", 1),
    "heading-level-2": ("Heading2", 2),
    "heading-level-3": ("Heading3", 3),
    "heading-level-4": ("Heading4", 4),
}


ROLE_CONFIDENCE_THRESHOLD = 0.85
BODY_RUN_KEYS = ("font_east_asia", "font_size_pt")
BODY_SPACING_KEYS = ("line_spacing_multiple", "line_spacing_pt")
BODY_INDENT_KEYS = ("first_line_indent_cm", "left_indent_cm", "right_indent_cm", "hanging_indent_cm")
BODY_FORMAT_KEYS = BODY_RUN_KEYS + BODY_SPACING_KEYS + BODY_INDENT_KEYS
DEFAULT_TABLE_RULE_MODE = "audit-only"
TABLE_RULE_MODES = {"skip", "audit-only", "auto-fix"}


def useful(value) -> bool:
    return value not in (None, "")


def numeric_equal(left, right, tolerance: float = 0.01) -> bool:
    try:
        return abs(float(left) - float(right)) <= tolerance
    except (TypeError, ValueError):
        return False


def values_match(left, right) -> bool:
    if not useful(left) or not useful(right):
        return True
    if numeric_equal(left, right):
        return True
    return str(left).strip().lower() == str(right).strip().lower()


def scalar(value: str | None):
    if value in (None, "", "null", "None"):
        return None
    if value in {"true", "false"}:
        return value == "true"
    try:
        if "." in value:
            return float(value)
        return int(value)
    except (TypeError, ValueError):
        return value


def resolve_table_rules_path(rule_dir: str | None, table_rules: str | None) -> Path | None:
    if table_rules:
        return Path(table_rules)
    if rule_dir:
        return Path(rule_dir) / "table-rules.yaml"
    return None


def parse_table_rules(path: Path | None) -> dict:
    if path is None or not path.exists():
        return {}
    data: dict = {}
    stack: list[tuple[int, dict]] = [(-1, data)]
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.split("#", 1)[0].rstrip()
        if not line.strip() or ":" not in line:
            continue
        indent = len(line) - len(line.lstrip(" "))
        key, value = line.strip().split(":", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        while stack and indent <= stack[-1][0]:
            stack.pop()
        parent = stack[-1][1]
        if value == "":
            section: dict = {}
            parent[key] = section
            stack.append((indent, section))
        else:
            parent[key] = scalar(value)
    return data


def table_rule_mode(table_rules: dict, rule_path: str) -> str:
    current = table_rules.get("tables") if isinstance(table_rules.get("tables"), dict) else {}
    mode = DEFAULT_TABLE_RULE_MODE
    for part in rule_path.split("."):
        if not isinstance(current, dict):
            break
        candidate = current.get(part)
        if isinstance(candidate, dict):
            raw_mode = candidate.get("mode")
            if isinstance(raw_mode, str) and raw_mode in TABLE_RULE_MODES:
                mode = raw_mode
            current = candidate
            continue
        break
    return mode


def table_rule_enabled(table_rules: dict, rule_path: str) -> bool:
    return table_rule_mode(table_rules, rule_path) != "skip"


def merge_effective_format(paragraph: dict, style_details: dict) -> dict:
    style = paragraph.get("style")
    style_detail = style_details.get(style, {}) if style else {}
    run = dict(style_detail.get("run_format") or {})
    para = dict(style_detail.get("paragraph_format") or {})
    run.update({k: v for k, v in (paragraph.get("run_format") or {}).items() if useful(v)})
    para.update({k: v for k, v in (paragraph.get("paragraph_format") or {}).items() if useful(v)})
    return {"run_format": run, "paragraph_format": para}


def issue_item_with_effective_format(paragraph: dict, style_details: dict) -> dict:
    effective = merge_effective_format(paragraph, style_details)
    item = dict(paragraph)
    item["run_format"] = effective["run_format"]
    item["paragraph_format"] = effective["paragraph_format"]
    return item


def most_common(values):
    filtered = [value for value in values if useful(value)]
    if not filtered:
        return None
    return Counter(filtered).most_common(1)[0][0]


def normal_style_baseline(style_details: dict) -> dict:
    style_detail = style_details.get("Normal") or style_details.get("normal") or {}
    run = style_detail.get("run_format") or {}
    para = style_detail.get("paragraph_format") or {}
    return {
        "font_east_asia": run.get("font_east_asia"),
        "font_size_pt": run.get("font_size_pt"),
        "line_spacing_multiple": para.get("line_spacing_multiple"),
        "line_spacing_pt": para.get("line_spacing_pt"),
        "first_line_indent_cm": para.get("first_line_indent_cm"),
        "left_indent_cm": para.get("left_indent_cm"),
        "right_indent_cm": para.get("right_indent_cm"),
        "hanging_indent_cm": para.get("hanging_indent_cm"),
    }


def representative_body_baseline(paragraphs: list[dict], style_details: dict) -> dict:
    body_paragraphs = [p for p in paragraphs if p.get("element_type") == "body-paragraph"]
    merged = [merge_effective_format(p, style_details) for p in body_paragraphs]
    representative = {
        "font_east_asia": most_common([item["run_format"].get("font_east_asia") for item in merged]),
        "font_size_pt": most_common([item["run_format"].get("font_size_pt") for item in merged]),
        "line_spacing_multiple": most_common([item["paragraph_format"].get("line_spacing_multiple") for item in merged]),
        "line_spacing_pt": most_common([item["paragraph_format"].get("line_spacing_pt") for item in merged]),
        "first_line_indent_cm": most_common([item["paragraph_format"].get("first_line_indent_cm") for item in merged]),
        "left_indent_cm": most_common([item["paragraph_format"].get("left_indent_cm") for item in merged]),
        "right_indent_cm": most_common([item["paragraph_format"].get("right_indent_cm") for item in merged]),
        "hanging_indent_cm": most_common([item["paragraph_format"].get("hanging_indent_cm") for item in merged]),
    }
    normal = normal_style_baseline(style_details)
    return {key: normal.get(key) if useful(normal.get(key)) else representative.get(key) for key in BODY_FORMAT_KEYS}


def baseline_has_minimum_fields(baseline: dict) -> bool:
    has_font = useful(baseline.get("font_east_asia"))
    has_size = useful(baseline.get("font_size_pt"))
    has_spacing = any(useful(baseline.get(key)) for key in BODY_SPACING_KEYS)
    has_indent = any(useful(baseline.get(key)) for key in BODY_INDENT_KEYS)
    return has_font and has_size and has_spacing and has_indent


def body_format_mismatches(paragraph: dict, baseline: dict, style_details: dict) -> dict:
    effective = merge_effective_format(paragraph, style_details)
    current = {
        "font_east_asia": effective["run_format"].get("font_east_asia"),
        "font_size_pt": effective["run_format"].get("font_size_pt"),
        "line_spacing_multiple": effective["paragraph_format"].get("line_spacing_multiple"),
        "line_spacing_pt": effective["paragraph_format"].get("line_spacing_pt"),
        "first_line_indent_cm": effective["paragraph_format"].get("first_line_indent_cm"),
        "left_indent_cm": effective["paragraph_format"].get("left_indent_cm"),
        "right_indent_cm": effective["paragraph_format"].get("right_indent_cm"),
        "hanging_indent_cm": effective["paragraph_format"].get("hanging_indent_cm"),
    }
    mismatches = {}
    for key in BODY_FORMAT_KEYS:
        if key == "line_spacing_pt" and useful(baseline.get("line_spacing_multiple")):
            continue
        if key == "line_spacing_multiple" and useful(baseline.get("line_spacing_pt")) and not useful(baseline.get("line_spacing_multiple")):
            continue
        if useful(baseline.get(key)) and useful(current.get(key)) and not values_match(current.get(key), baseline.get(key)):
            mismatches[key] = {"current": current.get(key), "expected": baseline.get(key)}
    return mismatches


def heading_requires_manual_review(paragraph: dict) -> tuple[bool, str]:
    confidence = paragraph.get("role_confidence")
    numbering_pattern = paragraph.get("numbering_pattern")
    classification_reason = paragraph.get("classification_reason")
    try:
        confidence_value = float(confidence)
    except (TypeError, ValueError):
        confidence_value = 0.0
    if confidence_value < ROLE_CONFIDENCE_THRESHOLD:
        return True, f"role_confidence={confidence!r} below {ROLE_CONFIDENCE_THRESHOLD}; reason={classification_reason!r}"
    if numbering_pattern == "ambiguous-numbered-item":
        return True, f"numbering_pattern={numbering_pattern!r}; reason={classification_reason!r}"
    return False, ""


def style_id_status(snapshot: dict, target_style_id: str) -> str:
    available_style_ids = snapshot.get("available_style_ids")
    if not isinstance(available_style_ids, list):
        return "unknown"
    return "available" if target_style_id in set(available_style_ids) else "missing"


def format_source(item: dict) -> str:
    if item.get("has_direct_format"):
        return "direct-format"
    if item.get("style"):
        return "style-definition"
    return "inherited"


def element_ref(item: dict) -> dict:
    return {
        "element_id": item.get("element_id"),
        "element_type": item.get("element_type"),
        "paragraph_index": item.get("paragraph_index"),
        "text_preview": item.get("text_preview"),
    }


def make_issue(
    task_id: str,
    index: int,
    item: dict,
    problem: str,
    issue_type: str,
    action_type: str,
    policy: str,
    severity: str = "medium",
    expected_role: str | None = None,
    expected_format: dict | None = None,
    risk_flags: list[str] | None = None,
    confidence: float = 0.86,
) -> dict:
    detected_role = item.get("element_type")
    return {
        "issue_id": f"{task_id}-I{index:03d}",
        "issue_type": issue_type,
        "severity": severity,
        "confidence": confidence,
        "element_ref": element_ref(item),
        "detected_role": detected_role,
        "expected_role": expected_role or detected_role,
        "problem": problem,
        "format_source": format_source(item),
        "current_format": {
            "style": item.get("style"),
            **(item.get("run_format") or {}),
            **(item.get("paragraph_format") or {}),
        },
        "expected_format": expected_format or {},
        "recommended_action": {
            "action_type": action_type,
            "auto_fix_policy": policy,
        },
        "risk_flags": risk_flags or [],
    }


def make_table_issue(task_id: str, index: int, table: dict, problem: str, policy: str, severity: str = "low") -> dict:
    return {
        "issue_id": f"{task_id}-I{index:03d}",
        "issue_type": "format_mismatch",
        "severity": severity,
        "confidence": 0.82,
        "element_ref": {
            "element_id": table.get("element_id"),
            "element_type": "normal-table",
            "table_index": table.get("table_index"),
            "paragraph_index": None,
            "text_preview": table.get("text_preview") or f"第 {table.get('table_index')} 个表格",
        },
        "detected_role": "normal-table",
        "expected_role": "normal-table",
        "problem": problem,
        "format_source": "mixed",
        "current_format": {
            "alignment": table.get("alignment"),
            "width": table.get("width"),
            "width_type": table.get("width_type"),
            "has_borders": table.get("has_borders"),
            "header_row_count": table.get("header_row_count"),
        },
        "expected_format": {"role": "table-header/table-body true cell format"},
        "recommended_action": {
            "action_type": "apply_table_cell_format",
            "auto_fix_policy": policy,
        },
        "risk_flags": ["merged_cells"] if table.get("has_merged_cells") or table.get("merged_cell_count", 0) else [],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--snapshot", required=True)
    parser.add_argument("--task-id", required=True)
    parser.add_argument("--phase", choices=["audit", "review"], default="audit")
    parser.add_argument("--output", required=True)
    parser.add_argument("--rule-dir")
    parser.add_argument("--table-rules")
    args = parser.parse_args()

    snapshot = json.loads(Path(args.snapshot).read_text(encoding="utf-8-sig"))
    table_rules = parse_table_rules(resolve_table_rules_path(args.rule_dir, args.table_rules))
    issues = []
    paragraphs = snapshot.get("paragraphs", [])
    style_details = snapshot.get("style_details") or {}
    body_baseline = representative_body_baseline(paragraphs, style_details)
    task_types = TASK_TYPES.get(args.task_id)
    checked = 0

    if args.task_id == "T02":
        checked = len(paragraphs)
        if not snapshot.get("has_toc_field"):
            pseudo = {"element_id": "toc-field-0001", "element_type": "toc-field", "paragraph_index": None, "text_preview": "未检测到自动目录", "style": None}
            issues.append(make_issue(args.task_id, len(issues) + 1, pseudo, "未检测到 Word 自动目录字段", "format_mismatch", "insert_or_replace_toc_field", "manual-review", "high"))
        for paragraph in paragraphs:
            if paragraph.get("element_type") == "toc-static-item":
                issues.append(make_issue(args.task_id, len(issues) + 1, paragraph, "检测到静态目录条目，替换前必须执行目录内容校验", "toc_content_mismatch", "toc_content_audit", "manual-review", "medium"))
    elif task_types:
        for paragraph in paragraphs:
            if paragraph.get("element_type") not in task_types:
                continue
            checked += 1
            element_type = paragraph.get("element_type", "")
            style = paragraph.get("style") or ""
            if element_type == "ambiguous-numbered-item":
                issues.append(
                    make_issue(
                        args.task_id,
                        len(issues) + 1,
                        paragraph,
                        "编号段落可能是标题或正文列表项，需人工确认后才能映射为标题",
                        "role_classification",
                        "manual_review_heading_role",
                        "manual-review",
                        "medium",
                        expected_format={
                            "numbering_pattern": paragraph.get("numbering_pattern"),
                            "role_confidence_threshold": ROLE_CONFIDENCE_THRESHOLD,
                        },
                        risk_flags=["ambiguous_numbered_item"],
                        confidence=paragraph.get("role_confidence") or 0.45,
                    )
                )
            elif element_type.startswith("heading-level-"):
                expected_style, outline_level = NATIVE_HEADING_STYLE.get(element_type, ("Heading1", 1))
                requires_review, review_reason = heading_requires_manual_review(paragraph)
                if requires_review:
                    issues.append(
                        make_issue(
                            args.task_id,
                            len(issues) + 1,
                            paragraph,
                            f"Heading role is not safe for automatic mapping: {review_reason}",
                            "role_classification",
                            "manual_review_heading_role",
                            "manual-review",
                            "medium",
                            expected_format={
                                "word_style_id": expected_style,
                                "outline_level": outline_level,
                                "role_confidence_threshold": ROLE_CONFIDENCE_THRESHOLD,
                            },
                            risk_flags=["ambiguous_role_classification"],
                            confidence=0.75,
                        )
                    )
                    continue
                if style not in {expected_style, f"Heading {outline_level}", f"标题 {outline_level}"}:
                    style_status = style_id_status(snapshot, expected_style)
                    if style_status != "available":
                        severity = "blocker" if style_status == "missing" else "medium"
                        issues.append(
                            make_issue(
                                args.task_id,
                                len(issues) + 1,
                                paragraph,
                                f"Cannot auto-map heading: target styleId {expected_style!r} is {style_status} in snapshot.available_style_ids",
                                "style_mapping",
                                "manual_review_heading_style_mapping",
                                "manual-review",
                                severity,
                                expected_format={
                                    "word_style_id": expected_style,
                                    "outline_level": outline_level,
                                    "available_style_ids_status": style_status,
                                },
                                risk_flags=["missing_target_style_id"] if style_status == "missing" else ["unknown_available_style_ids"],
                                confidence=0.8,
                            )
                        )
                        continue
                    issues.append(
                        make_issue(
                            args.task_id,
                            len(issues) + 1,
                            paragraph,
                            "高置信度标题未映射到 Word 原生标题样式",
                            "style_mapping",
                            "map_heading_native_style",
                            "auto-fix",
                            expected_format={"word_style_id": expected_style, "outline_level": outline_level},
                        )
                    )
            elif element_type in {"body-paragraph", "body-no-indent"}:
                effective_item = issue_item_with_effective_format(paragraph, style_details)
                if paragraph.get("has_direct_format"):
                    issues.append(
                        make_issue(
                            args.task_id,
                            len(issues) + 1,
                            effective_item,
                            "Body paragraph has direct formatting override; keep for manual review/direct-format strategy, do not auto-clean.",
                            "format_mismatch",
                            "review_body_direct_format_strategy",
                            "manual-review",
                            "low",
                            expected_format={"format_write_strategy": "style-definition-or-explicit-direct-format"},
                            risk_flags=["direct_format_override"],
                            confidence=0.82,
                        )
                    )
                if not baseline_has_minimum_fields(body_baseline):
                    issues.append(
                        make_issue(
                            args.task_id,
                            len(issues) + 1,
                            effective_item,
                            "Unable to determine a complete body format baseline from body-paragraph representatives or Normal style.",
                            "format_mismatch",
                            "manual_review_body_format_baseline",
                            "manual-review",
                            "medium",
                            expected_format=body_baseline,
                            risk_flags=["missing_body_format_baseline"],
                            confidence=0.72,
                        )
                    )
                    continue
                mismatches = body_format_mismatches(paragraph, body_baseline, style_details)
                if mismatches:
                    issues.append(
                        make_issue(
                            args.task_id,
                            len(issues) + 1,
                            effective_item,
                            "Body paragraph effective format differs from the representative body/Normal baseline.",
                            "format_mismatch",
                            "review_body_format_mismatch",
                            "manual-review",
                            "medium",
                            expected_format={key: value["expected"] for key, value in mismatches.items()},
                            risk_flags=["effective_format_mismatch"],
                            confidence=0.84,
                        )
                    )
    elif args.task_id == "T05":
        tables = snapshot.get("table_details") or snapshot.get("tables") or []
        checked = len(tables)
        for table in tables:
            cell_paragraphs = table.get("cell_paragraphs", [])
            if not cell_paragraphs:
                continue
            format_gaps = []
            if table_rule_enabled(table_rules, "border") and not table.get("has_borders"):
                format_gaps.append("边框")
            if table_rule_enabled(table_rules, "layout.alignment") and table.get("alignment") != "center":
                format_gaps.append("表格居中")
            if table_rule_enabled(table_rules, "header") and table.get("header_row_count", 0) < 1:
                format_gaps.append("表头重复")
            if table_rule_enabled(table_rules, "header.shading") and table.get("first_row_shading_cell_count", 0) == 0:
                format_gaps.append("表头底纹")
            if (
                table_rule_enabled(table_rules, "cell.vertical_alignment")
                and table.get("cell_count", 0)
                and table.get("vertically_centered_cell_count", 0) < table.get("cell_count", 0)
            ):
                format_gaps.append("单元格垂直居中")
            if table_rule_enabled(table_rules, "cell.margin") and any(table.get(f"cell_margin_{side}") in {None, ""} for side in ("top", "left", "bottom", "right")):
                format_gaps.append("单元格内边距")
            if not format_gaps:
                continue
            problem = "表格真实格式未统一：" + "、".join(format_gaps)
            policy = "manual-review" if table.get("has_merged_cells") or table.get("merged_cell_count", 0) else "auto-fix"
            if policy == "manual-review":
                problem += "；表格含合并单元格，本次不自动调整"
            issues.append(make_table_issue(args.task_id, len(issues) + 1, table, problem, policy, "low"))

    output = {
        "schema_version": "2.0.0",
        "task_id": args.task_id,
        "agent": TASK_OWNER.get(args.task_id, "docx-format-auditor"),
        "phase": args.phase,
        "source_snapshot": args.snapshot,
        "summary": {
            "checked_items": checked,
            "issues_found": len(issues),
            "auto_fixable": sum(1 for i in issues if i["recommended_action"]["auto_fix_policy"] == "auto-fix"),
            "manual_review": sum(1 for i in issues if i["recommended_action"]["auto_fix_policy"] != "auto-fix"),
            "blocked": sum(1 for i in issues if i["severity"] == "blocker"),
        },
        "issues": issues,
    }
    Path(args.output).write_text(json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
