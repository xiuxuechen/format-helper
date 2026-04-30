#!/usr/bin/env python3
"""按用户确认的会议方案格式规则生成语义审计。"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[4]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.validation.check_semantic_audit import validate_semantic_audit


TZ = timezone(timedelta(hours=8))
LEVEL1_PATTERN = re.compile(r"^[一二三四五六七八九十]+、")
TITLE_RULE = {
    "font_east_asia": "方正小标宋简体",
    "font_ascii": "Times New Roman",
    "font_size_pt": 22.0,
    "alignment": "center",
    "first_line_indent_cm": 0.0,
    "line_spacing_pt": 35.0,
}
LEVEL1_RULE = {
    "font_east_asia": "黑体",
    "font_ascii": "Times New Roman",
    "font_size_pt": 16.0,
    "bold": False,
    "alignment": "left",
    "first_line_indent_cm": 1.13,
    "line_spacing_pt": 28.0,
}
BODY_RULE = {
    "font_east_asia": "方正仿宋_GB2312",
    "font_ascii": "Times New Roman",
    "font_size_pt": 16.0,
    "first_line_indent_cm": 1.13,
    "line_spacing_pt": 28.0,
}


def load_json(path: Path) -> dict[str, Any]:
    """读取 JSON 文件。"""
    return json.loads(path.read_text(encoding="utf-8"))


def line_spacing_pt(paragraph: dict[str, Any]) -> float | None:
    """将 snapshot 中的 twips 行距转为磅值。"""
    raw = (paragraph.get("resolved_paragraph_format") or {}).get("line_spacing_raw")
    if raw in (None, ""):
        return None
    try:
        return round(float(raw) / 20, 2)
    except ValueError:
        return None


def current_format(paragraph: dict[str, Any]) -> dict[str, Any]:
    """提取当前可写回格式。"""
    p_format = paragraph.get("resolved_paragraph_format") or {}
    r_format = paragraph.get("resolved_run_format") or {}
    return {
        "font_east_asia": r_format.get("font_east_asia"),
        "font_ascii": r_format.get("font_ascii"),
        "font_size_pt": r_format.get("font_size_pt"),
        "bold": r_format.get("bold"),
        "alignment": p_format.get("alignment"),
        "first_line_indent_cm": p_format.get("first_line_indent_cm"),
        "line_spacing_pt": line_spacing_pt(paragraph),
    }


def differs(current: dict[str, Any], expected: dict[str, Any]) -> bool:
    """判断当前格式是否偏离期望格式。"""
    for key, expected_value in expected.items():
        current_value = current.get(key)
        if isinstance(expected_value, float):
            if current_value is None or abs(float(current_value) - expected_value) > 0.02:
                return True
            continue
        if current_value != expected_value:
            return True
    return False


def table_paragraph_indices(snapshot: dict[str, Any]) -> set[int]:
    """收集表格内段落下标，用于避免正文规则误改表格。"""
    indices: set[int] = set()
    for table in snapshot.get("tables", []):
        for cell in table.get("cells", []):
            for paragraph in cell.get("paragraphs", []):
                index = paragraph.get("paragraph_index")
                if isinstance(index, int):
                    indices.add(index)
    return indices


def role_for(paragraph: dict[str, Any], title_index: int, table_indices: set[int]) -> str | None:
    """识别已确认规则覆盖的段落角色。"""
    index = paragraph.get("paragraph_index")
    text = (paragraph.get("text_preview") or "").strip()
    if not text or index in table_indices:
        return None
    if index == title_index:
        return "document-title"
    if index in {2, 3}:
        return None
    if LEVEL1_PATTERN.match(text):
        return "level-1-heading"
    if isinstance(index, int) and index > 3:
        return "body-paragraph"
    return None


def expected_for(role: str) -> dict[str, Any]:
    """返回角色对应的期望格式。"""
    if role == "document-title":
        return TITLE_RULE
    if role == "level-1-heading":
        return LEVEL1_RULE
    return BODY_RULE


def build_item(paragraph: dict[str, Any], role: str, index: int) -> dict[str, Any]:
    """构造单段落语义审计项。"""
    before = current_format(paragraph)
    after = expected_for(role)
    text = paragraph.get("text_preview") or ""
    return {
        "issue_id": f"CMR-{index:04d}",
        "element_id": paragraph["element_id"],
        "semantic_role": role,
        "current_problem": f"{role} 格式与已确认标准规则不一致",
        "expected_role": role,
        "confidence": 0.95,
        "evidence": [
            "用户已确认：按标准文档统一标题、正文、一级标题；表格只审计",
            f"段落摘录：{text[:80]}",
            f"当前格式：{before}",
            f"期望格式：{after}",
        ],
        "recommended_action": {
            "action_type": "apply_body_direct_format",
            "auto_fix_policy": "auto-fix",
            "before": before,
            "after": after,
        },
        "risk_level": "low",
    }


def build_table_item(table: dict[str, Any], index: int) -> dict[str, Any]:
    """构造表格只审计项。"""
    return {
        "issue_id": f"CMR-T{index:03d}",
        "element_id": table.get("table_id") or f"table-{index:04d}",
        "semantic_role": "table",
        "current_problem": "待调整文档存在表格，但标准文档未提供表格规则",
        "expected_role": "table",
        "confidence": 0.9,
        "evidence": [
            "用户已确认：表格只审计",
            f"表格行数：{table.get('row_count')}",
            f"表格列数：{table.get('column_count')}",
        ],
        "recommended_action": {
            "action_type": "apply_table_cell_format",
            "auto_fix_policy": "audit-only",
            "before": {"table_id": table.get("table_id")},
            "after": {},
        },
        "risk_level": "medium",
    }


def build_audit(source_snapshot: str, snapshot: dict[str, Any], rule_profile_id: str) -> dict[str, Any]:
    """生成整份文档的语义审计。"""
    table_indices = table_paragraph_indices(snapshot)
    first_text = next(item for item in snapshot.get("paragraphs", []) if item.get("text_preview"))
    title_index = first_text["paragraph_index"]
    items: list[dict[str, Any]] = []
    for paragraph in snapshot.get("paragraphs", []):
        role = role_for(paragraph, title_index, table_indices)
        if role is None:
            continue
        before = current_format(paragraph)
        after = expected_for(role)
        if differs(before, after):
            items.append(build_item(paragraph, role, len(items) + 1))
    for table in snapshot.get("tables", []):
        items.append(build_table_item(table, len(items) + 1))
    return {
        "schema_version": "1.0.0",
        "source_snapshot": source_snapshot.replace("\\", "/"),
        "rule_profile_id": rule_profile_id,
        "generated_by": "codex",
        "generated_at": datetime.now(TZ).isoformat(),
        "items": items,
    }


def main() -> int:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="按已确认会议方案规则生成 semantic_audit.json")
    parser.add_argument("--issue-snapshot", required=True, type=Path)
    parser.add_argument("--rule-profile-id", default="meeting-plan-standard-20260306")
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()

    snapshot = load_json(args.issue_snapshot)
    audit = build_audit(str(args.issue_snapshot), snapshot, args.rule_profile_id)
    errors = validate_semantic_audit(audit)
    if errors:
        for error in errors:
            print(error)
        return 1
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(audit, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
