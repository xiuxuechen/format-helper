#!/usr/bin/env python3
"""从 semantic_audit.json 生成可追溯 repair_plan.yaml。"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[4]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.utils.simple_yaml import write_yaml
from scripts.validation.check_semantic_audit import validate_semantic_audit
from scripts.validation.check_repair_plan import WHITELIST_ACTIONS, validate_repair_plan


TZ = timezone(timedelta(hours=8))
FORMAT_STRATEGY = {
    "map_heading_native_style": "style-definition",
    "apply_body_style_definition": "style-definition",
    "apply_body_direct_format": "direct-format-override",
    "apply_table_cell_format": "direct-format-override",
    "apply_table_border": "direct-format-override",
    "toc_content_audit": "audit-only",
    "insert_or_replace_toc_field": "toc-field",
}
EXECUTION_ORDER = [
    "normalize_styles",
    "apply_page_section_rules",
    "apply_heading_styles",
    "apply_body_styles",
    "apply_table_safe_fixes",
    "toc_content_audit",
    "replace_or_insert_auto_toc",
    "refresh_fields_or_mark_for_update",
    "save_repaired_docx",
]


def load_json(path: Path) -> dict[str, Any]:
    """读取 JSON 文件。"""
    return json.loads(path.read_text(encoding="utf-8"))


def canonical_output_docx(source_docx: Path, requested_output: Path, now: datetime) -> Path:
    """生成符合运行目录约定的输出 Word 路径。"""
    if requested_output.suffix.lower() == ".docx":
        return requested_output
    timestamp = now.strftime("%Y%m%d%H%M")
    return requested_output / f"{source_docx.stem}{timestamp}.docx"


def normalize_path(path: Path | str) -> str:
    """将路径转为稳定斜杠格式。"""
    return str(path).replace("\\", "/")


def candidate_policy(item: dict[str, Any]) -> tuple[str, list[str]]:
    """判断候选动作是否允许自动修复。"""
    reasons: list[str] = []
    confidence = item.get("confidence")
    risk_level = item.get("risk_level")
    action = item.get("recommended_action") or {}
    action_type = action.get("action_type")
    requested_policy = action.get("auto_fix_policy")

    if requested_policy != "auto-fix":
        reasons.append(f"推荐策略为 {requested_policy or '空'}")
    if action_type not in WHITELIST_ACTIONS:
        reasons.append(f"动作 {action_type or '空'} 不在白名单内")
    if not isinstance(confidence, (int, float)) or confidence < 0.85:
        reasons.append("confidence 低于 0.85")
    if risk_level == "high":
        reasons.append("risk_level 为 high")
    if not item.get("evidence"):
        reasons.append("缺少语义证据")
    if not item.get("issue_id"):
        reasons.append("缺少 source issue")
    return ("manual-review" if reasons else "auto-fix", reasons)


def build_action(item: dict[str, Any], index: int) -> tuple[dict[str, Any], dict[str, Any] | None]:
    """构造修复动作和可选人工确认项。"""
    recommended = item.get("recommended_action") or {}
    action_type = recommended.get("action_type") or "manual_review"
    policy, reasons = candidate_policy(item)
    action = {
        "action_id": f"A{index:03d}",
        "source_issue_ids": [item["issue_id"], *item.get("format_issue_ids", [])],
        "action_type": action_type,
        "format_write_strategy": FORMAT_STRATEGY.get(action_type, "manual-review"),
        "target": {
            "element_id": item.get("element_id"),
            "expected_role": item.get("expected_role"),
        },
        "confidence": item.get("confidence"),
        "semantic_evidence": item.get("evidence") or [],
        "before": recommended.get("before", item.get("before", {})),
        "after": recommended.get("after", item.get("after", {})),
        "auto_fix_policy": policy,
        "risk_level": item.get("risk_level"),
        "status": "pending",
    }
    if policy == "auto-fix":
        return action, None

    manual_item = {
        "item_id": f"M{index:03d}",
        "source_issue_ids": action["source_issue_ids"],
        "element_ref": {
            "element_id": item.get("element_id"),
            "expected_role": item.get("expected_role"),
        },
        "reason": "；".join(reasons) if reasons else item.get("current_problem", "需要人工确认"),
        "required_decision": f"是否允许执行 {action_type}",
        "default_option": "manual-review",
    }
    return action, manual_item


def build_repair_plan(
    semantic_audit: dict[str, Any],
    source_docx: Path,
    working_docx: Path,
    output_docx: Path,
    snapshot: str,
    rule_id: str,
    rule_version: str,
    now: datetime,
) -> dict[str, Any]:
    """生成 repair_plan 数据结构。"""
    actions: list[dict[str, Any]] = []
    manual_items: list[dict[str, Any]] = []
    for index, item in enumerate(semantic_audit.get("items", []), start=1):
        action, manual_item = build_action(item, index)
        actions.append(action)
        if manual_item:
            manual_items.append(manual_item)

    return {
        "schema_version": "1.0.0",
        "repair_plan_id": f"docx-repair-plan-{now.strftime('%Y%m%d-%H%M%S')}",
        "created_at": now.isoformat(),
        "based_on_snapshot": snapshot,
        "rule_profile": {
            "id": rule_id,
            "version": rule_version,
        },
        "source_docx": normalize_path(source_docx),
        "working_docx": normalize_path(working_docx),
        "output_docx": normalize_path(output_docx),
        "conflict_resolution": [],
        "actions": actions,
        "manual_review_items": manual_items,
        "execution_order": EXECUTION_ORDER,
        "post_repair": {
            "generate_after_snapshot": True,
            "dispatch_second_round_review": True,
            "required_review_task_ids": ["T01", "T02", "T03", "T04", "T05", "T06"],
        },
    }


def main_from_args(argv: list[str] | None = None) -> int:
    """命令行入口，便于测试复用。"""
    parser = argparse.ArgumentParser(description="生成 repair_plan.yaml")
    parser.add_argument("--semantic-audit", required=True, type=Path)
    parser.add_argument("--format-audit", type=Path)
    parser.add_argument("--risk-policy", type=Path)
    parser.add_argument("--snapshot", required=True)
    parser.add_argument("--rule-id", required=True)
    parser.add_argument("--rule-version", default="1.0.0")
    parser.add_argument("--source-docx", required=True, type=Path)
    parser.add_argument("--working-docx", required=True, type=Path)
    parser.add_argument("--output-docx", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args(argv)

    semantic_audit = load_json(args.semantic_audit)
    semantic_errors = validate_semantic_audit(semantic_audit)
    if semantic_errors:
        for error in semantic_errors:
            print(error)
        return 1

    now = datetime.now(TZ)
    output_docx = canonical_output_docx(args.source_docx, args.output_docx, now)
    plan = build_repair_plan(
        semantic_audit=semantic_audit,
        source_docx=args.source_docx,
        working_docx=args.working_docx,
        output_docx=output_docx,
        snapshot=args.snapshot,
        rule_id=args.rule_id,
        rule_version=args.rule_version,
        now=now,
    )
    plan_errors = validate_repair_plan(plan)
    if plan_errors:
        for error in plan_errors:
            print(error)
        return 1

    write_yaml(args.output, plan)
    print(args.output)
    return 0


def main() -> int:
    """脚本入口。"""
    return main_from_args()


if __name__ == "__main__":
    raise SystemExit(main())
