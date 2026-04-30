"""repair_plan.yaml 的轻量校验工具。"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import sys

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.utils.simple_yaml import load_yaml


REQUIRED_ROOT_FIELDS = {
    "schema_version",
    "repair_plan_id",
    "created_at",
    "based_on_snapshot",
    "rule_profile",
    "source_docx",
    "working_docx",
    "output_docx",
    "actions",
    "manual_review_items",
    "execution_order",
    "post_repair",
}
REQUIRED_ACTION_FIELDS = {
    "action_id",
    "source_issue_ids",
    "action_type",
    "target",
    "confidence",
    "semantic_evidence",
    "auto_fix_policy",
    "risk_level",
    "status",
}
WHITELIST_ACTIONS = {
    "map_heading_native_style",
    "apply_body_style_definition",
    "apply_body_direct_format",
    "apply_table_cell_format",
    "apply_table_border",
    "toc_content_audit",
    "insert_or_replace_toc_field",
}
ALLOWED_RISK_LEVELS = {"low", "medium", "high"}
ALLOWED_POLICIES = {"auto-fix", "manual-review", "audit-only"}


def validate_repair_plan(data: dict[str, Any]) -> list[str]:
    """返回 repair_plan 的校验错误列表。"""
    errors: list[str] = []
    missing_root = sorted(REQUIRED_ROOT_FIELDS - data.keys())
    if missing_root:
        errors.append(f"缺少根字段: {', '.join(missing_root)}")
    if data.get("schema_version") != "1.0.0":
        errors.append("schema_version 必须为 1.0.0")

    rule_profile = data.get("rule_profile")
    if not isinstance(rule_profile, dict) or not rule_profile.get("id") or not rule_profile.get("version"):
        errors.append("rule_profile 必须包含 id 和 version")

    if data.get("source_docx") == data.get("output_docx"):
        errors.append("output_docx 不得覆盖 source_docx")
    if data.get("working_docx") == data.get("output_docx"):
        errors.append("output_docx 不得覆盖 working_docx")

    actions = data.get("actions")
    if not isinstance(actions, list):
        errors.append("actions 必须是列表")
        return errors

    seen_action_ids: set[str] = set()
    for index, action in enumerate(actions, start=1):
        if not isinstance(action, dict):
            errors.append(f"actions[{index}] 必须是对象")
            continue
        action_id = str(action.get("action_id") or f"actions[{index}]")
        if action_id in seen_action_ids:
            errors.append(f"{action_id} 重复")
        seen_action_ids.add(action_id)

        missing_action = sorted(REQUIRED_ACTION_FIELDS - action.keys())
        if missing_action:
            errors.append(f"{action_id} 缺少字段: {', '.join(missing_action)}")

        source_issue_ids = action.get("source_issue_ids")
        if not isinstance(source_issue_ids, list) or not source_issue_ids:
            errors.append(f"{action_id} source_issue_ids 必须是非空列表")

        target = action.get("target")
        if not isinstance(target, dict) or not target.get("element_id"):
            errors.append(f"{action_id} target.element_id 不能为空")

        confidence = action.get("confidence")
        if not isinstance(confidence, (int, float)) or confidence < 0 or confidence > 1:
            errors.append(f"{action_id} confidence 必须在 0 到 1 之间")
            continue

        evidence = action.get("semantic_evidence")
        if not isinstance(evidence, list) or not evidence or any(not str(value).strip() for value in evidence):
            errors.append(f"{action_id} semantic_evidence 必须至少包含一条非空证据")

        action_type = action.get("action_type")
        policy = action.get("auto_fix_policy")
        risk_level = action.get("risk_level")
        if action_type not in WHITELIST_ACTIONS and policy == "auto-fix":
            errors.append(f"{action_id} 未在白名单内，不得 auto-fix")
        if policy not in ALLOWED_POLICIES:
            errors.append(f"{action_id} auto_fix_policy 非法")
        if risk_level not in ALLOWED_RISK_LEVELS:
            errors.append(f"{action_id} risk_level 非法")
        if policy == "auto-fix" and confidence < 0.85:
            errors.append(f"{action_id} confidence < 0.85 时不得 auto-fix")
        if policy == "auto-fix" and risk_level == "high":
            errors.append(f"{action_id} risk_level=high 时不得 auto-fix")

    manual_items = data.get("manual_review_items")
    if not isinstance(manual_items, list):
        errors.append("manual_review_items 必须是列表")

    return errors


def validate_file(path: Path) -> list[str]:
    """读取并校验 repair_plan 文件。"""
    data = load_yaml(path)
    if not isinstance(data, dict):
        return ["repair_plan 根节点必须是对象"]
    return validate_repair_plan(data)


def main() -> int:
    """命令行入口。"""
    import argparse

    parser = argparse.ArgumentParser(description="校验 repair_plan.yaml")
    parser.add_argument("path", type=Path)
    args = parser.parse_args()
    errors = validate_file(args.path)
    if errors:
        for error in errors:
            print(error)
        return 1
    print("repair_plan 校验通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
