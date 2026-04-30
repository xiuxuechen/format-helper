"""semantic_audit.json 的轻量校验工具。"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


REQUIRED_ROOT_FIELDS = {
    "schema_version",
    "source_snapshot",
    "rule_profile_id",
    "generated_by",
    "generated_at",
    "items",
}
REQUIRED_ITEM_FIELDS = {
    "issue_id",
    "element_id",
    "semantic_role",
    "current_problem",
    "expected_role",
    "confidence",
    "evidence",
    "recommended_action",
    "risk_level",
}
ALLOWED_RISK_LEVELS = {"low", "medium", "high"}
ALLOWED_POLICIES = {"auto-fix", "manual-review", "audit-only"}


def load_json(path: Path) -> dict[str, Any]:
    """读取 JSON 文件。"""
    return json.loads(path.read_text(encoding="utf-8"))


def validate_semantic_audit(data: dict[str, Any]) -> list[str]:
    """返回 semantic_audit 的校验错误列表。"""
    errors: list[str] = []
    missing_root = sorted(REQUIRED_ROOT_FIELDS - data.keys())
    if missing_root:
        errors.append(f"缺少根字段: {', '.join(missing_root)}")
    if data.get("schema_version") != "1.0.0":
        errors.append("schema_version 必须为 1.0.0")
    if data.get("generated_by") != "codex":
        errors.append("generated_by 必须为 codex")

    items = data.get("items")
    if not isinstance(items, list) or not items:
        errors.append("items 必须是非空列表")
        return errors

    seen_issue_ids: set[str] = set()
    for index, item in enumerate(items, start=1):
        if not isinstance(item, dict):
            errors.append(f"items[{index}] 必须是对象")
            continue
        issue_id = str(item.get("issue_id") or f"items[{index}]")
        if issue_id in seen_issue_ids:
            errors.append(f"{issue_id} 重复")
        seen_issue_ids.add(issue_id)

        missing_item = sorted(REQUIRED_ITEM_FIELDS - item.keys())
        if missing_item:
            errors.append(f"{issue_id} 缺少字段: {', '.join(missing_item)}")

        confidence = item.get("confidence")
        if not isinstance(confidence, (int, float)) or confidence < 0 or confidence > 1:
            errors.append(f"{issue_id} confidence 必须在 0 到 1 之间")
            continue

        evidence = item.get("evidence")
        if not isinstance(evidence, list) or not evidence or any(not str(value).strip() for value in evidence):
            errors.append(f"{issue_id} evidence 必须至少包含一条非空证据")

        risk_level = item.get("risk_level")
        if risk_level not in ALLOWED_RISK_LEVELS:
            errors.append(f"{issue_id} risk_level 非法")

        action = item.get("recommended_action")
        if not isinstance(action, dict):
            errors.append(f"{issue_id} recommended_action 必须是对象")
            continue
        if not action.get("action_type"):
            errors.append(f"{issue_id} recommended_action.action_type 不能为空")
        policy = action.get("auto_fix_policy")
        if policy not in ALLOWED_POLICIES:
            errors.append(f"{issue_id} recommended_action.auto_fix_policy 非法")
        if policy == "auto-fix" and confidence < 0.85:
            errors.append(f"{issue_id} confidence < 0.85 时不得 auto-fix")
        if policy == "auto-fix" and risk_level == "high":
            errors.append(f"{issue_id} risk_level=high 时不得 auto-fix")

    return errors


def validate_file(path: Path) -> list[str]:
    """读取并校验 semantic_audit 文件。"""
    return validate_semantic_audit(load_json(path))


def main() -> int:
    """命令行入口。"""
    import argparse

    parser = argparse.ArgumentParser(description="校验 semantic_audit.json")
    parser.add_argument("path", type=Path)
    args = parser.parse_args()
    errors = validate_file(args.path)
    if errors:
        for error in errors:
            print(error)
        return 1
    print("semantic_audit 校验通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
