"""semantic_rule_draft.json 的轻量校验工具。"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


REQUIRED_ROOT_FIELDS = {
    "schema_version",
    "run_id",
    "created_at",
    "rule_id",
    "source_snapshot",
    "source",
    "document_type",
    "roles",
    "manual_confirmation",
    "validation",
}
REQUIRED_ROLE_FIELDS = {
    "role",
    "description",
    "evidence",
    "confidence",
    "format",
    "write_strategy",
    "requires_user_confirmation",
}
WRITE_STRATEGIES = {"style-definition", "direct-format", "audit-only"}


def load_json(path: Path) -> dict[str, Any]:
    """读取 JSON 文件。"""
    return json.loads(path.read_text(encoding="utf-8"))


def validate_semantic_rule_draft(data: dict[str, Any]) -> list[str]:
    """返回 semantic_rule_draft 的校验错误列表。"""
    errors: list[str] = []
    missing_root = sorted(REQUIRED_ROOT_FIELDS - data.keys())
    if missing_root:
        errors.append(f"缺少根字段: {', '.join(missing_root)}")

    if data.get("schema_version") != "1.0.0":
        errors.append("schema_version 必须为 1.0.0")

    roles = data.get("roles")
    if not isinstance(roles, list) or not roles:
        errors.append("roles 必须是非空列表")
        return errors

    for index, role in enumerate(roles, start=1):
        if not isinstance(role, dict):
            errors.append(f"roles[{index}] 必须是对象")
            continue
        role_name = role.get("role", f"roles[{index}]")
        missing_role = sorted(REQUIRED_ROLE_FIELDS - role.keys())
        if missing_role:
            errors.append(f"{role_name} 缺少字段: {', '.join(missing_role)}")

        evidence = role.get("evidence")
        if not isinstance(evidence, list) or not evidence or any(not str(item).strip() for item in evidence):
            errors.append(f"{role_name} evidence 必须至少包含一条非空证据")

        confidence = role.get("confidence")
        if not isinstance(confidence, (int, float)) or confidence < 0 or confidence > 1:
            errors.append(f"{role_name} confidence 必须在 0 到 1 之间")
            continue

        if role.get("write_strategy") not in WRITE_STRATEGIES:
            errors.append(f"{role_name} write_strategy 非法")

        requires_confirmation = role.get("requires_user_confirmation")
        if confidence < 0.85 and requires_confirmation is not True:
            errors.append(f"{role_name} confidence < 0.85 时必须 requires_user_confirmation=true")
        if confidence < 0.85 and not role.get("manual_confirmation_reason"):
            errors.append(f"{role_name} confidence < 0.85 时必须提供 manual_confirmation_reason")
        if role.get("write_strategy") == "audit-only" and requires_confirmation is not True:
            errors.append(f"{role_name} audit-only 必须进入人工确认")

    return errors


def validate_file(path: Path) -> list[str]:
    """读取并校验 semantic_rule_draft 文件。"""
    return validate_semantic_rule_draft(load_json(path))


def main() -> int:
    """命令行入口。"""
    import argparse

    parser = argparse.ArgumentParser(description="校验 semantic_rule_draft.json")
    parser.add_argument("path", type=Path)
    args = parser.parse_args()
    errors = validate_file(args.path)
    if errors:
        for error in errors:
            print(error)
        return 1
    print("semantic_rule_draft 校验通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
