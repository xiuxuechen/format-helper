"""review_result JSON 的轻量校验工具。"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


REQUIRED_FIELDS = {"schema_version", "task_id", "task_name", "status", "checked_at", "evidence", "issues"}
STATUSES = {"passed", "passed_with_warnings", "blocked"}


def load_json(path: Path) -> dict[str, Any]:
    """读取 JSON 文件。"""
    return json.loads(path.read_text(encoding="utf-8"))


def validate_review_result(data: dict[str, Any]) -> list[str]:
    """返回 review_result 的校验错误列表。"""
    errors: list[str] = []
    missing = sorted(REQUIRED_FIELDS - data.keys())
    if missing:
        errors.append(f"缺少字段: {', '.join(missing)}")
    if data.get("schema_version") != "1.0.0":
        errors.append("schema_version 必须为 1.0.0")
    task_id = data.get("task_id")
    if task_id not in {"T01", "T02", "T03", "T04", "T05", "T06"}:
        errors.append("task_id 必须为 T01-T06")
    if data.get("status") not in STATUSES:
        errors.append("status 非法")
    evidence = data.get("evidence")
    if not isinstance(evidence, list) or not evidence:
        errors.append("evidence 必须是非空列表")
    issues = data.get("issues")
    if not isinstance(issues, list):
        errors.append("issues 必须是列表")
    elif data.get("status") == "blocked" and not issues:
        errors.append("blocked 状态必须包含 issues")
    return errors


def validate_file(path: Path) -> list[str]:
    """读取并校验 review_result 文件。"""
    return validate_review_result(load_json(path))


def main() -> int:
    """命令行入口。"""
    import argparse

    parser = argparse.ArgumentParser(description="校验 review_result JSON")
    parser.add_argument("path", type=Path)
    args = parser.parse_args()
    errors = validate_file(args.path)
    if errors:
        for error in errors:
            print(error)
        return 1
    print("review_result 校验通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
