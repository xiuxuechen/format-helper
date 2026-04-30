"""final_acceptance.json 的轻量校验工具。"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


REQUIRED_FIELDS = {
    "schema_version",
    "accepted",
    "status",
    "created_at",
    "open_blockers",
    "manual_items_remaining",
    "output_docx_valid",
    "reports",
    "evidence",
}


def load_json(path: Path) -> dict[str, Any]:
    """读取 JSON 文件。"""
    return json.loads(path.read_text(encoding="utf-8"))


def validate_final_acceptance(data: dict[str, Any]) -> list[str]:
    """返回 final_acceptance 的校验错误列表。"""
    errors: list[str] = []
    missing = sorted(REQUIRED_FIELDS - data.keys())
    if missing:
        errors.append(f"缺少字段: {', '.join(missing)}")
    if data.get("schema_version") != "1.0.0":
        errors.append("schema_version 必须为 1.0.0")
    accepted = data.get("accepted")
    status = data.get("status")
    if not isinstance(accepted, bool):
        errors.append("accepted 必须是布尔值")
    if status not in {"accepted", "blocked"}:
        errors.append("status 非法")
    if accepted is True and status != "accepted":
        errors.append("accepted=true 时 status 必须为 accepted")
    if accepted is False and status != "blocked":
        errors.append("accepted=false 时 status 必须为 blocked")
    if accepted is True and data.get("open_blockers"):
        errors.append("accepted=true 时 open_blockers 必须为空")
    if not isinstance(data.get("output_docx_valid"), bool):
        errors.append("output_docx_valid 必须是布尔值")
    reports = data.get("reports")
    if not isinstance(reports, list) or not reports:
        errors.append("reports 必须是非空列表")
    evidence = data.get("evidence")
    if not isinstance(evidence, list) or not evidence:
        errors.append("evidence 必须是非空列表")
    return errors


def validate_file(path: Path) -> list[str]:
    """读取并校验 final_acceptance 文件。"""
    return validate_final_acceptance(load_json(path))


def main() -> int:
    """命令行入口。"""
    import argparse

    parser = argparse.ArgumentParser(description="校验 final_acceptance.json")
    parser.add_argument("path", type=Path)
    args = parser.parse_args()
    errors = validate_file(args.path)
    if errors:
        for error in errors:
            print(error)
        return 1
    print("final_acceptance 校验通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
