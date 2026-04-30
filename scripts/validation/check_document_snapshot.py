"""document_snapshot.json 的轻量校验工具。"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


REQUIRED_ROOT_FIELDS = {
    "schema_version",
    "snapshot_kind",
    "source_docx",
    "document_hash",
    "created_at",
    "paragraph_count",
    "table_count",
    "section_count",
    "paragraphs",
    "tables",
    "sections",
}
SNAPSHOT_KINDS = {"standard", "before", "after"}


def load_json(path: Path) -> dict[str, Any]:
    """读取 JSON 文件。"""
    return json.loads(path.read_text(encoding="utf-8"))


def validate_document_snapshot(data: dict[str, Any]) -> list[str]:
    """返回 document_snapshot 的校验错误列表。"""
    errors: list[str] = []
    missing_root = sorted(REQUIRED_ROOT_FIELDS - data.keys())
    if missing_root:
        errors.append(f"缺少根字段: {', '.join(missing_root)}")
    if data.get("schema_version") != "1.0.0":
        errors.append("schema_version 必须为 1.0.0")
    if data.get("snapshot_kind") not in SNAPSHOT_KINDS:
        errors.append("snapshot_kind 非法")
    for field in ("paragraph_count", "table_count", "section_count"):
        value = data.get(field)
        if not isinstance(value, int) or value < 0:
            errors.append(f"{field} 必须是非负整数")

    paragraphs = data.get("paragraphs")
    if not isinstance(paragraphs, list):
        errors.append("paragraphs 必须是列表")
    elif len(paragraphs) != data.get("paragraph_count"):
        errors.append("paragraphs 数量必须等于 paragraph_count")
    else:
        for index, paragraph in enumerate(paragraphs, start=1):
            if not isinstance(paragraph, dict) or not paragraph.get("element_id"):
                errors.append(f"paragraphs[{index}] 缺少 element_id")

    tables = data.get("tables")
    if not isinstance(tables, list):
        errors.append("tables 必须是列表")
    elif len(tables) != data.get("table_count"):
        errors.append("tables 数量必须等于 table_count")

    sections = data.get("sections")
    if not isinstance(sections, list) or len(sections) != data.get("section_count"):
        errors.append("sections 数量必须等于 section_count")
    return errors


def validate_file(path: Path) -> list[str]:
    """读取并校验 document_snapshot 文件。"""
    return validate_document_snapshot(load_json(path))


def main() -> int:
    """命令行入口。"""
    import argparse

    parser = argparse.ArgumentParser(description="校验 document_snapshot.json")
    parser.add_argument("path", type=Path)
    args = parser.parse_args()
    errors = validate_file(args.path)
    if errors:
        for error in errors:
            print(error)
        return 1
    print("document_snapshot 校验通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
