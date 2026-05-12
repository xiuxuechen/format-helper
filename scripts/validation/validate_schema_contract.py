"""v4 schema inventory 最小契约 validator（CODE-005）。

该模块只负责 CODE-005 的横向 schema inventory 准出校验，不替代后续
CODE-006 之后的业务级 validator。字段级权威仍以 41_SCHEMA_CONTRACTS.md 为准。

覆盖：
- 41-§2 canonical schema_id 与 schema 文件一致
- 41-§12 / 50-§3.5 examples 的 required 字段、const、enum
- SCHEMA_MIN_STRATEGY.md §3.1 的 semver、canonical alias、unknown enum blocking
- CODE-005 path_negative 示例的路径逃逸阻断
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[2]
SCHEMA_DIR = ROOT / "docs" / "v4" / "schemas"

SUPPORTED_SCHEMA_VERSION = "1.0.0"
CONTRACT_VERSION = "v4"
SEMVER_PATTERN = re.compile(r"^(\d+)\.(\d+)\.(\d+)$")

CANONICAL_ALIASES = {
    "state": "run-state",
    "repair-plan-draft": "repair-plan",
    "repair-plan-finalized": "repair-plan",
}

CODE005_MISSING_FIELD = "CODE005-MISSING-FIELD"
CODE005_CONST_MISMATCH = "CODE005-CONST-MISMATCH"
CODE005_INVALID_ENUM = "CODE005-INVALID-ENUM"
CODE005_SEMVER_INCOMPATIBLE = "CODE005-SEMVER-INCOMPATIBLE"
CODE005_PATH_ESCAPE = "CODE005-PATH-ESCAPE"
CODE005_SCHEMA_NOT_FOUND = "CODE005-SCHEMA-NOT-FOUND"

MISSING = object()


@dataclass
class ValidationResult:
    """最小契约校验结果。"""

    valid: bool = True
    errors: list[dict[str, str]] = field(default_factory=list)
    warnings: list[dict[str, str]] = field(default_factory=list)

    def add_error(self, code: str, field: str, message: str) -> None:
        """追加阻断错误。"""
        self.valid = False
        self.errors.append({"code": code, "field": field, "message": message})

    def add_warning(self, code: str, field: str, message: str) -> None:
        """追加兼容警告。"""
        self.warnings.append({"code": code, "field": field, "message": message})


def parse_semver(version: str) -> tuple[int, int, int] | None:
    """解析 semver 字符串。"""
    match = SEMVER_PATTERN.match(version)
    if match is None:
        return None
    return tuple(int(part) for part in match.groups())


def semver_compatible(actual: str, supported: str = SUPPORTED_SCHEMA_VERSION) -> bool:
    """检查 semver 是否兼容。major 不一致即不兼容。"""
    parsed_actual = parse_semver(actual)
    parsed_supported = parse_semver(supported)
    if parsed_actual is None or parsed_supported is None:
        return False
    return parsed_actual[0] == parsed_supported[0]


def canonical_schema_id(schema_id: str | None, result: ValidationResult) -> str | None:
    """将历史 schema_id 映射为 canonical id。"""
    if schema_id in CANONICAL_ALIASES:
        canonical = CANONICAL_ALIASES[schema_id]
        result.add_warning(
            "CODE005-CANONICAL-ALIAS",
            "schema_id",
            f"历史 schema_id='{schema_id}' 兼容读取；写回时应升级为 '{canonical}'",
        )
        return canonical
    return schema_id


def load_schema(schema_id: str) -> dict[str, Any] | None:
    """读取 CODE-005 最小 schema 定义。"""
    schema_path = SCHEMA_DIR / f"{schema_id}.schema.json"
    if not schema_path.exists():
        return None
    return json.loads(schema_path.read_text(encoding="utf-8"))


def get_nested(data: dict[str, Any], field_path: str) -> Any:
    """读取点号路径字段。"""
    current: Any = data
    for part in field_path.split("."):
        if not isinstance(current, dict) or part not in current:
            return MISSING
        current = current[part]
    return current


def check_path_policy(data: dict[str, Any], result: ValidationResult) -> None:
    """校验 CODE-005 path_negative 示例的路径逃逸信号。"""
    policy = data.get("path_policy")
    if not isinstance(policy, dict):
        return
    checked_path = str(policy.get("checked_path", ""))
    path_valid = policy.get("path_valid")
    normalized = checked_path.replace("\\", "/")
    if path_valid is False or normalized.startswith("../") or "/../" in normalized:
        result.add_error(
            CODE005_PATH_ESCAPE,
            "path_policy.checked_path",
            f"路径策略示例必须阻断逃逸路径：{checked_path}",
        )


def validate_schema_contract(data: dict[str, Any], expected_schema_id: str | None = None) -> ValidationResult:
    """按 CODE-005 最小契约校验任意 schema example。"""
    result = ValidationResult()
    raw_schema_id = data.get("schema_id")
    schema_id = canonical_schema_id(raw_schema_id, result)

    if expected_schema_id is not None and schema_id != expected_schema_id:
        result.add_error(
            CODE005_CONST_MISMATCH,
            "schema_id",
            f"schema_id 必须为 '{expected_schema_id}'，实际为 '{raw_schema_id}'",
        )
        return result

    if schema_id is None:
        result.add_error(CODE005_MISSING_FIELD, "schema_id", "缺少 required 字段: schema_id")
        return result

    schema = load_schema(schema_id)
    if schema is None:
        result.add_error(
            CODE005_SCHEMA_NOT_FOUND,
            "schema_id",
            f"未找到 canonical schema 定义：docs/v4/schemas/{schema_id}.schema.json",
        )
        return result

    for field_name in schema.get("required", []):
        if get_nested(data, field_name) is MISSING:
            result.add_error(
                CODE005_MISSING_FIELD,
                field_name,
                f"缺少 required 字段: {field_name}",
            )

    for field_name, definition in schema.get("properties", {}).items():
        value = get_nested(data, field_name)
        comparable_value = schema_id if field_name == "schema_id" else value
        if "const" in definition and comparable_value != definition["const"]:
            result.add_error(
                CODE005_CONST_MISMATCH,
                field_name,
                f"{field_name} 必须为 '{definition['const']}'，实际为 '{value}'",
            )
        if "enum" in definition and value is not MISSING and value is not None and value not in definition["enum"]:
            result.add_error(
                CODE005_INVALID_ENUM,
                field_name,
                f"{field_name}='{value}' 不在允许集合 {definition['enum']} 中",
            )

    version = data.get("schema_version")
    if not isinstance(version, str) or not semver_compatible(version):
        result.add_error(
            CODE005_SEMVER_INCOMPATIBLE,
            "schema_version",
            f"schema_version='{version}' 与支持版本 '{SUPPORTED_SCHEMA_VERSION}' 不兼容",
        )

    if data.get("contract_version") != CONTRACT_VERSION:
        result.add_error(
            CODE005_CONST_MISMATCH,
            "contract_version",
            f"contract_version 必须为 '{CONTRACT_VERSION}'",
        )

    check_path_policy(data, result)
    return result
