"""skill-result 最小 validator（CODE-005 / CODE-006A 核心闭环）。

实现 41_SCHEMA_CONTRACTS.md §5 的最小 validator，覆盖：
- Required 字段检查（按 SCHEMA_MIN_STRATEGY.md §3.2.1）
- Enum 闭合校验（closed enum）
- Nullable-but-required 逻辑
- Semver 兼容校验
- Canonical alias 映射
- Unknown enum blocking

参考：
- 41-§5.1: Required 字段
- 41-§5.3: Nullable 字段
- 41-§6: 状态组合规则
- 41-§7: Semver 兼容算法
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any


# 错误码
SR_MISSING_FIELD = "SR-MISSING-FIELD"
SR_INVALID_ENUM = "SR-INVALID-ENUM"
SR_NULLABLE_VIOLATION = "SR-NULLABLE-VIOLATION"
SR_SEMVER_INCOMPATIBLE = "SR-SEMVER-INCOMPATIBLE"
SR_CONSISTENCY_ERROR = "SR-CONSISTENCY-ERROR"
SR_SCHEMA_ID_MISMATCH = "SR-SCHEMA-ID-MISMATCH"


# v4 支持的 schema_version（simple semver: MAJOR.MINOR.PATCH）
SUPPORTED_SCHEMA_VERSION = "1.0.0"
CONTRACT_VERSION = "v4"

# Required 字段最小子集（参考 41-§5.1, SCHEMA_MIN_STRATEGY.md §3.2.1）
REQUIRED_FIELDS = [
    "schema_id",
    "schema_version",
    "contract_version",
    "result_id",
    "run_id",
    "order",
    "attempt",
    "idempotency_key",
    "stage",
    "status",
    "schema_valid",
    "gate_passed",
    "gate_check",
    "validation",
    "artifacts",
    "next_action",
    "runtime",
]

# Enum 枚举（参考 41-§4）
ALLOWED_STATUS = {"done", "waiting_user", "blocked", "synthetic_failure"}
ALLOWED_STAGE = {
    "init", "rule_selection", "fact_extraction", "semantic_strategy",
    "rule_packaging", "format_audit", "repair_planning", "manual_review",
    "repair_execution", "after_snapshot", "review", "toc_acceptance",
    "final_acceptance", "reporting", "completed",
}
ALLOWED_GATE_CHECK_STATUS = {"passed", "passed_with_warnings", "failed", "not_applicable"}
ALLOWED_NEXT_ACTION_KIND = {"run_skill", "wait_user", "retry", "manual_recover", "finalize", "stop"}

# Canonical alias 映射
SCHEMA_ID_ALIAS = {
    "state": "run-state",
}

# semver 正则
SEMVER_PATTERN = re.compile(r"^(\d+)\.(\d+)\.(\d+)$")


@dataclass
class ValidationResult:
    """校验结果。"""

    valid: bool
    errors: list[dict] = field(default_factory=list)
    warnings: list[dict] = field(default_factory=list)

    def add_error(self, code: str, field_path: str, message: str) -> None:
        """添加错误。"""
        self.valid = False
        self.errors.append({"code": code, "field": field_path, "message": message})

    def add_warning(self, code: str, field_path: str, message: str) -> None:
        """添加警告（不影响 valid）。"""
        self.warnings.append({"code": code, "field": field_path, "message": message})


def parse_semver(version: str) -> tuple[int, int, int] | None:
    """解析 semver 版本字符串。"""
    match = SEMVER_PATTERN.match(version)
    if match is None:
        return None
    return tuple(int(x) for x in match.groups())


def semver_compatible(actual: str, supported: str) -> tuple[bool, str]:
    """semver 兼容性检查（参考 41-§7）。

    Returns:
        (compatible, reason): 是否兼容 + 兼容原因
    """
    actual_parsed = parse_semver(actual)
    supported_parsed = parse_semver(supported)

    if actual_parsed is None or supported_parsed is None:
        return (False, "invalid_semver")

    a_major, a_minor, a_patch = actual_parsed
    s_major, s_minor, s_patch = supported_parsed

    if a_major != s_major:
        return (False, "major_mismatch")
    if a_minor > s_minor:
        return (True, "compatible_with_warnings_newer_minor")
    if a_minor < s_minor:
        return (True, "compatible_older_minor")
    if a_patch != s_patch:
        return (True, "compatible_patch_difference")
    return (True, "exact")


def validate_skill_result(result: dict[str, Any]) -> ValidationResult:
    """校验 skill-result 最小字段集。

    Args:
        result: skill-result 对象（dict）

    Returns:
        ValidationResult: 校验结果
    """
    validation = ValidationResult(valid=True)

    # 0. Canonical alias 映射（historical schema_id）
    schema_id = result.get("schema_id")
    if schema_id in SCHEMA_ID_ALIAS:
        canonical = SCHEMA_ID_ALIAS[schema_id]
        validation.add_warning(
            "SR-SCHEMA-ALIAS-UPGRADE",
            "schema_id",
            f"历史 schema_id='{schema_id}' 已废弃，请升级为 canonical '{canonical}'",
        )
        schema_id = canonical

    # 1. schema_id 必须为 skill-result
    if schema_id != "skill-result":
        validation.add_error(
            SR_SCHEMA_ID_MISMATCH,
            "schema_id",
            f"schema_id 必须为 'skill-result'，实际为 '{schema_id}'",
        )
        return validation

    # 2. Required 字段检查
    for req_field in REQUIRED_FIELDS:
        if req_field not in result:
            validation.add_error(
                SR_MISSING_FIELD,
                req_field,
                f"缺少 required 字段: {req_field}",
            )

    if not validation.valid:
        return validation

    # 3. contract_version 必须为 v4
    if result.get("contract_version") != CONTRACT_VERSION:
        validation.add_error(
            SR_INVALID_ENUM,
            "contract_version",
            f"contract_version 必须为 '{CONTRACT_VERSION}'，实际为 '{result.get('contract_version')}'",
        )

    # 4. semver 兼容校验
    schema_version = result.get("schema_version", "")
    compatible, reason = semver_compatible(schema_version, SUPPORTED_SCHEMA_VERSION)
    if not compatible:
        validation.add_error(
            SR_SEMVER_INCOMPATIBLE,
            "schema_version",
            f"schema_version='{schema_version}' 与支持版本 '{SUPPORTED_SCHEMA_VERSION}' 不兼容：{reason}",
        )
    elif reason == "compatible_with_warnings_newer_minor":
        validation.add_warning(
            "SR-SEMVER-WARNING",
            "schema_version",
            f"schema_version='{schema_version}' 高于支持版本，可能包含未知字段",
        )

    # 5. Enum 闭合校验
    status = result.get("status")
    if status not in ALLOWED_STATUS:
        validation.add_error(
            SR_INVALID_ENUM,
            "status",
            f"status='{status}' 不在允许集合 {sorted(ALLOWED_STATUS)} 中",
        )

    stage = result.get("stage")
    if stage not in ALLOWED_STAGE:
        validation.add_error(
            SR_INVALID_ENUM,
            "stage",
            f"stage='{stage}' 不在允许集合中",
        )

    gate_check = result.get("gate_check", {})
    gate_status = gate_check.get("status") if isinstance(gate_check, dict) else None
    if gate_status is not None and gate_status not in ALLOWED_GATE_CHECK_STATUS:
        validation.add_error(
            SR_INVALID_ENUM,
            "gate_check.status",
            f"gate_check.status='{gate_status}' 不在允许集合 {sorted(ALLOWED_GATE_CHECK_STATUS)} 中",
        )

    next_action = result.get("next_action", {})
    next_kind = next_action.get("kind") if isinstance(next_action, dict) else None
    if next_kind is not None and next_kind not in ALLOWED_NEXT_ACTION_KIND:
        validation.add_error(
            SR_INVALID_ENUM,
            "next_action.kind",
            f"next_action.kind='{next_kind}' 不在允许集合 {sorted(ALLOWED_NEXT_ACTION_KIND)} 中",
        )

    # 6. Nullable-but-required 逻辑（参考 41-§5.3, §6）
    error_obj = result.get("error", {})
    error_code = error_obj.get("code") if isinstance(error_obj, dict) else None

    if status in {"blocked", "synthetic_failure"}:
        if not error_code:
            validation.add_error(
                SR_NULLABLE_VIOLATION,
                "error.code",
                f"status={status} 时 error.code 必须非 null",
            )

    # 7. 一致性约束（参考 41-§5.1）
    if result.get("schema_valid") != result.get("validation", {}).get("schema_valid"):
        validation.add_error(
            SR_CONSISTENCY_ERROR,
            "schema_valid",
            "schema_valid 必须等于 validation.schema_valid",
        )

    if result.get("gate_passed") != gate_check.get("passed"):
        validation.add_error(
            SR_CONSISTENCY_ERROR,
            "gate_passed",
            "gate_passed 必须等于 gate_check.passed",
        )

    # 8. next_action.kind=run_skill 时 stage 和 skill_name 必须非 null
    if next_kind == "run_skill":
        if not next_action.get("stage"):
            validation.add_error(
                SR_NULLABLE_VIOLATION,
                "next_action.stage",
                "next_action.kind=run_skill 时 stage 必须非 null",
            )
        if not next_action.get("skill_name"):
            validation.add_error(
                SR_NULLABLE_VIOLATION,
                "next_action.skill_name",
                "next_action.kind=run_skill 时 skill_name 必须非 null",
            )

    # 9. gate_check.passed 与 status 一致性
    gate_passed = result.get("gate_passed")
    if status == "done" and gate_passed is not True:
        validation.add_error(
            SR_CONSISTENCY_ERROR,
            "gate_passed",
            "status=done 时 gate_passed 必须为 true",
        )
    if status in {"blocked", "synthetic_failure"} and gate_passed is True:
        validation.add_error(
            SR_CONSISTENCY_ERROR,
            "gate_passed",
            f"status={status} 时 gate_passed 不得为 true",
        )

    return validation
