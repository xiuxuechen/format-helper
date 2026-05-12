"""run-state 最小 validator（CODE-006）。

实现 41_SCHEMA_CONTRACTS.md §11.1 的最小 validator，覆盖：
- Required 字段检查（按 SCHEMA_MIN_STRATEGY.md §3.2.2）
- Enum 闭合校验（mode、workflow_mode、stage、status）
- Nullable-but-required 逻辑
- Semver 兼容校验
- Canonical alias 映射（state → run-state）
- 一致性约束

参考：
- 41-§11.1: run-state 字段契约
- 41-§4: Enum 定义
- 41-§7: Semver 兼容算法
- SCHEMA_MIN_STRATEGY.md §3.2.2: run-state 最小字段集
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any


# 错误码
RS_MISSING_FIELD = "RS-MISSING-FIELD"
RS_INVALID_ENUM = "RS-INVALID-ENUM"
RS_NULLABLE_VIOLATION = "RS-NULLABLE-VIOLATION"
RS_SEMVER_INCOMPATIBLE = "RS-SEMVER-INCOMPATIBLE"
RS_CONSISTENCY_ERROR = "RS-CONSISTENCY-ERROR"
RS_SCHEMA_ID_MISMATCH = "RS-SCHEMA-ID-MISMATCH"


# v4 支持的 schema_version
SUPPORTED_SCHEMA_VERSION = "1.0.0"
CONTRACT_VERSION = "v4"

# Required 字段最小子集（参考 41-§11.1, SCHEMA_MIN_STRATEGY.md §3.2.2）
REQUIRED_FIELDS = [
    "schema_id",
    "schema_version",
    "contract_version",
    "run_id",
    "run_dir",
    "mode",
    "workflow_mode",
    "stage",
    "status",
    "input_docx",
    "rule_id",
    "rule_ref",
    "safe_outputs",
    "skill_results",
    "last_result_id",
    "applied_result_id",
    "result_chain_head",
    "evidence_manifest_path",
    "evidence_manifest_generations",
    "final_acceptance_path",
    "reporting_result_path",
    "blockers",
    "next_action",
    "updated_at",
]

# Enum 枚举（参考 41-§4）
ALLOWED_MODE = {"build_rules", "audit_only", "repair", "resume"}
ALLOWED_WORKFLOW_MODE = {"build_rules", "audit_only", "repair"}
ALLOWED_NEXT_ACTION_KIND = {"run_skill", "wait_user", "retry", "manual_recover", "finalize", "stop"}
ALLOWED_STAGE = {
    "init", "rule_selection", "fact_extraction", "semantic_strategy",
    "rule_packaging", "format_audit", "repair_planning", "manual_review",
    "repair_execution", "after_snapshot", "review", "toc_acceptance",
    "final_acceptance", "reporting", "completed",
}
ALLOWED_STATUS = {"wip", "waiting_user", "blocked", "accepted", "accepted_with_warnings"}
NEXT_ACTION_REQUIRED_FIELDS = [
    "kind",
    "stage",
    "skill_name",
    "target_result_id",
    "target_error_code",
    "source_result_id",
    "override_reason",
    "resume_from_stage",
    "idempotency_key",
    "planned_idempotency_key",
    "reason",
    "required_inputs",
    "user_message",
]

# Canonical alias 映射（参考 41-§2）
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


def validate_run_state(state: dict[str, Any]) -> ValidationResult:
    """校验 run-state 最小字段集。

    Args:
        state: run-state 对象（dict）

    Returns:
        ValidationResult: 校验结果
    """
    validation = ValidationResult(valid=True)

    # 0. Canonical alias 映射（historical schema_id）
    schema_id = state.get("schema_id")
    if schema_id in SCHEMA_ID_ALIAS:
        canonical = SCHEMA_ID_ALIAS[schema_id]
        validation.add_warning(
            "RS-SCHEMA-ALIAS-UPGRADE",
            "schema_id",
            f"历史 schema_id='{schema_id}' 已废弃，请升级为 canonical '{canonical}'",
        )
        schema_id = canonical

    # 1. schema_id 必须为 run-state
    if schema_id != "run-state":
        validation.add_error(
            RS_SCHEMA_ID_MISMATCH,
            "schema_id",
            f"schema_id 必须为 'run-state'，实际为 '{schema_id}'",
        )
        return validation

    # 2. Required 字段检查
    for req_field in REQUIRED_FIELDS:
        if req_field not in state:
            validation.add_error(
                RS_MISSING_FIELD,
                req_field,
                f"缺少 required 字段: {req_field}",
            )

    if not validation.valid:
        return validation

    # 3. contract_version 必须为 v4
    if state.get("contract_version") != CONTRACT_VERSION:
        validation.add_error(
            RS_INVALID_ENUM,
            "contract_version",
            f"contract_version 必须为 '{CONTRACT_VERSION}'，实际为 '{state.get('contract_version')}'",
        )

    # 4. semver 兼容校验
    schema_version = state.get("schema_version", "")
    compatible, reason = semver_compatible(schema_version, SUPPORTED_SCHEMA_VERSION)
    if not compatible:
        validation.add_error(
            RS_SEMVER_INCOMPATIBLE,
            "schema_version",
            f"schema_version='{schema_version}' 与支持版本 '{SUPPORTED_SCHEMA_VERSION}' 不兼容：{reason}",
        )
    elif reason == "compatible_with_warnings_newer_minor":
        validation.add_warning(
            "RS-SEMVER-WARNING",
            "schema_version",
            f"schema_version='{schema_version}' 高于支持版本，可能包含未知字段",
        )

    # 5. Enum 闭合校验
    mode = state.get("mode")
    if mode not in ALLOWED_MODE:
        validation.add_error(
            RS_INVALID_ENUM,
            "mode",
            f"mode='{mode}' 不在允许集合 {sorted(ALLOWED_MODE)} 中",
        )

    workflow_mode = state.get("workflow_mode")
    if workflow_mode and workflow_mode not in ALLOWED_WORKFLOW_MODE:
        validation.add_error(
            RS_INVALID_ENUM,
            "workflow_mode",
            f"workflow_mode='{workflow_mode}' 不在允许集合 {sorted(ALLOWED_WORKFLOW_MODE)} 中",
        )

    stage = state.get("stage")
    if stage not in ALLOWED_STAGE:
        validation.add_error(
            RS_INVALID_ENUM,
            "stage",
            f"stage='{stage}' 不在允许集合中",
        )

    status = state.get("status")
    if status not in ALLOWED_STATUS:
        validation.add_error(
            RS_INVALID_ENUM,
            "status",
            f"status='{status}' 不在允许集合 {sorted(ALLOWED_STATUS)} 中",
        )

    # 6. Nullable-but-required 逻辑（参考 41-§11.1）
    # rule_id 可为 null；当本次运行使用或生成规则包时必须非 null
    # 这里只做基本检查，业务逻辑由 Gate 层处理

    # 7. 一致性约束
    # final_acceptance_path 在 stage 早于 final_acceptance 时必须为 null
    final_acceptance_path = state.get("final_acceptance_path")
    if stage in {"init", "rule_selection", "fact_extraction", "semantic_strategy",
                 "rule_packaging", "format_audit", "repair_planning", "manual_review",
                 "repair_execution", "after_snapshot", "review", "toc_acceptance"}:
        # 早于 final_acceptance 阶段
        if final_acceptance_path is not None:
            validation.add_error(
                RS_CONSISTENCY_ERROR,
                "final_acceptance_path",
                f"stage='{stage}' 早于 final_acceptance，final_acceptance_path 必须为 null",
            )

    # status=accepted/accepted_with_warnings 时 final_acceptance_path 必须非 null
    if status in {"accepted", "accepted_with_warnings"}:
        if final_acceptance_path is None:
            validation.add_error(
                RS_CONSISTENCY_ERROR,
                "final_acceptance_path",
                f"status='{status}' 时 final_acceptance_path 必须非 null",
            )

    # 8. reporting_result_path 在 stage 早于 reporting 时必须为 null
    reporting_result_path = state.get("reporting_result_path")
    if stage in {"init", "rule_selection", "fact_extraction", "semantic_strategy",
                 "rule_packaging", "format_audit", "repair_planning", "manual_review",
                 "repair_execution", "after_snapshot", "review", "toc_acceptance",
                 "final_acceptance"}:
        # 早于 reporting 阶段
        if reporting_result_path is not None:
            validation.add_error(
                RS_CONSISTENCY_ERROR,
                "reporting_result_path",
                f"stage='{stage}' 早于 reporting，reporting_result_path 必须为 null",
            )

    # 9. next_action 必须包含必需字段
    next_action = state.get("next_action", {})
    if not isinstance(next_action, dict):
        validation.add_error(
            RS_CONSISTENCY_ERROR,
            "next_action",
            "next_action 必须是对象",
        )
    else:
        # next_action.kind=run_skill 时 stage 和 skill_name 必须非 null
        next_kind = next_action.get("kind")
        if next_kind not in ALLOWED_NEXT_ACTION_KIND:
            validation.add_error(
                RS_INVALID_ENUM,
                "next_action.kind",
                f"next_action.kind='{next_kind}' 不在允许集合 {sorted(ALLOWED_NEXT_ACTION_KIND)} 中",
            )
        for required_key in NEXT_ACTION_REQUIRED_FIELDS:
            if required_key not in next_action:
                validation.add_error(
                    RS_MISSING_FIELD,
                    f"next_action.{required_key}",
                    f"next_action 缺少 required 字段: {required_key}",
                )
        if next_kind == "run_skill":
            if not next_action.get("stage"):
                validation.add_error(
                    RS_NULLABLE_VIOLATION,
                    "next_action.stage",
                    "next_action.kind=run_skill 时 stage 必须非 null",
                )
            if not next_action.get("skill_name"):
                validation.add_error(
                    RS_NULLABLE_VIOLATION,
                    "next_action.skill_name",
                    "next_action.kind=run_skill 时 skill_name 必须非 null",
                )
            if not next_action.get("idempotency_key"):
                validation.add_error(
                    RS_NULLABLE_VIOLATION,
                    "next_action.idempotency_key",
                    "next_action.kind=run_skill 时 idempotency_key 必须非 null",
                )
            if not next_action.get("planned_idempotency_key"):
                validation.add_error(
                    RS_NULLABLE_VIOLATION,
                    "next_action.planned_idempotency_key",
                    "next_action.kind=run_skill 时 planned_idempotency_key 必须非 null",
                )
        if next_kind in {"retry", "manual_recover"}:
            if not next_action.get("target_result_id") and not next_action.get("target_error_code"):
                validation.add_error(
                    RS_NULLABLE_VIOLATION,
                    "next_action.target_result_id",
                    f"next_action.kind={next_kind} 时 target_result_id 与 target_error_code 至少一项非 null",
                )
            if not next_action.get("resume_from_stage"):
                validation.add_error(
                    RS_NULLABLE_VIOLATION,
                    "next_action.resume_from_stage",
                    f"next_action.kind={next_kind} 时 resume_from_stage 必须非 null",
                )
            if not next_action.get("planned_idempotency_key"):
                validation.add_error(
                    RS_NULLABLE_VIOLATION,
                    "next_action.planned_idempotency_key",
                    f"next_action.kind={next_kind} 时 planned_idempotency_key 必须非 null",
                )
        if state.get("applied_result_id") is not None:
            if next_action.get("source_result_id") != state.get("applied_result_id"):
                validation.add_error(
                    RS_CONSISTENCY_ERROR,
                    "next_action.source_result_id",
                    "applied_result_id 非 null 时 next_action.source_result_id 必须等于 applied_result_id",
                )

    return validation
