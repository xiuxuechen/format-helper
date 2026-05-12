"""V4 必测子集的最小自动化 fixture/validator。

覆盖条款：
- 40-§7.3 V4-T18/V4-T28/V4-T36/V4-T39/V4-T43/V4-T45。
- 41-§3.4 路径策略、§3.13 ReviewCheck Object、§11.10 repair-execution-log、
  §11.12 semantic-audit、§11.14 review-result。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any


ACTION_RESULT_STATUSES = {"executed", "skipped", "rejected", "failed"}
REVIEW_STATUSES = {"passed", "passed_with_warnings", "failed"}
REVIEW_CHECK_STATUSES = {"passed", "failed", "skipped", "not_applicable"}
REVIEW_CHECK_REQUIRED_FIELDS = {
    "check_id",
    "action_id",
    "check_type",
    "source_action_status",
    "source_execution_status",
    "target",
    "attribute",
    "observed",
    "expected",
    "status",
    "blocking",
    "evidence_refs",
    "message",
}


def validate_resolved_run_relative_path(run_dir: Path, rel_path: str, *, resolved_path: Path | None = None) -> list[str]:
    """校验 run_relative 路径解析后仍留在 run_dir 内。

    `resolved_path` 参数用于无权限创建 Windows junction/symlink 时构造等价的解析结果 fixture。
    """
    errors: list[str] = []
    path = Path(rel_path)
    if path.is_absolute():
        errors.append("run_relative path must not be absolute")
    if ".." in path.parts:
        errors.append("run_relative path must not contain parent traversal")
    base = run_dir.resolve()
    resolved = (resolved_path or (base / path)).resolve()
    if resolved != base and base not in resolved.parents:
        errors.append("resolved path escapes run_dir, including symlink/junction target")
    return errors


def validate_source_docx_consistency(
    snapshot: dict[str, Any],
    *,
    semantic_role_map: dict[str, Any] | None = None,
    audit: dict[str, Any] | None = None,
) -> list[str]:
    """校验 source_docx hash/size/artifact_id 在 snapshot、role map、audit 间一致。"""
    errors: list[str] = []
    expected = {
        "source_docx_artifact_id": snapshot.get("source_docx_artifact_id"),
        "source_docx_sha256": snapshot.get("source_docx_sha256"),
        "source_docx_size_bytes": snapshot.get("source_docx_size_bytes"),
    }
    for label, payload in (("semantic_role_map", semantic_role_map), ("audit", audit)):
        if payload is None:
            continue
        for field_name, expected_value in expected.items():
            if payload.get(field_name) != expected_value:
                errors.append(f"{label}.{field_name} must match document-snapshot")
    return errors


def validate_repair_execution_log_minimal(log: dict[str, Any]) -> list[str]:
    """校验 repair-execution-log 的执行输入指纹和 action_results 状态。"""
    errors: list[str] = []
    required = {
        "schema_id",
        "contract_version",
        "run_id",
        "repair_plan_path",
        "repair_plan_sha256",
        "repair_plan_size_bytes",
        "working_docx_path",
        "working_docx_sha256",
        "working_docx_size_bytes",
        "action_results",
    }
    for field_name in sorted(required):
        if field_name not in log or log.get(field_name) in (None, ""):
            errors.append(f"{field_name} is required")
    if log.get("schema_id") != "repair-execution-log":
        errors.append("schema_id must be repair-execution-log")
    if log.get("contract_version") != "v4":
        errors.append("contract_version must be v4")
    for size_field in ("repair_plan_size_bytes", "working_docx_size_bytes"):
        if not isinstance(log.get(size_field), int) or log.get(size_field) <= 0:
            errors.append(f"{size_field} must be positive integer")
    action_results = log.get("action_results")
    if not isinstance(action_results, list):
        errors.append("action_results must be array")
        return errors
    for index, action in enumerate(action_results):
        if not isinstance(action, dict):
            errors.append(f"action_results[{index}] must be object")
            continue
        if not action.get("action_id"):
            errors.append(f"action_results[{index}].action_id is required")
        if action.get("status") not in ACTION_RESULT_STATUSES:
            errors.append(f"action_results[{index}].status is invalid")
        if action.get("status") == "blocked":
            errors.append("action_results[].status must not be blocked")
    return errors


def validate_review_result_minimal(review: dict[str, Any], repair_log: dict[str, Any]) -> list[str]:
    """校验 review-result 不复制 action_results，并用 ReviewCheck 覆盖执行动作。"""
    errors: list[str] = []
    if "action_results" in review:
        errors.append("review-result must not embed action_results")
    if review.get("schema_id") != "review-result":
        errors.append("schema_id must be review-result")
    if review.get("status") not in REVIEW_STATUSES:
        errors.append("review-result.status is invalid")
    if review.get("status") == "failed" and review.get("gate_check_status") != "failed":
        errors.append("failed review-result requires gate_check_status=failed")

    action_results = repair_log.get("action_results") or []
    action_ids = {item.get("action_id") for item in action_results if isinstance(item, dict) and item.get("action_id")}
    covered = set(review.get("covered_action_ids") or [])
    if action_ids and not action_ids.issubset(covered):
        errors.append("covered_action_ids must cover all source repair action_results")
    checks = review.get("checks")
    if not isinstance(checks, list):
        errors.append("checks must be array")
        checks = []
    for index, check in enumerate(checks):
        if not isinstance(check, dict):
            errors.append(f"checks[{index}] must be object")
            continue
        for field_name in sorted(REVIEW_CHECK_REQUIRED_FIELDS):
            if field_name not in check:
                errors.append(f"checks[{index}].{field_name} is required")
        if check.get("status") not in REVIEW_CHECK_STATUSES:
            errors.append(f"checks[{index}].status is invalid")
        if check.get("source_action_status") == "blocked" and check.get("source_execution_status") not in (None, "rejected"):
            errors.append("blocked source_action_status requires source_execution_status null or rejected")
        if check.get("source_execution_status") == "executed":
            if check.get("check_type") not in {"after_value_match", "snapshot_diff"}:
                errors.append("executed action requires after_value_match or snapshot_diff check")
        if check.get("blocking") is True and check.get("status") == "failed":
            if review.get("status") != "failed" or review.get("gate_check_status") != "failed":
                errors.append("blocking failed ReviewCheck requires failed review-result")
    failed_actions = [item for item in action_results if isinstance(item, dict) and item.get("status") == "failed"]
    if failed_actions and (review.get("status") != "failed" or review.get("gate_check_status") != "failed"):
        errors.append("failed source action_result requires failed review-result")
    return errors


def validate_semantic_audit_unresolved_roles(audit: dict[str, Any]) -> list[str]:
    """校验 semantic-audit unresolved 角色分支不得产生可执行 SuggestedAction。"""
    errors: list[str] = []
    if audit.get("schema_id") != "semantic-audit":
        errors.append("schema_id must be semantic-audit")
    findings = audit.get("findings")
    if not isinstance(findings, list):
        return errors + ["findings must be array"]
    for index, finding in enumerate(findings):
        if not isinstance(finding, dict):
            errors.append(f"findings[{index}] must be object")
            continue
        if finding.get("target_role_id") is not None:
            continue
        if finding.get("conclusion") != "uncertain":
            errors.append(f"findings[{index}].conclusion must be uncertain when target_role_id=null")
        if "target_rule_item_id" not in finding:
            errors.append(f"findings[{index}].target_rule_item_id is required-but-nullable")
        expected = finding.get("expected")
        if not isinstance(expected, dict) or "source_rule_ref" not in expected:
            errors.append(f"findings[{index}].expected.source_rule_ref is required-but-nullable")
        proposals = finding.get("manual_review_proposal_ids")
        if not isinstance(proposals, list) or not proposals:
            errors.append(f"findings[{index}].manual_review_proposal_ids must be non-empty")
        for action_index, action in enumerate(finding.get("suggested_actions") or []):
            if not isinstance(action, dict):
                errors.append(f"findings[{index}].suggested_actions[{action_index}] must be object")
                continue
            if action.get("execution_status") == "executable" or action.get("allowed_by_policy") is True:
                errors.append("unresolved semantic finding must not generate executable SuggestedAction")
    return errors


__all__ = [
    "validate_repair_execution_log_minimal",
    "validate_resolved_run_relative_path",
    "validate_review_result_minimal",
    "validate_semantic_audit_unresolved_roles",
    "validate_source_docx_consistency",
]
