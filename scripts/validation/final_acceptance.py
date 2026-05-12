"""v4 final_acceptance 与 reporting_result 最小生成和校验工具。

覆盖条款：
- 40-§7.4 final_acceptance 生成后不可变，reporting_result 只能后置引用。
- 41-§11.8 final_acceptance 四类 acceptance_type 与 pre_acceptance manifest 绑定。
- 41-§11.8.1 reporting_result 独立记录报告阶段结果。
- 50-§5.1 CODE-010 final acceptance 分支、不可变边界和报告后置引用。
"""

from __future__ import annotations

import json
import os
import re
from copy import deepcopy
from pathlib import Path
from typing import Any

from scripts.validation.common_predicates import is_reporting_result_post_only
from scripts.validation.evidence_manifest import GENERATION_PATHS, validate_evidence_manifest
from scripts.validation.skill_result_io import canonical_json, compute_file_sha256


FINAL_ACCEPTANCE_PATH = "logs/final_acceptance.json"
REPORTING_RESULT_PATH = "logs/reporting_result.json"
PRE_ACCEPTANCE_MANIFEST_PATH = GENERATION_PATHS["pre_acceptance"]

ACCEPTANCE_TYPES = {
    "final_delivery",
    "audit_only_terminal",
    "build_rules_terminal",
    "blocked_terminal",
}
FINAL_ACCEPTANCE_STATUSES = {"accepted", "accepted_with_warnings", "blocked"}
REPORTING_STATUSES = {"done", "blocked"}
BLOCKING_CATEGORIES = {
    "toc_failed",
    "high_risk_unconfirmed",
    "manual_review_pending",
    "evidence_chain_broken",
    "schema_invalid",
    "path_escape",
    "original_docx_modified",
    "repair_action_failed",
}
ALLOWED_WARNING_CATEGORIES = {"non_blocking_warning"}
FORBIDDEN_FINAL_REPORTING_FIELDS = {
    "report_refs",
    "report_path",
    "report_paths",
    "report_artifacts",
    "reporting_result_path",
    "reporting_manifest_ref",
}
NULL_OR_OMITTED_FINAL_DELIVERY_FIELDS = {
    "final_docx_path",
    "toc_acceptance_path",
    "repair_execution_log_path",
    "repair_plan_finalized_path",
    "after_snapshot_ref",
    "review_result_refs",
}
MANUAL_REVIEW_SUMMARY_REQUIRED_FIELDS = {
    "required",
    "status",
    "items_path",
    "items_sha256",
    "items_size_bytes",
    "pending_count",
    "blocking_count",
    "unresolved_manual_review_count",
    "high_risk_unconfirmed_count",
    "cleared_review_ids",
    "blocking_review_ids",
    "evidence_refs",
}
FINALIZED_PLAN_PATH_PATTERN = re.compile(r"^plans/repair_plan\.finalized\.r[0-9]+\.yaml$")
FINAL_DOCX_PATH_PATTERN = re.compile(r"^output/[^/\\]+[0-9]{12}(_r[0-9]{2})?\.docx$")
MANUAL_REVIEW_STATUSES = {"not_required", "pending", "cleared", "blocked"}
TOC_MODES = {"native_toc", "equivalent_visible_toc", "not_required"}
TOC_ACCEPTANCE_STATUSES = {"accepted", "accepted_with_warnings", "blocked"}
TOC_REQUIRED_FIELDS = {
    "schema_id",
    "schema_version",
    "contract_version",
    "run_id",
    "toc_required",
    "toc_mode",
    "office_refresh_attempted",
    "office_refresh_succeeded",
    "placeholder_removed",
    "toc_field_count",
    "visible_entry_count",
    "source_refs",
    "source_action_ids",
    "final_docx_path",
    "final_docx_sha256",
    "final_docx_size_bytes",
    "acceptance_status",
    "evidence_refs",
    "checked_at",
}
TOC_EXEMPTION_SOURCE_TYPES = {"toc_rule", "toc-rules", "manual_review_item", "repair_action", "repair_plan"}
REPORT_ARTIFACT_REQUIRED_FIELDS = {
    "artifact_id",
    "kind",
    "path",
    "path_kind",
    "schema_id",
    "schema_version",
    "sha256",
    "size_bytes",
    "required",
    "producer_result_id",
    "report_type",
    "audience",
    "language",
}


class FinalAcceptanceError(ValueError):
    """final_acceptance/reporting_result 契约错误。"""


def _resolve_run_relative_path(run_dir: Path, rel_path: str) -> Path:
    """解析 run-relative 路径并阻断路径穿越。"""
    if not rel_path or Path(rel_path).is_absolute():
        raise FinalAcceptanceError(f"path must be run-relative: {rel_path}")
    base = run_dir.resolve()
    resolved = (base / rel_path).resolve()
    if resolved != base and base not in resolved.parents:
        raise FinalAcceptanceError(f"path escapes run_dir: {rel_path}")
    return resolved


def _resolve_reference_path(run_dir: Path, rel_path: str) -> Path:
    """解析 run-relative 或 workspace-level 规则包引用路径。"""
    run_candidate = _resolve_run_relative_path(run_dir, rel_path)
    if run_candidate.exists() or not rel_path.startswith("format-rules/"):
        return run_candidate
    workspace_candidate = (run_dir.resolve().parent.parent / rel_path).resolve()
    workspace_root = run_dir.resolve().parent.parent
    if workspace_candidate != workspace_root and workspace_root in workspace_candidate.parents:
        return workspace_candidate
    return run_candidate


def _load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _write_json_atomic(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    tmp_path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(tmp_path, path)


def _file_ref(run_dir: Path, rel_path: str, *, field_prefix: str) -> dict[str, Any]:
    path = _resolve_run_relative_path(run_dir, rel_path)
    if not path.exists() or not path.is_file():
        raise FinalAcceptanceError(f"{field_prefix}_path does not exist: {rel_path}")
    return {
        f"{field_prefix}_path": rel_path,
        f"{field_prefix}_sha256": compute_file_sha256(path),
        f"{field_prefix}_size_bytes": path.stat().st_size,
    }


def _validate_file_ref(
    data: dict[str, Any],
    run_dir: Path,
    *,
    path_field: str,
    sha_field: str,
    size_field: str,
    errors: list[str],
) -> None:
    rel_path = data.get(path_field)
    if not rel_path:
        errors.append(f"{path_field} is required")
        return
    try:
        path = _resolve_run_relative_path(run_dir, str(rel_path))
    except FinalAcceptanceError as exc:
        errors.append(str(exc))
        return
    if not path.exists() or not path.is_file():
        errors.append(f"{path_field} does not exist: {rel_path}")
        return
    if data.get(sha_field) != compute_file_sha256(path):
        errors.append(f"{sha_field} must match real file sha256")
    if data.get(size_field) != path.stat().st_size:
        errors.append(f"{size_field} must match real file size")


def _validate_reference_hash(
    run_dir: Path | None,
    *,
    rel_path: Any,
    sha256: Any,
    field_prefix: str,
    errors: list[str],
) -> None:
    if not rel_path:
        errors.append(f"{field_prefix}_path is required")
        return
    if not sha256:
        errors.append(f"{field_prefix}_sha256 is required")
        return
    if run_dir is None:
        return
    try:
        path = _resolve_reference_path(run_dir, str(rel_path))
    except FinalAcceptanceError as exc:
        errors.append(str(exc))
        return
    if not path.exists() or not path.is_file():
        errors.append(f"{field_prefix}_path does not exist: {rel_path}")
        return
    if sha256 != compute_file_sha256(path):
        errors.append(f"{field_prefix}_sha256 must match real file sha256")


def _validate_object_file_ref(
    ref: Any,
    run_dir: Path | None,
    *,
    field_name: str,
    required_extra_fields: set[str],
    errors: list[str],
) -> None:
    if not isinstance(ref, dict):
        errors.append(f"{field_name} must be object")
        return
    for child in {"path", "sha256", "size_bytes"} | required_extra_fields:
        if child not in ref:
            errors.append(f"{field_name}.{child} is required")
    if run_dir is None or not ref.get("path"):
        return
    try:
        path = _resolve_reference_path(run_dir, str(ref["path"]))
    except FinalAcceptanceError as exc:
        errors.append(str(exc))
        return
    if not path.exists() or not path.is_file():
        errors.append(f"{field_name}.path does not exist: {ref.get('path')}")
        return
    if ref.get("sha256") != compute_file_sha256(path):
        errors.append(f"{field_name}.sha256 must match real file sha256")
    if ref.get("size_bytes") != path.stat().st_size:
        errors.append(f"{field_name}.size_bytes must match real file size")


def _validate_manual_review_summary(
    summary: Any,
    run_dir: Path | None,
    *,
    final_status: str | None,
    errors: list[str],
) -> None:
    if not isinstance(summary, dict):
        errors.append("manual_review_summary must be object")
        return
    for field_name in sorted(MANUAL_REVIEW_SUMMARY_REQUIRED_FIELDS):
        if field_name not in summary:
            errors.append(f"manual_review_summary.{field_name} is required")
    for count_field in ("pending_count", "blocking_count", "unresolved_manual_review_count", "high_risk_unconfirmed_count"):
        if not isinstance(summary.get(count_field), int):
            errors.append(f"manual_review_summary.{count_field} must be integer")
    if summary.get("status") not in MANUAL_REVIEW_STATUSES:
        errors.append("manual_review_summary.status is not allowed")
    if final_status in {"accepted", "accepted_with_warnings"}:
        for count_field in ("pending_count", "blocking_count", "high_risk_unconfirmed_count"):
            if isinstance(summary.get(count_field), int) and summary[count_field] > 0:
                errors.append(f"manual_review_summary.{count_field}>0 requires blocked final_acceptance")
    for list_field in ("cleared_review_ids", "blocking_review_ids", "evidence_refs"):
        if list_field in summary and not isinstance(summary.get(list_field), list):
            errors.append(f"manual_review_summary.{list_field} must be array")
    if run_dir is not None and summary.get("items_path"):
        try:
            items_path = _resolve_run_relative_path(run_dir, str(summary["items_path"]))
        except FinalAcceptanceError as exc:
            errors.append(str(exc))
            return
        if not items_path.exists() or not items_path.is_file():
            errors.append(f"manual_review_summary.items_path does not exist: {summary.get('items_path')}")
            return
        if summary.get("items_sha256") != compute_file_sha256(items_path):
            errors.append("manual_review_summary.items_sha256 must match manual_review_items file")
        if summary.get("items_size_bytes") != items_path.stat().st_size:
            errors.append("manual_review_summary.items_size_bytes must match manual_review_items file")


def _validate_final_docx_name(final_acceptance: dict[str, Any], errors: list[str]) -> None:
    rel_path = final_acceptance.get("final_docx_path")
    if not rel_path:
        return
    if not FINAL_DOCX_PATH_PATTERN.match(str(rel_path)):
        errors.append("final_docx_path must match output/{source_name}{yyyyMMddHHmm}(_rNN)?.docx")
    lowered = str(rel_path).lower()
    if lowered.endswith(".formatted.docx") or lowered.endswith("_with_toc.docx") or "/_internal/" in lowered:
        errors.append("final_docx_path must not expose internal status names")


def _validate_toc_acceptance(final_acceptance: dict[str, Any], run_dir: Path | None, errors: list[str]) -> None:
    if run_dir is None or not final_acceptance.get("toc_acceptance_path"):
        return
    try:
        toc_path = _resolve_run_relative_path(run_dir, str(final_acceptance["toc_acceptance_path"]))
    except FinalAcceptanceError as exc:
        errors.append(str(exc))
        return
    if not toc_path.exists() or not toc_path.is_file():
        return
    try:
        toc = _load_json(toc_path)
    except json.JSONDecodeError as exc:
        errors.append(f"toc_acceptance_path invalid json: {exc}")
        return
    toc_errors = validate_toc_acceptance_v4(toc, final_acceptance=final_acceptance)
    errors.extend(toc_errors)


def validate_toc_acceptance_v4(
    toc_acceptance: dict[str, Any],
    *,
    final_acceptance: dict[str, Any] | None = None,
) -> list[str]:
    """校验 toc-acceptance 三种 toc_mode 硬条件。"""
    errors: list[str] = []
    if not isinstance(toc_acceptance, dict):
        return ["toc_acceptance must be object"]
    for field_name in sorted(TOC_REQUIRED_FIELDS):
        if field_name not in toc_acceptance:
            errors.append(f"toc_acceptance.{field_name} is required")
    if toc_acceptance.get("schema_id") != "toc-acceptance":
        errors.append("toc_acceptance.schema_id must be toc-acceptance")
    if toc_acceptance.get("contract_version") != "v4":
        errors.append("toc_acceptance.contract_version must be v4")
    if toc_acceptance.get("acceptance_status") not in TOC_ACCEPTANCE_STATUSES:
        errors.append("toc_acceptance.acceptance_status is not allowed")
    mode = toc_acceptance.get("toc_mode")
    if mode not in TOC_MODES:
        errors.append("toc_acceptance.toc_mode is not allowed")
    source_refs = toc_acceptance.get("source_refs")
    evidence_refs = toc_acceptance.get("evidence_refs")
    if not isinstance(source_refs, list) or not source_refs:
        errors.append("toc_acceptance.source_refs must be non-empty")
        source_refs = []
    if not isinstance(evidence_refs, list) or not evidence_refs:
        errors.append("toc_acceptance.evidence_refs must be non-empty")
    if not isinstance(toc_acceptance.get("source_action_ids"), list):
        errors.append("toc_acceptance.source_action_ids must be array")
    else:
        derived_action_ids = [
            str(ref.get("item_id"))
            for ref in source_refs
            if isinstance(ref, dict) and ref.get("item_type") == "repair_action" and ref.get("item_id")
        ]
        if sorted(derived_action_ids) != sorted([str(item) for item in toc_acceptance.get("source_action_ids", [])]):
            errors.append("toc_acceptance.source_action_ids must derive from source_refs repair_action item_id")
    for count_field in ("toc_field_count", "visible_entry_count"):
        if not isinstance(toc_acceptance.get(count_field), int):
            errors.append(f"toc_acceptance.{count_field} must be integer")
    if toc_acceptance.get("final_docx_path") and final_acceptance is not None:
        if toc_acceptance.get("final_docx_path") != final_acceptance.get("final_docx_path"):
            errors.append("toc_acceptance.final_docx_path must match final_acceptance.final_docx_path")
        if toc_acceptance.get("final_docx_sha256") != final_acceptance.get("final_docx_sha256"):
            errors.append("toc_acceptance.final_docx_sha256 must match final_acceptance.final_docx_sha256")
        if toc_acceptance.get("final_docx_size_bytes") != final_acceptance.get("final_docx_size_bytes"):
            errors.append("toc_acceptance.final_docx_size_bytes must match final_acceptance.final_docx_size_bytes")

    hard_errors: list[str] = []
    if mode == "native_toc":
        for bool_field in ("toc_required", "office_refresh_attempted", "office_refresh_succeeded", "placeholder_removed"):
            if toc_acceptance.get(bool_field) is not True:
                hard_errors.append(f"toc_acceptance.{bool_field} must be true for native_toc")
        if isinstance(toc_acceptance.get("toc_field_count"), int) and toc_acceptance["toc_field_count"] <= 0:
            hard_errors.append("toc_acceptance.toc_field_count must be >0 for native_toc")
        if isinstance(toc_acceptance.get("visible_entry_count"), int) and toc_acceptance["visible_entry_count"] <= 0:
            hard_errors.append("toc_acceptance.visible_entry_count must be >0 for native_toc")
        if toc_acceptance.get("not_required_reason"):
            errors.append("toc_acceptance.not_required_reason must be empty unless toc_mode=not_required")
    elif mode == "equivalent_visible_toc":
        if toc_acceptance.get("toc_required") is not True:
            hard_errors.append("toc_acceptance.toc_required must be true for equivalent_visible_toc")
        if toc_acceptance.get("placeholder_removed") is not True:
            hard_errors.append("toc_acceptance.placeholder_removed must be true for equivalent_visible_toc")
        if isinstance(toc_acceptance.get("visible_entry_count"), int) and toc_acceptance["visible_entry_count"] <= 0:
            hard_errors.append("toc_acceptance.visible_entry_count must be >0 for equivalent_visible_toc")
        if toc_acceptance.get("not_required_reason"):
            errors.append("toc_acceptance.not_required_reason must be empty unless toc_mode=not_required")
    elif mode == "not_required":
        if toc_acceptance.get("toc_required") is not False:
            hard_errors.append("toc_acceptance.toc_required must be false for not_required")
        if not toc_acceptance.get("not_required_reason"):
            hard_errors.append("toc_acceptance.not_required_reason is required when toc_mode=not_required")
        source_types = {ref.get("item_type") for ref in source_refs if isinstance(ref, dict)}
        if not source_types.intersection(TOC_EXEMPTION_SOURCE_TYPES):
            hard_errors.append("toc_acceptance.not_required requires toc rule, manual review, repair action, or repair plan source_ref")
        for ref in source_refs:
            if not isinstance(ref, dict):
                continue
            operation = ref.get("operation") or ref.get("action_type")
            if ref.get("item_type") == "repair_action" and operation and operation != "toc_exemption":
                hard_errors.append("toc_acceptance.not_required repair_action source_ref must be toc_exemption")
        if toc_acceptance.get("office_refresh_attempted") and toc_acceptance.get("office_refresh_succeeded") is False:
            hard_errors.append("toc_acceptance.not_required must not be caused by Office refresh failure")

    if toc_acceptance.get("acceptance_status") in {"accepted", "accepted_with_warnings"}:
        errors.extend(hard_errors)
        if hard_errors:
            errors.append("toc_acceptance.acceptance_status=accepted or accepted_with_warnings requires all toc_mode hard conditions")
    if final_acceptance is not None:
        final_status = final_acceptance.get("status")
        if final_status in {"accepted", "accepted_with_warnings"} and toc_acceptance.get("acceptance_status") != "accepted":
            errors.append("toc_acceptance.acceptance_status must be accepted for accepted final_delivery")
    return errors


def _validate_manifest_generation_ref(
    ref: Any,
    run_dir: Path | None,
    *,
    field_name: str,
    expected_generation: str,
    errors: list[str],
) -> None:
    if ref is None:
        return
    if not isinstance(ref, dict):
        errors.append(f"{field_name} must be object or null")
        return
    expected_path = GENERATION_PATHS[expected_generation]
    for child in ("path", "role", "path_kind", "sha256", "size_bytes", "manifest_generation", "status"):
        if child not in ref:
            errors.append(f"{field_name}.{child} is required")
    if ref.get("manifest_generation") != expected_generation:
        errors.append(f"{field_name}.manifest_generation must be {expected_generation}")
    if ref.get("path") != expected_path:
        errors.append(f"{field_name}.path must be {expected_path}")
    if ref.get("path_kind") != "run_relative":
        errors.append(f"{field_name}.path_kind must be run_relative")
    if ref.get("role") != "artifact":
        errors.append(f"{field_name}.role must be artifact")
    if run_dir is None or not ref.get("path"):
        return
    try:
        path = _resolve_run_relative_path(run_dir, str(ref["path"]))
    except FinalAcceptanceError as exc:
        errors.append(str(exc))
        return
    if not path.exists() or not path.is_file():
        errors.append(f"{field_name}.path does not exist: {ref.get('path')}")
        return
    if ref.get("sha256") != compute_file_sha256(path):
        errors.append(f"{field_name}.sha256 must match manifest file")
    if ref.get("size_bytes") != path.stat().st_size:
        errors.append(f"{field_name}.size_bytes must match manifest file")
    try:
        manifest = _load_json(path)
    except json.JSONDecodeError as exc:
        errors.append(f"{field_name}.path invalid json: {exc}")
        return
    if ref.get("status") != manifest.get("status"):
        errors.append(f"{field_name}.status must match manifest status")


def _validate_pre_acceptance_manifest(final_acceptance: dict[str, Any], run_dir: Path, errors: list[str]) -> None:
    if final_acceptance.get("evidence_manifest_path") != PRE_ACCEPTANCE_MANIFEST_PATH:
        errors.append("evidence_manifest_path must be logs/evidence_manifest.pre_acceptance.json")
        return
    manifest_path = _resolve_run_relative_path(run_dir, PRE_ACCEPTANCE_MANIFEST_PATH)
    if not manifest_path.exists():
        errors.append("pre_acceptance evidence manifest is required")
        return
    manifest = _load_json(manifest_path)
    if final_acceptance.get("evidence_manifest_sha256") != compute_file_sha256(manifest_path):
        errors.append("evidence_manifest_sha256 must match pre_acceptance manifest file")
    if final_acceptance.get("evidence_manifest_size_bytes") != manifest_path.stat().st_size:
        errors.append("evidence_manifest_size_bytes must match pre_acceptance manifest file")
    validation = validate_evidence_manifest(manifest, run_dir=run_dir)
    if not validation.valid:
        errors.extend(f"pre_acceptance_manifest.{error}" for error in validation.errors)
    if manifest.get("manifest_generation") != "pre_acceptance":
        errors.append("evidence manifest generation must be pre_acceptance")


def _validate_common_final_acceptance(
    final_acceptance: dict[str, Any],
    errors: list[str],
) -> None:
    required = {
        "schema_id",
        "schema_version",
        "contract_version",
        "run_id",
        "acceptance_type",
        "status",
        "skill_results",
        "evidence_manifest_path",
        "evidence_manifest_sha256",
        "evidence_manifest_size_bytes",
        "manual_review_summary",
        "warnings",
        "blockers",
        "blocking_categories",
        "allowed_warning_categories",
        "evaluated_at",
    }
    for field_name in sorted(required):
        if field_name not in final_acceptance:
            errors.append(f"{field_name} is required")
    if final_acceptance.get("schema_id") != "final-acceptance":
        errors.append("schema_id must be final-acceptance")
    if final_acceptance.get("contract_version") != "v4":
        errors.append("contract_version must be v4")
    if final_acceptance.get("acceptance_type") not in ACCEPTANCE_TYPES:
        errors.append(f"acceptance_type is not allowed: {final_acceptance.get('acceptance_type')}")
    if final_acceptance.get("status") not in FINAL_ACCEPTANCE_STATUSES:
        errors.append(f"status is not allowed: {final_acceptance.get('status')}")
    if any(field in final_acceptance for field in FORBIDDEN_FINAL_REPORTING_FIELDS):
        errors.append("final_acceptance must not contain reporting fields")
    if final_acceptance.get("blockers") and final_acceptance.get("status") != "blocked":
        errors.append("blockers require status=blocked")
    if final_acceptance.get("status") == "accepted" and final_acceptance.get("blocking_categories"):
        errors.append("accepted final_acceptance must not have blocking_categories")
    if final_acceptance.get("status") == "accepted_with_warnings" and not final_acceptance.get("allowed_warning_categories"):
        errors.append("accepted_with_warnings requires allowed_warning_categories")
    for array_name in ("skill_results", "warnings", "blockers", "blocking_categories", "allowed_warning_categories"):
        if not isinstance(final_acceptance.get(array_name), list):
            errors.append(f"{array_name} must be array")
    categories = set(final_acceptance.get("blocking_categories") or [])
    warning_categories = set(final_acceptance.get("allowed_warning_categories") or [])
    if "reporting_incomplete" in categories or "reporting_incomplete" in warning_categories:
        errors.append("reporting_incomplete must not enter final_acceptance categories")
    unknown_blocking = categories - BLOCKING_CATEGORIES
    if unknown_blocking:
        errors.append(f"unknown blocking_categories: {sorted(unknown_blocking)}")
    unknown_warnings = warning_categories - ALLOWED_WARNING_CATEGORIES
    if unknown_warnings:
        errors.append(f"unknown allowed_warning_categories: {sorted(unknown_warnings)}")


def _validate_terminal_without_final_delivery_fields(final_acceptance: dict[str, Any], errors: list[str]) -> None:
    for field_name in NULL_OR_OMITTED_FINAL_DELIVERY_FIELDS:
        if final_acceptance.get(field_name) not in (None, [], {}):
            errors.append(f"{field_name} must be omitted or null outside final_delivery")


def _validate_final_delivery(final_acceptance: dict[str, Any], run_dir: Path | None, errors: list[str]) -> None:
    for path_field, sha_field, size_field in (
        ("final_docx_path", "final_docx_sha256", "final_docx_size_bytes"),
        ("toc_acceptance_path", "toc_acceptance_sha256", "toc_acceptance_size_bytes"),
        ("repair_execution_log_path", "repair_execution_log_sha256", "repair_execution_log_size_bytes"),
        ("repair_plan_finalized_path", "repair_plan_finalized_sha256", "repair_plan_finalized_size_bytes"),
    ):
        if run_dir is None:
            for field_name in (path_field, sha_field, size_field):
                if field_name not in final_acceptance:
                    errors.append(f"{field_name} is required for final_delivery")
        else:
            _validate_file_ref(
                final_acceptance,
                run_dir,
                path_field=path_field,
                sha_field=sha_field,
                size_field=size_field,
                errors=errors,
            )
    _validate_final_docx_name(final_acceptance, errors)
    _validate_toc_acceptance(final_acceptance, run_dir, errors)
    for field_name in ("original_docx_proof", "after_snapshot_ref", "review_result_refs"):
        if field_name not in final_acceptance:
            errors.append(f"{field_name} is required for final_delivery")
    if "original_docx_untouched" not in final_acceptance:
        errors.append("original_docx_untouched is required for final_delivery")
    elif final_acceptance.get("original_docx_untouched") is not True:
        errors.append("original_docx_untouched must be true for accepted final_delivery")
    proof = final_acceptance.get("original_docx_proof")
    if not isinstance(proof, dict):
        errors.append("original_docx_proof must be object")
    else:
        for field_name in ("initial_sha256", "current_sha256", "initial_size_bytes", "current_size_bytes"):
            if field_name not in proof:
                errors.append(f"original_docx_proof.{field_name} is required")
        if proof.get("initial_sha256") != proof.get("current_sha256") or proof.get("initial_size_bytes") != proof.get("current_size_bytes"):
            errors.append("original_docx_proof must prove original docx is untouched")
    plan_path = final_acceptance.get("repair_plan_finalized_path")
    if plan_path and not FINALIZED_PLAN_PATH_PATTERN.match(str(plan_path)):
        errors.append("repair_plan_finalized_path must match plans/repair_plan.finalized.r{plan_revision}.yaml")
    _validate_object_file_ref(
        final_acceptance.get("after_snapshot_ref"),
        run_dir,
        field_name="after_snapshot_ref",
        required_extra_fields={"artifact_id", "snapshot_id"},
        errors=errors,
    )
    review_refs = final_acceptance.get("review_result_refs")
    if not isinstance(review_refs, list) or not review_refs:
        errors.append("review_result_refs must contain at least one item for final_delivery")
        review_refs = []
    for index, review_ref in enumerate(final_acceptance.get("review_result_refs") or []):
        if not isinstance(review_ref, dict):
            errors.append(f"review_result_refs[{index}] must be object")
            continue
        _validate_object_file_ref(
            review_ref,
            run_dir,
            field_name=f"review_result_refs[{index}]",
            required_extra_fields={"review_id", "status"},
            errors=errors,
        )
        if review_ref.get("status") == "failed" and final_acceptance.get("status") != "blocked":
            errors.append("failed review_result_refs require status=blocked")


def _validate_audit_only_terminal(final_acceptance: dict[str, Any], run_dir: Path | None, errors: list[str]) -> None:
    for field_name in ("source_audit_refs", "source_snapshot_ref", "rule_ref_path", "rule_ref_sha256", "audit_summary"):
        if field_name not in final_acceptance:
            errors.append(f"{field_name} is required for audit_only_terminal")
    for index, audit_ref in enumerate(final_acceptance.get("source_audit_refs") or []):
        _validate_object_file_ref(
            audit_ref,
            run_dir,
            field_name=f"source_audit_refs[{index}]",
            required_extra_fields=set(),
            errors=errors,
        )
    _validate_object_file_ref(
        final_acceptance.get("source_snapshot_ref"),
        run_dir,
        field_name="source_snapshot_ref",
        required_extra_fields=set(),
        errors=errors,
    )
    _validate_reference_hash(
        run_dir,
        rel_path=final_acceptance.get("rule_ref_path"),
        sha256=final_acceptance.get("rule_ref_sha256"),
        field_prefix="rule_ref",
        errors=errors,
    )
    _validate_terminal_without_final_delivery_fields(final_acceptance, errors)


def _validate_build_rules_terminal(final_acceptance: dict[str, Any], run_dir: Path | None, errors: list[str]) -> None:
    for field_name in ("rule_ref_path", "rule_ref_sha256", "package_manifest_path", "package_manifest_sha256"):
        if field_name not in final_acceptance:
            errors.append(f"{field_name} is required for build_rules_terminal")
    _validate_reference_hash(
        run_dir,
        rel_path=final_acceptance.get("rule_ref_path"),
        sha256=final_acceptance.get("rule_ref_sha256"),
        field_prefix="rule_ref",
        errors=errors,
    )
    _validate_reference_hash(
        run_dir,
        rel_path=final_acceptance.get("package_manifest_path"),
        sha256=final_acceptance.get("package_manifest_sha256"),
        field_prefix="package_manifest",
        errors=errors,
    )
    if final_acceptance.get("rule_package_status") != "active":
        errors.append("rule_package_status must be active")
    if final_acceptance.get("activation_decision_status") != "approved":
        errors.append("activation_decision_status must be approved")
    manual_summary = final_acceptance.get("manual_review_summary") or {}
    for count_field in ("pending_count", "blocking_count", "unresolved_manual_review_count", "high_risk_unconfirmed_count"):
        if manual_summary.get(count_field) != 0:
            errors.append(f"manual_review_summary.{count_field} must be 0 for build_rules_terminal")
    _validate_terminal_without_final_delivery_fields(final_acceptance, errors)


def _validate_blocked_terminal(final_acceptance: dict[str, Any], errors: list[str]) -> None:
    if final_acceptance.get("status") != "blocked":
        errors.append("blocked_terminal requires status=blocked")
    for field_name in ("terminal_stage", "terminal_result_id", "terminal_blocker_refs"):
        if field_name not in final_acceptance:
            errors.append(f"{field_name} is required for blocked_terminal")
    if not final_acceptance.get("terminal_blocker_refs"):
        errors.append("blocked_terminal requires terminal_blocker_refs")
    _validate_terminal_without_final_delivery_fields(final_acceptance, errors)


def validate_final_acceptance_v4(final_acceptance: dict[str, Any], *, run_dir: Path | None = None) -> list[str]:
    """校验 v4 final_acceptance.json；返回错误列表。"""
    errors: list[str] = []
    if not isinstance(final_acceptance, dict):
        return ["final_acceptance must be object"]
    _validate_common_final_acceptance(final_acceptance, errors)
    _validate_manual_review_summary(
        final_acceptance.get("manual_review_summary"),
        run_dir,
        final_status=final_acceptance.get("status"),
        errors=errors,
    )
    if run_dir is not None:
        _validate_pre_acceptance_manifest(final_acceptance, run_dir, errors)
    acceptance_type = final_acceptance.get("acceptance_type")
    if acceptance_type == "final_delivery":
        _validate_final_delivery(final_acceptance, run_dir, errors)
    elif acceptance_type == "audit_only_terminal":
        _validate_audit_only_terminal(final_acceptance, run_dir, errors)
    elif acceptance_type == "build_rules_terminal":
        _validate_build_rules_terminal(final_acceptance, run_dir, errors)
    elif acceptance_type == "blocked_terminal":
        _validate_blocked_terminal(final_acceptance, errors)
    return errors


def build_final_acceptance(
    run_dir: Path,
    *,
    run_id: str,
    acceptance_type: str,
    status: str,
    skill_results: list[dict[str, Any]] | None = None,
    manual_review_summary: dict[str, Any] | None = None,
    warnings: list[dict[str, Any]] | None = None,
    blockers: list[dict[str, Any]] | None = None,
    blocking_categories: list[str] | None = None,
    allowed_warning_categories: list[str] | None = None,
    evaluated_at: str = "2026-05-08T00:00:00+08:00",
    branch_fields: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """构造 final_acceptance，并复算 pre_acceptance manifest 的真实 hash/size。"""
    manifest_path = _resolve_run_relative_path(run_dir, PRE_ACCEPTANCE_MANIFEST_PATH)
    if not manifest_path.exists():
        raise FinalAcceptanceError("pre_acceptance evidence manifest is required before final_acceptance")
    final_acceptance: dict[str, Any] = {
        "schema_id": "final-acceptance",
        "schema_version": "1.0.0",
        "contract_version": "v4",
        "run_id": run_id,
        "acceptance_type": acceptance_type,
        "status": status,
        "skill_results": deepcopy(skill_results or []),
        "evidence_manifest_path": PRE_ACCEPTANCE_MANIFEST_PATH,
        "evidence_manifest_sha256": compute_file_sha256(manifest_path),
        "evidence_manifest_size_bytes": manifest_path.stat().st_size,
        "manual_review_summary": deepcopy(
            manual_review_summary
            or {
                "required": False,
                "status": "not_required",
                "items_path": "plans/manual_review_items.json",
                "items_sha256": "",
                "items_size_bytes": 0,
                "pending_count": 0,
                "blocking_count": 0,
                "unresolved_manual_review_count": 0,
                "high_risk_unconfirmed_count": 0,
                "cleared_review_ids": [],
                "blocking_review_ids": [],
                "evidence_refs": [],
            }
        ),
        "warnings": deepcopy(warnings or []),
        "blockers": deepcopy(blockers or []),
        "blocking_categories": list(blocking_categories or []),
        "allowed_warning_categories": list(allowed_warning_categories or []),
        "evaluated_at": evaluated_at,
    }
    final_acceptance.update(deepcopy(branch_fields or {}))
    errors = validate_final_acceptance_v4(final_acceptance, run_dir=run_dir)
    if errors:
        raise FinalAcceptanceError(f"final_acceptance validation failed: {errors}")
    return final_acceptance


def write_final_acceptance(run_dir: Path, final_acceptance: dict[str, Any]) -> dict[str, Any]:
    """写入 logs/final_acceptance.json；存在不同内容时阻断不可变边界。"""
    errors = validate_final_acceptance_v4(final_acceptance, run_dir=run_dir)
    if errors:
        raise FinalAcceptanceError(f"final_acceptance validation failed: {errors}")
    path = _resolve_run_relative_path(run_dir, FINAL_ACCEPTANCE_PATH)
    if path.exists():
        existing = _load_json(path)
        if canonical_json(existing) != canonical_json(final_acceptance):
            raise FinalAcceptanceError("final_acceptance is immutable after generation")
    else:
        _write_json_atomic(path, final_acceptance)
    return {
        "path": FINAL_ACCEPTANCE_PATH,
        "sha256": compute_file_sha256(path),
        "size_bytes": path.stat().st_size,
        "final_acceptance": _load_json(path),
    }


def validate_reporting_result(reporting_result: dict[str, Any], *, run_dir: Path | None = None) -> list[str]:
    """校验 reporting_result.json；报告阶段不得反向修改 final_acceptance。"""
    errors: list[str] = []
    required = {
        "schema_id",
        "schema_version",
        "contract_version",
        "run_id",
        "reporting_id",
        "status",
        "final_acceptance_path",
        "final_acceptance_sha256",
        "final_acceptance_size_bytes",
        "post_acceptance_manifest_ref",
        "reporting_manifest_ref",
        "report_artifacts",
        "warnings",
        "blockers",
        "generated_at",
    }
    for field_name in sorted(required):
        if field_name not in reporting_result:
            errors.append(f"{field_name} is required")
    if reporting_result.get("schema_id") != "reporting-result":
        errors.append("schema_id must be reporting-result")
    if reporting_result.get("contract_version") != "v4":
        errors.append("contract_version must be v4")
    if reporting_result.get("status") not in REPORTING_STATUSES:
        errors.append(f"status is not allowed: {reporting_result.get('status')}")
    if reporting_result.get("status") == "done":
        if not reporting_result.get("report_artifacts"):
            errors.append("status=done requires report_artifacts")
        if not reporting_result.get("reporting_manifest_ref"):
            errors.append("status=done requires reporting_manifest_ref")
    if not isinstance(reporting_result.get("report_artifacts"), list):
        errors.append("report_artifacts must be array")
    else:
        for index, artifact in enumerate(reporting_result.get("report_artifacts") or []):
            if not isinstance(artifact, dict):
                errors.append(f"report_artifacts[{index}] must be object")
                continue
            for field_name in sorted(REPORT_ARTIFACT_REQUIRED_FIELDS):
                if field_name not in artifact:
                    errors.append(f"report_artifacts[{index}].{field_name} is required")
            if artifact.get("kind") != "report":
                errors.append(f"report_artifacts[{index}].kind must be report")
            if artifact.get("path") == REPORTING_RESULT_PATH:
                errors.append(f"report_artifacts[{index}] must not point to reporting_result.json itself")
            if artifact.get("path_kind") != "run_relative":
                errors.append(f"report_artifacts[{index}].path_kind must be run_relative")
            if artifact.get("required") is not True:
                errors.append(f"report_artifacts[{index}].required must be true")
            if run_dir is not None and artifact.get("path"):
                try:
                    report_path = _resolve_run_relative_path(run_dir, str(artifact["path"]))
                except FinalAcceptanceError as exc:
                    errors.append(str(exc))
                else:
                    if not report_path.exists() or not report_path.is_file():
                        errors.append(f"report_artifacts[{index}].path does not exist: {artifact.get('path')}")
                    else:
                        if artifact.get("sha256") != compute_file_sha256(report_path):
                            errors.append(f"report_artifacts[{index}].sha256 must match report file")
                        if artifact.get("size_bytes") != report_path.stat().st_size:
                            errors.append(f"report_artifacts[{index}].size_bytes must match report file")
    for array_name in ("warnings", "blockers"):
        if not isinstance(reporting_result.get(array_name), list):
            errors.append(f"{array_name} must be array")
    if run_dir is not None:
        if reporting_result.get("final_acceptance_path") != FINAL_ACCEPTANCE_PATH:
            errors.append("final_acceptance_path must be logs/final_acceptance.json")
        final_path = _resolve_run_relative_path(run_dir, FINAL_ACCEPTANCE_PATH)
        if not final_path.exists():
            errors.append("referenced final_acceptance.json does not exist")
        else:
            if reporting_result.get("final_acceptance_sha256") != compute_file_sha256(final_path):
                errors.append("final_acceptance_sha256 must match current final_acceptance file")
            if reporting_result.get("final_acceptance_size_bytes") != final_path.stat().st_size:
                errors.append("final_acceptance_size_bytes must match current final_acceptance file")
            if not is_reporting_result_post_only(final_path, reporting_result.get("final_acceptance_sha256")):
                errors.append("reporting_result must be post-only and final_acceptance immutable")
        _validate_manifest_generation_ref(
            reporting_result.get("post_acceptance_manifest_ref"),
            run_dir,
            field_name="post_acceptance_manifest_ref",
            expected_generation="post_acceptance",
            errors=errors,
        )
        _validate_manifest_generation_ref(
            reporting_result.get("reporting_manifest_ref"),
            run_dir,
            field_name="reporting_manifest_ref",
            expected_generation="reporting",
            errors=errors,
        )
    return errors


def build_reporting_result(
    run_dir: Path,
    *,
    run_id: str,
    reporting_id: str,
    status: str,
    post_acceptance_manifest_ref: dict[str, Any] | None,
    reporting_manifest_ref: dict[str, Any] | None,
    report_artifacts: list[dict[str, Any]] | None = None,
    warnings: list[dict[str, Any]] | None = None,
    blockers: list[dict[str, Any]] | None = None,
    generated_at: str = "2026-05-08T00:00:00+08:00",
) -> dict[str, Any]:
    """构造 reporting_result，并引用不可变 final_acceptance 文件。"""
    final_path = _resolve_run_relative_path(run_dir, FINAL_ACCEPTANCE_PATH)
    if not final_path.exists():
        raise FinalAcceptanceError("final_acceptance.json is required before reporting_result")
    reporting_result = {
        "schema_id": "reporting-result",
        "schema_version": "1.0.0",
        "contract_version": "v4",
        "run_id": run_id,
        "reporting_id": reporting_id,
        "status": status,
        "final_acceptance_path": FINAL_ACCEPTANCE_PATH,
        "final_acceptance_sha256": compute_file_sha256(final_path),
        "final_acceptance_size_bytes": final_path.stat().st_size,
        "post_acceptance_manifest_ref": deepcopy(post_acceptance_manifest_ref),
        "reporting_manifest_ref": deepcopy(reporting_manifest_ref),
        "report_artifacts": deepcopy(report_artifacts or []),
        "warnings": deepcopy(warnings or []),
        "blockers": deepcopy(blockers or []),
        "generated_at": generated_at,
    }
    errors = validate_reporting_result(reporting_result, run_dir=run_dir)
    if errors:
        raise FinalAcceptanceError(f"reporting_result validation failed: {errors}")
    return reporting_result


def write_reporting_result(run_dir: Path, reporting_result: dict[str, Any]) -> dict[str, Any]:
    """写入 logs/reporting_result.json，并断言 final_acceptance hash 不变。"""
    final_path = _resolve_run_relative_path(run_dir, FINAL_ACCEPTANCE_PATH)
    before_sha = compute_file_sha256(final_path) if final_path.exists() else None
    errors = validate_reporting_result(reporting_result, run_dir=run_dir)
    if errors:
        raise FinalAcceptanceError(f"reporting_result validation failed: {errors}")
    path = _resolve_run_relative_path(run_dir, REPORTING_RESULT_PATH)
    _write_json_atomic(path, reporting_result)
    after_sha = compute_file_sha256(final_path)
    if before_sha != after_sha:
        raise FinalAcceptanceError("reporting_result write must not modify final_acceptance")
    return {
        "path": REPORTING_RESULT_PATH,
        "sha256": compute_file_sha256(path),
        "size_bytes": path.stat().st_size,
        "reporting_result": _load_json(path),
    }


__all__ = [
    "FINAL_ACCEPTANCE_PATH",
    "REPORTING_RESULT_PATH",
    "PRE_ACCEPTANCE_MANIFEST_PATH",
    "FinalAcceptanceError",
    "build_final_acceptance",
    "build_reporting_result",
    "validate_final_acceptance_v4",
    "validate_reporting_result",
    "validate_toc_acceptance_v4",
    "write_final_acceptance",
    "write_reporting_result",
]
