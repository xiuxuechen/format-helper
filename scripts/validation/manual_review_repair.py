"""manual-review-items 与 repair-plan 闭环校验（CODE-009）。"""

from __future__ import annotations

# 修复动作白名单
WHITELIST_ACTIONS = {
    "map_heading_native_style",
    "apply_body_style_definition",
    "apply_body_direct_format",
    "apply_table_cell_format",
    "apply_table_border",
    "toc_content_audit",
    "insert_or_replace_toc_field",
}

import json
import os
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from scripts.utils.simple_yaml import write_yaml
from scripts.validation.skill_result_io import (
    canonical_json,
    compute_file_sha256,
    resolve_run_relative_path,
    sha256_text,
)


MANUAL_REVIEW_ITEMS_PATH = "plans/manual_review_items.json"
MANUAL_REVIEW_SUMMARY_PATH = "reports/MANUAL_CONFIRMATION.md"
DRAFT_REPAIR_PLAN_PATH = "plans/repair_plan.draft.yaml"

DECISION_STATUSES = {
    "pending",
    "approved",
    "modified",
    "rejected",
    "deferred_non_blocking",
    "not_applicable",
}

FINALIZED_ALLOWED_EXECUTION_STATUSES = {
    "executable",
    "skipped",
    "rejected",
    "blocked",
}

SELECTED_ACTION_DECISION_STATUSES = {
    "approved",
    "modified",
}

SELECTED_ACTION_NULL_STATUSES = {
    "rejected",
    "deferred_non_blocking",
    "not_applicable",
}

SELECTED_ACTION_SOURCE_KINDS = {
    "semantic_suggested_action",
    "format_suggested_action",
    "manual_review_proposal",
    "modified_action",
}

SELECTED_ACTION_REQUIRED_FIELDS = {
    "selected_action_id",
    "review_id",
    "decision_status",
    "source_kind",
    "source_suggested_action_id",
    "source_proposal_id",
    "base_action_id",
    "action_type",
    "operation",
    "target",
    "before_value",
    "after_value",
    "desired_value",
    "parameters",
    "risk_level",
    "auto_fix_policy",
    "requires_manual_review",
    "policy_match_ref",
    "policy_evidence_refs",
    "source_refs",
    "evidence_refs",
    "reason",
}


@dataclass
class ReviewValidationResult:
    """manual-review 或 repair-plan 校验结果。"""

    valid: bool
    errors: list[str] = field(default_factory=list)


def read_json(path: Path) -> dict[str, Any]:
    """读取 JSON 对象。"""
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"JSON 根节点必须是对象：{path}")
    return data


def write_json_atomic(path: Path, data: dict[str, Any]) -> None:
    """原子写入 UTF-8 JSON。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temp_path, path)


def count_manual_review_items(items: list[dict[str, Any]]) -> dict[str, int]:
    """统计人工确认清单状态。"""
    pending_count = 0
    blocking_count = 0
    unresolved_count = 0
    for item in items:
        decision = item.get("decision") if isinstance(item, dict) else {}
        status = decision.get("status") if isinstance(decision, dict) else None
        if status == "pending":
            pending_count += 1
        if item.get("blocking") is True:
            blocking_count += 1
            if status == "pending" or (status == "rejected" and decision.get("allows_continue") is False):
                unresolved_count += 1
    return {
        "pending_count": pending_count,
        "blocking_count": blocking_count,
        "unresolved_manual_review_count": unresolved_count,
    }


def compute_selected_action_id(selected_action: dict[str, Any]) -> str:
    """按内容派生 selected_action_id，排除自身避免自引用。"""
    payload = deepcopy(selected_action)
    payload.pop("selected_action_id", None)
    payload.pop("selected_action_sha256", None)
    digest = sha256_text(canonical_json(payload))
    return f"SA-{digest[:16]}"


def compute_selected_action_sha256(selected_action: dict[str, Any]) -> str:
    """计算 SelectedAction canonical hash。"""
    payload = deepcopy(selected_action)
    payload.pop("selected_action_sha256", None)
    payload.pop("decided_by", None)
    payload.pop("decided_at", None)
    return sha256_text(canonical_json(payload))


def normalize_selected_action(selected_action: dict[str, Any]) -> dict[str, Any]:
    """补齐 selected_action_id 并返回规范化 SelectedAction。"""
    normalized = deepcopy(selected_action)
    if not normalized.get("selected_action_id"):
        normalized["selected_action_id"] = compute_selected_action_id(normalized)
    return normalized


def validate_selected_action(
    selected_action: dict[str, Any] | None,
    *,
    decision_status: str,
    expected_sha256: str | None = None,
    risk_policy: dict[str, Any] | None = None,
    policy_sha256: str | None = None,
) -> list[str]:
    """校验 SelectedAction required-but-nullable、hash 和白名单重算。"""
    errors: list[str] = []
    if decision_status in SELECTED_ACTION_NULL_STATUSES:
        if selected_action is not None:
            return [f"decision_status={decision_status} requires selected_action=null"]
        if expected_sha256 is not None:
            return [f"decision_status={decision_status} requires selected_action_sha256=null"]
        return []
    if decision_status not in SELECTED_ACTION_DECISION_STATUSES:
        if selected_action is not None:
            errors.append(f"decision_status={decision_status} must not carry selected_action")
        return errors
    if not isinstance(selected_action, dict):
        return [f"decision_status={decision_status} requires selected_action"]
    for field_name in sorted(SELECTED_ACTION_REQUIRED_FIELDS - selected_action.keys()):
        errors.append(f"selected_action.{field_name} is required")
    if selected_action.get("decision_status") != decision_status:
        errors.append("selected_action.decision_status must equal decision status")
    if selected_action.get("source_kind") not in SELECTED_ACTION_SOURCE_KINDS:
        errors.append("selected_action.source_kind is not allowed")
    if decision_status == "approved" and selected_action.get("source_kind") == "modified_action":
        errors.append("approved selected_action must not use modified_action source_kind")
    if decision_status == "modified":
        if selected_action.get("source_kind") != "modified_action":
            errors.append("modified selected_action requires source_kind=modified_action")
        if not (
            selected_action.get("base_action_id")
            or selected_action.get("source_suggested_action_id")
            or selected_action.get("source_proposal_id")
        ):
            errors.append("modified selected_action requires base_action_id or source id")
    if not isinstance(selected_action.get("source_refs"), list) or not selected_action.get("source_refs"):
        errors.append("selected_action.source_refs must be non-empty array")
    if expected_sha256 is None:
        errors.append("selected_action_sha256 is required for approved/modified decision")
    elif expected_sha256 != compute_selected_action_sha256(selected_action):
        errors.append("selected_action_sha256 does not match selected_action canonical hash")
    if risk_policy is not None and policy_sha256 is not None:
        action_like = {
            "operation": selected_action.get("operation"),
            "action_type": selected_action.get("action_type"),
            "target": selected_action.get("target"),
            "allowed_by_policy": not selected_action.get("requires_manual_review"),
            "policy_match_ref": selected_action.get("policy_match_ref"),
        }
        errors.extend(f"selected_action.{error}" for error in validate_policy_match(action_like, risk_policy, policy_sha256))
    return errors


def apply_manual_review_decision(
    manual_review_items: dict[str, Any],
    *,
    review_id: str,
    decision_status: str,
    allows_continue: bool,
    decided_by: str,
    decided_at: str,
    selected_action: dict[str, Any] | None = None,
    comment: str | None = None,
) -> dict[str, Any]:
    """写入单个人工确认决策，并自动计算 selected_action_sha256。"""
    updated = deepcopy(manual_review_items)
    for item in updated.get("items", []):
        if item.get("review_id") != review_id:
            continue
        normalized_action = normalize_selected_action(selected_action) if selected_action is not None else None
        sha256 = compute_selected_action_sha256(normalized_action) if normalized_action is not None else None
        item["decision"] = {
            "status": decision_status,
            "allows_continue": allows_continue,
            "selected_action": normalized_action,
            "selected_action_sha256": sha256,
            "decided_by": decided_by,
            "decided_at": decided_at,
            "comment": comment,
        }
        validation = validate_manual_review_items(updated)
        if not validation.valid:
            raise ValueError(f"manual-review decision 未通过校验：{validation.errors}")
        return updated
    raise ValueError(f"review_id 不存在：{review_id}")


def build_manual_review_items(
    *,
    run_id: str,
    proposals: list[dict[str, Any]],
    generated_at: str = "2026-05-08T00:00:00+08:00",
    writer: str = "format-helper",
) -> dict[str, Any]:
    """由主控把 ManualReviewItemDraft 提升为权威 manual-review-items。"""
    items: list[dict[str, Any]] = []
    for index, proposal in enumerate(proposals, start=1):
        review_id = f"MRI-{index:03d}"
        item = {
            "review_id": review_id,
            "proposal_id": proposal.get("proposal_id"),
            "source_issue_ids": deepcopy(proposal.get("source_issue_ids", [])),
            "source_refs": deepcopy(proposal.get("source_refs", [])),
            "category": proposal.get("category"),
            "problem": proposal.get("problem"),
            "impact": proposal.get("impact"),
            "recommended_action": proposal.get("recommended_action"),
            "risk_level": proposal.get("risk_level"),
            "auto_fix_policy": proposal.get("auto_fix_policy"),
            "confidence": proposal.get("confidence"),
            "blocking": proposal.get("blocking"),
            "evidence_refs": deepcopy(proposal.get("evidence_refs", [])),
            "decision": {
                "status": "pending",
                "allows_continue": False if proposal.get("blocking") else True,
                "selected_action": None,
                "selected_action_sha256": None,
                "decided_by": None,
                "decided_at": None,
                "comment": None,
            },
        }
        items.append(item)
    counts = count_manual_review_items(items)
    return {
        "schema_id": "manual-review-items",
        "schema_version": "1.0.0",
        "contract_version": "v4",
        "run_id": run_id,
        "required": bool(items),
        "status": "pending" if counts["pending_count"] else "not_required",
        "items_path": MANUAL_REVIEW_ITEMS_PATH,
        "summary_path": MANUAL_REVIEW_SUMMARY_PATH,
        "generated_at": generated_at,
        "writer": writer,
        "items": items,
    }


def write_manual_review_items(
    run_dir: Path,
    manual_review_items: dict[str, Any],
    *,
    writer: str = "format-helper",
) -> dict[str, Any]:
    """主控唯一写入 plans/manual_review_items.json。"""
    if writer != "format-helper":
        raise ValueError("plans/manual_review_items.json 只能由 format-helper 主控写入")
    validation = validate_manual_review_items(manual_review_items)
    if not validation.valid:
        raise ValueError(f"manual-review-items 未通过校验：{validation.errors}")
    path = resolve_run_relative_path(run_dir, MANUAL_REVIEW_ITEMS_PATH)
    payload = deepcopy(manual_review_items)
    payload["items_path"] = MANUAL_REVIEW_ITEMS_PATH
    payload["writer"] = writer
    write_json_atomic(path, payload)
    return {
        "path": MANUAL_REVIEW_ITEMS_PATH,
        "sha256": compute_file_sha256(path),
        "size_bytes": path.stat().st_size,
        "counts": count_manual_review_items(payload.get("items", [])),
    }


def validate_manual_review_items(data: dict[str, Any]) -> ReviewValidationResult:
    """校验权威 manual-review-items 结构和决策状态。"""
    errors: list[str] = []
    required = {
        "schema_id",
        "schema_version",
        "contract_version",
        "run_id",
        "required",
        "status",
        "items_path",
        "summary_path",
        "generated_at",
        "items",
    }
    for field_name in sorted(required - data.keys()):
        errors.append(f"{field_name} is required")
    if data.get("schema_id") != "manual-review-items":
        errors.append("schema_id must be manual-review-items")
    if data.get("contract_version") != "v4":
        errors.append("contract_version must be v4")
    if data.get("items_path") != MANUAL_REVIEW_ITEMS_PATH:
        errors.append("items_path must be plans/manual_review_items.json")
    if data.get("writer") not in {None, "format-helper"}:
        errors.append("manual-review-items writer must be format-helper")

    items = data.get("items")
    if not isinstance(items, list):
        errors.append("items must be array")
        return ReviewValidationResult(False, errors)

    seen_review_ids: set[str] = set()
    for index, item in enumerate(items):
        if not isinstance(item, dict):
            errors.append(f"items[{index}] must be object")
            continue
        review_id = item.get("review_id")
        if not review_id:
            errors.append(f"items[{index}].review_id is required")
        elif review_id in seen_review_ids:
            errors.append(f"duplicate review_id: {review_id}")
        seen_review_ids.add(str(review_id))
        source_refs = item.get("source_refs")
        if not isinstance(source_refs, list) or not source_refs:
            errors.append(f"{review_id}.source_refs must be non-empty array")
        source_issue_ids = item.get("source_issue_ids")
        if not isinstance(source_issue_ids, list):
            errors.append(f"{review_id}.source_issue_ids must be array")
        if not isinstance(item.get("confidence"), (int, float)) or not 0 <= item.get("confidence") <= 1:
            errors.append(f"{review_id}.confidence must be between 0 and 1")
        decision = item.get("decision")
        if not isinstance(decision, dict):
            errors.append(f"{review_id}.decision must be object")
            continue
        status = decision.get("status")
        if status not in DECISION_STATUSES:
            errors.append(f"{review_id}.decision.status is not allowed: {status}")
        if "allows_continue" not in decision:
            errors.append(f"{review_id}.decision.allows_continue is required")
        if status == "pending" and item.get("blocking") is True and decision.get("allows_continue") is not False:
            errors.append(f"{review_id}.pending blocking decision must not allow continue")
        if status in {"approved", "modified"} and decision.get("selected_action") is None:
            errors.append(f"{review_id}.approved/modified requires selected_action")
        if status in {"rejected", "deferred_non_blocking", "not_applicable"} and decision.get("selected_action") is not None:
            errors.append(f"{review_id}.{status} must use selected_action=null")
        if status == "approved" and decision.get("selected_action") is None:
            errors.append(f"{review_id}.not_applicable must be used when no action is selected")
        errors.extend(
            f"{review_id}.{error}"
            for error in validate_selected_action(
                decision.get("selected_action"),
                decision_status=status,
                expected_sha256=decision.get("selected_action_sha256"),
            )
        )
    return ReviewValidationResult(not errors, errors)


def manual_review_items_ref(run_dir: Path, captured_at: str = "2026-05-08T00:00:00+08:00") -> dict[str, Any]:
    """从真实权威清单生成 repair-plan 引用对象。"""
    path = resolve_run_relative_path(run_dir, MANUAL_REVIEW_ITEMS_PATH)
    data = read_json(path)
    counts = count_manual_review_items(data.get("items", []))
    return {
        "ref_state": "finalized",
        "path": MANUAL_REVIEW_ITEMS_PATH,
        "sha256": compute_file_sha256(path),
        "size_bytes": path.stat().st_size,
        **counts,
        "captured_at": captured_at,
    }


def build_decision_snapshot(
    run_dir: Path,
    *,
    snapshot_id: str = "DS-001",
    captured_at: str = "2026-05-08T00:00:00+08:00",
) -> dict[str, Any]:
    """从权威 manual-review-items 生成最小决策快照。"""
    path = resolve_run_relative_path(run_dir, MANUAL_REVIEW_ITEMS_PATH)
    data = read_json(path)
    counts = count_manual_review_items(data.get("items", []))
    decisions = []
    approved_review_ids: list[str] = []
    modified_review_ids: list[str] = []
    rejected_review_ids: list[str] = []
    for item in data.get("items", []):
        decision = item.get("decision", {})
        status = decision.get("status")
        if status == "approved":
            approved_review_ids.append(item["review_id"])
        if status == "modified":
            modified_review_ids.append(item["review_id"])
        if status == "rejected":
            rejected_review_ids.append(item["review_id"])
        decisions.append(
            {
                "review_id": item.get("review_id"),
                "decision_status": status,
                "allows_continue": decision.get("allows_continue"),
                "selected_action": deepcopy(decision.get("selected_action")),
                "selected_action_sha256": decision.get("selected_action_sha256"),
                "decided_by": decision.get("decided_by"),
                "decided_at": decision.get("decided_at"),
                "source_refs": deepcopy(item.get("source_refs", [])),
            }
        )
    snapshot = {
        "snapshot_id": snapshot_id,
        "items_path": MANUAL_REVIEW_ITEMS_PATH,
        "items_sha256": compute_file_sha256(path),
        "items_size_bytes": path.stat().st_size,
        **counts,
        "approved_review_ids": approved_review_ids,
        "modified_review_ids": modified_review_ids,
        "rejected_review_ids": rejected_review_ids,
        "allows_continue": counts["unresolved_manual_review_count"] == 0,
        "decisions": decisions,
        "captured_at": captured_at,
    }
    validation_errors = validate_decision_snapshot(snapshot)
    if validation_errors:
        raise ValueError(f"decision_snapshot 未通过校验：{validation_errors}")
    return snapshot


def validate_decision_snapshot(
    snapshot: dict[str, Any],
    *,
    risk_policy: dict[str, Any] | None = None,
    policy_sha256: str | None = None,
) -> list[str]:
    """校验 decision_snapshot 的 selected_action required-but-nullable 与 hash。"""
    errors: list[str] = []
    for index, decision in enumerate(snapshot.get("decisions", [])):
        if not isinstance(decision, dict):
            errors.append(f"decisions[{index}] must be object")
            continue
        status = decision.get("decision_status")
        errors.extend(
            f"decisions[{index}].{error}"
            for error in validate_selected_action(
                decision.get("selected_action"),
                decision_status=status,
                expected_sha256=decision.get("selected_action_sha256"),
                risk_policy=risk_policy,
                policy_sha256=policy_sha256,
            )
        )
        if status in SELECTED_ACTION_DECISION_STATUSES:
            if not decision.get("decided_by") or not decision.get("decided_at"):
                errors.append(f"decisions[{index}].decided_by/decided_at are required")
    return errors


def compute_plan_revision(plan: dict[str, Any]) -> int:
    """由 finalized 输入 hash 集合确定性派生 plan_revision。"""
    payload = {
        "manual_review_items_sha256": plan.get("manual_review_items_ref", {}).get("sha256"),
        "decision_snapshot_items_sha256": plan.get("decision_snapshot", {}).get("items_sha256"),
        "risk_policy_sha256": plan.get("risk_policy_ref", {}).get("sha256"),
        "source_audit_sha256": [item.get("sha256") for item in plan.get("source_audit_refs", [])],
    }
    return int(sha256_text(canonical_json(payload))[:8], 16)


def validate_policy_match(action: dict[str, Any], risk_policy: dict[str, Any], policy_sha256: str) -> list[str]:
    """校验 action_whitelist 写回证明。"""
    errors: list[str] = []
    policy_ref = action.get("policy_match_ref")
    if not isinstance(policy_ref, dict):
        return ["policy_match_ref must be object"]
    if policy_ref.get("policy_sha256") != policy_sha256:
        errors.append("policy_match_ref.policy_sha256 must match risk-policy file")
    if action.get("allowed_by_policy") is True:
        if policy_ref.get("source_kind") != "action_whitelist":
            errors.append("allowed_by_policy=true requires policy_match_ref.source_kind=action_whitelist")
        if policy_ref.get("decision_kind") != "write_allowed":
            errors.append("action_whitelist policy must use decision_kind=write_allowed")
        whitelist_id = policy_ref.get("whitelist_id")
        whitelist = risk_policy.get("action_whitelist", [])
        matched = next((item for item in whitelist if item.get("whitelist_id") == whitelist_id), None)
        if not matched:
            errors.append("policy_match_ref.whitelist_id must resolve to risk-policy.action_whitelist")
        else:
            target = action.get("target", {})
            if matched.get("operation") != action.get("operation"):
                errors.append("whitelist operation does not match action.operation")
            if matched.get("action_type") != action.get("action_type"):
                errors.append("whitelist action_type does not match action.action_type")
            if matched.get("target_attribute") != target.get("attribute"):
                errors.append("whitelist target_attribute does not match action.target.attribute")
    return errors


def validate_repair_plan_v4(
    plan: dict[str, Any],
    *,
    run_dir: Path | None = None,
    risk_policy: dict[str, Any] | None = None,
) -> ReviewValidationResult:
    """校验 CODE-009 draft/finalized repair-plan 闭环。"""
    errors: list[str] = []
    required = {
        "schema_id",
        "schema_version",
        "contract_version",
        "run_id",
        "plan_id",
        "plan_state",
        "plan_revision",
        "source_audit_paths",
        "source_audit_refs",
        "risk_policy_path",
        "risk_policy_ref",
        "manual_review_items_ref",
        "decision_snapshot",
        "actions",
        "manual_review_required",
        "generated_at",
    }
    for field_name in sorted(required - plan.keys()):
        errors.append(f"{field_name} is required")
    if plan.get("schema_id") != "repair-plan":
        errors.append("schema_id must be repair-plan")
    if plan.get("contract_version") != "v4":
        errors.append("contract_version must be v4")

    actions = plan.get("actions")
    if not isinstance(actions, list):
        errors.append("actions must be array")
        actions = []
    plan_state = plan.get("plan_state")
    if plan_state == "draft":
        if plan.get("plan_revision") != 0:
            errors.append("draft plan_revision must be 0")
        ref_state = plan.get("manual_review_items_ref", {}).get("ref_state")
        if ref_state not in {"absent", "draft"}:
            errors.append("draft manual_review_items_ref.ref_state must be absent or draft")
        if plan.get("decision_snapshot") is not None:
            errors.append("draft decision_snapshot must be null")
        for action in actions:
            if action.get("execution_status") == "executable":
                errors.append(f"{action.get('action_id')}.draft action must not be executable")
    elif plan_state == "finalized":
        if not isinstance(plan.get("plan_revision"), int) or plan.get("plan_revision") <= 0:
            errors.append("finalized plan_revision must be positive integer")
        elif plan.get("plan_revision") != compute_plan_revision(plan):
            errors.append("finalized plan_revision must be derived from input hashes")
        if not plan.get("finalized_from_plan_id") or not plan.get("finalized_at"):
            errors.append("finalized plan requires finalized_from_plan_id and finalized_at")
        if plan.get("manual_review_items_ref", {}).get("ref_state") != "finalized":
            errors.append("finalized manual_review_items_ref.ref_state must be finalized")
        if not isinstance(plan.get("decision_snapshot"), dict):
            errors.append("finalized decision_snapshot must be object")
        else:
            snapshot = plan["decision_snapshot"]
            if snapshot.get("allows_continue") is not True:
                errors.append("finalized decision_snapshot.allows_continue must be true")
            for field_name in ("pending_count", "blocking_count", "unresolved_manual_review_count"):
                if snapshot.get(field_name) != 0:
                    errors.append(f"finalized decision_snapshot.{field_name} must be 0")
            policy_data_for_snapshot = risk_policy
            policy_sha_for_snapshot = plan.get("risk_policy_ref", {}).get("sha256")
            errors.extend(
                f"decision_snapshot.{error}"
                for error in validate_decision_snapshot(
                    snapshot,
                    risk_policy=policy_data_for_snapshot,
                    policy_sha256=policy_sha_for_snapshot,
                )
            )
        if run_dir is not None:
            errors.extend(_validate_finalized_refs(plan, run_dir))
        risk_policy_path = plan.get("risk_policy_path")
        policy_sha256 = plan.get("risk_policy_ref", {}).get("sha256")
        if run_dir is not None and risk_policy_path:
            policy_path = resolve_run_relative_path(run_dir, risk_policy_path)
            if not policy_path.exists():
                errors.append("risk_policy_path does not exist")
            elif compute_file_sha256(policy_path) != policy_sha256:
                errors.append("risk_policy_ref.sha256 must match risk_policy_path")
        policy_data = risk_policy
        if policy_data is None and run_dir is not None and risk_policy_path:
            try:
                policy_data = read_json(resolve_run_relative_path(run_dir, risk_policy_path))
            except (OSError, ValueError, json.JSONDecodeError):
                policy_data = None
        for action in actions:
            action_id = action.get("action_id")
            if action.get("execution_status") not in FINALIZED_ALLOWED_EXECUTION_STATUSES:
                errors.append(f"{action_id}.finalized execution_status is not allowed")
            if action.get("execution_status") == "executable":
                if action.get("allowed_by_policy") is not True:
                    errors.append(f"{action_id}.executable requires allowed_by_policy=true")
                if policy_data is None or policy_sha256 is None:
                    errors.append(f"{action_id}.executable requires risk-policy for whitelist recompute")
                else:
                    errors.extend(f"{action_id}.{error}" for error in validate_policy_match(action, policy_data, policy_sha256))
                if action.get("requires_manual_review") is True and action.get("manual_review_id"):
                    if not _manual_review_allows_continue(plan, action.get("manual_review_id")):
                        errors.append(f"{action_id}.manual_review_id does not allow continue")
            if action.get("allowed_by_policy") is True:
                policy_ref = action.get("policy_match_ref", {})
                if policy_ref.get("source_kind") != "action_whitelist":
                    errors.append(f"{action_id}.allowed_by_policy true only allowed by action_whitelist")
    else:
        errors.append("plan_state must be draft or finalized")
    return ReviewValidationResult(not errors, errors)


def _validate_finalized_refs(plan: dict[str, Any], run_dir: Path) -> list[str]:
    """校验 finalized plan 对 manual-review-items 的真实 hash/size/计数引用。"""
    errors: list[str] = []
    ref = plan.get("manual_review_items_ref", {})
    snapshot = plan.get("decision_snapshot", {})
    try:
        expected_ref = manual_review_items_ref(run_dir, ref.get("captured_at") or "2026-05-08T00:00:00+08:00")
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        return [f"manual_review_items_ref cannot resolve: {exc}"]
    for field_name in ("path", "sha256", "size_bytes", "pending_count", "blocking_count", "unresolved_manual_review_count"):
        if ref.get(field_name) != expected_ref.get(field_name):
            errors.append(f"manual_review_items_ref.{field_name} must match real manual-review-items")
    for ref_field, snapshot_field in (
        ("sha256", "items_sha256"),
        ("size_bytes", "items_size_bytes"),
        ("pending_count", "pending_count"),
        ("blocking_count", "blocking_count"),
        ("unresolved_manual_review_count", "unresolved_manual_review_count"),
    ):
        if ref.get(ref_field) != snapshot.get(snapshot_field):
            errors.append(f"decision_snapshot.{snapshot_field} must match manual_review_items_ref.{ref_field}")
    return errors


def _manual_review_allows_continue(plan: dict[str, Any], review_id: str) -> bool:
    snapshot = plan.get("decision_snapshot", {})
    for decision in snapshot.get("decisions", []):
        if decision.get("review_id") == review_id:
            return decision.get("allows_continue") is True
    return False


def finalized_plan_path(plan_revision: int) -> str:
    """生成 revisioned canonical finalized plan 路径。"""
    if not isinstance(plan_revision, int) or plan_revision <= 0:
        raise ValueError("plan_revision must be positive integer")
    return f"plans/repair_plan.finalized.r{plan_revision}.yaml"


def write_repair_plan(run_dir: Path, plan: dict[str, Any]) -> dict[str, Any]:
    """按 draft/finalized canonical 路径写入 repair-plan。"""
    validation = validate_repair_plan_v4(plan, run_dir=run_dir)
    if not validation.valid:
        raise ValueError(f"repair-plan 未通过校验：{validation.errors}")
    if plan.get("plan_state") == "draft":
        rel_path = DRAFT_REPAIR_PLAN_PATH
    elif plan.get("plan_state") == "finalized":
        rel_path = finalized_plan_path(plan["plan_revision"])
    else:
        raise ValueError("plan_state must be draft or finalized")
    path = resolve_run_relative_path(run_dir, rel_path)
    write_yaml(path, plan)
    return {"path": rel_path, "sha256": compute_file_sha256(path), "size_bytes": path.stat().st_size}


__all__ = [
    "DRAFT_REPAIR_PLAN_PATH",
    "MANUAL_REVIEW_ITEMS_PATH",
    "MANUAL_REVIEW_SUMMARY_PATH",
    "build_decision_snapshot",
    "build_manual_review_items",
    "apply_manual_review_decision",
    "compute_selected_action_id",
    "compute_selected_action_sha256",
    "compute_plan_revision",
    "count_manual_review_items",
    "finalized_plan_path",
    "manual_review_items_ref",
    "validate_manual_review_items",
    "validate_decision_snapshot",
    "validate_policy_match",
    "validate_repair_plan_v4",
    "validate_selected_action",
    "write_manual_review_items",
    "write_repair_plan",
]
