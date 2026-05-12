"""Gate 推进谓词与真值表判定（CODE-007）。

CODE-018 扩展：slot_facts / rule_confirmation_gate 槽位契约谓词。
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from scripts.validation.skill_result_io import validate_common_skill_result_contract
from scripts.validation.validate_run_state import validate_run_state


@dataclass
class GateDecision:
    """Gate 判定结果。"""

    can_advance: bool
    next_action_allowed: bool
    outcome: str
    reason: str
    blockers: list[str] = field(default_factory=list)
    recommended_state_status: str | None = None
    synthetic_failure_required: bool = False


def human_review_state(result: dict[str, Any]) -> str:
    """归一化 human_review 状态。"""
    review = result.get("human_review")
    if not isinstance(review, dict) or review.get("required") is False:
        return "not_required"
    status = review.get("status") or review.get("decision_status")
    if status in {"cleared", "not_required"}:
        return status
    if status in {"blocked", "rejected"}:
        return "blocked"
    items = review.get("items", [])
    if isinstance(items, list):
        for item in items:
            if not isinstance(item, dict):
                continue
            decision = item.get("decision", {})
            if item.get("blocking") is True and isinstance(decision, dict):
                if decision.get("status") == "pending":
                    return "pending"
                if decision.get("allows_continue") is False:
                    return "blocked"
    return "pending" if review.get("required") else "not_required"


def evidence_refs_resolve(result: dict[str, Any], manifest: dict[str, Any] | None) -> bool:
    """最小 evidence_refs 解析：空引用通过，非空必须能在 manifest 中解析。"""
    refs = result.get("evidence_refs", [])
    if not refs:
        return True
    if not isinstance(manifest, dict):
        return False
    known_ids: set[str] = set()
    for key in ("evidence_refs", "evidence", "items"):
        for item in manifest.get(key, []):
            if isinstance(item, dict):
                for id_key in ("evidence_id", "id"):
                    if item.get(id_key):
                        known_ids.add(item[id_key])
    for ref in refs:
        if isinstance(ref, str):
            if ref not in known_ids:
                return False
        elif isinstance(ref, dict):
            evidence_id = ref.get("evidence_id") or ref.get("id")
            if evidence_id and evidence_id not in known_ids:
                return False
    return True


def next_action_is_valid_for_stage(result: dict[str, Any]) -> bool:
    """校验 done 结果是否有可消费的下一步。"""
    next_action = result.get("next_action", {})
    if not isinstance(next_action, dict):
        return False
    if next_action.get("kind") in {"wait_user", "retry", "manual_recover"}:
        return False
    if next_action.get("kind") == "run_skill":
        return bool(next_action.get("stage") and next_action.get("skill_name"))
    return next_action.get("kind") in {"finalize", "stop"}


def can_advance(
    result: dict[str, Any],
    state: dict[str, Any],
    manifest: dict[str, Any] | None = None,
    *,
    run_dir: Path | None = None,
) -> bool:
    """41-§6.5 最小推进谓词。"""
    return evaluate_gate(result, state, manifest, run_dir=run_dir).can_advance


def evaluate_gate(
    result: dict[str, Any],
    state: dict[str, Any],
    manifest: dict[str, Any] | None = None,
    *,
    run_dir: Path | None = None,
) -> GateDecision:
    """按 40-§5.8 真值表和 41-§6.5 判定是否允许进入 next_action。"""
    blockers: list[str] = []
    schema_errors = validate_common_skill_result_contract(result, run_dir=run_dir)
    state_validation = validate_run_state(state)

    if schema_errors:
        return GateDecision(
            can_advance=False,
            next_action_allowed=False,
            outcome="blocked",
            reason="skill-result schema or artifact validation failed",
            blockers=schema_errors,
            recommended_state_status="blocked",
            synthetic_failure_required=result.get("status") == "synthetic_failure",
        )
    if not state_validation.valid:
        return GateDecision(
            can_advance=False,
            next_action_allowed=False,
            outcome="blocked",
            reason="run-state validation failed",
            blockers=[str(error) for error in state_validation.errors],
            recommended_state_status="blocked",
        )

    review_state = human_review_state(result)
    status = result.get("status")
    result_validation = result.get("validation", {})
    schema_valid = result.get("schema_valid") is True and result_validation.get("schema_valid") is True
    validation_checks_passed = all(
        result_validation.get(field) is True
        for field in ("schema_valid", "path_valid", "risk_policy_valid", "evidence_valid")
    )
    gate_passed = result.get("gate_passed") is True and result.get("gate_check", {}).get("passed") is True
    result_blockers = result.get("blockers", [])
    error_code = result.get("error", {}).get("code") if isinstance(result.get("error"), dict) else None
    manifest_status = manifest.get("status") if isinstance(manifest, dict) else None

    if status == "synthetic_failure":
        return GateDecision(
            can_advance=False,
            next_action_allowed=False,
            outcome="blocked",
            reason="synthetic_failure never advances",
            blockers=["synthetic_failure"],
            recommended_state_status="blocked",
            synthetic_failure_required=True,
        )
    if not schema_valid:
        return GateDecision(False, False, "blocked", "schema invalid", ["schema_invalid"], "blocked")
    if not validation_checks_passed:
        failed_fields = [
            field for field in ("path_valid", "risk_policy_valid", "evidence_valid")
            if result_validation.get(field) is not True
        ]
        return GateDecision(False, False, "blocked", "validation checks failed", failed_fields, "blocked")
    if status == "waiting_user" and result.get("gate_passed") is False and review_state == "pending":
        if result.get("next_action", {}).get("kind") == "wait_user" and not result_blockers and error_code is None:
            return GateDecision(False, False, "waiting_user", "manual review pending", [], "waiting_user")
        return GateDecision(False, False, "blocked", "waiting_user shape invalid", ["waiting_user_invalid"], "blocked")
    if status == "blocked":
        return GateDecision(False, False, "blocked", "result blocked", ["blocked_result"], "blocked")
    if status != "done":
        return GateDecision(False, False, "blocked", "status is not done", [f"status={status}"], "blocked")
    if not gate_passed:
        return GateDecision(False, False, "blocked", "gate did not pass", ["gate_not_passed"], "blocked")
    if result_blockers:
        return GateDecision(False, False, "blocked", "blockers present", ["blockers_present"], "blocked")
    if error_code is not None:
        return GateDecision(False, False, "blocked", "error code present", ["error_present"], "blocked")
    if review_state == "pending":
        return GateDecision(False, False, "waiting_user", "manual review pending with done result", ["manual_review_pending"], "waiting_user")
    if review_state == "blocked":
        return GateDecision(False, False, "blocked", "manual review blocked", ["manual_review_blocked"], "blocked")
    if state.get("schema_id") != "run-state":
        return GateDecision(False, False, "blocked", "state schema_id must be run-state", ["state_schema_id"], "blocked")
    if state.get("run_id") != result.get("run_id"):
        return GateDecision(False, False, "blocked", "run_id mismatch", ["run_id_mismatch"], "blocked")
    if state.get("stage") != result.get("stage"):
        return GateDecision(False, False, "blocked", "stage mismatch", ["stage_mismatch"], "blocked")
    if manifest_status == "broken":
        return GateDecision(False, False, "blocked", "manifest broken", ["manifest_broken"], "blocked")
    if not evidence_refs_resolve(result, manifest):
        return GateDecision(False, False, "blocked", "evidence refs unresolved", ["evidence_refs_unresolved"], "blocked")
    if not next_action_is_valid_for_stage(result):
        return GateDecision(False, False, "blocked", "next_action invalid for stage", ["next_action_invalid"], "blocked")

    # CODE-018: stage=rule_packaging 额外检查 slot facts (SLOT_CONTRACT_DESIGN.md §9.2)
    if result.get("stage") == "rule_packaging" and run_dir is not None:
        slot_facts_blockers = _check_slot_facts_for_stage(result, state, run_dir)
        if slot_facts_blockers:
            return GateDecision(
                can_advance=False,
                next_action_allowed=False,
                outcome="blocked",
                reason="slot facts unresolved or confirmation gate not cleared",
                blockers=slot_facts_blockers,
                recommended_state_status=state.get("status", "wip"),
            )

    # P1-3: waiting_user 通过 Gate 后提示主控将状态转回 wip
    next_status = state.get("status")
    if next_status == "waiting_user" and state.get("stage") == result.get("stage"):
        next_status = "wip"

    return GateDecision(
        can_advance=True,
        next_action_allowed=True,
        outcome="advance",
        reason="all gate predicates passed",
        recommended_state_status=next_status,
    )


# ── CODE-018：槽位契约谓词 ──────────────────────────────────────────


def _load_run_relative_json(run_dir: Path, rel_path: str) -> dict[str, Any] | None:
    """加载 run-relative JSON 文件，不存在或不可解析时返回 None。"""
    target = run_dir / rel_path
    if not target.exists() or not target.is_file():
        return None
    try:
        return json.loads(target.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return None


def _slot_facts_path() -> str:
    return "semantic/role_format_slot_facts.json"


def _rule_confirmation_gate_path() -> str:
    return "logs/rule_confirmation_gate.json"


def is_slot_facts_resolved(slot_facts: dict[str, Any] | None) -> bool:
    """slot_facts.gate_status == 'passed' 且无未解决必需槽位。"""
    if not isinstance(slot_facts, dict):
        return False
    if slot_facts.get("gate_status") != "passed":
        return False
    roles = slot_facts.get("roles", [])
    if not isinstance(roles, list):
        return False
    for role in roles:
        if not isinstance(role, dict):
            continue
        slot_summary = role.get("slot_summary", {})
        if not isinstance(slot_summary, dict):
            continue
        for _slot_name, summary in slot_summary.items():
            if not isinstance(summary, dict):
                continue
            if summary.get("status") in {"unresolved", "conflict"}:
                return False
    return True


def is_rule_confirmation_cleared(gate_json: dict[str, Any] | None) -> bool:
    """rule_confirmation_gate.status == 'cleared'。"""
    if not isinstance(gate_json, dict):
        return False
    return gate_json.get("status") == "cleared"


def validate_slot_contract_compliance(
    slot_facts: dict[str, Any] | None,
    contract: dict[str, Any] | None,
) -> "PredicateResult":
    """检查 slot_facts 是否满足 contract 中声明的 required_slots。

    返回 PredicateResult(errors=[]) 表示符合。
    """
    errors: list[str] = []
    if not isinstance(slot_facts, dict):
        return PredicateResult(False, ["slot_facts 缺失或格式错误"])
    if not isinstance(contract, dict):
        return PredicateResult(False, ["contract 缺失或格式错误"])

    role_contracts = contract.get("role_slot_contracts")
    if not isinstance(role_contracts, dict):
        return PredicateResult(False, ["contract.role_slot_contracts 缺失"])

    facts_roles = {r.get("role_kind"): r for r in slot_facts.get("roles", []) if isinstance(r, dict)}

    for role_kind, rc in role_contracts.items():
        if not isinstance(rc, dict):
            errors.append(f"contract.role_slot_contracts.{role_kind} is not an object")
            continue
        required_slots = rc.get("required_slots", [])
        if not isinstance(required_slots, list):
            continue
        if role_kind not in facts_roles:
            errors.append(f"role '{role_kind}' 在 slot_facts 中缺失")
            continue
        role = facts_roles[role_kind]
        slot_summary = role.get("slot_summary", {})
        if not isinstance(slot_summary, dict):
            errors.append(f"role '{role_kind}' slot_summary 缺失")
            continue
        for slot_name in required_slots:
            summary = slot_summary.get(slot_name)
            if not isinstance(summary, dict):
                errors.append(f"role '{role_kind}' required_slot '{slot_name}' 缺失")
                continue
            status = summary.get("status", "")
            if status not in {"resolved", "resolved_with_conflicts", "not_applicable", "user_confirmed"}:
                errors.append(
                    f"role '{role_kind}' required_slot '{slot_name}' status={status or 'missing'}，"
                    f"必须是 resolved/resolved_with_conflicts/not_applicable/user_confirmed"
                )

    return PredicateResult(len(errors) == 0, errors)


def _check_slot_facts_for_stage(
    result: dict[str, Any],
    state: dict[str, Any],
    run_dir: Path,
) -> list[str]:
    """在 rule_packaging 阶段检查 slot facts 与 confirmation gate。"""
    blockers: list[str] = []
    slot_facts = _load_run_relative_json(run_dir, _slot_facts_path())
    if slot_facts is None:
        blockers.append("slot_facts 不存在，无法推进 rule_packaging")
        return blockers

    status = state.get("status", "")
    if status == "waiting_user":
        gate_json = _load_run_relative_json(run_dir, _rule_confirmation_gate_path())
        if not is_rule_confirmation_cleared(gate_json):
            blockers.append("rule_confirmation_gate 未 cleared，用户决策未完成")
            if gate_json is None:
                blockers.append("rule_confirmation_gate.json 不存在")
    elif not is_slot_facts_resolved(slot_facts):
        blockers.append("slot_facts 存在 unresolved 必需槽位，无法自动推进")

    return blockers


@dataclass
class PredicateResult:
    """公共谓词结果。"""

    valid: bool
    errors: list[str] = field(default_factory=list)
