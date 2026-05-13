"""Gate 推进谓词与真值表判定（CODE-007）。

CODE-018 扩展：slot_facts / rule_confirmation_gate 槽位契约谓词。
"""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from scripts.validation.skill_result_io import validate_common_skill_result_contract
from scripts.utils.simple_yaml import load_yaml
from scripts.validation.validate_run_state import validate_run_state

TARGET_ROLE_MAP_PATH = "semantic/semantic_role_map.before.json"
ROLE_SLOT_CONTRACT_PATH = "docs/v4/schemas/role_slot_contract.yaml"
SHA256_HEX_PATTERN = re.compile(r"^[0-9a-f]{64}$")
ALLOWED_RESOLVER_REASON_CODES = {
    "ROLEMAP_CONSISTENT",
    "ROLEMAP_CONFLICT_WARNING",
    "ROLEMAP_CONFLICT_BLOCKED",
    "LEGACY_ROLE_DIAGNOSTIC",
    "STRUCTURE_RULE_USED",
    "STRUCTURE_RULE_WARNING",
    "STRUCTURE_RULE_BLOCKED",
    "USER_CONFIRMED",
    "UNRESOLVED_MAPPING",
    "LOW_CONFIDENCE",
    "MULTI_ROLE_CONFLICT",
}


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


def _required_slots_by_role(contract: dict[str, Any] | None = None) -> dict[str, set[str]]:
    """从契约提取每个角色的必需槽位集合。"""
    if not isinstance(contract, dict):
        return {}
    role_contracts = contract.get("role_slot_contracts", {})
    if not isinstance(role_contracts, dict):
        return {}
    required: dict[str, set[str]] = {}
    for role_kind, role_contract in role_contracts.items():
        if isinstance(role_contract, dict) and isinstance(role_contract.get("required_slots"), list):
            required[str(role_kind)] = {str(slot) for slot in role_contract["required_slots"]}
    return required


def is_slot_facts_resolved(slot_facts: dict[str, Any] | None, contract: dict[str, Any] | None = None) -> bool:
    """slot_facts.gate_status == passed 且必需槽位无未解决状态。"""
    if not isinstance(slot_facts, dict):
        return False
    if not isinstance(contract, dict):
        return False
    if slot_facts.get("gate_status") != "passed":
        return False
    roles = slot_facts.get("roles", [])
    if not isinstance(roles, list):
        return False
    required_by_role = _required_slots_by_role(contract)
    for role in roles:
        if not isinstance(role, dict):
            continue
        role_kind = str(role.get("role_kind") or "")
        required_slots = required_by_role.get(role_kind)
        slot_summary = role.get("slot_summary", {})
        if not isinstance(slot_summary, dict):
            continue
        for _slot_name, summary in slot_summary.items():
            if required_slots is not None and _slot_name not in required_slots:
                continue
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


def compute_file_sha256(path: Path) -> str:
    """计算文件 sha256。"""
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _load_target_role_map(run_dir: Path, rel_path: str) -> tuple[dict[str, Any] | None, str | None]:
    """读取 target_role_ref 指向的 role-map。"""
    target = run_dir / rel_path
    if not target.exists() or not target.is_file():
        return None, f"FH-SLOT-FACTS-INVALID: target_role_ref.path 文件不存在: {rel_path}"
    try:
        return json.loads(target.read_text(encoding="utf-8")), None
    except json.JSONDecodeError:
        return None, f"FH-SLOT-FACTS-INVALID: target_role_ref.path 文件不是合法 JSON: {rel_path}"


def _role_ids_from_role_map(role_map: dict[str, Any]) -> set[str]:
    """提取 semantic_role_map.before.json 中的 roles[].role_id。"""
    roles = role_map.get("roles", [])
    if not isinstance(roles, list):
        return set()
    return {
        str(role["role_id"])
        for role in roles
        if isinstance(role, dict) and isinstance(role.get("role_id"), str) and role.get("role_id")
    }


def _role_map_slot_kinds(role_map: dict[str, Any]) -> dict[str, str]:
    """提取 role-map role_id 到 slot_role_kind 的绑定。"""
    roles = role_map.get("roles", [])
    if not isinstance(roles, list):
        return {}
    result: dict[str, str] = {}
    for role in roles:
        if not isinstance(role, dict):
            continue
        role_id = role.get("role_id")
        slot_role_kind = role.get("slot_role_kind")
        if isinstance(role_id, str) and role_id and isinstance(slot_role_kind, str) and slot_role_kind:
            result[role_id] = slot_role_kind
    return result


def validate_target_role_refs(slot_facts: dict[str, Any] | None, run_dir: Path | None) -> "PredicateResult":
    """校验 role_format_slot_facts.roles[].target_role_ref 指向权威 role-map。"""
    errors: list[str] = []
    if not isinstance(slot_facts, dict):
        return PredicateResult(False, ["FH-SLOT-FACTS-INVALID: slot_facts 缺失或格式错误"])
    if run_dir is None:
        return PredicateResult(False, ["FH-SLOT-FACTS-INVALID: run_dir 缺失，无法复算 target_role_ref hash"])

    role_maps: dict[str, tuple[dict[str, Any] | None, set[str], dict[str, str], str | None]] = {}
    roles = slot_facts.get("roles", [])
    if not isinstance(roles, list):
        return PredicateResult(False, ["FH-SLOT-FACTS-INVALID: slot_facts.roles 缺失或格式错误"])

    for index, role in enumerate(roles):
        if not isinstance(role, dict):
            errors.append(f"FH-SLOT-FACTS-INVALID: roles[{index}] 不是对象")
            continue
        role_path = f"roles[{index}]"
        role_id = role.get("role_id")
        ref = role.get("target_role_ref")
        if not isinstance(ref, dict):
            errors.append(f"FH-SLOT-FACTS-INVALID: {role_path}.target_role_ref 必须是对象")
            continue

        ref_path = ref.get("path")
        path_kind = ref.get("path_kind")
        ref_role_id = ref.get("role_id")
        ref_sha256 = ref.get("sha256")
        if ref_path != TARGET_ROLE_MAP_PATH:
            errors.append(
                f"FH-SLOT-FACTS-INVALID: {role_path}.target_role_ref.path 必须为 {TARGET_ROLE_MAP_PATH}"
            )
        if path_kind != "run_relative":
            errors.append(f"FH-SLOT-FACTS-INVALID: {role_path}.target_role_ref.path_kind 必须为 run_relative")
        if not isinstance(ref_role_id, str) or not ref_role_id:
            errors.append(f"FH-SLOT-FACTS-INVALID: {role_path}.target_role_ref.role_id 缺失或无效")
        if role_id != ref_role_id:
            errors.append(f"FH-SLOT-FACTS-INVALID: {role_path}.role_id 必须与 target_role_ref.role_id 一致")
        sha256_valid = isinstance(ref_sha256, str) and SHA256_HEX_PATTERN.match(ref_sha256) is not None
        if not sha256_valid:
            errors.append(f"FH-SLOT-FACTS-INVALID: {role_path}.target_role_ref.sha256 必须为 64 位 hex")

        if ref_path != TARGET_ROLE_MAP_PATH or not isinstance(ref_path, str):
            continue
        if ref_path not in role_maps:
            role_map, load_error = _load_target_role_map(run_dir, ref_path)
            role_ids = _role_ids_from_role_map(role_map) if isinstance(role_map, dict) else set()
            slot_kinds = _role_map_slot_kinds(role_map) if isinstance(role_map, dict) else {}
            role_maps[ref_path] = (role_map, role_ids, slot_kinds, load_error)
        role_map, role_ids, slot_kinds, load_error = role_maps[ref_path]
        if load_error:
            errors.append(load_error)
            continue
        actual_sha256 = compute_file_sha256(run_dir / ref_path)
        if sha256_valid and ref_sha256 != actual_sha256:
            errors.append(
                f"FH-SLOT-FACTS-HASH-MISMATCH: {role_path}.target_role_ref.sha256 与 {ref_path} 实际 hash 不一致"
            )
        if ref_role_id not in role_ids:
            errors.append(
                f"FH-SLOT-FACTS-INVALID: {role_path}.target_role_ref.role_id 在 {ref_path} roles[].role_id 中不可解析"
            )
        if ref_role_id in role_ids and ref_role_id not in slot_kinds:
            errors.append(
                f"FH-SLOT-FACTS-INVALID: {role_path}.target_role_ref.role_id 缺少 {ref_path} roles[].slot_role_kind"
            )
        elif slot_kinds.get(ref_role_id) != role.get("role_kind"):
            errors.append(
                f"FH-SLOT-FACTS-INVALID: {role_path}.role_kind 必须与 {ref_path} roles[].slot_role_kind 一致"
            )

    return PredicateResult(len(errors) == 0, errors)


def validate_resolver_reasons(slot_facts: dict[str, Any] | None) -> "PredicateResult":
    """校验 resolver reasons[] 结构和 reason_code 枚举。"""
    errors: list[str] = []
    if not isinstance(slot_facts, dict):
        return PredicateResult(False, ["FH-SLOT-FACTS-INVALID: slot_facts 缺失或格式错误"])
    roles = slot_facts.get("roles", [])
    if not isinstance(roles, list):
        return PredicateResult(False, ["FH-SLOT-FACTS-INVALID: slot_facts.roles 缺失或格式错误"])

    for role_index, role in enumerate(roles):
        if not isinstance(role, dict):
            errors.append(f"FH-SLOT-FACTS-INVALID: roles[{role_index}] 不是对象")
            continue
        reasons = role.get("reasons")
        role_path = f"roles[{role_index}]"
        if not isinstance(reasons, list):
            errors.append(f"FH-SLOT-FACTS-INVALID: {role_path}.reasons 必须是数组")
            continue
        for reason_index, item in enumerate(reasons):
            item_path = f"{role_path}.reasons[{reason_index}]"
            if not isinstance(item, dict):
                errors.append(f"FH-SLOT-FACTS-INVALID: {item_path} 必须是对象")
                continue
            reason_code = item.get("reason_code")
            if reason_code not in ALLOWED_RESOLVER_REASON_CODES:
                errors.append(f"FH-SLOT-FACTS-INVALID: {item_path}.reason_code 未登记: {reason_code}")
            if not isinstance(item.get("message"), str) or not item.get("message"):
                errors.append(f"FH-SLOT-FACTS-INVALID: {item_path}.message 缺失或无效")
            if not isinstance(item.get("source"), str) or not item.get("source"):
                errors.append(f"FH-SLOT-FACTS-INVALID: {item_path}.source 缺失或无效")
            if "evidence_ref" not in item or not isinstance(item.get("evidence_ref"), (str, type(None))):
                errors.append(f"FH-SLOT-FACTS-INVALID: {item_path}.evidence_ref 缺失或无效")

    return PredicateResult(len(errors) == 0, errors)


def validate_slot_contract_compliance(
    slot_facts: dict[str, Any] | None,
    contract: dict[str, Any] | None,
) -> "PredicateResult":
    """检查 slot_facts 是否满足 contract 中声明的 required_slots。

    返回 PredicateResult(errors=[]) 表示符合。
    """
    errors: list[str] = []
    if not isinstance(slot_facts, dict):
        return PredicateResult(False, ["FH-SLOT-FACTS-INVALID: slot_facts 缺失或格式错误"])
    if not isinstance(contract, dict):
        return PredicateResult(False, ["FH-SLOT-CONTRACT-NOT-FOUND: contract 缺失或格式错误"])

    role_contracts = contract.get("role_slot_contracts")
    if not isinstance(role_contracts, dict):
        return PredicateResult(False, ["FH-SLOT-CONTRACT-NOT-FOUND: contract.role_slot_contracts 缺失"])

    facts_roles = {r.get("role_kind"): r for r in slot_facts.get("roles", []) if isinstance(r, dict)}

    for role_kind, rc in role_contracts.items():
        if not isinstance(rc, dict):
            errors.append(f"FH-SLOT-CONTRACT-NOT-FOUND: contract.role_slot_contracts.{role_kind} is not an object")
            continue
        required_slots = rc.get("required_slots", [])
        if not isinstance(required_slots, list):
            continue
        if role_kind not in facts_roles:
            errors.append(f"FH-SLOT-FACTS-MISSING: role '{role_kind}' 在 slot_facts 中缺失")
            continue
        role = facts_roles[role_kind]
        slot_summary = role.get("slot_summary", {})
        if not isinstance(slot_summary, dict):
            errors.append(f"FH-SLOT-FACTS-MISSING: role '{role_kind}' slot_summary 缺失")
            continue
        for slot_name in required_slots:
            summary = slot_summary.get(slot_name)
            if not isinstance(summary, dict):
                errors.append(f"FH-SLOT-FACTS-MISSING: role '{role_kind}' required_slot '{slot_name}' 缺失")
                continue
            status = summary.get("status", "")
            if status not in {"resolved", "resolved_with_conflicts", "not_applicable", "user_confirmed"}:
                code = "FH-SLOT-FACTS-CONFLICT" if status == "conflict" else "FH-SLOT-FACTS-UNRESOLVED"
                errors.append(
                    f"{code}: role '{role_kind}' required_slot '{slot_name}' status={status or 'missing'}，"
                    f"必须是 resolved/resolved_with_conflicts/not_applicable/user_confirmed"
                )
            if status in {"resolved", "resolved_with_conflicts"} and summary.get("confidence") == 0:
                errors.append(
                    f"FH-SLOT-FACTS-INVALID: role '{role_kind}' required_slot '{slot_name}' confidence=0 必须人工确认"
                )

    common_rules = contract.get("common_validation_rules")
    if not isinstance(common_rules, list) or not common_rules:
        errors.append("FH-SLOT-FACTS-INVALID: contract.common_validation_rules 缺失")
    else:
        for index, rule in enumerate(common_rules):
            if not isinstance(rule, dict):
                errors.append(f"FH-SLOT-FACTS-INVALID: common_validation_rules[{index}] 不是对象")
                continue
            if rule.get("condition") != "required_slot_confidence_eq_0":
                errors.append(
                    f"FH-SLOT-FACTS-INVALID: common_validation_rules[{index}].condition 未登记: {rule.get('condition')}"
                )
            if rule.get("severity") not in {"error", "warning"}:
                errors.append(f"FH-SLOT-FACTS-INVALID: common_validation_rules[{index}].severity 必须为 error 或 warning")
            if not isinstance(rule.get("blocks_confirmation_gate"), bool):
                errors.append(f"FH-SLOT-FACTS-INVALID: common_validation_rules[{index}].blocks_confirmation_gate 必须为 boolean")
            if rule.get("applies_to") != "all_roles":
                errors.append(f"FH-SLOT-FACTS-INVALID: common_validation_rules[{index}].applies_to 必须为 all_roles")
            if rule.get("applies_to_slots") != "required_slots":
                errors.append(f"FH-SLOT-FACTS-INVALID: common_validation_rules[{index}].applies_to_slots 必须为 required_slots")

    return PredicateResult(len(errors) == 0, errors)


def _load_contract_for_slot_facts(slot_facts: dict[str, Any], run_dir: Path) -> tuple[dict[str, Any] | None, str | None]:
    """加载 slot facts 声明的 role-slot contract。"""
    contract_ref = slot_facts.get("contract_ref") if isinstance(slot_facts, dict) else None
    contract_path = ROLE_SLOT_CONTRACT_PATH
    if isinstance(contract_ref, dict) and isinstance(contract_ref.get("contract_path"), str):
        contract_path = contract_ref["contract_path"]
    path = Path(contract_path)
    if not path.is_absolute():
        path = run_dir / path
        if not path.exists():
            path = Path(contract_path)
    if not path.exists() or not path.is_file():
        return None, f"FH-SLOT-CONTRACT-NOT-FOUND: contract 文件不存在: {contract_path}"
    try:
        data = load_yaml(path)
    except Exception as exc:
        return None, f"FH-SLOT-CONTRACT-NOT-FOUND: contract 无法解析: {exc}"
    if not isinstance(data, dict):
        return None, "FH-SLOT-CONTRACT-NOT-FOUND: contract 不是对象"
    return data, None


def _check_slot_facts_for_stage(
    result: dict[str, Any],
    state: dict[str, Any],
    run_dir: Path,
) -> list[str]:
    """在 rule_packaging 阶段检查 slot facts 与 confirmation gate。"""
    blockers: list[str] = []
    slot_facts = _load_run_relative_json(run_dir, _slot_facts_path())
    if slot_facts is None:
        blockers.append("FH-SLOT-FACTS-MISSING: slot_facts 不存在，无法推进 rule_packaging")
        return blockers

    target_role_refs = validate_target_role_refs(slot_facts, run_dir)
    if not target_role_refs.valid:
        blockers.extend(target_role_refs.errors)
        return blockers
    resolver_reasons = validate_resolver_reasons(slot_facts)
    if not resolver_reasons.valid:
        blockers.extend(resolver_reasons.errors)
        return blockers
    contract, contract_error = _load_contract_for_slot_facts(slot_facts, run_dir)
    if contract_error:
        blockers.append(contract_error)
        return blockers
    contract_result = validate_slot_contract_compliance(slot_facts, contract)
    if not contract_result.valid:
        blockers.extend(contract_result.errors)
        return blockers

    status = state.get("status", "")
    if status == "waiting_user":
        gate_json = _load_run_relative_json(run_dir, _rule_confirmation_gate_path())
        if not is_rule_confirmation_cleared(gate_json):
            blockers.append("rule_confirmation_gate 未 cleared，用户决策未完成")
            if gate_json is None:
                blockers.append("rule_confirmation_gate.json 不存在")
    elif not is_slot_facts_resolved(slot_facts, contract):
        blockers.append("FH-SLOT-FACTS-UNRESOLVED: slot_facts 存在 unresolved 必需槽位，无法自动推进")

    return blockers


@dataclass
class PredicateResult:
    """公共谓词结果。"""

    valid: bool
    errors: list[str] = field(default_factory=list)
