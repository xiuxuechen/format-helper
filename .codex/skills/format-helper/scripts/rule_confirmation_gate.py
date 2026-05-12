"""规则确认 Gate 摘要与槽位决策回填（CODE-016）。"""

from __future__ import annotations

import json
import os
import re
from copy import deepcopy
from pathlib import Path
from typing import Any

from scripts.validation.manual_review_repair import (
    build_manual_review_items,
    count_manual_review_items,
    write_manual_review_items,
)
from scripts.validation.skill_result_io import compute_file_sha256, resolve_run_relative_path


RULE_CONFIRMATION_GATE_PATH = "logs/rule_confirmation_gate.json"


def write_json_atomic(path: Path, data: dict[str, Any]) -> None:
    """原子写入 UTF-8 JSON。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temp_path, path)


def slot_proposal_from_blocker(blocker: dict[str, Any], facts_id: str) -> dict[str, Any]:
    """把 slot GateBlocker 提升为 manual-review-items 可消费的候选项。"""
    blocker_id = blocker["blocker_id"]
    slot_name = blocker["slot_name"]
    role_kind = blocker["role_kind"]
    source_ref = {
        "source_ref_id": f"SRC-{blocker_id}",
        "artifact_id": facts_id,
        "artifact_kind": "role_format_slot_facts",
        "schema_id": "role-format-slot-facts",
        "item_type": "slot_gate_blocker",
        "item_id": blocker_id,
        "attribute": slot_name,
        "value": None,
        "unit": None,
        "source_rule_ref": {"rule_id": blocker.get("rule_id")},
        "source_fact_ref": None,
        "source_role_ref": {"role_kind": role_kind},
        "locator": None,
        "evidence_refs": deepcopy(blocker.get("evidence_refs", [])),
    }
    return {
        "proposal_id": f"MRP-{blocker_id}",
        "source_issue_ids": [blocker_id],
        "source_refs": [source_ref],
        "category": "slot_confirmation",
        "problem": blocker.get("message") or f"{role_kind} 的 {slot_name} 需要人工确认",
        "impact": "阻断规则确认 Gate，未确认前不得自动激活规则包。",
        "recommended_action": "请从 suggested_options 中选择或输入确认值。",
        "risk_level": "high" if blocker.get("severity") == "error" else "medium",
        "auto_fix_policy": "require_confirmation",
        "confidence": 0.0,
        "blocking": blocker.get("severity") == "error",
        "evidence_refs": deepcopy(blocker.get("evidence_refs", [])),
        "suggested_options": deepcopy(blocker.get("suggested_options", [])),
    }


def normalize_slot_manual_review_proposal(proposal: dict[str, Any], facts_id: str) -> dict[str, Any]:
    """把 strategist 的 ManualReviewItemDraft 候选补齐为权威清单输入。"""
    if proposal.get("source_refs"):
        normalized = deepcopy(proposal)
    else:
        source = proposal.get("source", {})
        blocker_id = source.get("item_id") or proposal.get("proposal_id", "slot-gate-blocker")
        role_kind = source.get("role_kind")
        slot_name = source.get("slot_name")
        normalized = {
            "proposal_id": proposal.get("proposal_id") or f"MRP-{blocker_id}",
            "source_issue_ids": [blocker_id],
            "source_refs": [
                {
                    "source_ref_id": f"SRC-{blocker_id}",
                    "artifact_id": facts_id,
                    "artifact_kind": "role_format_slot_facts",
                    "schema_id": "role-format-slot-facts",
                    "item_type": source.get("item_type", "slot_gate_blocker"),
                    "item_id": blocker_id,
                    "attribute": slot_name,
                    "value": None,
                    "unit": None,
                    "source_rule_ref": None,
                    "source_fact_ref": None,
                    "source_role_ref": {"role_kind": role_kind},
                    "locator": None,
                    "evidence_refs": deepcopy(proposal.get("evidence_refs", [])),
                }
            ],
        }
    normalized.setdefault("source_issue_ids", [proposal.get("proposal_id", "slot-gate-blocker")])
    normalized.setdefault("category", "slot_confirmation")
    normalized.setdefault("problem", proposal.get("reason") or "槽位需要人工确认")
    normalized.setdefault("impact", "阻断规则确认 Gate，未确认前不得自动激活规则包。")
    normalized.setdefault("recommended_action", "请从 suggested_options 中选择或输入确认值。")
    normalized.setdefault("risk_level", "high" if proposal.get("blocking", True) else "medium")
    normalized.setdefault("auto_fix_policy", "require_confirmation")
    normalized.setdefault("confidence", 0.0)
    normalized.setdefault("blocking", bool(proposal.get("blocking", True)))
    normalized.setdefault("evidence_refs", deepcopy(proposal.get("evidence_refs", [])))
    normalized.setdefault("suggested_options", deepcopy(proposal.get("suggested_options", [])))
    return normalized


def build_manual_review_from_slot_facts(
    slot_facts: dict[str, Any],
    *,
    generated_at: str = "2026-05-09T00:00:00+08:00",
) -> dict[str, Any]:
    """由 format-helper 主控把 slot gate blockers 提升为权威人工确认清单。"""
    facts_id = slot_facts.get("facts_id") or "ART-SLOT-FACTS"
    source_proposals = slot_facts.get("manual_review_proposals") or [
        slot_proposal_from_blocker(blocker, facts_id) for blocker in slot_facts.get("gate_blockers", [])
    ]
    proposals = [normalize_slot_manual_review_proposal(proposal, facts_id) for proposal in source_proposals]
    return build_manual_review_items(run_id=slot_facts.get("run_id", "unknown"), proposals=proposals, generated_at=generated_at)


def manual_item_refs(manual_review_items: dict[str, Any]) -> list[dict[str, Any]]:
    """生成 rule_confirmation_gate 到 manual_review_items 的索引引用。"""
    refs = []
    for item in manual_review_items.get("items", []):
        refs.append(
            {
                "review_id": item.get("review_id"),
                "proposal_id": item.get("proposal_id"),
                "items_path": "plans/manual_review_items.json",
                "decision_status": item.get("decision", {}).get("status"),
                "allows_continue": item.get("decision", {}).get("allows_continue"),
            }
        )
    return refs


def gate_status_from_manual_items(slot_facts: dict[str, Any], manual_review_items: dict[str, Any]) -> str:
    """根据权威人工确认清单推导 Gate 摘要状态。"""
    if slot_facts.get("gate_status") == "passed":
        return "cleared"
    items = manual_review_items.get("items", [])
    if not items:
        return "blocked"
    counts = count_manual_review_items(items)
    if counts["pending_count"] > 0:
        return "pending"
    for item in items:
        decision = item.get("decision", {})
        if item.get("blocking") is True and decision.get("allows_continue") is not True:
            return "blocked"
    return "cleared"


def unresolved_slot_entries(slot_facts: dict[str, Any]) -> list[dict[str, Any]]:
    """抽取未解析槽位，供 Gate 摘要展示。"""
    entries = []
    for role in slot_facts.get("roles", []):
        for slot_name, summary in role.get("slot_summary", {}).items():
            if summary.get("status") != "unresolved":
                continue
            entries.append(
                {
                    "role_kind": role.get("role_kind"),
                    "role_id": role.get("role_id"),
                    "slot_name": slot_name,
                    "prompt": summary.get("confirmation_prompt"),
                    "source_fact_refs": deepcopy(summary.get("source_fact_refs", [])),
                    "triggered_rule_ids": deepcopy(summary.get("triggered_rule_ids", [])),
                }
            )
    return entries


def conflict_entries(slot_facts: dict[str, Any]) -> list[dict[str, Any]]:
    """抽取冲突槽位，供 Gate 摘要展示。"""
    entries = []
    for role in slot_facts.get("roles", []):
        for slot_name, summary in role.get("slot_summary", {}).items():
            if summary.get("status") not in {"conflict", "resolved_with_conflicts"}:
                continue
            entries.append(
                {
                    "role_kind": role.get("role_kind"),
                    "role_id": role.get("role_id"),
                    "slot_name": slot_name,
                    "status": summary.get("status"),
                    "mode_value": summary.get("mode_value"),
                    "mode_coverage": summary.get("mode_coverage"),
                    "value_histogram": deepcopy(summary.get("value_histogram", [])),
                    "conflicts": deepcopy(summary.get("conflicts", [])),
                    "prompt": summary.get("confirmation_prompt"),
                    "source_fact_refs": deepcopy(summary.get("source_fact_refs", [])),
                }
            )
    return entries


def build_rule_confirmation_gate(
    slot_facts: dict[str, Any],
    manual_review_items: dict[str, Any],
    *,
    generated_at: str = "2026-05-09T00:00:00+08:00",
) -> dict[str, Any]:
    """构造 rule_confirmation_gate.json 摘要，不承载用户决策权威。"""
    return {
        "schema_id": "rule-confirmation-gate",
        "schema_version": "1.0.0",
        "contract_version": "v4",
        "run_id": slot_facts.get("run_id", "unknown"),
        "gate_id": f"RCG-{slot_facts.get('facts_id', slot_facts.get('run_id', 'unknown'))}",
        "status": gate_status_from_manual_items(slot_facts, manual_review_items),
        "slot_facts_ref": {
            "path": "semantic/role_format_slot_facts.json",
            "path_kind": "run_relative",
            "facts_id": slot_facts.get("facts_id"),
            "schema_id": "role-format-slot-facts",
        },
        "manual_review_items_path": "plans/manual_review_items.json",
        "unresolved_slots": unresolved_slot_entries(slot_facts),
        "conflicts": conflict_entries(slot_facts),
        "manual_review_item_refs": manual_item_refs(manual_review_items),
        "evidence_refs": deepcopy(slot_facts.get("evidence_refs", [])),
        "generated_at": generated_at,
    }


def write_rule_confirmation_gate(run_dir: Path, gate: dict[str, Any]) -> dict[str, Any]:
    """写入 Gate 摘要并返回真实文件引用。"""
    errors = validate_rule_confirmation_gate(gate)
    if errors:
        raise ValueError(f"rule_confirmation_gate 未通过校验：{errors}")
    path = resolve_run_relative_path(run_dir, RULE_CONFIRMATION_GATE_PATH)
    write_json_atomic(path, gate)
    return {
        "path": RULE_CONFIRMATION_GATE_PATH,
        "sha256": compute_file_sha256(path),
        "size_bytes": path.stat().st_size,
    }


def create_rule_confirmation_gate_outputs(
    run_dir: Path,
    slot_facts: dict[str, Any] | None,
    *,
    generated_at: str = "2026-05-09T00:00:00+08:00",
) -> dict[str, Any]:
    """主控闭环：写权威 manual_review_items，再写 Gate 摘要。"""
    if not isinstance(slot_facts, dict):
        return {
            "error": "slot_facts 缺失或格式错误，无法创建 rule confirmation gate",
            "error_code": "FH-SLOT-FACTS-INVALID",
            "manual_review_items": None,
            "manual_review_items_ref": None,
            "gate": None,
            "gate_ref": None,
        }
    manual = build_manual_review_from_slot_facts(slot_facts, generated_at=generated_at)
    manual_ref = write_manual_review_items(run_dir, manual, writer="format-helper")
    gate = build_rule_confirmation_gate(slot_facts, manual, generated_at=generated_at)
    gate_ref = write_rule_confirmation_gate(run_dir, gate)
    return {"manual_review_items": manual, "manual_review_items_ref": manual_ref, "gate": gate, "gate_ref": gate_ref}


def selected_slot_value(decision: dict[str, Any]) -> Any:
    """从 manual_review decision 中提取用户确认值。"""
    selected_action = decision.get("selected_action") or {}
    parameters = selected_action.get("parameters") if isinstance(selected_action, dict) else {}
    if isinstance(parameters, dict):
        if "confirmed_value" in parameters:
            return parameters["confirmed_value"]
        if "value" in parameters:
            return parameters["value"]
    if isinstance(selected_action, dict) and "desired_value" in selected_action:
        return selected_action["desired_value"]
    return None


def next_facts_id(facts_id: str) -> str:
    """生成下一版 facts_id。"""
    match = re.match(r"^(.*-)(\d+)$", facts_id)
    if match:
        return f"{match.group(1)}{int(match.group(2)) + 1:0{len(match.group(2))}d}"
    return f"{facts_id}-r02"


def apply_rule_confirmation_decisions(
    slot_facts: dict[str, Any] | None,
    manual_review_items: dict[str, Any] | None,
    *,
    generated_at: str = "2026-05-09T00:00:00+08:00",
) -> dict[str, Any]:
    """把权威人工决策回填到新的 role_format_slot_facts revision。

    回填失败时返回包含 error 和 next_action=manual_recover 的 dict（不会抛异常）。
    """
    _error = lambda msg: {
        "error": msg,
        "error_code": "FH-SLOT-FACTS-INVALID",
        "next_action": {"kind": "manual_recover", "reason": msg},
    }
    if not isinstance(slot_facts, dict):
        return _error("slot_facts 缺失或格式错误，无法应用决策回填")
    if not isinstance(manual_review_items, dict):
        return _error("manual_review_items 缺失或格式错误，无法应用决策回填")
    try:
        updated = deepcopy(slot_facts)
        confirmed: dict[tuple[str, str], dict[str, Any]] = {}
        for item in manual_review_items.get("items", []):
            decision = item.get("decision", {})
            if decision.get("allows_continue") is not True or decision.get("status") not in {"approved", "modified"}:
                continue
            value = selected_slot_value(decision)
            if value is None:
                continue
            source_ref = next(
                (
                    ref
                    for ref in item.get("source_refs", [])
                    if ref.get("artifact_kind") == "role_format_slot_facts" and ref.get("item_type") == "slot_gate_blocker"
                ),
                None,
            )
            if not source_ref:
                continue
            role_kind = source_ref.get("source_role_ref", {}).get("role_kind")
            slot_name = source_ref.get("attribute")
            if role_kind and slot_name:
                confirmed[(role_kind, slot_name)] = {
                    "value": value,
                    "review_id": item.get("review_id"),
                    "decided_at": decision.get("decided_at"),
                }

        for role in updated.get("roles", []):
            for slot_name, summary in role.get("slot_summary", {}).items():
                key = (role.get("role_kind"), slot_name)
                if key not in confirmed:
                    continue
                decision_ref = confirmed[key]
                summary["status"] = "user_confirmed"
                summary["mode_value"] = decision_ref["value"]
                summary["mode_coverage"] = 1.0
                summary["primary_source"] = "user_confirmed"
                summary["confidence"] = 1.0
                summary["requires_confirmation"] = False
                summary["confirmation_prompt"] = ""
                source_rule_refs = list(summary.get("source_rule_refs", []))
                source_rule_refs.append(
                    {
                        "source_kind": "manual_review_item",
                        "review_id": decision_ref["review_id"],
                        "decided_at": decision_ref["decided_at"],
                    }
                )
                summary["source_rule_refs"] = source_rule_refs
        updated["facts_id"] = next_facts_id(str(slot_facts.get("facts_id", "RFSF-001")))
        updated["generated_at"] = generated_at
        unresolved_or_conflict = [
            (role.get("role_kind"), slot_name)
            for role in updated.get("roles", [])
            for slot_name, summary in role.get("slot_summary", {}).items()
            if summary.get("requires_confirmation") is True or summary.get("status") in {"unresolved", "conflict"}
        ]
        updated["gate_status"] = "passed" if not unresolved_or_conflict else "blocked"
        if updated["gate_status"] == "passed":
            updated["gate_blockers"] = []
            updated["manual_review_proposals"] = []
        return updated
    except Exception as exc:
        return _error(f"决策回填失败: {exc}")


def validate_rule_confirmation_gate(gate: dict[str, Any]) -> list[str]:
    """轻量校验 rule_confirmation_gate 关键契约。"""
    errors: list[str] = []
    required = {
        "schema_id",
        "schema_version",
        "contract_version",
        "run_id",
        "gate_id",
        "status",
        "slot_facts_ref",
        "manual_review_items_path",
        "unresolved_slots",
        "conflicts",
        "manual_review_item_refs",
        "evidence_refs",
        "generated_at",
    }
    for field_name in sorted(required - gate.keys()):
        errors.append(f"{field_name} is required")
    if gate.get("schema_id") != "rule-confirmation-gate":
        errors.append("schema_id must be rule-confirmation-gate")
    if gate.get("contract_version") != "v4":
        errors.append("contract_version must be v4")
    if gate.get("status") not in {"pending", "cleared", "blocked"}:
        errors.append("status is not allowed")
    if gate.get("manual_review_items_path") != "plans/manual_review_items.json":
        errors.append("manual_review_items_path must be plans/manual_review_items.json")
    if "decision" in gate:
        errors.append("rule_confirmation_gate must not carry decision")
    for ref in gate.get("manual_review_item_refs", []):
        if not ref.get("review_id"):
            errors.append("manual_review_item_refs[].review_id is required")
        if ref.get("items_path") != "plans/manual_review_items.json":
            errors.append("manual_review_item_refs[].items_path must be plans/manual_review_items.json")
    return errors


__all__ = [
    "RULE_CONFIRMATION_GATE_PATH",
    "apply_rule_confirmation_decisions",
    "build_manual_review_from_slot_facts",
    "build_rule_confirmation_gate",
    "create_rule_confirmation_gate_outputs",
    "validate_rule_confirmation_gate",
    "write_rule_confirmation_gate",
]
