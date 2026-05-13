#!/usr/bin/env python3
"""生成 role_format_slot_facts.json 的最小可落地实现。"""

from __future__ import annotations

import argparse
import hashlib
import json
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from scripts.utils.simple_yaml import load_yaml


TZ = timezone(timedelta(hours=8))
CONTRACT_RELATIVE_PATH = "docs/v4/schemas/role_slot_contract.yaml"
ROLE_MAP_RELATIVE_PATH = "semantic/semantic_role_map.before.json"
RESOLVED_STATUS = {"resolved", "resolved_with_conflicts", "not_applicable", "user_confirmed"}
ROLE_MAP_CONFIDENCE_THRESHOLD = 0.85
SUPPORTED_COMMON_CONDITIONS = {"required_slot_confidence_eq_0"}
STRUCTURE_ROLE_BY_FACT_KIND = {
    "table_cell_paragraph": "table-content",
    "page_setup": "section-page-setup",
    "header_paragraph": "header-footer",
    "footer_paragraph": "header-footer",
}
SUPPORTED_ROLE_FACT_KINDS = {
    "table-content": {"table_cell_paragraph"},
    "section-page-setup": {"page_setup"},
    "header-footer": {"header_paragraph", "footer_paragraph"},
}


def canonical_json(data: Any) -> str:
    """返回稳定 JSON 字符串。"""
    return json.dumps(data, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def sha256_file(path: Path) -> str:
    """计算文件 sha256。"""
    return hashlib.sha256(path.read_bytes()).hexdigest()


def unwrap_slot(value: Any) -> tuple[Any, str, float]:
    """兼容读取裸值和 {value, source, confidence}。"""
    if isinstance(value, dict) and "value" in value:
        return value.get("value"), str(value.get("source", "legacy")), float(value.get("confidence", 0.5))
    if value is None:
        return None, "unresolved", 0.0
    return value, "legacy", 0.5


def fact_id_for(item: dict[str, Any], index: int) -> tuple[str, str | None]:
    """返回 v4 fact_id，并保留 legacy element_id。"""
    fact_id = item.get("fact_id")
    legacy = item.get("element_id")
    if fact_id:
        return str(fact_id), str(legacy) if legacy else None
    if legacy:
        return f"legacy-{legacy}", str(legacy)
    return f"generated-fact-{index:05d}", None


def snapshot_items(snapshot: dict[str, Any]) -> list[dict[str, Any]]:
    """抽取本期支持的 snapshot 事实。"""
    items: list[dict[str, Any]] = []
    for index, paragraph in enumerate(snapshot.get("paragraphs", []), start=1):
        fact_id, legacy = fact_id_for(paragraph, index)
        item = dict(paragraph)
        item["fact_id"] = fact_id
        item["element_id_legacy"] = legacy
        item["fact_kind"] = item.get("fact_kind") or "paragraph"
        items.append(item)
    item_index = len(items)
    for table_index, table in enumerate(snapshot.get("tables", []), start=1):
        for cell in table.get("cells", []):
            for paragraph in cell.get("paragraphs", []):
                item_index += 1
                item = dict(paragraph)
                fact_id = item.get("fact_id") or (
                    f"table-{table_index:04d}-r{cell.get('row_index', 0):03d}-"
                    f"c{cell.get('column_index', 0):03d}-p{item_index:05d}"
                )
                item["fact_id"] = str(fact_id)
                item["element_id_legacy"] = item.get("element_id")
                item["fact_kind"] = "table_cell_paragraph"
                item["table_index"] = table_index
                item["cell_id"] = cell.get("cell_id")
                item["row_index"] = cell.get("row_index")
                item["column_index"] = cell.get("column_index")
                item["cell_format_summary"] = cell.get("format_summary") if isinstance(cell.get("format_summary"), dict) else {}
                if cell.get("vertical_alignment") is not None:
                    item["cell_format_summary"]["vertical_alignment"] = cell.get("vertical_alignment")
                items.append(item)
    for section_index, section in enumerate(snapshot.get("sections", []), start=1):
        page_setup = section.get("page_setup") if isinstance(section.get("page_setup"), dict) else section
        item = {
            "fact_id": section.get("fact_id") or f"section-{section_index:04d}-page-setup",
            "element_id_legacy": None,
            "fact_kind": "page_setup",
            "locator": {"section_index": section.get("section_index", section_index)},
            "text_preview": "",
            "resolved_page_setup": {
                "page_orientation": page_setup.get("page_orientation") or page_setup.get("orientation"),
                "page_width_twips": page_setup.get("page_width_twips"),
                "page_height_twips": page_setup.get("page_height_twips"),
                "margin_top_cm": page_setup.get("margin_top_cm"),
                "margin_bottom_cm": page_setup.get("margin_bottom_cm"),
                "margin_left_cm": page_setup.get("margin_left_cm"),
                "margin_right_cm": page_setup.get("margin_right_cm"),
                "header_distance_cm": page_setup.get("header_distance_cm"),
                "footer_distance_cm": page_setup.get("footer_distance_cm"),
                "page_number_format": page_setup.get("page_number_format") or page_setup.get("pg_num_type") or "none",
            },
        }
        items.append(item)
    return items


def role_map_roles(semantic_role_map: dict[str, Any] | None) -> dict[str, dict[str, Any]]:
    """按 slot role kind 建立 role-map 角色索引。"""
    if not isinstance(semantic_role_map, dict):
        return {}
    index: dict[str, dict[str, Any]] = {}
    for role in semantic_role_map.get("roles", []):
        if not isinstance(role, dict):
            continue
        role_kind = role.get("slot_role_kind")
        role_id = role.get("role_id")
        if role_kind and role_id:
            index[str(role_kind)] = role
    return index


def role_map_mappings(semantic_role_map: dict[str, Any] | None) -> list[dict[str, Any]]:
    """读取 role-map mappings。"""
    if not isinstance(semantic_role_map, dict):
        return []
    mappings = semantic_role_map.get("mappings", [])
    return [mapping for mapping in mappings if isinstance(mapping, dict)]


def reason(code: str, message: str, source: str, evidence_ref: str | None = None) -> dict[str, Any]:
    """构造稳定 resolver reason。"""
    return {
        "reason_code": code,
        "message": message,
        "source": source,
        "evidence_ref": evidence_ref,
    }


def choose_samples(
    items: list[dict[str, Any]],
    role_kind: str,
    role_id: str,
    mappings: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], float]:
    """选择角色样本；仅接受 role-map 可信映射。"""
    items_by_fact_id = {str(item.get("fact_id")): item for item in items if item.get("fact_id")}
    target_roles_by_fact: dict[str, set[str]] = defaultdict(set)
    for mapping in mappings:
        fact_id = str(mapping.get("fact_id") or mapping.get("source_fact_id") or "")
        target_role_id = mapping.get("target_role_id")
        if fact_id and isinstance(target_role_id, str) and target_role_id:
            target_roles_by_fact[fact_id].add(target_role_id)
    samples: list[dict[str, Any]] = []
    reasons: list[dict[str, Any]] = []
    confidences: list[float] = []
    for mapping in mappings:
        target_role_id = mapping.get("target_role_id")
        if target_role_id != role_id:
            continue
        confidence = float(mapping.get("confidence", 0.0) or 0.0)
        fact_id = str(mapping.get("fact_id") or mapping.get("source_fact_id") or "")
        if len(target_roles_by_fact.get(fact_id, set())) > 1:
            reasons.append(reason("MULTI_ROLE_CONFLICT", "同一事实被多个 role-map 角色竞争。", "role-map", fact_id or None))
            continue
        if mapping.get("mapping_kind") == "unresolved":
            reasons.append(reason("UNRESOLVED_MAPPING", "role-map 标记为未解析。", "role-map", fact_id or None))
            continue
        if confidence < ROLE_MAP_CONFIDENCE_THRESHOLD:
            reasons.append(reason("LOW_CONFIDENCE", "role-map 置信度低于阈值。", "role-map", fact_id or None))
            continue
        item = items_by_fact_id.get(fact_id)
        if item is None:
            reasons.append(reason("UNRESOLVED_MAPPING", "role-map 指向的事实不存在。", "role-map", fact_id or None))
            continue
        structure_role = STRUCTURE_ROLE_BY_FACT_KIND.get(str(item.get("fact_kind") or ""))
        if structure_role and structure_role != role_kind:
            reasons.append(reason("STRUCTURE_RULE_BLOCKED", "role-map 与结构事实角色冲突。", "structure-rule", fact_id))
            continue
        supported_fact_kinds = SUPPORTED_ROLE_FACT_KINDS.get(role_kind)
        if supported_fact_kinds is not None and item.get("fact_kind") not in supported_fact_kinds:
            reasons.append(reason("STRUCTURE_RULE_BLOCKED", "事实类型不属于该槽位角色允许范围。", "structure-rule", fact_id))
            continue
        samples.append(item)
        confidences.append(confidence)
    if samples:
        reasons.append(reason("ROLEMAP_CONSISTENT", "采用 role-map 可信映射。", "role-map", samples[0].get("fact_id")))
        return samples[:10], reasons, min(confidences) if confidences else 1.0

    reasons.append(reason("UNRESOLVED_MAPPING", "无可信 role-map 映射。", "role-map", None))
    return [], reasons, 0.0


def read_slot(item: dict[str, Any], slot_name: str) -> dict[str, Any]:
    """从事实中读取槽位。"""
    paragraph_format = item.get("resolved_paragraph_format") or item.get("paragraph_format") or {}
    run_format = item.get("resolved_run_format") or item.get("run_format") or {}
    page_setup = item.get("resolved_page_setup") or item.get("page_setup") or {}
    cell_format = item.get("cell_format_summary") if isinstance(item.get("cell_format_summary"), dict) else {}
    raw = paragraph_format.get(slot_name)
    if raw is None and slot_name in run_format:
        raw = run_format.get(slot_name)
    if raw is None and slot_name in page_setup:
        raw = page_setup.get(slot_name)
    if raw is None and slot_name in cell_format:
        raw = cell_format.get(slot_name)
    value, source, confidence = unwrap_slot(raw)
    return {"value": value, "source": source, "confidence": confidence, "evidence_refs": [], "notes": ""}


def build_sample(item: dict[str, Any], slots: list[str]) -> dict[str, Any]:
    """生成样本级槽位抽取。"""
    return {
        "fact_id": item["fact_id"],
        "locator": item.get("locator"),
        "element_id_legacy": item.get("element_id_legacy"),
        "fact_kind": item.get("fact_kind") or "paragraph",
        "text_preview": str(item.get("text_preview") or "")[:60],
        "extracted_slots": {slot: read_slot(item, slot) for slot in slots},
    }


def histogram_key(value: Any) -> str:
    """把槽位值转成可计数 key。"""
    return canonical_json(value)


def validate_common_validation_rules(contract: dict[str, Any]) -> None:
    """校验公共规则条件均已登记，避免生成端默默忽略。"""
    for index, rule in enumerate(contract.get("common_validation_rules", [])):
        if not isinstance(rule, dict):
            raise ValueError(f"common_validation_rules[{index}] 必须是对象")
        condition = rule.get("condition")
        if condition not in SUPPORTED_COMMON_CONDITIONS:
            raise ValueError(f"未知 common_validation_rules condition: {condition}")
        if rule.get("severity") not in {"error", "warning"}:
            raise ValueError(f"common_validation_rules[{index}].severity 必须为 error 或 warning")
        if rule.get("applies_to") != "all_roles":
            raise ValueError(f"common_validation_rules[{index}].applies_to 必须为 all_roles")
        if rule.get("applies_to_slots") != "required_slots":
            raise ValueError(f"common_validation_rules[{index}].applies_to_slots 必须为 required_slots")
        if not isinstance(rule.get("blocks_confirmation_gate"), bool):
            raise ValueError(f"common_validation_rules[{index}].blocks_confirmation_gate 必须为 boolean")


def apply_common_validation_rules(
    role_kind: str,
    slot_name: str,
    summary: dict[str, Any],
    required: bool,
    common_rules: list[dict[str, Any]],
) -> dict[str, Any]:
    """执行公共规则；命中阻断规则时转成人工确认。"""
    if not required:
        return summary
    updated = dict(summary)
    for rule in common_rules:
        if rule.get("applies_to") != "all_roles" or rule.get("applies_to_slots") != "required_slots":
            continue
        if rule.get("condition") != "required_slot_confidence_eq_0":
            continue
        if updated.get("confidence") != 0:
            continue
        rule_id = str(rule.get("rule_id") or "RSC-COMMON-ZERO-CONFIDENCE")
        triggered = list(updated.get("triggered_rule_ids") or [])
        if rule_id not in triggered:
            triggered.append(rule_id)
        updated["triggered_rule_ids"] = triggered
        source_rules = list(updated.get("source_rule_refs") or [])
        if rule_id not in source_rules:
            source_rules.append(rule_id)
        updated["source_rule_refs"] = source_rules
        if rule.get("blocks_confirmation_gate") is True or rule.get("severity") == "error":
            updated["status"] = "unresolved"
            updated["requires_confirmation"] = True
            updated["confirmation_prompt"] = str(rule.get("message") or f"请确认 {role_kind} 的 {slot_name}")
    return updated


def summarize_slot(
    role_kind: str,
    slot_name: str,
    slot_type: dict[str, Any],
    samples: list[dict[str, Any]],
    required: bool,
    confirmed_slots: dict[str, dict[str, Any]] | None = None,
    not_applicable_slots: set[str] | None = None,
) -> dict[str, Any]:
    """聚合单个槽位状态。"""
    confirmed_slots = confirmed_slots or {}
    not_applicable_slots = not_applicable_slots or set()
    status_key = f"{role_kind}.{slot_name}"
    if status_key in confirmed_slots:
        confirmed = confirmed_slots[status_key]
        return {
            "slot_name": slot_name,
            "unit": slot_type.get("unit"),
            "status": "user_confirmed",
            "mode_value": confirmed.get("value"),
            "mode_coverage": 1.0,
            "total_samples": len(samples),
            "value_histogram": [],
            "conflicts": [],
            "primary_source": "user_confirmed",
            "source_fact_refs": [],
            "source_rule_refs": [],
            "requires_confirmation": False,
            "confirmation_prompt": "",
            "triggered_rule_ids": [],
        }
    if status_key in not_applicable_slots:
        return {
            "slot_name": slot_name,
            "unit": slot_type.get("unit"),
            "status": "not_applicable",
            "mode_value": None,
            "mode_coverage": 1.0,
            "total_samples": len(samples),
            "value_histogram": [],
            "conflicts": [],
            "primary_source": "user_confirmed",
            "source_fact_refs": [],
            "source_rule_refs": [],
            "requires_confirmation": False,
            "confirmation_prompt": "",
            "triggered_rule_ids": [],
        }
    values: list[tuple[Any, str, str, float]] = []
    for sample in samples:
        slot = sample["extracted_slots"][slot_name]
        if slot["value"] is not None:
            values.append((slot["value"], slot["source"], sample["fact_id"], float(slot.get("confidence", 0.0) or 0.0)))
    total = len(samples)
    if not values:
        status = "unresolved"
        mode_value = None
        mode_coverage = 0.0
        mode_confidence = 0.0
        primary_source = "unresolved"
        histogram: list[dict[str, Any]] = []
        source_fact_refs: list[str] = []
    else:
        counter = Counter(histogram_key(value) for value, _, _, _ in values)
        mode_key, count = counter.most_common(1)[0]
        mode_value = json.loads(mode_key)
        mode_coverage = round(count / len(values), 3)
        mode_confidence = min(confidence for value, _, _, confidence in values if histogram_key(value) == mode_key)
        source_counter = Counter(source for value, source, _, _ in values if histogram_key(value) == mode_key)
        primary_source = source_counter.most_common(1)[0][0]
        if mode_coverage >= 0.9 and primary_source != "unresolved":
            status = "resolved"
        elif mode_coverage >= 0.6:
            status = "resolved_with_conflicts"
        else:
            status = "conflict"
        refs_by_key: dict[str, list[str]] = defaultdict(list)
        for value, _, fact_id, _ in values:
            key = histogram_key(value)
            if len(refs_by_key[key]) < 10:
                refs_by_key[key].append(fact_id)
        histogram = [
            {
                "value": json.loads(key),
                "count": value_count,
                "ratio": round(value_count / len(values), 3),
                "sample_fact_ids": refs_by_key[key],
            }
            for key, value_count in counter.most_common(20)
        ]
        source_fact_refs = refs_by_key.get(mode_key, [])[:10]
    rule_id = f"RSC-{role_kind.upper().replace('-', '-')}-UNRESOLVED" if status == "unresolved" else f"RSC-{role_kind.upper().replace('-', '-')}-CONFLICT"
    requires_confirmation = required and status in {"unresolved", "conflict"}
    return {
        "slot_name": slot_name,
        "unit": slot_type.get("unit"),
        "status": status,
        "mode_value": mode_value,
        "mode_coverage": mode_coverage,
        "confidence": round(mode_confidence, 3),
        "total_samples": total,
        "value_histogram": histogram,
        "conflicts": [bucket for bucket in histogram[1:]],
        "primary_source": primary_source,
        "source_fact_refs": source_fact_refs,
        "source_rule_refs": [],
        "requires_confirmation": requires_confirmation,
        "confirmation_prompt": f"请确认 {role_kind} 的 {slot_name}" if requires_confirmation else "",
        "triggered_rule_ids": [rule_id] if requires_confirmation else [],
    }


def blocker_for(role_kind: str, slot_name: str, summary: dict[str, Any]) -> dict[str, Any]:
    """生成 GateBlocker。"""
    rule_id = summary["triggered_rule_ids"][0] if summary.get("triggered_rule_ids") else f"RSC-{role_kind.upper()}-{slot_name.upper()}-BLOCKED"
    return {
        "blocker_id": f"BLK-{role_kind}-{slot_name}-{rule_id}",
        "role_kind": role_kind,
        "slot_name": slot_name,
        "rule_id": rule_id,
        "severity": "error",
        "message": f"{role_kind} 的 {slot_name} 未能确定，必须人工确认。",
        "suggested_options": [
            {"label": "使用众数", "value": summary.get("mode_value")},
            {"label": "人工输入", "value": None, "requires_input": True},
        ],
    }


def resolver_blocker_for(role_kind: str, reason_item: dict[str, Any]) -> dict[str, Any]:
    """把 resolver 冲突转换为稳定 GateBlocker。"""
    reason_code = str(reason_item.get("reason_code") or "UNRESOLVED_MAPPING")
    severity = "error"
    code_suffix = "CONFLICT" if reason_code in {"MULTI_ROLE_CONFLICT", "STRUCTURE_RULE_BLOCKED"} else "UNRESOLVED"
    rule_id = f"RSC-{role_kind.upper()}-{reason_code}"
    return {
        "blocker_id": f"BLK-{role_kind}-resolver-{reason_code}",
        "role_kind": role_kind,
        "slot_name": "*",
        "rule_id": rule_id,
        "severity": severity,
        "message": reason_item.get("message") or f"{role_kind} 的角色映射存在 {code_suffix.lower()}，必须人工确认。",
        "suggested_options": [
            {"label": "人工确认角色映射", "value": None, "requires_input": True},
        ],
        "error_code": f"FH-SLOT-FACTS-{code_suffix}",
    }


def proposal_for(blocker: dict[str, Any]) -> dict[str, Any]:
    """把 gate blocker 派生为 ManualReviewItemDraft 候选。"""
    return {
        "proposal_id": f"MRP-{blocker['blocker_id']}",
        "source": {
            "artifact_kind": "role_format_slot_facts",
            "item_type": "slot_gate_blocker",
            "item_id": blocker["blocker_id"],
            "role_kind": blocker["role_kind"],
            "slot_name": blocker["slot_name"],
        },
        "blocking": True,
        "reason": blocker["message"],
        "suggested_options": blocker["suggested_options"],
        "evidence_refs": [],
    }


def build_slot_facts(
    snapshot: dict[str, Any],
    contract: dict[str, Any],
    *,
    run_id: str,
    source_snapshot_path: str,
    source_snapshot_sha256: str,
    source_snapshot_artifact_id: str,
    contract_sha256: str,
    semantic_role_map: dict[str, Any] | None = None,
    semantic_role_map_sha256: str | None = None,
    semantic_role_map_path: str = ROLE_MAP_RELATIVE_PATH,
    generated_at: str | None = None,
    confirmed_slots: dict[str, dict[str, Any]] | None = None,
    not_applicable_slots: set[str] | None = None,
) -> dict[str, Any]:
    """构建 role_format_slot_facts 对象。"""
    if not isinstance(semantic_role_map, dict):
        raise ValueError("semantic_role_map.before.json 缺失，无法生成可校验的 target_role_ref")
    if not semantic_role_map_sha256:
        raise ValueError("semantic_role_map sha256 缺失，无法生成可校验的 target_role_ref")
    validate_common_validation_rules(contract)
    items = snapshot_items(snapshot)
    slot_registry = contract["slot_type_registry"]
    common_rules = contract.get("common_validation_rules", [])
    role_index = role_map_roles(semantic_role_map)
    mappings = role_map_mappings(semantic_role_map)
    roles: list[dict[str, Any]] = []
    gate_blockers: list[dict[str, Any]] = []
    for role_kind, role_contract in contract["role_slot_contracts"].items():
        slots = list(dict.fromkeys(role_contract["required_slots"] + role_contract.get("optional_slots", [])))
        role_map_role = role_index.get(role_kind)
        if not role_map_role:
            raise ValueError(f"semantic_role_map.before.json 缺少 slot_role_kind={role_kind} 的角色定义")
        role_id = str(role_map_role.get("role_id"))
        raw_samples, resolver_reasons, resolver_confidence = choose_samples(items, role_kind, role_id, mappings)
        samples = [build_sample(item, slots) for item in raw_samples]
        summaries: dict[str, Any] = {}
        confirmation_reasons: list[str] = [
            reason_item["reason_code"]
            for reason_item in resolver_reasons
            if reason_item.get("reason_code") in {"MULTI_ROLE_CONFLICT", "STRUCTURE_RULE_BLOCKED"}
        ]
        for reason_item in resolver_reasons:
            if reason_item.get("reason_code") in {"MULTI_ROLE_CONFLICT", "STRUCTURE_RULE_BLOCKED"}:
                gate_blockers.append(resolver_blocker_for(role_kind, reason_item))
        for slot in slots:
            required = slot in role_contract["required_slots"]
            summary = summarize_slot(
                role_kind,
                slot,
                slot_registry[slot],
                samples,
                required,
                confirmed_slots=confirmed_slots,
                not_applicable_slots=not_applicable_slots,
            )
            summary = apply_common_validation_rules(role_kind, slot, summary, required, common_rules)
            summaries[slot] = summary
            if summary["requires_confirmation"]:
                confirmation_reasons.extend(summary["triggered_rule_ids"])
                gate_blockers.append(blocker_for(role_kind, slot, summary))
        requires_confirmation = bool(confirmation_reasons)
        resolved_count = sum(1 for slot in role_contract["required_slots"] if summaries[slot]["status"] in RESOLVED_STATUS)
        required_total = max(len(role_contract["required_slots"]), 1)
        target_ref = {
            "path": semantic_role_map_path,
            "path_kind": "run_relative",
            "role_id": role_id,
            "sha256": semantic_role_map_sha256,
        }
        roles.append(
            {
                "role_kind": role_kind,
                "role_id": role_id,
                "target_role_ref": target_ref,
                "semantic_role_kind_ref": role_contract.get("semantic_role_kind_ref"),
                "category": role_contract["category"],
                "sample_count": len(samples),
                "samples": samples,
                "slot_summary": summaries,
                "role_confidence": round(min(resolved_count / required_total, resolver_confidence), 3),
                "requires_confirmation": requires_confirmation,
                "confirmation_reasons": sorted(set(confirmation_reasons)),
                "reasons": resolver_reasons,
            }
        )
    gate_status = "blocked" if gate_blockers else "passed"
    manual_review_proposals = [proposal_for(blocker) for blocker in gate_blockers]
    return {
        "schema_id": "role-format-slot-facts",
        "schema_version": "1.0.0",
        "contract_version": "v4",
        "run_id": run_id,
        "facts_id": f"RFSF-{run_id}-001",
        "source_snapshot_path": source_snapshot_path,
        "source_snapshot_sha256": source_snapshot_sha256,
        "source_snapshot_artifact_id": source_snapshot_artifact_id,
        "contract_ref": {
            "contract_path": CONTRACT_RELATIVE_PATH,
            "contract_sha256": contract_sha256,
            "contract_version": contract.get("schema_version", "1.0.0"),
            "path_kind": "workspace_relative",
        },
        "roles": roles,
        "gate_status": gate_status,
        "gate_blockers": gate_blockers,
        "manual_review_proposals": manual_review_proposals,
        "source_refs": [],
        "evidence_refs": [],
        "generated_at": generated_at or datetime.now(TZ).isoformat(),
    }


def build_arg_parser() -> argparse.ArgumentParser:
    """构造命令行参数解析器。"""
    parser = argparse.ArgumentParser(description="生成 role_format_slot_facts.json")
    parser.add_argument("--snapshot", required=True, type=Path)
    parser.add_argument("--contract", default=CONTRACT_RELATIVE_PATH, type=Path)
    parser.add_argument("--role-map", required=True, type=Path)
    parser.add_argument("--run-id", required=True)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--snapshot-artifact-id", default="ART-STANDARD-SNAPSHOT")
    return parser


def main() -> int:
    """命令行入口。"""
    parser = build_arg_parser()
    args = parser.parse_args()
    snapshot = json.loads(args.snapshot.read_text(encoding="utf-8"))
    contract = load_yaml(args.contract)
    semantic_role_map = json.loads(args.role_map.read_text(encoding="utf-8")) if args.role_map else None
    facts = build_slot_facts(
        snapshot,
        contract,
        run_id=args.run_id,
        source_snapshot_path=args.snapshot.as_posix(),
        source_snapshot_sha256=sha256_file(args.snapshot),
        source_snapshot_artifact_id=args.snapshot_artifact_id,
        contract_sha256=sha256_file(args.contract),
        semantic_role_map=semantic_role_map,
        semantic_role_map_sha256=sha256_file(args.role_map) if args.role_map else None,
    )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(facts, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
