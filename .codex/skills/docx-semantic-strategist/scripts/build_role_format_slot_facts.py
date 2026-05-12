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
RESOLVED_STATUS = {"resolved", "resolved_with_conflicts", "not_applicable", "user_confirmed"}


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
    """抽取 paragraph 事实；CODE-014 MVP 先覆盖段落事实。"""
    items: list[dict[str, Any]] = []
    for index, paragraph in enumerate(snapshot.get("paragraphs", []), start=1):
        fact_id, legacy = fact_id_for(paragraph, index)
        item = dict(paragraph)
        item["fact_id"] = fact_id
        item["element_id_legacy"] = legacy
        item["fact_kind"] = item.get("fact_kind") or "paragraph"
        items.append(item)
    return items


def choose_samples(items: list[dict[str, Any]], role_kind: str) -> list[dict[str, Any]]:
    """选择角色样本；测试可显式标 role_kind，正文角色保守使用所有段落。"""
    explicit = [item for item in items if item.get("role_kind") == role_kind or item.get("semantic_role_kind") == role_kind]
    if explicit:
        return explicit[:10]
    if role_kind == "body-paragraph":
        return [item for item in items if item.get("text_preview")][:10]
    return []


def read_slot(item: dict[str, Any], slot_name: str) -> dict[str, Any]:
    """从事实中读取槽位。"""
    paragraph_format = item.get("resolved_paragraph_format") or item.get("paragraph_format") or {}
    run_format = item.get("resolved_run_format") or item.get("run_format") or {}
    raw = paragraph_format.get(slot_name)
    if raw is None and slot_name in run_format:
        raw = run_format.get(slot_name)
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
    values: list[tuple[Any, str, str]] = []
    for sample in samples:
        slot = sample["extracted_slots"][slot_name]
        if slot["value"] is not None:
            values.append((slot["value"], slot["source"], sample["fact_id"]))
    total = len(samples)
    if not values:
        status = "unresolved"
        mode_value = None
        mode_coverage = 0.0
        primary_source = "unresolved"
        histogram: list[dict[str, Any]] = []
        source_fact_refs: list[str] = []
    else:
        counter = Counter(histogram_key(value) for value, _, _ in values)
        mode_key, count = counter.most_common(1)[0]
        mode_value = json.loads(mode_key)
        mode_coverage = round(count / len(values), 3)
        source_counter = Counter(source for value, source, _ in values if histogram_key(value) == mode_key)
        primary_source = source_counter.most_common(1)[0][0]
        if mode_coverage >= 0.9 and primary_source != "unresolved":
            status = "resolved"
        elif mode_coverage >= 0.6:
            status = "resolved_with_conflicts"
        else:
            status = "conflict"
        refs_by_key: dict[str, list[str]] = defaultdict(list)
        for value, _, fact_id in values:
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
    generated_at: str | None = None,
    confirmed_slots: dict[str, dict[str, Any]] | None = None,
    not_applicable_slots: set[str] | None = None,
) -> dict[str, Any]:
    """构建 role_format_slot_facts 对象。"""
    items = snapshot_items(snapshot)
    slot_registry = contract["slot_type_registry"]
    roles: list[dict[str, Any]] = []
    gate_blockers: list[dict[str, Any]] = []
    for role_kind, role_contract in contract["role_slot_contracts"].items():
        slots = list(dict.fromkeys(role_contract["required_slots"] + role_contract.get("optional_slots", [])))
        samples = [build_sample(item, slots) for item in choose_samples(items, role_kind)]
        summaries: dict[str, Any] = {}
        confirmation_reasons: list[str] = []
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
            summaries[slot] = summary
            if summary["requires_confirmation"]:
                confirmation_reasons.extend(summary["triggered_rule_ids"])
                gate_blockers.append(blocker_for(role_kind, slot, summary))
        requires_confirmation = bool(confirmation_reasons)
        resolved_count = sum(1 for slot in role_contract["required_slots"] if summaries[slot]["status"] in RESOLVED_STATUS)
        required_total = max(len(role_contract["required_slots"]), 1)
        roles.append(
            {
                "role_kind": role_kind,
                "role_id": f"{run_id}-{role_kind}",
                "target_role_ref": {
                    "path": "semantic/semantic_rule_draft.json",
                    "role_id": f"{run_id}-{role_kind}",
                    "sha256": None,
                },
                "semantic_role_kind_ref": role_contract.get("semantic_role_kind_ref"),
                "category": role_contract["category"],
                "sample_count": len(samples),
                "samples": samples,
                "slot_summary": summaries,
                "role_confidence": round(resolved_count / required_total, 3),
                "requires_confirmation": requires_confirmation,
                "confirmation_reasons": sorted(set(confirmation_reasons)),
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


def main() -> int:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="生成 role_format_slot_facts.json")
    parser.add_argument("--snapshot", required=True, type=Path)
    parser.add_argument("--contract", default=CONTRACT_RELATIVE_PATH, type=Path)
    parser.add_argument("--run-id", required=True)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--snapshot-artifact-id", default="ART-STANDARD-SNAPSHOT")
    args = parser.parse_args()
    snapshot = json.loads(args.snapshot.read_text(encoding="utf-8"))
    contract = load_yaml(args.contract)
    facts = build_slot_facts(
        snapshot,
        contract,
        run_id=args.run_id,
        source_snapshot_path=args.snapshot.as_posix(),
        source_snapshot_sha256=sha256_file(args.snapshot),
        source_snapshot_artifact_id=args.snapshot_artifact_id,
        contract_sha256=sha256_file(args.contract),
    )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(facts, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
