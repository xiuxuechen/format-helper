"""公共校验谓词分层入口（CODE-007A）。

CODE-018 扩展：is_slot_facts_resolved / is_rule_confirmation_cleared / validate_slot_contract_compliance。
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from scripts.validation.gate_predicates import (
    can_advance,
    is_rule_confirmation_cleared,
    is_slot_facts_resolved,
    validate_slot_contract_compliance,
)
from scripts.validation.skill_result_io import compute_file_sha256
from scripts.validation.validate_schema_contract import validate_schema_contract


@dataclass
class PredicateResult:
    """公共谓词结果。"""

    valid: bool
    errors: list[str] = field(default_factory=list)


def validate_schema(artifact_path: str | Path, schema_id: str) -> PredicateResult:
    """结构层谓词：只校验 schema/example 结构，不判定流程推进。"""
    path = Path(artifact_path)
    if not path.exists():
        return PredicateResult(False, [f"artifact missing: {artifact_path}"])
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        return PredicateResult(False, [f"invalid json: {exc}"])
    validation = validate_schema_contract(data, schema_id)
    return PredicateResult(
        valid=validation.valid,
        errors=[f"{item['field']}: {item['message']}" for item in validation.errors],
    )


def is_evidence_chain_intact(manifest: dict[str, Any]) -> bool:
    """证据层谓词：只判断 evidence manifest 是否断链。"""
    if not isinstance(manifest, dict):
        return False
    if manifest.get("status") == "broken":
        return False
    if manifest.get("blockers"):
        return False
    return True


def is_final_acceptance_immutable(path: str | Path, expected_sha256: str) -> bool:
    """终态不可变谓词：只校验 final_acceptance 文件 hash 是否保持不变。"""
    final_path = Path(path)
    if not final_path.exists() or not final_path.is_file():
        return False
    return compute_file_sha256(final_path) == expected_sha256


def is_reporting_result_post_only(final_acceptance_path: str | Path, original_sha256: str | None = None) -> bool:
    """报告后置谓词：reporting 阶段不得改写 final_acceptance。"""
    final_path = Path(final_acceptance_path)
    if not final_path.exists() or not final_path.is_file():
        return False
    try:
        final_acceptance = json.loads(final_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return False
    forbidden_reporting_fields = {
        "report_refs",
        "report_path",
        "report_paths",
        "report_artifacts",
        "reporting_result_path",
        "reporting_manifest_ref",
    }
    if any(field in final_acceptance for field in forbidden_reporting_fields):
        return False
    if original_sha256 is None:
        return True
    return is_final_acceptance_immutable(final_path, original_sha256)


__all__ = [
    "PredicateResult",
    "validate_schema",
    "can_advance",
    "is_evidence_chain_intact",
    "is_final_acceptance_immutable",
    "is_reporting_result_post_only",
    "is_slot_facts_resolved",
    "is_rule_confirmation_cleared",
    "validate_slot_contract_compliance",
]
