"""三代 evidence manifest 生成、写入与链路校验（CODE-008）。"""

from __future__ import annotations

import json
import os
import re
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from scripts.validation.skill_result_io import (
    FH_RULE_EXPECTED_FORMAT_UNREGISTERED,
    canonical_json,
    compute_file_sha256,
    resolve_run_relative_path,
    sha256_text,
)


GENERATION_PATHS = {
    "pre_acceptance": "logs/evidence_manifest.pre_acceptance.json",
    "post_acceptance": "logs/evidence_manifest.post_acceptance.json",
    "reporting": "logs/evidence_manifest.reporting.json",
}
OFFICECLI_EVIDENCE_MANIFEST_PATH = "logs/evidence_manifest.json"

PRE_ACCEPTANCE_FORBIDDEN_ARTIFACT_KINDS = {
    "final_acceptance",
    "report",
    "reporting_result",
}

PRE_ACCEPTANCE_FORBIDDEN_SCHEMA_IDS = {
    "final-acceptance",
    "reporting-result",
}

SUPPORTED_RELATION_OBJECT_TYPES = {
    "artifact",
    "evidence",
    "issue",
    "action",
    "review",
    "acceptance",
}

ALLOWED_RELATION_TYPES = {
    "derived_from", "executes", "verifies", "renders", "refreshes", "accepts",
    "supports", "fixes", "reviews", "blocks",  # legacy legacy
}
OFFICECLI_RELATION_TYPES = {"derived_from", "executes", "verifies", "renders", "refreshes", "accepts"}
OFFICECLI_ARTIFACT_KINDS = {
    "lock", "capability", "snapshot", "plan", "request", "result", "log",
    "review", "evidence", "toc_acceptance", "final_acceptance", "docx",
    "html", "png", "raw_xml", "executable", "license",
}
OFFICECLI_ARTIFACT_REQUIRED_FIELDS = {"artifact_id", "kind", "relative_path", "sha256", "size_bytes"}
OFFICECLI_ARTIFACT_ALLOWED_FIELDS = OFFICECLI_ARTIFACT_REQUIRED_FIELDS | {"schema_id", "schema_version"}

ALLOWED_MANIFEST_STATUSES = {
    "complete",
    "complete_with_warnings",
    "broken",
}

POST_ACCEPTANCE_FORBIDDEN_ARTIFACT_KINDS = {
    "report",
    "reporting_result",
}

REPORTING_REQUIRED_ARTIFACT_KINDS = {
    "report",
    "reporting_result",
}

RULE_PACKAGING_REQUIRED_ARTIFACTS = {
    "document_snapshot": {
        "schema_id": "officecli-document-snapshot",
        "path": "snapshots/officecli-document-snapshot.standard.json",
    },
    "semantic_role_map": {
        "schema_id": "semantic-role-map",
        "path": "semantic/semantic_role_map.before.json",
    },
    "role_format_slot_facts": {
        "schema_id": "role-format-slot-facts",
        "path": "semantic/role_format_slot_facts.json",
    },
    "rule_confirmation_gate": {
        "schema_id": "rule-confirmation-gate",
        "path": "logs/rule_confirmation_gate.json",
    },
}

EVIDENCE_REDUNDANT_FIELDS = [
    "path",
    "path_kind",
    "sha256",
    "size_bytes",
    "schema_id",
    "schema_version",
    "required",
    "producer_result_id",
]

TOP_LEVEL_REQUIRED_FIELDS = [
    "schema_id",
    "schema_version",
    "contract_version",
    "run_id",
    "status",
    "manifest_generation",
    "manifest_id",
    "generated_at",
    "manifest_sha256",
    "artifacts",
    "evidence",
    "relations",
    "warnings",
    "blockers",
]

EVIDENCE_REQUIRED_FIELDS = [
    "evidence_id",
    "artifact_id",
    "kind",
    "path",
    "path_kind",
    "sha256",
    "size_bytes",
    "schema_id",
    "schema_version",
    "required",
    "producer_result_id",
    "depends_on",
    "summary",
]


@dataclass
class ManifestValidationResult:
    """evidence manifest 校验结果。"""

    valid: bool
    status: str
    errors: list[str] = field(default_factory=list)


def _blocker(code: str, message: str, evidence_refs: list[str] | None = None) -> dict[str, Any]:
    """构造证据链阻断项。"""
    return {
        "code": code,
        "category": "evidence",
        "message": message,
        "impact": "证据链断裂，当前 Gate 不得继续推进。",
        "blocking": True,
        "user_action": None,
        "recovery": None,
        "evidence_refs": evidence_refs or [],
    }


def validate_rule_packaging_expected_artifacts(manifest: dict[str, Any]) -> ManifestValidationResult:
    """校验 rule_packaging 期望格式产物已完整登记到 evidence manifest。"""
    result = validate_evidence_manifest(manifest)
    errors = list(result.errors)
    artifacts = manifest.get("artifacts") if isinstance(manifest, dict) else None
    if not isinstance(artifacts, list):
        return ManifestValidationResult(False, "broken", errors or ["artifacts must be array"])

    by_kind = {
        artifact.get("kind"): artifact
        for artifact in artifacts
        if isinstance(artifact, dict) and artifact.get("kind")
    }
    for kind, expected in RULE_PACKAGING_REQUIRED_ARTIFACTS.items():
        artifact = by_kind.get(kind)
        if not isinstance(artifact, dict):
            errors.append(f"{FH_RULE_EXPECTED_FORMAT_UNREGISTERED}: missing artifact kind {kind}")
            continue
        if artifact.get("schema_id") != expected["schema_id"]:
            errors.append(
                f"{FH_RULE_EXPECTED_FORMAT_UNREGISTERED}: artifact kind {kind} schema_id must be {expected['schema_id']}"
            )
        if artifact.get("path") != expected["path"]:
            errors.append(f"{FH_RULE_EXPECTED_FORMAT_UNREGISTERED}: artifact kind {kind} path must be {expected['path']}")
        if artifact.get("path_kind") != "run_relative":
            errors.append(f"{FH_RULE_EXPECTED_FORMAT_UNREGISTERED}: artifact kind {kind} path_kind must be run_relative")
        if artifact.get("required") is not True:
            errors.append(f"{FH_RULE_EXPECTED_FORMAT_UNREGISTERED}: artifact kind {kind} must be required=true")

    return ManifestValidationResult(not errors, "broken" if errors else result.status, errors)


def compute_manifest_sha256(manifest: dict[str, Any]) -> str:
    """计算排除 manifest_sha256 自身后的 canonical hash。"""
    payload = deepcopy(manifest)
    payload.pop("manifest_sha256", None)
    return sha256_text(canonical_json(payload))


def artifact_from_file(
    run_dir: Path,
    *,
    artifact_id: str,
    kind: str,
    path: str,
    schema_id: str | None,
    schema_version: str | None = "1.0.0",
    required: bool = True,
    producer_result_id: str | None = None,
    description: str | None = None,
) -> dict[str, Any]:
    """从 run-relative 文件生成 Artifact Object，并回填真实 hash/size。"""
    artifact_path = resolve_run_relative_path(run_dir, path)
    if required and not artifact_path.exists():
        raise ValueError(f"required artifact 文件不存在：{path}")
    sha256 = compute_file_sha256(artifact_path) if artifact_path.exists() and artifact_path.is_file() else ""
    size_bytes = artifact_path.stat().st_size if artifact_path.exists() and artifact_path.is_file() else 0
    artifact = {
        "artifact_id": artifact_id,
        "kind": kind,
        "path": path,
        "path_kind": "run_relative",
        "schema_id": schema_id,
        "schema_version": schema_version,
        "sha256": sha256,
        "size_bytes": size_bytes,
        "required": required,
        "producer_result_id": producer_result_id,
    }
    if description is not None:
        artifact["description"] = description
    return artifact


def evidence_from_artifact(
    *,
    evidence_id: str,
    artifact: dict[str, Any],
    kind: str,
    depends_on: list[str] | None = None,
    summary: str = "",
) -> dict[str, Any]:
    """从 Artifact Object 派生 Evidence Object，保持冗余字段一致。"""
    evidence = {
        "evidence_id": evidence_id,
        "artifact_id": artifact.get("artifact_id"),
        "kind": kind,
        "depends_on": list(depends_on or []),
        "summary": summary,
    }
    for field_name in EVIDENCE_REDUNDANT_FIELDS:
        evidence[field_name] = artifact.get(field_name)
    return evidence


def build_evidence_manifest(
    *,
    run_id: str,
    generation: str,
    artifacts: list[dict[str, Any]],
    evidence: list[dict[str, Any]],
    manifest_id: str | None = None,
    result_ids: set[str] | None = None,
    relation_object_ids: dict[str, set[str]] | None = None,
    relations: list[dict[str, Any]] | None = None,
    warnings: list[dict[str, Any]] | None = None,
    blockers: list[dict[str, Any]] | None = None,
    generated_at: str = "2026-05-07T00:00:00+08:00",
) -> dict[str, Any]:
    """构造单代 evidence manifest 并写入自 hash。"""
    if generation not in GENERATION_PATHS:
        raise ValueError(f"未知 manifest_generation：{generation}")
    manifest = {
        "schema_id": "evidence-manifest",
        "schema_version": "2.0.0",
        "contract_version": "officecli",
        "run_id": run_id,
        "status": "complete_with_warnings" if warnings else "complete",
        "manifest_generation": generation,
        "manifest_id": manifest_id or f"EM-{run_id}-{generation}",
        "generated_at": generated_at,
        "manifest_sha256": None,
        "artifacts": deepcopy(artifacts),
        "evidence": deepcopy(evidence),
        "relations": deepcopy(relations or []),
        "warnings": deepcopy(warnings or []),
        "blockers": deepcopy(blockers or []),
    }
    if manifest.get("contract_version") == "officecli":
        import datetime
        now = datetime.datetime.now(datetime.timezone.utc).isoformat()
        present = {a.get("kind") for a in (artifacts or []) if isinstance(a, dict)}
        req = {"snapshot", "plan", "request", "result", "review"}
        missing = req - present
        manifest.setdefault("completeness", {
            "required_kinds": sorted(req), "present_kinds": sorted(present & req),
            "missing_kinds": sorted(missing),
        })
        manifest.setdefault("gate_check", {
            "gate_id": "evidence-manifest-officecli",
            "status": "passed" if not missing else "failed",
            "checked_at": now, "predicate_version": "1.0.0",
            "evidence_refs": [],
            "failed_codes": [f"missing:{k}" for k in sorted(missing)],
        })
    validation = validate_evidence_manifest(
        manifest, result_ids=result_ids, relation_object_ids=relation_object_ids,
    )
    if not validation.valid:
        manifest["status"] = "broken"
        manifest["blockers"] = [_blocker("EVIDENCE-MANIFEST-BROKEN", error) for error in validation.errors]
    manifest["manifest_sha256"] = compute_manifest_sha256(manifest)
    return manifest


def build_evidence_manifest_officecli(
    run_id: str, artifacts: list[dict[str, Any]],
    required_kinds: list[str] | None = None,
    relations: list[dict[str, Any]] | None = None,
    *, manifest_id: str | None = None, generated_at: str | None = None,
) -> dict[str, Any]:
    """独立构建严格 officecli evidence manifest，不复用 legacy 三代 manifest 对象。"""
    present = {a.get("kind") for a in artifacts if isinstance(a, dict)}
    required = set(
        ["snapshot", "plan", "request", "result", "review", "toc_acceptance"]
        if required_kinds is None else required_kinds
    )
    missing = required - present
    now = generated_at or datetime.now(timezone.utc).isoformat()
    manifest: dict[str, Any] = {
        "schema_id": "evidence-manifest",
        "schema_version": "2.0.0",
        "contract_version": "officecli",
        "run_id": run_id,
        "status": "broken" if missing else "complete",
        "manifest_id": manifest_id or f"EM-{run_id}",
        "generated_at": now,
        "manifest_sha256": None,
        "artifacts": deepcopy(artifacts),
        "relations": deepcopy(relations or []),
        "completeness": {
        "required_kinds": sorted(required),
        "present_kinds": sorted(present & required),
            "missing_kinds": sorted(missing),
        },
        "gate_check": {
            "gate_id": "evidence-manifest-officecli",
            "status": "failed" if missing else "passed",
            "checked_at": now,
            "predicate_version": "1.0.0",
            "evidence_refs": [],
            "failed_codes": [f"missing:{k}" for k in sorted(missing)],
        },
        "warnings": [],
        "blockers": ([ _blocker("EVIDENCE-MANIFEST-INCOMPLETE", f"缺少证据类型：{kind}") for kind in sorted(missing) ]),
    }
    manifest["manifest_sha256"] = compute_manifest_sha256(manifest)
    validation = validate_evidence_manifest_officecli(manifest)
    if not validation.valid:
        raise ValueError(f"officecli evidence manifest 未通过校验：{validation.errors}")
    return manifest


def _validate_officecli_artifact_ref(value: Any, field_name: str) -> list[str]:
    """严格校验 officecli ArtifactRef。"""
    errors: list[str] = []
    if not isinstance(value, dict):
        return [f"{field_name} must be ArtifactRef object"]
    missing = sorted(OFFICECLI_ARTIFACT_REQUIRED_FIELDS - set(value))
    unknown = sorted(set(value) - OFFICECLI_ARTIFACT_ALLOWED_FIELDS)
    if missing:
        errors.append(f"{field_name} missing fields: {', '.join(missing)}")
    if unknown:
        errors.append(f"{field_name} has unknown fields: {', '.join(unknown)}")
    if not isinstance(value.get("artifact_id"), str) or not value.get("artifact_id"):
        errors.append(f"{field_name}.artifact_id must be non-empty string")
    if value.get("kind") not in OFFICECLI_ARTIFACT_KINDS:
        errors.append(f"{field_name}.kind is not allowed: {value.get('kind')}")
    relative_path = value.get("relative_path")
    if not isinstance(relative_path, str) or not relative_path:
        errors.append(f"{field_name}.relative_path must be non-empty string")
    else:
        normalized = relative_path.replace("\\", "/")
        if Path(relative_path).is_absolute() or normalized.startswith("../") or "/../" in normalized:
            errors.append(f"{field_name}.relative_path must not escape: {relative_path}")
    if not isinstance(value.get("sha256"), str) or re.fullmatch(r"[a-f0-9]{64}", value.get("sha256", "")) is None:
        errors.append(f"{field_name}.sha256 must be lowercase sha256")
    size_bytes = value.get("size_bytes")
    if not isinstance(size_bytes, int) or isinstance(size_bytes, bool) or size_bytes < 0:
        errors.append(f"{field_name}.size_bytes must be non-negative integer")
    return errors


def validate_evidence_manifest_officecli(manifest: dict[str, Any]) -> ManifestValidationResult:
    """按 officecli schema 语义严格校验 evidence manifest。"""
    errors: list[str] = []
    required_fields = {
        "schema_id", "schema_version", "contract_version", "run_id", "status", "manifest_id",
        "generated_at", "artifacts", "relations", "completeness", "gate_check",
    }
    allowed_fields = required_fields | {"manifest_sha256", "warnings", "blockers"}
    if not isinstance(manifest, dict):
        return ManifestValidationResult(False, "broken", ["manifest must be object"])
    missing_fields = sorted(required_fields - set(manifest))
    unknown_fields = sorted(set(manifest) - allowed_fields)
    if missing_fields:
        errors.append(f"missing required fields: {', '.join(missing_fields)}")
    if unknown_fields:
        errors.append(f"unknown fields: {', '.join(unknown_fields)}")
    if manifest.get("schema_id") != "evidence-manifest":
        errors.append("schema_id must be evidence-manifest")
    if manifest.get("schema_version") != "2.0.0" or manifest.get("contract_version") != "officecli":
        errors.append("officecli evidence manifest version mismatch")
    if manifest.get("status") not in ALLOWED_MANIFEST_STATUSES:
        errors.append(f"status is not allowed: {manifest.get('status')}")
    artifacts = manifest.get("artifacts")
    artifact_ids: set[str] = set()
    if not isinstance(artifacts, list):
        errors.append("artifacts must be array")
        artifacts = []
    for index, artifact in enumerate(artifacts):
        errors.extend(_validate_officecli_artifact_ref(artifact, f"artifacts[{index}]"))
        if isinstance(artifact, dict) and isinstance(artifact.get("artifact_id"), str):
            if artifact["artifact_id"] in artifact_ids:
                errors.append(f"duplicate artifact_id: {artifact['artifact_id']}")
            artifact_ids.add(artifact["artifact_id"])
    relations = manifest.get("relations")
    if not isinstance(relations, list):
        errors.append("relations must be array")
        relations = []
    for index, relation in enumerate(relations):
        if not isinstance(relation, dict):
            errors.append(f"relations[{index}] must be object")
            continue
        required_relation = {"from_artifact_id", "to_artifact_id", "relation"}
        if set(relation) != required_relation:
            errors.append(f"relations[{index}] fields must be {sorted(required_relation)}")
        for key in ("from_artifact_id", "to_artifact_id"):
            if relation.get(key) not in artifact_ids:
                errors.append(f"relations[{index}].{key} does not resolve artifact")
        if relation.get("relation") not in OFFICECLI_RELATION_TYPES:
            errors.append(f"relations[{index}].relation is not allowed: {relation.get('relation')}")
    completeness = manifest.get("completeness")
    if not isinstance(completeness, dict) or set(completeness) != {"required_kinds", "present_kinds", "missing_kinds"}:
        errors.append("completeness fields are invalid")
    else:
        for key in ("required_kinds", "present_kinds", "missing_kinds"):
            if not isinstance(completeness.get(key), list) or any(not isinstance(item, str) for item in completeness[key]):
                errors.append(f"completeness.{key} must be string array")
        if completeness.get("missing_kinds") and manifest.get("status") != "broken":
            errors.append("status must be broken when completeness has missing_kinds")
    gate = manifest.get("gate_check")
    if not isinstance(gate, dict):
        errors.append("gate_check must be object")
    else:
        gate_required = {"gate_id", "status", "checked_at", "predicate_version", "evidence_refs", "failed_codes"}
        if set(gate) != gate_required:
            errors.append("gate_check fields are invalid")
        if gate.get("status") not in {"passed", "failed", "blocked"}:
            errors.append("gate_check.status is invalid")
        if completeness and completeness.get("missing_kinds") and gate.get("status") == "passed":
            errors.append("gate_check must not pass when completeness has missing_kinds")
    manifest_hash = manifest.get("manifest_sha256")
    if manifest_hash is not None and manifest_hash != compute_manifest_sha256(manifest):
        errors.append("manifest_sha256 does not match canonical manifest hash")
    return ManifestValidationResult(not errors, "broken" if errors else str(manifest.get("status")), errors)



def validate_evidence_manifest(
    manifest: dict[str, Any],
    *,
    run_dir: Path | None = None,
    result_ids: set[str] | None = None,
    relation_object_ids: dict[str, set[str]] | None = None,
) -> ManifestValidationResult:
    """校验 evidence manifest 的字段、hash、文件和 producer 链路。"""
    errors: list[str] = []
    if not isinstance(manifest, dict):
        return ManifestValidationResult(False, "broken", ["manifest must be object"])

    for field_name in TOP_LEVEL_REQUIRED_FIELDS:
        if field_name not in manifest:
            errors.append(f"{field_name} is required")

    if manifest.get("schema_id") != "evidence-manifest":
        errors.append("schema_id must be evidence-manifest")
    if manifest.get("contract_version") not in {"legacy", "officecli"}:
        errors.append("contract_version must be legacy or officecli")
    if manifest.get("status") not in ALLOWED_MANIFEST_STATUSES:
        errors.append(f"status is not allowed: {manifest.get('status')}")

    generation = manifest.get("manifest_generation")
    if generation not in GENERATION_PATHS:
        errors.append("manifest_generation must be pre_acceptance/post_acceptance/reporting")

    artifacts = manifest.get("artifacts")
    evidence_items = manifest.get("evidence")
    relations = manifest.get("relations")
    if not isinstance(artifacts, list):
        errors.append("artifacts must be array")
        artifacts = []
    if not isinstance(evidence_items, list):
        errors.append("evidence must be array")
        evidence_items = []
    if not isinstance(relations, list):
        errors.append("relations must be array")
        relations = []
    for array_name in ("warnings", "blockers"):
        if not isinstance(manifest.get(array_name), list):
            errors.append(f"{array_name} must be array")

    artifact_index: dict[str, dict[str, Any]] = {}
    for index, artifact in enumerate(artifacts):
        if not isinstance(artifact, dict):
            errors.append(f"artifacts[{index}] must be object")
            continue
        artifact_id = artifact.get("artifact_id")
        if not artifact_id:
            errors.append(f"artifacts[{index}].artifact_id is required")
            continue
        if artifact_id in artifact_index:
            errors.append(f"duplicate artifact_id: {artifact_id}")
        artifact_index[artifact_id] = artifact
        if generation == "pre_acceptance":
            if artifact.get("kind") in PRE_ACCEPTANCE_FORBIDDEN_ARTIFACT_KINDS:
                errors.append(f"pre_acceptance must not contain artifact kind {artifact.get('kind')}")
            if artifact.get("schema_id") in PRE_ACCEPTANCE_FORBIDDEN_SCHEMA_IDS:
                errors.append(f"pre_acceptance must not contain schema_id {artifact.get('schema_id')}")
        if generation == "post_acceptance" and artifact.get("kind") in POST_ACCEPTANCE_FORBIDDEN_ARTIFACT_KINDS:
            errors.append(f"post_acceptance must not contain artifact kind {artifact.get('kind')}")
        if artifact.get("required") is True:
            for field_name in ("path", "path_kind", "sha256", "size_bytes", "producer_result_id"):
                if field_name not in artifact or artifact.get(field_name) in {None, ""}:
                    errors.append(f"artifacts[{index}].{field_name} is required when required=true")
            if artifact.get("path_kind") != "run_relative":
                errors.append(f"artifacts[{index}].path_kind must be run_relative")
            if run_dir is not None and artifact.get("path"):
                try:
                    artifact_path = resolve_run_relative_path(run_dir, artifact["path"])
                except ValueError as exc:
                    errors.append(f"artifacts[{index}].path invalid: {exc}")
                else:
                    if not artifact_path.exists() or not artifact_path.is_file():
                        errors.append(f"artifacts[{index}].path does not exist: {artifact.get('path')}")
                    else:
                        if artifact.get("sha256") != compute_file_sha256(artifact_path):
                            errors.append(f"artifacts[{index}].sha256 does not match file")
                        if artifact.get("size_bytes") != artifact_path.stat().st_size:
                            errors.append(f"artifacts[{index}].size_bytes does not match file")
        producer_result_id = artifact.get("producer_result_id")
        if producer_result_id is not None and result_ids is not None and producer_result_id not in result_ids:
            errors.append(f"artifacts[{index}].producer_result_id does not resolve: {producer_result_id}")

    evidence_index: dict[str, dict[str, Any]] = {
        item["evidence_id"]: item
        for item in evidence_items
        if isinstance(item, dict) and item.get("evidence_id")
    }
    if len(evidence_index) != len(
        [item for item in evidence_items if isinstance(item, dict) and item.get("evidence_id")]
    ):
        errors.append("duplicate evidence_id")
    for index, item in enumerate(evidence_items):
        if not isinstance(item, dict):
            errors.append(f"evidence[{index}] must be object")
            continue
        for field_name in EVIDENCE_REQUIRED_FIELDS:
            if field_name not in item:
                errors.append(f"evidence[{index}].{field_name} is required")
        evidence_id = item.get("evidence_id")
        artifact = artifact_index.get(item.get("artifact_id"))
        if artifact is None:
            errors.append(f"evidence[{index}].artifact_id does not resolve: {item.get('artifact_id')}")
        else:
            for field_name in EVIDENCE_REDUNDANT_FIELDS:
                if item.get(field_name) != artifact.get(field_name):
                    errors.append(f"evidence[{index}].{field_name} does not match artifact")
        if item.get("required") is True and item.get("artifact_id") not in artifact_index:
            errors.append(f"evidence[{index}] required artifact is missing")
        producer_result_id = item.get("producer_result_id")
        if producer_result_id is not None and result_ids is not None and producer_result_id not in result_ids:
            errors.append(f"evidence[{index}].producer_result_id does not resolve: {producer_result_id}")
        depends_on = item.get("depends_on")
        if not isinstance(depends_on, list):
            errors.append(f"evidence[{index}].depends_on must be array")
        else:
            for dependency_id in depends_on:
                if dependency_id not in artifact_index and dependency_id not in evidence_index:
                    if result_ids is None or dependency_id not in result_ids:
                        errors.append(f"evidence[{index}].depends_on does not resolve: {dependency_id}")

    for index, relation in enumerate(relations):
        if not isinstance(relation, dict):
            errors.append(f"relations[{index}] must be object")
            continue
        for field_name in ("relation_id", "source_type", "source_id", "target_type", "target_id", "relation_type"):
            if not relation.get(field_name):
                errors.append(f"relations[{index}].{field_name} is required")
        for side in ("source", "target"):
            object_type = relation.get(f"{side}_type")
            object_id = relation.get(f"{side}_id")
            if object_type not in SUPPORTED_RELATION_OBJECT_TYPES:
                errors.append(f"relations[{index}].{side}_type is not supported in CODE-008: {object_type}")
                continue
            if object_type == "artifact" and object_id not in artifact_index:
                errors.append(f"relations[{index}].{side}_id does not resolve artifact: {object_id}")
            if object_type == "evidence" and object_id not in evidence_index:
                errors.append(f"relations[{index}].{side}_id does not resolve evidence: {object_id}")
            if object_type in {"issue", "action", "review", "acceptance"}:
                known_ids = (relation_object_ids or {}).get(object_type, set())
                if object_id not in known_ids:
                    errors.append(f"relations[{index}].{side}_id does not resolve {object_type}: {object_id}")
        if relation.get("relation_type") not in ALLOWED_RELATION_TYPES:
            errors.append(f"relations[{index}].relation_type is not allowed: {relation.get('relation_type')}")

    if generation == "reporting":
        artifact_kinds = {artifact.get("kind") for artifact in artifacts if isinstance(artifact, dict)}
        if not artifact_kinds.intersection(REPORTING_REQUIRED_ARTIFACT_KINDS):
            errors.append("reporting manifest must contain report or reporting_result artifact")

    manifest_sha256 = manifest.get("manifest_sha256")
    if manifest_sha256 is not None and manifest_sha256 != compute_manifest_sha256(manifest):
        errors.append("manifest_sha256 does not match canonical manifest hash")

    # OFFICECLI-011: officecli 必需字段强制校验
    if manifest.get("contract_version") == "officecli":
        if not isinstance(manifest.get("completeness"), dict):
            errors.append("officecli evidence manifest requires completeness")
        else:
            for k in ("required_kinds", "present_kinds", "missing_kinds"):
                if k not in manifest["completeness"]:
                    errors.append(f"officecli completeness missing {k}")
        if not isinstance(manifest.get("gate_check"), dict):
            errors.append("officecli evidence manifest requires gate_check")
        for rel in (relations or []):
            if isinstance(rel, dict):
                rt = rel.get("relation_type")
                if rt not in ALLOWED_RELATION_TYPES:
                    errors.append(f"relation_type not allowed: {rt}")
    status = "broken" if errors else str(manifest.get("status") or "complete")
    if errors and manifest.get("status") != "broken":
        errors.append("status must be broken when evidence chain has errors")
    return ManifestValidationResult(not errors, status, errors)


def write_evidence_manifest(
    run_dir: Path,
    manifest: dict[str, Any],
    *,
    result_ids: set[str] | None = None,
    relation_object_ids: dict[str, set[str]] | None = None,
) -> dict[str, Any]:
    """按 generation 固定路径原子写入 evidence manifest。"""
    generation = manifest.get("manifest_generation")
    if generation not in GENERATION_PATHS:
        raise ValueError(f"未知 manifest_generation：{generation}")
    manifest_to_write = deepcopy(manifest)
    manifest_to_write["manifest_sha256"] = compute_manifest_sha256(manifest_to_write)
    validation = validate_evidence_manifest(
        manifest_to_write,
        run_dir=run_dir,
        result_ids=result_ids or set(),
        relation_object_ids=relation_object_ids,
    )
    if not validation.valid:
        raise ValueError(f"evidence manifest 未通过校验：{validation.errors}")
    rel_path = GENERATION_PATHS[generation]
    path = resolve_run_relative_path(run_dir, rel_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(json.dumps(manifest_to_write, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temp_path, path)
    return {
        "path": rel_path,
        "sha256": compute_file_sha256(path),
        "size_bytes": path.stat().st_size,
        "manifest": manifest_to_write,
    }


def write_evidence_manifest_officecli(run_dir: Path, manifest: dict[str, Any]) -> dict[str, Any]:
    """严格校验并原子写入纯 officecli evidence manifest。"""
    manifest_to_write = deepcopy(manifest)
    manifest_to_write["manifest_sha256"] = compute_manifest_sha256(manifest_to_write)
    validation = validate_evidence_manifest_officecli(manifest_to_write)
    if not validation.valid:
        raise ValueError(f"officecli evidence manifest 未通过校验：{validation.errors}")
    path = resolve_run_relative_path(run_dir, OFFICECLI_EVIDENCE_MANIFEST_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(json.dumps(manifest_to_write, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temp_path, path)
    written = json.loads(path.read_text(encoding="utf-8"))
    readback = validate_evidence_manifest_officecli(written)
    if not readback.valid:
        raise ValueError(f"officecli evidence manifest 回读校验失败：{readback.errors}")
    return {
        "path": OFFICECLI_EVIDENCE_MANIFEST_PATH,
        "sha256": compute_file_sha256(path),
        "size_bytes": path.stat().st_size,
        "manifest": written,
    }


__all__ = [
    "GENERATION_PATHS",
    "OFFICECLI_EVIDENCE_MANIFEST_PATH",
    "ManifestValidationResult",
    "RULE_PACKAGING_REQUIRED_ARTIFACTS",
    "artifact_from_file",
    "build_evidence_manifest",
    "compute_manifest_sha256",
    "evidence_from_artifact",
    "validate_rule_packaging_expected_artifacts",
    "validate_evidence_manifest",
    "validate_evidence_manifest_officecli",
    "build_evidence_manifest_officecli",
    "write_evidence_manifest",
    "write_evidence_manifest_officecli",
]
