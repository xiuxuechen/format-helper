"""skill-result 统一读写与校验工具（CODE-006A）。"""

from __future__ import annotations

import hashlib
import json
import os
import re
from copy import deepcopy
from pathlib import Path
from typing import Any

from scripts.validation.validate_skill_result import validate_skill_result


FH_SKILL_RESULT_MISSING = "FH-SKILL-RESULT-MISSING"
FH_SKILL_RESULT_INVALID = "FH-SKILL-RESULT-INVALID"

# CODE-018: SLOT_CONTRACT_DESIGN.md §8.1 槽位相关错误码
FH_SLOT_CONTRACT_NOT_FOUND = "FH-SLOT-CONTRACT-NOT-FOUND"
FH_SLOT_FACTS_MISSING = "FH-SLOT-FACTS-MISSING"
FH_SLOT_FACTS_INVALID = "FH-SLOT-FACTS-INVALID"
FH_SLOT_FACTS_UNRESOLVED = "FH-SLOT-FACTS-UNRESOLVED"
FH_SLOT_FACTS_CONFLICT = "FH-SLOT-FACTS-CONFLICT"
FH_SLOT_FACTS_HASH_MISMATCH = "FH-SLOT-FACTS-HASH-MISMATCH"
FH_RULE_CONFIRM_GATE_PENDING = "FH-RULE-CONFIRM-GATE-PENDING"
FH_RULE_EXPECTED_FORMAT_UNREGISTERED = "FH-RULE-EXPECTED-FORMAT-UNREGISTERED"

NEXT_ACTION_REQUIRED_FIELDS = [
    "kind",
    "stage",
    "skill_name",
    "target_result_id",
    "target_error_code",
    "source_result_id",
    "override_reason",
    "resume_from_stage",
    "idempotency_key",
    "planned_idempotency_key",
    "reason",
    "required_inputs",
    "user_message",
]


def canonical_json(data: dict[str, Any]) -> str:
    """返回稳定 JSON 字符串，用于 hash 和幂等比较。"""
    return json.dumps(data, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def sha256_text(text: str) -> str:
    """计算 UTF-8 文本 sha256。"""
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def compute_result_file_sha256(result: dict[str, Any]) -> str:
    """计算 result canonical 内容 hash。"""
    return sha256_text(canonical_json(result))


def compute_file_sha256(path: Path) -> str:
    """计算文件真实字节 sha256。"""
    return hashlib.sha256(path.read_bytes()).hexdigest()


def compute_result_chain_head(previous_head: str | None, result: dict[str, Any], result_sha256: str | None = None) -> str:
    """按 41-§11.1 复算 result chain head。"""
    payload = {
        "previous_head": previous_head,
        "result_id": result.get("result_id"),
        "result_sha256": result_sha256 or compute_result_file_sha256(result),
        "stage": result.get("stage"),
        "status": result.get("status"),
        "gate_passed": result.get("gate_passed"),
    }
    return sha256_text(canonical_json(payload))


def resolve_run_relative_path(run_dir: Path, result_path: str | Path) -> Path:
    """解析 run-relative 路径并阻断绝对路径或目录逃逸。"""
    path = Path(result_path)
    if path.is_absolute() or ".." in path.parts:
        raise ValueError(f"result 路径必须是 run-relative 且不得逃逸 run_dir：{result_path}")
    resolved_run_dir = run_dir.resolve()
    resolved_path = (resolved_run_dir / path).resolve()
    if resolved_run_dir != resolved_path and resolved_run_dir not in resolved_path.parents:
        raise ValueError(f"result 路径逃逸 run_dir：{result_path}")
    return resolved_path


def safe_skill_name(skill_name: str) -> str:
    """将 skill 名称收敛为 result 文件名片段。"""
    cleaned = re.sub(r"[^A-Za-z0-9_.-]+", "-", skill_name).strip("-")
    if not cleaned:
        raise ValueError("skill_name 不能生成安全文件名")
    return cleaned


def validate_common_skill_result_contract(
    result: dict[str, Any],
    *,
    run_dir: Path | None = None,
) -> list[str]:
    """校验 CODE-006A 公共闭环字段与状态信封约束。"""
    errors: list[str] = []
    validation = validate_skill_result(result)
    if not validation.valid:
        errors.extend(f"{item['field']}: {item['message']}" for item in validation.errors)

    for field in ("result_id", "run_id", "idempotency_key", "stage", "status"):
        if not result.get(field):
            errors.append(f"{field} must be non-empty")
    for field in ("order", "attempt"):
        value = result.get(field)
        if not isinstance(value, int) or value <= 0:
            errors.append(f"{field} must be positive integer")

    runtime = result.get("runtime")
    if not isinstance(runtime, dict):
        errors.append("runtime must be object")
    else:
        for field in ("started_at", "duration_ms", "platform"):
            if field not in runtime:
                errors.append(f"runtime.{field} is required")
        if "ended_at" not in runtime and "finished_at" not in runtime:
            errors.append("runtime.ended_at or runtime.finished_at is required")

    next_action = result.get("next_action")
    if not isinstance(next_action, dict):
        errors.append("next_action must be object")
    else:
        for field in NEXT_ACTION_REQUIRED_FIELDS:
            if field not in next_action:
                errors.append(f"next_action.{field} is required")
        if next_action.get("source_result_id") != result.get("result_id"):
            errors.append("skill-result.next_action.source_result_id must equal result_id")
        if next_action.get("kind") in {"run_skill", "retry", "manual_recover"}:
            if not next_action.get("idempotency_key"):
                errors.append("next_action.idempotency_key is required")
            if not next_action.get("planned_idempotency_key"):
                errors.append("next_action.planned_idempotency_key is required")
        if next_action.get("kind") in {"retry", "manual_recover"}:
            if not next_action.get("target_result_id") and not next_action.get("target_error_code"):
                errors.append("retry/manual_recover requires target_result_id or target_error_code")

    artifacts = result.get("artifacts")
    if not isinstance(artifacts, list):
        errors.append("artifacts must be array")
    else:
        for index, artifact in enumerate(artifacts):
            if not isinstance(artifact, dict):
                errors.append(f"artifacts[{index}] must be object")
                continue
            required_fields = {
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
            }
            for field in required_fields:
                if field not in artifact:
                    errors.append(f"artifacts[{index}].{field} is required")
            if artifact.get("required") is True:
                for field in ("path", "path_kind", "sha256", "size_bytes"):
                    if artifact.get(field) in {None, ""}:
                        errors.append(f"artifacts[{index}].{field} is required when artifact.required=true")
                path_value = artifact.get("path")
                path_kind = artifact.get("path_kind")
                if path_kind != "run_relative":
                    errors.append(f"artifacts[{index}].path_kind must be run_relative for required artifact")
                if path_value:
                    try:
                        artifact_path = resolve_run_relative_path(run_dir, path_value) if run_dir is not None else Path(path_value)
                    except ValueError as exc:
                        errors.append(f"artifacts[{index}].path is not run-relative: {exc}")
                    else:
                        if run_dir is not None:
                            if not artifact_path.exists():
                                errors.append(f"artifacts[{index}].path does not exist: {path_value}")
                            elif artifact_path.is_file():
                                actual_sha256 = compute_file_sha256(artifact_path)
                                actual_size = artifact_path.stat().st_size
                                if artifact.get("sha256") != actual_sha256:
                                    errors.append(f"artifacts[{index}].sha256 does not match file")
                                if artifact.get("size_bytes") != actual_size:
                                    errors.append(f"artifacts[{index}].size_bytes does not match file")
            if artifact.get("producer_result_id") not in {None, result.get("result_id")}:
                errors.append(f"artifacts[{index}].producer_result_id must equal result_id or null")

    for optional_array in ("warnings", "evidence_refs"):
        if optional_array in result and not isinstance(result[optional_array], list):
            errors.append(f"{optional_array} must be array when present")
    if "metrics" in result:
        try:
            json.dumps(result["metrics"], ensure_ascii=False)
        except (TypeError, ValueError):
            errors.append("metrics must be JSON serializable")

    return errors


def load_skill_result_file(
    result_path: str | Path,
    run_dir: Path | None = None,
    expected_sha256: str | None = None,
) -> tuple[dict[str, Any], str]:
    """读取 result 文件、执行公共校验，并返回 result 与真实文件 sha256。"""
    path = resolve_run_relative_path(run_dir, result_path) if run_dir is not None else Path(result_path)
    if not path.exists():
        raise ValueError(f"result 文件不存在，无法计算真实 sha256：{result_path}")
    actual_sha256 = compute_file_sha256(path)
    if expected_sha256 is not None and expected_sha256 != actual_sha256:
        raise ValueError("result_sha256 must match real result file sha256")
    try:
        result = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"result 文件不是合法 JSON：{result_path}") from exc
    errors = validate_common_skill_result_contract(result, run_dir=run_dir)
    if errors:
        raise ValueError(f"skill-result 未通过公共契约校验：{errors}")
    return result, actual_sha256


def build_synthetic_failure(
    *,
    run_id: str,
    order: int,
    target_stage: str,
    target_skill_name: str,
    missing_result_path: str,
    attempt: int = 1,
    error_code: str = FH_SKILL_RESULT_MISSING,
) -> dict[str, Any]:
    """生成缺失或无效 result 的 synthetic failure 信封。"""
    idempotency_key = sha256_text(
        f"{run_id}:{target_stage}:{target_skill_name}:attempt{attempt}:{error_code}:{missing_result_path}"
    )
    result_id = f"SF-{sha256_text(idempotency_key)[:12]}"
    return {
        "schema_id": "skill-result",
        "schema_version": "1.0.0",
        "contract_version": "v4",
        "result_id": result_id,
        "run_id": run_id,
        "order": order,
        "attempt": attempt,
        "idempotency_key": idempotency_key,
        "stage": target_stage,
        "status": "synthetic_failure",
        "schema_valid": True,
        "gate_passed": False,
        "gate_check": {
            "status": "failed",
            "passed": False,
            "blockers": [
                {
                    "code": error_code,
                    "message": f"目标 result 缺失或不可解析：{missing_result_path}",
                    "stage": target_stage,
                    "skill_name": target_skill_name,
                }
            ],
        },
        "validation": {
            "schema_valid": True,
            "path_valid": False,
            "evidence_valid": False,
            "errors": [{"field": missing_result_path, "code": error_code}],
        },
        "artifacts": [],
        "blockers": [
            {
                "code": error_code,
                "message": f"目标 result 缺失或不可解析：{missing_result_path}",
                "stage": target_stage,
                "skill_name": target_skill_name,
            }
        ],
        "error": {
            "code": error_code,
            "message": f"目标 result 缺失或不可解析：{missing_result_path}",
            "severity": "block",
        },
        "next_action": {
            "kind": "manual_recover",
            "stage": target_stage,
            "skill_name": target_skill_name,
            "target_result_id": result_id,
            "target_error_code": error_code,
            "source_result_id": result_id,
            "override_reason": None,
            "resume_from_stage": target_stage,
            "idempotency_key": sha256_text(f"{idempotency_key}:next_action"),
            "planned_idempotency_key": sha256_text(f"{idempotency_key}:manual_recover"),
            "reason": "内部 skill 未写出可用 result，需人工恢复或重试",
            "required_inputs": [{"path": missing_result_path, "path_kind": "run_relative"}],
            "user_message": "内部步骤结果缺失，已生成 synthetic failure 并阻断推进。",
        },
        "runtime": {
            "started_at": "2026-05-07T00:00:00+08:00",
            "finished_at": "2026-05-07T00:00:00+08:00",
            "duration_ms": 0,
            "platform": "codex",
        },
    }


def skill_result_rel_path(order: int, skill_name: str) -> str:
    """生成 canonical skill-result 相对路径。"""
    if not isinstance(order, int) or order <= 0:
        raise ValueError("order must be positive integer")
    return f"logs/skill_results/{order:03d}_{safe_skill_name(skill_name)}.result.json"


def write_skill_result_atomic(
    run_dir: Path,
    result: dict[str, Any],
    *,
    skill_name: str | None = None,
    result_rel_path: str | None = None,
) -> dict[str, Any]:
    """校验并原子写入 skill-result，支持同幂等键重放识别。"""
    errors = validate_common_skill_result_contract(result, run_dir=run_dir)
    if errors:
        raise ValueError(f"skill-result 未通过公共契约校验：{errors}")
    rel_path = result_rel_path or skill_result_rel_path(int(result["order"]), skill_name or result.get("skill_name") or "skill-result")
    path = resolve_run_relative_path(run_dir, rel_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    serialized = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    if path.exists():
        existing, existing_sha256 = load_skill_result_file(rel_path, run_dir)
        if existing.get("idempotency_key") == result.get("idempotency_key") and canonical_json(existing) == canonical_json(result):
            return {
                "path": rel_path,
                "sha256": existing_sha256,
                "idempotent_replay": True,
                "result": existing,
            }
        raise ValueError(f"skill-result 已存在且不是等价幂等重放：{rel_path}")

    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(serialized, encoding="utf-8")
    os.replace(temp_path, path)
    return {
        "path": rel_path,
        "sha256": compute_file_sha256(path),
        "idempotent_replay": False,
        "result": deepcopy(result),
    }
