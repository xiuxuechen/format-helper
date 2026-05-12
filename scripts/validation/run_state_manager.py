"""run-state 写入、链路和恢复工具（CODE-006）。

实现范围限定为 50_DEV_PLAN.md 的 CODE-006：
- `logs/state.yaml` 原子写入前校验
- `result_chain_head` 复算
- `next_action` 与 `planned_idempotency_key` 校验
- `run_id` 恢复判定
- 缺失或不可解析 result 时生成 synthetic failure 输入对象
"""

from __future__ import annotations

import json
import os
from copy import deepcopy
from pathlib import Path
from typing import Any

from scripts.validation.skill_result_io import (
    FH_SKILL_RESULT_INVALID,
    FH_SKILL_RESULT_MISSING,
    build_synthetic_failure,
    canonical_json,
    compute_file_sha256,
    compute_result_chain_head,
    compute_result_file_sha256,
    load_skill_result_file,
    resolve_run_relative_path,
    sha256_text,
    write_skill_result_atomic,
)
from scripts.validation.validate_run_state import validate_run_state
from scripts.validation.validate_skill_result import validate_skill_result


def load_result_file(
    result_path: str | Path,
    run_dir: Path | None = None,
    expected_sha256: str | None = None,
) -> tuple[dict[str, Any], str]:
    """兼容 CODE-006 调用名，实际委托 CODE-006A 公共工具。"""
    return load_skill_result_file(result_path, run_dir, expected_sha256)


def ensure_result_matches_planned_key(state: dict[str, Any], result: dict[str, Any]) -> None:
    """应用 result 前校验上一状态 planned key。"""
    prior_next = state.get("next_action", {})
    if not isinstance(prior_next, dict):
        return
    if prior_next.get("kind") not in {"run_skill", "retry", "manual_recover"}:
        return
    planned_key = prior_next.get("planned_idempotency_key")
    if planned_key and result.get("idempotency_key") != planned_key:
        raise ValueError("result.idempotency_key must match previous next_action.planned_idempotency_key")


def apply_result_to_state(
    state: dict[str, Any],
    result: dict[str, Any],
    result_path: str | Path,
    *,
    run_dir: Path | None = None,
    result_sha256: str | None = None,
    enforce_planned_key: bool = True,
) -> dict[str, Any]:
    """把已通过校验的 result 原子应用到 run-state 候选对象。"""
    file_result, resolved_result_sha256 = load_result_file(result_path, run_dir, result_sha256)
    if canonical_json(file_result) != canonical_json(result):
        raise ValueError("传入 result 必须与 result 文件内容一致")
    result_validation = validate_skill_result(result)
    if not result_validation.valid:
        raise ValueError(f"skill-result 未通过校验：{result_validation.errors}")
    if enforce_planned_key:
        ensure_result_matches_planned_key(state, result)
    next_state = deepcopy(state)
    result_id = result["result_id"]
    result_rel_path = str(result_path).replace("\\", "/")
    next_state["stage"] = result["stage"]
    next_state["status"] = "blocked" if result["status"] in {"blocked", "synthetic_failure"} else "wip"
    next_state.setdefault("skill_results", [])
    if result_rel_path not in next_state["skill_results"]:
        next_state["skill_results"].append(result_rel_path)
    next_state["last_result_id"] = result_id
    next_state["applied_result_id"] = result_id
    next_state["result_chain_head"] = compute_result_chain_head(
        state.get("result_chain_head"),
        result,
        resolved_result_sha256,
    )
    next_state["safe_outputs"] = list(result.get("artifacts", []))
    next_action = deepcopy(result.get("next_action", {}))
    next_action["source_result_id"] = result_id
    next_state["next_action"] = next_action
    next_state["blockers"] = list(result.get("blockers", []))
    validation = validate_run_state(next_state)
    if not validation.valid:
        raise ValueError(f"run-state 候选对象未通过校验：{validation.errors}")
    return next_state


STAGE_ORDER = [
    "init",
    "rule_selection",
    "fact_extraction",
    "semantic_strategy",
    "rule_packaging",
    "format_audit",
    "repair_planning",
    "manual_review",
    "repair_execution",
    "after_snapshot",
    "review",
    "toc_acceptance",
    "final_acceptance",
    "reporting",
    "completed",
]


def stage_index(stage: str | None) -> int:
    """返回阶段顺序；未知阶段返回极大值。"""
    try:
        return STAGE_ORDER.index(stage or "")
    except ValueError:
        return 10_000


def validate_next_action_contract(
    *,
    state: dict[str, Any],
    result_index: dict[str, dict[str, Any]] | None = None,
    artifact_index: set[str] | None = None,
) -> list[str]:
    """校验 next_action 与恢复幂等约束。"""
    errors: list[str] = []
    next_action = state.get("next_action", {})
    if not isinstance(next_action, dict):
        return ["next_action must be object"]
    kind = next_action.get("kind")
    source_result_id = next_action.get("source_result_id")
    applied_result_id = state.get("applied_result_id")

    if kind in {"retry", "manual_recover"}:
        target_result_id = next_action.get("target_result_id")
        target_error_code = next_action.get("target_error_code")
        if not target_result_id and not target_error_code:
            errors.append("retry/manual_recover requires target_result_id or target_error_code")
        if not next_action.get("resume_from_stage"):
            errors.append("retry/manual_recover requires resume_from_stage")
        elif stage_index(next_action.get("resume_from_stage")) > stage_index(state.get("stage")):
            errors.append("resume_from_stage must not be later than current state.stage")
        if result_index is not None:
            target_result = result_index.get(target_result_id) if target_result_id else None
            if target_result_id and target_result is None:
                errors.append(f"target_result_id is not resolvable: {target_result_id}")
            if target_error_code:
                candidate_results = [target_result] if target_result is not None else list(result_index.values())
                if not any(result_has_error_code(candidate, target_error_code) for candidate in candidate_results if candidate):
                    errors.append(f"target_error_code is not resolvable: {target_error_code}")

    if kind in {"run_skill", "retry", "manual_recover"}:
        if not next_action.get("idempotency_key"):
            errors.append(f"{kind} requires idempotency_key")
        if not next_action.get("planned_idempotency_key"):
            errors.append(f"{kind} requires planned_idempotency_key")

    required_inputs = next_action.get("required_inputs", [])
    if not isinstance(required_inputs, list):
        errors.append("next_action.required_inputs must be array")
    elif artifact_index is not None:
        for item in required_inputs:
            if isinstance(item, str) and item not in artifact_index:
                errors.append(f"required_input is not resolvable: {item}")
            elif isinstance(item, dict):
                artifact_id = item.get("artifact_id")
                if artifact_id and artifact_id not in artifact_index:
                    errors.append(f"required_input artifact_id is not resolvable: {artifact_id}")
                path = item.get("path")
                if path is not None and not is_safe_run_relative_path(path):
                    errors.append(f"required_input path is not run-relative: {path}")

    if applied_result_id is not None and source_result_id != applied_result_id:
        errors.append("next_action.source_result_id must equal applied_result_id")

    if result_index and applied_result_id in result_index:
        source_result = result_index[applied_result_id]
        result_next = source_result.get("next_action", {})
        if kind == "run_skill" and next_action.get("planned_idempotency_key") != result_next.get("planned_idempotency_key"):
            errors.append("run-state planned_idempotency_key must match applied result next_action")
        copied_fields = {
            "kind",
            "stage",
            "skill_name",
            "target_result_id",
            "target_error_code",
            "resume_from_stage",
            "planned_idempotency_key",
            "required_inputs",
        }
        differs = any(next_action.get(field) != result_next.get(field) for field in copied_fields)
        if differs and not next_action.get("override_reason"):
            errors.append("overridden next_action requires override_reason")

    return errors


def result_has_error_code(result: dict[str, Any], error_code: str) -> bool:
    """检查 result 的 error 或 blockers 是否包含目标错误码。"""
    error = result.get("error")
    if isinstance(error, dict) and error.get("code") == error_code:
        return True
    for blocker in result.get("blockers", []):
        if isinstance(blocker, dict) and blocker.get("code") == error_code:
            return True
    gate_check = result.get("gate_check")
    if isinstance(gate_check, dict):
        for blocker in gate_check.get("blockers", []):
            if isinstance(blocker, dict) and blocker.get("code") == error_code:
                return True
    return False


def is_safe_run_relative_path(path_value: Any) -> bool:
    """判断 Path Object 的 path 是否为安全 run-relative 路径。"""
    if not isinstance(path_value, str) or not path_value:
        return False
    path = Path(path_value)
    return not path.is_absolute() and ".." not in path.parts


def collect_artifact_ids(state: dict[str, Any], result_index: dict[str, dict[str, Any]] | None = None) -> set[str]:
    """汇总 run-state.safe_outputs 与 skill-result.artifacts 中的 artifact_id。"""
    artifact_ids = {
        item.get("artifact_id")
        for item in state.get("safe_outputs", [])
        if isinstance(item, dict) and item.get("artifact_id")
    }
    if result_index:
        for result in result_index.values():
            for artifact in result.get("artifacts", []):
                if isinstance(artifact, dict) and artifact.get("artifact_id"):
                    artifact_ids.add(artifact["artifact_id"])
    return artifact_ids


def replay_result_chain(run_dir: Path, state: dict[str, Any]) -> dict[str, Any]:
    """按 skill_results 文件顺序复算 result_chain_head。"""
    previous_head: str | None = None
    result_index: dict[str, dict[str, Any]] = {}
    for result_rel_path in state.get("skill_results", []):
        try:
            result, file_hash = load_skill_result_file(result_rel_path, run_dir)
        except ValueError as exc:
            message = str(exc)
            if "不存在" in message:
                failure_kind = "missing"
            elif "合法 JSON" in message:
                failure_kind = "invalid_json"
            elif "run-relative" in message or "逃逸" in message:
                failure_kind = "path_invalid"
            else:
                failure_kind = "invalid_schema"
            return {
                "valid": False,
                "reason": f"result file invalid: {result_rel_path}: {exc}",
                "computed_head": previous_head,
                "result_index": result_index,
                "failure_kind": failure_kind,
                "failure_path": result_rel_path,
                "failure_order": len(result_index) + 1,
            }
        previous_head = compute_result_chain_head(previous_head, result, file_hash)
        result_index[result["result_id"]] = result
    return {
        "valid": previous_head == state.get("result_chain_head"),
        "reason": "ok" if previous_head == state.get("result_chain_head") else "result_chain_head mismatch",
        "computed_head": previous_head,
        "result_index": result_index,
        "failure_kind": None if previous_head == state.get("result_chain_head") else "chain_mismatch",
    }


def materialize_synthetic_failure_for_replay_error(
    *,
    state: dict[str, Any],
    run_dir: Path,
    replay: dict[str, Any],
) -> dict[str, Any]:
    """为缺失或不可解析 result 写出 synthetic failure 并返回更新后的 state。"""
    failure_path = replay.get("failure_path") or "logs/skill_results/unknown_missing.result.json"
    next_action = state.get("next_action", {}) if isinstance(state.get("next_action"), dict) else {}
    synthetic = build_synthetic_failure(
        run_id=state["run_id"],
        order=int(replay.get("failure_order") or len(state.get("skill_results", [])) + 1),
        target_stage=next_action.get("stage") or state.get("stage") or "init",
        target_skill_name=next_action.get("skill_name") or "unknown-skill",
        missing_result_path=str(failure_path),
        error_code=FH_SKILL_RESULT_MISSING if replay.get("failure_kind") == "missing" else FH_SKILL_RESULT_INVALID,
    )
    write_info = write_skill_result_atomic(
        run_dir,
        synthetic,
        result_rel_path=f"logs/skill_results/{synthetic['order']:03d}_synthetic_failure.result.json",
    )
    result_rel = write_info["path"]
    base_state = deepcopy(state)
    failure_path = replay.get("failure_path")
    base_state["skill_results"] = [
        path for path in base_state.get("skill_results", []) if path != failure_path
    ]
    base_state["result_chain_head"] = replay.get("computed_head")
    return apply_result_to_state(
        base_state,
        synthetic,
        result_rel,
        run_dir=run_dir,
        enforce_planned_key=False,
    )


def decide_resume_action(
    state: dict[str, Any],
    discovered_last_result_id: str | None = None,
    run_dir: Path | None = None,
) -> dict[str, Any]:
    """根据 state 判定恢复动作。"""
    replay = None
    if run_dir is not None:
        replay = replay_result_chain(run_dir, state)
        if not replay["valid"]:
            next_state = None
            synthetic_result_path = None
            if replay.get("failure_kind") in {"missing", "invalid_json", "invalid_schema", "path_invalid"}:
                next_state = materialize_synthetic_failure_for_replay_error(
                    state=state,
                    run_dir=run_dir,
                    replay=replay,
                )
                synthetic_result_path = next_state["skill_results"][-1]
                atomic_write_state(next_state, run_dir / "logs" / "state.yaml", run_dir=run_dir)
            return {
                "resume_allowed": False,
                "next_action_kind": "manual_recover",
                "reason": replay["reason"],
                "resume_from_result_id": state.get("applied_result_id"),
                "computed_head": replay["computed_head"],
                "synthetic_failure_path": synthetic_result_path,
                "next_state": next_state,
            }

    last_result_id = discovered_last_result_id or state.get("last_result_id")
    applied_result_id = state.get("applied_result_id")
    if last_result_id != applied_result_id:
        return {
            "resume_allowed": False,
            "next_action_kind": "manual_recover",
            "reason": "last_result_id 与 applied_result_id 不一致，只能从已应用结果继续或生成 synthetic failure",
            "resume_from_result_id": applied_result_id,
        }
    next_action = state.get("next_action", {})
    result_index = replay["result_index"] if replay else None
    contract_errors = validate_next_action_contract(
        state=state,
        result_index=result_index,
        artifact_index=collect_artifact_ids(state, result_index),
    )
    if contract_errors:
        return {
            "resume_allowed": False,
            "next_action_kind": "manual_recover",
            "reason": "; ".join(contract_errors),
            "resume_from_result_id": applied_result_id,
        }
    return {
        "resume_allowed": next_action.get("kind") not in {"stop", None},
        "next_action_kind": next_action.get("kind"),
        "reason": next_action.get("reason", "从 run-state.next_action 恢复"),
        "resume_from_result_id": applied_result_id,
    }


def atomic_write_state(state: dict[str, Any], state_path: Path, *, run_dir: Path | None = None) -> None:
    """校验并原子写入 state.yaml。

    为避免引入 YAML 依赖，这里写入 JSON-compatible YAML；YAML 解析器可直接读取。
    """
    state_to_write = deepcopy(state)
    if state_to_write.get("schema_id") == "state":
        state_to_write["schema_id"] = "run-state"
    validation = validate_run_state(state_to_write)
    if not validation.valid:
        raise ValueError(f"run-state 未通过校验：{validation.errors}")
    resolved_run_dir = run_dir or state_path.parent.parent
    result_index = None
    if state_to_write.get("skill_results"):
        replay = replay_result_chain(resolved_run_dir, state_to_write)
        if not replay["valid"]:
            raise ValueError(f"run-state result_chain_head 未通过复算：{replay['reason']}")
        result_index = replay["result_index"]
    contract_errors = validate_next_action_contract(
        state=state_to_write,
        result_index=result_index,
        artifact_index=collect_artifact_ids(state_to_write, result_index),
    )
    if contract_errors:
        raise ValueError(f"run-state next_action 未通过恢复协议校验：{contract_errors}")
    state_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = state_path.with_suffix(state_path.suffix + ".tmp")
    temp_path.write_text(json.dumps(state_to_write, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temp_path, state_path)
