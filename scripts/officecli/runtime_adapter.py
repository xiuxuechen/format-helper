#!/usr/bin/env python3
"""OfficeCLI runtime adapter — 执行 batch、管理 checkpoint、结果对齐。

不在此模块中执行语义判断或风险升级。所有操作均按 execution request 的
确定性指引执行。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from jsonschema import Draft202012Validator, FormatChecker

from scripts.officecli.contracts import RAW_SET_ACTIONS

UTC = timezone.utc
ROOT = Path(__file__).resolve().parents[2]

# §14 / §21.8 错误码
FH_OFFICECLI_TIMEOUT = "FH-OFFICECLI-TIMEOUT"
FH_OFFICECLI_NONJSON_OUTPUT = "FH-OFFICECLI-NONJSON-OUTPUT"
FH_OFFICECLI_RESULT_MISMATCH = "FH-OFFICECLI-RESULT-MISMATCH"
FH_OFFICECLI_IDEMPOTENCY_CONFLICT = "FH-OFFICECLI-IDEMPOTENCY-CONFLICT"
DFR_OFFICECLI_BATCH_FAILED = "DFR-OFFICECLI-BATCH-FAILED"
DFR_OFFICECLI_L3_NOT_AUTHORIZED = "DFR-OFFICECLI-L3-NOT-AUTHORIZED"
FH_OFFICECLI_REQUEST_INVALID = "FH-OFFICECLI-REQUEST-INVALID"
SINGLE_NODE_XPATH_V1 = re.compile(
    r"^/[A-Za-z_][A-Za-z0-9_.-]*:[A-Za-z_][A-Za-z0-9_.-]*\[1\]"
    r"(?:/[A-Za-z_][A-Za-z0-9_.-]*:[A-Za-z_][A-Za-z0-9_.-]*\[[1-9][0-9]*\])*$"
)
REQUEST_SCHEMA_PATH = ROOT / "contracts" / "officecli" / "schemas" / "officecli-execution-request.schema.json"

# §21.8 不可重试错误码
NON_RETRYABLE_CODES = {
    FH_OFFICECLI_NONJSON_OUTPUT,
    FH_OFFICECLI_RESULT_MISMATCH,
    FH_OFFICECLI_IDEMPOTENCY_CONFLICT,
    "DFR-OFFICECLI-L3-NOT-AUTHORIZED",
    "DFR-OFFICECLI-VALIDATE-FAILED",
    "DFR-OFFICECLI-POSTCONDITION-FAILED",
}


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def utc_now() -> str:
    return datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def is_retryable(code: str) -> bool:
    """§21.8: 纯函数判定 retryable。"""
    if code in NON_RETRYABLE_CODES:
        return False
    return code in {FH_OFFICECLI_TIMEOUT, DFR_OFFICECLI_BATCH_FAILED}


def parse_single_json_stdout(stdout: str) -> Any:
    text = stdout.strip()
    if not text:
        raise ValueError("stdout empty")
    decoder = json.JSONDecoder()
    value, index = decoder.raw_decode(text)
    if text[index:].strip():
        raise ValueError("trailing garbage in stdout")
    return value


def run_officecli(executable: Path, args: list[str], timeout_seconds: int = 120) -> subprocess.CompletedProcess:
    env = os.environ.copy()
    env["OFFICECLI_SKIP_UPDATE"] = "1"
    env["OFFICECLI_NO_AUTO_RESIDENT"] = "1"
    return subprocess.run(
        [str(executable), *args],
        text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        timeout=timeout_seconds, check=False, env=env,
    )


def align_operation_results(
    native_results: list[dict[str, Any]],
    request_operations: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """按 index 对齐 OfficeCLI 返回结果与 request operations。

    检测：重复索引、越界索引、缺口（首个失败前）、命令不一致 → RESULT_MISMATCH。
    """
    indexed: dict[int, dict[str, Any]] = {}
    for nr in native_results:
        idx = nr.get("index")
        if not isinstance(idx, int):
            raise ValueError(f"{FH_OFFICECLI_RESULT_MISMATCH}: native result missing index")
        if idx in indexed:
            raise ValueError(f"{FH_OFFICECLI_RESULT_MISMATCH}: duplicate index {idx}")
        indexed[idx] = nr

    results: list[dict[str, Any]] = []
    first_failure_seen = False
    max_native_idx = max(indexed.keys()) if indexed else -1

    for op_idx, op in enumerate(request_operations):
        nr = indexed.get(op_idx)
        if nr is None:
            if first_failure_seen or op_idx > max_native_idx:
                results.append({
                    "operation_id": op["operation_id"],
                    "source_action_id": op.get("source_action_id", ""),
                    "index": op_idx,
                    "status": "not_run",
                    "native_success": None, "native_output": None,
                    "native_error": None,
                    "before_target_fingerprint": None,
                    "after_target_fingerprint": None,
                    "postconditions_passed": False,
                    "duration_ms": 0,
                })
            else:
                raise ValueError(
                    f"{FH_OFFICECLI_RESULT_MISMATCH}: missing index {op_idx} before first failure"
                )
        else:
            native_success = bool(nr.get("success", False))
            if not native_success and not first_failure_seen:
                first_failure_seen = True
            # §11.1: 校验 native result 命令一致性
            native_cmd = nr.get("command") or nr.get("operation")
            if native_cmd and native_cmd != op.get("command"):
                raise ValueError(
                    f"{FH_OFFICECLI_RESULT_MISMATCH}: command mismatch at index {op_idx}"
                )
            results.append({
                "operation_id": op["operation_id"],
                "source_action_id": op.get("source_action_id", ""),
                "index": op_idx,
                "status": "executed" if native_success else "failed",
                "native_success": native_success,
                "native_output": nr.get("output"),
                "native_error": nr.get("error"),
                "before_target_fingerprint": op.get("target_binding", {}).get("fingerprint") if isinstance(op.get("target_binding"), dict) else None,
                "after_target_fingerprint": None,
                "postconditions_passed": native_success,
                "duration_ms": 0,
            })

    for idx in indexed:
        if idx >= len(request_operations):
            raise ValueError(f"{FH_OFFICECLI_RESULT_MISMATCH}: index {idx} out of range")
    return results


def execute_batch(
    executable: Path,
    working_docx: Path,
    batch: dict[str, Any],
    checkpoint_dir: Path,
    artifact_dir: Path,
    run_dir: Path,
    timeout_seconds: int = 120,
    checkpoint_callback: Any | None = None,
) -> dict[str, Any]:
    """执行单个 batch：checkpoint → OfficeCLI batch → 对齐结果。"""
    batch_id = batch.get("batch_id", "unknown")
    seq = batch.get("sequence", 0)
    started = utc_now()
    start_ms = int(time.time() * 1000)

    # checkpoint
    checkpoint_dir.mkdir(parents=True, exist_ok=True)
    checkpoint_path = checkpoint_dir / f"{batch_id}.before.docx"
    shutil.copy2(working_docx, checkpoint_path)
    working_before_hash = sha256_file(working_docx)
    if checkpoint_callback is not None:
        checkpoint_callback(batch, checkpoint_path)
    artifact_dir.mkdir(parents=True, exist_ok=True)
    stdout_path = artifact_dir / f"stdout-{batch_id}.txt"
    stderr_path = artifact_dir / f"stderr-{batch_id}.txt"

    # native batch JSON
    native_ref = batch.get("officecli_batch_ref", {})
    native_path = _resolve_request_artifact(run_dir, native_ref) or Path("")

    args = ["batch", str(working_docx), "--input", str(native_path), "--stop-on-error", "--json"]

    try:
        proc = run_officecli(executable, args, timeout_seconds)
    except subprocess.TimeoutExpired:
        stdout_path.write_text("", encoding="utf-8")
        stderr_path.write_text("batch timeout", encoding="utf-8")
        shutil.copy2(checkpoint_path, working_docx)
        return _failed_batch_result(batch_id, seq, started, working_before_hash,
                                    FH_OFFICECLI_TIMEOUT, "batch timeout", None, None,
                                    docx_path=working_docx, run_dir=run_dir,
                                    stdout_path=stdout_path, stderr_path=stderr_path)

    end_ms = int(time.time() * 1000)
    finished = utc_now()

    stdout_path.write_text(proc.stdout or "", encoding="utf-8")
    stderr_path.write_text(proc.stderr or "", encoding="utf-8")

    if proc.returncode != 0:
        shutil.copy2(checkpoint_path, working_docx)
        return _failed_batch_result(batch_id, seq, started, working_before_hash,
                                    DFR_OFFICECLI_BATCH_FAILED, f"exit code {proc.returncode}",
                                    proc.returncode, proc.stderr, docx_path=working_docx,
                                    run_dir=run_dir, stdout_path=stdout_path, stderr_path=stderr_path)

    try:
        parsed = parse_single_json_stdout(proc.stdout)
    except ValueError:
        shutil.copy2(checkpoint_path, working_docx)
        return _failed_batch_result(batch_id, seq, started, working_before_hash,
                                    FH_OFFICECLI_NONJSON_OUTPUT, "stdout not valid JSON",
                                    proc.returncode, proc.stderr, docx_path=working_docx,
                                    run_dir=run_dir, stdout_path=stdout_path, stderr_path=stderr_path)

    # OfficeCLI envelope: {success, data: {results: [...]}}
    envelope_success = bool(parsed.get("success", False)) if isinstance(parsed, dict) else False
    data = parsed.get("data", {}) if isinstance(parsed, dict) else {}
    native_results = data.get("results", []) if isinstance(data, dict) else []

    try:
        op_results = align_operation_results(native_results, batch.get("operations", []))
    except ValueError as exc:
        shutil.copy2(checkpoint_path, working_docx)
        return _failed_batch_result(batch_id, seq, started, working_before_hash,
                                    FH_OFFICECLI_RESULT_MISMATCH, str(exc),
                                    proc.returncode, proc.stderr, docx_path=working_docx,
                                    run_dir=run_dir, stdout_path=stdout_path, stderr_path=stderr_path)

    business_success = envelope_success and all(item.get("status") == "executed" for item in op_results)
    if not business_success:
        shutil.copy2(checkpoint_path, working_docx)

    working_after_hash = sha256_file(working_docx) if working_docx.exists() else ""
    duration = end_ms - start_ms

    payload = {
        "batch_id": batch_id,
        "sequence": seq,
        "started_at": started,
        "finished_at": finished,
        "duration_ms": duration,
        "exit_code": proc.returncode,
        "native_success": business_success,
        "status": "done" if business_success else "failed",
        "stdout_artifact_ref": _artifact_ref_for_file(run_dir, stdout_path, f"stdout-{batch_id}", "log", None, None),
        "stderr_artifact_ref": _artifact_ref_for_file(run_dir, stderr_path, f"stderr-{batch_id}", "log", None, None),
        "working_before_sha256": working_before_hash,
        "working_after_sha256": working_after_hash,
        "operation_results": op_results,
    }
    if not business_success:
        payload.update({
            "_error_code": DFR_OFFICECLI_BATCH_FAILED,
            "_error_message": "OfficeCLI batch business result failed",
            "_native_stderr": proc.stderr,
        })
    return payload


def _failed_batch_result(
    batch_id: str, seq: int, started: str, before_hash: str,
    code: str, message: str, exit_code: int | None, stderr: str | None,
    docx_path: Path | None = None, run_dir: Path | None = None,
    stdout_path: Path | None = None, stderr_path: Path | None = None,
) -> dict[str, Any]:
    """构造失败 batch result，所有操作标记为 not_run。"""
    finished = utc_now()
    after_hash = "0" * 64
    if docx_path and docx_path.exists():
        try:
            after_hash = sha256_file(docx_path)
        except OSError:
            pass
    payload = {
        "batch_id": batch_id,
        "sequence": seq,
        "started_at": started,
        "finished_at": finished,
        "duration_ms": 0,
        "exit_code": -1 if exit_code is None else exit_code,
        "native_success": False,
        "status": "failed",
        "working_before_sha256": before_hash,
        "working_after_sha256": after_hash,
        "operation_results": [],  # caller fills with not_run
        "stdout_artifact_ref": _artifact_ref_for_file(run_dir, stdout_path, f"stdout-{batch_id}", "log", None, None) if run_dir and stdout_path else None,
        "stderr_artifact_ref": _artifact_ref_for_file(run_dir, stderr_path, f"stderr-{batch_id}", "log", None, None) if run_dir and stderr_path else None,
        "_error_code": code,
        "_error_message": message,
        "_native_stderr": stderr,
    }
    return payload


def _canonical_sha256(value: Any) -> str:
    return sha256_bytes(json.dumps(
        value, ensure_ascii=False, sort_keys=True, separators=(",", ":"),
    ).encode("utf-8"))


def _validate_ref_integrity(run_dir: Path, ref: dict[str, Any], field: str) -> tuple[Path | None, list[str]]:
    errors: list[str] = []
    if not isinstance(ref, dict):
        return None, [f"{field}: ArtifactRef 缺失"]
    path = _resolve_request_artifact(run_dir, ref)
    if path is None or not path.is_file():
        return path, [f"{field}: 引用文件不存在或路径非法"]
    if sha256_file(path) != ref.get("sha256") or path.stat().st_size != ref.get("size_bytes"):
        errors.append(f"{field}: hash/size 不匹配")
    return path, errors


def _strip_runtime_batch_fields(batch: dict[str, Any]) -> dict[str, Any]:
    """移除 request 执行时补充的 batch 引用字段，用于与 plan 重建结果比对。"""
    payload = json.loads(json.dumps(batch, ensure_ascii=False))
    payload.pop("officecli_batch_ref", None)
    payload.pop("checkpoint_ref", None)
    return payload


def _validate_request_matches_finalized_plan(
    run_dir: Path, request: dict[str, Any],
) -> list[str]:
    errors: list[str] = []
    plan_path, ref_errors = _validate_ref_integrity(run_dir, request.get("plan_ref"), "plan_ref")
    errors.extend(f"{FH_OFFICECLI_REQUEST_INVALID}: {item}" for item in ref_errors)
    if plan_path is None or ref_errors:
        return errors
    try:
        from scripts.officecli.request_builder import (
            load_repair_plan, plan_to_batches, validate_finalized_plan_for_request,
        )
        plan = load_repair_plan(plan_path)
        snapshot_path = _resolve_request_artifact(run_dir, request.get("snapshot_ref") or {}) or Path()
        capability_path = _resolve_request_artifact(run_dir, request.get("capability_manifest_ref") or {}) or Path()
        plan_errors = validate_finalized_plan_for_request(
            plan,
            plan_path=plan_path,
            run_id=str(request.get("run_id", "")),
            plan_revision=str(request.get("plan_revision", "")),
            snapshot_path=snapshot_path,
            capability_manifest_path=capability_path,
        )
        errors.extend(f"{FH_OFFICECLI_REQUEST_INVALID}: {item}" for item in plan_errors)
        expected_batches = plan_to_batches(
            plan.get("actions", []),
            str(request.get("plan_sha256", "")),
            str((request.get("working_docx_before_ref") or {}).get("sha256", "")),
        )
        actual_batches = [
            _strip_runtime_batch_fields(batch)
            for batch in request.get("batches", [])
        ]
        if expected_batches != actual_batches:
            errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: request batches 与 finalized repair-plan 重建结果不一致")
        if request.get("plan_sha256") != sha256_file(plan_path):
            errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: plan_sha256 与 plan_ref 文件不一致")
        if plan.get("snapshot_ref") != request.get("snapshot_ref"):
            errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: snapshot_ref 与 finalized repair-plan 不一致")
        if plan.get("capability_manifest_ref") != request.get("capability_manifest_ref"):
            errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: capability_manifest_ref 与 finalized repair-plan 不一致")
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: finalized repair-plan 不可校验: {exc}")
    return errors


def _expected_raw_confirmation(raw: dict[str, Any]) -> dict[str, Any]:
    return {
        "part": raw.get("part"),
        "xpath": raw.get("xpath"),
        "action": raw.get("action"),
        "xml_sha256": raw.get("xml_sha256"),
        "expected_match_count": raw.get("expected_match_count"),
        "precondition_raw_sha256": raw.get("precondition_raw_sha256"),
    }


def _validate_l3_operation(
    executable: Path, working_docx: Path, run_dir: Path,
    request: dict[str, Any], operation: dict[str, Any],
    *, validate_live_raw: bool = True,
) -> list[str]:
    errors: list[str] = []
    raw = operation.get("raw")
    if operation.get("command") != "raw-set" or not isinstance(raw, dict):
        return [f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: L3_WRITE 必须为 raw-set"]
    xpath = str(raw.get("xpath", ""))
    if SINGLE_NODE_XPATH_V1.fullmatch(xpath) is None:
        errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: XPath 不符合 single_node_xpath_v1")
    xml = str(raw.get("xml", ""))
    if sha256_bytes(xml.encode("utf-8")) != raw.get("xml_sha256"):
        errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: XML hash 不匹配")
    if raw.get("expected_match_count") != 1:
        errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: expected_match_count 必须为 1")
    if raw.get("action") not in RAW_SET_ACTIONS:
        errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: raw action 不在 OfficeCLI v1.0.113 canonical 白名单")

    confirmation_ref = operation.get("manual_confirmation_ref")
    confirmation_path, confirmation_errors = _validate_ref_integrity(
        run_dir, confirmation_ref, "manual_confirmation_ref",
    )
    errors.extend(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: {item}" for item in confirmation_errors)
    if confirmation_path and not confirmation_errors:
        try:
            confirmation = json.loads(confirmation_path.read_text(encoding="utf-8"))
            review_id = raw.get("manual_review_id")
            item = next(
                (candidate for candidate in confirmation.get("items", [])
                 if candidate.get("review_id") == review_id),
                None,
            )
            decision = item.get("decision", {}) if isinstance(item, dict) else {}
            selected_action = decision.get("selected_action")
            if (
                decision.get("status") not in {"approved", "modified"}
                or decision.get("allows_continue") is not True
                or decision.get("selected_action_sha256") != raw.get("decision_snapshot_sha256")
            ):
                errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: 人工确认未授权或决策 hash 不匹配")
            if not isinstance(selected_action, dict):
                errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: 人工确认缺少 selected_action")
            else:
                from scripts.validation.manual_review_repair import compute_selected_action_sha256
                if compute_selected_action_sha256(selected_action) != decision.get("selected_action_sha256"):
                    errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: selected_action canonical hash 不匹配")
                if selected_action.get("review_id") != review_id:
                    errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: selected_action.review_id 不匹配")
                if selected_action.get("operation") != "raw-set":
                    errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: selected_action.operation 必须为 raw-set")
                parameters = selected_action.get("parameters")
                if not isinstance(parameters, dict) or parameters.get("officecli_raw") != _expected_raw_confirmation(raw):
                    errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: selected_action 未绑定 raw-set payload")
        except (OSError, json.JSONDecodeError):
            errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: 人工确认文件不可解析")

    snapshot_path, snapshot_errors = _validate_ref_integrity(
        run_dir, request.get("snapshot_ref"), "snapshot_ref",
    )
    errors.extend(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: {item}" for item in snapshot_errors)
    part_name = str(raw.get("part", "")).lstrip("/")
    if snapshot_path and not snapshot_errors:
        try:
            snapshot = json.loads(snapshot_path.read_text(encoding="utf-8"))
            part = next(
                (candidate for candidate in snapshot.get("parts", [])
                 if str(candidate.get("part_name", "")).lstrip("/") == part_name),
                None,
            )
            if not part or part.get("sha256") != raw.get("precondition_raw_sha256"):
                errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: before snapshot raw hash 不匹配")
        except (OSError, json.JSONDecodeError):
            errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: before snapshot 不可解析")

    if validate_live_raw and not errors:
        proc = run_officecli(executable, ["raw", str(working_docx), f"/{part_name}"], 120)
        if proc.returncode != 0 or sha256_bytes((proc.stdout or "").encode("utf-8")) != raw.get("precondition_raw_sha256"):
            errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: 执行前 current raw hash 不匹配")
    return errors


def validate_execution_request_preflight(
    executable: Path, request: dict[str, Any], run_dir: Path,
    working_state_ref: dict[str, Any] | None = None,
    completed_batch_ids: set[str] | None = None,
) -> list[str]:
    """执行前统一校验 Schema、hash、Gate、引用和 L3 授权。"""
    schema = json.loads(REQUEST_SCHEMA_PATH.read_text(encoding="utf-8"))
    validator = Draft202012Validator(schema, format_checker=FormatChecker())
    errors = [
        f"{FH_OFFICECLI_REQUEST_INVALID}: {error.json_path}: {error.message}"
        for error in sorted(validator.iter_errors(request), key=lambda item: item.json_path)
    ]
    expected_request_sha = _canonical_sha256({
        key: value for key, value in request.items() if key != "request_sha256"
    })
    if request.get("request_sha256") != expected_request_sha:
        errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: request_sha256 不匹配")
    if request.get("gate_check", {}).get("status") != "passed":
        errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: request Gate 未通过")
    if request.get("runtime_id") not in {
        "win-x64", "win-arm64", "linux-x64-gnu", "linux-arm64-gnu",
        "linux-x64-musl", "linux-arm64-musl", "osx-x64", "osx-arm64",
    }:
        errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: runtime_id 非法")
    errors.extend(_validate_request_matches_finalized_plan(run_dir, request))

    for field in (
        "plan_ref", "snapshot_ref",
        "lock_ref", "capability_manifest_ref", "officecli_executable_ref",
    ):
        _path, ref_errors = _validate_ref_integrity(run_dir, request.get(field), field)
        errors.extend(f"{FH_OFFICECLI_REQUEST_INVALID}: {item}" for item in ref_errors)
    working_ref = request.get("working_docx_before_ref") or {}
    working_docx = _resolve_request_artifact(run_dir, working_ref)
    expected_working = working_state_ref or working_ref
    if working_docx is None or not working_docx.is_file():
        errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: working_docx_before_ref: 引用文件不存在或路径非法")
    elif (
        sha256_file(working_docx) != expected_working.get("sha256")
        or working_docx.stat().st_size != expected_working.get("size_bytes")
    ):
        errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: working_docx_before_ref: hash/size 不匹配")
    executable_ref = request.get("officecli_executable_ref") or {}
    if executable.is_file() and sha256_file(executable) != executable_ref.get("sha256"):
        errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: executable hash 不匹配")

    for batch in request.get("batches", []):
        native_path, native_errors = _validate_ref_integrity(
            run_dir, batch.get("officecli_batch_ref"), f"{batch.get('batch_id')}.officecli_batch_ref",
        )
        errors.extend(f"{FH_OFFICECLI_REQUEST_INVALID}: {item}" for item in native_errors)
        if native_path and not native_errors:
            try:
                from scripts.officecli.request_builder import build_native_batch_item
                native_payload = json.loads(native_path.read_text(encoding="utf-8"))
                expected_native = [
                    build_native_batch_item(operation)
                    for operation in batch.get("operations", [])
                ]
                if native_payload != expected_native:
                    errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: governance/native batch 不一致")
            except (OSError, json.JSONDecodeError, ValueError):
                errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: native batch 不可解析")
        if batch.get("checkpoint_ref") is not None:
            errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: checkpoint_ref 执行前必须为 null")
        l3_operations = [
            operation for operation in batch.get("operations", [])
            if operation.get("risk_class") == "L3_WRITE" or operation.get("command") == "raw-set"
        ]
        if l3_operations and (len(batch.get("operations", [])) != 1 or len(l3_operations) != 1):
            errors.append(f"{DFR_OFFICECLI_L3_NOT_AUTHORIZED}: L3_WRITE 必须独占 batch")
        if l3_operations and working_docx:
            errors.extend(_validate_l3_operation(
                executable, working_docx, run_dir, request, l3_operations[0],
                validate_live_raw=batch.get("batch_id") not in (completed_batch_ids or set()),
            ))
        for predicate in batch.get("preconditions", []):
            if predicate.get("type") in {"artifact_exists", "hash_equals"}:
                target_id = predicate.get("target_ref")
                refs = [
                    request.get(field) for field in (
                        "plan_ref", "working_docx_before_ref", "snapshot_ref",
                        "lock_ref", "capability_manifest_ref", "officecli_executable_ref",
                    )
                ]
                refs.extend(
                    candidate.get("officecli_batch_ref")
                    for candidate in request.get("batches", [])
                )
                target = next((ref for ref in refs if isinstance(ref, dict) and ref.get("artifact_id") == target_id), None)
                _path, predicate_errors = _validate_ref_integrity(run_dir, target, predicate.get("predicate_id", "predicate"))
                errors.extend(f"{FH_OFFICECLI_REQUEST_INVALID}: {item}" for item in predicate_errors)
                if predicate.get("type") == "hash_equals" and isinstance(target, dict) and target.get("sha256") != predicate.get("expected"):
                    errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: {predicate.get('predicate_id')} hash_equals 不满足")
            else:
                errors.append(f"{FH_OFFICECLI_REQUEST_INVALID}: 不支持的执行前 predicate {predicate.get('type')}")
    return errors


def execute_request(
    executable: Path,
    request: dict[str, Any],
    run_dir: Path,
    timeout_seconds: int = 120,
    request_path: Path | None = None,
    attempt_no: int = 1,
) -> dict[str, Any]:
    """runtime adapter 主入口：执行完整 execution request。"""
    started = utc_now()
    start_ms = int(time.time() * 1000)
    working_docx_path = _resolve_request_artifact(run_dir, request.get("working_docx_before_ref", {})) or Path("")

    batches = request.get("batches", [])
    batch_results: list[dict[str, Any]] = []
    overall_status = "done"
    failed_batch_id = None
    failed_operation_id = None
    retryable = False
    error: dict[str, Any] | None = None

    checkpoint_dir = run_dir / "output" / "_internal" / "checkpoints"
    artifact_dir = run_dir / "output" / "_internal" / "officecli"

    if request_path is not None and request_path.is_file():
        progress_request_ref = _artifact_ref_for_file(
            run_dir, request_path, request.get("request_id", request_path.stem),
            "request", "officecli-execution-request", "2.0.0",
        )
    else:
        progress_request_ref = None

    completed_batch_ids: list[str] = []

    def persist_progress(batch: dict[str, Any], checkpoint_path: Path) -> None:
        if progress_request_ref is None:
            return
        _write_execution_in_progress(
            run_dir, request, progress_request_ref, batch, checkpoint_path,
            completed_batch_ids,
        )

    for batch in batches:
        result = execute_batch(
            executable, working_docx_path, batch,
            checkpoint_dir, artifact_dir, run_dir, timeout_seconds,
            persist_progress,
        )
        # 确保操作结果挂载
        if not result.get("operation_results"):
            ops = batch.get("operations", [])
            result["operation_results"] = [
                {"operation_id": op.get("operation_id", "?"),
                 "source_action_id": op.get("source_action_id", ""),
                 "index": i, "status": "not_run",
                 "native_success": None, "native_output": None,
                 "native_error": None,
                 "before_target_fingerprint": None,
                 "after_target_fingerprint": None,
                 "postconditions_passed": False,
                 "duration_ms": 0}
                for i, op in enumerate(ops)
            ]
        # 读取内部字段（在 pop 前）
        saved_error_code = result.get("_error_code", DFR_OFFICECLI_BATCH_FAILED)
        saved_error_message = result.get("_error_message", "batch failed")
        saved_native_stderr = result.get("_native_stderr")

        # 从 batch result 中移除内部字段
        result.pop("_error_code", None)
        result.pop("_error_message", None)
        result.pop("_native_stderr", None)
        batch_results.append(result)
        if result["status"] == "done":
            completed_batch_ids.append(result["batch_id"])
            completed_checkpoint = checkpoint_dir / f"{result['batch_id']}.after.docx"
            shutil.copy2(working_docx_path, completed_checkpoint)
            persist_progress(batch, completed_checkpoint)

        if result["status"] == "failed":
            overall_status = "failed"
            failed_batch_id = result["batch_id"]
            op_results = result.get("operation_results", [])
            for opr in op_results:
                if opr.get("status") == "failed":
                    failed_operation_id = opr.get("operation_id")
                    break
            error_code = saved_error_code
            retryable = is_retryable(error_code)
            error = {
                "code": error_code,
                "reason_code": error_code,
                "message": saved_error_message,
                "stage": f"batch-{result['batch_id']}",
                "retryable": retryable,
                "failed_artifact_ref": None,
                "native_exit_code": result.get("exit_code"),
                "native_error": saved_native_stderr,
                "stderr_artifact_ref": None,
            }
            break

    end_ms = int(time.time() * 1000)
    finished = utc_now()

    if request_path is not None and request_path.is_file():
        request_ref = _artifact_ref_for_file(
            run_dir, request_path, request.get("request_id", request_path.stem),
            "request", "officecli-execution-request", "2.0.0",
        )
    else:
        request_bytes = json.dumps(
            request, ensure_ascii=False, sort_keys=True, separators=(",", ":"),
        ).encode("utf-8")
        request_ref = {
            "artifact_id": request.get("request_id", "?"),
            "kind": "request",
            "relative_path": f"plans/officecli-execution-request.r{request.get('plan_revision', '?')}.json",
            "sha256": sha256_bytes(request_bytes),
            "size_bytes": len(request_bytes),
            "schema_id": "officecli-execution-request",
            "schema_version": "2.0.0",
        }
    working_after_ref = None
    if overall_status == "done" and working_docx_path.is_file():
        working_after_ref = _artifact_ref_for_file(
            run_dir, working_docx_path,
            f"DOCX-{request.get('run_id', '?')}-AFTER", "docx", None, None,
        )

    payload = {
        "schema_id": "officecli-execution-result",
        "schema_version": "2.0.0",
        "result_id": f"RES-{request.get('run_id', '?')}-{request.get('plan_revision', '?')}-A{attempt_no:03d}",
        "run_id": request.get("run_id", ""),
        "request_ref": request_ref,
        "created_at": utc_now(),
        "extensions": {},
        "officecli_version": "1.0.113",
        "runtime_id": request.get("runtime_id", ""),
        "executable_sha256": sha256_file(executable) if executable.exists() else "",
        "started_at": started,
        "finished_at": finished,
        "duration_ms": end_ms - start_ms,
        "status": overall_status,
        "working_docx_before_ref": request.get("working_docx_before_ref"),
        "working_docx_after_ref": working_after_ref,
        "batch_results": batch_results,
        "failed_batch_id": failed_batch_id,
        "failed_operation_id": failed_operation_id,
        "retryable": retryable,
        "error": error if error else {
            "code": "NONE", "reason_code": "NONE", "message": "",
            "stage": "done", "retryable": False,
            "failed_artifact_ref": None, "native_exit_code": None,
            "native_error": None, "stderr_artifact_ref": None,
        },
        "stdout_artifacts": [item["stdout_artifact_ref"] for item in batch_results],
        "stderr_artifacts": [item["stderr_artifact_ref"] for item in batch_results],
        "gate_check": {
            "gate_id": "officecli-execution-result-v5",
            "status": "passed" if overall_status == "done" else "failed",
            "checked_at": finished,
            "predicate_version": "1.0.0",
            "evidence_refs": [
                ref["artifact_id"]
                for item in batch_results
                for ref in (item["stdout_artifact_ref"], item["stderr_artifact_ref"])
            ],
            "failed_codes": [error["code"]] if error else [],
        },
    }
    return payload


def _resume_handler(run_dir: Path, executable: Path) -> int:
    """resume — 校验 lock/capability/plan hash 一致性后从 checkpoint 恢复。"""
    repair_log_path = run_dir / "logs" / "repair_execution_log.json"
    if not repair_log_path.exists():
        sys.stdout.write(json.dumps({"ok": False, "error": "FH-OFFICECLI-RESUME-INTEGRITY-FAILED",
                                      "message": "No execution log found"}) + "\n")
        return 2
    log = json.loads(repair_log_path.read_text(encoding="utf-8"))
    status = log.get("current_status", "")
    resume_allowed = {"execution_in_progress", "execution_failed_retryable", "executed_ready"}
    if status not in resume_allowed:
        sys.stdout.write(json.dumps({"ok": False, "status": status, "action": "manual_recover"}) + "\n")
        return 1
    if status == "executed_ready":
        sys.stdout.write(json.dumps({
            "ok": True, "status": status, "action": "post_write_qa",
            "message": "执行已完成，禁止重复 batch；继续写后 QA。",
        }, ensure_ascii=False, sort_keys=True) + "\n")
        return 0
    if status == "execution_failed_retryable" and log.get("resume_policy", {}).get("max_additional_attempts", 0) <= 0:
        sys.stdout.write(json.dumps({
            "ok": False, "status": status, "action": "manual_recover",
            "error": "FH-OFFICECLI-RESUME-ATTEMPTS-EXHAUSTED",
        }, ensure_ascii=False, sort_keys=True) + "\n")
        return 1

    request_ref = log.get("request_ref") or {}
    request_path = run_dir / str(request_ref.get("relative_path", ""))
    if not request_path.is_file() or sha256_file(request_path) != request_ref.get("sha256"):
        sys.stdout.write(json.dumps({"ok": False, "error": "FH-OFFICECLI-RESUME-INTEGRITY-FAILED"}) + "\n")
        return 2
    request = json.loads(request_path.read_text(encoding="utf-8"))
    request_core = {key: value for key, value in request.items() if key != "request_sha256"}
    request_digest = sha256_bytes(json.dumps(request_core, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8"))
    if request_digest != request.get("request_sha256") or request.get("plan_ref") != log.get("plan_ref"):
        sys.stdout.write(json.dumps({"ok": False, "error": "FH-OFFICECLI-RESUME-INTEGRITY-FAILED", "message": "request or plan reference mismatch"}) + "\n")
        return 2

    # §15.5: 复算所有执行前引用，禁止按 mtime 或环境猜测恢复。
    for key in ("plan_ref", "snapshot_ref", "lock_ref", "capability_manifest_ref"):
        ref = request.get(key) or {}
        path = _resolve_request_artifact(run_dir, ref)
        if not path or not path.is_file() or sha256_file(path) != ref.get("sha256"):
            sys.stdout.write(json.dumps({
                "ok": False, "error": "FH-OFFICECLI-RESUME-INTEGRITY-FAILED",
                "message": f"{key} missing or hash mismatch",
            }) + "\n")
            return 2
    for batch in request.get("batches", []):
        native_ref = batch.get("officecli_batch_ref") or {}
        native_path = _resolve_request_artifact(run_dir, native_ref)
        if not native_path or not native_path.is_file() or sha256_file(native_path) != native_ref.get("sha256"):
            sys.stdout.write(json.dumps({"ok": False, "error": "FH-OFFICECLI-RESUME-INTEGRITY-FAILED", "message": "native batch missing or hash mismatch"}) + "\n")
            return 2
    executable_ref = request.get("officecli_executable_ref") or {}
    if not executable.is_file() or sha256_file(executable) != executable_ref.get("sha256"):
        sys.stdout.write(json.dumps({"ok": False, "error": "FH-OFFICECLI-RESUME-INTEGRITY-FAILED", "message": "executable hash mismatch"}) + "\n")
        return 2
    version_proc = run_officecli(executable, ["--version"], 30)
    if version_proc.returncode != 0 or version_proc.stdout.strip().lstrip("v") != "1.0.113":
        sys.stdout.write(json.dumps({"ok": False, "error": "FH-OFFICECLI-RESUME-INTEGRITY-FAILED", "message": "OfficeCLI version mismatch"}) + "\n")
        return 2

    attempts = log.get("attempts") or []
    checkpoint_ref = attempts[-1].get("checkpoint_ref") if attempts else None
    if checkpoint_ref is None:
        required_refs = log.get("resume_policy", {}).get("required_artifact_refs", [])
        checkpoint_ref = next((ref for ref in reversed(required_refs) if str(ref.get("artifact_id", "")).startswith("checkpoint-")), None)
    if checkpoint_ref:
        checkpoint_path = run_dir / checkpoint_ref.get("relative_path", "")
        if not checkpoint_path.is_file() or sha256_file(checkpoint_path) != checkpoint_ref.get("sha256"):
            sys.stdout.write(json.dumps({"ok": False, "error": "FH-OFFICECLI-RESUME-INTEGRITY-FAILED", "message": "checkpoint missing or hash mismatch"}) + "\n")
            return 2
        working_path = _resolve_request_artifact(run_dir, request.get("working_docx_before_ref") or {})
        if working_path is None:
            return 2
        shutil.copy2(checkpoint_path, working_path)

    # §15.5: 跳过已成功 batch，仅从失败 batch 继续
    prev_results: list[dict[str, Any]] = []
    result_refs = log.get("result_refs") or []
    if result_refs:
        latest_result_path = run_dir / result_refs[-1].get("relative_path", "")
        if not latest_result_path.is_file() or sha256_file(latest_result_path) != result_refs[-1].get("sha256"):
            sys.stdout.write(json.dumps({
                "ok": False, "error": "FH-OFFICECLI-RESUME-INTEGRITY-FAILED",
                "message": "latest result missing or hash mismatch",
            }) + "\n")
            return 2
        prev_results = json.loads(latest_result_path.read_text(encoding="utf-8")).get("batch_results", [])
    succeeded_ids = {br.get("batch_id") for br in prev_results if br.get("status") == "done"}
    succeeded_ids.update(log.get("extensions", {}).get("completed_batch_ids", []))

    preflight_errors = validate_execution_request_preflight(
        executable, request, run_dir, working_state_ref=checkpoint_ref,
        completed_batch_ids={str(item) for item in succeeded_ids if item},
    )
    if preflight_errors:
        sys.stdout.write(json.dumps({
            "ok": False, "error": FH_OFFICECLI_REQUEST_INVALID,
            "details": preflight_errors,
        }, ensure_ascii=False) + "\n")
        return 2

    if succeeded_ids:
        remaining = [b for b in request.get("batches", []) if b.get("batch_id") not in succeeded_ids]
        request["batches"] = remaining
    result = execute_request(executable, request, run_dir, request_path=request_path, attempt_no=len(log.get("attempts") or []) + 1)
    out_path = run_dir / "logs" / "officecli-execution-result.resume.json"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(result, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n",
                        encoding="utf-8")
    _write_repair_execution_log(run_dir, request, result, out_path, existing=log)
    summary = {"ok": result["status"] == "done", "resumed": True, "status": result["status"]}
    sys.stdout.write(json.dumps(summary, ensure_ascii=False, sort_keys=True) + "\n")
    return 0 if result["status"] == "done" else 1


def _artifact_ref_for_file(run_dir: Path, path: Path, artifact_id: str, kind: str, schema_id: str | None, schema_version: str | None) -> dict[str, Any]:
    """为 run 内文件构造真实 ArtifactRef。"""
    return {
        "artifact_id": artifact_id,
        "kind": kind,
        "relative_path": str(path.resolve().relative_to(run_dir.resolve())).replace("\\", "/"),
        "sha256": sha256_file(path),
        "size_bytes": path.stat().st_size,
        "schema_id": schema_id,
        "schema_version": schema_version,
    }


def _resolve_request_artifact(run_dir: Path, ref: dict[str, Any]) -> Path | None:
    """解析 request 中历史绝对路径或 workspace/run 相对路径。"""
    raw = str(ref.get("relative_path", ""))
    if not raw:
        return None
    path = Path(raw)
    if path.is_absolute():
        return None
    candidates = [run_dir / path, ROOT / path, run_dir.parent.parent / path]
    return next((candidate for candidate in candidates if candidate.exists()), candidates[0])


def _write_execution_in_progress(
    run_dir: Path, request: dict[str, Any], request_ref: dict[str, Any],
    batch: dict[str, Any], checkpoint_path: Path,
    completed_batch_ids: list[str] | None = None,
) -> None:
    """在启动 native batch 前持久化可恢复状态与真实 checkpoint。"""
    path = run_dir / "logs" / "repair_execution_log.json"
    previous = json.loads(path.read_text(encoding="utf-8")) if path.is_file() else {}
    checkpoint_ref = _artifact_ref_for_file(
        run_dir, checkpoint_path, f"checkpoint-{batch.get('batch_id')}",
        "docx", None, None,
    )
    log = {
        "schema_id": "repair-execution-log", "schema_version": "2.0.0",
        "created_at": previous.get("created_at", utc_now()),
        "extensions": {"completed_batch_ids": list(completed_batch_ids or [])},
        "run_id": request.get("run_id"), "plan_ref": request.get("plan_ref"),
        "request_ref": request_ref, "result_refs": previous.get("result_refs", []),
        "attempts": previous.get("attempts", []),
        "current_status": "execution_in_progress",
        "resume_policy": {
            "resume_allowed": True,
            "allowed_platforms": [request.get("runtime_id")],
            "required_artifact_refs": [request.get("plan_ref"), request_ref, checkpoint_ref],
            "max_additional_attempts": int(previous.get("resume_policy", {}).get("max_additional_attempts", 1)),
            "blocked_reason_code": None,
        },
        "gate_check": {
            "gate_id": "repair-execution-log-v5", "status": "blocked",
            "checked_at": utc_now(), "predicate_version": "1.0.0",
            "evidence_refs": [checkpoint_ref["artifact_id"]],
            "failed_codes": ["FH-OFFICECLI-EXECUTION-IN-PROGRESS"],
        },
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(".json.tmp")
    temp.write_text(json.dumps(log, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n", encoding="utf-8")
    os.replace(temp, path)


def _write_repair_execution_log(
    run_dir: Path, request: dict[str, Any], result: dict[str, Any], result_path: Path,
    *, existing: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """生成或追加 repair_execution_log v2，供 resume、复核和报告共同消费。"""
    result_ref = _artifact_ref_for_file(
        run_dir, result_path, result.get("result_id", result_path.stem),
        "result", "officecli-execution-result", "2.0.0",
    )
    previous = existing or {}
    attempts = list(previous.get("attempts") or [])
    source_status = previous.get("current_status", "plan_finalized")
    if result.get("status") == "done":
        target_status = "executed_ready"
    elif result.get("retryable"):
        target_status = "execution_failed_retryable"
    else:
        target_status = "execution_failed_final"
    failed_batch_id = result.get("failed_batch_id")
    checkpoint_path = run_dir / "output" / "_internal" / "checkpoints" / f"{failed_batch_id}.before.docx" if failed_batch_id else None
    checkpoint_ref = (
        _artifact_ref_for_file(run_dir, checkpoint_path, f"checkpoint-{failed_batch_id}", "docx", None, None)
        if checkpoint_path and checkpoint_path.is_file() else None
    )
    attempts.append({
        "attempt_no": len(attempts) + 1,
        "started_at": result.get("started_at"),
        "finished_at": result.get("finished_at"),
        "checkpoint_ref": checkpoint_ref,
        "request_ref": result.get("request_ref"),
        "result_ref": result_ref,
        "outcome": result.get("status"),
        "retry_reason_code": None if target_status == "executed_ready" else result.get("error", {}).get("reason_code"),
        "source_status": source_status,
        "target_status": target_status,
    })
    result_refs = list(previous.get("result_refs") or []) + [result_ref]
    retryable = target_status == "execution_failed_retryable"
    previous_remaining = int(previous.get("resume_policy", {}).get("max_additional_attempts", 1))
    completed_attempts_before = len(previous.get("attempts") or [])
    remaining_attempts = max(0, previous_remaining - 1) if completed_attempts_before > 0 else previous_remaining
    log = {
        "schema_id": "repair-execution-log",
        "schema_version": "2.0.0",
        "created_at": previous.get("created_at", utc_now()),
        "extensions": {},
        "run_id": request.get("run_id"),
        "plan_ref": request.get("plan_ref"),
        "request_ref": result.get("request_ref"),
        "result_refs": result_refs,
        "attempts": attempts,
        "current_status": target_status,
        "resume_policy": {
            "resume_allowed": target_status == "executed_ready" or (retryable and remaining_attempts > 0),
            "allowed_platforms": [result.get("runtime_id")] if result.get("runtime_id") else [],
            "required_artifact_refs": [request.get("plan_ref"), result.get("request_ref"), result_ref],
            "max_additional_attempts": remaining_attempts,
            "blocked_reason_code": None if target_status == "executed_ready" else result.get("error", {}).get("code"),
        },
        "gate_check": {
            "gate_id": "repair-execution-log-v5",
            "status": "passed" if target_status == "executed_ready" else "failed",
            "checked_at": utc_now(),
            "predicate_version": "1.0.0",
            "evidence_refs": [ref["artifact_id"] for ref in result_refs],
            "failed_codes": [] if target_status == "executed_ready" else [result.get("error", {}).get("code", "FH-OFFICECLI-EXECUTION-FAILED")],
        },
    }
    path = run_dir / "logs" / "repair_execution_log.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(".json.tmp")
    temp.write_text(json.dumps(log, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n", encoding="utf-8")
    os.replace(temp, path)
    return log


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="OfficeCLI runtime adapter")
    sub = parser.add_subparsers(dest="command", required=True)
    exe = sub.add_parser("execute", help="执行 officecli-execution-request")
    exe.add_argument("--run-dir", required=True, type=Path)
    exe.add_argument("--request", required=True, type=Path)
    exe.add_argument("--officecli-executable", type=Path)
    exe.add_argument("--timeout", type=int, default=120)
    exe.add_argument("--output", type=Path)
    resume = sub.add_parser("resume", help="恢复中断的执行（§15.5）")
    resume.add_argument("--run-dir", required=True, type=Path)
    resume.add_argument("--officecli-executable", type=Path)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        if args.command not in {"execute", "resume"}:
            print("Unknown command", file=sys.stderr)
            return 2

        run_dir = args.run_dir.resolve()
        executable = (args.officecli_executable or
                      Path(".cache/officecli/v1.0.113/win-x64/officecli.exe")).resolve()

        if args.command == "resume":
            return _resume_handler(run_dir, executable)

        request = json.loads(args.request.read_text(encoding="utf-8"))
        existing_log_path = run_dir / "logs" / "repair_execution_log.json"
        existing_log = json.loads(existing_log_path.read_text(encoding="utf-8")) if existing_log_path.is_file() else None
        if existing_log:
            existing_status = existing_log.get("current_status")
            if existing_status == "executed_ready":
                sys.stdout.write(json.dumps({"ok": True, "status": existing_status, "action": "post_write_qa"}) + "\n")
                return 0
            sys.stdout.write(json.dumps({"ok": False, "status": existing_status, "action": "use_resume"}) + "\n")
            return 1

        preflight_errors = validate_execution_request_preflight(executable, request, run_dir)
        if preflight_errors:
            sys.stdout.write(json.dumps({
                "ok": False, "error": FH_OFFICECLI_REQUEST_INVALID,
                "details": preflight_errors,
            }, ensure_ascii=False) + "\n")
            return 2

        result = execute_request(executable, request, run_dir, args.timeout, args.request, attempt_no=1)

        out_path = (args.output or
                    run_dir / "logs" /
                    f"officecli-execution-result.r{request.get('plan_revision', '?')}.a001.json")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(
            json.dumps(result, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n",
            encoding="utf-8",
        )
        progress_path = run_dir / "logs" / "repair_execution_log.json"
        progress_log = json.loads(progress_path.read_text(encoding="utf-8")) if progress_path.is_file() else None
        _write_repair_execution_log(run_dir, request, result, out_path, existing=progress_log)

        summary = {"ok": result["status"] == "done", "result": str(out_path),
                    "status": result["status"], "retryable": result["retryable"]}
        sys.stdout.write(json.dumps(summary, ensure_ascii=False, sort_keys=True) + "\n")
        return 0 if result["status"] == "done" else 1
    except Exception as exc:
        sys.stdout.write(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False) + "\n")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
