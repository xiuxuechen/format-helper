#!/usr/bin/env python3
"""将 finalized repair plan 转换为 officecli-execution-request。

该模块不直接调用 OfficeCLI；只生成确定性的 request JSON。
governance 字段（operation_id/source_action_id/risk_class/target_binding/expected_result/idempotency_key）
不写入原生 batch JSON。"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from jsonschema import Draft202012Validator, FormatChecker

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.officecli.contracts import RAW_SET_ACTIONS
from scripts.officecli.runtime_resolver import EXPECTED_RUNTIME_IDS

CANONICAL_SEP = (",", ":")
BATCH_MAX_OPERATIONS = 12
DEFAULT_TIMEOUT_SECONDS = 120
UTC = timezone.utc
REPAIR_PLAN_SCHEMA_PATH = ROOT / "contracts" / "officecli" / "schemas" / "repair-plan.schema.json"

# §8.2 生产 batch 允许的命令
PRODUCTION_BATCH_COMMANDS = {"set", "add", "remove", "move", "swap", "raw-set"}

# §8.2 governance 字段 — 禁止写入原生 batch JSON
GOVERNANCE_ONLY_FIELDS = {
    "operation_id", "source_action_id", "risk_class",
    "target_binding", "expected_result", "idempotency_key",
}

# 治理 → OfficeCLI 原生映射
GOVERNANCE_TO_NATIVE: dict[str, str] = {
    "command": "command",
    "path": "path",
    "element_type": "type",
    "properties": "props",
    "destination_path": "to",
    "index": "index",
}


def canonical_json_bytes(value: Any) -> bytes:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=CANONICAL_SEP).encode("utf-8")


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def utc_now() -> str:
    return datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def load_repair_plan(plan_path: Path) -> dict[str, Any]:
    """读取 JSON/YAML repair-plan。"""
    if plan_path.suffix.lower() == ".json":
        return json.loads(plan_path.read_text(encoding="utf-8"))
    from scripts.utils.simple_yaml import load_yaml
    return load_yaml(plan_path)


def _load_risk_policy(plan: dict[str, Any], plan_path: Path) -> dict[str, Any] | None:
    """读取 finalized plan 引用的 risk-policy，供白名单闭环复算。"""
    risk_policy_path = plan.get("risk_policy_path")
    if not isinstance(risk_policy_path, str) or not risk_policy_path:
        return None
    run_dir = plan_path.parent.parent if plan_path.parent.name == "plans" else plan_path.parent
    path = (run_dir / risk_policy_path).resolve()
    if not path.is_file():
        return None
    if path.suffix.lower() == ".json":
        return json.loads(path.read_text(encoding="utf-8"))
    from scripts.utils.simple_yaml import load_yaml
    return load_yaml(path)


def _normalize_plan_revision(value: str | int) -> str:
    if isinstance(value, int):
        return f"{value:03d}"
    if isinstance(value, str) and value.isdigit():
        return f"{int(value):03d}"
    return str(value)


def validate_finalized_plan_for_request(
    plan: dict[str, Any],
    *,
    plan_path: Path,
    run_id: str,
    plan_revision: str,
    snapshot_path: Path,
    capability_manifest_path: Path,
) -> list[str]:
    """校验 execution request 来源必须是确定的 finalized repair-plan。"""
    errors: list[str] = []
    schema = json.loads(REPAIR_PLAN_SCHEMA_PATH.read_text(encoding="utf-8"))
    validator = Draft202012Validator(schema, format_checker=FormatChecker())
    errors.extend(
        f"repair-plan schema {error.json_path}: {error.message}"
        for error in sorted(validator.iter_errors(plan), key=lambda item: item.json_path)
    )
    from scripts.validation.manual_review_repair import validate_repair_plan_officecli
    risk_policy = _load_risk_policy(plan, plan_path)
    officecli_result = validate_repair_plan_officecli(plan, risk_policy=risk_policy)
    errors.extend(f"repair-plan officecli: {item}" for item in officecli_result.errors)
    if plan.get("plan_state") != "finalized":
        errors.append("repair-plan plan_state must be finalized")
    if plan.get("run_id") != run_id:
        errors.append("repair-plan run_id must match request run_id")
    if _normalize_plan_revision(plan.get("plan_revision", "")) != _normalize_plan_revision(plan_revision):
        errors.append("repair-plan plan_revision must match request plan_revision")
    snapshot_ref = plan.get("snapshot_ref") or {}
    if snapshot_path.is_file() and snapshot_ref.get("sha256") != sha256_file(snapshot_path):
        errors.append("repair-plan snapshot_ref.sha256 must match snapshot file")
    capability_ref = plan.get("capability_manifest_ref") or {}
    if capability_manifest_path.is_file() and capability_ref.get("sha256") != sha256_file(capability_manifest_path):
        errors.append("repair-plan capability_manifest_ref.sha256 must match capability manifest file")
    return errors


def artifact_ref(path: Path, kind: str, schema_id: str | None = None,
                 schema_version: str | None = None, artifact_id: str | None = None,
                 base_dir: Path | None = None) -> dict[str, Any]:
    file_hash = sha256_file(path)
    base = (base_dir or path.parent).resolve()
    relative_path = str(path.resolve().relative_to(base)).replace("\\", "/")
    return {
        "artifact_id": artifact_id or f"{kind}-{file_hash[:12]}",
        "kind": kind,
        "relative_path": relative_path,
        "sha256": file_hash,
        "size_bytes": path.stat().st_size,
        "schema_id": schema_id,
        "schema_version": schema_version,
    }


def build_native_batch_item(operation: dict[str, Any]) -> dict[str, Any]:
    """将 governance operation 转换为 OfficeCLI 原生 batch item。

    governance 字段（operation_id/source_action_id/risk_class 等）被删除。
    null 属性表示不下发该 key。
    """
    native: dict[str, Any] = {}
    for gov_key, native_key in GOVERNANCE_TO_NATIVE.items():
        if gov_key == "element_type":
            # 仅 add 命令时下发 type
            if operation.get("command") == "add" and operation.get("element_type"):
                native["type"] = operation["element_type"]
            continue
        if gov_key == "destination_path":
            if operation.get("command") == "move" and operation.get("destination_path"):
                native["to"] = operation["destination_path"]
            continue
        if gov_key == "index":
            val = operation.get("index")
            if isinstance(val, int):
                native["index"] = val
            continue
        if gov_key == "properties":
            props = operation.get("properties", {})
            if isinstance(props, dict) and props:
                native["props"] = {str(k): str(v) if not isinstance(v, str) else v
                                   for k, v in props.items() if v is not None}
            continue
        value = operation.get(gov_key)
        if value is not None:
            native[native_key] = value

    # raw-set 命令传递 raw 子字段
    if operation.get("command") == "raw-set":
        raw = operation.get("raw", {})
        if isinstance(raw, dict):
            if raw.get("action") not in RAW_SET_ACTIONS:
                raise ValueError(f"raw-set action invalid: {raw.get('action')}")
            for raw_key in ("part", "xpath", "action", "xml"):
                if raw.get(raw_key) is not None:
                    native[raw_key] = raw[raw_key]

    return native


def compute_idempotency_key(
    plan_sha256: str,
    working_docx_before_sha256: str,
    sequence: int,
    operation: dict[str, Any],
) -> str:
    """§8.3: 确定性幂等键。"""
    op_canonical = canonical_json_bytes({
        k: v for k, v in sorted(operation.items())
        if k in {"command", "path", "element_type", "properties",
                 "index", "destination_path", "raw"}
    }).decode("utf-8")
    payload = f"{plan_sha256}\n{working_docx_before_sha256}\n{sequence}\n{op_canonical}"
    return sha256_bytes(payload.encode("utf-8"))


def action_to_operation(
    action: dict[str, Any],
    index: int,
    batch_sequence: int,
    plan_sha256: str,
    working_docx_before_sha256: str,
) -> dict[str, Any] | None:
    """将单个 RepairAction 转换为 Operation。skip 非 executable 动作。"""
    if action.get("execution_status") != "executable" or action.get("status") != "executable":
        return None
    backend = action.get("backend_action") or {}
    binding = action.get("target_binding") or {}
    command = backend.get("command", "")
    if command not in PRODUCTION_BATCH_COMMANDS:
        return None
    if command == "raw-set":
        raw_action = backend.get("raw") or {}
        if not isinstance(raw_action, dict) or raw_action.get("action") not in RAW_SET_ACTIONS:
            raise ValueError(f"{action.get('action_id')}: raw-set action invalid")

    operation = {
        "operation_id": f"OP-{batch_sequence:03d}-{index:03d}",
        "source_action_id": action.get("action_id", ""),
        "command": command,
        "path": backend.get("path", binding.get("path", "")),
        "element_type": backend.get("element_type"),
        "risk_class": action.get("risk_class", "L2"),
        "properties": backend.get("properties", {}),
        "target_binding": binding,
        "expected_result": {
            "predicate_id": f"POST-OP-{batch_sequence:03d}-{index:03d}",
            "type": "property_equals",
            "target_ref": backend.get("path", binding.get("path", "")),
            "json_pointer": None,
            "expected": True,
            "failure_code": "DFR-OFFICECLI-POSTCONDITION-FAILED",
        },
        "index": backend.get("index"),
        "destination_path": backend.get("destination_path"),
        "raw": backend.get("raw") if command == "raw-set" else None,
        "manual_confirmation_ref": action.get("manual_confirmation_ref"),
        "idempotency_key": "",
    }
    operation["idempotency_key"] = compute_idempotency_key(
        plan_sha256, working_docx_before_sha256, batch_sequence, operation,
    )
    return operation


def plan_to_batches(
    actions: list[dict[str, Any]],
    plan_sha256: str,
    working_docx_before_sha256: str,
) -> list[dict[str, Any]]:
    """将 actions[] 分组为 batches。写操作按原序分组，每批 ≤12。"""
    # 过滤并扁平化为 operations
    operations: list[dict[str, Any]] = []
    for action in actions:
        op = action_to_operation(action, len(operations) + 1, 0, plan_sha256, working_docx_before_sha256)
        if op is not None:
            operations.append(op)

    if not operations:
        return []

    # L3_WRITE 必须独占 batch；其余操作每批最多 12 个。
    grouped_operations: list[list[dict[str, Any]]] = []
    current_group: list[dict[str, Any]] = []
    for operation in operations:
        if operation.get("risk_class") == "L3_WRITE":
            if current_group:
                grouped_operations.append(current_group)
                current_group = []
            grouped_operations.append([operation])
            continue
        current_group.append(operation)
        if len(current_group) == BATCH_MAX_OPERATIONS:
            grouped_operations.append(current_group)
            current_group = []
    if current_group:
        grouped_operations.append(current_group)

    batches: list[dict[str, Any]] = []
    for batch_ops in grouped_operations:
        seq = len(batches) + 1
        # 重新编号操作
        for j, op in enumerate(batch_ops, start=1):
            op["operation_id"] = f"OP-{seq:03d}-{j:03d}"
            op["idempotency_key"] = compute_idempotency_key(
                plan_sha256, working_docx_before_sha256, seq, op,
            )
        batch: dict[str, Any] = {
            "batch_id": f"BATCH-{seq:03d}",
            "sequence": seq,
            "atomicity": "stop_on_first_error",
            "timeout_seconds": DEFAULT_TIMEOUT_SECONDS,
            "max_operations": len(batch_ops),
            "preconditions": [],
            "operations": batch_ops,
            "postconditions": [
                {
                    "predicate_id": f"POST-B-{seq:03d}",
                    "type": "validate_clean",
                    "target_ref": "working_docx",
                    "json_pointer": None,
                    "expected": True,
                    "failure_code": "DFR-OFFICECLI-VALIDATE-FAILED",
                }
            ],
        }
        batches.append(batch)
    return batches


def build_execution_request(
    run_id: str,
    plan_path: Path,
    plan_revision: str,
    working_docx_path: Path,
    snapshot_path: Path,
    lock_path: Path,
    capability_manifest_path: Path,
    officecli_executable_path: Path,
    runtime_id: str,
    request_id: str | None = None,
    artifact_root: Path | None = None,
) -> dict[str, Any]:
    """execution request builder 主入口：构建完整 execution request。"""
    plan = load_repair_plan(plan_path)
    plan_sha = sha256_file(plan_path)
    working_sha = sha256_file(working_docx_path)
    plan_errors = validate_finalized_plan_for_request(
        plan,
        plan_path=plan_path,
        run_id=run_id,
        plan_revision=plan_revision,
        snapshot_path=snapshot_path,
        capability_manifest_path=capability_manifest_path,
    )
    if plan_errors:
        raise ValueError("finalized repair-plan 校验失败: " + "; ".join(plan_errors))

    now = utc_now()
    req_id = request_id or f"REQ-{run_id}-{plan_revision}"
    root = (artifact_root or Path(os.path.commonpath([
        str(path.resolve()) for path in (
            plan_path, working_docx_path, snapshot_path, lock_path,
            capability_manifest_path, officecli_executable_path,
        )
    ]))).resolve()
    if root.is_file():
        root = root.parent

    batches = plan_to_batches(
        plan.get("actions", []), plan_sha, working_sha,
    )

    req_core = {
        "schema_id": "officecli-execution-request",
        "schema_version": "2.0.0",
        "request_id": req_id,
        "run_id": run_id,
        "created_at": now,
        "extensions": {},
        "plan_ref": artifact_ref(plan_path, "plan", "repair-plan", "2.0.0", "repair-plan", root),
        "plan_sha256": plan_sha,
        "plan_revision": plan_revision,
        "working_docx_before_ref": artifact_ref(working_docx_path, "docx", None, None, "working-docx-before", root),
        "snapshot_ref": artifact_ref(snapshot_path, "snapshot", "officecli-document-snapshot", "2.0.0", "before-snapshot", root) if snapshot_path.exists() else None,
        "lock_ref": artifact_ref(lock_path, "lock", "officecli-lock", "1.0.0", "officecli-lock", root) if lock_path.exists() else None,
        "capability_manifest_ref": artifact_ref(capability_manifest_path, "capability", "officecli-capability-manifest", "1.0.0", "officecli-capability", root) if capability_manifest_path.exists() else None,
        "runtime_id": runtime_id,
        "officecli_executable_ref": artifact_ref(officecli_executable_path, "executable", None, None, "officecli-executable", root),
        "environment": {
            "OFFICECLI_SKIP_UPDATE": "1",
            "OFFICECLI_NO_AUTO_RESIDENT": "1",
            "locale": "C.UTF-8",
            "timezone": "UTC",
        },
        "batches": batches,
        "gate_check": {
            "gate_id": "officecli-execution-request-officecli",
            "status": "passed" if batches else "blocked",
            "checked_at": now,
            "predicate_version": "1.0.0",
            "evidence_refs": [],
            "failed_codes": [] if batches else ["DFR-OFFICECLI-NO-EXECUTABLE-ACTIONS"],
        },
    }
    # checkpoint 在每个 batch 启动前才可确定；request 必须显式 null，
    # runtime 随即把真实 hash/size 写入 repair_execution_log。
    for batch in req_core["batches"]:
        batch["checkpoint_ref"] = None
    req_core["request_sha256"] = sha256_bytes(canonical_json_bytes({
        k: v for k, v in req_core.items() if k != "request_sha256"
    }))
    return req_core


def write_native_batches(request: dict[str, Any], output_dir: Path, artifact_root: Path | None = None) -> list[Path]:
    """为每个 batch 生成原生 OfficeCLI batch JSON 文件。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    paths: list[Path] = []
    for batch in request.get("batches", []):
        native_items = [build_native_batch_item(op) for op in batch.get("operations", [])]
        batch_id = batch.get("batch_id", "batch")
        path = output_dir / f"officecli-batch-{batch_id}.json"
        path.write_bytes(canonical_json_bytes(native_items) + b"\n")
        batch["officecli_batch_ref"] = artifact_ref(path, "request", None, None, f"native-{batch_id}", artifact_root or output_dir)
        paths.append(path)
    return paths


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="OfficeCLI execution request builder")
    sub = parser.add_subparsers(dest="command", required=True)
    build = sub.add_parser("build", help="构建 officecli-execution-request")
    build.add_argument("--run-dir", required=True, type=Path)
    build.add_argument("--run-id", required=True)
    build.add_argument("--repair-plan", required=True, type=Path)
    build.add_argument("--plan-revision", required=True)
    build.add_argument("--working-docx", type=Path)
    build.add_argument("--snapshot", type=Path)
    build.add_argument("--lock", type=Path, default=Path("tools/officecli/officecli.lock.json"))
    build.add_argument("--capability-manifest", type=Path, default=Path("tools/officecli/officecli-capability-manifest.json"))
    build.add_argument("--officecli-executable", type=Path)
    build.add_argument("--runtime-id")
    build.add_argument("--output", type=Path)
    build.add_argument("--native-dir", type=Path)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        if args.command != "build":
            print("Unknown command", file=sys.stderr)
            return 2

        run_dir = args.run_dir.resolve()
        working_docx = (args.working_docx or run_dir / "input" / "working.docx").resolve()
        snapshot = (args.snapshot or run_dir / "snapshots" / "officecli-document-snapshot.before.json").resolve()
        lock_path = args.lock.resolve()
        manifest_path = args.capability_manifest.resolve()
        officecli_bin = (args.officecli_executable or Path(".cache/officecli/officecli.exe")).resolve()

        from scripts.officecli.runtime_resolver import detect_runtime_id
        runtime_id = args.runtime_id or detect_runtime_id()
        if runtime_id not in EXPECTED_RUNTIME_IDS:
            print(f"Unsupported runtime_id: {runtime_id}", file=sys.stderr)
            return 2

        request = build_execution_request(
            run_id=args.run_id,
            plan_path=args.repair_plan.resolve(),
            plan_revision=args.plan_revision,
            working_docx_path=working_docx,
            snapshot_path=snapshot,
            lock_path=lock_path,
            capability_manifest_path=manifest_path,
            officecli_executable_path=officecli_bin,
            runtime_id=runtime_id,
            artifact_root=ROOT,
        )

        # 写原生 batch 文件
        native_dir = (args.native_dir or run_dir / "output" / "_internal" / "officecli" / "batches").resolve()
        write_native_batches(request, native_dir, ROOT)
        request["request_sha256"] = sha256_bytes(canonical_json_bytes({
            key: value for key, value in request.items() if key != "request_sha256"
        }))

        # 写 execution request
        out_path = (args.output or run_dir / "plans" /
                    f"officecli-execution-request.r{args.plan_revision}.json").resolve()
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(canonical_json_bytes(request) + b"\n")

        result = {"ok": True, "request": str(out_path), "request_id": request["request_id"],
                   "batch_count": len(request["batches"])}
        sys.stdout.write(json.dumps(result, ensure_ascii=False, sort_keys=True) + "\n")
        return 0
    except Exception as exc:
        sys.stdout.write(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False) + "\n")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
