#!/usr/bin/env python3
"""基于 OfficeCLI 事实产物生成 review-result 2.0.0。"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from scripts.utils.simple_yaml import load_yaml


UTC = timezone.utc
MIN_RENDER_PAGE_BYTES = 50_000
SINGLE_NODE_XPATH_V1 = re.compile(
    r"^/[A-Za-z_][A-Za-z0-9_.-]*:[A-Za-z_][A-Za-z0-9_.-]*\[1\]"
    r"(?:/[A-Za-z_][A-Za-z0-9_.-]*:[A-Za-z_][A-Za-z0-9_.-]*\[[1-9][0-9]*\])*$"
)


def utc_now() -> str:
    return datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def artifact_ref(run_dir: Path, path: Path, artifact_id: str, kind: str, schema_id: str | None) -> dict[str, Any]:
    resolved = path.resolve()
    return {
        "artifact_id": artifact_id,
        "kind": kind,
        "relative_path": str(resolved.relative_to(run_dir.resolve())).replace("\\", "/"),
        "sha256": sha256_file(resolved),
        "size_bytes": resolved.stat().st_size,
        "schema_id": schema_id,
        "schema_version": "2.0.0" if schema_id else None,
    }


def latest_finalized_plan(run_dir: Path) -> Path:
    candidates = sorted(run_dir.glob("plans/repair_plan.finalized.r*.yaml"))
    if not candidates:
        raise FileNotFoundError("缺少 plans/repair_plan.finalized.rNNN.yaml")
    return candidates[-1]


def select_render_pages(run_dir: Path) -> list[Path]:
    qa_pages = sorted((run_dir / "output" / "_internal" / "preview" / "pages").glob("page-*.png"))
    if qa_pages:
        return qa_pages
    candidates: list[tuple[float, list[Path]]] = []
    for render_dir in run_dir.glob("render*"):
        pages = sorted(render_dir.glob("page-*.png")) if render_dir.is_dir() else []
        if pages:
            candidates.append((max(page.stat().st_mtime for page in pages), pages))
    return max(candidates, key=lambda item: item[0])[1] if candidates else []


def flatten_operation_results(run_dir: Path, execution_log: dict[str, Any]) -> tuple[dict[str, list[dict[str, Any]]], list[dict[str, Any]]]:
    by_action: dict[str, list[dict[str, Any]]] = {}
    result_refs: list[dict[str, Any]] = []
    for result_ref in execution_log.get("result_refs", []):
        result_path = run_dir / result_ref["relative_path"]
        if not result_path.is_file() or sha256_file(result_path) != result_ref.get("sha256"):
            raise ValueError(f"执行结果引用缺失或 hash 不符：{result_ref.get('relative_path')}")
        result_refs.append(result_ref)
        result = load_json(result_path)
        for batch in result.get("batch_results", []):
            for operation in batch.get("operation_results", []):
                action_id = operation.get("source_action_id")
                if action_id:
                    by_action.setdefault(action_id, []).append(operation)
    return by_action, result_refs


def _expected_changes(action: dict[str, Any]) -> list[str]:
    backend = action.get("backend_action") or {}
    changes = [f"command={backend.get('command')}", f"path={backend.get('path')}"]
    for key, value in sorted((backend.get("properties") or {}).items()):
        changes.append(f"{key}={json.dumps(value, ensure_ascii=False, sort_keys=True)}")
    return changes


def _node_by_path(snapshot: dict[str, Any], path: str | None) -> dict[str, Any] | None:
    return next((node for node in snapshot.get("nodes", []) if node.get("officecli_path") == path), None)


def _part_evidence_is_valid(run_dir: Path, part: dict[str, Any] | None) -> bool:
    if not isinstance(part, dict) or not isinstance(part.get("evidence_ref"), dict):
        return False
    ref = part["evidence_ref"]
    relative = Path(str(ref.get("relative_path", "")))
    candidates = [run_dir / relative, run_dir.parent.parent / relative]
    path = next((candidate for candidate in candidates if candidate.is_file()), None)
    return bool(
        path
        and sha256_file(path) == ref.get("sha256") == part.get("sha256")
        and path.stat().st_size == ref.get("size_bytes") == part.get("size_bytes")
    )


def _action_result(
    action: dict[str, Any], operations: list[dict[str, Any]],
    evidence_refs: list[dict[str, Any]], before_snapshot: dict[str, Any],
    after_snapshot: dict[str, Any], run_dir: Path,
) -> dict[str, Any]:
    binding = action.get("target_binding") or {}
    node_ref = binding.get("path")
    policy = action.get("auto_fix_policy")
    execution_status = action.get("execution_status") or action.get("status")
    failure_codes: list[str] = []
    before_node = _node_by_path(before_snapshot, node_ref)
    after_node = _node_by_path(after_snapshot, node_ref)
    if policy != "auto-fix":
        status = "manual_required" if policy == "manual-review" else "not_executed"
    elif execution_status != "executable":
        status = "not_executed"
        failure_codes.append("FH-REVIEW-ACTION-NOT-EXECUTABLE")
    elif not operations:
        status = "not_executed"
        failure_codes.append("FH-REVIEW-ACTION-MISSING-RESULT")
    elif any(item.get("status") == "failed" or not item.get("postconditions_passed") for item in operations):
        status = "failed"
        failure_codes.append("FH-REVIEW-ACTION-FAILED")
    elif all(item.get("status") == "executed" and item.get("postconditions_passed") for item in operations):
        status = "passed"
    else:
        status = "not_executed"
        failure_codes.append("FH-REVIEW-ACTION-NOT-RUN")
    backend = action.get("backend_action") or {}
    observed = [f"{item.get('operation_id')}:{item.get('status')}" for item in operations]
    command = backend.get("command")
    if status == "passed" and command == "set":
        if after_node is None:
            status = "failed"
            failure_codes.append("FH-REVIEW-AFTER-NODE-MISSING")
        else:
            actual_properties = {**(after_node.get("attributes") or {}), **(after_node.get("effective_format") or {})}
            mismatches = []
            for key, expected in (backend.get("properties") or {}).items():
                actual = actual_properties.get(key)
                observed.append(f"{key}={json.dumps(actual, ensure_ascii=False, sort_keys=True)}")
                if actual != expected:
                    mismatches.append(key)
            if mismatches:
                status = "failed"
                failure_codes.append("FH-REVIEW-AFTER-PROPERTY-MISMATCH")
    elif status == "passed" and command == "raw-set":
        raw_contract = backend.get("raw") or {}
        part_name = str(raw_contract.get("part", "")).lstrip("/")
        before_part = next((part for part in before_snapshot.get("parts", []) if str(part.get("part_name", "")).lstrip("/") == part_name), None)
        after_part = next((part for part in after_snapshot.get("parts", []) if str(part.get("part_name", "")).lstrip("/") == part_name), None)
        xml = str(raw_contract.get("xml", ""))
        payload_hash_valid = hashlib.sha256(xml.encode("utf-8")).hexdigest() == raw_contract.get("xml_sha256")
        xpath_valid = SINGLE_NODE_XPATH_V1.fullmatch(str(raw_contract.get("xpath", ""))) is not None
        precondition_valid = before_part and before_part.get("sha256") == raw_contract.get("precondition_raw_sha256")
        evidence_valid = _part_evidence_is_valid(run_dir, before_part) and _part_evidence_is_valid(run_dir, after_part)
        if (
            raw_contract.get("expected_match_count") != 1
            or not xpath_valid
            or not payload_hash_valid
            or not precondition_valid
            or not evidence_valid
            or before_part.get("sha256") == after_part.get("sha256")
        ):
            status = "failed"
            failure_codes.append("FH-REVIEW-RAW-WRITE-NOT-VERIFIED")
        else:
            observed.append(f"raw-part-sha256={after_part.get('sha256')}")
    elif status == "passed" and command == "remove" and after_node is not None:
        status = "failed"
        failure_codes.append("FH-REVIEW-REMOVED-NODE-STILL-PRESENT")
    elif status == "passed" and command in {"add", "move", "swap"}:
        output_paths = [item.get("native_output", {}).get("path") for item in operations if isinstance(item.get("native_output"), dict)]
        if not output_paths or any(_node_by_path(after_snapshot, path) is None for path in output_paths):
            status = "failed"
            failure_codes.append("FH-REVIEW-ACTION-UNVERIFIABLE")
    return {
        "action_id": action["action_id"],
        "status": status,
        "before_node_ref": before_node.get("node_id") if before_node else None,
        "after_node_ref": after_node.get("node_id") if after_node and status == "passed" else None,
        "expected_changes": _expected_changes(action),
        "observed_changes": observed,
        "unexpected_changes": [] if status in {"passed", "manual_required"} else observed or ["未找到执行结果"],
        "evidence_refs": evidence_refs,
        "failure_codes": failure_codes,
    }


def build_review(run_dir: Path) -> dict[str, Any]:
    plan_path = latest_finalized_plan(run_dir)
    execution_log_path = run_dir / "logs" / "repair_execution_log.json"
    before_path = run_dir / "snapshots" / "officecli-document-snapshot.before.json"
    after_path = run_dir / "snapshots" / "officecli-document-snapshot.after.json"
    required = [execution_log_path, before_path, after_path]
    missing = [str(path.relative_to(run_dir)) for path in required if not path.is_file()]
    if missing:
        raise FileNotFoundError("缺少二轮复核输入：" + "、".join(missing))

    plan = load_yaml(plan_path)
    execution_log = load_json(execution_log_path)
    qa_path = run_dir / "logs" / "post_write_qa.json"
    qa = load_json(qa_path) if qa_path.is_file() else None
    before_snapshot = load_json(before_path)
    after_snapshot = load_json(after_path)
    operations_by_action, result_refs = flatten_operation_results(run_dir, execution_log)
    plan_ref = artifact_ref(run_dir, plan_path, f"PLAN-{plan.get('plan_id', plan_path.stem)}", "plan", "repair-plan")
    before_ref = artifact_ref(run_dir, before_path, before_snapshot.get("snapshot_id", "SNAP-BEFORE"), "snapshot", "officecli-document-snapshot")
    after_ref = artifact_ref(run_dir, after_path, after_snapshot.get("snapshot_id", "SNAP-AFTER"), "snapshot", "officecli-document-snapshot")
    action_evidence = [*result_refs, before_ref, after_ref]
    action_results = [
        _action_result(
            action, operations_by_action.get(action["action_id"], []), action_evidence,
            before_snapshot, after_snapshot, run_dir,
        )
        for action in plan.get("actions", [])
    ]
    failed_codes: list[str] = []
    if before_snapshot.get("gate_check", {}).get("status") != "passed":
        failed_codes.append("FH-REVIEW-BEFORE-SNAPSHOT-BLOCKED")
    if after_snapshot.get("gate_check", {}).get("status") != "passed":
        failed_codes.append("FH-REVIEW-AFTER-SNAPSHOT-BLOCKED")
    if execution_log.get("current_status") != "review_ready":
        failed_codes.append("FH-REVIEW-EXECUTION-NOT-READY")
    if not isinstance(qa, dict) or qa.get("status") != "passed":
        failed_codes.append("FH-REVIEW-POST-WRITE-QA-NOT-PASSED")
    failed_codes.extend(code for item in action_results for code in item["failure_codes"])
    if any(item["status"] == "manual_required" for item in action_results):
        failed_codes.append("FH-REVIEW-MANUAL-ACTION-UNRESOLVED")
    render_pages = select_render_pages(run_dir)
    if not render_pages:
        failed_codes.append("FH-REVIEW-RENDER-MISSING")
    elif any(page.stat().st_size < MIN_RENDER_PAGE_BYTES for page in render_pages):
        failed_codes.append("FH-REVIEW-RENDER-SUSPECT")
    statuses = ("passed", "failed", "not_executed", "manual_required")
    counts = {status: sum(1 for item in action_results if item["status"] == status) for status in statuses}
    return {
        "schema_id": "review-result",
        "schema_version": "2.0.0",
        "created_at": utc_now(),
        "extensions": {},
        "review_id": f"REV-{plan.get('run_id', run_dir.name)}-FINAL",
        "run_id": plan.get("run_id", run_dir.name),
        "plan_ref": plan_ref,
        "before_snapshot_ref": before_ref,
        "after_snapshot_ref": after_ref,
        "action_results": action_results,
        "summary": {
            "total_actions": len(action_results),
            "passed_count": counts["passed"],
            "failed_count": counts["failed"],
            "not_executed_count": counts["not_executed"],
            "manual_required_count": counts["manual_required"],
            "render_page_count": len(render_pages),
        },
        "gate_check": {
            "gate_id": "review-result-v5",
            "status": "passed" if not failed_codes else "blocked",
            "checked_at": utc_now(),
            "predicate_version": "1.0.0",
            "evidence_refs": [ref["artifact_id"] for ref in action_evidence],
            "failed_codes": sorted(set(failed_codes)),
        },
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="生成 review-result 2.0.0")
    parser.add_argument("--run-dir", required=True, type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)
    output = args.output or args.run_dir / "review_results" / "final_review.json"
    review = build_review(args.run_dir.resolve())
    output.parent.mkdir(parents=True, exist_ok=True)
    temp = output.with_suffix(output.suffix + ".tmp")
    temp.write_text(json.dumps(review, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temp, output)
    sys.stdout.write(json.dumps({"ok": True, "review": str(output), "status": review["gate_check"]["status"]}, ensure_ascii=False) + "\n")
    return 0 if review["gate_check"]["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
