#!/usr/bin/env python3
"""写后质量保证 — validate → issues → after snapshot → html → stats → screenshots。

固定顺序（§12.5），任一步失败则阻塞。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

UTC = timezone.utc
MAX_HTML_BYTES = 64 * 1024 * 1024
MAX_PAGES = 500

FH_OFFICECLI_RENDERER_UNAVAILABLE = "FH-OFFICECLI-RENDERER-UNAVAILABLE"


def utc_now() -> str:
    return datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run_officecli(executable: Path, args: list[str], timeout_seconds: int = 120,
                  capture_text: bool = True) -> subprocess.CompletedProcess:
    env = os.environ.copy()
    env["OFFICECLI_SKIP_UPDATE"] = "1"
    env["OFFICECLI_NO_AUTO_RESIDENT"] = "1"
    kwargs: dict[str, Any] = {"timeout": timeout_seconds, "check": False, "env": env}
    if capture_text:
        kwargs["text"] = True
        kwargs["stdout"] = subprocess.PIPE
        kwargs["stderr"] = subprocess.PIPE
    return subprocess.run([str(executable), *args], **kwargs)


def parse_json_stdout(stdout: str) -> Any:
    text = stdout.strip()
    if not text:
        raise ValueError("stdout empty")
    decoder = json.JSONDecoder()
    value, index = decoder.raw_decode(text)
    if text[index:].strip():
        raise ValueError("trailing garbage")
    return value


def unwrap_data(value: Any) -> Any:
    if isinstance(value, dict) and "success" in value and "data" in value:
        if value.get("success") is not True:
            raise ValueError("envelope success=false")
        return value["data"]
    return value


def validate_data_is_clean(data: Any) -> bool:
    """validate 必须明确无错误；未知形状 fail-closed。"""
    if not isinstance(data, dict):
        return False
    if data.get("valid") is False or data.get("clean") is False:
        return False
    for key in ("errors", "blocking_errors", "invalid_parts"):
        value = data.get(key)
        if isinstance(value, list) and value:
            return False
    return data.get("valid") is True or data.get("clean") is True or data.get("errors") == []


def issues_data_is_nonblocking(data: Any) -> bool:
    """issues 结果中任何 blocking/error/critical 项均阻断。"""
    if isinstance(data, dict):
        if data.get("blocking_count", 0) or data.get("error_count", 0):
            return False
        items = data.get("issues", data.get("items", []))
    elif isinstance(data, list):
        items = data
    else:
        return False
    if not isinstance(items, list):
        return False
    return not any(
        isinstance(item, dict)
        and (item.get("blocking") is True or str(item.get("severity", "")).lower() in {"blocking", "blocker", "error", "critical"})
        for item in items
    )


def check_resident_conflict(executable: Path, docx: Path) -> dict[str, Any]:
    """§12.5.1: 检查文件是否被 resident 持有。"""
    proc = run_officecli(executable, ["get", str(docx), "/document", "--json"], timeout_seconds=10)
    if proc.returncode != 0 and "resident" in (proc.stderr or "").lower():
        return {"ok": False, "error": "FH-OFFICECLI-RESIDENT-CONFLICT"}
    return {"ok": True}


def run_validate(executable: Path, docx: Path) -> dict[str, Any]:
    proc = run_officecli(executable, ["validate", str(docx), "--json"])
    if proc.returncode != 0:
        return {"ok": False, "exit_code": proc.returncode, "stderr": proc.stderr}
    try:
        parsed = parse_json_stdout(proc.stdout)
        data = unwrap_data(parsed)
    except ValueError as exc:
        return {"ok": False, "exit_code": proc.returncode, "stderr": proc.stderr, "error": str(exc)}
    return {"ok": validate_data_is_clean(data), "data": data,
            "error": None if validate_data_is_clean(data) else "validate result is not clean"}


def run_issues(executable: Path, docx: Path) -> dict[str, Any]:
    proc = run_officecli(executable, ["view", str(docx), "issues", "--json"])
    if proc.returncode != 0:
        return {"ok": False, "exit_code": proc.returncode, "stderr": proc.stderr}
    try:
        parsed = parse_json_stdout(proc.stdout)
        data = unwrap_data(parsed)
    except ValueError as exc:
        return {"ok": False, "exit_code": proc.returncode, "stderr": proc.stderr, "error": str(exc)}
    return {"ok": issues_data_is_nonblocking(data), "data": data,
            "error": None if issues_data_is_nonblocking(data) else "blocking issues found"}


def run_html_preview(executable: Path, docx: Path, output_dir: Path) -> dict[str, Any]:
    proc = run_officecli(executable, ["view", str(docx), "html"])
    if proc.returncode != 0:
        return {"ok": False, "exit_code": proc.returncode, "stderr": proc.stderr}
    stdout_bytes = proc.stdout.encode("utf-8") if proc.stdout else b""
    if not stdout_bytes or len(stdout_bytes) > MAX_HTML_BYTES:
        return {"ok": False, "error": f"HTML stdout empty or exceeds {MAX_HTML_BYTES} bytes"}
    output_dir.mkdir(parents=True, exist_ok=True)
    html_path = output_dir / "document.html"
    # §21.5: 原子写入
    tmp_path = output_dir / "document.html.tmp"
    tmp_path.write_bytes(stdout_bytes)
    tmp_path.replace(html_path)
    return {"ok": True, "html_path": str(html_path), "sha256": sha256_file(html_path),
            "size_bytes": len(stdout_bytes)}


def run_stats(executable: Path, docx: Path) -> dict[str, Any]:
    proc = run_officecli(executable, ["view", str(docx), "stats", "--page-count", "--json"])
    if proc.returncode != 0:
        return {"ok": False, "exit_code": proc.returncode, "stderr": proc.stderr}
    try:
        parsed = parse_json_stdout(proc.stdout)
        data = unwrap_data(parsed)
    except ValueError as exc:
        return {"ok": False, "exit_code": proc.returncode, "stderr": proc.stderr, "error": str(exc)}
    pages = data.get("pages") if isinstance(data, dict) else None
    if not isinstance(pages, int) or pages <= 0 or pages > MAX_PAGES:
        return {"ok": False, "error": f"page count invalid or exceeds {MAX_PAGES}: {pages}"}
    return {"ok": True, "pages": pages}


def run_screenshots(executable: Path, docx: Path, pages: int, output_dir: Path) -> list[dict[str, Any]]:
    render_dir = output_dir / "pages"
    render_dir.mkdir(parents=True, exist_ok=True)
    results: list[dict[str, Any]] = []
    for n in range(1, pages + 1):
        out_file = render_dir / f"page-{n:04d}.png"
        proc = run_officecli(executable, [
            "view", str(docx), "screenshot", "--page", str(n),
            "--render", "html", "--out", str(out_file),
        ])
        ok = proc.returncode == 0 and out_file.exists()
        results.append({
            "page": n, "ok": ok, "path": str(out_file),
            "sha256": sha256_file(out_file) if ok else None,
            "exit_code": proc.returncode,
        })
        if not ok:
            break
    return results


def _qa_blocked(run_id: str, executed_docx: Path, now: str, errors: list[str],
                evidence: dict[str, Any], run_dir: Path) -> dict[str, Any]:
    """构造阻塞 QA 结果。"""
    result = {
        "schema_id": "post-write-qa-result",
        "schema_version": "1.0.0",
        "run_id": run_id,
        "executed_docx": str(executed_docx),
        "checked_at": now,
        "status": "blocked",
        "errors": errors,
        "evidence": evidence,
    }
    qa_path = run_dir / "logs" / "post_write_qa.json"
    qa_path.parent.mkdir(parents=True, exist_ok=True)
    qa_path.write_text(json.dumps(result, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n", encoding="utf-8")
    return result


def run_full_qa(
    executable: Path,
    executed_docx: Path,
    run_dir: Path,
    snapshot_executable: Path | None = None,
    capability_manifest: Path | None = None,
) -> dict[str, Any]:
    """写后 QA 主入口：执行完整写后 QA 序列。"""
    run_id = run_dir.name
    now = utc_now()
    output_dir = run_dir / "output" / "_internal"
    preview_dir = output_dir / "preview"
    errors: list[str] = []
    evidence: dict[str, Any] = {}

    # 0. resident conflict check (§12.5.1)
    resident = check_resident_conflict(executable, executed_docx)
    evidence["resident_check"] = resident
    if not resident["ok"]:
        errors.append(resident["error"])
        # resident conflict is blocking — return immediately
        return _qa_blocked(run_id, executed_docx, now, errors, evidence, run_dir)

    # 1. validate
    v = run_validate(executable, executed_docx)
    evidence["validate"] = v
    if not v["ok"]:
        errors.append(f"DFR-OFFICECLI-VALIDATE-FAILED: validate failed")
        return _qa_blocked(run_id, executed_docx, now, errors, evidence, run_dir)

    # 2. issues
    iss = run_issues(executable, executed_docx)
    evidence["issues"] = iss
    if not iss["ok"]:
        errors.append(f"FH-OFFICECLI-NONJSON-OUTPUT: issues scan failed")
        return _qa_blocked(run_id, executed_docx, now, errors, evidence, run_dir)

    # 3. after snapshot (delegate to snapshot_adapter if available)
    after_snapshot_path = run_dir / "snapshots" / "officecli-document-snapshot.after.json"
    if snapshot_executable and capability_manifest:
        try:
            from scripts.officecli.snapshot_adapter import (
                collect_snapshot_inputs_with_officecli, build_snapshot,
            )
            collected = collect_snapshot_inputs_with_officecli(
                snapshot_executable, executed_docx,
                capability_manifest, run_dir, "after",
            )
            nodes = collected["raw_nodes"]
            warnings = collected["warnings"]
            parts = collected["parts"]
            snap = build_snapshot(
                run_id=run_id, kind="after", source_docx=executed_docx,
                officecli_executable=snapshot_executable,
                capability_manifest=capability_manifest,
                artifact_root=run_dir.parent if run_dir.parent.name == "format_runs" else run_dir.parent,
                raw_nodes=nodes, parts=parts, warnings=warnings,
            )
            after_snapshot_path.parent.mkdir(parents=True, exist_ok=True)
            import json as _json
            after_snapshot_path.write_text(_json.dumps(snap, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n", encoding="utf-8")
            evidence["after_snapshot"] = {"ok": True, "path": str(after_snapshot_path)}
        except Exception as exc:
            evidence["after_snapshot"] = {"ok": False, "error": str(exc)}
            errors.append(f"after snapshot failed: {exc}")
    else:
        evidence["after_snapshot"] = {"ok": False, "error": "snapshot adapter not configured"}
        errors.append("FH-OFFICECLI-SNAPSHOT-NOT-CONFIGURED: after snapshot is required")

    # 4. HTML preview
    html = run_html_preview(executable, executed_docx, preview_dir)
    evidence["html"] = html
    if not html["ok"]:
        errors.append(f"HTML preview failed: {html.get('error', html)}")

    # 5. stats page-count
    stats = run_stats(executable, executed_docx)
    evidence["stats"] = stats
    if not stats["ok"]:
        errors.append(f"stats page-count failed: {stats.get('error', stats)}")

    # 6. screenshots
    screenshots = []
    if stats["ok"]:
        screenshots = run_screenshots(executable, executed_docx, stats["pages"], preview_dir)
        evidence["screenshots"] = screenshots
        if not all(s["ok"] for s in screenshots):
            errors.append("screenshots incomplete or failed")

    result = {
        "schema_id": "post-write-qa-result",
        "schema_version": "1.0.0",
        "run_id": run_id,
        "executed_docx": str(executed_docx),
        "checked_at": now,
        "status": "passed" if not errors else "blocked",
        "errors": errors,
        "evidence": evidence,
    }

    qa_path = run_dir / "logs" / "post_write_qa.json"
    qa_path.parent.mkdir(parents=True, exist_ok=True)
    qa_path.write_text(json.dumps(result, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n", encoding="utf-8")

    if result["status"] == "passed":
        execution_log_path = run_dir / "logs" / "repair_execution_log.json"
        if not execution_log_path.is_file():
            result["status"] = "blocked"
            result["errors"].append("FH-OFFICECLI-EXECUTION-LOG-MISSING")
            qa_path.write_text(json.dumps(result, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n", encoding="utf-8")
        else:
            execution_log = json.loads(execution_log_path.read_text(encoding="utf-8"))
            if execution_log.get("current_status") != "executed_ready":
                result["status"] = "blocked"
                result["errors"].append("FH-OFFICECLI-EXECUTION-STATE-INVALID")
                qa_path.write_text(json.dumps(result, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n", encoding="utf-8")
            else:
                execution_log["current_status"] = "review_ready"
                execution_log["gate_check"]["checked_at"] = utc_now()
                temp_log = execution_log_path.with_suffix(".json.tmp")
                temp_log.write_text(json.dumps(execution_log, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n", encoding="utf-8")
                os.replace(temp_log, execution_log_path)

    return result


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="写后 QA")
    sub = parser.add_subparsers(dest="command", required=True)
    run = sub.add_parser("run", help="执行写后 QA 序列")
    run.add_argument("--run-dir", required=True, type=Path)
    run.add_argument("--executed-docx", required=True, type=Path)
    run.add_argument("--officecli-executable", required=True, type=Path)
    run.add_argument("--capability-manifest", type=Path)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        if args.command != "run":
            print("Unknown command", file=sys.stderr)
            return 2

        run_dir = args.run_dir.resolve()
        executed = args.executed_docx.resolve()
        executable = args.officecli_executable.resolve()

        snap_exe = executable if executable.exists() else None
        cap_manifest = args.capability_manifest.resolve() if args.capability_manifest else None

        result = run_full_qa(
            executable=executable, executed_docx=executed,
            run_dir=run_dir, snapshot_executable=snap_exe,
            capability_manifest=cap_manifest,
        )

        sys.stdout.write(json.dumps({"ok": result["status"] == "passed", "status": result["status"],
                                      "errors": result["errors"]}, ensure_ascii=False) + "\n")
        return 0 if result["status"] == "passed" else 1
    except Exception as exc:
        sys.stdout.write(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False) + "\n")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
