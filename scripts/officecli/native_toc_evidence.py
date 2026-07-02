"""V5-014 专用 Windows runner 原生 TOC 证据生成。"""

from __future__ import annotations

import argparse
import hashlib
import inspect
import json
import os
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.officecli.runtime_resolver import ensure_officecli
from scripts.officecli.toc_refresh_adapter import probe_viewer, refresh_toc


def _viewer_id(viewer_name: str) -> str:
    """将 COM 产品名归一化为 CI 契约 viewer id。"""
    normalized = viewer_name.strip().lower()
    if normalized in {"word", "microsoft word"}:
        return "word"
    if normalized in {"wps", "wps writer"}:
        return "wps"
    return normalized


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _run(command: list[str], env: dict[str, str], timeout: int) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command, text=True, encoding="utf-8", errors="replace",
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout,
        check=False, env=env,
    )


def _officecli_command_succeeded(proc: subprocess.CompletedProcess[str]) -> bool:
    """OfficeCLI 有些 advisory warning 会返回非零，但 JSON success 仍为 true。"""
    if proc.returncode == 0:
        return True
    try:
        payload = json.loads((proc.stdout or "").strip())
    except json.JSONDecodeError:
        return False
    if not isinstance(payload, dict) or payload.get("success") is not True:
        return False
    warnings = payload.get("warnings")
    if not isinstance(warnings, list) or not warnings:
        return False
    return all(isinstance(item, dict) and item.get("code") in {"warning", "advisory"} for item in warnings)


def _probe_required_viewer(viewer_probe: Callable[..., dict[str, Any]], required_viewer: str) -> dict[str, Any]:
    """兼容旧测试注入，同时避免吞掉 probe 内部真实 TypeError。"""
    try:
        signature = inspect.signature(viewer_probe)
    except (TypeError, ValueError):
        return viewer_probe(required_viewer=required_viewer)
    parameters = signature.parameters
    accepts_keyword = any(param.kind == inspect.Parameter.VAR_KEYWORD for param in parameters.values())
    if "required_viewer" in parameters or accepts_keyword:
        return viewer_probe(required_viewer=required_viewer)
    return viewer_probe()


def collect_native_toc_evidence(
    *, workspace_root: Path, lock_path: Path, run_dir: Path,
    required_viewer: str,
    ensure_func: Callable[..., dict[str, Any]] = ensure_officecli,
    command_runner: Callable[[list[str], dict[str, str], int], subprocess.CompletedProcess[str]] = _run,
    viewer_probe: Callable[[], dict[str, Any]] = probe_viewer,
    refresh_func: Callable[..., dict[str, Any]] = refresh_toc,
) -> dict[str, Any]:
    """创建 TOC fixture 并执行 Word/WPS 原生刷新验收。"""
    run_dir.mkdir(parents=True, exist_ok=True)
    (run_dir / "input").mkdir(exist_ok=True)
    (run_dir / "output" / "_internal").mkdir(parents=True, exist_ok=True)
    (run_dir / "logs").mkdir(exist_ok=True)
    resolution = ensure_func(
        lock_path=lock_path, workspace_root=workspace_root,
        runtime_id="win-x64", offline=False, skip_version_check=False,
    )
    executable = Path(resolution["executable_path"])
    input_docx = run_dir / "input" / "native-toc.docx"
    output_docx = run_dir / "output" / "_internal" / "native-toc-refreshed.docx"
    env = os.environ.copy()
    env.update({"OFFICECLI_SKIP_UPDATE": "1", "OFFICECLI_NO_AUTO_RESIDENT": "1"})
    commands = [
        [str(executable), "create", str(input_docx)],
        [str(executable), "add", str(input_docx), "/body", "--type", "toc", "--prop", "levels=1-3", "--prop", "title=目录", "--index", "0", "--json"],
        [str(executable), "add", str(input_docx), "/body", "--type", "paragraph", "--prop", "text=第一章", "--json"],
        [str(executable), "add", str(input_docx), "/body", "--type", "paragraph", "--prop", "text=正文内容", "--json"],
    ]
    command_results: list[dict[str, Any]] = []
    for command in commands:
        proc = command_runner(command, env, 120)
        command_results.append({"command": command, "exit_code": proc.returncode, "stdout": proc.stdout, "stderr": proc.stderr})
        if not _officecli_command_succeeded(proc):
            raise RuntimeError(f"TOC fixture 命令失败：{command[1]}, exit={proc.returncode}")
    viewer = _probe_required_viewer(viewer_probe, required_viewer)
    if not viewer.get("ok"):
        raise RuntimeError(f"Word/WPS viewer 不可用：{viewer}")
    if _viewer_id(str(viewer.get("viewer", ""))) != _viewer_id(required_viewer):
        raise RuntimeError(f"viewer 不匹配：要求 {required_viewer}，实际 {viewer.get('viewer')}")
    viewer["native_toc_fixture_prepare_outline"] = True
    acceptance = refresh_func(
        input_docx, output_docx, viewer,
        officecli_executable=str(executable),
    )
    acceptance["run_id"] = run_dir.name
    page_count = acceptance.get("page_count")
    if not isinstance(page_count, int) or page_count <= 0:
        raise RuntimeError("原生 TOC 验收缺少有效 page_count")
    screenshot_refs: list[dict[str, Any]] = []
    for page_number in range(1, page_count + 1):
        page_path = run_dir / "output" / "_internal" / f"toc-page-{page_number:04d}.png"
        command = [str(executable), "view", str(output_docx), "screenshot", "-o", str(page_path), "--page", str(page_number)]
        proc = command_runner(command, env, 120)
        command_results.append({"command": command, "exit_code": proc.returncode, "stdout": proc.stdout, "stderr": proc.stderr})
        if proc.returncode != 0 or not page_path.is_file() or page_path.stat().st_size == 0:
            raise RuntimeError(f"TOC 第 {page_number} 页截图失败")
        screenshot_refs.append({
            "artifact_id": f"toc-page-{page_number:04d}",
            "kind": "png",
            "relative_path": str(page_path.relative_to(run_dir)).replace("\\", "/"),
            "sha256": _sha256_file(page_path),
            "size_bytes": page_path.stat().st_size,
            "schema_id": None,
            "schema_version": None,
        })
    if len(screenshot_refs) != page_count:
        raise RuntimeError("TOC page_count 与截图证据数量不一致")
    acceptance["evidence_refs"] = list(acceptance.get("evidence_refs") or []) + screenshot_refs
    acceptance_path = run_dir / "logs" / "toc_acceptance.json"
    acceptance_path.write_text(json.dumps(acceptance, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if acceptance.get("status") != "passed" or acceptance.get("gate_check", {}).get("status") != "passed":
        raise RuntimeError(f"原生 TOC 验收失败：{acceptance.get('error')}")
    evidence = {
        "schema_id": "officecli-native-toc-evidence",
        "schema_version": "1.0.0",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "resolution": resolution,
        "viewer": viewer,
        "fixture_commands": command_results,
        "toc_acceptance_path": str(acceptance_path.relative_to(run_dir)).replace("\\", "/"),
        "toc_acceptance": acceptance,
        "page_screenshots": screenshot_refs,
        "status": "passed",
    }
    evidence_path = run_dir / "logs" / "native_toc.platform-evidence.json"
    evidence_path.write_text(json.dumps(evidence, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return evidence


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="生成 Word/WPS 原生 TOC CI 证据")
    parser.add_argument("--workspace-root", type=Path, default=Path.cwd())
    parser.add_argument("--lock", type=Path, default=Path("tools/officecli/officecli.lock.json"))
    parser.add_argument("--run-dir", type=Path, required=True)
    parser.add_argument("--viewer", choices=["word", "wps"], required=True)
    args = parser.parse_args(argv)
    try:
        evidence = collect_native_toc_evidence(
            workspace_root=args.workspace_root.resolve(), lock_path=args.lock.resolve(),
            run_dir=args.run_dir.resolve(), required_viewer=args.viewer,
        )
        sys.stdout.write(json.dumps({"ok": True, "status": evidence["status"]}) + "\n")
        return 0
    except Exception as exc:
        sys.stderr.write(f"V5-014 原生 TOC 证据失败：{exc}\n")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
