"""V5-014 静态生产路径与 win/mac 必过平台证据 Gate。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.officecli.runtime_resolver import EXPECTED_RUNTIME_IDS, load_lock, select_asset


FORBIDDEN_PYTHON_TOKENS = ("import zipfile", "xml.etree", "from lxml", "python-docx", "from docx")
REQUIRED_SMOKE_COMMANDS = ["version", "create", "add", "get", "set", "validate", "screenshot"]
RETIRED_SKILLS = {"docx-format-repairer"}
REQUIRED_RELEASE_RUNTIME_IDS = {"win-x64", "osx-arm64"}
EXPECTED_RUNNER = {
    "win-x64": ("windows", {"amd64", "x86_64"}),
    "win-arm64": ("windows", {"arm64", "aarch64"}),
    "linux-x64-gnu": ("linux", {"x86_64", "amd64"}),
    "linux-arm64-gnu": ("linux", {"aarch64", "arm64"}),
    "linux-x64-musl": ("linux", {"x86_64", "amd64"}),
    "linux-arm64-musl": ("linux", {"aarch64", "arm64"}),
    "osx-x64": ("darwin", {"x86_64", "amd64"}),
    "osx-arm64": ("darwin", {"arm64", "aarch64"}),
}


def _validate_smoke_command_argv(runtime_id: str, commands: list[dict[str, Any]]) -> list[str]:
    """校验平台 smoke 证据中的关键 argv，防止手工旧证据绕过生成器契约。"""
    errors: list[str] = []
    by_name = {item.get("name"): item for item in commands}
    create_command = by_name.get("create", {}).get("command")
    if (
        not isinstance(create_command, list)
        or len(create_command) < 4
        or create_command[1] != "create"
        or "--force" not in create_command[3:]
    ):
        errors.append(f"{runtime_id}: create smoke command must include --force")
    return errors


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def scan_production_paths(root: Path) -> list[str]:
    """扫描活跃 v5 Python/Skill，退役目录不计入生产路径。"""
    errors: list[str] = []
    python_files = list((root / "scripts" / "officecli").glob("*.py"))
    skill_root = root / ".codex" / "skills"
    for skill_dir in skill_root.iterdir():
        if not skill_dir.is_dir() or skill_dir.name in RETIRED_SKILLS:
            continue
        python_files.extend(skill_dir.glob("scripts/*.py"))
        skill_file = skill_dir / "SKILL.md"
        if skill_file.exists():
            content = skill_file.read_text(encoding="utf-8")
            if re.search(r"(?im)^\s*(python|py)\s+.*scripts[/\\]ooxml", content):
                errors.append(f"{skill_file.relative_to(root)} invokes retired scripts/ooxml")
            if re.search(r"(?im)^\s*(python|py)\s+.*apply_repair_plan\.py", content):
                errors.append(f"{skill_file.relative_to(root)} invokes retired repair backend")
    for path in python_files:
        if path.resolve() == Path(__file__).resolve():
            continue
        content = path.read_text(encoding="utf-8")
        for token in FORBIDDEN_PYTHON_TOKENS:
            if token in content:
                errors.append(f"{path.relative_to(root)} contains forbidden production token: {token}")
    reporter = root / ".codex" / "skills" / "docx-format-reporter" / "scripts" / "render_final_reports.py"
    if "officecli-document-snapshot.before.json" not in reporter.read_text(encoding="utf-8"):
        errors.append("reporter does not consume officecli snapshot v2")
    workflow = root / ".github" / "workflows" / "officecli-v5.yml"
    if not workflow.exists():
        errors.append("missing .github/workflows/officecli-v5.yml")
    else:
        workflow_text = workflow.read_text(encoding="utf-8")
        for runtime_id in sorted(REQUIRED_RELEASE_RUNTIME_IDS):
            if runtime_id not in workflow_text:
                errors.append(f"workflow missing runtime_id: {runtime_id}")
        for required_label in ("native-toc-evidence", "officecli-windows-word", "officecli-windows-wps"):
            if required_label not in workflow_text:
                errors.append(f"workflow missing dedicated native TOC requirement: {required_label}")
    return errors


def _load_platform_evidence(evidence_root: Path) -> dict[str, tuple[dict[str, Any], Path]]:
    found: dict[str, tuple[dict[str, Any], Path]] = {}
    for path in evidence_root.rglob("*.platform-evidence.json"):
        payload = json.loads(path.read_text(encoding="utf-8"))
        runtime_id = payload.get("runtime_id")
        if runtime_id in found:
            if runtime_id in REQUIRED_RELEASE_RUNTIME_IDS:
                raise ValueError(f"重复平台证据：{runtime_id}")
            continue
        if isinstance(runtime_id, str):
            found[runtime_id] = (payload, path)
    return found


def validate_platform_evidence(evidence_root: Path, lock_path: Path, capability_path: Path) -> list[str]:
    """验证 win/mac 必过平台证据集合满足最终发布 Gate。"""
    errors: list[str] = []
    lock = load_lock(lock_path)
    capability = json.loads(capability_path.read_text(encoding="utf-8"))
    expected_capability_hash = capability.get("aggregate_sha256")
    expected_lock_file_hash = _sha256_file(lock_path)
    expected_capability_file_hash = _sha256_file(capability_path)
    try:
        found = _load_platform_evidence(evidence_root)
    except (OSError, json.JSONDecodeError, ValueError) as exc:
        return [f"平台证据读取失败：{exc}"]
    missing = sorted(REQUIRED_RELEASE_RUNTIME_IDS - set(found))
    extra = sorted(set(found) - EXPECTED_RUNTIME_IDS)
    if missing:
        errors.append(f"缺少发布必过平台证据：{', '.join(missing)}")
    if extra:
        errors.append(f"存在未知平台证据：{', '.join(extra)}")
    for runtime_id in sorted(REQUIRED_RELEASE_RUNTIME_IDS & set(found)):
        payload, evidence_path = found[runtime_id]
        evidence_dir = evidence_path.parent
        resolution = payload.get("resolution", {})
        asset = select_asset(lock, runtime_id)
        if payload.get("status") != "passed":
            errors.append(f"{runtime_id}: status must be passed")
        if resolution.get("runtime_id") != runtime_id:
            errors.append(f"{runtime_id}: resolution runtime mismatch")
        if resolution.get("officecli_version") != "1.0.113" or resolution.get("version") != "1.0.113":
            errors.append(f"{runtime_id}: OfficeCLI version mismatch")
        if resolution.get("sha256") != asset["sha256"] or resolution.get("size_bytes") != asset["size_bytes"]:
            errors.append(f"{runtime_id}: locked asset hash/size mismatch")
        if payload.get("capability_aggregate_sha256") != expected_capability_hash:
            errors.append(f"{runtime_id}: capability hash mismatch")
        if payload.get("lock_sha256") != expected_lock_file_hash:
            errors.append(f"{runtime_id}: lock file hash mismatch")
        if payload.get("capability_file_sha256") != expected_capability_file_hash:
            errors.append(f"{runtime_id}: capability file hash mismatch")
        runner = payload.get("runner")
        if not isinstance(runner, dict) or not all(runner.get(key) for key in ("system", "release", "machine", "python")):
            errors.append(f"{runtime_id}: runner information incomplete")
        elif str(runner["system"]).lower() != EXPECTED_RUNNER[runtime_id][0] or str(runner["machine"]).lower() not in EXPECTED_RUNNER[runtime_id][1]:
            errors.append(f"{runtime_id}: runner system/machine mismatch")
        elif runtime_id.startswith("linux-"):
            expected_libc = "musl" if runtime_id.endswith("-musl") else "glibc"
            if runner.get("libc") != expected_libc:
                errors.append(f"{runtime_id}: runner libc mismatch")
            if expected_libc == "musl" and (
                runner.get("distribution_id") != "alpine"
                or not runner.get("distribution_version")
                or runner.get("distribution_version") == "unknown"
            ):
                errors.append(f"{runtime_id}: Alpine distribution evidence missing")
        environment = payload.get("environment")
        if not isinstance(environment, dict) or environment.get("OFFICECLI_SKIP_UPDATE") != "1" or environment.get("OFFICECLI_NO_AUTO_RESIDENT") != "1":
            errors.append(f"{runtime_id}: OfficeCLI update/resident environment is unsafe")
        commands = payload.get("commands")
        names = [item.get("name") for item in commands] if isinstance(commands, list) else []
        if names != REQUIRED_SMOKE_COMMANDS:
            errors.append(f"{runtime_id}: smoke command sequence mismatch")
        elif any(item.get("exit_code") != 0 or item.get("business_success") is not True for item in commands):
            errors.append(f"{runtime_id}: smoke command failed")
        else:
            errors.extend(_validate_smoke_command_argv(runtime_id, commands))
            version_stdout = evidence_dir / str(commands[0].get("stdout", {}).get("path", ""))
            if not version_stdout.is_file() or version_stdout.read_text(encoding="utf-8").strip().lstrip("v") != "1.0.113":
                errors.append(f"{runtime_id}: version stdout mismatch")
            for item in commands:
                for stream in ("stdout", "stderr"):
                    ref = item.get(stream)
                    if not isinstance(ref, dict) or not re.fullmatch(r"[a-f0-9]{64}", str(ref.get("sha256", ""))) or not isinstance(ref.get("size_bytes"), int):
                        errors.append(f"{runtime_id}: {item.get('name')} {stream} evidence invalid")
                        continue
                    artifact_path = evidence_dir / str(ref.get("path", ""))
                    if not artifact_path.is_file() or _sha256_file(artifact_path) != ref["sha256"] or artifact_path.stat().st_size != ref["size_bytes"]:
                        errors.append(f"{runtime_id}: {item.get('name')} {stream} artifact mismatch")
        for artifact_name in ("smoke_docx", "screenshot"):
            artifact = payload.get(artifact_name)
            if not isinstance(artifact, dict) or artifact.get("size_bytes", 0) <= 0 or not re.fullmatch(r"[a-f0-9]{64}", str(artifact.get("sha256", ""))):
                errors.append(f"{runtime_id}: invalid {artifact_name} evidence")
                continue
            artifact_path = evidence_dir / str(artifact.get("path", ""))
            if not artifact_path.is_file() or _sha256_file(artifact_path) != artifact["sha256"] or artifact_path.stat().st_size != artifact["size_bytes"]:
                errors.append(f"{runtime_id}: {artifact_name} artifact mismatch")
    return errors


def main(argv: list[str] | None = None) -> int:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="OfficeCLI v5 release gate")
    sub = parser.add_subparsers(dest="command", required=True)
    static = sub.add_parser("static")
    static.add_argument("--root", type=Path, default=Path.cwd())
    platform_gate = sub.add_parser("platform")
    platform_gate.add_argument("--evidence-root", required=True, type=Path)
    platform_gate.add_argument("--lock", required=True, type=Path)
    platform_gate.add_argument("--capability", required=True, type=Path)
    args = parser.parse_args(argv)
    if args.command == "static":
        errors = scan_production_paths(args.root.resolve())
    else:
        errors = validate_platform_evidence(args.evidence_root, args.lock, args.capability)
    sys.stdout.write(json.dumps({"ok": not errors, "errors": errors}, ensure_ascii=False) + "\n")
    return 0 if not errors else 2


if __name__ == "__main__":
    raise SystemExit(main())
