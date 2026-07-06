"""OfficeCLI 静态生产路径、win/mac 平台证据与 native TOC 证据 Gate。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any

from jsonschema import Draft202012Validator, FormatChecker

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.officecli.runtime_resolver import EXPECTED_RUNTIME_IDS, load_lock, select_asset


FORBIDDEN_PYTHON_TOKENS = ("import zipfile", "xml.etree", "from lxml", "python-docx", "from docx")
REQUIRED_SMOKE_COMMANDS = ["version", "create", "add", "get", "set", "validate", "screenshot"]
RETIRED_SKILLS = {"docx-format-repairer"}
REQUIRED_RELEASE_RUNTIME_IDS = {"win-x64", "osx-arm64"}
REQUIRED_NATIVE_TOC_VIEWERS = {"word", "wps"}
TOC_ACCEPTANCE_SCHEMA = ROOT / "contracts" / "officecli" / "schemas" / "toc-acceptance.schema.json"
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


def _sha256_utf8_lf_file(path: Path) -> str:
    text = path.read_text(encoding="utf-8")
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def scan_production_paths(root: Path) -> list[str]:
    """扫描活跃 OfficeCLI Python/Skill，退役目录不计入生产路径。"""
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
    workflow = root / ".github" / "workflows" / "officecli-platform-evidence.yml"
    if not workflow.exists():
        errors.append("missing .github/workflows/officecli-platform-evidence.yml")
    else:
        workflow_text = workflow.read_text(encoding="utf-8")
        for runtime_id in sorted(REQUIRED_RELEASE_RUNTIME_IDS):
            if runtime_id not in workflow_text:
                errors.append(f"workflow missing runtime_id: {runtime_id}")
        for required_label in ("native-toc-evidence", "run_native_toc_dedicated", "officecli-windows-word", "officecli-windows-wps"):
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
    expected_lock_file_hash = _sha256_utf8_lf_file(lock_path)
    expected_capability_file_hash = _sha256_utf8_lf_file(capability_path)
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


def _viewer_id(viewer_name: str) -> str:
    normalized = viewer_name.strip().lower()
    if normalized in {"word", "microsoft word"}:
        return "word"
    if normalized in {"wps", "wps writer"}:
        return "wps"
    return normalized


def _load_native_toc_evidence(evidence_root: Path) -> dict[str, tuple[dict[str, Any], Path]]:
    found: dict[str, tuple[dict[str, Any], Path]] = {}
    for path in evidence_root.rglob("native_toc.platform-evidence.json"):
        payload = json.loads(path.read_text(encoding="utf-8"))
        viewer_payload = payload.get("viewer")
        if not isinstance(viewer_payload, dict):
            continue
        viewer_id = _viewer_id(str(viewer_payload.get("viewer", "")))
        if viewer_id in found:
            raise ValueError(f"重复 native TOC 证据：{viewer_id}")
        if viewer_id:
            found[viewer_id] = (payload, path)
    return found


def _validate_toc_acceptance_payload(viewer_id: str, payload: dict[str, Any], evidence_dir: Path, errors: list[str]) -> None:
    """校验 native TOC 内嵌和落盘 toc_acceptance 的发布证据强度。"""
    toc_path_value = payload.get("toc_acceptance_path")
    if not isinstance(toc_path_value, str) or not toc_path_value:
        errors.append(f"{viewer_id}: toc_acceptance_path required")
        return
    toc_path = evidence_dir / toc_path_value
    if not toc_path.is_file():
        errors.append(f"{viewer_id}: toc_acceptance_path missing")
        return
    try:
        file_acceptance = json.loads(toc_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        errors.append(f"{viewer_id}: toc_acceptance_path unreadable: {exc}")
        return
    acceptance = payload.get("toc_acceptance")
    if not isinstance(acceptance, dict):
        errors.append(f"{viewer_id}: toc_acceptance missing")
        return
    if acceptance != file_acceptance:
        errors.append(f"{viewer_id}: embedded toc_acceptance does not match toc_acceptance_path")
    schema = json.loads(TOC_ACCEPTANCE_SCHEMA.read_text(encoding="utf-8"))
    validator = Draft202012Validator(schema, format_checker=FormatChecker())
    schema_errors = sorted(validator.iter_errors(acceptance), key=lambda error: list(error.path))
    for error in schema_errors:
        errors.append(f"{viewer_id}: toc_acceptance schema error: {error.message}")
    if acceptance.get("status") != "passed" or acceptance.get("gate_check", {}).get("status") != "passed":
        errors.append(f"{viewer_id}: toc_acceptance must pass")
    if _viewer_id(str(acceptance.get("viewer", ""))) != viewer_id:
        errors.append(f"{viewer_id}: toc_acceptance viewer mismatch")
    if not isinstance(acceptance.get("before_sha256"), str) or not re.fullmatch(r"[a-f0-9]{64}", acceptance["before_sha256"]):
        errors.append(f"{viewer_id}: before_sha256 required")
    if not isinstance(acceptance.get("after_sha256"), str) or not re.fullmatch(r"[a-f0-9]{64}", acceptance["after_sha256"]):
        errors.append(f"{viewer_id}: after_sha256 required")
    if acceptance.get("before_sha256") == acceptance.get("after_sha256"):
        errors.append(f"{viewer_id}: before_sha256 and after_sha256 must differ")
    for key in ("field_update_count", "toc_update_count"):
        if not isinstance(acceptance.get(key), int) or acceptance[key] <= 0:
            errors.append(f"{viewer_id}: {key} must be positive")
    page_count = acceptance.get("page_count")
    if not isinstance(page_count, int) or page_count <= 0:
        errors.append(f"{viewer_id}: page_count must be positive")
    visible_entries = acceptance.get("visible_entries")
    if not isinstance(visible_entries, list) or not visible_entries:
        errors.append(f"{viewer_id}: visible_entries required")
    else:
        for entry in visible_entries:
            if not isinstance(entry, dict) or not isinstance(entry.get("page_number"), int) or entry["page_number"] <= 0:
                errors.append(f"{viewer_id}: visible_entries page_number must be positive")
    screenshot_refs = payload.get("page_screenshots")
    if isinstance(page_count, int) and isinstance(screenshot_refs, list) and len(screenshot_refs) != page_count:
        errors.append(f"{viewer_id}: page_count must equal page_screenshots count")
    evidence_refs = acceptance.get("evidence_refs")
    if not isinstance(evidence_refs, list) or not isinstance(screenshot_refs, list):
        errors.append(f"{viewer_id}: toc_acceptance evidence_refs and page_screenshots required")
        return
    evidence_keys = {
        (
            ref.get("relative_path"),
            ref.get("sha256"),
            ref.get("size_bytes"),
        )
        for ref in evidence_refs
        if isinstance(ref, dict)
    }
    for ref in screenshot_refs:
        if not isinstance(ref, dict):
            continue
        key = (ref.get("relative_path"), ref.get("sha256"), ref.get("size_bytes"))
        if key not in evidence_keys:
            errors.append(f"{viewer_id}: toc_acceptance evidence_refs must cover page_screenshots")


def validate_native_toc_evidence(evidence_root: Path, lock_path: Path) -> list[str]:
    """验证 Word/WPS native TOC 证据，可来自 CI dedicated runner 或维护者本机。"""
    errors: list[str] = []
    lock = load_lock(lock_path)
    win_asset = select_asset(lock, "win-x64")
    try:
        found = _load_native_toc_evidence(evidence_root)
    except (OSError, json.JSONDecodeError, ValueError) as exc:
        return [f"native TOC 证据读取失败：{exc}"]
    missing = sorted(REQUIRED_NATIVE_TOC_VIEWERS - set(found))
    if missing:
        errors.append(f"缺少 native TOC 证据：{', '.join(missing)}")
    for viewer_id in sorted(REQUIRED_NATIVE_TOC_VIEWERS & set(found)):
        payload, evidence_path = found[viewer_id]
        evidence_dir = evidence_path.parent.parent
        if payload.get("schema_id") != "officecli-native-toc-evidence":
            errors.append(f"{viewer_id}: schema_id must be officecli-native-toc-evidence")
        if payload.get("status") != "passed":
            errors.append(f"{viewer_id}: status must be passed")
        resolution = payload.get("resolution")
        if not isinstance(resolution, dict):
            errors.append(f"{viewer_id}: resolution missing")
        else:
            if resolution.get("runtime_id") != "win-x64":
                errors.append(f"{viewer_id}: native TOC must use win-x64 OfficeCLI")
            if resolution.get("officecli_version") != "1.0.113" or resolution.get("version") != "1.0.113":
                errors.append(f"{viewer_id}: OfficeCLI version mismatch")
            if resolution.get("sha256") != win_asset["sha256"] or resolution.get("size_bytes") != win_asset["size_bytes"]:
                errors.append(f"{viewer_id}: locked win-x64 asset hash/size mismatch")
        viewer = payload.get("viewer")
        if not isinstance(viewer, dict) or viewer.get("ok") is not True:
            errors.append(f"{viewer_id}: viewer probe must be ok")
        elif _viewer_id(str(viewer.get("viewer", ""))) != viewer_id or not viewer.get("version"):
            errors.append(f"{viewer_id}: viewer identity/version mismatch")
        _validate_toc_acceptance_payload(viewer_id, payload, evidence_dir, errors)
        screenshots = payload.get("page_screenshots")
        if not isinstance(screenshots, list) or not screenshots:
            errors.append(f"{viewer_id}: page_screenshots required")
        else:
            for ref in screenshots:
                if not isinstance(ref, dict) or not re.fullmatch(r"[a-f0-9]{64}", str(ref.get("sha256", ""))) or not isinstance(ref.get("size_bytes"), int):
                    errors.append(f"{viewer_id}: screenshot ref invalid")
                    continue
                artifact_path = evidence_dir / str(ref.get("relative_path", ""))
                if not artifact_path.is_file() or _sha256_file(artifact_path) != ref["sha256"] or artifact_path.stat().st_size != ref["size_bytes"]:
                    errors.append(f"{viewer_id}: screenshot artifact mismatch")
    return errors


def main(argv: list[str] | None = None) -> int:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="OfficeCLI release gate")
    sub = parser.add_subparsers(dest="command", required=True)
    static = sub.add_parser("static")
    static.add_argument("--root", type=Path, default=Path.cwd())
    platform_gate = sub.add_parser("platform")
    platform_gate.add_argument("--evidence-root", required=True, type=Path)
    platform_gate.add_argument("--lock", required=True, type=Path)
    platform_gate.add_argument("--capability", required=True, type=Path)
    native_toc_gate = sub.add_parser("native-toc")
    native_toc_gate.add_argument("--evidence-root", required=True, type=Path)
    native_toc_gate.add_argument("--lock", required=True, type=Path)
    args = parser.parse_args(argv)
    if args.command == "static":
        errors = scan_production_paths(args.root.resolve())
    elif args.command == "platform":
        errors = validate_platform_evidence(args.evidence_root, args.lock, args.capability)
    else:
        errors = validate_native_toc_evidence(args.evidence_root, args.lock)
    sys.stdout.write(json.dumps({"ok": not errors, "errors": errors}, ensure_ascii=False) + "\n")
    return 0 if not errors else 2


if __name__ == "__main__":
    raise SystemExit(main())
