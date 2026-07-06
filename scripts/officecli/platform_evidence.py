"""生成 OfficeCLI 平台运行证据。"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.officecli.runtime_resolver import ensure_officecli


CommandRunner = Callable[[list[str], dict[str, str], int], subprocess.CompletedProcess[str]]


def sha256_file(path: Path) -> str:
    """计算文件 SHA-256。"""
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sha256_utf8_lf_file(path: Path) -> str:
    """计算 UTF-8 文本在 LF 换行规范化后的 SHA-256。"""
    text = path.read_text(encoding="utf-8")
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def _run(command: list[str], env: dict[str, str], timeout_seconds: int) -> subprocess.CompletedProcess[str]:
    """执行单条 OfficeCLI 命令并捕获文本输出。"""
    return subprocess.run(
        command,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=timeout_seconds,
        check=False,
        env=env,
    )


def detect_libc() -> str:
    """返回可审计的 libc 分类。"""
    if platform.system().lower() != "linux":
        return "not_applicable"
    if Path("/etc/alpine-release").is_file():
        return "musl"
    name, _version = platform.libc_ver()
    normalized = name.lower()
    if normalized in {"glibc", "gnu libc"}:
        return "glibc"
    if "musl" in normalized:
        return "musl"
    return "unknown"


def detect_linux_distribution() -> tuple[str, str]:
    """读取 Linux 发行版证据；非 Linux 返回 not_applicable。"""
    if platform.system().lower() != "linux":
        return "not_applicable", "not_applicable"
    values: dict[str, str] = {}
    os_release = Path("/etc/os-release")
    if os_release.is_file():
        for line in os_release.read_text(encoding="utf-8", errors="replace").splitlines():
            if "=" in line:
                key, value = line.split("=", 1)
                values[key] = value.strip().strip('"')
    distro_id = values.get("ID", "unknown").lower()
    version = values.get("VERSION_ID", "")
    alpine_release = Path("/etc/alpine-release")
    if alpine_release.is_file():
        distro_id = "alpine"
        version = alpine_release.read_text(encoding="utf-8", errors="replace").strip()
    return distro_id, version or "unknown"


def _write_command_artifacts(output_dir: Path, index: int, name: str, proc: subprocess.CompletedProcess[str]) -> dict[str, Any]:
    """写出命令 stdout/stderr 并返回可审计摘要。"""
    stdout_path = output_dir / f"{index:02d}-{name}.stdout.txt"
    stderr_path = output_dir / f"{index:02d}-{name}.stderr.txt"
    stdout_path.write_text(proc.stdout or "", encoding="utf-8")
    stderr_path.write_text(proc.stderr or "", encoding="utf-8")
    return {
        "name": name,
        "exit_code": proc.returncode,
        "stdout": {
            "path": stdout_path.name,
            "sha256": sha256_file(stdout_path),
            "size_bytes": stdout_path.stat().st_size,
        },
        "stderr": {
            "path": stderr_path.name,
            "sha256": sha256_file(stderr_path),
            "size_bytes": stderr_path.stat().st_size,
        },
    }


def collect_platform_evidence(
    *,
    workspace_root: Path,
    lock_path: Path,
    capability_path: Path,
    runtime_id: str,
    output_dir: Path,
    offline: bool = False,
    ensure_func: Callable[..., dict[str, Any]] = ensure_officecli,
    command_runner: CommandRunner = _run,
) -> dict[str, Any]:
    """解析固定 runtime 并执行最小 DOCX create/get/set/validate/screenshot。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    resolution = ensure_func(
        lock_path=lock_path,
        workspace_root=workspace_root,
        runtime_id=runtime_id,
        offline=offline,
        skip_version_check=False,
    )
    executable = Path(resolution["executable_path"])
    smoke_docx = output_dir / "officecli-smoke.docx"
    screenshot = output_dir / "officecli-smoke.png"
    env = os.environ.copy()
    env.update({
        "OFFICECLI_SKIP_UPDATE": "1",
        "OFFICECLI_NO_AUTO_RESIDENT": "1",
        "LC_ALL": "C.UTF-8",
        "TZ": "UTC",
    })
    commands = [
        ("version", [str(executable), "--version"]),
        ("create", [str(executable), "create", str(smoke_docx), "--force"]),
        ("add", [str(executable), "add", str(smoke_docx), "/body", "--type", "paragraph", "--prop", "text=OfficeCLI smoke", "--json"]),
        ("get", [str(executable), "get", str(smoke_docx), "/body/p[1]", "--depth", "2", "--json"]),
        ("set", [str(executable), "set", str(smoke_docx), "/body/p[1]", "--prop", "alignment=center", "--json"]),
        ("validate", [str(executable), "validate", str(smoke_docx), "--json"]),
        ("screenshot", [str(executable), "view", str(smoke_docx), "screenshot", "-o", str(screenshot), "--page", "1"]),
    ]
    command_results: list[dict[str, Any]] = []
    for index, (name, command) in enumerate(commands, start=1):
        proc = command_runner(command, env, 120)
        record = _write_command_artifacts(output_dir, index, name, proc)
        record["command"] = command
        if name in {"add", "get", "set", "validate"}:
            try:
                payload = json.loads((proc.stdout or "").strip())
                record["business_success"] = isinstance(payload, dict) and payload.get("success") is True
            except json.JSONDecodeError:
                record["business_success"] = False
        else:
            record["business_success"] = bool((proc.stdout or "").strip())
        if name == "version":
            record["business_success"] = (proc.stdout or "").strip().lstrip("v") == "1.0.113"
        command_results.append(record)
        if proc.returncode != 0 or record["business_success"] is not True:
            raise RuntimeError(f"OfficeCLI 平台证据命令失败：{name}, exit={proc.returncode}")
    if not smoke_docx.is_file() or smoke_docx.stat().st_size == 0:
        raise RuntimeError("OfficeCLI smoke DOCX 未生成")
    if not screenshot.is_file() or screenshot.stat().st_size == 0:
        raise RuntimeError("OfficeCLI screenshot 未生成")
    capability = json.loads(capability_path.read_text(encoding="utf-8"))
    distro_id, distro_version = detect_linux_distribution()
    evidence = {
        "schema_id": "officecli-platform-evidence",
        "schema_version": "1.0.0",
        "runtime_id": runtime_id,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "runner": {
            "system": platform.system(),
            "release": platform.release(),
            "machine": platform.machine(),
            "python": platform.python_version(),
            "libc": detect_libc(),
            "distribution_id": distro_id,
            "distribution_version": distro_version,
        },
        "resolution": resolution,
        "lock_sha256": sha256_utf8_lf_file(lock_path),
        "capability_file_sha256": sha256_utf8_lf_file(capability_path),
        "capability_aggregate_sha256": capability.get("aggregate_sha256"),
        "environment": {
            "OFFICECLI_SKIP_UPDATE": "1",
            "OFFICECLI_NO_AUTO_RESIDENT": "1",
            "locale": "C.UTF-8",
            "timezone": "UTC",
        },
        "commands": command_results,
        "smoke_docx": {
            "path": smoke_docx.name,
            "sha256": sha256_file(smoke_docx),
            "size_bytes": smoke_docx.stat().st_size,
        },
        "screenshot": {
            "path": screenshot.name,
            "sha256": sha256_file(screenshot),
            "size_bytes": screenshot.stat().st_size,
        },
        "status": "passed",
    }
    evidence_path = output_dir / f"{runtime_id}.platform-evidence.json"
    evidence_path.write_text(json.dumps(evidence, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return evidence


def main(argv: list[str] | None = None) -> int:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="生成 OfficeCLI 平台证据")
    parser.add_argument("--workspace-root", type=Path, default=Path.cwd())
    parser.add_argument("--lock", type=Path, default=Path("tools/officecli/officecli.lock.json"))
    parser.add_argument("--capability", type=Path, default=Path("tools/officecli/officecli-capability-manifest.json"))
    parser.add_argument("--runtime-id", required=True)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--offline", action="store_true")
    args = parser.parse_args(argv)
    try:
        evidence = collect_platform_evidence(
            workspace_root=args.workspace_root.resolve(),
            lock_path=args.lock.resolve(),
            capability_path=args.capability.resolve(),
            runtime_id=args.runtime_id,
            output_dir=args.output_dir.resolve(),
            offline=args.offline,
        )
        sys.stdout.write(json.dumps({"ok": True, "runtime_id": evidence["runtime_id"], "status": evidence["status"]}) + "\n")
        return 0
    except Exception as exc:
        sys.stderr.write(f"平台证据失败：{exc}\n")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
