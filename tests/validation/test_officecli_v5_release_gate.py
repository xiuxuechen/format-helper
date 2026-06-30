"""V5-014 发布 Gate 测试。"""

from __future__ import annotations

import hashlib
import json
import tempfile
import unittest
from pathlib import Path

from scripts.officecli.runtime_resolver import EXPECTED_RUNTIME_IDS, load_lock, select_asset
from scripts.officecli.v5_release_gate import scan_production_paths, validate_platform_evidence


ROOT = Path(__file__).resolve().parents[2]
LOCK = ROOT / "tools" / "officecli" / "officecli.lock.json"
CAPABILITY = ROOT / "tools" / "officecli" / "officecli-capability-manifest.json"


class TestOfficeCliV5ReleaseGate(unittest.TestCase):
    """覆盖生产路径扫描与八平台聚合。"""

    def test_repository_production_paths_are_officecli_only(self):
        self.assertEqual(scan_production_paths(ROOT), [])

    def test_eight_platform_evidence_passes_and_missing_platform_blocks(self):
        runner_by_runtime = {
            "win-x64": ("Windows", "AMD64"), "win-arm64": ("Windows", "ARM64"),
            "linux-x64-gnu": ("Linux", "x86_64"), "linux-arm64-gnu": ("Linux", "aarch64"),
            "linux-x64-musl": ("Linux", "x86_64"), "linux-arm64-musl": ("Linux", "aarch64"),
            "osx-x64": ("Darwin", "x86_64"), "osx-arm64": ("Darwin", "arm64"),
        }
        lock = load_lock(LOCK)
        capability_hash = json.loads(CAPABILITY.read_text(encoding="utf-8"))["aggregate_sha256"]
        lock_file_hash = hashlib.sha256(LOCK.read_bytes()).hexdigest()
        capability_file_hash = hashlib.sha256(CAPABILITY.read_bytes()).hexdigest()
        with tempfile.TemporaryDirectory() as tmp:
            evidence_root = Path(tmp)
            for runtime_id in EXPECTED_RUNTIME_IDS:
                asset = select_asset(lock, runtime_id)
                runtime_dir = evidence_root / runtime_id
                runtime_dir.mkdir()
                commands = []
                command_by_name = {
                    "version": ["officecli", "--version"],
                    "create": ["officecli", "create", "smoke.docx", "--force"],
                    "add": ["officecli", "add", "smoke.docx", "/body", "--type", "paragraph", "--json"],
                    "get": ["officecli", "get", "smoke.docx", "/body/p[1]", "--json"],
                    "set": ["officecli", "set", "smoke.docx", "/body/p[1]", "--prop", "alignment=center", "--json"],
                    "validate": ["officecli", "validate", "smoke.docx", "--json"],
                    "screenshot": ["officecli", "view", "smoke.docx", "screenshot", "-o", "smoke.png", "--page", "1"],
                }
                for index, name in enumerate(["version", "create", "add", "get", "set", "validate", "screenshot"], start=1):
                    stdout_path = runtime_dir / f"{index:02d}-{name}.stdout.txt"
                    stderr_path = runtime_dir / f"{index:02d}-{name}.stderr.txt"
                    stdout_path.write_text("1.0.113" if name == "version" else "ok", encoding="utf-8")
                    stderr_path.write_text("", encoding="utf-8")
                    commands.append({
                        "name": name, "exit_code": 0, "business_success": True,
                        "command": command_by_name[name],
                        "stdout": {"path": stdout_path.name, "sha256": hashlib.sha256(stdout_path.read_bytes()).hexdigest(), "size_bytes": stdout_path.stat().st_size},
                        "stderr": {"path": stderr_path.name, "sha256": hashlib.sha256(stderr_path.read_bytes()).hexdigest(), "size_bytes": stderr_path.stat().st_size},
                    })
                smoke_path = runtime_dir / "smoke.docx"
                screenshot_path = runtime_dir / "smoke.png"
                smoke_path.write_bytes(b"docx")
                screenshot_path.write_bytes(b"png")
                payload = {
                    "runtime_id": runtime_id,
                    "status": "passed",
                    "resolution": {
                        "runtime_id": runtime_id,
                        "officecli_version": "1.0.113",
                        "version": "1.0.113",
                        "sha256": asset["sha256"],
                        "size_bytes": asset["size_bytes"],
                    },
                    "capability_aggregate_sha256": capability_hash,
                    "lock_sha256": lock_file_hash,
                    "capability_file_sha256": capability_file_hash,
                    "runner": {"system": runner_by_runtime[runtime_id][0], "release": "test", "machine": runner_by_runtime[runtime_id][1], "python": "3.12", "libc": ("musl" if runtime_id.endswith("-musl") else "glibc") if runtime_id.startswith("linux-") else "not_applicable", "distribution_id": "alpine" if runtime_id.endswith("-musl") else ("ubuntu" if runtime_id.startswith("linux-") else "not_applicable"), "distribution_version": "3.20" if runtime_id.endswith("-musl") else ("24.04" if runtime_id.startswith("linux-") else "not_applicable")},
                    "environment": {"OFFICECLI_SKIP_UPDATE": "1", "OFFICECLI_NO_AUTO_RESIDENT": "1"},
                    "commands": commands,
                    "smoke_docx": {"path": smoke_path.name, "sha256": hashlib.sha256(smoke_path.read_bytes()).hexdigest(), "size_bytes": smoke_path.stat().st_size},
                    "screenshot": {"path": screenshot_path.name, "sha256": hashlib.sha256(screenshot_path.read_bytes()).hexdigest(), "size_bytes": screenshot_path.stat().st_size},
                }
                (runtime_dir / f"{runtime_id}.platform-evidence.json").write_text(json.dumps(payload), encoding="utf-8")
            self.assertEqual(validate_platform_evidence(evidence_root, LOCK, CAPABILITY), [])
            win_x64_path = evidence_root / "win-x64" / "win-x64.platform-evidence.json"
            win_x64_payload = json.loads(win_x64_path.read_text(encoding="utf-8"))
            win_x64_payload["commands"][1]["command"] = ["officecli", "create", "smoke.docx"]
            win_x64_path.write_text(json.dumps(win_x64_payload), encoding="utf-8")
            errors = validate_platform_evidence(evidence_root, LOCK, CAPABILITY)
            self.assertTrue(any("create smoke command must include --force" in error for error in errors))
            win_x64_payload["commands"][1]["command"] = ["officecli", "--force", "create", "smoke.docx"]
            win_x64_path.write_text(json.dumps(win_x64_payload), encoding="utf-8")
            errors = validate_platform_evidence(evidence_root, LOCK, CAPABILITY)
            self.assertTrue(any("create smoke command must include --force" in error for error in errors))
            win_x64_payload["commands"][1]["command"] = ["officecli", "create", "smoke.docx", "--force"]
            win_x64_path.write_text(json.dumps(win_x64_payload), encoding="utf-8")
            (evidence_root / "win-arm64" / "win-arm64.platform-evidence.json").unlink()
            errors = validate_platform_evidence(evidence_root, LOCK, CAPABILITY)
            self.assertTrue(any("win-arm64" in error for error in errors))


if __name__ == "__main__":
    unittest.main()
