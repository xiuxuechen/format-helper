"""V5-014 平台证据生成测试。"""

from __future__ import annotations

import json
import subprocess
import tempfile
import unittest
from pathlib import Path

from scripts.officecli.platform_evidence import collect_platform_evidence


class TestOfficeCliPlatformEvidence(unittest.TestCase):
    """验证平台证据成功与失败闭环。"""

    def test_collects_required_smoke_evidence(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            lock = root / "lock.json"
            capability = root / "capability.json"
            executable = root / "officecli.exe"
            lock.write_text("{}", encoding="utf-8")
            capability.write_text(json.dumps({"aggregate_sha256": "a" * 64}), encoding="utf-8")
            executable.write_bytes(b"binary")

            def ensure_func(**_kwargs):
                return {
                    "runtime_id": "win-x64", "officecli_version": "1.0.113",
                    "executable_path": str(executable), "sha256": "b" * 64,
                    "size_bytes": executable.stat().st_size, "version": "1.0.113",
                }

            seen_commands = []

            def runner(command, _env, _timeout):
                seen_commands.append(command)
                if "create" in command:
                    Path(command[2]).write_bytes(b"docx")
                if "screenshot" in command:
                    Path(command[command.index("-o") + 1]).write_bytes(b"png")
                stdout = "1.0.113\n" if command[1] == "--version" else '{"success":true}\n'
                return subprocess.CompletedProcess(command, 0, stdout=stdout, stderr="")

            evidence = collect_platform_evidence(
                workspace_root=root, lock_path=lock, capability_path=capability,
                runtime_id="win-x64", output_dir=root / "evidence",
                ensure_func=ensure_func, command_runner=runner,
            )
            self.assertEqual(evidence["status"], "passed")
            self.assertEqual([item["name"] for item in evidence["commands"]], ["version", "create", "add", "get", "set", "validate", "screenshot"])
            create_command = next(command for command in seen_commands if command[1] == "create")
            self.assertIn("--force", create_command)
            self.assertTrue((root / "evidence" / "win-x64.platform-evidence.json").exists())

    def test_failed_command_blocks_evidence(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            lock = root / "lock.json"
            capability = root / "capability.json"
            executable = root / "officecli"
            lock.write_text("{}", encoding="utf-8")
            capability.write_text(json.dumps({"aggregate_sha256": "a" * 64}), encoding="utf-8")
            executable.write_bytes(b"binary")

            def ensure_func(**_kwargs):
                return {"runtime_id": "linux-x64-gnu", "executable_path": str(executable)}

            def runner(command, _env, _timeout):
                if command[1] == "--version":
                    return subprocess.CompletedProcess(command, 0, stdout="1.0.113\n", stderr="")
                return subprocess.CompletedProcess(command, 1, stdout="", stderr="failed")

            with self.assertRaisesRegex(RuntimeError, "create"):
                collect_platform_evidence(
                    workspace_root=root, lock_path=lock, capability_path=capability,
                    runtime_id="linux-x64-gnu", output_dir=root / "evidence",
                    ensure_func=ensure_func, command_runner=runner,
                )

    def test_exit_zero_success_false_is_blocked(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            lock = root / "lock.json"
            capability = root / "capability.json"
            executable = root / "officecli"
            lock.write_text("{}", encoding="utf-8")
            capability.write_text(json.dumps({"aggregate_sha256": "a" * 64}), encoding="utf-8")
            executable.write_bytes(b"binary")

            def runner(command, _env, _timeout):
                if command[1] == "--version":
                    return subprocess.CompletedProcess(command, 0, stdout="1.0.113\n", stderr="")
                if command[1] == "create":
                    Path(command[2]).write_bytes(b"docx")
                    return subprocess.CompletedProcess(command, 0, stdout="created\n", stderr="")
                return subprocess.CompletedProcess(command, 0, stdout='{"success":false}\n', stderr="")

            with self.assertRaisesRegex(RuntimeError, "add"):
                collect_platform_evidence(
                    workspace_root=root, lock_path=lock, capability_path=capability,
                    runtime_id="linux-x64-gnu", output_dir=root / "evidence",
                    ensure_func=lambda **_kwargs: {"executable_path": str(executable)},
                    command_runner=runner,
                )


if __name__ == "__main__":
    unittest.main()
