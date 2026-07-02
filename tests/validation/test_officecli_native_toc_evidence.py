"""V5-014 原生 TOC evidence harness 测试。"""

from __future__ import annotations

import subprocess
import tempfile
import unittest
from pathlib import Path

from scripts.officecli.native_toc_evidence import _officecli_command_succeeded, collect_native_toc_evidence


class TestNativeTocEvidence(unittest.TestCase):
    def test_passed_native_refresh_generates_evidence(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            executable = root / "officecli.exe"
            executable.write_bytes(b"binary")
            lock = root / "lock.json"
            lock.write_text("{}", encoding="utf-8")

            def ensure_func(**_kwargs):
                return {"runtime_id": "win-x64", "officecli_version": "1.0.113", "executable_path": str(executable)}

            def runner(command, _env, _timeout):
                if command[1] == "create":
                    Path(command[2]).write_bytes(b"docx")
                if "screenshot" in command:
                    Path(command[command.index("-o") + 1]).write_bytes(b"png")
                return subprocess.CompletedProcess(command, 0, stdout="{}", stderr="")

            def refresh(input_docx, output_docx, viewer, officecli_executable):
                output_docx.write_bytes(input_docx.read_bytes() + b"-refreshed")
                return {
                    "status": "passed", "viewer": viewer["viewer"],
                    "gate_check": {"status": "passed"}, "page_count": 1,
                    "evidence_refs": [],
                    "visible_entries": [{"level": 1, "text": "第一章", "page_number": 1}],
                }

            evidence = collect_native_toc_evidence(
                workspace_root=root, lock_path=lock, run_dir=root / "run",
                required_viewer="word",
                ensure_func=ensure_func, command_runner=runner,
                viewer_probe=lambda: {"ok": True, "viewer": "word", "version": "test"},
                refresh_func=refresh,
            )
            self.assertEqual(evidence["status"], "passed")
            self.assertTrue((root / "run" / "logs" / "toc_acceptance.json").exists())
            self.assertEqual(len(evidence["page_screenshots"]), 1)

    def test_missing_viewer_blocks(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            executable = root / "officecli.exe"
            executable.write_bytes(b"binary")
            lock = root / "lock.json"
            lock.write_text("{}", encoding="utf-8")

            def runner(command, _env, _timeout):
                if command[1] == "create":
                    Path(command[2]).write_bytes(b"docx")
                if "screenshot" in command:
                    Path(command[command.index("-o") + 1]).write_bytes(b"png")
                return subprocess.CompletedProcess(command, 0, stdout="{}", stderr="")

            with self.assertRaisesRegex(RuntimeError, "viewer"):
                collect_native_toc_evidence(
                    workspace_root=root, lock_path=lock, run_dir=root / "run",
                    required_viewer="word",
                    ensure_func=lambda **_kwargs: {"executable_path": str(executable)},
                    command_runner=runner, viewer_probe=lambda: {"ok": False, "reason": "missing"},
                )

    def test_wrong_viewer_blocks(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            executable = root / "officecli.exe"
            executable.write_bytes(b"binary")
            lock = root / "lock.json"
            lock.write_text("{}", encoding="utf-8")

            def runner(command, _env, _timeout):
                if command[1] == "create":
                    Path(command[2]).write_bytes(b"docx")
                if "screenshot" in command:
                    Path(command[command.index("-o") + 1]).write_bytes(b"png")
                return subprocess.CompletedProcess(command, 0, stdout="{}", stderr="")

            with self.assertRaisesRegex(RuntimeError, "viewer 不匹配"):
                collect_native_toc_evidence(
                    workspace_root=root, lock_path=lock, run_dir=root / "run",
                    required_viewer="wps",
                    ensure_func=lambda **_kwargs: {"executable_path": str(executable)},
                    command_runner=runner,
                    viewer_probe=lambda: {"ok": True, "viewer": "word", "version": "test"},
                )

    def test_product_viewer_names_match_ci_ids(self):
        from scripts.officecli.native_toc_evidence import _viewer_id
        self.assertEqual("word", _viewer_id("Microsoft Word"))
        self.assertEqual("wps", _viewer_id("WPS Writer"))

    def test_advisory_exit_with_json_success_is_accepted(self):
        proc = subprocess.CompletedProcess(
            ["officecli", "add"], 2,
            stdout='{"success": true, "warnings": [{"code": "advisory"}]}',
            stderr="",
        )
        self.assertTrue(_officecli_command_succeeded(proc))

    def test_nonzero_success_without_warning_is_rejected(self):
        proc = subprocess.CompletedProcess(
            ["officecli", "add"], 2,
            stdout='{"success": true}',
            stderr="",
        )
        self.assertFalse(_officecli_command_succeeded(proc))

    def test_nonzero_failed_json_is_rejected(self):
        proc = subprocess.CompletedProcess(
            ["officecli", "add"], 2,
            stdout='{"success": false, "warnings": [{"code": "advisory"}]}',
            stderr="",
        )
        self.assertFalse(_officecli_command_succeeded(proc))

    def test_nonzero_plain_text_is_rejected(self):
        proc = subprocess.CompletedProcess(
            ["officecli", "add"], 2,
            stdout="not json",
            stderr="",
        )
        self.assertFalse(_officecli_command_succeeded(proc))

    def test_required_viewer_is_forwarded_to_probe(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            executable = root / "officecli.exe"
            executable.write_bytes(b"binary")
            lock = root / "lock.json"
            lock.write_text("{}", encoding="utf-8")
            seen: list[str] = []

            def runner(command, _env, _timeout):
                if command[1] == "create":
                    Path(command[2]).write_bytes(b"docx")
                if "screenshot" in command:
                    Path(command[command.index("-o") + 1]).write_bytes(b"png")
                return subprocess.CompletedProcess(command, 0, stdout="{}", stderr="")

            def refresh(input_docx, output_docx, viewer, officecli_executable):
                output_docx.write_bytes(input_docx.read_bytes() + b"-refreshed")
                return {
                    "status": "passed", "viewer": viewer["viewer"],
                    "gate_check": {"status": "passed"}, "page_count": 1,
                    "evidence_refs": [],
                    "visible_entries": [{"level": 1, "text": "第一章", "page_number": 1}],
                }

            collect_native_toc_evidence(
                workspace_root=root, lock_path=lock, run_dir=root / "run",
                required_viewer="wps",
                ensure_func=lambda **_kwargs: {"executable_path": str(executable)},
                command_runner=runner,
                viewer_probe=lambda required_viewer: seen.append(required_viewer) or {"ok": True, "viewer": "wps", "version": "test"},
                refresh_func=refresh,
            )
            self.assertEqual(["wps"], seen)


if __name__ == "__main__":
    unittest.main()
