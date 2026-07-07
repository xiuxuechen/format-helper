"""OFFICECLI-014 发布 Gate 测试。"""

from __future__ import annotations

import hashlib
import json
import tempfile
import unittest
from pathlib import Path

from scripts.officecli.runtime_resolver import load_lock, select_asset
from scripts.officecli.release_gate import (
    REQUIRED_RELEASE_RUNTIME_IDS,
    _sha256_utf8_lf_file,
    scan_production_paths,
    validate_native_toc_evidence,
    validate_platform_evidence,
)


ROOT = Path(__file__).resolve().parents[2]
LOCK = ROOT / "tools" / "officecli" / "officecli.lock.json"
CAPABILITY = ROOT / "tools" / "officecli" / "officecli-capability-manifest.json"


class TestOfficeCliReleaseGate(unittest.TestCase):
    """覆盖生产路径扫描与 Windows/Apple Silicon Mac 必过平台聚合。"""

    def test_repository_production_paths_are_officecli_only(self):
        self.assertEqual(scan_production_paths(ROOT), [])

    def test_required_win_mac_platform_evidence_passes_and_missing_required_blocks(self):
        runner_by_runtime = {
            "win-x64": ("Windows", "AMD64"), "win-arm64": ("Windows", "ARM64"),
            "linux-x64-gnu": ("Linux", "x86_64"), "linux-arm64-gnu": ("Linux", "aarch64"),
            "linux-x64-musl": ("Linux", "x86_64"), "linux-arm64-musl": ("Linux", "aarch64"),
            "osx-x64": ("Darwin", "x86_64"), "osx-arm64": ("Darwin", "arm64"),
        }
        lock = load_lock(LOCK)
        capability_hash = json.loads(CAPABILITY.read_text(encoding="utf-8"))["aggregate_sha256"]
        lock_file_hash = _sha256_utf8_lf_file(LOCK)
        capability_file_hash = _sha256_utf8_lf_file(CAPABILITY)
        with tempfile.TemporaryDirectory() as tmp:
            evidence_root = Path(tmp)
            for runtime_id in REQUIRED_RELEASE_RUNTIME_IDS | {"linux-x64-gnu", "osx-x64"}:
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
            linux_path = evidence_root / "linux-x64-gnu" / "linux-x64-gnu.platform-evidence.json"
            linux_payload = json.loads(linux_path.read_text(encoding="utf-8"))
            linux_payload["status"] = "failed"
            linux_payload["resolution"]["sha256"] = "0" * 64
            linux_path.write_text(json.dumps(linux_payload), encoding="utf-8")
            osx_x64_path = evidence_root / "osx-x64" / "osx-x64.platform-evidence.json"
            osx_x64_payload = json.loads(osx_x64_path.read_text(encoding="utf-8"))
            osx_x64_payload["status"] = "failed"
            osx_x64_payload["resolution"]["sha256"] = "0" * 64
            osx_x64_path.write_text(json.dumps(osx_x64_payload), encoding="utf-8")
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
            (evidence_root / "osx-arm64" / "osx-arm64.platform-evidence.json").unlink()
            errors = validate_platform_evidence(evidence_root, LOCK, CAPABILITY)
            self.assertTrue(any("osx-arm64" in error for error in errors))
            unknown_dir = evidence_root / "unknown-runtime"
            unknown_dir.mkdir()
            (unknown_dir / "unknown-runtime.platform-evidence.json").write_text(
                json.dumps({"runtime_id": "unknown-runtime"}),
                encoding="utf-8",
            )
            errors = validate_platform_evidence(evidence_root, LOCK, CAPABILITY)
            self.assertTrue(any("存在未知平台证据：unknown-runtime" in error for error in errors))

    def test_text_hashes_ignore_checkout_newline_differences(self):
        """Windows checkout 的 CRLF 不应导致 lock/capability 文本 hash 漂移。"""
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            crlf_lock = tmp_path / "officecli.lock.json"
            crlf_capability = tmp_path / "officecli-capability-manifest.json"
            crlf_lock.write_text(LOCK.read_text(encoding="utf-8").replace("\n", "\r\n"), encoding="utf-8", newline="")
            crlf_capability.write_text(CAPABILITY.read_text(encoding="utf-8").replace("\n", "\r\n"), encoding="utf-8", newline="")
            self.assertEqual(_sha256_utf8_lf_file(LOCK), _sha256_utf8_lf_file(crlf_lock))
            self.assertEqual(_sha256_utf8_lf_file(CAPABILITY), _sha256_utf8_lf_file(crlf_capability))

    def test_native_toc_evidence_accepts_word_and_wps_local_bundle(self):
        """本机 Word/WPS 真实 evidence 可作为发布候选 native TOC Gate 输入。"""
        lock = load_lock(LOCK)
        asset = select_asset(lock, "win-x64")
        with tempfile.TemporaryDirectory() as tmp:
            evidence_root = Path(tmp)
            for viewer_id, viewer_name in (("word", "Microsoft Word"), ("wps", "WPS Writer")):
                run_dir = evidence_root / viewer_id
                logs_dir = run_dir / "logs"
                image_dir = run_dir / "output" / "_internal"
                logs_dir.mkdir(parents=True)
                image_dir.mkdir(parents=True)
                page_path = image_dir / "toc-page-0001.png"
                page_path.write_bytes(f"{viewer_id}-png".encode("utf-8"))
                page_ref = {
                    "artifact_id": "toc-page-0001",
                    "kind": "png",
                    "relative_path": "output/_internal/toc-page-0001.png",
                    "sha256": hashlib.sha256(page_path.read_bytes()).hexdigest(),
                    "size_bytes": page_path.stat().st_size,
                    "schema_id": None,
                    "schema_version": None,
                }
                acceptance = {
                    "schema_id": "toc-acceptance",
                    "schema_version": "2.0.0",
                    "run_id": viewer_id,
                    "required": True,
                    "status": "passed",
                    "viewer": viewer_name,
                    "viewer_version": "12.0",
                    "platform": "windows",
                    "input_ref": None,
                    "output_ref": None,
                    "before_sha256": "a" * 64,
                    "after_sha256": "b" * 64,
                    "field_update_count": 1,
                    "toc_update_count": 1,
                    "page_count": 1,
                    "visible_entries": [{"level": 1, "text": "第一章\t1", "page_number": 1}],
                    "evidence_refs": [page_ref],
                    "gate_check": {
                        "gate_id": "toc-acceptance-officecli",
                        "status": "passed",
                        "checked_at": "2026-01-01T00:00:00Z",
                        "predicate_version": "1.0.0",
                        "evidence_refs": [],
                        "failed_codes": [],
                    },
                    "error": {"code": "NONE", "reason_code": "none", "message": "", "retryable": False, "viewer": None},
                }
                payload = {
                    "schema_id": "officecli-native-toc-evidence",
                    "schema_version": "1.0.0",
                    "status": "passed",
                    "resolution": {
                        "runtime_id": "win-x64",
                        "officecli_version": "1.0.113",
                        "version": "1.0.113",
                        "sha256": asset["sha256"],
                        "size_bytes": asset["size_bytes"],
                    },
                    "viewer": {"ok": True, "viewer": viewer_name, "version": "12.0"},
                    "toc_acceptance_path": "logs/toc_acceptance.json",
                    "toc_acceptance": acceptance,
                    "page_screenshots": [page_ref],
                }
                (logs_dir / "toc_acceptance.json").write_text(json.dumps(acceptance), encoding="utf-8")
                (logs_dir / "native_toc.platform-evidence.json").write_text(json.dumps(payload), encoding="utf-8")
            self.assertEqual(validate_native_toc_evidence(evidence_root, LOCK), [])
            word_toc_path = evidence_root / "word" / "logs" / "toc_acceptance.json"
            word_acceptance = json.loads(word_toc_path.read_text(encoding="utf-8"))
            word_acceptance["field_update_count"] = 0
            word_toc_path.write_text(json.dumps(word_acceptance), encoding="utf-8")
            errors = validate_native_toc_evidence(evidence_root, LOCK)
            self.assertTrue(any("embedded toc_acceptance does not match" in error for error in errors))
            word_payload_path = evidence_root / "word" / "logs" / "native_toc.platform-evidence.json"
            word_payload = json.loads(word_payload_path.read_text(encoding="utf-8"))
            word_payload["toc_acceptance"] = word_acceptance
            word_payload_path.write_text(json.dumps(word_payload), encoding="utf-8")
            errors = validate_native_toc_evidence(evidence_root, LOCK)
            self.assertTrue(any("field_update_count must be positive" in error for error in errors))
            word_acceptance["field_update_count"] = 1
            word_toc_path.write_text(json.dumps(word_acceptance), encoding="utf-8")
            word_payload["toc_acceptance"] = word_acceptance
            word_payload_path.write_text(json.dumps(word_payload), encoding="utf-8")
            (evidence_root / "wps" / "logs" / "native_toc.platform-evidence.json").unlink()
            errors = validate_native_toc_evidence(evidence_root, LOCK)
            self.assertTrue(any("缺少 native TOC 证据：wps" in error for error in errors))

    def test_native_toc_evidence_rejects_broken_screenshot_hash(self):
        """native TOC Gate 必须复算截图 hash，不能只信 JSON。"""
        lock = load_lock(LOCK)
        asset = select_asset(lock, "win-x64")
        with tempfile.TemporaryDirectory() as tmp:
            evidence_root = Path(tmp)
            for viewer_id, viewer_name in (("word", "Microsoft Word"), ("wps", "WPS Writer")):
                run_dir = evidence_root / viewer_id
                logs_dir = run_dir / "logs"
                image_dir = run_dir / "output" / "_internal"
                logs_dir.mkdir(parents=True)
                image_dir.mkdir(parents=True)
                page_path = image_dir / "toc-page-0001.png"
                page_path.write_bytes(b"png")
                page_ref = {
                    "artifact_id": "toc-page-0001",
                    "kind": "png",
                    "relative_path": "output/_internal/toc-page-0001.png",
                    "sha256": hashlib.sha256(page_path.read_bytes()).hexdigest(),
                    "size_bytes": page_path.stat().st_size,
                    "schema_id": None,
                    "schema_version": None,
                }
                acceptance = {
                    "schema_id": "toc-acceptance",
                    "schema_version": "2.0.0",
                    "run_id": viewer_id,
                    "required": True,
                    "status": "passed",
                    "viewer": viewer_name,
                    "viewer_version": "12.0",
                    "platform": "windows",
                    "input_ref": None,
                    "output_ref": None,
                    "before_sha256": "a" * 64,
                    "after_sha256": "b" * 64,
                    "field_update_count": 1,
                    "toc_update_count": 1,
                    "page_count": 1,
                    "visible_entries": [{"level": 1, "text": "第一章\t1", "page_number": 1}],
                    "evidence_refs": [page_ref],
                    "error": {"code": "NONE", "reason_code": "none", "message": "", "retryable": False, "viewer": None},
                    "gate_check": {
                        "gate_id": "toc-acceptance-officecli",
                        "status": "passed",
                        "checked_at": "2026-01-01T00:00:00Z",
                        "predicate_version": "1.0.0",
                        "evidence_refs": [],
                        "failed_codes": [],
                    },
                }
                payload = {
                    "schema_id": "officecli-native-toc-evidence",
                    "schema_version": "1.0.0",
                    "status": "passed",
                    "resolution": {
                        "runtime_id": "win-x64",
                        "officecli_version": "1.0.113",
                        "version": "1.0.113",
                        "sha256": asset["sha256"],
                        "size_bytes": asset["size_bytes"],
                    },
                    "viewer": {"ok": True, "viewer": viewer_name, "version": "12.0"},
                    "toc_acceptance_path": "logs/toc_acceptance.json",
                    "toc_acceptance": acceptance,
                    "page_screenshots": [dict(page_ref)],
                }
                if viewer_id == "word":
                    payload["page_screenshots"][0]["sha256"] = "0" * 64
                (logs_dir / "toc_acceptance.json").write_text(json.dumps(acceptance), encoding="utf-8")
                (logs_dir / "native_toc.platform-evidence.json").write_text(json.dumps(payload), encoding="utf-8")
            errors = validate_native_toc_evidence(evidence_root, LOCK)
            self.assertTrue(any("word: screenshot artifact mismatch" in error for error in errors))


if __name__ == "__main__":
    unittest.main()
