"""V5-010 TOC refresh adapter 验证。"""

from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path
from typing import Any

from jsonschema import Draft202012Validator, FormatChecker

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

TOC_SCHEMA_PATH = ROOT / "docs" / "v5" / "schemas" / "toc-acceptance.schema.json"


def _load_adapter():
    import importlib.util
    path = ROOT / "scripts" / "officecli" / "toc_refresh_adapter.py"
    spec = importlib.util.spec_from_file_location("v5_toc_adapter", path)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(module)
    return module


def _validate_schema(instance: dict[str, Any]) -> list[str]:
    schema = json.loads(TOC_SCHEMA_PATH.read_text(encoding="utf-8"))
    validator = Draft202012Validator(schema, format_checker=FormatChecker())
    return [error.message for error in sorted(validator.iter_errors(instance), key=lambda error: list(error.path))]


class ProbeViewerTest(unittest.TestCase):
    def test_probe_function_exists(self):
        mod = _load_adapter()
        self.assertTrue(callable(mod.probe_viewer))

    def test_probe_returns_dict(self):
        mod = _load_adapter()
        result = mod.probe_viewer()
        self.assertIsInstance(result, dict)
        self.assertIn("ok", result)

    def test_is_windows_detected(self):
        mod = _load_adapter()
        self.assertIsInstance(mod.is_windows(), bool)

    def test_validate_result_requires_clean_data(self):
        mod = _load_adapter()
        self.assertTrue(mod.validate_result_is_clean({"success": True, "data": {"valid": True, "errors": []}}))
        self.assertFalse(mod.validate_result_is_clean({"success": True, "data": {"valid": False}}))

    def test_validate_result_allows_native_style_uipriority_metadata_only(self):
        mod = _load_adapter()
        payload = {
            "success": False,
            "data": {
                "valid": False,
                "errors": [{
                    "type": "Schema",
                    "description": "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:uiPriority'.",
                    "path": "/w:styles[1]/w:style[1]",
                    "part": "/word/styles.xml",
                }],
            },
        }
        self.assertTrue(mod.validate_result_is_clean(payload))
        payload["data"]["errors"][0]["description"] = "The element has unexpected child element 'w:bad'."
        self.assertFalse(mod.validate_result_is_clean(payload))

    def test_validate_result_allows_uipriority_warning_shape_only(self):
        mod = _load_adapter()
        payload = {
            "success": False,
            "warnings": [
                {"message": "Found 1 validation error(s):", "code": "warning"},
                {"message": "[Schema] The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:uiPriority'.", "code": "warning"},
                {"message": "Path: /w:styles[1]/w:style[1]", "code": "warning"},
                {"message": "Part: /word/styles.xml", "code": "warning"},
            ],
        }
        self.assertTrue(mod.validate_result_is_clean(payload))
        payload["warnings"][0]["message"] = "Found 2 validation error(s):"
        self.assertFalse(mod.validate_result_is_clean(payload))
        payload["warnings"][0]["message"] = "Found 1 validation error(s):"
        payload["warnings"].append({"message": "Part: /word/document.xml", "code": "warning"})
        self.assertFalse(mod.validate_result_is_clean(payload))

    def test_fixture_heading_prepare_is_evidence_only(self):
        source = (ROOT / "scripts" / "officecli" / "toc_refresh_adapter.py").read_text(encoding="utf-8")
        self.assertIn('viewer_info.get("native_toc_fixture_prepare_outline") is True', source)
        self.assertIn("prepare_fixture_headings(doc)", source)

    def test_probe_timeout_is_bounded_by_stage_timeout(self):
        mod = _load_adapter()

        def command_factory(_result_path, state_path):
            script = (
                "import json,pathlib,sys,time;"
                "pathlib.Path(sys.argv[1]).write_text(json.dumps({'stage':'probe_word','stage_started_at':time.time()-5,'worker_pid':999}), encoding='utf-8');"
                "time.sleep(30)"
            )
            return [sys.executable, "-c", script, str(state_path)]

        result = mod._run_probe_with_timeout(
            Path("/tmp/toc-probe"),
            stage_timeout_seconds=0.1,
            worker_command_factory=command_factory,
        )
        self.assertFalse(result["ok"])
        self.assertEqual("viewer_busy", result["reason_code"])
        self.assertIn("cleanup_failed=true", result["error"])

    def test_remaining_total_timeout_counts_from_probe_start(self):
        mod = _load_adapter()
        started = mod.time.monotonic() - 5
        remaining = mod._remaining_total_timeout(started, total_timeout_seconds=10)
        self.assertLess(remaining, 6)
        self.assertGreater(remaining, 4)

    def test_probe_viewer_quits_created_instance_on_version_failure(self):
        mod = _load_adapter()

        class BrokenApp:
            def __init__(self):
                self.quit_called = False

            @property
            def Version(self):
                raise RuntimeError("version failed")

            def Quit(self):
                self.quit_called = True

        broken = BrokenApp()

        class FakeClient:
            def DispatchEx(self, _progid):
                return broken

        import types
        fake_win32com = types.SimpleNamespace(client=FakeClient())
        original = sys.modules.get("win32com")
        sys.modules["win32com"] = fake_win32com
        try:
            result = mod.probe_viewer()
        finally:
            if original is None:
                del sys.modules["win32com"]
            else:
                sys.modules["win32com"] = original
        self.assertFalse(result["ok"])
        self.assertTrue(broken.quit_called)


class ReasonCodesTest(unittest.TestCase):
    def test_all_codes_have_messages(self):
        mod = _load_adapter()
        for code in mod.REASON_CODES:
            self.assertIsInstance(mod.REASON_CODES[code], str)
            self.assertTrue(len(mod.REASON_CODES[code]) > 0)

    def test_retryable_codes_are_subset(self):
        mod = _load_adapter()
        for code in mod.RETRYABLE_TOC_CODES:
            self.assertIn(code, mod.REASON_CODES)

    def test_blocked_result_has_all_fields(self):
        mod = _load_adapter()
        result = mod._toc_blocked(Path("/tmp/test.docx"), "test_stage",
                                  reason_code="viewer_unavailable",
                                  message="test message")
        self.assertEqual("blocked", result["status"])
        self.assertEqual("toc-acceptance", result["schema_id"])
        self.assertEqual("2.0.0", result["schema_version"])
        self.assertIsNotNone(result["error"])
        self.assertEqual("viewer_unavailable", result["error"]["reason_code"])
        self.assertEqual("blocked", result["gate_check"]["status"])

    def test_blocked_result_reason_code_enum(self):
        mod = _load_adapter()
        for code in mod.REASON_CODES:
            result = mod._toc_blocked(Path("/tmp/test.docx"), "s",
                                      reason_code=code, message="test")
            self.assertEqual(code, result["error"]["reason_code"])

    def test_readonly_exception_maps_to_readonly_recommended(self):
        mod = _load_adapter()
        for message in (
            "Document is read-only recommended",
            "Document is read only recommended",
            "Document is readonly recommended",
        ):
            self.assertEqual(
                "readonly_recommended",
                mod._classify_com_exception_message(message),
            )


class WorkerTimeoutTest(unittest.TestCase):
    def test_refresh_path_uses_dispatch_ex_for_owned_com_instance(self):
        source = (ROOT / "scripts" / "officecli" / "toc_refresh_adapter.py").read_text(encoding="utf-8")
        self.assertIn("win32com.client.DispatchEx(progid)", source)
        self.assertIn('"Word.Application"', source)
        self.assertIn('"kwps.Application"', source)
        self.assertNotIn("win32com.client.Dispatch(", source)

    def test_worker_result_is_returned(self):
        mod = _load_adapter()

        def command_factory(result_path, _state_path):
            payload = json.dumps({
                "schema_id": "toc-acceptance",
                "schema_version": "2.0.0",
                "run_id": "r1",
                "required": True,
                "status": "passed",
                "viewer": "Microsoft Word",
                "viewer_version": "16.0",
                "platform": "windows",
                "before_sha256": "a" * 64,
                "after_sha256": "b" * 64,
                "field_update_count": 1,
                "toc_update_count": 1,
                "page_count": 1,
                "visible_entries": [{"level": 1, "text": "第一章", "page_number": 1}],
                "evidence_refs": [],
                "error": {"code": "NONE", "reason_code": "none", "message": "", "retryable": False, "viewer": None},
                "gate_check": {
                    "gate_id": "toc-acceptance-v5",
                    "status": "passed",
                    "checked_at": "2026-01-01T00:00:00Z",
                    "predicate_version": "1.0.0",
                    "evidence_refs": [],
                    "failed_codes": [],
                },
            }, ensure_ascii=False)
            script = "import pathlib,sys; pathlib.Path(sys.argv[1]).write_text(sys.argv[2], encoding='utf-8')"
            return [sys.executable, "-c", script, str(result_path), payload]

        result = mod._run_toc_worker_with_timeout(
            Path("/tmp/test.docx"),
            Path("/tmp/out.docx"),
            {"ok": True},
            stage_timeout_seconds=5,
            worker_command_factory=command_factory,
        )
        self.assertEqual("passed", result["status"])

    def test_worker_timeout_terminates_adapter_process(self):
        mod = _load_adapter()

        def command_factory(_result_path, state_path):
            script = (
                "import json,pathlib,sys,time;"
                "pathlib.Path(sys.argv[1]).write_text(json.dumps({'stage':'open_hidden','stage_started_at':time.time()-5,'worker_pid':999}), encoding='utf-8');"
                "time.sleep(30)"
            )
            return [sys.executable, "-c", script, str(state_path)]

        result = mod._run_toc_worker_with_timeout(
            Path("/tmp/test.docx"),
            Path("/tmp/out.docx"),
            {"ok": True},
            stage_timeout_seconds=0.1,
            worker_command_factory=command_factory,
        )
        self.assertEqual("blocked", result["status"])
        self.assertEqual("viewer_busy", result["error"]["reason_code"])
        self.assertIn("worker_pid=", result["error"]["message"])
        self.assertIn("cleanup_failed=", result["error"]["message"])
        self.assertEqual([], _validate_schema(result))

    def test_office_stage_without_application_pid_marks_cleanup_failed(self):
        mod = _load_adapter()

        def command_factory(_result_path, state_path):
            script = (
                "import json,pathlib,sys,time;"
                "pathlib.Path(sys.argv[1]).write_text(json.dumps({'stage':'open_hidden','stage_started_at':time.time()-5,'worker_pid':999}), encoding='utf-8');"
                "time.sleep(30)"
            )
            return [sys.executable, "-c", script, str(state_path)]

        result = mod._run_toc_worker_with_timeout(
            Path("/tmp/test.docx"),
            Path("/tmp/out.docx"),
            {"ok": True},
            stage_timeout_seconds=0.1,
            worker_command_factory=command_factory,
        )
        self.assertIn("application_pid=unknown", result["error"]["message"])
        self.assertIn("cleanup_failed=true", result["error"]["message"])

    def test_worker_state_write_is_readable(self):
        mod = _load_adapter()
        state_path = Path("/tmp/toc-worker-state-test.json")
        mod._write_worker_state(state_path, "open_hidden", application_pid=123)
        state = mod._read_worker_state(state_path)
        self.assertEqual("open_hidden", state["stage"])
        self.assertEqual(123, state["application_pid"])
        state_path.unlink(missing_ok=True)

    def test_warning_evidence_ref_is_schema_valid(self):
        mod = _load_adapter()
        out = Path("/tmp/toc-warning-test/toc-refresh.docx")
        out.parent.mkdir(parents=True, exist_ok=True)
        refs = mod._write_warning_evidence(out, [{
            "code": "automation_security_unavailable",
            "severity": "warning",
            "message": "Application.AutomationSecurity is unavailable",
            "stage": "open_application",
        }])
        self.assertEqual(1, len(refs))
        result = mod._toc_blocked(Path("/tmp/test.docx"), "s", reason_code="viewer_busy", message="m")
        result["evidence_refs"] = refs
        self.assertEqual([], _validate_schema(result))

    def test_blocked_result_can_carry_warning_evidence_ref(self):
        mod = _load_adapter()
        out = Path("/tmp/toc-warning-blocked-test/toc-refresh.docx")
        out.parent.mkdir(parents=True, exist_ok=True)
        refs = mod._write_warning_evidence(out, [{
            "code": "automation_security_unavailable",
            "severity": "warning",
            "message": "Application.AutomationSecurity is unavailable",
            "stage": "open_application",
        }])
        result = mod._toc_blocked(
            Path("/tmp/test.docx"),
            "com_error",
            reason_code="api_incompatible",
            message="failed",
            evidence_refs=refs,
        )
        self.assertEqual(["toc-refresh-warnings"], result["gate_check"]["evidence_refs"])
        self.assertEqual([], _validate_schema(result))

    def test_timeout_result_carries_warning_evidence_from_state(self):
        mod = _load_adapter()
        out = Path("/tmp/toc-warning-timeout-test/toc-refresh.docx")
        out.parent.mkdir(parents=True, exist_ok=True)
        refs = mod._write_warning_evidence(out, [{
            "code": "automation_security_unavailable",
            "severity": "warning",
            "message": "Application.AutomationSecurity is unavailable",
            "stage": "open_application",
        }])

        class FakeProcess:
            pid = 4321

            def poll(self):
                return 1

            def communicate(self, timeout=None):
                return "", ""

        result = mod._timeout_worker_result(
            Path("/tmp/test.docx"),
            FakeProcess(),
            {"stage": "update_all_fields", "warning_evidence_refs": refs},
            "stage_timeout",
            "timeout",
        )
        self.assertEqual(["toc-refresh-warnings"], result["gate_check"]["evidence_refs"])
        self.assertEqual([], _validate_schema(result))

    def test_worker_exit_without_result_blocks(self):
        mod = _load_adapter()

        def command_factory(_result_path, _state_path):
            return [sys.executable, "-c", "pass"]

        result = mod._run_toc_worker_with_timeout(
            Path("/tmp/test.docx"),
            Path("/tmp/out.docx"),
            {"ok": True},
            stage_timeout_seconds=5,
            worker_command_factory=command_factory,
        )
        self.assertEqual("blocked", result["status"])
        self.assertEqual("process_start_failed", result["error"]["reason_code"])

    def test_worker_invalid_contract_blocks(self):
        mod = _load_adapter()

        def command_factory(result_path, _state_path):
            script = "import pathlib,sys; pathlib.Path(sys.argv[1]).write_text('{\"status\":\"passed\"}', encoding='utf-8')"
            return [sys.executable, "-c", script, str(result_path)]

        result = mod._run_toc_worker_with_timeout(
            Path("/tmp/test.docx"),
            Path("/tmp/out.docx"),
            {"ok": True},
            stage_timeout_seconds=5,
            worker_command_factory=command_factory,
        )
        self.assertEqual("blocked", result["status"])
        self.assertEqual("process_start_failed", result["error"]["reason_code"])
        self.assertIn("'schema_id' is a required property", result["error"]["message"])

    def test_default_worker_timeout_is_stage_timeout(self):
        mod = _load_adapter()
        import inspect
        signature = inspect.signature(mod._run_toc_worker_with_timeout)
        self.assertEqual(mod.STAGE_TIMEOUT, signature.parameters["stage_timeout_seconds"].default)
        self.assertEqual(mod.TOTAL_TIMEOUT, signature.parameters["total_timeout_seconds"].default)


class SchemaTest(unittest.TestCase):
    def test_schema_exists(self):
        self.assertTrue(TOC_SCHEMA_PATH.exists())

    def test_passed_toc_valid(self):
        r = {
            "schema_id": "toc-acceptance",
            "schema_version": "2.0.0",
            "run_id": "r1",
            "required": True,
            "status": "passed",
            "viewer": "Microsoft Word",
            "viewer_version": "16.0",
            "platform": "windows",
            "input_ref": None,
            "output_ref": None,
            "before_sha256": "a" * 64,
            "after_sha256": "b" * 64,
            "field_update_count": 5,
            "toc_update_count": 1,
            "page_count": 10,
            "visible_entries": [
                {"level": 1, "text": "第一章", "page_number": 1}
            ],
            "evidence_refs": [],
            "error": {
                "code": "NONE", "reason_code": "none",
                "message": "", "retryable": False, "viewer": None,
            },
            "gate_check": {
                "gate_id": "g", "status": "passed",
                "checked_at": "2026-01-01T00:00:00Z",
                "predicate_version": "1.0.0",
                "evidence_refs": [], "failed_codes": []
            },
        }
        self.assertEqual([], _validate_schema(r))
        self.assertNotEqual("viewer_unavailable", r["error"]["reason_code"])

    def test_blocked_toc_valid(self):
        r = {
            "schema_id": "toc-acceptance",
            "schema_version": "2.0.0",
            "run_id": "r1",
            "required": True,
            "status": "blocked",
            "viewer": None, "viewer_version": None, "platform": None,
            "input_ref": None, "output_ref": None,
            "before_sha256": None, "after_sha256": None,
            "field_update_count": None, "toc_update_count": None,
            "page_count": None,
            "visible_entries": [],
            "evidence_refs": [],
            "error": {
                "code": "DFR-TOC-VIEWER_UNAVAILABLE",
                "reason_code": "viewer_unavailable",
                "message": "none",
                "retryable": False,
                "viewer": None,
            },
            "gate_check": {
                "gate_id": "g", "status": "blocked",
                "checked_at": "2026-01-01T00:00:00Z",
                "predicate_version": "1.0.0",
                "evidence_refs": [], "failed_codes": ["DFR-TOC-VIEWER_UNAVAILABLE"]
            },
        }
        self.assertEqual([], _validate_schema(r))

    def test_not_required_toc_valid(self):
        r = {
            "schema_id": "toc-acceptance",
            "schema_version": "2.0.0",
            "run_id": "r1",
            "required": False,
            "status": "not_required",
            "viewer": None, "viewer_version": None, "platform": None,
            "input_ref": None, "output_ref": None,
            "before_sha256": None, "after_sha256": None,
            "field_update_count": None, "toc_update_count": None,
            "page_count": None,
            "visible_entries": [],
            "evidence_refs": [],
            "error": None,
            "gate_check": {
                "gate_id": "g", "status": "passed",
                "checked_at": "2026-01-01T00:00:00Z",
                "predicate_version": "1.0.0",
                "evidence_refs": [], "failed_codes": []
            },
        }
        self.assertEqual([], _validate_schema(r))


if __name__ == "__main__":
    unittest.main(verbosity=2)
