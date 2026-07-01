"""V5-006 repair-plan v5 schema 和 planner 验证。"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

SCHEMA_PATH = ROOT / "docs" / "v5" / "schemas" / "repair-plan.schema.json"


def _minimal_v5_plan(plan_state: str = "draft", **overrides: Any) -> dict[str, Any]:
    """构造最小 v5 repair-plan。"""
    plan: dict[str, Any] = {
        "schema_id": "repair-plan",
        "schema_version": "2.0.0",
        "contract_version": "v5",
        "run_id": "run-test",
        "plan_id": "RP-run-test-20260616-000000",
        "plan_state": plan_state,
        "plan_revision": 0 if plan_state == "draft" else 42,
        "execution_backend": "officecli",
        "backend_version": "1.0.113",
        "created_at": "2026-06-16T00:00:00+08:00",
        "extensions": {},
        "snapshot_ref": {
            "artifact_id": "before-snapshot",
            "kind": "snapshot",
            "relative_path": "snapshots/officecli-document-snapshot.before.json",
            "sha256": "a" * 64,
            "size_bytes": 1000,
        },
        "capability_manifest_ref": {
            "artifact_id": "officecli-capability",
            "kind": "capability",
            "relative_path": "tools/officecli/officecli-capability-manifest.json",
            "sha256": "b" * 64,
            "size_bytes": 5000,
        },
        "source_audit_paths": ["semantic/semantic_audit.json"],
        "source_audit_refs": [],
        "risk_policy_path": "format-rules/test/risk-policy.yaml",
        "risk_policy_ref": None,
        "decision_snapshot": None if plan_state == "draft" else {
            "allows_continue": True,
            "decisions": [],
            "decided_at": "2026-06-16T00:00:00+08:00",
        },
        "manual_review_items_ref": {
            "ref_state": "draft" if plan_state == "draft" else "finalized",
            "path": "plans/manual_review_items.json",
            "pending_count": 0,
            "blocking_count": 0,
            "total_count": 0,
        },
        "actions": [],
        "manual_review_required": False,
        "execution_order": ["normalize_styles"],
        "post_repair": {},
        "generated_at": "2026-06-16T00:00:00+08:00",
    }
    if plan_state == "finalized":
        plan["finalized_from_plan_id"] = "RP-run-test-draft"
        plan["finalized_at"] = "2026-06-16T00:00:00+08:00"
    plan.update(overrides)
    return plan


def _validate_against_schema(instance: dict[str, Any]) -> list[str]:
    """使用 draft 2020-12 子集校验。"""
    schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
    errors: list[str] = []

    def _check(value: Any, sch: dict[str, Any], path: str) -> None:
        if "$ref" in sch:
            ref_name = sch["$ref"].rsplit("/", 1)[-1]
            sch = schema.get("$defs", {}).get(ref_name, sch)
        if "const" in sch and value != sch["const"]:
            errors.append(f"{path}: const mismatch (expected {sch['const']})")
            return
        if "enum" in sch and value not in sch["enum"]:
            errors.append(f"{path}: enum mismatch")
            return
        expected_type = sch.get("type")
        if expected_type and not _type_match(value, expected_type):
            errors.append(f"{path}: type mismatch")
            return
        if isinstance(value, str):
            if "minLength" in sch and len(value) < sch["minLength"]:
                errors.append(f"{path}: minLength")
            if "pattern" in sch:
                import re
                if not re.fullmatch(sch["pattern"], value):
                    errors.append(f"{path}: pattern mismatch")
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            if "minimum" in sch and value < sch["minimum"]:
                errors.append(f"{path}: minimum")
        if isinstance(value, dict):
            req = sch.get("required", [])
            for r in req:
                if r not in value:
                    errors.append(f"{path}.{r}: required")
            props = sch.get("properties", {})
            add = sch.get("additionalProperties")
            if add is False:
                for k in value:
                    if k not in props:
                        errors.append(f"{path}.{k}: additional property")
            for k, v in value.items():
                if k in props:
                    _check(v, props[k], f"{path}.{k}")
        if isinstance(value, list) and "items" in sch:
            for i, item in enumerate(value):
                _check(item, sch["items"], f"{path}[{i}]")

    def _type_match(value: Any, expected: Any) -> bool:
        types = expected if isinstance(expected, list) else [expected]
        for t in types:
            if t == "null" and value is None:
                return True
            if t == "object" and isinstance(value, dict):
                return True
            if t == "array" and isinstance(value, list):
                return True
            if t == "string" and isinstance(value, str):
                return True
            if t == "integer" and isinstance(value, int) and not isinstance(value, bool):
                return True
            if t == "number" and isinstance(value, (int, float)) and not isinstance(value, bool):
                return True
            if t == "boolean" and isinstance(value, bool):
                return True
        return False

    _check(instance, schema, "$")
    return errors


class V5RepairPlanSchemaTest(unittest.TestCase):
    """v5 repair-plan schema 验证。"""

    def test_schema_exists(self):
        self.assertTrue(SCHEMA_PATH.exists())

    def test_minimal_draft_passes(self):
        plan = _minimal_v5_plan("draft")
        self.assertEqual([], _validate_against_schema(plan))

    def test_minimal_draft_passes_draft202012_validator(self):
        from jsonschema import Draft202012Validator
        schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
        plan = _minimal_v5_plan("draft")
        self.assertEqual([], list(Draft202012Validator(schema).iter_errors(plan)))

    def test_minimal_finalized_passes(self):
        plan = _minimal_v5_plan("finalized")
        self.assertEqual([], _validate_against_schema(plan))

    def test_minimal_finalized_passes_draft202012_validator(self):
        from jsonschema import Draft202012Validator
        schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
        plan = _minimal_v5_plan("finalized")
        self.assertEqual([], list(Draft202012Validator(schema).iter_errors(plan)))

    def test_draft_revision_must_be_zero(self):
        from scripts.validation.manual_review_repair import validate_repair_plan_v5
        plan = _minimal_v5_plan("draft", plan_revision=5)
        result = validate_repair_plan_v5(plan)
        self.assertFalse(result.valid)
        self.assertTrue(any("plan_revision" in e for e in result.errors), result.errors)

    def test_finalized_revision_must_be_positive(self):
        from scripts.validation.manual_review_repair import validate_repair_plan_v5
        plan = _minimal_v5_plan("finalized", plan_revision=0)
        result = validate_repair_plan_v5(plan)
        self.assertFalse(result.valid)
        self.assertTrue(any("plan_revision" in e for e in result.errors), result.errors)

    def test_finalized_requires_finalized_fields(self):
        from scripts.validation.manual_review_repair import validate_repair_plan_v5
        plan = _minimal_v5_plan("finalized")
        del plan["finalized_from_plan_id"]
        result = validate_repair_plan_v5(plan)
        self.assertFalse(result.valid)
        self.assertTrue(any("finalized_from_plan_id" in e for e in result.errors), result.errors)

    def test_missing_execution_backend(self):
        plan = _minimal_v5_plan("draft")
        del plan["execution_backend"]
        errors = _validate_against_schema(plan)
        self.assertTrue(any("execution_backend" in e for e in errors), errors)

    def test_wrong_contract_version(self):
        plan = _minimal_v5_plan("draft", contract_version="v4")
        errors = _validate_against_schema(plan)
        self.assertTrue(any("contract_version" in e for e in errors), errors)

    def test_action_with_valid_risk_class(self):
        plan = _minimal_v5_plan("draft")
        plan["actions"] = [{
            "action_id": "A001",
            "source_issue_ids": ["K001"],
            "action_type": "apply_body_direct_format",
            "operation": "set",
            "target": {"element_id": "p-00001"},
            "target_binding": {
                "node_id": "N-" + "1" * 24,
                "path": "/body/p[1]",
                "fingerprint": "f" * 64,
            },
            "confidence": 0.95,
            "auto_fix_policy": "auto-fix",
            "risk_level": "low",
            "risk_class": "L1",
            "status": "pending",
        }]
        self.assertEqual([], _validate_against_schema(plan))

    def test_action_invalid_risk_class(self):
        plan = _minimal_v5_plan("draft")
        plan["actions"] = [{
            "action_id": "A001",
            "source_issue_ids": ["K001"],
            "action_type": "test",
            "target": {},
            "confidence": 0.9,
            "auto_fix_policy": "manual-review",
            "risk_level": "low",
            "risk_class": "INVALID",
            "status": "pending",
        }]
        errors = _validate_against_schema(plan)
        self.assertTrue(any("risk_class" in e for e in errors), errors)

    def test_backend_action_command_enum(self):
        plan = _minimal_v5_plan("finalized")
        plan["actions"] = [{
            "action_id": "A001",
            "source_issue_ids": ["K001"],
            "action_type": "apply_body_direct_format",
            "target": {},
            "target_binding": {
                "node_id": "N-" + "1" * 24,
                "path": "/body/p[1]",
                "fingerprint": "f" * 64,
            },
            "confidence": 0.95,
            "auto_fix_policy": "auto-fix",
            "risk_level": "low",
            "risk_class": "L1",
            "status": "executable",
            "execution_status": "executable",
            "backend_action": {
                "command": "invalid_cmd",
                "path": "/body/p[1]",
                "element_type": "paragraph",
                "properties": {"style": "1"},
                "index": None,
                "destination_path": None,
                "raw": None,
            },
        }]
        errors = _validate_against_schema(plan)
        self.assertTrue(any("command" in e for e in errors), errors)

    def test_backend_action_properties_scalars_only(self):
        from scripts.validation.manual_review_repair import validate_repair_plan_v5
        plan = _minimal_v5_plan("finalized")
        plan["actions"] = [{
            "action_id": "A001",
            "source_issue_ids": ["K001"],
            "action_type": "apply_body_direct_format",
            "target": {},
            "target_binding": {
                "node_id": "N-" + "1" * 24,
                "path": "/body/p[1]",
                "fingerprint": "f" * 64,
            },
            "confidence": 0.95,
            "auto_fix_policy": "auto-fix",
            "risk_level": "low",
            "risk_class": "L1",
            "status": "executable",
            "execution_status": "executable",
            "allowed_by_policy": True,
            "policy_match_ref": {"source_kind": "action_whitelist"},
            "backend_action": {
                "command": "set",
                "path": "/body/p[1]",
                "element_type": "paragraph",
                "properties": {"nested": {"key": "value"}},
                "index": None,
                "destination_path": None,
                "raw": None,
            },
        }]
        result = validate_repair_plan_v5(plan)
        self.assertFalse(result.valid)
        self.assertTrue(any("must be scalar" in e for e in result.errors), result.errors)

    def test_l3_write_requires_raw_fields(self):
        plan = _minimal_v5_plan("finalized")
        plan["actions"] = [{
            "action_id": "A001",
            "source_issue_ids": ["K001"],
            "action_type": "apply_table_border",
            "operation": "raw-set",
            "target": {},
            "target_binding": {
                "node_id": "N-" + "1" * 24,
                "path": "/body/tbl[1]/row[1]/cell[1]",
                "fingerprint": "f" * 64,
            },
            "confidence": 0.95,
            "auto_fix_policy": "manual-review",
            "risk_level": "high",
            "risk_class": "L3_WRITE",
            "status": "executable",
            "execution_status": "executable",
            "manual_confirmation_ref": {
                "artifact_id": "mri-001",
                "kind": "review",
                "relative_path": "plans/manual_review_items.json",
                "sha256": "c" * 64,
                "size_bytes": 200,
            },
            "backend_action": {
                "command": "raw-set",
                "path": "/body/tbl[1]/row[1]/cell[1]",
                "element_type": "cell",
                "properties": {},
                "index": None,
                "destination_path": None,
                "raw": {
                    "part": "document",
                    "xpath": "/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]",
                    "action": "replace",
                    "xml": "<w:tc xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>",
                    "xml_sha256": "d" * 64,
                    "expected_match_count": 1,
                    "precondition_raw_sha256": "e" * 64,
                    "manual_review_id": "MRI-001",
                    "decision_snapshot_sha256": "f" * 64,
                },
            },
        }]
        self.assertEqual([], _validate_against_schema(plan))

    def test_l3_write_raw_action_uses_canonical_enum(self):
        from jsonschema import Draft202012Validator
        plan = _minimal_v5_plan("finalized")
        plan["actions"] = [{
            "action_id": "A001",
            "source_issue_ids": ["K001"],
            "action_type": "apply_table_border",
            "operation": "raw-set",
            "target": {},
            "target_binding": {
                "node_id": "N-" + "1" * 24,
                "path": "/body/tbl[1]/row[1]/cell[1]",
                "fingerprint": "f" * 64,
            },
            "confidence": 0.95,
            "auto_fix_policy": "manual-review",
            "risk_level": "high",
            "risk_class": "L3_WRITE",
            "status": "executable",
            "execution_status": "executable",
            "manual_confirmation_ref": {
                "artifact_id": "mri-001",
                "kind": "review",
                "relative_path": "plans/manual_review_items.json",
                "sha256": "c" * 64,
                "size_bytes": 200,
            },
            "backend_action": {
                "command": "raw-set",
                "path": "/body/tbl[1]/row[1]/cell[1]",
                "element_type": "cell",
                "properties": {},
                "index": None,
                "destination_path": None,
                "raw": {
                    "part": "document",
                    "xpath": "/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]",
                    "action": "before",
                    "xml": "<w:tc xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>",
                    "xml_sha256": "d" * 64,
                    "expected_match_count": 1,
                    "precondition_raw_sha256": "e" * 64,
                    "manual_review_id": "MRI-001",
                    "decision_snapshot_sha256": "f" * 64,
                },
            },
        }]
        schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
        self.assertTrue(list(Draft202012Validator(schema).iter_errors(plan)))

    def test_target_binding_node_id_pattern(self):
        plan = _minimal_v5_plan("draft")
        plan["actions"] = [{
            "action_id": "A001",
            "source_issue_ids": ["K001"],
            "action_type": "test",
            "target": {},
            "target_binding": {
                "node_id": "invalid-format",
                "path": "/body/p[1]",
                "fingerprint": "f" * 64,
            },
            "confidence": 0.9,
            "auto_fix_policy": "manual-review",
            "risk_level": "low",
            "risk_class": "L2",
            "status": "pending",
        }]
        errors = _validate_against_schema(plan)
        self.assertTrue(any("node_id" in e for e in errors), errors)

    def test_artifact_ref_sha256_pattern(self):
        plan = _minimal_v5_plan("draft")
        plan["snapshot_ref"]["sha256"] = "bad-hash"
        errors = _validate_against_schema(plan)
        self.assertTrue(any("sha256" in e for e in errors), errors)


class V5RepairPlanBackendActionTest(unittest.TestCase):
    """backend_action 映射正确性测试。"""

    def _load_build_module(self):
        import importlib.util
        path = ROOT / ".codex" / "skills" / "docx-repair-planner" / "scripts" / "build_repair_plan.py"
        spec = importlib.util.spec_from_file_location("test_build_repair_plan_v5", path)
        module = importlib.util.module_from_spec(spec)
        assert spec and spec.loader
        spec.loader.exec_module(module)
        return module

    def test_backend_action_map_covers_all_whitelist(self):
        mod = self._load_build_module()
        from scripts.validation.manual_review_repair import WHITELIST_ACTIONS
        for action_type in WHITELIST_ACTIONS:
            self.assertIn(action_type, mod.BACKEND_ACTION_MAP,
                          f"{action_type} missing from BACKEND_ACTION_MAP")

    def test_map_heading_is_set_paragraph(self):
        mod = self._load_build_module()
        cfg = mod.BACKEND_ACTION_MAP["map_heading_native_style"]
        self.assertEqual("set", cfg["command"])
        self.assertEqual("paragraph", cfg["element_type"])

    def test_toc_audit_is_query_toc(self):
        mod = self._load_build_module()
        cfg = mod.BACKEND_ACTION_MAP["toc_content_audit"]
        self.assertEqual("query", cfg["command"])
        self.assertTrue(cfg["read_only"])

    def test_risk_class_forbidden_border_elevates_l3(self):
        mod = self._load_build_module()
        risk = mod.assign_risk_class(
            "apply_table_border",
            {"border.top": "single"},
            mod.BACKEND_ACTION_MAP.get("apply_table_border"),
        )
        self.assertEqual("L3_WRITE", risk)

    def test_risk_class_shd_fill_elevates_l3(self):
        mod = self._load_build_module()
        risk = mod.assign_risk_class(
            "apply_body_direct_format",
            {"shd.fill": "FF0000"},
            mod.BACKEND_ACTION_MAP.get("apply_body_direct_format"),
        )
        self.assertEqual("L3_WRITE", risk)

    def test_risk_class_body_format_is_l1(self):
        mod = self._load_build_module()
        risk = mod.assign_risk_class(
            "apply_body_direct_format",
            {"font_size_pt": 12},
            mod.BACKEND_ACTION_MAP.get("apply_body_direct_format"),
        )
        self.assertEqual("L1", risk)

    def test_canonical_plan_output_path_honors_requested_directory(self):
        mod = self._load_build_module()
        draft = mod.canonical_plan_output_path(
            Path("format_runs/run-test/plans/repair_plan.draft.yaml"),
            "draft",
            0,
        )
        finalized = mod.canonical_plan_output_path(
            Path("format_runs/run-test/plans/repair_plan.finalized.yaml"),
            "finalized",
            7,
        )
        self.assertEqual(Path("format_runs/run-test/plans/repair_plan.draft.yaml"), draft)
        self.assertEqual(Path("format_runs/run-test/plans/repair_plan.finalized.r007.yaml"), finalized)

    def test_main_from_args_writes_finalized_plan_next_to_requested_output(self):
        mod = self._load_build_module()
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            work_cwd = root / "cwd"
            semantic_audit = root / "semantic_audit.json"
            snapshot = root / "snapshots" / "officecli-document-snapshot.before.json"
            capability = root / "tools" / "officecli-capability-manifest.json"
            output = root / "format_runs" / "run-test" / "plans" / "repair_plan.finalized.yaml"
            work_cwd.mkdir()
            semantic_audit.write_text(json.dumps({"items": []}), encoding="utf-8")
            snapshot.parent.mkdir(parents=True)
            snapshot.write_text(json.dumps({"nodes": []}), encoding="utf-8")
            capability.parent.mkdir(parents=True)
            capability.write_text(json.dumps({"schema_id": "officecli-capability-manifest"}), encoding="utf-8")

            old_cwd = Path.cwd()
            try:
                os.chdir(work_cwd)
                result = mod.main_from_args([
                    "--semantic-audit", str(semantic_audit),
                    "--snapshot", str(snapshot),
                    "--capability-manifest", str(capability),
                    "--run-id", "run-test",
                    "--plan-state", "finalized",
                    "--rule-id", "rule-test",
                    "--source-docx", str(root / "input" / "source.docx"),
                    "--working-docx", str(root / "input" / "working.docx"),
                    "--output-docx", str(root / "output" / "repaired.docx"),
                    "--output", str(output),
                ])
            finally:
                os.chdir(old_cwd)

            self.assertEqual(0, result)
            finalized = list(output.parent.glob("repair_plan.finalized.r*.yaml"))
            self.assertEqual(1, len(finalized))
            self.assertFalse((work_cwd / "plans").exists())


if __name__ == "__main__":
    unittest.main(verbosity=2)
