"""CODE-011 回归覆盖输入与扫描工具。

覆盖条款：
- 40-§7.2/§7.3：V4 回归测试与 V4-T01 至 V4-T45 可执行测试矩阵。
- 50-§3.3：最低必测集合、not_automated_with_reason 治理和模板扫描清单。
- 50-§3.8：60_TEST_PLAN.md 输入材料。
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any


MINIMUM_REQUIRED_TEST_IDS = {
    "V4-T04",
    "V4-T08",
    "V4-T09",
    "V4-T16",
    "V4-T16A",
    "V4-T17",
    "V4-T18",
    "V4-T19",
    "V4-T20",
    "V4-T24",
    "V4-T28",
    "V4-T29",
    "V4-T30",
    "V4-T31",
    "V4-T32",
    "V4-T33",
    "V4-T34",
    "V4-T35",
    "V4-T36",
    "V4-T37",
    "V4-T38",
    "V4-T39",
    "V4-T40",
    "V4-T41",
    "V4-T42",
    "V4-T43",
    "V4-T44",
    "V4-T45",
}
ALL_V4_TEST_IDS = {f"V4-T{index:02d}" for index in range(1, 46)} | {"V4-T16A"} | {f"V4-TS{index:02d}" for index in range(1, 11)}

VERIFICATION_TYPES = {
    "static_scan",
    "schema_example",
    "schema_invalid",
    "synthetic_fixture",
    "office_integration",
    "manual_gate",
    "not_automated_with_reason",
    "pending_implementation",
}
IMPLEMENTATION_STATUSES = {"automated", "pending_implementation", "not_automated_with_reason"}
VALID_NOT_AUTOMATED_REASONS = {
    "requires_office_com",
    "requires_real_word_rendering",
    "requires_human_semantic_judgment",
    "requires_external_license_or_environment",
}
FORBIDDEN_NOT_AUTOMATED_REASON_TEXT = {"实现复杂", "暂未开发", "pending", "todo"}
REQUIRED_SKILL_NAMES = {
    "docx-fact-extractor",
    "docx-format-auditor",
    "docx-format-repairer",
    "docx-format-reporter",
    "docx-repair-planner",
    "docx-rule-packager",
    "docx-semantic-strategist",
    "format-helper",
}
REQUIRED_TEMPLATE_SECTIONS = ["任务清单", "当前阶段", "执行结果", "交付物", "阻塞/人工确认", "下一步", "验收自检"]
INTERNAL_SKILL_PATTERN = re.compile(r"\bdocx-[a-z-]+\b")


@dataclass(frozen=True)
class CoverageInput:
    """CODE-011 给 CODE-011A/60_TEST_PLAN 使用的覆盖输入行。"""

    test_id: str
    owner_task: str
    verification_type: str
    evidence_path: str
    gate_relevance: bool
    status: str
    reason: str | None = None


CODE_011_REGRESSION_INPUTS: list[CoverageInput] = [
    CoverageInput("V4-T04", "CODE-010A", "synthetic_fixture", "tests/validation/test_final_acceptance.py", True, "automated"),
    CoverageInput("V4-T08", "CODE-011", "static_scan", "tests/validation/test_regression_coverage.py", True, "automated"),
    CoverageInput("V4-T09", "CODE-011", "synthetic_fixture", "tests/validation/test_regression_coverage.py", True, "automated"),
    CoverageInput("V4-T16", "CODE-006A", "synthetic_fixture", "tests/validation/test_run_state_manager.py", True, "automated"),
    CoverageInput("V4-T16A", "CODE-006", "synthetic_fixture", "tests/validation/test_skill_result_io.py", True, "automated"),
    CoverageInput("V4-T17", "CODE-006A", "synthetic_fixture", "tests/validation/test_run_state_manager.py", True, "automated"),
    CoverageInput("V4-T18", "CODE-011", "synthetic_fixture", "tests/validation/test_v4_required_fixtures.py", True, "automated"),
    CoverageInput("V4-T19", "CODE-001", "synthetic_fixture", "tests/validation/test_ensure_run_directories.py", True, "automated"),
    CoverageInput("V4-T20", "CODE-007", "synthetic_fixture", "tests/validation/test_gate_predicates.py", True, "automated"),
    CoverageInput("V4-T24", "CODE-011", "synthetic_fixture", "tests/validation/test_regression_coverage.py", True, "automated"),
    CoverageInput("V4-T28", "CODE-011", "synthetic_fixture", "tests/validation/test_v4_required_fixtures.py", True, "automated"),
    CoverageInput("V4-T29", "CODE-009", "synthetic_fixture", "tests/validation/test_manual_review_repair.py", True, "automated"),
    CoverageInput("V4-T30", "CODE-010", "synthetic_fixture", "tests/validation/test_final_acceptance.py", True, "automated"),
    CoverageInput("V4-T31", "CODE-010A", "synthetic_fixture", "tests/validation/test_final_acceptance.py", True, "automated"),
    CoverageInput("V4-T32", "CODE-011", "static_scan", "tests/validation/test_regression_coverage.py", True, "automated"),
    CoverageInput("V4-T33", "CODE-006A", "synthetic_fixture", "tests/validation/test_run_state_manager.py", True, "automated"),
    CoverageInput("V4-T34", "CODE-009A", "synthetic_fixture", "tests/validation/test_manual_review_repair.py", True, "automated"),
    CoverageInput("V4-T35", "CODE-008A", "synthetic_fixture", "tests/validation/test_evidence_manifest.py", True, "automated"),
    CoverageInput("V4-T36", "CODE-011", "synthetic_fixture", "tests/validation/test_v4_required_fixtures.py", True, "automated"),
    CoverageInput("V4-T37", "CODE-010", "synthetic_fixture", "tests/validation/test_final_acceptance.py", True, "automated"),
    CoverageInput("V4-T38", "CODE-010", "synthetic_fixture", "tests/validation/test_final_acceptance.py", True, "automated"),
    CoverageInput("V4-T39", "CODE-011", "synthetic_fixture", "tests/validation/test_v4_required_fixtures.py", True, "automated"),
    CoverageInput("V4-T40", "CODE-005", "schema_example", "tests/validation/test_schema_inventory_scan.py", True, "automated"),
    CoverageInput("V4-T41", "CODE-009A", "synthetic_fixture", "tests/validation/test_manual_review_repair.py", True, "automated"),
    CoverageInput("V4-T42", "CODE-009A", "synthetic_fixture", "tests/validation/test_manual_review_repair.py", True, "automated"),
    CoverageInput("V4-T43", "CODE-011", "synthetic_fixture", "tests/validation/test_v4_required_fixtures.py", True, "automated"),
    CoverageInput("V4-T44", "CODE-008A", "synthetic_fixture", "tests/validation/test_evidence_manifest.py", True, "automated"),
    CoverageInput("V4-T45", "CODE-011", "synthetic_fixture", "tests/validation/test_v4_required_fixtures.py", True, "automated"),
]


def coverage_inputs_as_dicts() -> list[dict[str, Any]]:
    """返回可序列化的 CODE-011 覆盖输入材料。"""
    return [item.__dict__.copy() for item in CODE_011_REGRESSION_INPUTS]


def validate_regression_inputs(items: list[CoverageInput] | None = None) -> list[str]:
    """校验最低必测集合、枚举和未实现项原因。"""
    rows = items or CODE_011_REGRESSION_INPUTS
    errors: list[str] = []
    by_id = {item.test_id: item for item in rows}
    missing = sorted(MINIMUM_REQUIRED_TEST_IDS - by_id.keys())
    if missing:
        errors.append(f"minimum required V4 tests missing: {missing}")
    for item in rows:
        if item.verification_type not in VERIFICATION_TYPES:
            errors.append(f"{item.test_id}.verification_type is invalid: {item.verification_type}")
        if item.status not in IMPLEMENTATION_STATUSES:
            errors.append(f"{item.test_id}.status is invalid: {item.status}")
        if item.status == "pending_implementation" and not item.reason:
            errors.append(f"{item.test_id}.pending_implementation requires reason")
        if item.status == "not_automated_with_reason":
            if item.reason not in VALID_NOT_AUTOMATED_REASONS:
                errors.append(f"{item.test_id}.not_automated_with_reason is not allowed: {item.reason}")
            if item.reason and any(token in item.reason for token in FORBIDDEN_NOT_AUTOMATED_REASON_TEXT):
                errors.append(f"{item.test_id}.not_automated_with_reason uses forbidden reason text")
        if item.status == "automated" and not item.evidence_path.startswith("tests/"):
            errors.append(f"{item.test_id}.automated evidence_path must point to tests/")
    return errors


def scan_skill_templates(skill_root: Path) -> list[str]:
    """扫描 8 个 skill 的触发、固定输出分块和失败模板。"""
    errors: list[str] = []
    found = {path.parent.name: path for path in skill_root.glob("*/SKILL.md")}
    missing_skills = sorted(REQUIRED_SKILL_NAMES - found.keys())
    if missing_skills:
        errors.append(f"missing SKILL.md files: {missing_skills}")
    for skill_name in sorted(REQUIRED_SKILL_NAMES & found.keys()):
        content = found[skill_name].read_text(encoding="utf-8")
        if "触发" not in content and "使用时机" not in content:
            errors.append(f"{skill_name}: missing trigger/使用时机")
        for section in REQUIRED_TEMPLATE_SECTIONS:
            if section not in content:
                errors.append(f"{skill_name}: missing output section {section}")
        if "失败" not in content and "blocked" not in content:
            errors.append(f"{skill_name}: missing failure/blocked branch")
    return errors


def scan_format_helper_markdown_gate(root: Path) -> list[str]:
    """扫描 format-helper 不得通过 Markdown 模板驱动 Gate。"""
    errors: list[str] = []
    format_helper_root = root / ".codex" / "skills" / "format-helper"
    forbidden_patterns = (
        re.compile(r"parse\s*\([^)]*\.md", re.IGNORECASE),
        re.compile(r"Gate[^\\n]+Markdown", re.IGNORECASE),
    )
    for path in format_helper_root.rglob("*"):
        if path.is_file() and path.suffix in {".py", ".md"}:
            content = path.read_text(encoding="utf-8", errors="ignore")
            for pattern in forbidden_patterns:
                if pattern.search(content):
                    errors.append(f"{path.relative_to(root)} matches forbidden Markdown Gate pattern")
    return errors


def scan_plain_user_output(output_text: str) -> list[str]:
    """扫描普通用户输出不得泄漏内部 docx-* skill 名称或内部 JSON 正文。"""
    errors: list[str] = []
    if INTERNAL_SKILL_PATTERN.search(output_text):
        errors.append("plain output must not expose internal docx-* skill names")
    if re.search(r'\{\s*"schema_id"\s*:', output_text):
        errors.append("plain output must not embed internal JSON body")
    return errors


def scan_repair_log_filename_references(root: Path) -> list[str]:
    """扫描 repair execution log 统一命名。"""
    errors: list[str] = []
    allowed = "logs/repair_execution_log.json"
    forbidden_names = {"repair_log.json", "repair-execution-log.json", "repairExecutionLog.json"}
    for base in (root / "docs" / "v4", root / "tests", root / "scripts"):
        for path in base.rglob("*"):
            if path.is_file() and path.suffix in {".md", ".py", ".json", ".yaml", ".yml"}:
                content = path.read_text(encoding="utf-8", errors="ignore")
                for forbidden in forbidden_names:
                    if forbidden in content and allowed not in content:
                        errors.append(f"{path.relative_to(root)} contains legacy repair log name {forbidden}")
    return errors


# ── CODE-018 槽位契约关键字扫描 ──────────────────────────────────────


SLOT_KEYWORD_CHECKS = {
    "docx-semantic-strategist": {
        "keywords": ["role_format_slot_facts", "role_slot_contract", "FH-SLOT-FACTS-UNRESOLVED"],
        "description": "SLOT_CONTRACT_DESIGN.md §9.4",
    },
    "docx-rule-packager": {
        "keywords": ["slot_facts_ref", "resolved_slot_count"],
        "description": "SLOT_CONTRACT_DESIGN.md §9.4",
    },
    "format-helper": {
        "keywords": ["rule_confirmation_gate", "waiting_user_on_unresolved_slots"],
        "description": "SLOT_CONTRACT_DESIGN.md §9.4",
    },
}


def scan_slot_keywords(skill_root: Path) -> list[str]:
    """CODE-018: 扫描 SKILL.md 中的槽位契约关键字。"""
    errors: list[str] = []
    for skill_name, check in SLOT_KEYWORD_CHECKS.items():
        skill_md = skill_root / skill_name / "SKILL.md"
        if not skill_md.exists():
            errors.append(f"{skill_name}/SKILL.md 不存在")
            continue
        content = skill_md.read_text(encoding="utf-8")
        for keyword in check["keywords"]:
            if keyword not in content:
                errors.append(
                    f"{skill_name}/SKILL.md 缺少槽位契约关键字「{keyword}」（{check['description']}）"
                )
    return errors


def load_coverage_matrix(path: Path) -> dict[str, Any]:
    """读取 coverage_matrix.yaml。

    当前文件使用 JSON-compatible YAML，避免为本地 schema 校验引入额外依赖。
    文件不存在或 JSON 无效时返回带错误标识的最小 dict。
    """
    try:
        if not path.exists():
            return {"schema_id": "coverage-matrix", "_load_error": f"文件不存在: {path}"}
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, PermissionError) as exc:
        return {"schema_id": "coverage-matrix", "_load_error": str(exc)}


def validate_coverage_matrix(matrix: dict[str, Any]) -> list[str]:
    """校验 CODE-011A 覆盖矩阵本地 schema 和完整性。"""
    errors: list[str] = []
    required_top = {"schema_id", "schema_version", "contract_version", "generated_at", "items"}
    for field_name in sorted(required_top):
        if field_name not in matrix:
            errors.append(f"{field_name} is required")
    if matrix.get("schema_id") != "coverage-matrix":
        errors.append("schema_id must be coverage-matrix")
    if matrix.get("contract_version") != "v4":
        errors.append("contract_version must be v4")
    items = matrix.get("items")
    if not isinstance(items, list):
        return errors + ["items must be array"]
    by_id: dict[str, dict[str, Any]] = {}
    required_item = {"test_id", "owner_task", "verification_type", "evidence_path", "gate_relevance"}
    for index, item in enumerate(items):
        if not isinstance(item, dict):
            errors.append(f"items[{index}] must be object")
            continue
        for field_name in sorted(required_item):
            if field_name not in item:
                errors.append(f"items[{index}].{field_name} is required")
        test_id = item.get("test_id")
        if test_id in by_id:
            errors.append(f"duplicate test_id: {test_id}")
        if test_id:
            by_id[str(test_id)] = item
        if item.get("verification_type") not in VERIFICATION_TYPES:
            errors.append(f"{test_id}.verification_type is invalid")
        if not isinstance(item.get("gate_relevance"), bool):
            errors.append(f"{test_id}.gate_relevance must be boolean")
        if item.get("status") not in IMPLEMENTATION_STATUSES:
            errors.append(f"{test_id}.status is invalid")
        if item.get("status") == "pending_implementation" and not item.get("pending_reason"):
            errors.append(f"{test_id}.pending_implementation requires pending_reason")
        if item.get("verification_type") == "office_integration":
            if not item.get("office_capability"):
                errors.append(f"{test_id}.office_capability is required for office_integration")
            if item.get("status") == "not_automated_with_reason" and not item.get("manual_validation_path"):
                errors.append(f"{test_id}.manual_validation_path is required for not_automated_with_reason")
        if item.get("status") == "not_automated_with_reason":
            reason = item.get("not_automated_reason")
            if reason not in VALID_NOT_AUTOMATED_REASONS:
                errors.append(f"{test_id}.not_automated_reason is invalid")
            if not item.get("risk_acceptance_path"):
                errors.append(f"{test_id}.risk_acceptance_path is required")
        if item.get("status") == "automated" and not str(item.get("evidence_path", "")).startswith("tests/"):
            errors.append(f"{test_id}.automated evidence_path must point to tests/")
    missing_all = sorted(ALL_V4_TEST_IDS - by_id.keys())
    if missing_all:
        errors.append(f"coverage matrix missing V4 tests: {missing_all}")
    extra = sorted(set(by_id) - ALL_V4_TEST_IDS)
    if extra:
        errors.append(f"coverage matrix contains unknown tests: {extra}")
    missing_minimum = sorted(MINIMUM_REQUIRED_TEST_IDS - by_id.keys())
    if missing_minimum:
        errors.append(f"minimum required V4 tests missing: {missing_minimum}")
    return errors


__all__ = [
    "CODE_011_REGRESSION_INPUTS",
    "ALL_V4_TEST_IDS",
    "MINIMUM_REQUIRED_TEST_IDS",
    "SLOT_KEYWORD_CHECKS",
    "coverage_inputs_as_dicts",
    "load_coverage_matrix",
    "scan_format_helper_markdown_gate",
    "scan_plain_user_output",
    "scan_repair_log_filename_references",
    "scan_skill_templates",
    "scan_slot_keywords",
    "validate_regression_inputs",
    "validate_coverage_matrix",
]
