#!/usr/bin/env python3
"""生成最终交付报告，并在需要时回写最终验收状态。"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[4]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.reporting.human_readable import markdown_list, render_template, safe_markdown_text, status_marker
from scripts.utils.simple_yaml import load_yaml, write_yaml
from scripts.validation.human_readable_report import (
    FINAL_REPORT_REQUIRED_SECTIONS,
    assert_human_readable_report,
)


TZ = timezone(timedelta(hours=8))
TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "templates" / "FINAL_REPORT.template.md"
LEGACY_REPAIR_PLAN_WARNING = "未发现 repair_plan.yaml，已按现有执行产物兼容渲染。"


def load_json(path: Path) -> dict[str, Any]:
    """读取 JSON 文件。"""
    return json.loads(path.read_text(encoding="utf-8"))


def load_json_if_exists(path: Path) -> dict[str, Any] | None:
    """读取可选 JSON 文件。"""
    if not path.exists():
        return None
    return load_json(path)


def load_yaml_if_exists(path: Path) -> dict[str, Any] | None:
    """读取可选 YAML 文件。"""
    if not path.exists():
        return None
    return load_yaml(path)


def write_json_atomic(path: Path, payload: dict[str, Any]) -> None:
    """原子写入 JSON 文件。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    tmp_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(tmp_path, path)


def review_files(run_dir: Path) -> list[Path]:
    """列出二轮复核结果文件。"""
    return sorted((run_dir / "review_results").glob("T*.review.json"))


def load_reviews(run_dir: Path) -> list[dict[str, Any]]:
    """读取二轮复核结果。"""
    return [load_json(path) for path in review_files(run_dir)]


def collect_blockers(reviews: list[dict[str, Any]]) -> list[str]:
    """收集复核中的阻断项。"""
    blockers: list[str] = []
    for review in reviews:
        if review.get("status") != "blocked":
            continue
        task_name = review.get("task_name") or review.get("task_id") or "复核项"
        issues = review.get("issues") or []
        if not issues:
            blockers.append(f"{task_name} 未通过复核。")
            continue
        for issue in issues:
            description = issue.get("description") or "存在未通过项"
            blockers.append(f"{task_name}：{description}")
    return blockers


def detect_mode_details(run_dir: Path, final_acceptance: dict[str, Any] | None = None) -> tuple[str, bool]:
    """识别当前运行模式，并返回是否为明确识别结果。"""
    final_acceptance = final_acceptance or {}
    acceptance_type = str(final_acceptance.get("acceptance_type") or "").strip()
    if acceptance_type == "build_rules_terminal":
        return "extract-rule", True
    if acceptance_type == "audit_only_terminal":
        return "audit-only", True
    if acceptance_type == "final_delivery":
        return "repair", True

    state = load_yaml_if_exists(run_dir / "logs" / "state.yaml") or {}
    stage = str(state.get("stage") or "").strip()
    workflow_mode = str(state.get("workflow_mode") or state.get("mode") or "").strip()

    if stage in {"rule_packaging", "rule_confirmation"} or workflow_mode in {"extract-rule", "build_rules"}:
        return "extract-rule", True
    if workflow_mode in {"audit-only", "audit_only"}:
        return "audit-only", True
    if workflow_mode in {"repair", "final_delivery"}:
        return "repair", True

    if (run_dir / "semantic").exists() and not (run_dir / "logs" / "repair_execution.json").exists():
        return "extract-rule", True
    if (run_dir / "logs" / "repair_execution.json").exists() or (run_dir / "snapshots" / "document_snapshot.after.json").exists():
        return "repair", True
    if (run_dir / "snapshots" / "document_snapshot.before.json").exists():
        return "audit-only", True
    return "repair", False


def detect_mode(run_dir: Path, final_acceptance: dict[str, Any] | None = None) -> str:
    """返回当前运行模式。"""
    return detect_mode_details(run_dir, final_acceptance)[0]


def load_rule_profile(run_dir: Path, repair_plan: dict[str, Any] | None) -> dict[str, Any]:
    """尽量提取规则包信息。"""
    if repair_plan and isinstance(repair_plan.get("rule_profile"), dict):
        return repair_plan["rule_profile"]
    rule_ref = load_json_if_exists(run_dir / "logs" / "rule_ref.json") or {}
    if isinstance(rule_ref.get("rule_profile"), dict):
        return rule_ref["rule_profile"]
    return {}


def infer_input_doc(run_dir: Path, execution_log: dict[str, Any] | None) -> str:
    """推断输入文档路径。"""
    if execution_log:
        for key in ("working_docx", "source_docx", "input_docx"):
            if execution_log.get(key):
                return str(execution_log[key])
    input_dir = run_dir / "input"
    if input_dir.exists():
        files = sorted(path for path in input_dir.iterdir() if path.is_file())
        if files:
            return str(files[0]).replace("\\", "/")
    return "未指定"


def infer_output_doc(
    run_dir: Path,
    execution_log: dict[str, Any] | None,
    final_acceptance: dict[str, Any],
    mode: str,
) -> str:
    """推断输出文档路径。"""
    if execution_log and execution_log.get("output_docx"):
        return str(execution_log["output_docx"])
    if final_acceptance.get("final_docx_path"):
        return str((run_dir / str(final_acceptance["final_docx_path"])).resolve()).replace("\\", "/")
    if mode == "audit-only":
        return "本次未生成修复后文档"
    return "未指定"


def normalize_status(final_acceptance: dict[str, Any]) -> str:
    """统一最终状态。"""
    status = str(final_acceptance.get("status") or "").strip()
    if status:
        return status
    return "accepted" if final_acceptance.get("accepted") else "blocked"


def load_or_build_final_acceptance(run_dir: Path, *, mode: str) -> dict[str, Any]:
    """优先读取已有 final_acceptance，不存在时按运行产物构建兼容对象。"""
    existing = load_json_if_exists(run_dir / "logs" / "final_acceptance.json")
    if existing:
        return existing

    execution_log = load_json_if_exists(run_dir / "logs" / "repair_execution.json")
    repair_plan = load_yaml_if_exists(run_dir / "plans" / "repair_plan.yaml") or {}
    before_snapshot = load_json_if_exists(run_dir / "snapshots" / "document_snapshot.before.json")
    after_snapshot = load_json_if_exists(run_dir / "snapshots" / "document_snapshot.after.json")
    reviews = load_reviews(run_dir)

    blockers = collect_blockers(reviews)
    if mode == "repair":
        if execution_log is None:
            blockers.append("repair 模式缺少执行日志。")
        if before_snapshot is None:
            blockers.append("repair 模式缺少修复前快照。")
        if after_snapshot is None:
            blockers.append("repair 模式缺少修复后快照。")
        if not reviews:
            blockers.append("repair 模式缺少二轮复核结果。")
        if execution_log and not bool(execution_log.get("output_docx_valid", False)):
            blockers.append("输出 DOCX 未通过有效性检查。")

    accepted = not blockers
    output_doc = infer_output_doc(run_dir, execution_log, {}, mode)
    reports = []
    final_report_path = run_dir / "reports" / "FINAL_ACCEPTANCE_REPORT.md"
    if final_report_path.exists():
        reports.append(str(final_report_path).replace("\\", "/"))
    return {
        "schema_version": "1.0.0",
        "accepted": accepted,
        "status": "accepted" if accepted else "blocked",
        "created_at": datetime.now(TZ).isoformat(),
        "open_blockers": blockers,
        "manual_items_remaining": repair_plan.get("manual_review_items", []),
        "output_docx_valid": bool(execution_log.get("output_docx_valid", False)) if execution_log else False,
        "reports": reports,
        "evidence": [f"输出文件路径：{output_doc}"],
    }


def build_final_report_view_model(
    run_dir: Path,
    final_acceptance: dict[str, Any],
    *,
    mode: str = "repair",
    use_icons: bool = True,
    mode_recognized: bool = True,
) -> dict[str, object]:
    """从运行目录和 final_acceptance 构建最终交付报告 view model。"""
    execution_log = load_json_if_exists(run_dir / "logs" / "repair_execution.json")
    repair_plan = load_yaml_if_exists(run_dir / "plans" / "repair_plan.yaml")
    before_snapshot = load_json_if_exists(run_dir / "snapshots" / "document_snapshot.before.json")
    after_snapshot = load_json_if_exists(run_dir / "snapshots" / "document_snapshot.after.json")
    reviews = load_reviews(run_dir)
    rule_profile = load_rule_profile(run_dir, repair_plan)

    missing_items: list[str] = []
    if mode == "repair":
        if execution_log is None:
            missing_items.append("repair 模式缺少 logs/repair_execution.json")
        if before_snapshot is None:
            missing_items.append("repair 模式缺少 snapshots/document_snapshot.before.json")
        if after_snapshot is None:
            missing_items.append("repair 模式缺少 snapshots/document_snapshot.after.json")
        if not reviews:
            missing_items.append("repair 模式缺少 review_results/T*.review.json")
        if missing_items:
            raise ValueError("报告渲染不满足用户可读性 Gate：" + "；".join(missing_items))

    status = normalize_status(final_acceptance)
    if status == "accepted":
        status_label = "已通过"
    elif status == "accepted_with_warnings":
        status_label = "存在风险"
    else:
        status_label = "未通过"

    input_doc = infer_input_doc(run_dir, execution_log)
    output_doc = infer_output_doc(run_dir, execution_log, final_acceptance, mode)
    scope_text = rule_profile.get("scope") or rule_profile.get("applicable_scope") or "未声明规则适用范围"

    input_and_rule_lines = [
        f"输入文档：{safe_markdown_text(input_doc, max_length=None)}",
        f"输出文档：{safe_markdown_text(output_doc, max_length=None)}",
        f"运行 ID：{safe_markdown_text(run_dir.name, max_length=None)}",
        f"规则包：{safe_markdown_text(rule_profile.get('id', '未指定'), max_length=None)}",
        f"规则适用范围：{safe_markdown_text(scope_text, max_length=None)}",
    ]

    audit_summary_lines: list[str] = []
    if before_snapshot:
        audit_summary_lines.append(
            "输入快照：{paragraphs} 段，{tables} 张表。".format(
                paragraphs=before_snapshot.get("paragraph_count", 0),
                tables=before_snapshot.get("table_count", 0),
            )
        )
    else:
        audit_summary_lines.append("未发现格式问题")

    if execution_log and isinstance(execution_log.get("counts"), dict):
        counts = execution_log["counts"]
        count_parts = []
        for label, key in (
            ("标题段落", "heading_paragraphs"),
            ("正文段落", "body_paragraphs"),
            ("列表段落", "list_paragraphs"),
            ("表格单元格", "table_cells"),
        ):
            if key in counts:
                count_parts.append(f"{label} {counts[key]}")
        if count_parts:
            audit_summary_lines.append("涉及范围：" + "，".join(count_parts) + "。")

    if mode == "audit-only":
        repair_summary_lines = ["本次未执行自动修复"]
        before_after_lines = ["无可展示的修复前后对比项，原因：本次为仅审计流程，未生成修复后文档"]
    else:
        repair_summary_lines = []
        if repair_plan and isinstance(repair_plan.get("actions"), list):
            repair_summary_lines.append(f"计划修复项：{len(repair_plan.get('actions', []))}")
        if execution_log:
            actions_total = execution_log.get("actions_total")
            if actions_total is None and isinstance(execution_log.get("counts"), dict):
                actions_total = sum(int(value) for value in execution_log["counts"].values() if isinstance(value, int))
            if actions_total is not None:
                repair_summary_lines.append(f"处理统计项：{actions_total}")
            if execution_log.get("actions_executed") is not None:
                repair_summary_lines.append(f"成功项：{execution_log.get('actions_executed')}")
            if execution_log.get("actions_rejected") is not None:
                repair_summary_lines.append(f"失败项：{execution_log.get('actions_rejected')}")
            if execution_log.get("actions_skipped") is not None:
                repair_summary_lines.append(f"跳过项：{execution_log.get('actions_skipped')}")
        if not repair_summary_lines:
            repair_summary_lines.append("本次未执行自动修复")

        before_after_lines = []
        if before_snapshot and after_snapshot:
            before_after_lines.extend(
                [
                    "段落数：{before} -> {after}".format(
                        before=before_snapshot.get("paragraph_count", 0),
                        after=after_snapshot.get("paragraph_count", 0),
                    ),
                    "表格数：{before} -> {after}".format(
                        before=before_snapshot.get("table_count", 0),
                        after=after_snapshot.get("table_count", 0),
                    ),
                ]
            )
        else:
            before_after_lines.append("无可展示的修复前后对比项，原因：缺少 before 或 after 快照")

    manual_items = final_acceptance.get("manual_items_remaining")
    if not isinstance(manual_items, list) and repair_plan:
        manual_items = repair_plan.get("manual_review_items", [])
    if manual_items:
        unfixed_lines = [
            safe_markdown_text(item.get("reason") or item.get("description") or "存在未修复项", max_length=120)
            for item in manual_items
            if isinstance(item, dict)
        ]
        if not unfixed_lines:
            unfixed_lines = ["无未修复项"]
    else:
        unfixed_lines = ["无未修复项"]

    risk_lines: list[str] = []
    open_blockers = final_acceptance.get("open_blockers")
    if isinstance(open_blockers, list):
        for item in open_blockers:
            risk_lines.append(safe_markdown_text(item, max_length=120))
    if mode == "repair" and repair_plan is None:
        risk_lines.append(LEGACY_REPAIR_PLAN_WARNING)
    if mode == "repair" and not mode_recognized:
        risk_lines.append("未识别流程模式，已按 repair 模式验收。")
    if not reviews and mode == "audit-only":
        risk_lines.append("本次未生成二轮复核产物。")
    if not risk_lines:
        risk_lines = ["无已知剩余风险"]

    acceptance_evidence_lines = [
        f"输出文件路径：{safe_markdown_text(output_doc, max_length=None)}",
    ]
    if execution_log and execution_log.get("source_sha256") and execution_log.get("output_sha256"):
        acceptance_evidence_lines.append("原始文件未覆盖证明：执行日志中保留源文件与输出文件 hash。")
    elif any("原始文件" in str(item) for item in final_acceptance.get("evidence", [])):
        acceptance_evidence_lines.append("原始文件未覆盖证明：见 final_acceptance 证据记录。")
    else:
        acceptance_evidence_lines.append("原始文件未覆盖证明：未发现原始文件被覆盖迹象。")
    if reviews:
        blocked_count = sum(1 for review in reviews if review.get("status") == "blocked")
        passed_count = len(reviews) - blocked_count
        acceptance_evidence_lines.append(f"二轮复核摘要：共 {len(reviews)} 项，其中通过 {passed_count} 项，阻断 {blocked_count} 项。")
    elif mode == "audit-only":
        acceptance_evidence_lines.append("二轮复核摘要：本次为仅审计流程，未生成二轮复核产物。")
    else:
        acceptance_evidence_lines.append("二轮复核摘要：未生成二轮复核产物。")

    if status == "accepted":
        final_conclusion = "本次格式治理已完成，核心格式项已通过验收。"
        next_steps_lines = ["无需进一步操作"]
    elif status == "accepted_with_warnings":
        final_conclusion = "本次处理已完成，但仍存在需要关注的风险或限制。"
        next_steps_lines = ["请关注风险和限制章节后再交付使用。"]
    else:
        final_conclusion = "本次处理未通过最终验收，请先处理阻断项。"
        next_steps_lines = ["请先处理阻断项后重新生成报告"]

    return {
        "status_marker": status_marker(status, use_icons=use_icons),
        "status_label": status_label,
        "final_conclusion": final_conclusion,
        "input_and_rule_section": markdown_list(input_and_rule_lines, empty_text="无输入与规则来源"),
        "audit_summary_section": markdown_list(audit_summary_lines, empty_text="未发现格式问题"),
        "repair_summary_section": markdown_list(repair_summary_lines, empty_text="本次未执行自动修复"),
        "before_after_section": markdown_list(before_after_lines, empty_text="无可展示的修复前后对比项"),
        "unfixed_items_section": markdown_list(unfixed_lines, empty_text="无未修复项"),
        "remaining_risks_section": markdown_list(risk_lines, empty_text="无已知剩余风险"),
        "acceptance_evidence_section": markdown_list(acceptance_evidence_lines, empty_text="无验收证据"),
        "next_steps_section": markdown_list(next_steps_lines, empty_text="无需进一步操作"),
    }


def render_diff_summary(run_dir: Path) -> str:
    """渲染内部差异摘要。"""
    before = load_json_if_exists(run_dir / "snapshots" / "document_snapshot.before.json")
    after = load_json_if_exists(run_dir / "snapshots" / "document_snapshot.after.json")
    if before is None or after is None:
        return "## 快照差异\n\n- 无可展示的修复前后对比项。"
    return "\n".join(
        [
            "## 快照差异",
            "",
            f"- before hash：`{before.get('document_hash')}`",
            f"- after hash：`{after.get('document_hash')}`",
            f"- 段落数：{before.get('paragraph_count')} -> {after.get('paragraph_count')}",
            f"- 表格数：{before.get('table_count')} -> {after.get('table_count')}",
        ]
    )


def render_repair_log(run_dir: Path) -> str:
    """渲染内部修复日志。"""
    log = load_json_if_exists(run_dir / "logs" / "repair_execution.json")
    if log is None:
        return "## 执行摘要\n\n- 本次未执行自动修复。"
    lines = [
        "## 执行摘要",
        "",
        f"- 输入文件：`{log.get('source_docx') or log.get('working_docx')}`",
        f"- 输出文件：`{log.get('output_docx')}`",
        f"- 输出有效：{log.get('output_docx_valid')}",
    ]
    if isinstance(log.get("counts"), dict):
        lines.extend(["", "## 统计", ""])
        for key, value in log["counts"].items():
            lines.append(f"- {key}：{value}")
    return "\n".join(lines)


def write_optional_reports(run_dir: Path) -> None:
    """按现有产物补写内部追溯报告。"""
    reports_dir = run_dir / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)
    if (run_dir / "logs" / "repair_execution.json").exists():
        (reports_dir / "REPAIR_LOG.md").write_text("# 修复日志\n\n" + render_repair_log(run_dir).rstrip() + "\n", encoding="utf-8")
    if (run_dir / "snapshots" / "document_snapshot.before.json").exists():
        (reports_dir / "DIFF_SUMMARY.md").write_text("# 差异摘要\n\n" + render_diff_summary(run_dir).rstrip() + "\n", encoding="utf-8")


def write_blocked_state(run_dir: Path, final_acceptance: dict[str, Any], *, error_message: str) -> None:
    """在报告失败时写 blocked 状态，不覆盖旧报告。"""
    blocked = dict(final_acceptance)
    blocked["accepted"] = False
    blocked["status"] = "blocked"
    blocked.setdefault("created_at", datetime.now(TZ).isoformat())
    open_blockers = list(blocked.get("open_blockers") or [])
    open_blockers.append(error_message)
    blocked["open_blockers"] = open_blockers

    final_path = run_dir / "logs" / "final_acceptance.json"
    write_json_atomic(final_path, blocked)

    execution_log = load_json_if_exists(run_dir / "logs" / "repair_execution.json")
    output_docx = execution_log.get("output_docx") if execution_log else None
    write_yaml(
        run_dir / "logs" / "state.yaml",
        {
            "schema_version": "1.0.0",
            "run_id": run_dir.name,
            "run_dir": str(run_dir).replace("\\", "/"),
            "stage": "final_acceptance",
            "state": "blocked",
            "updated_at": datetime.now(TZ).isoformat(),
            "output_docx": output_docx,
            "final_acceptance": str(final_path).replace("\\", "/"),
            "next_action": "retry_report_render",
        },
    )


def render_reports(run_dir: Path) -> dict[str, Any]:
    """生成最终交付报告。extract-rule 模式不生成该报告。"""
    reports_dir = run_dir / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)

    seed_final_acceptance = load_json_if_exists(run_dir / "logs" / "final_acceptance.json") or {}
    mode, mode_recognized = detect_mode_details(run_dir, seed_final_acceptance)
    if mode == "extract-rule":
        raise ValueError("extract-rule 模式不生成最终交付报告，请查看 RULE_SUMMARY.md")

    final_acceptance = load_or_build_final_acceptance(run_dir, mode=mode)
    write_optional_reports(run_dir)

    try:
        final_view_model = build_final_report_view_model(
            run_dir,
            final_acceptance,
            mode=mode,
            use_icons=True,
            mode_recognized=mode_recognized,
        )
        final_report_content = render_template(TEMPLATE_PATH.read_text(encoding="utf-8"), final_view_model)
        assert_human_readable_report(
            final_report_content,
            report_kind="final_report",
            required_sections=FINAL_REPORT_REQUIRED_SECTIONS,
        )
    except ValueError as exc:
        write_blocked_state(run_dir, final_acceptance, error_message=str(exc))
        raise

    final_report_path = reports_dir / "FINAL_ACCEPTANCE_REPORT.md"
    temp_report = final_report_path.with_suffix(final_report_path.suffix + ".tmp")
    temp_report.write_text(final_report_content.rstrip() + "\n", encoding="utf-8")
    os.replace(temp_report, final_report_path)

    updated_final = dict(final_acceptance)
    updated_final.setdefault("created_at", datetime.now(TZ).isoformat())
    updated_final["reports"] = list(dict.fromkeys(list(updated_final.get("reports") or []) + [str(final_report_path).replace("\\", "/")]))

    final_path = run_dir / "logs" / "final_acceptance.json"
    write_json_atomic(final_path, updated_final)

    execution_log = load_json_if_exists(run_dir / "logs" / "repair_execution.json")
    output_docx = execution_log.get("output_docx") if execution_log else None
    write_yaml(
        run_dir / "logs" / "state.yaml",
        {
            "schema_version": "1.0.0",
            "run_id": run_dir.name,
            "run_dir": str(run_dir).replace("\\", "/"),
            "stage": "final_acceptance",
            "state": normalize_status(updated_final),
            "updated_at": str(updated_final.get("created_at")),
            "output_docx": output_docx,
            "final_acceptance": str(final_path).replace("\\", "/"),
            "next_action": "done" if normalize_status(updated_final) == "accepted" else "manual_recover",
        },
    )
    return updated_final


def main() -> int:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="生成最终交付报告")
    parser.add_argument("--run-dir", required=True, type=Path)
    args = parser.parse_args()
    render_reports(args.run_dir)
    print(args.run_dir / "reports")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
