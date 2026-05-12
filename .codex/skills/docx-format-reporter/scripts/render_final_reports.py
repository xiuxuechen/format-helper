#!/usr/bin/env python3
"""生成 CODE-006 最终报告、final_acceptance.json 和 state.yaml。"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[4]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.utils.simple_yaml import load_yaml, write_yaml
from scripts.validation.final_acceptance import validate_final_acceptance_v4


TZ = timezone(timedelta(hours=8))
REPORT_NAMES = [
    "AUDIT_REPORT.md",
    "REVIEW_REPORT.md",
    "MANUAL_CONFIRMATION.md",
    "DIFF_SUMMARY.md",
    "REPAIR_LOG.md",
    "FINAL_ACCEPTANCE_REPORT.md",
]


def load_json(path: Path) -> dict[str, Any]:
    """读取 JSON 文件。"""
    return json.loads(path.read_text(encoding="utf-8"))


def review_files(run_dir: Path) -> list[Path]:
    """列出复核结果文件。"""
    return sorted((run_dir / "review_results").glob("T*.review.json"))


def load_reviews(run_dir: Path) -> list[dict[str, Any]]:
    """读取复核结果。"""
    return [load_json(path) for path in review_files(run_dir)]


def collect_blockers(reviews: list[dict[str, Any]]) -> list[str]:
    """收集阻塞项。"""
    blockers: list[str] = []
    for review in reviews:
        if review.get("status") != "blocked":
            continue
        for issue in review.get("issues", []):
            blockers.append(f"{review.get('task_id')} {issue.get('description')}")
    return blockers


def build_final_acceptance(run_dir: Path) -> dict[str, Any]:
    """生成最终验收 JSON 数据。"""
    repair_plan = load_yaml(run_dir / "plans" / "repair_plan.yaml")
    execution_log = load_json(run_dir / "logs" / "repair_execution.json")
    reviews = load_reviews(run_dir)
    blockers = collect_blockers(reviews)
    output_valid = bool(execution_log.get("output_docx_valid"))
    if not output_valid:
        blockers.append("输出 DOCX 未通过 OOXML 有效性检查")
    accepted = not blockers and len(reviews) == 6
    reports = [str(run_dir / "reports" / name).replace("\\", "/") for name in REPORT_NAMES]
    evidence = [
        f"输出 DOCX：{execution_log.get('output_docx')}",
        f"执行动作：{execution_log.get('actions_executed')} executed, {execution_log.get('actions_skipped')} skipped, {execution_log.get('actions_rejected')} rejected",
        f"二轮复核任务数：{len(reviews)}",
        "document_snapshot.after.json 已生成",
    ]
    return {
        "schema_version": "1.0.0",
        "accepted": accepted,
        "status": "accepted" if accepted else "blocked",
        "created_at": datetime.now(TZ).isoformat(),
        "open_blockers": blockers,
        "manual_items_remaining": repair_plan.get("manual_review_items", []),
        "output_docx_valid": output_valid,
        "reports": reports,
        "evidence": evidence,
    }


def write_report(path: Path, title: str, body: str) -> None:
    """写入 Markdown 报告。"""
    path.write_text(f"# {title}\n\n{body.rstrip()}\n", encoding="utf-8")


def render_audit_report(run_dir: Path, final: dict[str, Any]) -> str:
    """渲染审计报告。"""
    repair_plan = load_yaml(run_dir / "plans" / "repair_plan.yaml")
    actions = repair_plan.get("actions", [])
    auto_fix = [item for item in actions if item.get("auto_fix_policy") == "auto-fix"]
    manual = repair_plan.get("manual_review_items", [])
    return "\n".join(
        [
            "## 审计范围",
            "",
            f"- 运行目录：`{run_dir}`",
            f"- 修复动作总数：{len(actions)}",
            f"- 自动修复动作：{len(auto_fix)}",
            f"- 人工确认项：{len(manual)}",
            "",
            "## 结论",
            "",
            "修复计划已按白名单、置信度、风险等级和语义证据生成；低置信度表格动作保留为人工确认项。",
        ]
    )


def render_review_report(reviews: list[dict[str, Any]]) -> str:
    """渲染复核报告。"""
    lines = ["## T01-T06 复核结果", ""]
    for review in reviews:
        lines.append(f"- {review['task_id']} {review['task_name']}：{review['status']}")
        for issue in review.get("issues", []):
            lines.append(f"  - {issue.get('severity')}: {issue.get('description')}")
    return "\n".join(lines)


def render_manual_confirmation(final: dict[str, Any]) -> str:
    """渲染人工确认清单。"""
    items = final.get("manual_items_remaining", [])
    if not items:
        return "## 人工确认项\n\n无。"
    lines = ["## 人工确认项", ""]
    for item in items:
        lines.append(f"- {item.get('item_id')}：{item.get('reason')}")
        lines.append(f"  - 默认处理：{item.get('default_option', 'manual-review')}")
        lines.append(f"  - 需要确认：{item.get('required_decision', '是否允许自动处理')}")
    return "\n".join(lines)


def render_diff_summary(run_dir: Path) -> str:
    """渲染差异摘要。"""
    before = load_json(run_dir / "snapshots" / "document_snapshot.before.json")
    after = load_json(run_dir / "snapshots" / "document_snapshot.after.json")
    return "\n".join(
        [
            "## 快照差异",
            "",
            f"- before hash：`{before.get('document_hash')}`",
            f"- after hash：`{after.get('document_hash')}`",
            f"- 段落数：{before.get('paragraph_count')} -> {after.get('paragraph_count')}",
            f"- 表格数：{before.get('table_count')} -> {after.get('table_count')}",
            f"- 节数量：{before.get('section_count')} -> {after.get('section_count')}",
        ]
    )


def render_repair_log(run_dir: Path) -> str:
    """渲染修复日志。"""
    log = load_json(run_dir / "logs" / "repair_execution.json")
    lines = [
        "## 执行摘要",
        "",
        f"- 工作副本：`{log.get('working_docx')}`",
        f"- 输出文件：`{log.get('output_docx')}`",
        f"- 动作总数：{log.get('actions_total')}",
        f"- 已执行：{log.get('actions_executed')}",
        f"- 已跳过：{log.get('actions_skipped')}",
        f"- 已拒绝：{log.get('actions_rejected')}",
        "",
        "## 动作明细",
        "",
    ]
    for item in log.get("actions", []):
        lines.append(f"- {item.get('action_id')} {item.get('action_type', '')}：{item.get('status')}，{item.get('reason')}")
    return "\n".join(lines)


def render_final_acceptance(final: dict[str, Any]) -> str:
    """渲染最终验收报告。"""
    lines = [
        "## 验收结论",
        "",
        f"- 结论：{'通过' if final.get('accepted') else '未通过'}",
        f"- 状态：{final.get('status')}",
        f"- 输出 DOCX 有效：{final.get('output_docx_valid')}",
        f"- 剩余人工确认项：{len(final.get('manual_items_remaining', []))}",
        "",
        "## 验收依据",
        "",
    ]
    for item in final.get("evidence", []):
        lines.append(f"- {item}")
    if final.get("open_blockers"):
        lines.extend(["", "## 阻塞项", ""])
        for item in final["open_blockers"]:
            lines.append(f"- {item}")
    return "\n".join(lines)


def render_reports(run_dir: Path) -> dict[str, Any]:
    """生成全部报告和验收状态。"""
    reports_dir = run_dir / "reports"
    reports_dir.mkdir(parents=True, exist_ok=True)
    final = build_final_acceptance(run_dir)
    reviews = load_reviews(run_dir)
    report_bodies = {
        "AUDIT_REPORT.md": ("审计报告", render_audit_report(run_dir, final)),
        "REVIEW_REPORT.md": ("复核报告", render_review_report(reviews)),
        "MANUAL_CONFIRMATION.md": ("人工确认清单", render_manual_confirmation(final)),
        "DIFF_SUMMARY.md": ("差异摘要", render_diff_summary(run_dir)),
        "REPAIR_LOG.md": ("修复日志", render_repair_log(run_dir)),
        "FINAL_ACCEPTANCE_REPORT.md": ("最终验收报告", render_final_acceptance(final)),
    }
    for name, (title, body) in report_bodies.items():
        write_report(reports_dir / name, title, body)

    final_path = run_dir / "logs" / "final_acceptance.json"
    final_path.write_text(json.dumps(final, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    errors = validate_final_acceptance_v4(final)
    if errors:
        raise SystemExit("\n".join(errors))
    write_yaml(
        run_dir / "logs" / "state.yaml",
        {
            "schema_version": "1.0.0",
            "run_dir": str(run_dir).replace("\\", "/"),
            "state": final["status"],
            "updated_at": final["created_at"],
            "output_docx": load_json(run_dir / "logs" / "repair_execution.json").get("output_docx"),
            "final_acceptance": str(final_path).replace("\\", "/"),
        },
    )
    return final


def main() -> int:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="生成 CODE-006 最终报告")
    parser.add_argument("--run-dir", required=True, type=Path)
    args = parser.parse_args()
    render_reports(args.run_dir)
    print(args.run_dir / "reports")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
