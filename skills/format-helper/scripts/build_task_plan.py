#!/usr/bin/env python3
"""根据文档快照和规则版本生成 PLAN.yaml。"""
from __future__ import annotations

import argparse
from datetime import datetime, timezone, timedelta
from pathlib import Path


TZ = timezone(timedelta(hours=8))


TASKS = [
    ("T01", "title-agent", "总题目、封面、副标题", ["cover-title", "subtitle"]),
    ("T02", "toc-agent", "静态目录识别、自动目录生成、目录层级检查", ["toc-title", "toc-static-item", "toc-field", "heading-level-1", "heading-level-2", "heading-level-3"]),
    ("T03", "heading-agent", "标题层级、编号连续性、标准标题样式和大纲级别", ["heading-level-1", "heading-level-2", "heading-level-3", "heading-level-4", "ambiguous-numbered-item"]),
    ("T04", "body-agent", "正文段落、字体、字号、行距、首行缩进、直接格式覆盖", ["body-paragraph", "body-no-indent", "emphasis-run"]),
    ("T05", "table-agent", "普通表格、专栏表格、附表、表头、表格正文、横向页表格", ["normal-table", "special-panel", "appendix-table", "table-title", "table-header", "table-body"]),
    ("T06", "page-agent", "页面设置、节、横向页、页眉页脚、页码、脚注", ["page-section", "header-footer", "footnote"]),
]


def yaml_list(items, indent: int) -> str:
    space = " " * indent
    return "\n".join(f"{space}- {item}" for item in items)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--run-dir", required=True)
    parser.add_argument("--source-docx", required=True)
    parser.add_argument("--working-docx", required=True)
    parser.add_argument("--snapshot", required=True)
    parser.add_argument("--rule-id", required=True)
    parser.add_argument("--rule-version", default="1.0.0")
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    now = datetime.now(TZ)
    run_dir = Path(args.run_dir)
    plan_id = f"docx-format-plan-{now.strftime('%Y%m%d-%H%M%S')}"
    lines = [
        "schema_version: 1.0.0",
        f"plan_id: {plan_id}",
        f"created_at: {now.isoformat()}",
        "",
        "document:",
        f"  source_path: {args.source_docx}",
        f"  working_copy_path: {args.working_docx}",
        f"  baseline_snapshot_path: {args.snapshot}",
        "  document_type_guess: null",
        "  page_count_estimate: null",
        "",
        "rule_profile:",
        f"  id: {args.rule_id}",
        f"  version: {args.rule_version}",
        "  confirmed_by: user",
        f"  confirmed_at: {now.isoformat()}",
        "",
        "global_objectives:",
        "  auto_toc: true",
        "  style_driven: true",
        "  preserve_original_document: true",
        "  output_repaired_copy: true",
        "  output_audit_report: true",
        "  output_review_report: true",
        "",
        "global_constraints:",
        "  child_agents_may_modify_docx: false",
        "  master_thread_only_writes_docx: true",
        "  require_second_round_review: true",
        "  require_final_structure_validation: true",
        "  require_final_render_validation: true",
        "",
        "artifacts:",
        f"  audit_results_dir: {run_dir / 'audit_results'}",
        f"  repair_plan_path: {run_dir / 'plans' / 'repair_plan.yaml'}",
        f"  repaired_docx_path: {run_dir / 'output'}",
        f"  after_snapshot_path: {run_dir / 'snapshots' / 'document_snapshot.after.json'}",
        f"  review_results_dir: {run_dir / 'review_results'}",
        f"  final_report_path: {run_dir / 'reports' / 'FINAL_ACCEPTANCE_REPORT.md'}",
        "",
        "tasks:",
    ]

    for task_id, owner, scope, element_types in TASKS:
        lines.extend(
            [
                f"  - task_id: {task_id}",
                f"    owner: {owner}",
                f"    scope: {scope}",
                "    element_types:",
                yaml_list(element_types, 6),
                f"    audit_output_path: {run_dir / 'audit_results' / (task_id + '.audit.json')}",
                f"    review_output_path: {run_dir / 'review_results' / (task_id + '.review.json')}",
                "    status:",
                "      audit: pending",
                "      repair: pending",
                "      review: pending",
            ]
        )

    lines.extend(
        [
            "",
            "acceptance:",
        "  pass_condition:",
        "    - all_required_tasks_review_status_in: [passed, manual_accepted]",
        "    - no_blocker_issues",
        "    - automatic_toc_present",
        "    - required_heading_outline_levels_set",
        "    - required_real_format_checks_passed",
            "  failure_condition:",
            "    - any_required_task_review_status_in: [failed, blocked]",
            "    - repaired_docx_unreadable",
            "    - automatic_toc_missing",
        ]
    )
    Path(args.output).write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
