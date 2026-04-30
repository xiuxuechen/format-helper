#!/usr/bin/env python3
"""从合并审计结果生成 repair_plan.yaml 初稿。"""
from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone, timedelta
from pathlib import Path


TZ = timezone(timedelta(hours=8))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--merged-audit", required=True)
    parser.add_argument("--plan-id", required=True)
    parser.add_argument("--snapshot", required=True)
    parser.add_argument("--rule-id", required=True)
    parser.add_argument("--rule-version", default="1.0.0")
    parser.add_argument("--source-docx", required=True)
    parser.add_argument("--working-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    data = json.loads(Path(args.merged_audit).read_text(encoding="utf-8"))
    now = datetime.now(TZ)
    lines = [
        "schema_version: 1.0.0",
        f"repair_plan_id: docx-repair-plan-{now.strftime('%Y%m%d-%H%M%S')}",
        f"created_at: {now.isoformat()}",
        f"based_on_plan_id: {args.plan_id}",
        f"based_on_snapshot: {args.snapshot}",
        "",
        "rule_profile:",
        f"  id: {args.rule_id}",
        f"  version: {args.rule_version}",
        "",
        f"source_docx: {args.source_docx}",
        f"working_docx: {args.working_docx}",
        f"output_docx: {args.output_docx}",
        "",
        "conflict_resolution: []",
        "actions:",
    ]

    manual_items = []
    action_index = 1
    for issue in data.get("issues", []):
        action = issue.get("recommended_action", {})
        policy = action.get("auto_fix_policy")
        issue_id = issue.get("issue_id", "")
        if policy == "auto-fix":
            action_type = action.get("action_type", "manual_review")
            strategy = "direct-format-override" if action_type.endswith("direct_format") else "style-definition"
            lines.extend(
                [
                    f"  - action_id: A{action_index:03d}",
                    "    source_issue_ids:",
                    f"      - {issue_id}",
                    f"    action_type: {action_type}",
                    f"    format_write_strategy: {strategy}",
                    "    target:",
                    f"      element_id: {issue.get('element_ref', {}).get('element_id', '')}",
                    f"      expected_role: {issue.get('expected_role', '')}",
                    "    before: {}",
                    "    after: {}",
                    "    auto_fix_policy: auto-fix",
                    f"    risk_level: {issue.get('severity', 'medium')}",
                    "    status: pending",
                ]
            )
            action_index += 1
        else:
            manual_items.append(issue)

    if action_index == 1:
        lines.append("  []")

    lines.append("manual_review_items:")
    if manual_items:
        for index, issue in enumerate(manual_items, start=1):
            lines.extend(
                [
                    f"  - item_id: M{index:03d}",
                    "    source_issue_ids:",
                    f"      - {issue.get('issue_id', '')}",
                    "    element_ref:",
                    f"      element_id: {issue.get('element_ref', {}).get('element_id', '')}",
                    f"    reason: {issue.get('problem', '需要人工确认')}",
                    "    required_decision: 是否允许自动处理该项",
                ]
            )
    else:
        lines.append("  []")

    lines.extend(
        [
            "execution_order:",
            "  - normalize_styles",
            "  - apply_page_section_rules",
            "  - apply_heading_styles",
            "  - apply_body_styles",
            "  - apply_table_safe_fixes",
            "  - replace_or_insert_auto_toc",
            "  - refresh_fields_or_mark_for_update",
            "  - save_repaired_docx",
            "post_repair:",
            "  generate_after_snapshot: true",
            "  dispatch_second_round_review: true",
            "  required_review_task_ids: [T01, T02, T03, T04, T05, T06]",
        ]
    )
    Path(args.output).write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
