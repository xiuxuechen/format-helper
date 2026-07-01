#!/usr/bin/env python3
"""生成 CODE-006 T01-T06 二轮复核结果。"""

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

from scripts.utils.simple_yaml import load_yaml
TZ = timezone(timedelta(hours=8))
REVIEW_TASKS = [
    ("T01", "输出 DOCX OOXML 完整性复核"),
    ("T02", "before/after 快照完整性复核"),
    ("T03", "修复执行日志复核"),
    ("T04", "自动修复动作追溯复核"),
    ("T05", "人工确认项留痕复核"),
    ("T06", "渲染证据复核"),
]
MIN_RENDER_PAGE_BYTES = 50_000


def load_json(path: Path) -> dict[str, Any]:
    """读取 JSON 文件。"""
    return json.loads(path.read_text(encoding="utf-8"))


def latest_finalized_plan(run_dir: Path) -> Path:
    """返回最新 finalized revision，禁止静默读取 repair_plan.yaml 别名。"""
    candidates = sorted(run_dir.glob("plans/repair_plan.finalized.r*.yaml"))
    if not candidates:
        raise FileNotFoundError("缺少 plans/repair_plan.finalized.rNNN.yaml")
    return candidates[-1]


def snapshot_v2_errors(snapshot: dict[str, Any], expected_kind: str) -> list[str]:
    """校验二轮复核所需的 OfficeCLI snapshot v2 最小 Gate。"""
    errors: list[str] = []
    if snapshot.get("schema_id") != "officecli-document-snapshot":
        errors.append("schema_id 必须为 officecli-document-snapshot")
    if snapshot.get("schema_version") != "2.0.0" or snapshot.get("contract_version") != "v5":
        errors.append("snapshot 必须为 v5/2.0.0")
    if snapshot.get("kind") != expected_kind:
        errors.append(f"snapshot.kind 必须为 {expected_kind}")
    if snapshot.get("gate_check", {}).get("status") != "passed":
        errors.append("snapshot gate_check 未通过")
    if not isinstance(snapshot.get("nodes"), list) or not isinstance(snapshot.get("indexes"), dict):
        errors.append("snapshot nodes/indexes 缺失")
    return errors


def select_render_pages(run_dir: Path) -> tuple[Path, list[Path]]:
    """选择最新且包含页面图片的渲染证据目录。"""
    candidates: list[tuple[float, Path, list[Path]]] = []
    for render_dir in run_dir.glob("render*"):
        if not render_dir.is_dir():
            continue
        pages = sorted(render_dir.glob("page-*.png"))
        if not pages:
            continue
        latest_page_time = max(page.stat().st_mtime for page in pages)
        candidates.append((latest_page_time, render_dir, pages))
    if not candidates:
        return run_dir / "render", []
    _, render_dir, pages = max(candidates, key=lambda item: item[0])
    return render_dir, pages


def make_result(task_id: str, task_name: str, status: str, evidence: list[str], issues: list[dict[str, Any]]) -> dict[str, Any]:
    """构造复核结果。"""
    return {
        "schema_version": "1.0.0",
        "task_id": task_id,
        "task_name": task_name,
        "status": status,
        "checked_at": datetime.now(TZ).isoformat(),
        "evidence": evidence,
        "issues": issues,
    }


def build_reviews(run_dir: Path) -> list[dict[str, Any]]:
    """基于运行目录生成 T01-T06 复核结果。"""
    repair_plan_path = latest_finalized_plan(run_dir)
    execution_log_path = run_dir / "logs" / "repair_execution.json"
    before_snapshot_path = run_dir / "snapshots" / "officecli-document-snapshot.before.json"
    after_snapshot_path = run_dir / "snapshots" / "officecli-document-snapshot.after.json"
    repair_plan = load_yaml(repair_plan_path)
    execution_log = load_json(execution_log_path)
    before_snapshot = load_json(before_snapshot_path)
    after_snapshot = load_json(after_snapshot_path)
    output_docx = Path(execution_log.get("output_docx") or execution_log.get("working_docx") or "")

    reviews: list[dict[str, Any]] = []

    valid_docx = bool(execution_log.get("output_docx_valid")) and output_docx.is_file()
    reviews.append(
        make_result(
            "T01",
            "输出 DOCX OOXML 完整性复核",
            "passed" if valid_docx else "blocked",
            [f"输出文件：{output_docx}", f"OfficeCLI validate 通过：{valid_docx}"],
            [] if valid_docx else [{"issue_id": "T01-I001", "severity": "blocker", "description": "输出 DOCX 未通过 OfficeCLI 写后校验"}],
        )
    )

    before_errors = snapshot_v2_errors(before_snapshot, "before")
    after_errors = snapshot_v2_errors(after_snapshot, "after")
    snapshot_status = "passed" if not before_errors and not after_errors else "blocked"
    reviews.append(
        make_result(
            "T02",
            "before/after 快照完整性复核",
            snapshot_status,
            [
                f"before 节点数：{before_snapshot.get('document', {}).get('node_count')}",
                f"after 节点数：{after_snapshot.get('document', {}).get('node_count')}",
                f"after 快照类型：{after_snapshot.get('kind')}",
            ],
            [
                {"issue_id": f"T02-I{index:03d}", "severity": "blocker", "description": error}
                for index, error in enumerate(before_errors + after_errors, start=1)
            ],
        )
    )

    rejected = int(execution_log.get("actions_rejected", 0))
    output_valid = bool(execution_log.get("output_docx_valid"))
    execution_blocked = rejected > 0 or not output_valid
    reviews.append(
        make_result(
            "T03",
            "修复执行日志复核",
            "blocked" if execution_blocked else "passed",
            [
                f"动作总数：{execution_log.get('actions_total')}",
                f"已执行：{execution_log.get('actions_executed')}",
                f"跳过：{execution_log.get('actions_skipped')}",
                f"拒绝：{execution_log.get('actions_rejected')}",
            ],
            []
            if not execution_blocked
            else [{"issue_id": "T03-I001", "severity": "blocker", "description": "存在拒绝动作或输出 DOCX 无效"}],
        )
    )

    auto_actions = [action for action in repair_plan.get("actions", []) if action.get("auto_fix_policy") == "auto-fix"]
    executed_ids = {
        item.get("action_id")
        for item in execution_log.get("actions", [])
        if item.get("status") == "executed"
    }
    missing_executed = [action.get("action_id") for action in auto_actions if action.get("action_id") not in executed_ids]
    reviews.append(
        make_result(
            "T04",
            "自动修复动作追溯复核",
            "passed" if not missing_executed else "blocked",
            [
                f"auto-fix 动作数：{len(auto_actions)}",
                f"已执行动作：{', '.join(sorted(executed_ids)) if executed_ids else '无'}",
            ],
            [
                {"issue_id": "T04-I001", "severity": "blocker", "description": f"auto-fix 动作未执行：{', '.join(missing_executed)}"}
            ]
            if missing_executed
            else [],
        )
    )

    manual_items = repair_plan.get("manual_review_items", [])
    reviews.append(
        make_result(
            "T05",
            "人工确认项留痕复核",
            "passed_with_warnings" if manual_items else "passed",
            [f"人工确认项数量：{len(manual_items)}"],
            [
                {
                    "issue_id": item.get("item_id", f"T05-I{index:03d}"),
                    "severity": "warning",
                    "description": item.get("reason", "存在人工确认项"),
                }
                for index, item in enumerate(manual_items, start=1)
            ],
        )
    )

    render_dir, render_pages = select_render_pages(run_dir)
    blank_like_pages = [page.name for page in render_pages if page.stat().st_size < MIN_RENDER_PAGE_BYTES]
    render_issues = []
    if not render_pages:
        render_issues.append({"issue_id": "T06-I001", "severity": "blocker", "description": "未找到渲染页 PNG"})
    if blank_like_pages:
        render_issues.append(
            {
                "issue_id": "T06-I002",
                "severity": "blocker",
                "description": f"疑似空白渲染页：{', '.join(blank_like_pages)}",
            }
        )
    reviews.append(
        make_result(
            "T06",
            "渲染证据复核",
            "passed" if not render_issues else "blocked",
            [
                f"渲染目录：{render_dir}",
                f"渲染页数量：{len(render_pages)}",
                f"疑似空白页数量：{len(blank_like_pages)}",
            ],
            render_issues,
        )
    )

    return reviews


def main() -> int:
    """兼容旧入口，实际执行统一的 v5 review-result 构建器。"""
    from scripts.officecli.review_builder import main as review_main

    parser = argparse.ArgumentParser(description="生成 review-result 2.0.0")
    parser.add_argument("--run-dir", required=True, type=Path)
    parser.add_argument("--output-dir", type=Path)
    args = parser.parse_args()
    output = (args.output_dir / "final_review.json") if args.output_dir else None
    argv = ["--run-dir", str(args.run_dir)]
    if output is not None:
        argv.extend(["--output", str(output)])
    return review_main(argv)


if __name__ == "__main__":
    raise SystemExit(main())
