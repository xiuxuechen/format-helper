#!/usr/bin/env python3
"""从格式治理产物生成用户可读中文报告。"""
from __future__ import annotations

import argparse
import json
from pathlib import Path


TASK_LABELS = {
    "T01": "封面与题名检查",
    "T02": "目录检查",
    "T03": "标题结构检查",
    "T04": "正文格式检查",
    "T05": "表格检查",
    "T06": "页面设置检查",
}

STYLE_LABELS = {
    "Normal": "普通正文格式",
    "Heading1": "Word 一级标题样式",
    "Heading2": "Word 二级标题样式",
    "Heading3": "Word 三级标题样式",
    "Heading 1": "Word 一级标题样式",
    "Heading 2": "Word 二级标题样式",
    "Heading 3": "Word 三级标题样式",
    "标题 1": "Word 一级标题样式",
    "标题 2": "Word 二级标题样式",
    "标题 3": "Word 三级标题样式",
    "table-original": "原表格文字格式",
    None: "未检测到明确格式",
    "": "未检测到明确格式",
}

ACTION_LABELS = {
    "map_heading_native_style": "映射为 Word 原生标题样式",
    "apply_body_style_definition": "统一正文基础样式定义",
    "apply_body_direct_format": "按确认策略写入局部正文格式",
    "apply_table_cell_format": "统一表格单元格格式",
    "toc_content_audit": "校验目录内容差异",
    "insert_or_replace_toc_field": "替换为可自动更新目录",
    "manual_review": "需要人工或专项能力确认",
}

SEVERITY_LABELS = {
    "blocker": "阻塞",
    "high": "高",
    "medium": "中",
    "low": "低",
}


def read_json(path: Path) -> dict:
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def read_template(name: str) -> str:
    path = Path(__file__).resolve().parents[1] / "references" / "report_templates" / name
    return path.read_text(encoding="utf-8")


def render_template(name: str, values: dict[str, str]) -> str:
    output = read_template(name)
    for key, value in values.items():
        output = output.replace("{{" + key + "}}", str(value))
    return output


def task_label(task_id: str) -> str:
    return TASK_LABELS.get(task_id, "其他检查")


def style_label(style: str | None) -> str:
    if style in STYLE_LABELS:
        return STYLE_LABELS[style]
    return "原文自带格式"


def action_label(action_type: str | None) -> str:
    return ACTION_LABELS.get(action_type or "", "按规则处理")


def severity_label(severity: str | None) -> str:
    return SEVERITY_LABELS.get(severity or "", "未分级")


def preview_text(element_ref: dict) -> str:
    text = markdown_cell(element_ref.get("text_preview") or "").strip()
    return text if text else "无文本摘录"


def markdown_cell(value) -> str:
    return str(value or "").replace("|", "\\|").replace("\n", " ").strip()


def issue_current(issue: dict) -> str:
    current = issue.get("current_format", {})
    return style_label(current.get("style"))


def issue_expected(issue: dict) -> str:
    expected = issue.get("expected_format", {})
    expected_style = expected.get("style")
    if expected_style:
        return style_label(expected_style)
    problem = issue.get("problem", "")
    if "自动目录" in problem or issue.get("element_ref", {}).get("element_type") == "toc-field":
        return "文档应包含可自动更新目录"
    return "符合已确认规则"


def issue_action(issue: dict) -> str:
    action = issue.get("recommended_action", {})
    problem = issue.get("problem", "")
    if "自动目录" in problem:
        return "生成或插入可自动更新目录"
    return action_label(action.get("action_type"))


def collect_audits(run_dir: Path) -> list[dict]:
    rows = []
    for path in sorted(path for path in (run_dir / "audit_results").glob("*.audit.json") if path.name != "merged.audit.json"):
        data = read_json(path)
        rows.append(
            {
                "task_id": data.get("task_id", path.stem.split(".")[0]),
                "label": task_label(data.get("task_id", path.stem.split(".")[0])),
                "summary": data.get("summary", {}),
                "issues": data.get("issues", []),
            }
        )
    return rows


def collect_reviews(run_dir: Path) -> list[dict]:
    rows = []
    for path in sorted((run_dir / "review_results").glob("*.review.json")):
        data = read_json(path)
        rows.append(
            {
                "task_id": data.get("task_id", path.stem.split(".")[0]),
                "label": task_label(data.get("task_id", path.stem.split(".")[0])),
                "summary": data.get("summary", {}),
                "issues": data.get("issues", []),
            }
        )
    return rows


def load_run_log(run_dir: Path) -> dict:
    path = run_dir / "logs" / "run_log.yaml"
    data = {}
    if not path.exists():
        return data
    for raw in path.read_text(encoding="utf-8").splitlines():
        if ":" not in raw or raw.startswith(" "):
            continue
        key, value = raw.split(":", 1)
        data[key.strip()] = value.strip()
    return data


def load_repair_log(run_dir: Path) -> dict:
    path = run_dir / "logs" / "repair_log.yaml"
    data = {}
    if not path.exists():
        return data
    for raw in path.read_text(encoding="utf-8").splitlines():
        if ":" not in raw or raw.startswith(" "):
            continue
        key, value = raw.split(":", 1)
        data[key.strip()] = value.strip()
    return data


def load_snapshot(path: Path) -> dict:
    return read_json(path)


def load_page_map(run_dir: Path, snapshot: dict) -> dict:
    raw = read_json(run_dir / "logs" / "page_map.after.json")
    paragraphs = {str(key): value for key, value in (raw.get("paragraphs") or {}).items()}
    tables = {}
    for table in snapshot.get("table_details", []):
        table_index = table.get("table_index")
        paragraph_indexes = [
            item.get("paragraph_index")
            for item in table.get("cell_paragraphs", [])
            if item.get("paragraph_index") is not None
        ]
        if table_index is None or not paragraph_indexes:
            continue
        first_paragraph = str(min(paragraph_indexes))
        page = paragraphs.get(first_paragraph)
        if page is not None:
            tables[str(table_index)] = page
    return {
        "page_count": raw.get("page_count"),
        "paragraphs": paragraphs,
        "tables": tables,
    }


def page_label(page) -> str:
    if page in (None, "", "null"):
        return "页码待确认"
    return f"第 {page} 页"


def issue_page_label(element_ref: dict, page_map: dict | None = None) -> str:
    page_map = page_map or {}
    table_index = element_ref.get("table_index")
    if table_index is not None:
        return page_label((page_map.get("tables") or {}).get(str(table_index)))
    paragraph_index = element_ref.get("paragraph_index")
    if paragraph_index is not None:
        return page_label((page_map.get("paragraphs") or {}).get(str(paragraph_index)))
    return "全文"


def changed_paragraph_styles(before: dict, after: dict) -> list[tuple[int, str, str, str]]:
    before_items = {p.get("element_id"): p for p in before.get("paragraphs", [])}
    after_items = {p.get("element_id"): p for p in after.get("paragraphs", [])}
    changes = []
    for element_id, before_item in before_items.items():
        after_item = after_items.get(element_id)
        if not after_item:
            continue
        before_style = before_item.get("style")
        after_style = after_item.get("style")
        if before_style == after_style:
            continue
        changes.append(
            (
                before_item.get("paragraph_index"),
                (before_item.get("text_preview") or "").replace("|", "\\|"),
                style_label(before_style),
                style_label(after_style),
            )
        )
    return changes


def issue_category(issue: dict) -> tuple[str, str]:
    action = issue.get("recommended_action", {})
    action_type = action.get("action_type")
    problem = issue.get("problem") or ""
    issue_type = issue.get("issue_type")
    risk_flags = set(issue.get("risk_flags") or [])
    element_type = issue.get("element_ref", {}).get("element_type")
    if action_type == "insert_or_replace_toc_field" or "自动目录" in problem or element_type == "toc-field":
        return "自动目录缺失", "需要确认目录内容与标题范围后再生成或替换自动目录。"
    if "merged_cells" in risk_flags or "合并单元格" in problem:
        return "复杂合并表格", "表格含合并单元格，需确认是否允许统一表格文字、内边距、垂直居中和表头策略。"
    if issue_type == "role_classification" or "Heading role" in problem or "heading role" in problem.lower():
        return "标题/编号角色不确定", "需确认该段落是标题、正文编号还是列表项，再决定是否纳入标题样式和目录。"
    if action_type in {"review_body_direct_format_strategy", "apply_body_direct_format"} or "direct formatting" in problem.lower() or "直接格式" in problem:
        return "正文直接格式覆盖", "段落存在直接格式覆盖，需确认是否保留局部强调或统一为正文规范。"
    if action_type == "review_body_format_mismatch" or "format mismatch" in problem.lower() or "基线" in problem:
        return "正文格式与基线不一致", "正文有效格式与代表性正文/Normal 基线不一致，需确认是否批量统一。"
    if action_type == "map_heading_native_style" or issue_type == "heading_style":
        return "标题样式未规范化", "标题未使用对应 Word 原生标题样式，可能影响目录生成和结构识别。"
    if action_type == "apply_body_style_definition" or issue_type == "body_style":
        return "正文样式定义不一致", "正文基础样式与已确认规则不一致，需要统一字体、字号、段距或缩进。"
    if action_type == "apply_table_cell_format" or element_type == "table-cell":
        return "表格文字格式不一致", "表格单元格文字格式与规则不一致，需要统一表格内文字和段落格式。"
    if "toc" in problem.lower() or "目录" in problem:
        return "目录内容需复核", "目录内容、层级或更新状态与文档标题结构不一致，需要复核后处理。"
    if issue.get("severity") in {"high", "blocker"}:
        return "高风险格式问题", "存在影响交付验收的格式问题，需要优先确认处理。"
    return "其他格式问题", "格式与已确认规则不一致，需要按规则处理或人工确认。"


def issue_problem_text(issue: dict) -> str:
    return issue_category(issue)[1]


def pages_text(pages: list[str], limit: int = 10) -> str:
    shown = "、".join(pages[:limit])
    if len(pages) > limit:
        shown += f" 等 {len(pages)} 处"
    return shown or "页码待确认"


def group_issues(issues: list[dict], page_map: dict | None = None) -> dict[str, dict]:
    grouped: dict[str, dict] = {}
    for issue in issues:
        name, decision = issue_category(issue)
        bucket = grouped.setdefault(name, {"decision": decision, "count": 0, "sample": issue, "pages": []})
        bucket["count"] += 1
        page = issue_page_label(issue.get("element_ref", {}), page_map)
        if page not in bucket["pages"]:
            bucket["pages"].append(page)
        if name == "复杂合并表格" and issue.get("element_ref", {}).get("table_index") == 21:
            bucket["sample"] = issue
    return grouped


def render_issue_case_table(issues: list[dict], limit: int = 30, page_map: dict | None = None) -> str:
    if not issues:
        return "无。\n"
    lines = [
        "| 序号 | 页码 | 原文摘录 | 问题分类 | 问题现象 | 文档现状 | 规则要求 | 风险 | 建议处理 |",
        "| --- | --- | --- | --- | --- | --- | --- | --- | --- |",
    ]
    for index, issue in enumerate(issues[:limit], start=1):
        element_ref = issue.get("element_ref", {})
        category, _decision = issue_category(issue)
        lines.append(
            "| {index} | {position} | {preview} | {category} | {problem} | {current} | {expected} | {severity} | {action} |".format(
                index=index,
                position=issue_page_label(element_ref, page_map),
                preview=preview_text(element_ref),
                category=category,
                problem=markdown_cell(issue_problem_text(issue)),
                current=issue_current(issue),
                expected=issue_expected(issue),
                severity=severity_label(issue.get("severity")),
                action=issue_action(issue),
            )
        )
    if len(issues) > limit:
        lines.append(f"\n仅展示前 {limit} 条代表性问题，其余同类问题已纳入机器追溯文件。")
    return "\n".join(lines) + "\n"


def render_issue_summary(issues: list[dict], page_map: dict | None = None, heading_level: int = 3) -> str:
    if not issues:
        return "无。\n"
    grouped = group_issues(issues, page_map)
    marker = "#" * heading_level
    lines = [
        f"{marker} 问题分类汇总",
        "",
        "| 分类 | 数量 | 涉及页码 | 处理说明 |",
        "| --- | ---: | --- | --- |",
    ]
    for name, bucket in grouped.items():
        lines.append(f"| {name} | {bucket['count']} | {pages_text(bucket['pages'])} | {bucket['decision']} |")
    lines.extend(["", f"{marker} 分类代表案例", ""])
    samples = [bucket["sample"] for bucket in grouped.values()]
    lines.append(render_issue_case_table(samples, limit=len(samples), page_map=page_map).rstrip())
    return "\n".join(lines) + "\n"


def render_change_table(changes: list[tuple[int, str, str, str]], limit: int = 30, page_map: dict | None = None) -> str:
    if not changes:
        return "无。\n"
    grouped: dict[str, dict] = {}
    for change in changes:
        paragraph_index, text, before_style, after_style = change
        name = f"{before_style} → {after_style}"
        phenomenon = f"段落样式由“{before_style}”调整为“{after_style}”。"
        bucket = grouped.setdefault(name, {"phenomenon": phenomenon, "count": 0, "sample": change, "pages": []})
        bucket["count"] += 1
        page = page_label((page_map or {}).get("paragraphs", {}).get(str(paragraph_index)))
        if page not in bucket["pages"]:
            bucket["pages"].append(page)
    lines = [
        "### 变更分类汇总",
        "",
        "| 分类 | 数量 | 涉及页码 | 变更现象 |",
        "| --- | ---: | --- | --- |",
    ]
    for name, bucket in grouped.items():
        lines.append(f"| {name} | {bucket['count']} | {pages_text(bucket['pages'])} | {bucket['phenomenon']} |")
    lines.extend(
        [
            "",
            "### 分类代表案例",
            "",
            "| 序号 | 页码 | 原文摘录 | 变更分类 | 变更现象 | 输入文档状态 | 输出文档状态 |",
            "| --- | --- | --- | --- | --- | --- | --- |",
        ]
    )
    samples = [bucket["sample"] for bucket in grouped.values()]
    for index, (paragraph_index, text, before_style, after_style) in enumerate(samples[:limit], start=1):
        page = page_label((page_map or {}).get("paragraphs", {}).get(str(paragraph_index)))
        category = f"{before_style} → {after_style}"
        phenomenon = f"段落样式由“{before_style}”调整为“{after_style}”。"
        lines.append(
            f"| {index} | {page} | {markdown_cell(text) or '无文本摘录'} | {category} | {phenomenon} | {before_style} | {after_style} |"
        )
    if len(samples) > limit:
        lines.append(f"\n仅展示前 {limit} 类代表性变更，其余同类变更已纳入机器追溯文件。")
    return "\n".join(lines) + "\n"


def render_manual_confirmation(issues: list[dict], page_map: dict | None = None) -> str:
    if not issues:
        return "无。\n"
    grouped = group_issues(issues, page_map)
    lines = [
        "## 1. 问题分类汇总",
        "",
        "| 分类 | 数量 | 涉及页码 | 需要确认的决策 |",
        "| --- | ---: | --- | --- |",
    ]
    for name, bucket in grouped.items():
        lines.append(f"| {name} | {bucket['count']} | {pages_text(bucket['pages'])} | {bucket['decision']} |")
    lines.extend(["", "## 2. 分类代表案例", ""])
    samples = [bucket["sample"] for bucket in grouped.values()]
    lines.append(render_issue_case_table(samples, limit=len(samples), page_map=page_map).rstrip())
    lines.extend(["", "完整机器追溯请查看 review_results/*.review.json 和 plans/repair_plan.yaml。"])
    return "\n".join(lines) + "\n"


def render_audit_group_summary(audits: list[dict]) -> str:
    if not audits:
        return "无。\n"
    lines = []
    for row in audits:
        summary = row["summary"]
        lines.append(
            "- {label}：检查 {checked} 项，发现 {issues} 项问题，可自动处理 {auto} 项，需要人工确认 {manual} 项。".format(
                label=row["label"],
                checked=summary.get("checked_items", 0),
                issues=summary.get("issues_found", 0),
                auto=summary.get("auto_fixable", 0),
                manual=summary.get("manual_review", 0),
            )
        )
    return "\n".join(lines) + "\n"


def render_review_group_summary(reviews: list[dict]) -> str:
    if not reviews:
        return "当前尚未生成复核结果。\n"
    lines = []
    for row in reviews:
        summary = row["summary"]
        checked = summary.get("checked_items", 0)
        issues_found = summary.get("issues_found", 0)
        if checked == 0:
            status = "本轮未发现可检查内容"
        elif issues_found == 0:
            status = "已通过复核"
        else:
            status = f"仍有 {issues_found} 项未解决"
        lines.append(f"- {row['label']}：复核 {checked} 项，剩余问题 {issues_found} 项，结论：{status}。")
    return "\n".join(lines) + "\n"


def render_rule_reference(run_dir: Path) -> str:
    rule_dir = run_dir / "rules" / "selected_rule"
    profile_path = rule_dir / "profile.yaml"
    rule_summary_path = rule_dir / "RULE_SUMMARY.md"
    return "\n".join(
        [
            f"- 规则目录：{rule_dir}",
            f"- 规则配置：{profile_path}",
            f"- 规则说明：{rule_summary_path}",
            "- 说明：完整规则说明不再嵌入最终验收报告，以避免报告正文出现非问题分类表格。",
        ]
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--run-dir", required=True)
    args = parser.parse_args()
    run_dir = Path(args.run_dir)
    reports = run_dir / "reports"
    audits = collect_audits(run_dir)
    reviews = collect_reviews(run_dir)
    repair_plan = run_dir / "plans" / "repair_plan.yaml"
    repair_log = run_dir / "logs" / "repair_log.yaml"
    rule_summary = run_dir / "rules" / "selected_rule" / "RULE_SUMMARY.md"
    run_log = load_run_log(run_dir)
    repair_log_data = load_repair_log(run_dir)
    before_snapshot = load_snapshot(run_dir / "snapshots" / "document_snapshot.before.json")
    after_snapshot = load_snapshot(run_dir / "snapshots" / "document_snapshot.after.json")
    page_map = load_page_map(run_dir, after_snapshot)
    changes = changed_paragraph_styles(before_snapshot, after_snapshot)

    total_issues = 0
    manual_count = 0
    auto_fixable = 0
    all_issues = []
    for row in audits:
        summary = row["summary"]
        total_issues += summary.get("issues_found", 0)
        manual_count += summary.get("manual_review", 0)
        auto_fixable += summary.get("auto_fixable", 0)
        all_issues.extend(row["issues"])

    source_docx = run_log.get("source_docx", "未记录")
    write(
        reports / "AUDIT_REPORT.md",
        render_template(
            "AUDIT_REPORT_TEMPLATE.md",
            {
                "SOURCE_DOCX": source_docx,
                "RULE_NAME": run_log.get("rule_id", "未记录"),
                "CHECKED_COUNT": sum(row["summary"].get("checked_items", 0) for row in audits),
                "TOTAL_ISSUES": total_issues,
                "AUTO_FIXABLE": auto_fixable,
                "MANUAL_COUNT": manual_count,
                "GROUP_TABLE": render_audit_group_summary(audits),
                "ISSUE_TABLE": render_issue_summary(all_issues, page_map=page_map),
            },
        ),
    )

    remaining_issues = []
    for row in reviews:
        summary = row["summary"]
        remaining_issues.extend(row["issues"])
    write(
        reports / "REVIEW_REPORT.md",
        render_template(
            "REVIEW_REPORT_TEMPLATE.md",
            {"REVIEW_TABLE": render_review_group_summary(reviews), "REMAINING_ISSUES": render_issue_summary(remaining_issues, page_map=page_map)},
        ),
    )

    manual_items = [
        issue
        for issue in remaining_issues
        if issue.get("recommended_action", {}).get("auto_fix_policy") != "auto-fix"
        or issue.get("severity") in {"high", "blocker"}
        or "自动目录" in (issue.get("problem") or "")
    ]
    write(
        reports / "MANUAL_CONFIRMATION.md",
        render_template(
            "MANUAL_CONFIRMATION_TEMPLATE.md",
            {"MANUAL_ITEMS": render_manual_confirmation(manual_items, page_map=page_map) if manual_items else "无。\n"},
        ),
    )

    toc_before = "已存在可自动更新目录" if before_snapshot.get("has_toc_field") else "未检测到可自动更新目录"
    toc_after = "已存在可自动更新目录" if after_snapshot.get("has_toc_field") else "仍未检测到可自动更新目录"
    write(
        reports / "DIFF_SUMMARY.md",
        render_template(
            "DIFF_SUMMARY_TEMPLATE.md",
            {
                "CHANGE_COUNT": len(changes),
                "TOC_BEFORE": toc_before,
                "TOC_AFTER": toc_after,
                "CHANGE_TABLE": render_change_table(changes, page_map=page_map),
            },
        ),
    )

    actions_total = repair_log_data.get("actions_total", "未记录")
    actions_applied = repair_log_data.get("actions_applied", "未记录")
    actions_skipped = repair_log_data.get("actions_skipped", "未记录")
    repair_note = repair_log_data.get("note", "未记录")
    write(
        reports / "REPAIR_LOG.md",
        render_template(
            "REPAIR_LOG_TEMPLATE.md",
            {
                "ACTIONS_TOTAL": actions_total,
                "ACTIONS_APPLIED": actions_applied,
                "ACTIONS_SKIPPED": actions_skipped,
                "REPAIR_NOTE": repair_note,
                "CHANGE_COUNT": len(changes),
                "UNRESOLVED_ITEMS": "- 自动目录仍未生成，需要后续专项处理或人工确认。\n" if not after_snapshot.get("has_toc_field") else "无。\n",
            },
        ),
    )

    repair_plan_status = "存在" if repair_plan.exists() else "缺失"
    remaining_review_issues = 0
    for row in reviews:
        remaining_review_issues += row["summary"].get("issues_found", 0)
    conclusion = "通过" if remaining_review_issues == 0 else "未通过"
    next_step = "可以交付输出文档。" if conclusion == "通过" else "先处理仍未解决的问题，尤其是自动目录生成，再重新复核。"
    write(
        reports / "FINAL_ACCEPTANCE_REPORT.md",
        render_template(
            "FINAL_ACCEPTANCE_REPORT_TEMPLATE.md",
            {
                "CONCLUSION": conclusion,
                "REMAINING_COUNT": remaining_review_issues,
                "RULE_SUMMARY": render_rule_reference(run_dir),
                "FAILED_REASONS": render_issue_summary(remaining_issues, page_map=page_map),
                "REPAIR_PLAN_STATUS": repair_plan_status,
                "AUDIT_GROUP_COUNT": len(audits),
                "REVIEW_GROUP_COUNT": len(reviews),
                "TOC_STATUS": "已通过" if remaining_review_issues == 0 else "仍需处理或人工确认",
                "OUTPUT_PATH": run_dir / "output",
                "NEXT_STEP": next_step,
            },
        ),
    )
    print(reports)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
