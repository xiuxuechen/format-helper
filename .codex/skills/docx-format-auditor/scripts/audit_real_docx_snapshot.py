#!/usr/bin/env python3
"""基于真实 DOCX 快照和已知问题清单生成语义审计。"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[4]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.validation.check_semantic_audit import validate_semantic_audit


TZ = timezone(timedelta(hours=8))


def load_json(path: Path) -> dict[str, Any]:
    """读取 JSON 文件。"""
    return json.loads(path.read_text(encoding="utf-8"))


def normalize_text(value: str) -> str:
    """去除标点和空白，便于跨文档匹配摘录。"""
    value = (value or "").replace("\\", "")
    return re.sub(r"[\s\u3000，。、“”‘’\"'：:；;（）()\[\]【】、,.…]+", "", value)


def parse_known_issues(path: Path) -> list[dict[str, str]]:
    """解析问题清单 Markdown 表格。"""
    text = path.read_text(encoding="utf-8-sig")
    issues: list[dict[str, str]] = []
    current_title = ""
    for raw_line in text.splitlines():
        line = raw_line.strip()
        title_match = re.match(r"^###\s*\d+\.\s*(.+)$", line)
        if title_match:
            current_title = title_match.group(1).strip()
            continue
        if not current_title or not line.startswith("|") or "---" in line:
            continue
        columns = [column.strip().strip('"') for column in line.strip("|").split("|")]
        if len(columns) < 4 or not columns[0].isdigit():
            continue
        issues.append(
            {
                "category": current_title,
                "sequence": columns[0],
                "location": columns[1],
                "excerpt": columns[2],
                "symptom": columns[3],
            }
        )
    return issues


def quoted_tokens(value: str) -> list[str]:
    """提取引号内关键词，失败时回退到长文本片段。"""
    value = (value or "").replace("\\", "")
    tokens = re.findall(r"[“\"']([^”\"']{2,})[”\"']", value)
    parts = [part.strip() for part in re.split(r"[，。；;、]+", value or "") if len(part.strip()) >= 4]
    result: list[str] = []
    for item in [*tokens, *parts]:
        if item not in result:
            result.append(item)
    return result[:6] if result else [value]


def paragraph_text(paragraph: dict[str, Any]) -> str:
    """读取段落预览文本。"""
    return str(paragraph.get("text_preview") or "")


def match_paragraphs(snapshot: dict[str, Any], issue: dict[str, str]) -> list[dict[str, Any]]:
    """根据问题摘录匹配段落。"""
    excerpt_patterns = [normalize_text(issue.get("excerpt", ""))]
    for token in quoted_tokens(issue.get("excerpt", "")):
        excerpt_patterns.append(normalize_text(token))
    excerpt_patterns = [item for item in excerpt_patterns if len(item) >= 4]
    location_patterns = [normalize_text(issue.get("location", ""))]
    for token in quoted_tokens(issue.get("location", "")):
        location_patterns.append(normalize_text(token))
    location_patterns = [item for item in location_patterns if len(item) >= 4]

    def find_by_patterns(patterns: list[str]) -> list[dict[str, Any]]:
        found: list[dict[str, Any]] = []
        for paragraph in snapshot.get("paragraphs", []):
            text = normalize_text(paragraph_text(paragraph))
            if not text:
                continue
            if any(pattern and pattern in text for pattern in patterns):
                found.append(paragraph)
        return found

    matches = find_by_patterns(excerpt_patterns)
    if matches:
        return matches
    haystacks = location_patterns
    matches: list[dict[str, Any]] = []
    for paragraph in snapshot.get("paragraphs", []):
        text = normalize_text(paragraph_text(paragraph))
        if not text:
            continue
        if any(pattern and pattern in text for pattern in haystacks):
            matches.append(paragraph)
    return matches


def iter_cells(snapshot: dict[str, Any]) -> list[dict[str, Any]]:
    """展开所有表格单元格。"""
    cells: list[dict[str, Any]] = []
    for table in snapshot.get("tables", []):
        for cell in table.get("cells", []):
            item = dict(cell)
            item["table_id"] = table.get("element_id")
            cells.append(item)
    return cells


def match_cells(snapshot: dict[str, Any], issue: dict[str, str]) -> list[dict[str, Any]]:
    """根据问题摘录匹配表格单元格。"""
    tokens = [normalize_text(token) for token in quoted_tokens(issue.get("excerpt", ""))]
    tokens.extend(normalize_text(part) for part in re.split(r"[，、；;]+", issue.get("excerpt", "")) if len(part.strip()) >= 2)
    tokens = [token for token in tokens if token]
    matches: list[dict[str, Any]] = []
    for cell in iter_cells(snapshot):
        text = normalize_text(str(cell.get("text_preview") or ""))
        if text and any(token in text or text in token for token in tokens):
            matches.append(cell)
    return matches


def format_signature(paragraph: dict[str, Any]) -> tuple[Any, ...]:
    """抽取用于一致性对比的有效格式签名。"""
    paragraph_format_data = paragraph.get("resolved_paragraph_format") or paragraph.get("paragraph_format") or {}
    run_format_data = paragraph.get("resolved_run_format") or paragraph.get("run_format") or {}
    return (
        paragraph_format_data.get("first_line_indent_cm"),
        paragraph_format_data.get("left_indent_cm"),
        paragraph_format_data.get("line_spacing_raw"),
        paragraph_format_data.get("line_spacing_rule"),
        run_format_data.get("font_east_asia"),
        run_format_data.get("font_size_pt"),
        run_format_data.get("bold"),
    )


def paragraph_element_id(paragraph: dict[str, Any], fallback: str) -> str:
    """读取段落 element_id。"""
    return str(paragraph.get("element_id") or fallback)


def make_audit_item(
    issue_id: str,
    element_id: str,
    role: str,
    problem: str,
    expected_role: str,
    confidence: float,
    evidence: list[str],
    action_type: str,
    policy: str,
    risk_level: str,
    before: dict[str, Any] | None = None,
    after: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """构造 semantic_audit item。"""
    return {
        "issue_id": issue_id,
        "element_id": element_id,
        "semantic_role": role,
        "current_problem": problem,
        "expected_role": expected_role,
        "confidence": confidence,
        "evidence": evidence,
        "recommended_action": {
            "action_type": action_type,
            "auto_fix_policy": policy,
            "before": before or {},
            "after": after or {},
        },
        "risk_level": risk_level,
    }


def adjacent_blank_detected(snapshot: dict[str, Any], paragraphs: list[dict[str, Any]]) -> bool:
    """判断匹配段落附近是否存在空段。"""
    all_paragraphs = snapshot.get("paragraphs", [])
    by_index = {item.get("paragraph_index"): item for item in all_paragraphs}
    for paragraph in paragraphs:
        index = paragraph.get("paragraph_index")
        if not isinstance(index, int):
            continue
        for nearby in range(index - 2, index + 3):
            candidate = by_index.get(nearby)
            if candidate is not None and not paragraph_text(candidate).strip():
                return True
    return False


def build_audit(
    standard_snapshot: dict[str, Any],
    issue_snapshot: dict[str, Any],
    known_issues: list[dict[str, str]],
    source_snapshot: str,
    rule_profile_id: str,
    now: datetime,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """生成 semantic_audit 和已知问题覆盖报告。"""
    del standard_snapshot
    audit_items: list[dict[str, Any]] = []
    coverage_items: list[dict[str, Any]] = []

    for index, known_issue in enumerate(known_issues, start=1):
        category = known_issue["category"]
        paragraphs = match_paragraphs(issue_snapshot, known_issue)
        cells = match_cells(issue_snapshot, known_issue) if "表格" in category else []
        issue_code = f"K{index:03d}"
        status = "covered"
        evidence: list[str] = [
            f"问题类型：{category}",
            f"位置：{known_issue['location']}",
            f"摘录：{known_issue['excerpt']}",
            f"问题现象：{known_issue['symptom']}",
        ]
        action_type = "apply_body_direct_format"
        role = "body-paragraph"
        expected_role = "body-paragraph"
        element_id = "document"
        before: dict[str, Any] = {}
        after: dict[str, Any] = {}

        if "表格" in category:
            action_type = "apply_table_cell_format"
            role = "table-cell"
            expected_role = "table-cell"
            if cells:
                element_id = str(cells[0].get("cell_id") or cells[0].get("table_id") or "table")
                bold_cells = [cell for cell in cells if (cell.get("format_summary") or {}).get("has_bold")]
                evidence.append(f"匹配单元格数量：{len(cells)}")
                evidence.append(f"检测到加粗单元格数量：{len(bold_cells)}")
                before = {"matched_cell_count": len(cells), "bold_cell_count": len(bold_cells)}
                after = {"expected_bold": False}
            else:
                status = "unmatched"
                evidence.append("未匹配到表格单元格")
        elif "空行" in category:
            role = "blank-line"
            expected_role = "compact-section-spacing"
            action_type = "apply_body_direct_format"
            if paragraphs:
                element_id = paragraph_element_id(paragraphs[0], "paragraph")
                evidence.append(f"匹配段落数量：{len(paragraphs)}")
                evidence.append(f"附近空段落：{adjacent_blank_detected(issue_snapshot, paragraphs)}")
            else:
                status = "unmatched"
                evidence.append("未匹配到相关段落")
        else:
            if paragraphs:
                element_id = paragraph_element_id(paragraphs[0], "paragraph")
                signatures = {format_signature(paragraph) for paragraph in paragraphs}
                evidence.append(f"匹配段落数量：{len(paragraphs)}")
                evidence.append(f"有效格式签名数量：{len(signatures)}")
                before = {"format_signatures": [list(item) for item in sorted(signatures, key=str)]}
            else:
                status = "unmatched"
                evidence.append("未匹配到相关段落")

        if status == "unmatched":
            confidence = 0.55
            risk_level = "high"
            policy = "audit-only"
        else:
            confidence = 0.88
            risk_level = "medium"
            policy = "manual-review"

        audit_items.append(
            make_audit_item(
                issue_id=issue_code,
                element_id=element_id,
                role=role,
                problem=f"{category}：{known_issue['symptom']}",
                expected_role=expected_role,
                confidence=confidence,
                evidence=evidence,
                action_type=action_type,
                policy=policy,
                risk_level=risk_level,
                before=before,
                after=after,
            )
        )
        coverage_items.append(
            {
                "issue_id": issue_code,
                "category": category,
                "status": status,
                "matched_paragraph_count": len(paragraphs),
                "matched_cell_count": len(cells),
                "audit_item_id": issue_code,
                "location": known_issue["location"],
                "excerpt": known_issue["excerpt"],
            }
        )

    audit = {
        "schema_version": "1.0.0",
        "source_snapshot": source_snapshot,
        "rule_profile_id": rule_profile_id,
        "generated_by": "codex",
        "generated_at": now.isoformat(),
        "items": audit_items,
    }
    coverage = {
        "schema_version": "1.0.0",
        "generated_at": now.isoformat(),
        "source_snapshot": source_snapshot,
        "known_issue_count": len(known_issues),
        "covered_count": sum(1 for item in coverage_items if item["status"] == "covered"),
        "unmatched_count": sum(1 for item in coverage_items if item["status"] == "unmatched"),
        "items": coverage_items,
    }
    return audit, coverage


def main() -> int:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="生成真实 DOCX 语义审计和问题清单覆盖报告")
    parser.add_argument("--standard-snapshot", required=True, type=Path)
    parser.add_argument("--issue-snapshot", required=True, type=Path)
    parser.add_argument("--known-issues", required=True, type=Path)
    parser.add_argument("--rule-profile-id", default="real-docx-standard")
    parser.add_argument("--semantic-audit-output", required=True, type=Path)
    parser.add_argument("--coverage-output", required=True, type=Path)
    args = parser.parse_args()

    now = datetime.now(TZ)
    known_issues = parse_known_issues(args.known_issues)
    audit, coverage = build_audit(
        standard_snapshot=load_json(args.standard_snapshot),
        issue_snapshot=load_json(args.issue_snapshot),
        known_issues=known_issues,
        source_snapshot=str(args.issue_snapshot).replace("\\", "/"),
        rule_profile_id=args.rule_profile_id,
        now=now,
    )
    errors = validate_semantic_audit(audit)
    if errors:
        for error in errors:
            print(error)
        return 1

    args.semantic_audit_output.parent.mkdir(parents=True, exist_ok=True)
    args.coverage_output.parent.mkdir(parents=True, exist_ok=True)
    args.semantic_audit_output.write_text(json.dumps(audit, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.coverage_output.write_text(json.dumps(coverage, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(args.semantic_audit_output)
    print(args.coverage_output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
