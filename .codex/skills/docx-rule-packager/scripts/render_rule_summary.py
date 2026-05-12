"""从语义规则草案或槽位事实生成 RULE_SUMMARY.md。"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
from typing import Any


RESOLVED_STATUSES = {"resolved", "resolved_with_conflicts", "not_applicable", "user_confirmed"}


def format_bool(value: Any) -> str:
    """将布尔值渲染为中文。"""
    if value is True:
        return "是"
    if value is False:
        return "否"
    return "未指定"


def alignment_label(value: Any) -> str:
    """渲染对齐方式中文标签。"""
    labels = {
        "left": "左对齐",
        "center": "居中",
        "right": "右对齐",
        "justify": "两端对齐",
        "distributed": "分散对齐",
        "both": "两端对齐",
        "distribute": "分散对齐",
    }
    return labels.get(value, format_value(value))


def strategy_label(value: str) -> str:
    """渲染写入策略中文标签。"""
    labels = {
        "style-definition": "样式定义",
        "direct-format": "直接格式",
        "audit-only": "仅审计",
    }
    return labels.get(value, value)


def format_value(value: Any, suffix: str = "") -> str:
    """渲染格式值。"""
    if value is None or value == "":
        return "未指定"
    return f"{value}{suffix}"


def render_format(format_rule: dict[str, Any]) -> str:
    """渲染用户可读格式说明。"""
    parts = []
    if "font_east_asia" in format_rule:
        parts.append(f"字体：{format_value(format_rule.get('font_east_asia'))}")
    if "font_size_pt" in format_rule:
        parts.append(f"字号：{format_value(format_rule.get('font_size_pt'), 'pt')}")
    if "bold" in format_rule:
        parts.append(f"加粗：{format_bool(format_rule.get('bold'))}")
    if "first_line_indent_cm" in format_rule:
        parts.append(f"首行缩进：{format_value(format_rule.get('first_line_indent_cm'), 'cm')}")
    if "line_spacing_multiple" in format_rule:
        parts.append(f"行距：{format_value(format_rule.get('line_spacing_multiple'), ' 倍')}")
    if "outline_level" in format_rule:
        parts.append(f"大纲级别：{format_value(format_rule.get('outline_level'))}")
    if "alignment" in format_rule:
        parts.append(f"对齐：{alignment_label(format_rule.get('alignment'))}")
    return "；".join(parts) if parts else "未指定"


def unwrap_slot_value(value: Any) -> Any:
    """兼容带来源对象和值本身两种槽位值形态。"""
    if isinstance(value, dict) and set(value.keys()) & {"value", "source", "confidence"}:
        return value.get("value")
    return value


def format_slot_value(slot_name: str, value: Any, unit: Any = None) -> str:
    """按槽位类型渲染用户可读值。"""
    raw_value = unwrap_slot_value(value)
    if raw_value is None or raw_value == "":
        return "未指定"
    if slot_name in {"alignment"}:
        return alignment_label(raw_value)
    if slot_name == "vertical_alignment":
        labels = {"top": "顶端", "center": "中部", "bottom": "底端"}
        return labels.get(raw_value, format_value(raw_value))
    if unit:
        unit_labels = {"pt": "pt", "cm": "cm", "multiple": "倍", "twip": "twip", "half_pt": "半磅"}
        suffix = unit_labels.get(unit, str(unit))
        return f"{raw_value}{suffix}"
    return str(raw_value)


def load_contract(path: Path | None) -> dict[str, Any] | None:
    """读取 JSON/YAML 契约文件。"""
    if path is None:
        return None
    text = path.read_text(encoding="utf-8")
    if path.suffix.lower() == ".json":
        return json.loads(text)
    try:
        import yaml  # type: ignore
    except ImportError as exc:  # pragma: no cover - 只在缺依赖环境触发
        raise RuntimeError("读取 role_slot_contract.yaml 需要 PyYAML") from exc
    return yaml.safe_load(text)


def role_required_slots(role: dict[str, Any], contract: dict[str, Any] | None) -> list[str]:
    """从契约获取必需槽位；缺契约时回退到当前 role 的 slot_summary。"""
    role_kind = role.get("role_kind", "")
    contracts = (contract or {}).get("role_slot_contracts", {})
    if role_kind in contracts:
        return list(contracts[role_kind].get("required_slots", []))
    return list(role.get("slot_summary", {}).keys())


def role_optional_slots(role: dict[str, Any], contract: dict[str, Any] | None) -> list[str]:
    """从契约获取可选槽位。"""
    role_kind = role.get("role_kind", "")
    contracts = (contract or {}).get("role_slot_contracts", {})
    if role_kind in contracts:
        return list(contracts[role_kind].get("optional_slots", []))
    return []


def slot_confidence(summary: dict[str, Any]) -> float:
    """推导槽位置信度。"""
    if isinstance(summary.get("confidence"), (int, float)):
        return float(summary["confidence"])
    coverage = summary.get("mode_coverage")
    if isinstance(coverage, (int, float)):
        return float(coverage)
    return 0.0


def format_required_slot_cells(role: dict[str, Any], required_slots: list[str]) -> str:
    """渲染一个角色的 required_slots 摘要。"""
    slot_summary = role.get("slot_summary", {})
    cells = []
    for slot_name in required_slots:
        summary = slot_summary.get(slot_name, {})
        cells.append(
            "{slot}={value}（置信度 {confidence:.2f}）".format(
                slot=slot_name,
                value=format_slot_value(slot_name, summary.get("mode_value"), summary.get("unit")),
                confidence=slot_confidence(summary),
            )
        )
    return "<br>".join(cells) if cells else "无 required_slots。"


def format_histogram(histogram: Any, slot_name: str, unit: Any = None) -> str:
    """渲染冲突分布。"""
    if not histogram:
        return "无 histogram。"
    parts = []
    for item in histogram:
        if isinstance(item, dict):
            value = item.get("value")
            count = item.get("count", item.get("sample_count", item.get("frequency", 0)))
            ratio = item.get("ratio", item.get("coverage"))
            suffix = f"，占比 {float(ratio):.2f}" if isinstance(ratio, (int, float)) else ""
            parts.append(f"{format_slot_value(slot_name, value, unit)}：{count}{suffix}")
        else:
            parts.append(str(item))
    return "；".join(parts)


def extract_slot_rows(
    slot_facts: dict[str, Any],
    contract: dict[str, Any] | None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    """抽取已确定、未确定和冲突三类展示行。"""
    resolved_rows: list[dict[str, Any]] = []
    unresolved_rows: list[dict[str, Any]] = []
    conflict_rows: list[dict[str, Any]] = []

    for role in slot_facts.get("roles", []):
        required_slots = role_required_slots(role, contract)
        optional_slots = role_optional_slots(role, contract)
        slot_summary = role.get("slot_summary", {})
        required_statuses = [slot_summary.get(slot, {}).get("status", "unresolved") for slot in required_slots]
        if required_slots and all(status in RESOLVED_STATUSES for status in required_statuses):
            resolved_rows.append({"role": role, "required_slots": required_slots})
        for slot_name, summary in slot_summary.items():
            status = summary.get("status")
            is_required = slot_name in required_slots
            is_optional = slot_name in optional_slots or not is_required
            if status == "unresolved":
                unresolved_rows.append({"role": role, "slot_name": slot_name, "summary": summary})
            if status == "conflict" and is_required:
                conflict_rows.append({"role": role, "slot_name": slot_name, "summary": summary, "severity": "blocking"})
            elif status == "resolved_with_conflicts" and is_required:
                conflict_rows.append({"role": role, "slot_name": slot_name, "summary": summary, "severity": "warning"})
            elif status == "conflict" and is_optional:
                conflict_rows.append({"role": role, "slot_name": slot_name, "summary": summary, "severity": "warning"})
    return resolved_rows, unresolved_rows, conflict_rows


def render_slot_facts_summary(
    slot_facts: dict[str, Any] | None,
    contract: dict[str, Any] | None = None,
) -> tuple[str, dict[str, Any]]:
    """从 role_format_slot_facts.json 渲染 RULE_SUMMARY.md 和 metrics。

    若 slot_facts 为 None 或格式错误，返回错误占位文本并记录 metrics 为全 0。
    """
    if not isinstance(slot_facts, dict):
        return (
            "# 规则摘要\n\n## 错误\n\n规则提取产物 `role_format_slot_facts.json` 缺失或格式错误，无法渲染规则摘要。",
            {"sections_rendered": 0, "error": "slot_facts_missing"},
        )
    resolved_rows, unresolved_rows, conflict_rows = extract_slot_rows(slot_facts, contract)
    gate_blockers = slot_facts.get("gate_blockers", [])
    roles = slot_facts.get("roles", [])

    blocking_conflict_count = sum(1 for row in conflict_rows if row["severity"] == "blocking")
    warning_conflict_count = sum(1 for row in conflict_rows if row["severity"] == "warning")
    metrics = {
        "sections_rendered": 6,
        "resolved_rule_row_count": len(resolved_rows),
        "unresolved_slot_count": len(unresolved_rows),
        "blocking_conflict_count": blocking_conflict_count,
        "warning_conflict_count": warning_conflict_count,
        "gate_blocker_count": len(gate_blockers),
        "resolved_slot_count": sum(
            1
            for role in roles
            for summary in role.get("slot_summary", {}).values()
            if summary.get("status") in {"resolved", "resolved_with_conflicts", "not_applicable"}
        ),
        "conflict_slot_count": blocking_conflict_count + warning_conflict_count,
        "user_confirmed_slot_count": sum(
            1
            for role in roles
            for summary in role.get("slot_summary", {}).values()
            if summary.get("status") == "user_confirmed"
        ),
    }

    contract_ref = slot_facts.get("contract_ref", {})
    lines = [
        f"# {slot_facts.get('facts_id', 'role-format-slot-facts')} 规则摘要",
        "",
        "## 1. 规则来源与适用范围",
        "",
        f"- 运行 ID：`{slot_facts.get('run_id', '未指定')}`",
        f"- 来源快照：`{slot_facts.get('source_snapshot_path', '未指定')}`",
        f"- 槽位契约：`{contract_ref.get('contract_path', '未指定')}`",
        f"- Gate 状态：{slot_facts.get('gate_status', '未指定')}",
        "- 说明：本摘要仅供人工阅读，机器权威为 `role_format_slot_facts.json` 与规则包结构化产物。",
        "",
        "## 2. 已确定的格式规则",
        "",
        "| 角色 | required_slots 取值 |",
        "| --- | --- |",
    ]
    if resolved_rows:
        for row in resolved_rows:
            role = row["role"]
            lines.append(
                f"| {role.get('role_kind', 'unknown')} | {format_required_slot_cells(role, row['required_slots'])} |"
            )
    else:
        lines.append("| 无 | 无 |")

    lines.extend(["", "## 3. ⚠️ 未能确定的格式属性", ""])
    if unresolved_rows:
        for row in unresolved_rows:
            role = row["role"]
            summary = row["summary"]
            prompt = summary.get("confirmation_prompt") or "需要人工确认。"
            lines.append(f"- {role.get('role_kind')} × {row['slot_name']}：{prompt}")
    else:
        lines.append("无。")

    lines.extend(["", "## 4. ⚠️ 存在冲突的格式属性", ""])
    if conflict_rows:
        for row in conflict_rows:
            role = row["role"]
            summary = row["summary"]
            severity = "阻断" if row["severity"] == "blocking" else "警告"
            histogram = format_histogram(summary.get("value_histogram"), row["slot_name"], summary.get("unit"))
            lines.append(f"- {role.get('role_kind')} × {row['slot_name']}（{severity}）：{histogram}")
    else:
        lines.append("无。")

    lines.extend(["", "## 5. 人工确认清单", ""])
    if gate_blockers:
        for blocker in gate_blockers:
            options = "、".join(str(option) for option in blocker.get("suggested_options", [])) or "无建议选项"
            lines.append(
                f"- {blocker.get('blocker_id')}：{blocker.get('role_kind')} × {blocker.get('slot_name')}；"
                f"{blocker.get('message')}；建议选项：{options}"
            )
    else:
        lines.append("无。")

    lines.extend(["", "## 6. 规则证据", ""])
    evidence_written = False
    for role in roles:
        samples = role.get("samples", [])
        if not samples:
            continue
        evidence_written = True
        lines.append(f"### {role.get('role_kind', 'unknown')}")
        for sample in samples:
            lines.append(f"- `{sample.get('fact_id', 'unknown')}`：{sample.get('text_preview', '')}")
        lines.append("")
    if not evidence_written:
        lines.append("无。")
        lines.append("")

    return "\n".join(lines), metrics


def render_rule_summary(
    slot_facts: dict[str, Any],
    rule_package: dict[str, Any] | None,
    contract: dict[str, Any] | None,
    output_path: Path,
) -> dict[str, Any]:
    """从槽位事实渲染 RULE_SUMMARY.md，并返回 skill-result.metrics 所需字段。"""
    del rule_package
    content, metrics = render_slot_facts_summary(slot_facts, contract)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    temp_path.write_text(content, encoding="utf-8")
    os.replace(temp_path, output_path)
    data = output_path.read_bytes()
    metrics.update(
        {
            "output_path": str(output_path),
            "sha256": hashlib.sha256(data).hexdigest(),
            "size_bytes": len(data),
        }
    )
    return metrics


def render_summary(draft: dict[str, Any]) -> str:
    """渲染 RULE_SUMMARY.md 正文。"""
    lines = [
        f"# {draft['rule_id']} 规则摘要",
        "",
        "## 1. 规则来源与适用范围",
        "",
        f"- 文档类型：{draft['document_type']}",
        f"- 来源快照：`{draft['source_snapshot']}`",
        f"- 规则状态：draft",
        "",
        "## 2. 核心语义规则",
        "",
        "| 角色 | 说明 | 可见格式规则 | 写入策略 | 置信度 | 人工确认 |",
        "| --- | --- | --- | --- | --- | --- |",
    ]
    for role in draft["roles"]:
        lines.append(
            "| {role} | {description} | {format_rule} | {strategy} | {confidence:.2f} | {confirm} |".format(
                role=role["role"],
                description=role["description"],
                format_rule=render_format(role["format"]),
                strategy=strategy_label(role["write_strategy"]),
                confidence=float(role["confidence"]),
                confirm="是" if role["requires_user_confirmation"] else "否",
            )
        )

    lines.extend(["", "## 3. 规则证据", ""])
    for role in draft["roles"]:
        lines.append(f"### {role['description']}")
        for evidence in role["evidence"]:
            lines.append(f"- {evidence}")
        lines.append("")

    lines.extend(["## 4. 需人工确认", ""])
    if draft.get("manual_confirmation"):
        for item in draft["manual_confirmation"]:
            lines.append(f"- {item['item']}：{item['reason']}")
    else:
        lines.append("无。")

    lines.extend(
        [
            "",
            "## 5. 自动处理边界",
            "",
            "- 低于 0.85 置信度的规则不得自动升级为强规则。",
            "- 高风险、复杂结构和仅审计项必须进入人工确认。",
            "- 后续修复动作必须继续通过 schema、风险策略和白名单校验。",
            "",
        ]
    )
    return "\n".join(lines)


def main() -> int:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="从槽位事实或 semantic_rule_draft.json 生成 RULE_SUMMARY.md")
    parser.add_argument("--draft", type=Path)
    parser.add_argument("--slot-facts", type=Path)
    parser.add_argument("--contract", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()

    if args.slot_facts:
        if args.contract is None:
            parser.error("使用 --slot-facts 时必须同时提供 --contract")
        slot_facts = json.loads(args.slot_facts.read_text(encoding="utf-8"))
        metrics = render_rule_summary(slot_facts, {}, load_contract(args.contract), args.output)
        print(json.dumps(metrics, ensure_ascii=False, indent=2))
        return 0
    if args.draft:
        draft = json.loads(args.draft.read_text(encoding="utf-8"))
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(render_summary(draft), encoding="utf-8")
        print(args.output)
        return 0
    parser.error("必须提供 --slot-facts 或 --draft")
    return 0


def main_from_test(draft: Path, output: Path) -> int:
    """测试入口：按 CLI 等价逻辑生成摘要。"""
    data = json.loads(draft.read_text(encoding="utf-8"))
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(render_summary(data), encoding="utf-8")
    return 0


def main_from_slot_facts_test(slot_facts: Path, output: Path, contract: Path | None = None) -> dict[str, Any]:
    """测试入口：按槽位事实生成摘要并返回 metrics。"""
    data = json.loads(slot_facts.read_text(encoding="utf-8"))
    return render_rule_summary(data, {}, load_contract(contract), output)


if __name__ == "__main__":
    raise SystemExit(main())
