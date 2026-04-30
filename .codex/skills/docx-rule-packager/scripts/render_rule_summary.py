"""从 semantic_rule_draft.json 生成 RULE_SUMMARY.md。"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


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
    parser = argparse.ArgumentParser(description="从 semantic_rule_draft.json 生成 RULE_SUMMARY.md")
    parser.add_argument("--draft", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()

    draft = json.loads(args.draft.read_text(encoding="utf-8"))
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(render_summary(draft), encoding="utf-8")
    print(args.output)
    return 0


def main_from_test(draft: Path, output: Path) -> int:
    """测试入口：按 CLI 等价逻辑生成摘要。"""
    data = json.loads(draft.read_text(encoding="utf-8"))
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(render_summary(data), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
