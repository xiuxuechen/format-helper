#!/usr/bin/env python3
"""执行 V2 repair_plan.yaml 中允许的安全修复动作。"""
from __future__ import annotations

import argparse
import shutil
import zipfile
from datetime import datetime, timezone, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET


TZ = timezone(timedelta(hours=8))
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = NS["w"]
ET.register_namespace("w", W)


def parse_scalar_yaml(path: Path) -> dict:
    data = {}
    for raw in path.read_text(encoding="utf-8").splitlines():
        if ":" not in raw or raw.startswith(" "):
            continue
        key, value = raw.split(":", 1)
        data[key.strip()] = value.strip().strip('"').strip("'")
    return data


def parse_rule_profile_id(path: Path) -> str | None:
    in_rule_profile = False
    for raw in path.read_text(encoding="utf-8").splitlines():
        stripped = raw.strip()
        if stripped == "rule_profile:":
            in_rule_profile = True
            continue
        if in_rule_profile and raw and not raw.startswith(" "):
            in_rule_profile = False
        if in_rule_profile and stripped.startswith("id:"):
            return stripped.split(":", 1)[1].strip().strip('"').strip("'")
    return None


def rule_file(repair_plan: Path, name: str) -> Path:
    run_dir = repair_plan.parent.parent
    selected = run_dir / "rules" / "selected_rule" / name
    if selected.exists():
        return selected
    rule_id = parse_rule_profile_id(repair_plan)
    if not rule_id:
        return selected
    return Path("format_rules") / "docx" / "rule_profiles" / rule_id / name


def scalar(value: str | None):
    if value in (None, "", "null", "None"):
        return None
    if value in {"true", "false"}:
        return value == "true"
    try:
        if "." in value:
            return float(value)
        return int(value)
    except (TypeError, ValueError):
        return value


def parse_style_map(path: Path) -> dict[str, dict]:
    if not path.exists():
        return {}
    styles: dict[str, dict] = {}
    current: dict | None = None
    current_role: str | None = None
    in_format = False
    for raw in path.read_text(encoding="utf-8").splitlines():
        if raw.startswith("  ") and not raw.startswith("    ") and raw.strip().endswith(":"):
            current_role = raw.strip()[:-1]
            current = {"format": {}}
            styles[current_role] = current
            in_format = False
            continue
        if current is None:
            continue
        stripped = raw.strip()
        if stripped == "format:":
            in_format = True
            continue
        if raw.startswith("    ") and not raw.startswith("      "):
            in_format = False
        if ":" not in stripped:
            continue
        key, value = stripped.split(":", 1)
        value = value.strip().strip('"').strip("'")
        if in_format:
            current["format"][key] = scalar(value)
        elif key in {"word_style_id", "word_style_name", "outline_level", "write_strategy"}:
            current[key] = scalar(value)
    return styles


def parse_table_rules(path: Path) -> dict:
    data = {
        "header": {},
        "body": {},
        "layout": {},
        "row_height": {"mode": "audit-only"},
        "border": {"mode": "audit-only"},
    }
    if not path.exists():
        return data
    section = None
    for raw in path.read_text(encoding="utf-8").splitlines():
        stripped = raw.strip()
        if raw.startswith("  ") and not raw.startswith("    ") and stripped.endswith(":"):
            name = stripped[:-1]
            if name in data:
                section = name
            continue
        if section and raw.startswith("    ") and ":" in stripped:
            key, value = stripped.split(":", 1)
            data[section][key] = scalar(value.strip().strip('"').strip("'"))
    return data


def parse_auto_fix_actions(path: Path) -> list[dict]:
    actions: list[dict] = []
    current: dict | None = None
    in_target = False
    for raw in path.read_text(encoding="utf-8").splitlines():
        stripped = raw.strip()
        if raw.startswith("  - action_id:"):
            if current:
                actions.append(current)
            current = {"action_id": stripped.split(":", 1)[1].strip()}
            in_target = False
            continue
        if current is None:
            continue
        if stripped == "target:":
            in_target = True
            continue
        if raw.startswith("    ") and not raw.startswith("      "):
            in_target = False
        if in_target and stripped.startswith("element_id:"):
            current["element_id"] = stripped.split(":", 1)[1].strip()
        elif in_target and stripped.startswith("expected_role:"):
            current["expected_role"] = stripped.split(":", 1)[1].strip()
        elif stripped.startswith("action_type:"):
            current["action_type"] = stripped.split(":", 1)[1].strip()
        elif stripped.startswith("format_write_strategy:"):
            current["format_write_strategy"] = stripped.split(":", 1)[1].strip()
        elif stripped.startswith("auto_fix_policy:"):
            current["auto_fix_policy"] = stripped.split(":", 1)[1].strip()
    if current:
        actions.append(current)
    return [item for item in actions if item.get("auto_fix_policy") == "auto-fix"]


def set_attr(node, name: str, value) -> None:
    node.set(f"{{{W}}}{name}", str(value))


def w_attr(node, name: str):
    return node.get(f"{{{W}}}{name}")


def get_or_create(parent, name: str, prepend: bool = False):
    node = parent.find(f"./w:{name}", NS)
    if node is None:
        node = ET.SubElement(parent, f"{{{W}}}{name}")
        if prepend and len(parent) > 1:
            parent.remove(node)
            parent.insert(0, node)
    return node


def half_points(value) -> str | None:
    if value in (None, "", "null"):
        return None
    return str(int(round(float(value) * 2)))


def set_run_bool(parent, name: str, enabled: bool) -> None:
    node = parent.find(f"./w:{name}", NS)
    if enabled and node is None:
        ET.SubElement(parent, f"{{{W}}}{name}")
    elif not enabled and node is not None:
        parent.remove(node)


def paragraph_style_id(paragraph) -> str | None:
    pstyle = paragraph.find("./w:pPr/w:pStyle", NS)
    if pstyle is None:
        return None
    return w_attr(pstyle, "val")


def paragraph_uses_style(paragraph, style_id: str, include_implicit_normal: bool = False) -> bool:
    current_style = paragraph_style_id(paragraph)
    if current_style == style_id:
        return True
    return include_implicit_normal and current_style is None


def is_normal_style(style_id: str) -> bool:
    return style_id.lower() == "normal" or style_id in {"正文", "正文基础"}


def apply_style_paragraph_format(style, rule: dict, outline_level=None) -> None:
    ppr = get_or_create(style, "pPr")
    line_spacing_multiple = rule.get("line_spacing_multiple")
    line_spacing_pt = rule.get("line_spacing_pt")
    if line_spacing_multiple is not None or line_spacing_pt is not None:
        spacing = get_or_create(ppr, "spacing")
        if line_spacing_multiple is not None:
            set_attr(spacing, "line", int(round(float(line_spacing_multiple) * 240)))
            set_attr(spacing, "lineRule", "auto")
        elif line_spacing_pt is not None:
            set_attr(spacing, "line", int(round(float(line_spacing_pt) * 20)))
            set_attr(spacing, "lineRule", "exact")
    if outline_level is not None:
        set_attr(get_or_create(ppr, "outlineLvl"), "val", int(outline_level) - 1)


def clear_direct_format_conflicts(paragraph, rule: dict) -> dict:
    detail = {"count": 0, "removed_tags": {}}

    def record_removed(tag: str) -> None:
        detail["count"] += 1
        detail["removed_tags"][tag] = detail["removed_tags"].get(tag, 0) + 1

    ppr = paragraph.find("./w:pPr", NS)
    if ppr is not None and (rule.get("line_spacing_multiple") is not None or rule.get("line_spacing_pt") is not None):
        spacing = ppr.find("./w:spacing", NS)
        if spacing is not None:
            ppr.remove(spacing)
            record_removed("spacing")
    run_tags = []
    if rule.get("font_east_asia") or rule.get("font_ascii"):
        run_tags.append("rFonts")
    if half_points(rule.get("font_size_pt")):
        run_tags.extend(["sz", "szCs"])
    if rule.get("bold") is not None:
        run_tags.append("b")
    if not run_tags:
        return detail
    for run in paragraph.findall("./w:r", NS):
        rpr = run.find("./w:rPr", NS)
        if rpr is None:
            continue
        for tag in run_tags:
            node = rpr.find(f"./w:{tag}", NS)
            if node is not None:
                rpr.remove(node)
                record_removed(tag)
        if len(rpr) == 0:
            run.remove(rpr)
    return detail


def format_conflict_details(details: list[dict], max_items: int = 20) -> list[str]:
    if not details:
        return ["cleaned_conflict_details: none"]
    lines = ["cleaned_conflict_details:"]
    for item in details[:max_items]:
        tag_summary = ", ".join(
            f"{tag} x{count}" if count != 1 else tag
            for tag, count in sorted(item["removed_tags"].items())
        )
        lines.extend(
            [
                f"  - action_id: {item['action_id']}",
                f"    element_id: {item['element_id']}",
                f"    style_id: {item['style_id']}",
                f"    removed_count: {item['removed_count']}",
                f"    removed_tags: {tag_summary or 'none'}",
            ]
        )
    if len(details) > max_items:
        lines.append(f"cleaned_conflict_details_truncated: {len(details) - max_items}")
    return lines


def collect_story_paragraphs(entries: dict[str, bytes], names: list[str]) -> dict[str, list]:
    stories: dict[str, list] = {}
    for name in names:
        content = entries.get(name)
        if not content:
            continue
        try:
            root = ET.fromstring(content)
        except ET.ParseError:
            continue
        stories[name] = root.findall(".//w:p", NS)
    return stories


def audit_body_style_definition_risk(document_root, entries: dict[str, bytes], style_id: str) -> list[str]:
    risky: list[str] = []
    include_implicit_normal = is_normal_style(style_id)
    for index, paragraph in enumerate(document_root.findall(".//w:tc//w:p", NS), start=1):
        if paragraph_uses_style(paragraph, style_id, include_implicit_normal):
            risky.append(f"word/document.xml:table-cell-p-{index}")
    story_names = [
        name
        for name in entries
        if name.startswith("word/header") and name.endswith(".xml")
        or name.startswith("word/footer") and name.endswith(".xml")
        or name in {"word/footnotes.xml", "word/endnotes.xml"}
    ]
    for name, paragraphs in collect_story_paragraphs(entries, sorted(story_names)).items():
        for index, paragraph in enumerate(paragraphs, start=1):
            if paragraph_uses_style(paragraph, style_id, include_implicit_normal):
                risky.append(f"{name}:p-{index}")
    return risky


def apply_run_format_rule(paragraph, rule: dict) -> None:
    font = rule.get("font_east_asia") or rule.get("font_ascii")
    ascii_font = rule.get("font_ascii") or font
    size = half_points(rule.get("font_size_pt"))
    bold = rule.get("bold")
    for run in paragraph.findall("./w:r", NS):
        rpr = run.find("./w:rPr", NS)
        if rpr is None:
            rpr = ET.Element(f"{{{W}}}rPr")
            run.insert(0, rpr)
        if font:
            rfonts = get_or_create(rpr, "rFonts")
            set_attr(rfonts, "eastAsia", font)
            set_attr(rfonts, "ascii", ascii_font or font)
            set_attr(rfonts, "hAnsi", ascii_font or font)
        if size:
            sz = get_or_create(rpr, "sz")
            set_attr(sz, "val", size)
            sz_cs = get_or_create(rpr, "szCs")
            set_attr(sz_cs, "val", size)
        if bold is not None:
            set_run_bool(rpr, "b", bool(bold))


def apply_paragraph_format_rule(paragraph, rule: dict) -> None:
    ppr = get_or_create(paragraph, "pPr", prepend=True)
    line_spacing_multiple = rule.get("line_spacing_multiple")
    line_spacing_pt = rule.get("line_spacing_pt")
    if line_spacing_multiple is not None or line_spacing_pt is not None:
        spacing = get_or_create(ppr, "spacing")
        if line_spacing_multiple is not None:
            set_attr(spacing, "line", int(round(float(line_spacing_multiple) * 240)))
            set_attr(spacing, "lineRule", "auto")
        elif line_spacing_pt is not None:
            set_attr(spacing, "line", int(round(float(line_spacing_pt) * 20)))
            set_attr(spacing, "lineRule", "exact")


def set_paragraph_style(paragraph, style_id: str, outline_level=None) -> None:
    ppr = get_or_create(paragraph, "pPr", prepend=True)
    pstyle = get_or_create(ppr, "pStyle")
    set_attr(pstyle, "val", style_id)
    if outline_level is not None:
        outline = get_or_create(ppr, "outlineLvl")
        set_attr(outline, "val", int(outline_level) - 1)


def find_style(styles_root, style_id: str):
    for style in styles_root.findall("./w:style", NS):
        if style.get(f"{{{W}}}styleId") == style_id:
            return style
    return None


def apply_style_definition(styles_root, style_id: str, rule: dict, outline_level=None) -> bool:
    style = find_style(styles_root, style_id)
    if style is None:
        return False
    rpr = get_or_create(style, "rPr")
    font = rule.get("font_east_asia") or rule.get("font_ascii")
    ascii_font = rule.get("font_ascii") or font
    if font:
        rfonts = get_or_create(rpr, "rFonts")
        set_attr(rfonts, "eastAsia", font)
        set_attr(rfonts, "ascii", ascii_font or font)
        set_attr(rfonts, "hAnsi", ascii_font or font)
    size = half_points(rule.get("font_size_pt"))
    if size:
        set_attr(get_or_create(rpr, "sz"), "val", size)
        set_attr(get_or_create(rpr, "szCs"), "val", size)
    if rule.get("bold") is not None:
        set_run_bool(rpr, "b", bool(rule.get("bold")))
    apply_style_paragraph_format(style, rule, outline_level)
    return True


def apply_table_cell_format(table, table_rules: dict) -> None:
    layout = table_rules.get("layout", {})
    header_rule = table_rules.get("header", {})
    body_rule = table_rules.get("body", {})
    tbl_pr = get_or_create(table, "tblPr", prepend=True)
    if layout.get("width"):
        tbl_w = get_or_create(tbl_pr, "tblW")
        set_attr(tbl_w, "w", layout["width"])
        set_attr(tbl_w, "type", layout.get("width_type") or "pct")
    if layout.get("alignment"):
        set_attr(get_or_create(tbl_pr, "jc"), "val", layout["alignment"])
    margins = get_or_create(tbl_pr, "tblCellMar")
    for side, key in (("top", "cell_margin_top"), ("bottom", "cell_margin_bottom"), ("left", "cell_margin_left"), ("right", "cell_margin_right")):
        if layout.get(key) is None:
            continue
        margin = get_or_create(margins, side)
        set_attr(margin, "w", layout[key])
        set_attr(margin, "type", "dxa")
    for row_index, row in enumerate(table.findall("./w:tr", NS)):
        tr_pr = get_or_create(row, "trPr", prepend=True)
        if row_index == 0 and str(layout.get("header_repeat", "true")).lower() == "true":
            get_or_create(tr_pr, "tblHeader")
        rule = header_rule if row_index == 0 else body_rule
        for paragraph in row.findall(".//w:p", NS):
            apply_paragraph_format_rule(paragraph, rule)
            apply_run_format_rule(paragraph, rule)
        for cell in row.findall("./w:tc", NS):
            tc_pr = get_or_create(cell, "tcPr", prepend=True)
            if layout.get("vertical_alignment"):
                set_attr(get_or_create(tc_pr, "vAlign"), "val", layout["vertical_alignment"])
            if row_index == 0 and header_rule.get("shading_fill"):
                shading = get_or_create(tc_pr, "shd")
                set_attr(shading, "val", "clear")
                set_attr(shading, "color", "auto")
                set_attr(shading, "fill", header_rule["shading_fill"])


def apply_repairs(input_docx: Path, output_docx: Path, actions: list[dict], styles: dict, table_rules: dict) -> dict:
    shutil.copy2(input_docx, output_docx)
    with zipfile.ZipFile(output_docx, "r") as zf:
        entries = {item.filename: zf.read(item.filename) for item in zf.infolist() if not item.is_dir()}
    document_root = ET.fromstring(entries["word/document.xml"])
    styles_root = ET.fromstring(entries["word/styles.xml"]) if "word/styles.xml" in entries else None
    paragraphs = document_root.findall(".//w:p", NS)
    tables = document_root.findall(".//w:tbl", NS)

    applied = 0
    tables_applied = 0
    skipped = 0
    blocked = 0
    missing_styles: set[str] = set()
    blocked_reasons: list[str] = []
    cleaned_direct_conflicts = 0
    cleaned_conflict_details: list[dict] = []

    heading_targets: dict[str, dict] = {}
    heading_actions: list[tuple[dict, int, str, dict]] = []
    body_style_actions: list[dict] = []
    style_definitions_applied: set[str] = set()

    for action in actions:
        action_type = action.get("action_type")
        element_id = action.get("element_id", "")
        expected_role = action.get("expected_role") or "body-paragraph"
        if action_type == "map_heading_native_style" and element_id.startswith("p-"):
            rule = styles.get(expected_role, {})
            style_id = rule.get("word_style_id")
            if not style_id:
                skipped += 1
                continue
            try:
                index = int(element_id.rsplit("-", 1)[1]) - 1
                paragraphs[index]
            except (IndexError, ValueError):
                skipped += 1
                continue
            heading_actions.append((action, index, style_id, rule))
            plan = heading_targets.setdefault(style_id, {"indices": set(), "rule": rule, "roles": set()})
            plan["indices"].add(index)
            plan["roles"].add(expected_role)
        elif action_type == "apply_body_style_definition":
            body_style_actions.append(action)

    if styles_root is None and (heading_actions or body_style_actions):
        for _, _, style_id, _ in heading_actions:
            missing_styles.add(style_id)
        body_style_id = styles.get("body-paragraph", {}).get("word_style_id") or "Normal"
        missing_styles.add(body_style_id)
        blocked += len(heading_actions) + len(body_style_actions)
        blocked_reasons.append("word/styles.xml:missing")
        heading_actions = []
        body_style_actions = []

    safe_heading_style_defs: set[str] = set()
    blocked_heading_style_defs: set[str] = set()
    if styles_root is not None:
        for style_id, plan in heading_targets.items():
            if find_style(styles_root, style_id) is None:
                missing_styles.add(style_id)
                blocked += len(plan["indices"])
                blocked_heading_style_defs.add(style_id)
                blocked_reasons.append(f"{style_id}:missing-native-style")
                continue
            used_indices = {
                index
                for index, paragraph in enumerate(paragraphs)
                if paragraph_style_id(paragraph) == style_id
            }
            outside_targets = used_indices - plan["indices"]
            if outside_targets:
                blocked_heading_style_defs.add(style_id)
                blocked_reasons.append(f"{style_id}:used-outside-targets:{len(outside_targets)}")
                continue
            if style_id not in style_definitions_applied:
                apply_style_definition(styles_root, style_id, plan["rule"].get("format", {}), plan["rule"].get("outline_level"))
                style_definitions_applied.add(style_id)
            safe_heading_style_defs.add(style_id)

    if styles_root is not None and body_style_actions:
        body_rule = styles.get("body-paragraph", {})
        body_style_id = body_rule.get("word_style_id") or "Normal"
        if find_style(styles_root, body_style_id) is None:
            blocked += len(body_style_actions)
            missing_styles.add(body_style_id)
            blocked_reasons.append(f"{body_style_id}:missing-native-style")
        else:
            risky_body_uses = audit_body_style_definition_risk(document_root, entries, body_style_id)
            if risky_body_uses:
                blocked += len(body_style_actions)
                blocked_reasons.append(f"{body_style_id}:body-style-risk:{len(risky_body_uses)}")
            elif body_style_id not in style_definitions_applied:
                apply_style_definition(styles_root, body_style_id, body_rule.get("format", {}), body_rule.get("outline_level"))
                style_definitions_applied.add(body_style_id)
                applied += len(body_style_actions)

    for action, index, style_id, rule in heading_actions:
        if style_id in missing_styles:
            continue
        if style_id in blocked_heading_style_defs:
            blocked += 1
            continue
        set_paragraph_style(paragraphs[index], style_id, rule.get("outline_level"))
        if style_id in safe_heading_style_defs:
            conflict_detail = clear_direct_format_conflicts(paragraphs[index], rule.get("format", {}))
            cleaned_direct_conflicts += conflict_detail["count"]
            if conflict_detail["count"]:
                cleaned_conflict_details.append(
                    {
                        "action_id": action.get("action_id", ""),
                        "element_id": action.get("element_id", ""),
                        "style_id": style_id,
                        "removed_count": conflict_detail["count"],
                        "removed_tags": conflict_detail["removed_tags"],
                    }
                )
        applied += 1

    for action in actions:
        action_type = action.get("action_type")
        element_id = action.get("element_id", "")
        try:
            if action_type in {"map_heading_native_style", "apply_body_style_definition"}:
                continue
            if action_type == "apply_body_direct_format":
                if action.get("format_write_strategy") != "direct-format-override" or not element_id.startswith("p-"):
                    skipped += 1
                    continue
                index = int(element_id.rsplit("-", 1)[1]) - 1
                rule = styles.get("body-paragraph", {}).get("format", {})
                apply_paragraph_format_rule(paragraphs[index], rule)
                apply_run_format_rule(paragraphs[index], rule)
                applied += 1
            elif action_type == "apply_table_cell_format" and element_id.startswith("t-"):
                index = int(element_id.rsplit("-", 1)[1]) - 1
                apply_table_cell_format(tables[index], table_rules)
                tables_applied += 1
            else:
                skipped += 1
        except (IndexError, ValueError):
            skipped += 1

    entries["word/document.xml"] = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)
    if styles_root is not None:
        entries["word/styles.xml"] = ET.tostring(styles_root, encoding="utf-8", xml_declaration=True)
    temp_output = output_docx.with_suffix(".tmp.docx")
    with zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in entries.items():
            zf.writestr(name, content)
    temp_output.replace(output_docx)
    return {
        "applied": applied,
        "tables_applied": tables_applied,
        "skipped": skipped,
        "blocked": blocked,
        "missing_styles": sorted(missing_styles),
        "style_definitions_applied": sorted(style_definitions_applied),
        "blocked_reasons": blocked_reasons,
        "cleaned_direct_conflicts": cleaned_direct_conflicts,
        "cleaned_conflict_details": cleaned_conflict_details,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repair-plan", required=True)
    parser.add_argument("--log", required=True)
    args = parser.parse_args()

    plan_path = Path(args.repair_plan)
    plan = parse_scalar_yaml(plan_path)
    working_value = plan.get("working_docx", "")
    output_value = plan.get("output_docx", "")
    if not working_value:
        raise SystemExit("repair_plan.yaml 缺少 working_docx")
    if not output_value:
        raise SystemExit("repair_plan.yaml 缺少 output_docx")
    working = Path(working_value)
    output = Path(output_value)
    if not working.exists():
        raise SystemExit(f"工作副本不存在：{working}")

    output.parent.mkdir(parents=True, exist_ok=True)
    actions = parse_auto_fix_actions(plan_path)
    styles = parse_style_map(rule_file(plan_path, "style-map.yaml"))
    table_rules = parse_table_rules(rule_file(plan_path, "table-rules.yaml"))
    result = apply_repairs(working, output, actions, styles, table_rules)
    status = "applied_v2_repairs" if result["blocked"] == 0 else "blocked_style_definition_safety"
    log = [
        f"repair_plan: {plan_path}",
        f"executed_at: {datetime.now(TZ).isoformat()}",
        f"working_docx: {working}",
        f"output_docx: {output}",
        f"status: {status}",
        f"actions_total: {len(actions)}",
        f"actions_applied: {result['applied']}",
        f"tables_applied: {result['tables_applied']}",
        f"actions_skipped: {result['skipped']}",
        f"actions_blocked: {result['blocked']}",
        f"missing_native_styles: {', '.join(result['missing_styles']) if result['missing_styles'] else 'none'}",
        f"style_definitions_applied: {', '.join(result['style_definitions_applied']) if result['style_definitions_applied'] else 'none'}",
        f"style_definition_blocked_reasons: {'; '.join(result['blocked_reasons']) if result['blocked_reasons'] else 'none'}",
        f"cleaned_direct_conflicts: {result['cleaned_direct_conflicts']}",
        *format_conflict_details(result["cleaned_conflict_details"]),
        "note: V2 修复器不会创建任何缺失样式；缺失原生样式进入阻塞或人工确认。",
    ]
    Path(args.log).write_text("\n".join(log) + "\n", encoding="utf-8")
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
