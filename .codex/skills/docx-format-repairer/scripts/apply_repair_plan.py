#!/usr/bin/env python3
"""执行 repair_plan.yaml 中允许的 v3 白名单修复动作。"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[4]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.utils.simple_yaml import load_yaml
from scripts.validation.check_repair_plan import validate_repair_plan


TZ = timezone(timedelta(hours=8))
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = NS["w"]
ET.register_namespace("w", W)
EXECUTABLE_ACTIONS = {
    "map_heading_native_style",
    "apply_body_direct_format",
}


def w_attr(node: ET.Element | None, name: str) -> str | None:
    """读取 w:* 命名空间属性。"""
    if node is None:
        return None
    return node.get(f"{{{W}}}{name}")


def set_attr(node: ET.Element, name: str, value: Any) -> None:
    """写入 w:* 命名空间属性。"""
    node.set(f"{{{W}}}{name}", str(value))


def get_or_create(parent: ET.Element, name: str, prepend: bool = False) -> ET.Element:
    """获取或创建子节点。"""
    node = parent.find(f"./w:{name}", NS)
    if node is None:
        node = ET.SubElement(parent, f"{{{W}}}{name}")
        if prepend and len(parent) > 1:
            parent.remove(node)
            parent.insert(0, node)
    return node


def bool_run_prop(parent: ET.Element, name: str, enabled: bool) -> None:
    """设置 run 布尔格式。"""
    node = get_or_create(parent, name)
    if enabled:
        node.attrib.pop(f"{{{W}}}val", None)
    else:
        set_attr(node, "val", "0")


def cm_to_twips(value: Any) -> str | None:
    """厘米转 twips。"""
    if value in (None, ""):
        return None
    return str(int(round(float(value) * 567)))


def half_points(value: Any) -> str | None:
    """磅值转半磅。"""
    if value in (None, ""):
        return None
    return str(int(round(float(value) * 2)))


def paragraph_index(element_id: str) -> int:
    """从 p-00001 提取段落下标。"""
    if not element_id.startswith("p-"):
        raise ValueError(f"仅支持段落元素：{element_id}")
    return int(element_id.rsplit("-", 1)[1]) - 1


def style_ids(styles_root: ET.Element | None) -> set[str]:
    """收集文档中已存在的 Word 样式 ID。"""
    if styles_root is None:
        return set()
    return {
        style.get(f"{{{W}}}styleId", "")
        for style in styles_root.findall("./w:style", NS)
        if style.get(f"{{{W}}}styleId")
    }


def set_paragraph_style(paragraph: ET.Element, style_id: str, outline_level: Any = None) -> None:
    """设置段落样式，不创建样式定义。"""
    ppr = get_or_create(paragraph, "pPr", prepend=True)
    set_attr(get_or_create(ppr, "pStyle"), "val", style_id)
    if outline_level is not None:
        set_attr(get_or_create(ppr, "outlineLvl"), "val", int(outline_level) - 1)


def apply_paragraph_format(paragraph: ET.Element, values: dict[str, Any]) -> None:
    """应用段落直接格式。"""
    ppr = get_or_create(paragraph, "pPr", prepend=True)
    alignment = values.get("alignment")
    if alignment:
        set_attr(get_or_create(ppr, "jc"), "val", "both" if alignment == "justify" else alignment)

    spacing_values = ("line_spacing_multiple", "line_spacing_pt", "space_before_pt", "space_after_pt")
    if any(values.get(key) is not None for key in spacing_values):
        spacing = get_or_create(ppr, "spacing")
        if values.get("line_spacing_multiple") is not None:
            set_attr(spacing, "line", int(round(float(values["line_spacing_multiple"]) * 240)))
            set_attr(spacing, "lineRule", "auto")
        if values.get("line_spacing_pt") is not None:
            set_attr(spacing, "line", int(round(float(values["line_spacing_pt"]) * 20)))
            set_attr(spacing, "lineRule", "exact")
        if values.get("space_before_pt") is not None:
            set_attr(spacing, "before", int(round(float(values["space_before_pt"]) * 20)))
        if values.get("space_after_pt") is not None:
            set_attr(spacing, "after", int(round(float(values["space_after_pt"]) * 20)))

    indent_values = {
        "firstLine": cm_to_twips(values.get("first_line_indent_cm")),
        "left": cm_to_twips(values.get("left_indent_cm")),
        "right": cm_to_twips(values.get("right_indent_cm")),
        "hanging": cm_to_twips(values.get("hanging_indent_cm")),
    }
    if any(value is not None for value in indent_values.values()):
        ind = get_or_create(ppr, "ind")
        for key, value in indent_values.items():
            if value is None:
                continue
            if key == "firstLine":
                ind.attrib.pop(f"{{{W}}}hanging", None)
            if key == "hanging":
                ind.attrib.pop(f"{{{W}}}firstLine", None)
            set_attr(ind, key, value)


def apply_run_format(paragraph: ET.Element, values: dict[str, Any]) -> None:
    """应用 run 直接格式。"""
    font = values.get("font_east_asia") or values.get("font_ascii")
    ascii_font = values.get("font_ascii") or font
    size = half_points(values.get("font_size_pt"))
    bold = values.get("bold")
    for run in paragraph.findall("./w:r", NS):
        rpr = get_or_create(run, "rPr", prepend=True)
        if font:
            fonts = get_or_create(rpr, "rFonts")
            set_attr(fonts, "eastAsia", font)
            set_attr(fonts, "ascii", ascii_font or font)
            set_attr(fonts, "hAnsi", ascii_font or font)
        if size:
            set_attr(get_or_create(rpr, "sz"), "val", size)
            set_attr(get_or_create(rpr, "szCs"), "val", size)
        if bold is not None:
            bool_run_prop(rpr, "b", bool(bold))
            bool_run_prop(rpr, "bCs", bool(bold))


def apply_action(
    action: dict[str, Any],
    paragraphs: list[ET.Element],
    available_style_ids: set[str],
) -> tuple[str, str]:
    """执行单个动作，返回状态和说明。"""
    action_type = action.get("action_type")
    if action_type not in EXECUTABLE_ACTIONS:
        return "skipped", "首阶段修复器尚未实现该白名单动作的写回逻辑"

    element_id = action.get("target", {}).get("element_id", "")
    index = paragraph_index(element_id)
    if index < 0 or index >= len(paragraphs):
        return "rejected", "目标段落不存在"
    paragraph = paragraphs[index]
    after = action.get("after") or {}

    if action_type == "map_heading_native_style":
        style_id = after.get("style_id") or after.get("style")
        if not style_id:
            return "rejected", "缺少 after.style_id"
        if style_id not in available_style_ids:
            return "rejected", f"文档中不存在原生样式 {style_id}"
        set_paragraph_style(paragraph, style_id, after.get("outline_level"))
        return "executed", f"已设置段落样式 {style_id}"

    if action_type == "apply_body_direct_format":
        apply_paragraph_format(paragraph, after)
        apply_run_format(paragraph, after)
        return "executed", "已应用段落直接格式"

    return "skipped", "未处理动作"


def validate_output_docx(path: Path) -> bool:
    """验证输出 DOCX 基本 OOXML 结构。"""
    try:
        with zipfile.ZipFile(path, "r") as archive:
            archive.getinfo("word/document.xml")
            archive.testzip()
        return True
    except (KeyError, zipfile.BadZipFile):
        return False


def apply_plan(plan: dict[str, Any], plan_path: Path, log_path: Path) -> dict[str, Any]:
    """执行修复计划。"""
    working_docx = Path(plan["working_docx"])
    output_docx = Path(plan["output_docx"])
    source_docx = Path(plan["source_docx"])
    if output_docx.resolve() in {source_docx.resolve(), working_docx.resolve()}:
        raise SystemExit("output_docx 不得覆盖 source_docx 或 working_docx")
    if not working_docx.exists():
        raise SystemExit(f"工作副本不存在：{working_docx}")

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(working_docx, output_docx)
    with zipfile.ZipFile(output_docx, "r") as archive:
        entries = {item.filename: archive.read(item.filename) for item in archive.infolist() if not item.is_dir()}

    document_root = ET.fromstring(entries["word/document.xml"])
    styles_root = ET.fromstring(entries["word/styles.xml"]) if "word/styles.xml" in entries else None
    paragraphs = document_root.findall(".//w:p", NS)
    available_style_ids = style_ids(styles_root)

    action_results: list[dict[str, Any]] = []
    for action in plan.get("actions", []):
        if action.get("auto_fix_policy") != "auto-fix":
            action_results.append(
                {
                    "action_id": action.get("action_id"),
                    "status": "skipped",
                    "reason": "非 auto-fix 动作不执行",
                }
            )
            continue
        try:
            status, reason = apply_action(action, paragraphs, available_style_ids)
        except (ValueError, IndexError) as exc:
            status, reason = "rejected", str(exc)
        action_results.append(
            {
                "action_id": action.get("action_id"),
                "action_type": action.get("action_type"),
                "element_id": action.get("target", {}).get("element_id"),
                "status": status,
                "reason": reason,
            }
        )

    entries["word/document.xml"] = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)
    temp_output = output_docx.with_suffix(".tmp.docx")
    with zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)
    temp_output.replace(output_docx)

    result = {
        "repair_plan": str(plan_path),
        "executed_at": datetime.now(TZ).isoformat(),
        "working_docx": str(working_docx),
        "output_docx": str(output_docx),
        "actions_total": len(plan.get("actions", [])),
        "actions_executed": sum(1 for item in action_results if item["status"] == "executed"),
        "actions_skipped": sum(1 for item in action_results if item["status"] == "skipped"),
        "actions_rejected": sum(1 for item in action_results if item["status"] == "rejected"),
        "output_docx_valid": validate_output_docx(output_docx),
        "actions": action_results,
    }
    log_path.parent.mkdir(parents=True, exist_ok=True)
    log_path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return result


def main_from_args(argv: list[str] | None = None) -> int:
    """命令行入口，便于测试复用。"""
    parser = argparse.ArgumentParser(description="执行 repair_plan.yaml 白名单动作")
    parser.add_argument("--repair-plan", required=True, type=Path)
    parser.add_argument("--log", required=True, type=Path)
    args = parser.parse_args(argv)

    plan = load_yaml(args.repair_plan)
    if not isinstance(plan, dict):
        print("repair_plan 根节点必须是对象")
        return 1
    errors = validate_repair_plan(plan)
    if errors:
        for error in errors:
            print(error)
        return 1
    result = apply_plan(plan, args.repair_plan, args.log)
    if not result["output_docx_valid"]:
        print("输出 DOCX OOXML 校验失败")
        return 1
    print(result["output_docx"])
    return 0


def main() -> int:
    """脚本入口。"""
    return main_from_args()


if __name__ == "__main__":
    raise SystemExit(main())
