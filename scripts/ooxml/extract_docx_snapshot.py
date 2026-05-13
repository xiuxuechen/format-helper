#!/usr/bin/env python3
"""从 DOCX 提取客观 OOXML 事实快照。"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


TZ = timezone(timedelta(hours=8))
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = NS["w"]
FORMAT_EMPTY_VALUES = {None, ""}
SOURCE_CONFIDENCE = {
    "direct": 0.98,
    "style_inherit": 0.9,
    "doc_defaults": 0.85,
    "theme": 0.75,
    "word_ui_default": 0.6,
    "unresolved": 0.0,
    "legacy": 0.5,
}
PARAGRAPH_SOURCE_SLOTS = {
    "alignment",
    "outline_level",
    "first_line_indent_cm",
    "left_indent_cm",
    "right_indent_cm",
    "line_spacing_raw",
    "line_spacing_rule",
    "line_spacing_multiple",
    "line_spacing_pt",
    "space_before_pt",
    "space_after_pt",
}
RUN_SOURCE_SLOTS = {
    "font_east_asia",
    "font_ascii",
    "font_size_pt",
    "bold",
}


def w_attr(node: ET.Element | None, name: str) -> str | None:
    """读取 w:* 命名空间属性。"""
    if node is None:
        return None
    return node.get(f"{{{W}}}{name}")


def text_of(node: ET.Element) -> str:
    """提取节点内文本。"""
    return "".join(text.text or "" for text in node.findall(".//w:t", NS))


def int_attr(node: ET.Element | None, name: str) -> int | None:
    """读取整数属性。"""
    value = w_attr(node, name)
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def twips_to_cm(value: int | None) -> float | None:
    """twips 转厘米。"""
    if value is None:
        return None
    return round(value / 567, 2)


def half_points_to_pt(value: str | None) -> float | None:
    """半磅转磅。"""
    if value is None:
        return None
    try:
        return int(value) / 2
    except ValueError:
        return None


def paragraph_format(paragraph: ET.Element) -> dict[str, Any]:
    """抽取段落格式事实。"""
    ppr = paragraph.find("./w:pPr", NS)
    return paragraph_format_from_ppr(ppr)


def paragraph_format_from_ppr(ppr: ET.Element | None) -> dict[str, Any]:
    """从 pPr 节点抽取段落格式事实。"""
    spacing = ppr.find("./w:spacing", NS) if ppr is not None else None
    ind = ppr.find("./w:ind", NS) if ppr is not None else None
    jc = ppr.find("./w:jc", NS) if ppr is not None else None
    outline = ppr.find("./w:outlineLvl", NS) if ppr is not None else None
    result: dict[str, Any] = {
        "alignment": w_attr(jc, "val"),
        "outline_level": None,
        "first_line_indent_cm": twips_to_cm(int_attr(ind, "firstLine")),
        "left_indent_cm": twips_to_cm(int_attr(ind, "left")),
        "right_indent_cm": twips_to_cm(int_attr(ind, "right")),
        "line_spacing_raw": w_attr(spacing, "line"),
        "line_spacing_rule": w_attr(spacing, "lineRule"),
        "space_before_pt": None,
        "space_after_pt": None,
    }
    outline_value = int_attr(outline, "val")
    if outline_value is not None:
        result["outline_level"] = outline_value + 1
    before = int_attr(spacing, "before")
    after = int_attr(spacing, "after")
    if before is not None:
        result["space_before_pt"] = before / 20
    if after is not None:
        result["space_after_pt"] = after / 20
    apply_line_spacing_derived_slots(result)
    return result


def run_format(paragraph: ET.Element) -> dict[str, Any]:
    """抽取首个 run 格式事实。"""
    run = paragraph.find("./w:r", NS)
    rpr = run.find("./w:rPr", NS) if run is not None else None
    return run_format_from_rpr(rpr)


def run_format_from_rpr(rpr: ET.Element | None) -> dict[str, Any]:
    """从 rPr 节点抽取字符格式事实。"""
    fonts = rpr.find("./w:rFonts", NS) if rpr is not None else None
    size = rpr.find("./w:sz", NS) if rpr is not None else None
    bold = rpr.find("./w:b", NS) if rpr is not None else None
    return {
        "font_east_asia": w_attr(fonts, "eastAsia"),
        "font_ascii": w_attr(fonts, "ascii"),
        "font_size_pt": half_points_to_pt(w_attr(size, "val")),
        "bold": None if bold is None else w_attr(bold, "val") not in {"0", "false", "False"},
    }


def paragraph_style(paragraph: ET.Element) -> str | None:
    """读取段落样式 ID。"""
    return w_attr(paragraph.find("./w:pPr/w:pStyle", NS), "val")


def run_style(run: ET.Element | None) -> str | None:
    """读取 run 样式 ID。"""
    if run is None:
        return None
    return w_attr(run.find("./w:rPr/w:rStyle", NS), "val")


def non_empty_format(data: dict[str, Any]) -> dict[str, Any]:
    """剔除空值，保留显式 false 和 0。"""
    return {key: value for key, value in data.items() if value not in FORMAT_EMPTY_VALUES}


def merge_formats(*formats: dict[str, Any]) -> dict[str, Any]:
    """按顺序合并格式，后者覆盖前者的非空值。"""
    merged: dict[str, Any] = {}
    for item in formats:
        merged.update(non_empty_format(item))
    return merged


def source_value(value: Any, source: str) -> dict[str, Any]:
    """构造带来源的槽位值对象。"""
    return {"value": value, "source": source, "confidence": SOURCE_CONFIDENCE[source]}


def non_empty_source_format(data: dict[str, Any], source: str) -> dict[str, dict[str, Any]]:
    """把裸格式值转成带 source 的非空格式。"""
    return {key: source_value(value, source) for key, value in data.items() if value not in FORMAT_EMPTY_VALUES}


def merge_source_formats(*formats: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    """按来源优先级合并格式对象，后者覆盖前者。"""
    merged: dict[str, dict[str, Any]] = {}
    for item in formats:
        merged.update(item)
    return merged


def ensure_source_slots(data: dict[str, dict[str, Any]], slot_names: set[str]) -> dict[str, dict[str, Any]]:
    """确保所有登记槽位显式存在，缺失时标为 unresolved。"""
    result = dict(data)
    for slot_name in sorted(slot_names):
        if slot_name not in result:
            result[slot_name] = source_value(None, "unresolved")
    return result


def apply_line_spacing_derived_slots(data: dict[str, Any]) -> None:
    """从 line + lineRule 推导倍数或固定磅值行距。"""
    raw = data.get("line_spacing_raw")
    if raw in FORMAT_EMPTY_VALUES:
        return
    try:
        line_value = float(raw)
    except (TypeError, ValueError):
        return
    rule = data.get("line_spacing_rule") or "auto"
    if rule == "auto":
        data["line_spacing_multiple"] = round(line_value / 240.0, 3)
    elif rule in {"exact", "atLeast"}:
        data["line_spacing_pt"] = round(line_value / 20.0, 3)


class StyleResolver:
    """解析 styles.xml 中样式继承后的有效格式。"""

    def __init__(self, root: ET.Element | None):
        self.styles: dict[str, dict[str, Any]] = {}
        self.default_paragraph_format: dict[str, Any] = {}
        self.default_run_format: dict[str, Any] = {}
        if root is None:
            return
        self.default_paragraph_format = non_empty_format(
            paragraph_format_from_ppr(root.find("./w:docDefaults/w:pPrDefault/w:pPr", NS))
        )
        self.default_run_format = non_empty_format(
            run_format_from_rpr(root.find("./w:docDefaults/w:rPrDefault/w:rPr", NS))
        )
        for style in root.findall("./w:style", NS):
            style_id = w_attr(style, "styleId")
            if not style_id:
                continue
            self.styles[style_id] = {
                "style_id": style_id,
                "style_type": w_attr(style, "type"),
                "style_name": w_attr(style.find("./w:name", NS), "val"),
                "based_on": w_attr(style.find("./w:basedOn", NS), "val"),
                "paragraph_format": non_empty_format(paragraph_format_from_ppr(style.find("./w:pPr", NS))),
                "run_format": non_empty_format(run_format_from_rpr(style.find("./w:rPr", NS))),
            }

    def style_name(self, style_id: str | None) -> str | None:
        """读取样式名。"""
        if not style_id:
            return None
        item = self.styles.get(style_id)
        return item.get("style_name") if item else None

    def resolve_style(self, style_id: str | None, seen: set[str] | None = None) -> tuple[dict[str, Any], dict[str, Any]]:
        """解析单个样式的继承链。"""
        if not style_id or style_id not in self.styles:
            return {}, {}
        seen = seen or set()
        if style_id in seen:
            return {}, {}
        seen.add(style_id)
        item = self.styles[style_id]
        base_p, base_r = self.resolve_style(item.get("based_on"), seen)
        return (
            merge_formats(base_p, item.get("paragraph_format", {})),
            merge_formats(base_r, item.get("run_format", {})),
        )

    def resolve_style_with_source(
        self, style_id: str | None, seen: set[str] | None = None
    ) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
        """解析样式继承链，并保留 style_inherit 来源。"""
        if not style_id or style_id not in self.styles:
            return {}, {}
        seen = seen or set()
        if style_id in seen:
            return {}, {}
        seen.add(style_id)
        item = self.styles[style_id]
        base_p, base_r = self.resolve_style_with_source(item.get("based_on"), seen)
        return (
            merge_source_formats(base_p, non_empty_source_format(item.get("paragraph_format", {}), "style_inherit")),
            merge_source_formats(base_r, non_empty_source_format(item.get("run_format", {}), "style_inherit")),
        )

    def resolved_paragraph_format(self, paragraph: ET.Element) -> dict[str, Any]:
        """解析段落最终有效格式。"""
        style_p, _ = self.resolve_style(paragraph_style(paragraph))
        direct_p = paragraph_format(paragraph)
        return merge_formats(self.default_paragraph_format, style_p, direct_p)

    def resolved_paragraph_format_with_source(self, paragraph: ET.Element) -> dict[str, dict[str, Any]]:
        """解析段落最终格式，并为每个槽位写明来源。"""
        style_p, _ = self.resolve_style_with_source(paragraph_style(paragraph))
        direct_p = non_empty_source_format(paragraph_format(paragraph), "direct")
        defaults = non_empty_source_format(self.default_paragraph_format, "doc_defaults")
        return ensure_source_slots(merge_source_formats(defaults, style_p, direct_p), PARAGRAPH_SOURCE_SLOTS)

    def resolved_run_format(self, paragraph: ET.Element, run: ET.Element | None = None) -> dict[str, Any]:
        """解析 run 最终有效格式。"""
        if run is None:
            run = paragraph.find("./w:r", NS)
        _, paragraph_style_r = self.resolve_style(paragraph_style(paragraph))
        _, run_style_r = self.resolve_style(run_style(run))
        direct_r = run_format_from_rpr(run.find("./w:rPr", NS) if run is not None else None)
        return merge_formats(self.default_run_format, paragraph_style_r, run_style_r, direct_r)

    def resolved_run_format_with_source(
        self, paragraph: ET.Element, run: ET.Element | None = None
    ) -> dict[str, dict[str, Any]]:
        """解析 run 最终格式，并为每个槽位写明来源。"""
        if run is None:
            run = paragraph.find("./w:r", NS)
        _, paragraph_style_r = self.resolve_style_with_source(paragraph_style(paragraph))
        _, run_style_r = self.resolve_style_with_source(run_style(run))
        direct_r = non_empty_source_format(run_format_from_rpr(run.find("./w:rPr", NS) if run is not None else None), "direct")
        defaults = non_empty_source_format(self.default_run_format, "doc_defaults")
        return ensure_source_slots(merge_source_formats(defaults, paragraph_style_r, run_style_r, direct_r), RUN_SOURCE_SLOTS)


def table_cell_role(row_index: int, column_index: int) -> str:
    """给单元格提供保守的角色候选。"""
    if row_index == 1:
        return "header-row-cell"
    if column_index == 1:
        return "field-name-cell"
    return "value-cell"


def run_items(paragraph: ET.Element, resolver: StyleResolver, with_source: bool = True) -> list[dict[str, Any]]:
    """抽取段落下所有 run 的文本和格式。"""
    items: list[dict[str, Any]] = []
    for index, run in enumerate(paragraph.findall("./w:r", NS), start=1):
        text = text_of(run).strip()
        direct = run_format_from_rpr(run.find("./w:rPr", NS))
        resolved = (
            resolver.resolved_run_format_with_source(paragraph, run)
            if with_source
            else resolver.resolved_run_format(paragraph, run)
        )
        items.append(
            {
                "run_index": index,
                "text_preview": text[:80],
                "run_style_id": run_style(run),
                "run_format": direct,
                "resolved_run_format": resolved,
            }
        )
    return items


def cell_format_summary(paragraphs: list[dict[str, Any]], vertical_alignment: str | None = None) -> dict[str, Any]:
    """汇总单元格格式，用于表格正文审计。"""
    fonts: set[str] = set()
    sizes: set[float] = set()
    has_bold = False
    for paragraph in paragraphs:
        for run in paragraph.get("runs", []):
            resolved = run.get("resolved_run_format") or {}
            direct = run.get("run_format") or {}
            resolved_bold = format_slot_value(resolved.get("bold"))
            direct_bold = format_slot_value(direct.get("bold"))
            if resolved_bold is True or direct_bold is True:
                has_bold = True
            font = format_slot_value(resolved.get("font_east_asia")) or format_slot_value(direct.get("font_east_asia"))
            size = format_slot_value(resolved.get("font_size_pt")) or format_slot_value(direct.get("font_size_pt"))
            if font:
                fonts.add(str(font))
            if isinstance(size, (int, float)):
                sizes.add(float(size))
    return {
        "has_bold": has_bold,
        "font_east_asia_values": sorted(fonts),
        "font_size_pt_values": sorted(sizes),
        "vertical_alignment": vertical_alignment,
    }


def format_slot_value(value: Any) -> Any:
    """兼容读取裸值和 {value, source, confidence} 槽位对象。"""
    if isinstance(value, dict) and "value" in value:
        return value.get("value")
    return value


def table_info(
    table: ET.Element,
    index: int,
    paragraph_index_by_id: dict[int, int],
    resolver: StyleResolver,
    with_source: bool = True,
) -> dict[str, Any]:
    """抽取表格基础事实。"""
    rows = table.findall("./w:tr", NS)
    column_count = 0
    cells: list[dict[str, Any]] = []
    for row_index, row in enumerate(rows, start=1):
        row_cells = row.findall("./w:tc", NS)
        column_count = max(column_count, len(row_cells))
        for column_index, cell in enumerate(row_cells, start=1):
            paragraph_items: list[dict[str, Any]] = []
            cell_id = f"table-{index:04d}-r{row_index:03d}-c{column_index:03d}"
            vertical_alignment = w_attr(cell.find("./w:tcPr/w:vAlign", NS), "val")
            for paragraph_offset, paragraph in enumerate(cell.findall(".//w:p", NS), start=1):
                text = text_of(paragraph).strip()
                paragraph_index = paragraph_index_by_id.get(id(paragraph))
                paragraph_items.append(
                    {
                        "fact_id": f"{cell_id}-p{paragraph_offset:03d}",
                        "fact_kind": "table_cell_paragraph",
                        "locator": {
                            "table_index": index,
                            "row_index": row_index,
                            "column_index": column_index,
                            "paragraph_index": paragraph_index,
                        },
                        "paragraph_index": paragraph_index,
                        "text_preview": text[:120],
                        "style_id": paragraph_style(paragraph),
                        "style_name": resolver.style_name(paragraph_style(paragraph)),
                        "paragraph_format": paragraph_format(paragraph),
                        "resolved_paragraph_format": (
                            resolver.resolved_paragraph_format_with_source(paragraph)
                            if with_source
                            else resolver.resolved_paragraph_format(paragraph)
                        ),
                        "run_format": run_format(paragraph),
                        "resolved_run_format": (
                            resolver.resolved_run_format_with_source(paragraph)
                            if with_source
                            else resolver.resolved_run_format(paragraph)
                        ),
                        "runs": run_items(paragraph, resolver, with_source=with_source),
                    }
                )
            cells.append(
                {
                    "cell_id": cell_id,
                    "row_index": row_index,
                    "column_index": column_index,
                    "cell_role_candidate": table_cell_role(row_index, column_index),
                    "vertical_alignment": vertical_alignment,
                    "text_preview": text_of(cell).strip()[:120],
                    "paragraphs": paragraph_items,
                    "format_summary": cell_format_summary(paragraph_items, vertical_alignment),
                }
            )
    for row in rows:
        column_count = max(column_count, len(row.findall("./w:tc", NS)))
    return {
        "element_id": f"table-{index:04d}",
        "table_index": index,
        "row_count": len(rows),
        "column_count": column_count,
        "text_preview": text_of(table).strip()[:80],
        "cells": cells,
    }


def page_number_format(section: ET.Element | None) -> str:
    """抽取 sectPr/w:pgNumType 页码格式；缺失时显式标为 none。"""
    if section is None:
        return "unresolved"
    page_number = section.find("./w:pgNumType", NS)
    if page_number is None:
        return "none"
    return w_attr(page_number, "fmt") or "unresolved"


def page_setup_from_section(section: ET.Element | None) -> dict[str, Any]:
    """抽取分节页面设置事实，不解析 header/footer relationship。"""
    page_size = section.find("./w:pgSz", NS) if section is not None else None
    margin = section.find("./w:pgMar", NS) if section is not None else None
    return {
        "page_orientation": w_attr(page_size, "orient") or "portrait",
        "page_width_twips": int_attr(page_size, "w"),
        "page_height_twips": int_attr(page_size, "h"),
        "margin_top_cm": twips_to_cm(int_attr(margin, "top")),
        "margin_bottom_cm": twips_to_cm(int_attr(margin, "bottom")),
        "margin_left_cm": twips_to_cm(int_attr(margin, "left")),
        "margin_right_cm": twips_to_cm(int_attr(margin, "right")),
        "header_distance_cm": twips_to_cm(int_attr(margin, "header")),
        "footer_distance_cm": twips_to_cm(int_attr(margin, "footer")),
        "page_number_format": page_number_format(section),
    }


def section_info(root: ET.Element) -> list[dict[str, Any]]:
    """抽取节信息。"""
    sections = root.findall(".//w:sectPr", NS)
    if not sections:
        sections = []
    result: list[dict[str, Any]] = []
    for index, section in enumerate(sections, start=1):
        page_setup = page_setup_from_section(section)
        result.append(
            {
                "fact_id": f"section-{index:04d}-page-setup",
                "fact_kind": "page_setup",
                "section_index": index,
                "locator": {"section_index": index},
                "page_setup": page_setup,
                "page_width_twips": page_setup["page_width_twips"],
                "page_height_twips": page_setup["page_height_twips"],
                "orientation": page_setup["page_orientation"],
                "margin_top_cm": page_setup["margin_top_cm"],
                "margin_bottom_cm": page_setup["margin_bottom_cm"],
                "margin_left_cm": page_setup["margin_left_cm"],
                "margin_right_cm": page_setup["margin_right_cm"],
            }
        )
    if not result:
        result.append(
            {
                "fact_id": "section-0001-page-setup",
                "fact_kind": "page_setup",
                "section_index": 1,
                "locator": {"section_index": 1},
                "page_setup": page_setup_from_section(None),
            }
        )
    return result


def extract_snapshot(docx_path: Path, snapshot_kind: str, with_source: bool = True) -> dict[str, Any]:
    """提取 DOCX 快照。

    Raises:
        FileNotFoundError: docx_path 不存在
        zipfile.BadZipFile: 文件不是合法 ZIP/OOXML 格式
        ValueError: ZIP 内缺少 word/document.xml
    """
    if not docx_path.exists():
        raise FileNotFoundError(f"DOCX 文件不存在: {docx_path}")
    try:
        digest = hashlib.sha256(docx_path.read_bytes()).hexdigest()
    except PermissionError as exc:
        raise PermissionError(f"无法读取 DOCX 文件: {docx_path}") from exc
    try:
        with zipfile.ZipFile(docx_path, "r") as archive:
            namelist = archive.namelist()
            if "word/document.xml" not in namelist:
                raise ValueError(f"DOCX 缺少 word/document.xml: {docx_path}")
            document_xml = archive.read("word/document.xml")
            styles_xml = archive.read("word/styles.xml") if "word/styles.xml" in namelist else None
            archive.testzip()
    except zipfile.BadZipFile:
        raise zipfile.BadZipFile(f"文件不是合法 ZIP/OOXML 格式: {docx_path}")
    root = ET.fromstring(document_xml)
    styles_root = ET.fromstring(styles_xml) if styles_xml else None
    resolver = StyleResolver(styles_root)
    paragraphs = root.findall(".//w:p", NS)
    paragraph_index_by_id = {id(paragraph): index for index, paragraph in enumerate(paragraphs, start=1)}
    tables = root.findall(".//w:tbl", NS)
    paragraph_items = []
    for index, paragraph in enumerate(paragraphs, start=1):
        text = text_of(paragraph).strip()
        style_id = paragraph_style(paragraph)
        paragraph_items.append(
            {
                "element_id": f"p-{index:05d}",
                "paragraph_index": index,
                "text_preview": text[:120],
                "style_id": style_id,
                "style_name": resolver.style_name(style_id),
                "paragraph_format": paragraph_format(paragraph),
                "resolved_paragraph_format": (
                    resolver.resolved_paragraph_format_with_source(paragraph)
                    if with_source
                    else resolver.resolved_paragraph_format(paragraph)
                ),
                "run_format": run_format(paragraph),
                "resolved_run_format": (
                    resolver.resolved_run_format_with_source(paragraph)
                    if with_source
                    else resolver.resolved_run_format(paragraph)
                ),
            }
        )
    table_items = [
        table_info(table, index, paragraph_index_by_id, resolver, with_source=with_source)
        for index, table in enumerate(tables, start=1)
    ]
    sections = section_info(root)
    return {
        "schema_version": "1.0.0",
        "snapshot_kind": snapshot_kind,
        "source_docx": str(docx_path).replace("\\", "/"),
        "document_hash": f"sha256:{digest}",
        "created_at": datetime.now(TZ).isoformat(),
        "paragraph_count": len(paragraph_items),
        "table_count": len(table_items),
        "section_count": len(sections),
        "style_count": len(resolver.styles),
        "resolved_format_with_source": with_source,
        "paragraphs": paragraph_items,
        "tables": table_items,
        "sections": sections,
    }


def main() -> int:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="提取 DOCX 事实快照")
    parser.add_argument("docx_path", type=Path)
    parser.add_argument("--snapshot-kind", required=True, choices=["standard", "before", "after"])
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--without-source", action="store_true", help="输出旧版裸 resolved_* 格式")
    args = parser.parse_args()
    snapshot = extract_snapshot(args.docx_path, args.snapshot_kind, with_source=not args.without_source)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    content = json.dumps(snapshot, ensure_ascii=False, indent=2) + "\n"
    temp_path = args.output.with_suffix(args.output.suffix + ".tmp")
    temp_path.write_text(content, encoding="utf-8")
    os.replace(temp_path, args.output)
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
