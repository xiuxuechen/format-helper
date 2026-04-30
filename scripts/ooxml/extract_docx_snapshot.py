#!/usr/bin/env python3
"""从 DOCX 提取客观 OOXML 事实快照。"""

from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


TZ = timezone(timedelta(hours=8))
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = NS["w"]
FORMAT_EMPTY_VALUES = {None, ""}


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

    def resolved_paragraph_format(self, paragraph: ET.Element) -> dict[str, Any]:
        """解析段落最终有效格式。"""
        style_p, _ = self.resolve_style(paragraph_style(paragraph))
        direct_p = paragraph_format(paragraph)
        return merge_formats(self.default_paragraph_format, style_p, direct_p)

    def resolved_run_format(self, paragraph: ET.Element, run: ET.Element | None = None) -> dict[str, Any]:
        """解析 run 最终有效格式。"""
        if run is None:
            run = paragraph.find("./w:r", NS)
        _, paragraph_style_r = self.resolve_style(paragraph_style(paragraph))
        _, run_style_r = self.resolve_style(run_style(run))
        direct_r = run_format_from_rpr(run.find("./w:rPr", NS) if run is not None else None)
        return merge_formats(self.default_run_format, paragraph_style_r, run_style_r, direct_r)


def table_cell_role(row_index: int, column_index: int) -> str:
    """给单元格提供保守的角色候选。"""
    if row_index == 1:
        return "header-row-cell"
    if column_index == 1:
        return "field-name-cell"
    return "value-cell"


def run_items(paragraph: ET.Element, resolver: StyleResolver) -> list[dict[str, Any]]:
    """抽取段落下所有 run 的文本和格式。"""
    items: list[dict[str, Any]] = []
    for index, run in enumerate(paragraph.findall("./w:r", NS), start=1):
        text = text_of(run).strip()
        direct = run_format_from_rpr(run.find("./w:rPr", NS))
        resolved = resolver.resolved_run_format(paragraph, run)
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


def cell_format_summary(paragraphs: list[dict[str, Any]]) -> dict[str, Any]:
    """汇总单元格格式，用于表格正文审计。"""
    fonts: set[str] = set()
    sizes: set[float] = set()
    has_bold = False
    for paragraph in paragraphs:
        for run in paragraph.get("runs", []):
            resolved = run.get("resolved_run_format") or {}
            direct = run.get("run_format") or {}
            if resolved.get("bold") is True or direct.get("bold") is True:
                has_bold = True
            font = resolved.get("font_east_asia") or direct.get("font_east_asia")
            size = resolved.get("font_size_pt") or direct.get("font_size_pt")
            if font:
                fonts.add(str(font))
            if isinstance(size, (int, float)):
                sizes.add(float(size))
    return {
        "has_bold": has_bold,
        "font_east_asia_values": sorted(fonts),
        "font_size_pt_values": sorted(sizes),
    }


def table_info(table: ET.Element, index: int, paragraph_index_by_id: dict[int, int], resolver: StyleResolver) -> dict[str, Any]:
    """抽取表格基础事实。"""
    rows = table.findall("./w:tr", NS)
    column_count = 0
    cells: list[dict[str, Any]] = []
    for row_index, row in enumerate(rows, start=1):
        row_cells = row.findall("./w:tc", NS)
        column_count = max(column_count, len(row_cells))
        for column_index, cell in enumerate(row_cells, start=1):
            paragraph_items: list[dict[str, Any]] = []
            for paragraph in cell.findall(".//w:p", NS):
                text = text_of(paragraph).strip()
                paragraph_items.append(
                    {
                        "paragraph_index": paragraph_index_by_id.get(id(paragraph)),
                        "text_preview": text[:120],
                        "style_id": paragraph_style(paragraph),
                        "style_name": resolver.style_name(paragraph_style(paragraph)),
                        "paragraph_format": paragraph_format(paragraph),
                        "resolved_paragraph_format": resolver.resolved_paragraph_format(paragraph),
                        "run_format": run_format(paragraph),
                        "resolved_run_format": resolver.resolved_run_format(paragraph),
                        "runs": run_items(paragraph, resolver),
                    }
                )
            cells.append(
                {
                    "cell_id": f"table-{index:04d}-r{row_index:03d}-c{column_index:03d}",
                    "row_index": row_index,
                    "column_index": column_index,
                    "cell_role_candidate": table_cell_role(row_index, column_index),
                    "text_preview": text_of(cell).strip()[:120],
                    "paragraphs": paragraph_items,
                    "format_summary": cell_format_summary(paragraph_items),
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


def section_info(root: ET.Element) -> list[dict[str, Any]]:
    """抽取节信息。"""
    sections = root.findall(".//w:sectPr", NS)
    if not sections:
        sections = []
    result: list[dict[str, Any]] = []
    for index, section in enumerate(sections, start=1):
        page_size = section.find("./w:pgSz", NS)
        margin = section.find("./w:pgMar", NS)
        result.append(
            {
                "section_index": index,
                "page_width_twips": int_attr(page_size, "w"),
                "page_height_twips": int_attr(page_size, "h"),
                "orientation": w_attr(page_size, "orient") or "portrait",
                "margin_top_cm": twips_to_cm(int_attr(margin, "top")),
                "margin_bottom_cm": twips_to_cm(int_attr(margin, "bottom")),
                "margin_left_cm": twips_to_cm(int_attr(margin, "left")),
                "margin_right_cm": twips_to_cm(int_attr(margin, "right")),
            }
        )
    if not result:
        result.append({"section_index": 1})
    return result


def extract_snapshot(docx_path: Path, snapshot_kind: str) -> dict[str, Any]:
    """提取 DOCX 快照。"""
    digest = hashlib.sha256(docx_path.read_bytes()).hexdigest()
    with zipfile.ZipFile(docx_path, "r") as archive:
        document_xml = archive.read("word/document.xml")
        styles_xml = archive.read("word/styles.xml") if "word/styles.xml" in archive.namelist() else None
        archive.testzip()
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
                "resolved_paragraph_format": resolver.resolved_paragraph_format(paragraph),
                "run_format": run_format(paragraph),
                "resolved_run_format": resolver.resolved_run_format(paragraph),
            }
        )
    table_items = [table_info(table, index, paragraph_index_by_id, resolver) for index, table in enumerate(tables, start=1)]
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
    args = parser.parse_args()
    snapshot = extract_snapshot(args.docx_path, args.snapshot_kind)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(snapshot, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
