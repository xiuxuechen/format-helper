#!/usr/bin/env python3
"""生成 DOCX 结构画像。"""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
W = f"{{{NS['w']}}}"


def xml_from_docx(docx: Path, name: str):
    with zipfile.ZipFile(docx) as zf:
        try:
            return ET.fromstring(zf.read(name))
        except KeyError:
            return None


def attr(node, name: str):
    return node.get(f"{W}{name}") if node is not None else None


def twips_to_cm(value: str | None):
    if value is None:
        return None
    try:
        return round(int(value) / 567, 2)
    except ValueError:
        return None


def half_points(value: str | None):
    if value is None:
        return None
    try:
        return round(int(value) / 2, 1)
    except ValueError:
        return None


def point_spacing(value: str | None):
    if value is None:
        return None
    try:
        return round(int(value) / 20, 1)
    except ValueError:
        return None


def line_spacing_value(value: str | None, rule: str | None) -> dict:
    if value is None:
        return {
            "line_spacing_raw": None,
            "line_spacing_rule": rule,
            "line_spacing_multiple": None,
            "line_spacing_pt": None,
        }
    try:
        numeric = int(value)
    except ValueError:
        return {
            "line_spacing_raw": value,
            "line_spacing_rule": rule,
            "line_spacing_multiple": None,
            "line_spacing_pt": None,
        }
    if rule == "auto":
        return {
            "line_spacing_raw": value,
            "line_spacing_rule": rule,
            "line_spacing_multiple": round(numeric / 240, 2),
            "line_spacing_pt": None,
        }
    return {
        "line_spacing_raw": value,
        "line_spacing_rule": rule,
        "line_spacing_multiple": None,
        "line_spacing_pt": round(numeric / 20, 1),
    }


def bool_prop(node) -> bool | None:
    if node is None:
        return None
    value = attr(node, "val")
    if value is None:
        return True
    return value not in {"0", "false", "False"}


def text_of(paragraph) -> str:
    return "".join(t.text or "" for t in paragraph.findall(".//w:t", NS)).strip()


def style_of(paragraph) -> str:
    node = paragraph.find("./w:pPr/w:pStyle", NS)
    return node.get(f"{{{NS['w']}}}val") if node is not None else "Normal"


def run_format_from(node) -> dict:
    rpr = node.find("./w:rPr", NS) if node is not None else None
    fonts = rpr.find("./w:rFonts", NS) if rpr is not None else None
    size = rpr.find("./w:sz", NS) if rpr is not None else None
    underline = rpr.find("./w:u", NS) if rpr is not None else None
    color = rpr.find("./w:color", NS) if rpr is not None else None
    highlight = rpr.find("./w:highlight", NS) if rpr is not None else None
    vert_align = rpr.find("./w:vertAlign", NS) if rpr is not None else None
    return {
        "font_ascii": attr(fonts, "ascii"),
        "font_east_asia": attr(fonts, "eastAsia"),
        "font_size_pt": half_points(attr(size, "val")),
        "bold": bool_prop(rpr.find("./w:b", NS)) if rpr is not None else None,
        "italic": bool_prop(rpr.find("./w:i", NS)) if rpr is not None else None,
        "underline": attr(underline, "val"),
        "color": attr(color, "val"),
        "highlight": attr(highlight, "val"),
        "strike": bool_prop(rpr.find("./w:strike", NS)) if rpr is not None else None,
        "small_caps": bool_prop(rpr.find("./w:smallCaps", NS)) if rpr is not None else None,
        "all_caps": bool_prop(rpr.find("./w:caps", NS)) if rpr is not None else None,
        "vertical_align": attr(vert_align, "val"),
    }


def paragraph_format_from(paragraph) -> dict:
    ppr = paragraph.find("./w:pPr", NS)
    ind = ppr.find("./w:ind", NS) if ppr is not None else None
    spacing = ppr.find("./w:spacing", NS) if ppr is not None else None
    jc = ppr.find("./w:jc", NS) if ppr is not None else None
    outline = ppr.find("./w:outlineLvl", NS) if ppr is not None else None
    tabs = ppr.findall("./w:tabs/w:tab", NS) if ppr is not None else []
    paragraph_format = {
        "alignment": attr(jc, "val"),
        "first_line_indent_cm": twips_to_cm(attr(ind, "firstLine")),
        "left_indent_cm": twips_to_cm(attr(ind, "left")),
        "right_indent_cm": twips_to_cm(attr(ind, "right")),
        "hanging_indent_cm": twips_to_cm(attr(ind, "hanging")),
        "space_before_pt": point_spacing(attr(spacing, "before")),
        "space_after_pt": point_spacing(attr(spacing, "after")),
        "keep_next": bool_prop(ppr.find("./w:keepNext", NS)) if ppr is not None else None,
        "keep_lines": bool_prop(ppr.find("./w:keepLines", NS)) if ppr is not None else None,
        "page_break_before": bool_prop(ppr.find("./w:pageBreakBefore", NS)) if ppr is not None else None,
        "widow_control": bool_prop(ppr.find("./w:widowControl", NS)) if ppr is not None else None,
        "outline_level": attr(outline, "val"),
        "tab_count": len(tabs),
    }
    paragraph_format.update(line_spacing_value(attr(spacing, "line"), attr(spacing, "lineRule")))
    return paragraph_format


def first_run_format(paragraph) -> dict:
    run = paragraph.find("./w:r", NS)
    return run_format_from(run)


def run_format_summary(paragraph) -> dict:
    runs = paragraph.findall("./w:r", NS)
    formats = [run_format_from(run) for run in runs]
    if not formats:
        return run_format_from(None)

    summary = {}
    for key in formats[0]:
        values = [item.get(key) for item in formats]
        non_null_values = [value for value in values if value is not None]
        unique_non_null_values = []
        for value in non_null_values:
            if value not in unique_non_null_values:
                unique_non_null_values.append(value)
        if len(unique_non_null_values) == 1 and len(non_null_values) == len(values):
            summary[key] = unique_non_null_values[0]
        elif unique_non_null_values:
            summary[key] = unique_non_null_values[0]
        else:
            summary[key] = None
    return summary


def run_format_variants(paragraph) -> list[dict]:
    variants = []
    seen = set()
    for run in paragraph.findall("./w:r", NS):
        run_format = run_format_from(run)
        key = tuple(sorted(run_format.items()))
        if key in seen:
            continue
        seen.add(key)
        variants.append(run_format)
    return variants


def has_mixed_run_format(paragraph) -> bool:
    return len(run_format_variants(paragraph)) > 1


DIRECT_PARAGRAPH_FORMAT_TAGS = {
    "adjustRightInd",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "cnfStyle",
    "contextualSpacing",
    "divId",
    "framePr",
    "ind",
    "jc",
    "keepLines",
    "keepNext",
    "kinsoku",
    "mirrorIndents",
    "numPr",
    "outlineLvl",
    "overflowPunct",
    "pageBreakBefore",
    "pBdr",
    "sectPr",
    "shd",
    "snapToGrid",
    "spacing",
    "suppressAutoHyphens",
    "suppressLineNumbers",
    "suppressOverlap",
    "tabs",
    "textAlignment",
    "textboxTightWrap",
    "textDirection",
    "topLinePunct",
    "widowControl",
    "wordWrap",
}


DIRECT_RUN_FORMAT_TAGS = {
    "b",
    "bCs",
    "bdr",
    "caps",
    "color",
    "cs",
    "dstrike",
    "eastAsianLayout",
    "effect",
    "em",
    "emboss",
    "fitText",
    "highlight",
    "i",
    "iCs",
    "imprint",
    "kern",
    "lang",
    "noProof",
    "oMath",
    "outline",
    "position",
    "rFonts",
    "rtl",
    "shadow",
    "shd",
    "smallCaps",
    "snapToGrid",
    "spacing",
    "specVanish",
    "strike",
    "sz",
    "szCs",
    "u",
    "vanish",
    "vertAlign",
    "w",
    "webHidden",
}


def local_tag(node) -> str:
    return node.tag.rsplit("}", 1)[-1]


def direct_format_tags(paragraph) -> list[str]:
    tags = []
    ppr = paragraph.find("./w:pPr", NS)
    if ppr is not None:
        for child in list(ppr):
            tag = local_tag(child)
            if tag in DIRECT_PARAGRAPH_FORMAT_TAGS:
                tags.append(f"pPr/{tag}")

    for run in paragraph.findall("./w:r", NS):
        rpr = run.find("./w:rPr", NS)
        if rpr is None:
            continue
        for child in list(rpr):
            tag = local_tag(child)
            if tag in DIRECT_RUN_FORMAT_TAGS:
                tags.append(f"rPr/{tag}")

    return sorted(set(tags))


def has_direct_format(paragraph) -> bool:
    return bool(direct_format_tags(paragraph))


def parse_styles(docx: Path) -> dict:
    styles = xml_from_docx(docx, "word/styles.xml")
    result = {}
    if styles is None:
        return result
    for style in styles.findall(".//w:style", NS):
        style_id = attr(style, "styleId")
        if not style_id:
            continue
        name = style.find("./w:name", NS)
        ppr = style.find("./w:pPr", NS)
        result[style_id] = {
            "name": attr(name, "val"),
            "run_format": run_format_from(style),
            "paragraph_format": paragraph_format_from(style),
        }
        if ppr is None:
            result[style_id]["paragraph_format"] = {}
    return result


def parse_sections(document) -> list[dict]:
    sections = []
    for section in document.findall(".//w:sectPr", NS):
        page_size = section.find("./w:pgSz", NS)
        margins = section.find("./w:pgMar", NS)
        cols = section.find("./w:cols", NS)
        doc_grid = section.find("./w:docGrid", NS)
        text_direction = section.find("./w:textDirection", NS)
        v_align = section.find("./w:vAlign", NS)
        header_refs = section.findall("./w:headerReference", NS)
        footer_refs = section.findall("./w:footerReference", NS)
        sections.append(
            {
                "width_cm": twips_to_cm(attr(page_size, "w")),
                "height_cm": twips_to_cm(attr(page_size, "h")),
                "orientation": attr(page_size, "orient") or "portrait",
                "margin_top_cm": twips_to_cm(attr(margins, "top")),
                "margin_bottom_cm": twips_to_cm(attr(margins, "bottom")),
                "margin_left_cm": twips_to_cm(attr(margins, "left")),
                "margin_right_cm": twips_to_cm(attr(margins, "right")),
                "header_cm": twips_to_cm(attr(margins, "header")),
                "footer_cm": twips_to_cm(attr(margins, "footer")),
                "gutter_cm": twips_to_cm(attr(margins, "gutter")),
                "column_count": attr(cols, "num"),
                "column_space_cm": twips_to_cm(attr(cols, "space")),
                "text_direction": attr(text_direction, "val"),
                "vertical_alignment": attr(v_align, "val"),
                "doc_grid_type": attr(doc_grid, "type"),
                "header_reference_types": [attr(item, "type") for item in header_refs],
                "footer_reference_types": [attr(item, "type") for item in footer_refs],
                "different_first_page": section.find("./w:titlePg", NS) is not None,
            }
        )
    return sections


def parse_settings(docx: Path) -> dict:
    settings = xml_from_docx(docx, "word/settings.xml")
    if settings is None:
        return {}
    return {
        "even_and_odd_headers": settings.find(".//w:evenAndOddHeaders", NS) is not None,
        "track_revisions": settings.find(".//w:trackRevisions", NS) is not None,
        "update_fields": settings.find(".//w:updateFields", NS) is not None,
        "document_protection": settings.find(".//w:documentProtection", NS) is not None,
        "default_tab_stop_twips": attr(settings.find(".//w:defaultTabStop", NS), "val"),
    }


def parse_numbering(docx: Path) -> dict:
    numbering = xml_from_docx(docx, "word/numbering.xml")
    if numbering is None:
        return {"abstract_numbering_count": 0, "numbering_instance_count": 0, "levels_count": 0}
    return {
        "abstract_numbering_count": len(numbering.findall(".//w:abstractNum", NS)),
        "numbering_instance_count": len(numbering.findall(".//w:num", NS)),
        "levels_count": len(numbering.findall(".//w:lvl", NS)),
    }


def border_summary(borders) -> dict:
    if borders is None:
        return {"has_borders": False, "border_types": [], "border_values": {}}
    return {
        "has_borders": True,
        "border_types": [item.tag.rsplit("}", 1)[-1] for item in list(borders)],
        "border_values": {
            item.tag.rsplit("}", 1)[-1]: {
                "val": attr(item, "val"),
                "sz": attr(item, "sz"),
                "color": attr(item, "color"),
                "space": attr(item, "space"),
            }
            for item in list(borders)
        },
    }


def table_cell_margin_summary(tbl_pr) -> dict:
    margins = tbl_pr.find("./w:tblCellMar", NS) if tbl_pr is not None else None
    result = {}
    for side in ("top", "left", "bottom", "right"):
        node = margins.find(f"./w:{side}", NS) if margins is not None else None
        result[f"cell_margin_{side}"] = attr(node, "w")
    return result


def parse_tables(document, paragraph_indexes: dict[int, int]) -> list[dict]:
    result = []
    for index, table in enumerate(document.findall(".//w:tbl", NS), start=1):
        rows = table.findall("./w:tr", NS)
        grid_cols = table.findall("./w:tblGrid/w:gridCol", NS)
        tbl_pr = table.find("./w:tblPr", NS)
        tbl_width = tbl_pr.find("./w:tblW", NS) if tbl_pr is not None else None
        tbl_jc = tbl_pr.find("./w:jc", NS) if tbl_pr is not None else None
        tbl_layout = tbl_pr.find("./w:tblLayout", NS) if tbl_pr is not None else None
        borders = tbl_pr.find("./w:tblBorders", NS) if tbl_pr is not None else None
        cells = table.findall(".//w:tc", NS)
        merged_cells = 0
        vertically_centered_cells = 0
        cell_paragraphs = []
        text_preview = ""
        for row_index, row in enumerate(rows, start=1):
            for column_index, cell in enumerate(row.findall("./w:tc", NS), start=1):
                tc_pr = cell.find("./w:tcPr", NS)
                if tc_pr is not None and (tc_pr.find("./w:gridSpan", NS) is not None or tc_pr.find("./w:vMerge", NS) is not None):
                    merged_cells += 1
                v_align = tc_pr.find("./w:vAlign", NS) if tc_pr is not None else None
                if attr(v_align, "val") == "center":
                    vertically_centered_cells += 1
                for paragraph in cell.findall("./w:p", NS):
                    text = text_of(paragraph)
                    if not text:
                        continue
                    style = style_of(paragraph)
                    classification = classify_details(text, style)
                    if not text_preview:
                        text_preview = text[:120]
                    cell_paragraphs.append(
                        {
                            "paragraph_index": paragraph_indexes.get(id(paragraph)),
                            "row_index": row_index,
                            "column_index": column_index,
                            "text_preview": text[:120],
                            "style": style,
                            "element_type": classification["element_type"],
                            "role_confidence": classification["role_confidence"],
                            "numbering_pattern": classification["numbering_pattern"],
                            "classification_reason": classification["classification_reason"],
                            "paragraph_format": paragraph_format_from(paragraph),
                            "run_format": run_format_summary(paragraph),
                            "first_run_format": first_run_format(paragraph),
                            "run_format_variants": run_format_variants(paragraph),
                            "mixed_run_format": has_mixed_run_format(paragraph),
                            "has_direct_format": has_direct_format(paragraph),
                            "direct_format_tags": direct_format_tags(paragraph),
                        }
                    )
        header_rows = sum(1 for row in rows if row.find("./w:trPr/w:tblHeader", NS) is not None)
        cant_split_rows = sum(1 for row in rows if row.find("./w:trPr/w:cantSplit", NS) is not None)
        row_height_count = sum(1 for row in rows if row.find("./w:trPr/w:trHeight", NS) is not None)
        first_row = rows[0] if rows else None
        first_row_shading_count = 0
        first_row_shading_fills = []
        if first_row is not None:
            for cell in first_row.findall("./w:tc", NS):
                shading = cell.find("./w:tcPr/w:shd", NS)
                if shading is not None and attr(shading, "fill"):
                    first_row_shading_count += 1
                    first_row_shading_fills.append(attr(shading, "fill"))
        result.append(
            {
                "element_id": f"t-{index:05d}",
                "table_index": index,
                "row_count": len(rows),
                "cell_count": len(cells),
                "grid_column_count": len(grid_cols),
                "width": attr(tbl_width, "w"),
                "width_type": attr(tbl_width, "type"),
                "alignment": attr(tbl_jc, "val"),
                "layout_type": attr(tbl_layout, "type"),
                "merged_cell_count": merged_cells,
                "has_merged_cells": merged_cells > 0,
                "header_row_count": header_rows,
                "cant_split_row_count": cant_split_rows,
                "row_height_count": row_height_count,
                "vertically_centered_cell_count": vertically_centered_cells,
                "first_row_shading_cell_count": first_row_shading_count,
                "first_row_shading_fills": first_row_shading_fills,
                "text_preview": text_preview,
                "cell_paragraphs": cell_paragraphs,
                **table_cell_margin_summary(tbl_pr),
                **border_summary(borders),
            }
        )
    return result


def package_part_summary(docx: Path) -> dict:
    with zipfile.ZipFile(docx) as zf:
        names = zf.namelist()
    return {
        "header_part_count": len([name for name in names if name.startswith("word/header") and name.endswith(".xml")]),
        "footer_part_count": len([name for name in names if name.startswith("word/footer") and name.endswith(".xml")]),
        "footnotes_present": "word/footnotes.xml" in names,
        "endnotes_present": "word/endnotes.xml" in names,
        "comments_present": "word/comments.xml" in names,
        "media_count": len([name for name in names if name.startswith("word/media/")]),
    }


CHINESE_NUMBER_CHARS = r"\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u5343\u4e07\u96f6\u3007\u4e24"


def numbering_pattern(text: str) -> str:
    if re.match(rf"^[{CHINESE_NUMBER_CHARS}]+[\u3001\uff0e\.]", text):
        return "chinese-level-1"
    if re.match(rf"^[\(\uff08][{CHINESE_NUMBER_CHARS}]+[\)\uff09]", text):
        return "chinese-level-2"
    if re.match(r"^\d+(?:\.\d+)+", text):
        return "multilevel-decimal"
    if re.match(r"^\d+[\.\uff0e\u3001\)]", text):
        return "decimal"
    return "none"


def classify_details(text: str, style: str) -> dict:
    pattern = numbering_pattern(text)
    if not text:
        return {
            "element_type": "empty",
            "role_confidence": 1.0,
            "numbering_pattern": pattern,
            "classification_reason": "paragraph has no text",
        }
    if text == "目录":
        return {
            "element_type": "toc-title",
            "role_confidence": 0.95,
            "numbering_pattern": pattern,
            "classification_reason": "text matches table-of-contents title",
        }
    if re.search(r"\.{3,}\s*\d+$", text):
        return {
            "element_type": "toc-static-item",
            "role_confidence": 0.9,
            "numbering_pattern": pattern,
            "classification_reason": "text ends with dot leaders and page number",
        }
    if style.lower().startswith("heading"):
        level = re.sub(r"\D", "", style) or "1"
        return {
            "element_type": f"heading-level-{level}",
            "role_confidence": 0.95,
            "numbering_pattern": pattern,
            "classification_reason": f"paragraph style is {style}",
        }
    if pattern == "chinese-level-1":
        return {
            "element_type": "heading-level-1",
            "role_confidence": 0.85,
            "numbering_pattern": pattern,
            "classification_reason": "text starts with Chinese level-1 numbering",
        }
    if pattern == "chinese-level-2":
        return {
            "element_type": "heading-level-2",
            "role_confidence": 0.8,
            "numbering_pattern": pattern,
            "classification_reason": "text starts with parenthesized Chinese numbering",
        }
    if pattern in {"decimal", "multilevel-decimal"}:
        return {
            "element_type": "ambiguous-numbered-item",
            "role_confidence": 0.45,
            "numbering_pattern": pattern,
            "classification_reason": "Arabic numeric numbering is ambiguous without a heading style",
        }
    return {
        "element_type": "body-paragraph",
        "role_confidence": 0.75,
        "numbering_pattern": pattern,
        "classification_reason": "no heading style or recognized heading numbering",
    }


def classify(text: str, style: str) -> str:
    return classify_details(text, style)["element_type"]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx")
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    docx = Path(args.docx)
    document = xml_from_docx(docx, "word/document.xml")
    if document is None:
        raise SystemExit("不是有效的 DOCX：缺少 word/document.xml")

    paragraphs = []
    paragraph_indexes = {}
    style_counter = Counter()
    type_counter = Counter()
    style_details = parse_styles(docx)
    for index, paragraph in enumerate(document.findall(".//w:p", NS), start=1):
        paragraph_indexes[id(paragraph)] = index
        text = text_of(paragraph)
        style = style_of(paragraph)
        classification = classify_details(text, style)
        element_type = classification["element_type"]
        if text:
            style_counter[style] += 1
            type_counter[element_type] += 1
        paragraphs.append(
            {
                "element_id": f"p-{index:05d}",
                "paragraph_index": index,
                "text_preview": text[:120],
                "style": style,
                "element_type": element_type,
                "role_confidence": classification["role_confidence"],
                "numbering_pattern": classification["numbering_pattern"],
                "classification_reason": classification["classification_reason"],
                "paragraph_format": paragraph_format_from(paragraph),
                "run_format": run_format_summary(paragraph),
                "first_run_format": first_run_format(paragraph),
                "run_format_variants": run_format_variants(paragraph),
                "mixed_run_format": has_mixed_run_format(paragraph),
                "has_direct_format": has_direct_format(paragraph),
                "direct_format_tags": direct_format_tags(paragraph),
            }
        )

    tables = document.findall(".//w:tbl", NS)
    sections = parse_sections(document)
    has_toc_field = "TOC" in ET.tostring(document, encoding="unicode")
    package_summary = package_part_summary(docx)
    settings = parse_settings(docx)
    numbering = parse_numbering(docx)
    table_details = parse_tables(document, paragraph_indexes)
    hyperlinks = document.findall(".//w:hyperlink", NS)
    field_instructions = [node.text or "" for node in document.findall(".//w:instrText", NS)]
    page_field_count = sum(1 for text in field_instructions if "PAGE" in text.upper())
    revision_count = len(document.findall(".//w:ins", NS)) + len(document.findall(".//w:del", NS))
    snapshot = {
        "schema_version": "1.0.0",
        "source_path": str(docx),
        "paragraph_count": len(paragraphs),
        "non_empty_paragraph_count": sum(1 for p in paragraphs if p["text_preview"]),
        "table_count": len(tables),
        "table_details": table_details,
        "section_count": len(sections),
        "sections": sections,
        "has_toc_field": has_toc_field,
        "field_summary": {
            "field_instruction_count": len(field_instructions),
            "page_field_count": page_field_count,
        },
        "package_summary": package_summary,
        "settings": settings,
        "numbering_summary": numbering,
        "hyperlink_count": len(hyperlinks),
        "revision_count": revision_count,
        "style_summary": dict(style_counter),
        "style_details": style_details,
        "available_style_ids": sorted(style_details),
        "element_type_summary": dict(type_counter),
        "paragraphs": paragraphs,
    }
    Path(args.output).write_text(json.dumps(snapshot, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
