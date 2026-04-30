#!/usr/bin/env python3
"""优化表格与分页布局，减少孤立标题、孤立标点和空白页。"""

from __future__ import annotations

import argparse
import json
import shutil
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


TZ = timezone(timedelta(hours=8))
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
W = NS["w"]
ET.register_namespace("w", W)
TITLE_TEXTS = {"四、会议议程", "预算表"}
BUDGET_SECTION_TITLE = "建议预算及执行口径"
BUDGET_TABLE_TITLE = "预算表"
COMPACT_LINE_SPACING = "440"
RISK_CONTROL_PREFIX = "五、风险控制："
RISK_CONTROL_CHAR_SPACING = "-8"
RISK_CONTROL_FONT_SIZE = "31"


def qn(name: str) -> str:
    """返回 w 命名空间标签。"""
    return f"{{{W}}}{name}"


def set_attr(node: ET.Element, name: str, value: Any) -> None:
    """写入 w:* 属性。"""
    node.set(qn(name), str(value))


def get_or_create(parent: ET.Element, name: str, prepend: bool = False) -> ET.Element:
    """获取或创建子节点。"""
    node = parent.find(f"./w:{name}", NS)
    if node is None:
        node = ET.SubElement(parent, qn(name))
        if prepend and len(parent) > 1:
            parent.remove(node)
            parent.insert(0, node)
    return node


def paragraph_text(paragraph: ET.Element) -> str:
    """提取段落可见文本。"""
    return "".join(text.text or "" for text in paragraph.findall(".//w:t", NS)).strip()


def node_text(node: ET.Element) -> str:
    """提取节点内全部可见文本。"""
    return "".join(text.text or "" for text in node.findall(".//w:t", NS)).strip()


def is_empty_paragraph(paragraph: ET.Element) -> bool:
    """判断段落是否没有可见内容或复杂对象。"""
    if paragraph.tag != qn("p") or paragraph_text(paragraph):
        return False
    visible_nodes = (
        ".//w:drawing",
        ".//w:pict",
        ".//w:object",
        ".//w:fldChar",
        ".//w:instrText",
        ".//w:br",
    )
    return not any(paragraph.findall(pattern, NS) for pattern in visible_nodes)


def table_column_count(table: ET.Element) -> int:
    """推断表格列数。"""
    grid_cols = table.findall("./w:tblGrid/w:gridCol", NS)
    if grid_cols:
        return len(grid_cols)
    first_row = table.find("./w:tr", NS)
    if first_row is None:
        return 1
    return max(1, len(first_row.findall("./w:tc", NS)))


def table_width(table: ET.Element) -> str:
    """汇总表格网格宽度。"""
    widths: list[int] = []
    for col in table.findall("./w:tblGrid/w:gridCol", NS):
        value = col.get(qn("w"))
        if value and value.isdigit():
            widths.append(int(value))
    if widths:
        return str(sum(widths))
    return "0"


def make_caption_row(title: str, column_count: int, width: str) -> ET.Element:
    """创建跨列标题行。"""
    row = ET.Element(qn("tr"))
    row_pr = ET.SubElement(row, qn("trPr"))
    cant_split = ET.SubElement(row_pr, qn("cantSplit"))
    set_attr(cant_split, "val", "1")

    cell = ET.SubElement(row, qn("tc"))
    cell_pr = ET.SubElement(cell, qn("tcPr"))
    cell_width = ET.SubElement(cell_pr, qn("tcW"))
    set_attr(cell_width, "w", width)
    set_attr(cell_width, "type", "dxa" if width != "0" else "auto")
    if column_count > 1:
        grid_span = ET.SubElement(cell_pr, qn("gridSpan"))
        set_attr(grid_span, "val", column_count)
    v_align = ET.SubElement(cell_pr, qn("vAlign"))
    set_attr(v_align, "val", "center")

    paragraph = ET.SubElement(cell, qn("p"))
    p_pr = ET.SubElement(paragraph, qn("pPr"))
    spacing = ET.SubElement(p_pr, qn("spacing"))
    set_attr(spacing, "before", "0")
    set_attr(spacing, "after", "0")
    set_attr(spacing, "line", "560")
    set_attr(spacing, "lineRule", "exact")
    jc = ET.SubElement(p_pr, qn("jc"))
    set_attr(jc, "val", "center")
    run = ET.SubElement(paragraph, qn("r"))
    r_pr = ET.SubElement(run, qn("rPr"))
    fonts = ET.SubElement(r_pr, qn("rFonts"))
    set_attr(fonts, "ascii", "Times New Roman")
    set_attr(fonts, "hAnsi", "Times New Roman")
    set_attr(fonts, "eastAsia", "黑体")
    size = ET.SubElement(r_pr, qn("sz"))
    set_attr(size, "val", "32")
    size_cs = ET.SubElement(r_pr, qn("szCs"))
    set_attr(size_cs, "val", "32")
    text = ET.SubElement(run, qn("t"))
    text.text = title
    return row


def mark_repeat_header(row: ET.Element) -> None:
    """将表格原表头行标记为重复表头。"""
    row_pr = get_or_create(row, "trPr", prepend=True)
    header = get_or_create(row_pr, "tblHeader")
    set_attr(header, "val", "1")


def compact_paragraph_line_spacing(paragraph: ET.Element) -> bool:
    """压缩指定段落行距，用于消除孤立标点页。"""
    p_pr = get_or_create(paragraph, "pPr", prepend=True)
    spacing = get_or_create(p_pr, "spacing")
    old_line = spacing.get(qn("line"))
    old_rule = spacing.get(qn("lineRule"))
    if old_line == COMPACT_LINE_SPACING and old_rule == "exact":
        return False
    set_attr(spacing, "line", COMPACT_LINE_SPACING)
    set_attr(spacing, "lineRule", "exact")
    if spacing.get(qn("before")) is None:
        set_attr(spacing, "before", "0")
    if spacing.get(qn("after")) is None:
        set_attr(spacing, "after", "0")
    return True


def condense_run_spacing(paragraph: ET.Element) -> int:
    """轻微压缩段落字符间距和字号，避免句号单独成行。"""
    changed = 0
    for run in paragraph.findall("./w:r", NS):
        r_pr = get_or_create(run, "rPr", prepend=True)
        spacing = get_or_create(r_pr, "spacing")
        if spacing.get(qn("val")) == RISK_CONTROL_CHAR_SPACING:
            spacing_changed = False
        else:
            set_attr(spacing, "val", RISK_CONTROL_CHAR_SPACING)
            spacing_changed = True
        size = get_or_create(r_pr, "sz")
        size_cs = get_or_create(r_pr, "szCs")
        size_changed = size.get(qn("val")) != RISK_CONTROL_FONT_SIZE
        size_cs_changed = size_cs.get(qn("val")) != RISK_CONTROL_FONT_SIZE
        set_attr(size, "val", RISK_CONTROL_FONT_SIZE)
        set_attr(size_cs, "val", RISK_CONTROL_FONT_SIZE)
        if spacing_changed or size_changed or size_cs_changed:
            changed += 1
    return changed


def find_table_title(table: ET.Element) -> str:
    """提取表格第一行标题。"""
    first_row = table.find("./w:tr", NS)
    if first_row is None:
        return ""
    return node_text(first_row)


def bind_table_titles(body: ET.Element) -> list[dict[str, Any]]:
    """将表格前标题并入表格首行。"""
    actions: list[dict[str, Any]] = []
    children = list(body)
    index = 0
    while index < len(children) - 1:
        current = children[index]
        next_node = children[index + 1]
        if current.tag != qn("p") or next_node.tag != qn("tbl"):
            index += 1
            continue
        title = paragraph_text(current)
        if title not in TITLE_TEXTS:
            index += 1
            continue

        if find_table_title(next_node) == title:
            actions.append(
                {
                    "type": "table-title-binding",
                    "title": title,
                    "status": "skipped",
                    "reason": "表格首行已包含标题",
                }
            )
            index += 1
            continue

        caption_row = make_caption_row(title, table_column_count(next_node), table_width(next_node))
        next_node.insert(0, caption_row)
        original_header = next_node.findall("./w:tr", NS)
        if len(original_header) > 1:
            mark_repeat_header(original_header[1])
        body.remove(current)
        actions.append(
            {
                "type": "table-title-binding",
                "title": title,
                "status": "executed",
                "reason": "已将表格前标题并入表格首行，避免标题孤立分页",
            }
        )
        children = list(body)
    return actions


def compact_budget_intro(body: ET.Element) -> list[dict[str, Any]]:
    """压缩预算表前说明段落，避免末尾标点独占一页。"""
    children = list(body)
    start_index: int | None = None
    table_index: int | None = None
    for index, child in enumerate(children):
        if child.tag == qn("p") and paragraph_text(child) == BUDGET_SECTION_TITLE:
            start_index = index
            continue
        if start_index is not None and child.tag == qn("tbl") and find_table_title(child) == BUDGET_TABLE_TITLE:
            table_index = index
            break
    if start_index is None or table_index is None or table_index <= start_index + 1:
        return [
            {
                "type": "budget-intro-pagination",
                "status": "skipped",
                "reason": "未找到预算说明段落与预算表的相邻范围",
            }
        ]

    changed = 0
    targets = 0
    for child in children[start_index + 1 : table_index]:
        if child.tag != qn("p") or not paragraph_text(child):
            continue
        targets += 1
        if compact_paragraph_line_spacing(child):
            changed += 1
    return [
        {
            "type": "budget-intro-pagination",
            "status": "executed" if changed else "skipped",
            "targets": targets,
            "changed": changed,
            "line_spacing": COMPACT_LINE_SPACING,
            "reason": "已压缩预算表前说明段落行距，消除孤立标点页",
        }
    ]


def condense_risk_control_paragraph(body: ET.Element) -> list[dict[str, Any]]:
    """压缩风险控制段落字符间距，消除单独成行的末尾标点。"""
    for child in list(body):
        if child.tag != qn("p"):
            continue
        text = paragraph_text(child)
        if not text.startswith(RISK_CONTROL_PREFIX):
            continue
        changed = condense_run_spacing(child)
        return [
            {
                "type": "risk-control-punctuation",
                "status": "executed" if changed else "skipped",
                "changed_runs": changed,
                "char_spacing": RISK_CONTROL_CHAR_SPACING,
                "font_size_half_points": RISK_CONTROL_FONT_SIZE,
                "reason": "已轻微压缩风险控制段落字符间距和字号，避免末尾句号单独成行",
            }
        ]
    return [
        {
            "type": "risk-control-punctuation",
            "status": "skipped",
            "reason": "未找到风险控制段落",
        }
    ]


def remove_trailing_empty_paragraphs(body: ET.Element) -> list[dict[str, Any]]:
    """删除节属性前的文末纯空段落。"""
    removed = 0
    while True:
        children = list(body)
        if len(children) < 2:
            break
        last_content = children[-2] if children[-1].tag == qn("sectPr") else children[-1]
        if not is_empty_paragraph(last_content):
            break
        body.remove(last_content)
        removed += 1
    return [
        {
            "type": "trailing-empty-paragraph",
            "status": "executed" if removed else "skipped",
            "removed": removed,
            "reason": "已删除文末纯空段落，避免生成空白尾页",
        }
    ]


def optimize_document_xml(xml: bytes) -> tuple[bytes, list[dict[str, Any]]]:
    """优化 document.xml 并返回动作日志。"""
    root = ET.fromstring(xml)
    body = root.find("./w:body", NS)
    if body is None:
        raise ValueError("缺少 w:body")
    actions: list[dict[str, Any]] = []
    actions.extend(bind_table_titles(body))
    actions.extend(compact_budget_intro(body))
    actions.extend(condense_risk_control_paragraph(body))
    actions.extend(remove_trailing_empty_paragraphs(body))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), actions


def valid_docx(path: Path) -> bool:
    """检查 DOCX 基本结构。"""
    try:
        with zipfile.ZipFile(path, "r") as archive:
            archive.getinfo("word/document.xml")
            return archive.testzip() is None
    except (KeyError, zipfile.BadZipFile):
        return False


def optimize_docx(input_docx: Path, output_docx: Path, backup_docx: Path | None, log_path: Path) -> dict[str, Any]:
    """执行分页优化。"""
    source_docx = input_docx
    if input_docx.resolve() == output_docx.resolve():
        if backup_docx is None:
            raise SystemExit("覆盖输出时必须提供 --backup")
        backup_docx.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(input_docx, backup_docx)
        source_docx = backup_docx

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source_docx, "r") as archive:
        entries = {item.filename: archive.read(item.filename) for item in archive.infolist() if not item.is_dir()}
    entries["word/document.xml"], actions = optimize_document_xml(entries["word/document.xml"])

    temp_output = output_docx.with_suffix(".tmp.docx")
    with zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in entries.items():
            archive.writestr(name, content)
    temp_output.replace(output_docx)

    result = {
        "optimized_at": datetime.now(TZ).isoformat(),
        "input_docx": str(input_docx),
        "backup_docx": str(backup_docx) if backup_docx else None,
        "output_docx": str(output_docx),
        "actions_total": len(actions),
        "actions_executed": sum(1 for item in actions if item["status"] == "executed"),
        "output_docx_valid": valid_docx(output_docx),
        "actions": actions,
    }
    log_path.parent.mkdir(parents=True, exist_ok=True)
    log_path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return result


def main() -> int:
    """命令行入口。"""
    parser = argparse.ArgumentParser(description="优化表格标题分页")
    parser.add_argument("--input-docx", required=True, type=Path)
    parser.add_argument("--output-docx", required=True, type=Path)
    parser.add_argument("--backup", type=Path)
    parser.add_argument("--log", required=True, type=Path)
    args = parser.parse_args()
    result = optimize_docx(args.input_docx, args.output_docx, args.backup, args.log)
    if not result["output_docx_valid"]:
        print("输出 DOCX OOXML 校验失败")
        return 1
    print(args.output_docx)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
