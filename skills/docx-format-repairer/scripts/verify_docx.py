#!/usr/bin/env python3
"""校验 DOCX 是否可作为 ZIP/OOXML 打开。"""
from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx")
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    docx = Path(args.docx)
    result = {"docx": str(docx), "readable": False, "has_document_xml": False, "error": None}
    try:
        with zipfile.ZipFile(docx) as zf:
            names = set(zf.namelist())
            result["readable"] = True
            result["has_document_xml"] = "word/document.xml" in names
    except Exception as exc:
        result["error"] = str(exc)
    Path(args.output).write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(args.output)
    return 0 if result["readable"] and result["has_document_xml"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
