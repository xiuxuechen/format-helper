#!/usr/bin/env python3
"""创建 DOCX 结构快照。"""
from __future__ import annotations

import runpy
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1].parent / "docx-rule-extractor" / "scripts" / "inspect_docx_profile.py"

if __name__ == "__main__":
    runpy.run_path(str(SCRIPT), run_name="__main__")
