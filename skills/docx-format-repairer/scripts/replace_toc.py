#!/usr/bin/env python3
"""生成自动目录替换动作说明。"""
from __future__ import annotations

import argparse
import json
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--levels", default="1-3")
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    Path(args.output).write_text(json.dumps({"action": "insert_or_replace_toc_field", "require_toc_content_audit": True, "levels": args.levels}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
