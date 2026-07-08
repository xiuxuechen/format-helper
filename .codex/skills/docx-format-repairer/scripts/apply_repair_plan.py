#!/usr/bin/env python3
"""退役写回入口。

该文件仅保留历史命令路径的明确阻断信息。生产写回必须使用
`scripts/officecli/request_builder.py` 和 `scripts/officecli/runtime_adapter.py`。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


ERROR_CODE = "FH-OFFICECLI-LEGACY-BACKEND-RETIRED"
MESSAGE = "旧 DOCX 写回后端已退役，请使用 OfficeCLI runtime_adapter 执行写回。"


def retired_result() -> dict[str, object]:
    """返回稳定的退役阻断结果。"""
    return {
        "ok": False,
        "status": "blocked",
        "error": {
            "code": ERROR_CODE,
            "message": MESSAGE,
            "retryable": False,
        },
    }


def main_from_args(argv: list[str] | None = None) -> int:
    """兼容旧命令参数，但始终阻断执行。"""
    parser = argparse.ArgumentParser(description="retired DOCX repair backend")
    parser.add_argument("--repair-plan", type=Path)
    parser.add_argument("--log", type=Path)
    args = parser.parse_args(argv)
    result = retired_result()
    if args.log:
        args.log.parent.mkdir(parents=True, exist_ok=True)
        args.log.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    sys.stdout.write(json.dumps(result, ensure_ascii=False) + "\n")
    return 2


def main() -> int:
    """命令行入口。"""
    return main_from_args()


if __name__ == "__main__":
    raise SystemExit(main())
