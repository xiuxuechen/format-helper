#!/usr/bin/env python3
"""合并专项审计 JSON。"""
from __future__ import annotations

import argparse
import json
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--audit-dir", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    audit_dir = Path(args.audit_dir)
    merged = {
        "schema_version": "1.0.0",
        "phase": "merge",
        "summary": {"files": 0, "issues_found": 0, "auto_fixable": 0, "manual_review": 0, "blocked": 0},
        "issues": [],
    }

    for path in sorted(audit_dir.glob("*.audit.json")):
        data = json.loads(path.read_text(encoding="utf-8"))
        merged["summary"]["files"] += 1
        for issue in data.get("issues", []):
            merged["issues"].append(issue)
            merged["summary"]["issues_found"] += 1
            policy = issue.get("recommended_action", {}).get("auto_fix_policy")
            if policy == "auto-fix":
                merged["summary"]["auto_fixable"] += 1
            elif policy == "blocked":
                merged["summary"]["blocked"] += 1
            else:
                merged["summary"]["manual_review"] += 1

    Path(args.output).write_text(json.dumps(merged, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
