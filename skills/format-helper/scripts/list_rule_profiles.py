#!/usr/bin/env python3
"""列出 format_rules/docx/rule_profiles 下的规则版本。"""
from __future__ import annotations

import argparse
import json
from pathlib import Path


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def parse_simple_yaml(text: str) -> dict:
    data = {}
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or ":" not in line:
            continue
        key, value = line.split(":", 1)
        value = value.strip().strip('"').strip("'")
        data[key.strip()] = value
    return data


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", default="format_rules/docx/rule_profiles")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    root = Path(args.root)
    profiles = []
    if root.exists():
        for item in sorted(p for p in root.iterdir() if p.is_dir()):
            profile_path = item / "profile.yaml"
            summary_path = item / "RULE_SUMMARY.md"
            meta = parse_simple_yaml(read_text(profile_path)) if profile_path.exists() else {}
            profiles.append(
                {
                    "id": meta.get("id", item.name),
                    "name": meta.get("name", item.name),
                    "version": meta.get("version", ""),
                    "status": meta.get("status", "unknown"),
                    "description": meta.get("description", ""),
                    "path": str(item),
                    "has_rule_summary": summary_path.exists(),
                }
            )

    if args.json:
        print(json.dumps(profiles, ensure_ascii=False, indent=2))
    else:
        for profile in profiles:
            print(f"{profile['id']}\t{profile['version']}\t{profile['status']}\t{profile['description']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
