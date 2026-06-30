"""OfficeCLI DOCX 能力清单生成与校验。

生成器以 OfficeCLI v1.0.113 固定源码中的 DOCX help target 为迭代输入，逐项消费
`officecli help docx {help_target} --json` 的结果，并生成可在 CI 中
比对的 capability manifest。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional

from scripts.officecli.runtime_resolver import (
    FH_OFFICECLI_LOCK_INVALID,
    OfficeCliRuntimeError,
    load_lock,
)


FH_OFFICECLI_CAPABILITY_INVALID = "FH-OFFICECLI-CAPABILITY-INVALID"
FH_OFFICECLI_HELP_FAILED = "FH-OFFICECLI-HELP-FAILED"


DOCX_HELP_TARGETS = [
    ("abstractNum", "abstractNum"),
    ("body", "body"),
    ("bookmark", "bookmark"),
    ("chart-axis", "chart-axis"),
    ("chart-series", "chart-series"),
    ("chart", "chart"),
    ("comment", "comment"),
    ("document", "document"),
    ("endnote", "endnote"),
    ("equation", "equation"),
    ("field", "field"),
    ("fieldchar", "fieldChar"),
    ("footer", "footer"),
    ("footnote", "footnote"),
    ("formfield", "formfield"),
    ("header", "header"),
    ("hyperlink", "hyperlink"),
    ("instrtext", "instrText"),
    ("level", "level"),
    ("num", "num"),
    ("numbering", "numbering"),
    ("ole", "ole"),
    ("pagebreak", "pagebreak"),
    ("paragraph", "paragraph"),
    ("permStart", "permStart"),
    ("picture", "picture"),
    ("ptab", "ptab"),
    ("raw", "raw"),
    ("revision", "revision"),
    ("run", "run"),
    ("sdt", "sdt"),
    ("section", "section"),
    ("style", "style"),
    ("styles", "styles"),
    ("table-cell", "cell"),
    ("table-column", "column"),
    ("table-row", "row"),
    ("table", "table"),
    ("toc", "toc"),
    ("watermark", "watermark"),
]


DOCX_ELEMENT_OPERATIONS = {
    "abstractNum": ["add", "set", "get", "query", "remove"],
    "body": ["get", "query"],
    "bookmark": ["add", "set", "get", "query", "remove"],
    "chart": ["add", "set", "get", "query", "remove"],
    "chart-axis": ["set", "get"],
    "chart-series": ["add", "set", "get", "remove"],
    "comment": ["add", "set", "get", "query", "remove"],
    "document": ["set", "get", "query"],
    "endnote": ["add", "set", "get", "query", "remove"],
    "equation": ["add", "set", "get", "query", "remove"],
    "field": ["add", "set", "get", "query", "remove"],
    "fieldChar": ["set", "get", "query", "remove"],
    "footer": ["add", "set", "get", "query", "remove"],
    "footnote": ["add", "set", "get", "query", "remove"],
    "formfield": ["add", "set", "get", "query", "remove"],
    "header": ["add", "set", "get", "query", "remove"],
    "hyperlink": ["add", "set", "get", "query", "remove"],
    "instrText": ["set", "get", "query", "remove"],
    "level": ["add", "set", "get", "remove"],
    "num": ["add", "set", "get", "query", "remove"],
    "numbering": ["get", "query"],
    "ole": ["add", "set", "get", "query", "remove"],
    "pagebreak": ["add", "set", "get", "query", "remove"],
    "paragraph": ["add", "set", "get", "query", "remove"],
    "permStart": ["add", "get", "remove"],
    "picture": ["add", "set", "get", "query", "remove"],
    "ptab": ["add", "set", "get", "query", "remove"],
    "raw": [],
    "revision": ["set", "get", "query"],
    "run": ["add", "set", "get", "remove"],
    "sdt": ["add", "set", "get", "query", "remove"],
    "section": ["add", "set", "get", "query", "remove"],
    "style": ["add", "set", "get", "query", "remove"],
    "styles": ["add", "get", "query"],
    "table": ["add", "set", "get", "query", "remove"],
    "cell": ["add", "set", "get", "query", "remove"],
    "column": ["add", "remove"],
    "row": ["add", "set", "get", "query", "remove"],
    "toc": ["add", "set", "get", "query", "remove"],
    "watermark": ["add", "set", "get", "query", "remove"],
}


ROOT_COMMAND_POLICIES = {
    "create": ("conditional", "fixture_or_explicit_migration_only"),
    "import": ("conditional", "fixture_or_explicit_migration_only"),
    "open": ("deny", "resident_session_not_used_in_format_helper"),
    "close": ("deny", "resident_session_not_used_in_format_helper"),
    "save": ("deny", "resident_session_not_used_in_format_helper"),
    "get": ("allow", "docx_read"),
    "query": ("allow", "docx_read"),
    "set": ("allow", "docx_l1_l2_write"),
    "add": ("allow", "docx_l1_l2_write"),
    "remove": ("allow", "docx_l1_l2_write"),
    "move": ("allow", "docx_l1_l2_write"),
    "swap": ("allow", "docx_l1_l2_write"),
    "batch": ("allow", "deterministic_batch"),
    "dump": ("allow", "snapshot_support"),
    "raw": ("allow", "l3_read_allowed"),
    "raw-set": ("allow", "manual_confirmation_required"),
    "add-part": ("deny", "out_of_scope_for_format_helper"),
    "validate": ("allow", "post_write_gate"),
    "view": ("allow", "preview_and_evidence"),
    "merge": ("conditional", "explicit_migration_only"),
    "refresh": ("conditional", "not_native_toc_acceptance"),
    "watch": ("deny", "interactive_mode_not_allowed"),
    "goto": ("deny", "interactive_mode_not_allowed"),
    "mark": ("deny", "interactive_mode_not_allowed"),
    "plugins": ("deny", "plugin_management_not_allowed"),
    "mcp": ("deny", "mcp_not_formal_automation_interface"),
    "skills": ("deny", "skill_management_not_allowed"),
    "update": ("deny", "auto_update_disabled"),
    "uninstall": ("deny", "runtime_mutation_not_allowed"),
}


VIEW_MODES = {
    "text": (True, "text/plain"),
    "annotated": (True, "text/plain"),
    "outline": (True, "application/json"),
    "stats": (True, "application/json"),
    "issues": (True, "application/json"),
    "html": (True, "text/html"),
    "screenshot": (True, "image/png"),
    "pdf": (True, "application/pdf"),
    "forms": (True, "application/json"),
    "svg": (False, "image/svg+xml"),
}


def canonical_json_bytes(value: Any) -> bytes:
    """生成稳定 JSON bytes。"""
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")


def sha256_bytes(value: bytes) -> str:
    """计算 bytes 的 SHA-256。"""
    return hashlib.sha256(value).hexdigest()


def unwrap_help_payload(payload: dict) -> dict:
    """兼容 OfficeCLI JSON envelope 或直接 payload。"""
    if isinstance(payload, dict) and "success" in payload and "data" in payload:
        if payload.get("success") is not True:
            raise OfficeCliRuntimeError(
                FH_OFFICECLI_CAPABILITY_INVALID,
                "OfficeCLI help JSON envelope 标记失败",
                {"payload": payload},
            )
        data = payload["data"]
        if not isinstance(data, dict):
            raise OfficeCliRuntimeError(
                FH_OFFICECLI_CAPABILITY_INVALID,
                "OfficeCLI help JSON data 不是对象",
            )
        return data
    return payload


def normalize_element_help(element: str, payload: dict, raw_hash: str) -> dict:
    """归一化单个元素 help JSON。"""
    data = unwrap_help_payload(payload)
    data_element = data.get("element", element)
    if data_element != element:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_CAPABILITY_INVALID,
            "OfficeCLI help 元素名与请求不一致",
            {"expected": element, "actual": data_element},
        )
    operations = data.get("operations", {op: True for op in DOCX_ELEMENT_OPERATIONS[element]})
    if isinstance(operations, dict):
        allowed_operations = [op for op, enabled in operations.items() if enabled is True]
    else:
        allowed_operations = list(operations)
    missing_ops = [op for op in DOCX_ELEMENT_OPERATIONS[element] if op not in allowed_operations]
    extra_ops = [op for op in allowed_operations if op not in DOCX_ELEMENT_OPERATIONS[element]]
    if missing_ops or extra_ops:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_CAPABILITY_INVALID,
            "OfficeCLI help 操作集合与规范不一致",
            {"element": element, "missing_operations": missing_ops, "extra_operations": extra_ops},
        )
    return {
        "element": element,
        "elementAliases": data.get("elementAliases", []),
        "operations": operations,
        "allowed_operations": allowed_operations,
        "properties": data.get("properties", []),
        "children": data.get("children", []),
        "note": data.get("note"),
        "raw_help_sha256": raw_hash,
    }


def load_help_json_from_dir(help_dir: Path, help_target: str) -> tuple[dict, str]:
    """从目录读取元素 help JSON。"""
    candidates = [
        help_dir / f"{help_target}.json",
        help_dir / f"{help_target.replace('-', '_')}.json",
    ]
    for path in candidates:
        if path.exists():
            raw = path.read_bytes()
            return json.loads(raw.decode("utf-8")), sha256_bytes(raw)
    raise OfficeCliRuntimeError(
        FH_OFFICECLI_CAPABILITY_INVALID,
        "缺少元素 help JSON",
        {"help_target": help_target, "help_dir": str(help_dir)},
    )


def load_help_json_from_binary(executable: Path, help_target: str, timeout_seconds: int = 30) -> tuple[dict, str]:
    """调用固定 OfficeCLI 二进制读取元素 help JSON。"""
    try:
        proc = subprocess.run(
            [str(executable), "help", "docx", help_target, "--json"],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout_seconds,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_HELP_FAILED,
            "OfficeCLI help 调用失败",
            {"help_target": help_target, "reason": str(exc)},
        )
    if proc.returncode != 0:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_HELP_FAILED,
            "OfficeCLI help 返回非零退出码",
            {"help_target": help_target, "exit_code": proc.returncode, "stderr": proc.stderr},
        )
    raw = proc.stdout.encode("utf-8")
    try:
        return json.loads(proc.stdout), sha256_bytes(raw)
    except json.JSONDecodeError as exc:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_HELP_FAILED,
            "OfficeCLI help stdout 不是 JSON",
            {"help_target": help_target, "reason": str(exc)},
        )


def build_global_commands() -> list[dict]:
    """构建固定源码命令分类快照。"""
    commands = []
    for name in sorted(ROOT_COMMAND_POLICIES):
        policy, reason = ROOT_COMMAND_POLICIES[name]
        commands.append(
            {
                "name": name,
                "docx_supported": policy in {"allow", "conditional"},
                "automation_surface": "cli-json",
                "risk_class": "L3_WRITE" if name == "raw-set" else "L2" if policy == "allow" else "DENY",
                "production_policy": policy,
                "reason": reason,
            }
        )
    return commands


def build_view_modes() -> list[dict]:
    """构建 view 模式分类快照。"""
    return [
        {
            "mode": mode,
            "docx_supported": docx_supported,
            "output_type": output_type,
            "production_policy": "allow" if docx_supported else "deny",
        }
        for mode, (docx_supported, output_type) in sorted(VIEW_MODES.items())
    ]


def build_manifest(
    lock: dict,
    help_dir: Optional[Path] = None,
    executable: Optional[Path] = None,
    generated_at: Optional[str] = None,
    generator_version: str = "0.1.0",
) -> dict:
    """生成 OfficeCLI capability manifest。"""
    if help_dir is None and executable is None:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_CAPABILITY_INVALID,
            "必须提供 help_dir 或 executable",
        )
    elements = []
    for help_target, element in DOCX_HELP_TARGETS:
        if help_dir is not None:
            payload, raw_hash = load_help_json_from_dir(help_dir, help_target)
        else:
            payload, raw_hash = load_help_json_from_binary(Path(executable), help_target)
        normalized = normalize_element_help(element, payload, raw_hash)
        normalized["help_target"] = help_target
        elements.append(normalized)
    aggregate_source = {
        "elements": elements,
        "global_commands": build_global_commands(),
        "view_modes": build_view_modes(),
    }
    return {
        "schema_id": "officecli-capability-manifest",
        "schema_version": "1.0.0",
        "officecli_version": lock["officecli_version"],
        "source_commit": lock["source_commit"],
        "generated_at": generated_at
        or datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
        "generator_version": generator_version,
        "global_commands": aggregate_source["global_commands"],
        "view_modes": aggregate_source["view_modes"],
        "elements": elements,
        "raw_read_allowed": True,
        "raw_write_policy": "manual_confirmation_required",
        "aggregate_sha256": sha256_bytes(canonical_json_bytes(aggregate_source)),
    }


def verify_manifest(manifest: dict, lock: dict) -> list[str]:
    """校验 manifest 与锁文件及固定覆盖基线一致。"""
    errors = []
    for key, expected in {
        "schema_id": "officecli-capability-manifest",
        "schema_version": "1.0.0",
        "officecli_version": lock["officecli_version"],
        "source_commit": lock["source_commit"],
        "raw_write_policy": "manual_confirmation_required",
    }.items():
        if manifest.get(key) != expected:
            errors.append(f"{key} 应为 {expected}，实际为 {manifest.get(key)}")
    element_names = [item.get("element") for item in manifest.get("elements", [])]
    help_targets = [item.get("help_target") for item in manifest.get("elements", [])]
    expected_names = [element for _help_target, element in DOCX_HELP_TARGETS]
    expected_help_targets = [help_target for help_target, _element in DOCX_HELP_TARGETS]
    if element_names != expected_names:
        errors.append("DOCX 元素顺序或集合与 OfficeCLI v1.0.113 固定 help 清单不一致")
    if help_targets != expected_help_targets:
        errors.append("DOCX help_target 顺序或集合与 OfficeCLI v1.0.113 固定 help 清单不一致")
    for item in manifest.get("elements", []):
        element = item.get("element")
        if element not in DOCX_ELEMENT_OPERATIONS:
            errors.append(f"未知元素：{element}")
            continue
        allowed_operations = item.get("allowed_operations")
        if allowed_operations is None:
            raw_operations = item.get("operations", [])
            if isinstance(raw_operations, dict):
                allowed_operations = [op for op, enabled in raw_operations.items() if enabled is True]
            else:
                allowed_operations = list(raw_operations)
        missing = [op for op in DOCX_ELEMENT_OPERATIONS[element] if op not in allowed_operations]
        extra = [op for op in allowed_operations if op not in DOCX_ELEMENT_OPERATIONS[element]]
        if missing:
            errors.append(f"{element} 缺少操作：{','.join(missing)}")
        if extra:
            errors.append(f"{element} 存在未分类操作：{','.join(extra)}")
    command_names = [item.get("name") for item in manifest.get("global_commands", [])]
    expected_command_names = sorted(ROOT_COMMAND_POLICIES)
    if command_names != expected_command_names:
        errors.append("全局命令集合或顺序与固定源码分类不一致")
    commands = {item.get("name"): item for item in manifest.get("global_commands", [])}
    for name, (policy, reason) in ROOT_COMMAND_POLICIES.items():
        command = commands.get(name)
        if command is None:
            errors.append(f"全局命令未分类：{name}")
            continue
        if command.get("production_policy") != policy or command.get("reason") != reason:
            errors.append(f"全局命令 {name} 的策略或原因与固定分类不一致")
    view_modes = {item.get("mode"): item for item in manifest.get("view_modes", [])}
    view_mode_names = [item.get("mode") for item in manifest.get("view_modes", [])]
    if view_mode_names != sorted(VIEW_MODES):
        errors.append("view mode 集合或顺序与固定分类不一致")
    if view_modes.get("svg", {}).get("docx_supported") is not False:
        errors.append("svg view 模式必须登记为 docx_supported=false")
    aggregate_source = {
        "elements": manifest.get("elements", []),
        "global_commands": manifest.get("global_commands", []),
        "view_modes": manifest.get("view_modes", []),
    }
    expected_aggregate = sha256_bytes(canonical_json_bytes(aggregate_source))
    if manifest.get("aggregate_sha256") != expected_aggregate:
        errors.append("aggregate_sha256 与 manifest 内容不一致")
    return errors


def write_json(path: Path, value: dict) -> None:
    """以 UTF-8 写出稳定 JSON。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(canonical_json_bytes(value) + b"\n")


def parse_args(argv: Optional[list] = None) -> argparse.Namespace:
    """解析命令行参数。"""
    parser = argparse.ArgumentParser(description="OfficeCLI capability manifest 工具")
    sub = parser.add_subparsers(dest="command", required=True)
    generate = sub.add_parser("generate", help="生成 capability manifest")
    generate.add_argument("--lock", required=True, type=Path)
    generate.add_argument("--help-dir", type=Path)
    generate.add_argument("--executable", type=Path)
    generate.add_argument("--out", required=True, type=Path)
    generate.add_argument("--generated-at")
    verify = sub.add_parser("verify", help="校验 capability manifest")
    verify.add_argument("--lock", required=True, type=Path)
    verify.add_argument("--manifest", required=True, type=Path)
    return parser.parse_args(argv)


def main(argv: Optional[list] = None) -> int:
    """CLI 入口。"""
    args = parse_args(argv)
    try:
        lock = load_lock(args.lock)
        if args.command == "generate":
            manifest = build_manifest(
                lock,
                help_dir=args.help_dir,
                executable=args.executable,
                generated_at=args.generated_at,
            )
            write_json(args.out, manifest)
            result = {"ok": True, "manifest": str(args.out), "aggregate_sha256": manifest["aggregate_sha256"]}
        elif args.command == "verify":
            manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
            errors = verify_manifest(manifest, lock)
            result = {"ok": not errors, "errors": errors}
        else:
            raise OfficeCliRuntimeError(FH_OFFICECLI_LOCK_INVALID, f"未知命令：{args.command}")
        sys.stdout.write(json.dumps(result, ensure_ascii=False, sort_keys=True) + "\n")
        return 0 if result["ok"] else 2
    except OfficeCliRuntimeError as exc:
        sys.stdout.write(json.dumps(exc.to_json(), ensure_ascii=False, sort_keys=True) + "\n")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
