"""OfficeCLI document snapshot v2 适配器。

该模块不直接读取或修改 DOCX OOXML；所有文档事实必须来自 OfficeCLI
JSON 输出或测试 fixture。旧 `scripts/ooxml` 路径只作为迁移输入保留。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

from scripts.officecli.capability_manifest import canonical_json_bytes


FH_OFFICECLI_SNAPSHOT_INVALID = "FH-OFFICECLI-SNAPSHOT-INVALID"
FH_OFFICECLI_SNAPSHOT_LIMIT = "FH-OFFICECLI-SNAPSHOT-LIMIT"
FH_OFFICECLI_NONJSON_OUTPUT = "FH-OFFICECLI-NONJSON-OUTPUT"
FH_OFFICECLI_PATH_ESCAPE = "FH-OFFICECLI-PATH-ESCAPE"
FH_OFFICECLI_SNAPSHOT_COMPLETENESS = "FH-OFFICECLI-SNAPSHOT-COMPLETENESS"
FH_OFFICECLI_SNAPSHOT_RAW_FAILED = "FH-OFFICECLI-SNAPSHOT-RAW-FAILED"
FH_OFFICECLI_SNAPSHOT_DUMP_WARNING = "FH-OFFICECLI-SNAPSHOT-DUMP-WARNING"


SNAPSHOT_KINDS = {"standard", "before", "after", "post_toc"}
MAX_NODES = 200_000
MAX_STDOUT_BYTES = 256 * 1024 * 1024
ROOT_PATHS = ["/document", "/body", "/styles", "/numbering"]
REQUIRED_ROOT_PATHS = {"/document", "/body", "/styles"}
RAW_PARTS = {
    "/document": ("document", "/word/document.xml", True),
    "/styles": ("styles", "/word/styles.xml", True),
    "/numbering": ("numbering", "/word/numbering.xml", False),
    "/settings": ("settings", "/word/settings.xml", False),
}
BLOCKING_DUMP_TARGETS = {
    "body",
    "styles",
    "numbering",
    "section",
    "header",
    "footer",
    "toc",
    "field",
    "table",
    "form",
}


class SnapshotAdapterError(Exception):
    """OfficeCLI snapshot 适配错误。"""

    def __init__(self, code: str, message: str, detail: Optional[dict] = None):
        self.code = code
        self.message = message
        self.detail = detail or {}
        super().__init__(f"[{code}] {message}")

    def to_json(self) -> dict:
        """转换为稳定 JSON 错误对象。"""
        return {
            "ok": False,
            "error": {
                "code": self.code,
                "message": self.message,
                "detail": self.detail,
            },
        }


def sha256_bytes(value: bytes) -> str:
    """计算 bytes 的 SHA-256。"""
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    """计算文件 SHA-256。"""
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def utc_now() -> str:
    """返回 UTC RFC3339 时间。"""
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def safe_relative_path(path: Path, root: Path) -> str:
    """生成 repo/run 相对路径并阻止路径逃逸。"""
    resolved_path = path.resolve()
    resolved_root = root.resolve()
    try:
        relative = resolved_path.relative_to(resolved_root)
    except ValueError as exc:
        raise SnapshotAdapterError(
            FH_OFFICECLI_PATH_ESCAPE,
            "artifact 路径逃逸根目录",
            {"path": str(path), "root": str(root)},
        ) from exc
    value = str(relative).replace("\\", "/")
    if value.startswith("../") or value == "..":
        raise SnapshotAdapterError(FH_OFFICECLI_PATH_ESCAPE, "artifact 相对路径不合法", {"relative_path": value})
    return value


def artifact_ref(
    path: Path,
    root: Path,
    kind: str,
    schema_id: Optional[str],
    schema_version: Optional[str],
    artifact_id: Optional[str] = None,
) -> dict:
    """构造 ArtifactRef。"""
    if not path.exists():
        raise SnapshotAdapterError(
            FH_OFFICECLI_SNAPSHOT_INVALID,
            "artifact 文件不存在",
            {"path": str(path)},
        )
    relative_path = safe_relative_path(path, root)
    file_hash = sha256_file(path)
    return {
        "artifact_id": artifact_id or f"{kind}-{file_hash[:12]}",
        "kind": kind,
        "relative_path": relative_path,
        "sha256": file_hash,
        "size_bytes": path.stat().st_size,
        "schema_id": schema_id,
        "schema_version": schema_version,
    }


def parse_single_json_stdout(stdout: str) -> Any:
    """严格解析 stdout 中唯一 JSON 值。"""
    text = stdout.strip()
    if not text:
        raise SnapshotAdapterError(FH_OFFICECLI_NONJSON_OUTPUT, "OfficeCLI stdout 为空")
    decoder = json.JSONDecoder()
    try:
        value, index = decoder.raw_decode(text)
    except json.JSONDecodeError as exc:
        raise SnapshotAdapterError(
            FH_OFFICECLI_NONJSON_OUTPUT,
            "OfficeCLI stdout 不是 JSON",
            {"reason": str(exc)},
        ) from exc
    if text[index:].strip():
        raise SnapshotAdapterError(FH_OFFICECLI_NONJSON_OUTPUT, "OfficeCLI stdout 包含 JSON 外垃圾字符")
    return value


def run_officecli_json_with_size(executable: Path, args: list[str], timeout_seconds: int = 120) -> tuple[Any, int]:
    """执行 OfficeCLI JSON 命令并返回解析值和 stdout bytes。"""
    env = os.environ.copy()
    env["OFFICECLI_SKIP_UPDATE"] = "1"
    env["OFFICECLI_NO_AUTO_RESIDENT"] = "1"
    try:
        proc = subprocess.run(
            [str(executable), *args],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout_seconds,
            check=False,
            env=env,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise SnapshotAdapterError(
            FH_OFFICECLI_SNAPSHOT_INVALID,
            "OfficeCLI 子进程执行失败",
            {"args": args, "reason": str(exc)},
        ) from exc
    if proc.returncode != 0:
        raise SnapshotAdapterError(
            FH_OFFICECLI_SNAPSHOT_INVALID,
            "OfficeCLI 子进程返回非零退出码",
            {"args": args, "exit_code": proc.returncode, "stderr": proc.stderr},
        )
    stdout_bytes = proc.stdout.encode("utf-8")
    return parse_single_json_stdout(proc.stdout), len(stdout_bytes)


def run_officecli_json(executable: Path, args: list[str], timeout_seconds: int = 120) -> Any:
    """执行 OfficeCLI 并解析 JSON stdout。"""
    value, _stdout_size = run_officecli_json_with_size(executable, args, timeout_seconds)
    return value


def run_officecli_text(executable: Path, args: list[str], timeout_seconds: int = 120) -> str:
    """执行 OfficeCLI 文本命令，主要用于 raw XML 输出。"""
    env = os.environ.copy()
    env["OFFICECLI_SKIP_UPDATE"] = "1"
    env["OFFICECLI_NO_AUTO_RESIDENT"] = "1"
    try:
        proc = subprocess.run(
            [str(executable), *args],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout_seconds,
            check=False,
            env=env,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise SnapshotAdapterError(
            FH_OFFICECLI_SNAPSHOT_INVALID,
            "OfficeCLI 子进程执行失败",
            {"args": args, "reason": str(exc)},
        ) from exc
    if proc.returncode != 0:
        raise SnapshotAdapterError(
            FH_OFFICECLI_SNAPSHOT_RAW_FAILED,
            "OfficeCLI 子进程返回非零退出码",
            {"args": args, "exit_code": proc.returncode, "stderr": proc.stderr},
        )
    return proc.stdout


def unwrap_officecli_data(value: Any) -> Any:
    """兼容 OfficeCLI success/data envelope。"""
    if isinstance(value, dict) and "success" in value and "data" in value:
        if value.get("success") is not True:
            raise SnapshotAdapterError(
                FH_OFFICECLI_SNAPSHOT_INVALID,
                "OfficeCLI JSON envelope 标记失败",
                {"value": value},
            )
        return value["data"]
    return value


def load_manifest(manifest_path: Path) -> dict:
    """读取 capability manifest。"""
    return json.loads(manifest_path.read_text(encoding="utf-8"))


def load_manifest_element_names(manifest_path: Path) -> set[str]:
    """读取 capability manifest 中的元素集合。"""
    manifest = load_manifest(manifest_path)
    return {item["element"] for item in manifest.get("elements", [])}


def load_queryable_element_names(manifest_path: Path) -> list[str]:
    """读取 capability manifest 中支持 query 的元素。"""
    manifest = load_manifest(manifest_path)
    result = []
    for item in manifest.get("elements", []):
        if "query" in item.get("allowed_operations", []):
            result.append(item["element"])
    return result


def require_officecli_path(value: str) -> None:
    """校验 OfficeCLI path。"""
    if not isinstance(value, str) or not value.startswith("/") or "\n" in value or "\r" in value:
        raise SnapshotAdapterError(
            FH_OFFICECLI_SNAPSHOT_INVALID,
            "OfficeCLI path 不合法",
            {"officecli_path": value},
        )


def raw_node_path(raw_node: dict) -> str:
    """从 OfficeCLI JSON 中提取 canonical path。"""
    value = raw_node.get("officecli_path") or raw_node.get("path") or raw_node.get("canonicalPath")
    require_officecli_path(value)
    return value


def raw_node_type(raw_node: dict) -> str:
    """从 OfficeCLI JSON 中提取节点类型。"""
    value = raw_node.get("node_type") or raw_node.get("type") or raw_node.get("element")
    if not isinstance(value, str) or not value:
        raise SnapshotAdapterError(FH_OFFICECLI_SNAPSHOT_INVALID, "节点缺少类型", {"node": raw_node})
    return value


def raw_node_text(raw_node: dict) -> Optional[str]:
    """从 OfficeCLI JSON 中提取文本。"""
    value = raw_node.get("text")
    if value is None:
        return None
    return str(value)


def raw_node_attributes(raw_node: dict) -> dict:
    """提取 OfficeCLI 节点属性。"""
    value = raw_node.get("attributes")
    if value is None:
        value = raw_node.get("properties")
    return value if isinstance(value, dict) else {}


def raw_child_paths(raw_node: dict) -> list[str]:
    """提取直接 child path。"""
    children = raw_node.get("child_paths") or raw_node.get("children") or []
    result = []
    for child in children:
        if isinstance(child, str):
            require_officecli_path(child)
            result.append(child)
        elif isinstance(child, dict):
            child_path = raw_node_path(child)
            result.append(child_path)
    return result


def extract_paths_from_value(value: Any) -> set[str]:
    """从任意 OfficeCLI JSON 值中递归提取 path。"""
    paths = set()
    if isinstance(value, dict):
        for key in ("officecli_path", "path", "canonicalPath"):
            item = value.get(key)
            if isinstance(item, str) and item.startswith("/"):
                require_officecli_path(item)
                paths.add(item)
        for child in value.values():
            paths.update(extract_paths_from_value(child))
    elif isinstance(value, list):
        for item in value:
            paths.update(extract_paths_from_value(item))
    return paths


def coerce_nodes_from_get_response(value: Any, requested_path: str) -> list[dict]:
    """把 OfficeCLI get 响应转换为节点对象列表。"""
    data = unwrap_officecli_data(value)
    if isinstance(data, dict):
        node = dict(data)
        node.setdefault("path", requested_path)
        return [node]
    if isinstance(data, list):
        nodes = []
        for item in data:
            if isinstance(item, dict):
                node = dict(item)
                node.setdefault("path", requested_path)
                nodes.append(node)
        return nodes
    raise SnapshotAdapterError(
        FH_OFFICECLI_SNAPSHOT_INVALID,
        "OfficeCLI get 返回值不是对象或对象数组",
        {"path": requested_path},
    )


def node_identity_for_conflict_check(raw_node: dict) -> tuple[str, Optional[str]]:
    """返回用于重复 path 冲突检查的类型和 native identity。"""
    path = raw_node_path(raw_node)
    node_type = raw_node_type(raw_node)
    attributes = raw_node_attributes(raw_node)
    return node_type, extract_native_identity(node_type, path, attributes)


def default_json_runner(executable: Path, args: list[str]) -> tuple[Any, int]:
    """默认 OfficeCLI JSON runner。"""
    return run_officecli_json_with_size(executable, args)


def default_raw_runner(executable: Path, args: list[str]) -> str:
    """默认 OfficeCLI raw runner。"""
    return run_officecli_text(executable, args)


def warning_item(code: str, severity: str, message: str, source_command: Optional[str] = None) -> dict:
    """构造 snapshot warning。"""
    return {
        "code": code,
        "severity": severity,
        "message": message,
        "json_pointer": None,
        "source_command": source_command,
    }


def collect_bfs_nodes_with_officecli(
    executable: Path,
    source_docx: Path,
    command_runner=None,
) -> tuple[list[dict], list[dict]]:
    """使用 OfficeCLI get 执行 FIFO BFS 遍历。"""
    runner = command_runner or (lambda args: default_json_runner(executable, args))
    queue = list(ROOT_PATHS)
    queued = set(queue)
    visited: dict[str, dict] = {}
    warnings = []
    stdout_total = 0

    while queue:
        path = queue.pop(0)
        args = ["get", str(source_docx), path, "--depth", "3", "--json"]
        try:
            value, stdout_size = runner(args)
        except SnapshotAdapterError as exc:
            severity = "blocking" if path in REQUIRED_ROOT_PATHS else "warning"
            warnings.append(
                warning_item(
                    "FH-OFFICECLI-SNAPSHOT-ROOT-UNAVAILABLE",
                    severity,
                    f"根路径不可用：{path}",
                    " ".join(args),
                )
            )
            if severity == "blocking":
                raise
            continue
        stdout_total += stdout_size
        if stdout_total > MAX_STDOUT_BYTES:
            raise SnapshotAdapterError(
                FH_OFFICECLI_SNAPSHOT_LIMIT,
                "snapshot 累计 stdout 超过上限",
                {"limit_bytes": MAX_STDOUT_BYTES, "actual_bytes": stdout_total},
            )
        for raw_node in coerce_nodes_from_get_response(value, path):
            current_path = raw_node_path(raw_node)
            if current_path in visited:
                old_type, old_native = node_identity_for_conflict_check(visited[current_path])
                new_type, new_native = node_identity_for_conflict_check(raw_node)
                if old_type != new_type or old_native != new_native:
                    raise SnapshotAdapterError(
                        FH_OFFICECLI_SNAPSHOT_INVALID,
                        "BFS 发现同一路径类型或 native identity 冲突",
                        {
                            "path": current_path,
                            "old_type": old_type,
                            "new_type": new_type,
                            "old_native_identity": old_native,
                            "new_native_identity": new_native,
                        },
                    )
            else:
                visited[current_path] = raw_node
            for child_path in raw_child_paths(raw_node):
                if child_path not in queued and child_path not in visited:
                    queue.append(child_path)
                    queued.add(child_path)
            if len(visited) > MAX_NODES:
                raise SnapshotAdapterError(
                    FH_OFFICECLI_SNAPSHOT_LIMIT,
                    "snapshot BFS 节点数超过上限",
                    {"node_count": len(visited), "limit": MAX_NODES},
                )
    return list(visited.values()), warnings


def verify_query_completeness(
    executable: Path,
    source_docx: Path,
    capability_manifest: Path,
    bfs_nodes: list[dict],
    command_runner=None,
) -> None:
    """校验 queryable 子集必须被 BFS 覆盖。"""
    runner = command_runner or (lambda args: default_json_runner(executable, args))
    bfs_paths_by_type: dict[str, set[str]] = {}
    for node in bfs_nodes:
        bfs_paths_by_type.setdefault(raw_node_type(node), set()).add(raw_node_path(node))
    for element in load_queryable_element_names(capability_manifest):
        args = ["query", str(source_docx), "--type", element, "--json"]
        value, _stdout_size = runner(args)
        query_paths = extract_paths_from_value(unwrap_officecli_data(value))
        missing = sorted(query_paths - bfs_paths_by_type.get(element, set()))
        if missing:
            raise SnapshotAdapterError(
                FH_OFFICECLI_SNAPSHOT_COMPLETENESS,
                "query 返回 BFS 未见的同类型路径",
                {"element": element, "missing_paths": missing},
            )


def dump_warning_target(raw_warning: Any) -> str:
    """从 dump warning 中提取目标字符串。"""
    if isinstance(raw_warning, dict):
        candidates = [
            raw_warning.get("target"),
            raw_warning.get("element"),
            raw_warning.get("type"),
            raw_warning.get("path"),
            raw_warning.get("part"),
            raw_warning.get("message"),
        ]
        return " ".join(str(item) for item in candidates if item is not None).lower()
    return str(raw_warning).lower()


def extract_dump_warning_values(value: Any) -> list[Any]:
    """解析 dump 输出中的 skipped、unsupported 和 warnings。"""
    data = unwrap_officecli_data(value)
    if not isinstance(data, dict):
        return []
    result = []
    for key in ("skipped", "unsupported", "warnings"):
        item = data.get(key)
        if isinstance(item, list):
            result.extend(item)
        elif item:
            result.append(item)
    return result


def collect_dump_warnings(
    executable: Path,
    source_docx: Path,
    command_runner=None,
) -> list[dict]:
    """执行 OfficeCLI dump 并分类 warning。"""
    runner = command_runner or (lambda args: default_json_runner(executable, args))
    args = ["dump", str(source_docx), "--json"]
    value, _stdout_size = runner(args)
    warnings = []
    for raw_warning in extract_dump_warning_values(value):
        target = dump_warning_target(raw_warning)
        severity = "blocking" if any(token in target for token in BLOCKING_DUMP_TARGETS) else "warning"
        message = raw_warning.get("message") if isinstance(raw_warning, dict) else str(raw_warning)
        warnings.append(
            warning_item(
                FH_OFFICECLI_SNAPSHOT_DUMP_WARNING,
                severity,
                str(message),
                " ".join(args),
            )
        )
    return warnings


def raw_part_file_name(part_name: str) -> str:
    """生成 raw part 输出文件名。"""
    return f"{part_name.strip('/').replace('/', '_')}.xml"


def collect_raw_parts(
    executable: Path,
    source_docx: Path,
    run_dir: Path,
    snapshot_kind: str = "before",
    raw_runner=None,
) -> tuple[list[dict], list[dict]]:
    """通过 OfficeCLI raw 采集固定 part 证据。"""
    runner = raw_runner or (lambda args: default_raw_runner(executable, args))
    raw_dir = run_dir / "output" / "_internal" / "officecli" / "raw" / snapshot_kind
    raw_dir.mkdir(parents=True, exist_ok=True)
    parts = []
    warnings = []
    for part_path, (part_name, package_uri, required) in RAW_PARTS.items():
        args = ["raw", str(source_docx), part_path]
        try:
            text = runner(args)
        except SnapshotAdapterError:
            severity = "blocking" if required else "warning"
            warnings.append(
                warning_item(
                    FH_OFFICECLI_SNAPSHOT_RAW_FAILED,
                    severity,
                    f"raw part 采集失败：{part_path}",
                    " ".join(args),
                )
            )
            continue
        output_path = raw_dir / raw_part_file_name(part_name)
        output_path.write_text(text, encoding="utf-8", newline="")
        evidence_ref = artifact_ref(
            output_path,
            infer_workspace_root(run_dir),
            "raw_xml",
            None,
            None,
            f"raw-{snapshot_kind}-{part_name}",
        )
        parts.append(
            {
                "part_name": part_name,
                "package_uri": package_uri,
                "sha256": evidence_ref["sha256"],
                "size_bytes": evidence_ref["size_bytes"],
                "required": required,
                "evidence_ref": evidence_ref,
            }
        )
    return parts, warnings


def collect_snapshot_inputs_with_officecli(
    executable: Path,
    source_docx: Path,
    capability_manifest: Path,
    run_dir: Path,
    snapshot_kind: str = "before",
    command_runner=None,
    raw_runner=None,
) -> dict:
    """采集构建 snapshot v2 所需的 OfficeCLI 输入。"""
    raw_nodes, bfs_warnings = collect_bfs_nodes_with_officecli(executable, source_docx, command_runner)
    verify_query_completeness(executable, source_docx, capability_manifest, raw_nodes, command_runner)
    dump_warnings = collect_dump_warnings(executable, source_docx, command_runner)
    parts, raw_warnings = collect_raw_parts(executable, source_docx, run_dir, snapshot_kind, raw_runner)
    return {
        "raw_nodes": raw_nodes,
        "parts": parts,
        "warnings": [*bfs_warnings, *dump_warnings, *raw_warnings],
    }


def collect_nodes_with_officecli(executable: Path, source_docx: Path) -> list[dict]:
    """兼容旧测试入口：仅返回 BFS 节点。"""
    raw_nodes, _warnings = collect_bfs_nodes_with_officecli(executable, source_docx)
    return raw_nodes


def parent_path_from_path(path: str) -> Optional[str]:
    """根据 OfficeCLI path 推导 parent path。"""
    if path == "/document":
        return None
    prefix = path.rsplit("/", 1)[0]
    return prefix or "/document"


def ordinal_from_path(path: str) -> int:
    """从路径尾部 `[N]` 推导 0-based ordinal，无法推导时返回 0。"""
    tail = path.rsplit("/", 1)[-1]
    if tail.endswith("]") and "[" in tail:
        raw = tail.rsplit("[", 1)[-1][:-1]
        if raw.isdigit() and int(raw) > 0:
            return int(raw) - 1
    return 0


def part_name_from_path(path: str) -> str:
    """按路径推导语义 part name。"""
    if path.startswith("/header"):
        return "header"
    if path.startswith("/footer"):
        return "footer"
    if path.startswith("/styles"):
        return "styles"
    if path.startswith("/numbering"):
        return "numbering"
    return "document"


def extract_native_identity(node_type: str, path: str, attributes: dict) -> Optional[str]:
    """提取可跨快照重绑定的原生 identity。"""
    for key in ("id", "paraId", "styleId", "commentId", "bookmarkId"):
        if key in attributes and attributes[key] not in {None, ""}:
            return f"{node_type}:{key}={attributes[key]}"
    if "@id=" in path or "@paraId=" in path or "@styleId=" in path:
        return f"{node_type}:{path}"
    return None


def extract_logical_identity(node_type: str, path: str, attributes: dict) -> Optional[str]:
    """提取业务逻辑 identity。"""
    for key in ("name", "styleId", "id"):
        if key in attributes and attributes[key] not in {None, ""}:
            return f"{node_type}:{key}={attributes[key]}"
    return None if path == "/document" else f"{node_type}:{path}"


def content_fingerprint(node_type: str, text: Optional[str], attributes: dict) -> str:
    """计算节点内容指纹。"""
    payload = {
        "node_type": node_type,
        "text": text,
        "attributes": attributes,
    }
    return sha256_bytes(canonical_json_bytes(payload))


def stable_selector_for_node(node_type: str, path: str, text: Optional[str], attributes: dict) -> dict:
    """构造 stable selector。"""
    native_identity = extract_native_identity(node_type, path, attributes)
    logical_identity = extract_logical_identity(node_type, path, attributes)
    if native_identity is not None:
        kind = "native_id"
        value = native_identity
        rebindable = True
    elif logical_identity is not None and any(key in attributes for key in ("name", "styleId", "id")):
        kind = "semantic_key"
        value = logical_identity
        rebindable = True
    else:
        kind = "positional"
        value = path
        rebindable = False
    return {
        "kind": kind,
        "value": value,
        "rebindable": rebindable,
        "content_fingerprint": content_fingerprint(node_type, text, attributes),
    }


def normalize_node(raw_node: dict, source_hash: str, allowed_types: set[str]) -> dict:
    """把 OfficeCLI 节点 JSON 归一化为 snapshot v2 node。"""
    path = raw_node_path(raw_node)
    node_type = raw_node_type(raw_node)
    if node_type not in allowed_types and node_type != "document":
        raise SnapshotAdapterError(
            FH_OFFICECLI_SNAPSHOT_INVALID,
            "节点类型未登记在 capability manifest",
            {"node_type": node_type, "path": path},
        )
    text = raw_node_text(raw_node)
    attributes = raw_node_attributes(raw_node)
    node_hash_input = f"{source_hash}\n{path}\n{node_type}".encode("utf-8")
    native_identity = extract_native_identity(node_type, path, attributes)
    logical_identity = extract_logical_identity(node_type, path, attributes)
    return {
        "node_id": "N-" + sha256_bytes(node_hash_input)[:24],
        "officecli_path": path,
        "node_type": node_type,
        "parent_path": raw_node.get("parent_path") or parent_path_from_path(path),
        "ordinal": int(raw_node.get("ordinal", ordinal_from_path(path))),
        "part_name": raw_node.get("part_name") or part_name_from_path(path),
        "text": text,
        "text_sha256": sha256_bytes(text.encode("utf-8")) if text is not None else None,
        "attributes": attributes,
        "effective_format": raw_node.get("effective_format") if isinstance(raw_node.get("effective_format"), dict) else {},
        "effective_sources": raw_node.get("effective_sources") if isinstance(raw_node.get("effective_sources"), dict) else {},
        "child_paths": raw_child_paths(raw_node),
        "stable_selector": stable_selector_for_node(node_type, path, text, attributes),
        "raw_evidence_ref": raw_node.get("raw_evidence_ref"),
        "logical_identity": logical_identity,
        "parent_logical_identity": raw_node.get("parent_logical_identity"),
        "native_identity": native_identity,
    }


def build_indexes(nodes: list[dict]) -> dict:
    """构造 snapshot 索引。"""
    indexes = {
        "by_type": {},
        "by_style_id": {},
        "by_native_id": {},
        "by_logical_identity": {},
    }
    for node in nodes:
        indexes["by_type"].setdefault(node["node_type"], []).append(node["node_id"])
        style_id = node["attributes"].get("styleId") or node["attributes"].get("style")
        if style_id:
            indexes["by_style_id"].setdefault(str(style_id), []).append(node["node_id"])
        if node["native_identity"]:
            indexes["by_native_id"].setdefault(node["native_identity"], []).append(node["node_id"])
        if node["logical_identity"]:
            indexes["by_logical_identity"].setdefault(node["logical_identity"], []).append(node["node_id"])
    return indexes


def document_summary(nodes: list[dict], parts: list[dict]) -> dict:
    """构造 document 摘要。"""
    node_types = {node["node_type"] for node in nodes}
    return {
        "format": "docx",
        "root_path": "/document",
        "node_count": len(nodes),
        "part_count": len(parts),
        "has_toc": "toc" in node_types,
        "has_forms": "formfield" in node_types or "sdt" in node_types,
        "has_revisions": "revision" in node_types,
        "has_protection": any(node["attributes"].get("protection") for node in nodes),
    }


def build_gate_check(created_at: str, evidence_refs: list[dict], failed_codes: Optional[list[str]] = None) -> dict:
    """构造 snapshot GateCheck。"""
    failures = failed_codes or []
    return {
        "gate_id": "officecli-document-snapshot-v2",
        "status": "passed" if not failures else "blocked",
        "checked_at": created_at,
        "predicate_version": "1.0.0",
        "evidence_refs": evidence_refs,
        "failed_codes": failures,
    }


def build_snapshot(
    run_id: str,
    kind: str,
    source_docx: Path,
    officecli_executable: Path,
    capability_manifest: Path,
    artifact_root: Path,
    raw_nodes: list[dict],
    parts: Optional[list[dict]] = None,
    warnings: Optional[list[dict]] = None,
    created_at: Optional[str] = None,
) -> dict:
    """从 OfficeCLI 节点 JSON 构造 snapshot v2。"""
    if kind not in SNAPSHOT_KINDS:
        raise SnapshotAdapterError(FH_OFFICECLI_SNAPSHOT_INVALID, "snapshot kind 不合法", {"kind": kind})
    if len(raw_nodes) > MAX_NODES:
        raise SnapshotAdapterError(FH_OFFICECLI_SNAPSHOT_LIMIT, "snapshot 节点数超过上限", {"node_count": len(raw_nodes)})
    created = created_at or utc_now()
    source_hash = sha256_file(source_docx)
    allowed_types = load_manifest_element_names(capability_manifest)
    nodes = [normalize_node(item, source_hash, allowed_types) for item in raw_nodes]
    normalized_parts = parts or []
    part_evidence_by_name = {
        part["part_name"]: part["evidence_ref"]
        for part in normalized_parts
        if isinstance(part, dict) and isinstance(part.get("evidence_ref"), dict)
    }
    for node in nodes:
        if node["raw_evidence_ref"] is None and node["part_name"] in part_evidence_by_name:
            node["raw_evidence_ref"] = part_evidence_by_name[node["part_name"]]
    node_paths = [node["officecli_path"] for node in nodes]
    if len(node_paths) != len(set(node_paths)):
        raise SnapshotAdapterError(FH_OFFICECLI_SNAPSHOT_INVALID, "snapshot 存在重复 OfficeCLI path")
    normalized_warnings = warnings or []
    evidence_refs = [
        artifact_ref(source_docx, artifact_root, "docx", None, None, "source-docx"),
        artifact_ref(officecli_executable, artifact_root, "executable", None, None, "officecli-executable"),
        artifact_ref(capability_manifest, artifact_root, "capability", "officecli-capability-manifest", "1.0.0", "officecli-capability"),
        *part_evidence_by_name.values(),
    ]
    failed_codes = [item["code"] for item in normalized_warnings if item.get("severity") == "blocking"]
    snapshot_core = {
        "run_id": run_id,
        "kind": kind,
        "source_hash": source_hash,
        "nodes": nodes,
        "parts": normalized_parts,
    }
    snapshot_id = "S-" + sha256_bytes(canonical_json_bytes(snapshot_core))[:24]
    return {
        "schema_id": "officecli-document-snapshot",
        "schema_version": "2.0.0",
        "contract_version": "v5",
        "snapshot_id": snapshot_id,
        "kind": kind,
        "run_id": run_id,
        "officecli_version": "1.0.113",
        "created_at": created,
        "extensions": {},
        "source_docx_ref": evidence_refs[0],
        "officecli_executable_ref": evidence_refs[1],
        "capability_manifest_ref": evidence_refs[2],
        "snapshot_source_hash": source_hash,
        "document": document_summary(nodes, normalized_parts),
        "nodes": nodes,
        "parts": normalized_parts,
        "indexes": build_indexes(nodes),
        "warnings": normalized_warnings,
        "gate_check": build_gate_check(created, evidence_refs, failed_codes),
    }


def write_json_atomic(path: Path, value: dict) -> None:
    """原子写出 UTF-8 JSON。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_name(path.name + ".tmp")
    temp_path.write_bytes(canonical_json_bytes(value) + b"\n")
    temp_path.replace(path)


def default_snapshot_path(run_dir: Path, kind: str) -> Path:
    """返回标准 snapshot 输出路径。"""
    return run_dir / "snapshots" / f"officecli-document-snapshot.{kind}.json"


def infer_workspace_root(run_dir: Path) -> Path:
    """从 run_dir 推导工作区根目录。"""
    resolved = run_dir.resolve()
    if resolved.parent.name == "format_runs":
        return resolved.parent.parent
    return resolved.parent


def parse_args(argv: Optional[list] = None) -> argparse.Namespace:
    """解析命令行参数。"""
    parser = argparse.ArgumentParser(description="OfficeCLI document snapshot v2 适配器")
    sub = parser.add_subparsers(dest="command", required=True)
    build = sub.add_parser("build", help="构造 officecli-document-snapshot v2")
    build.add_argument("--run-dir", required=True, type=Path)
    build.add_argument("--run-id")
    build.add_argument("--kind", required=True, choices=sorted(SNAPSHOT_KINDS))
    build.add_argument("--source-docx", type=Path)
    build.add_argument("--officecli-executable", required=True, type=Path)
    build.add_argument("--capability-manifest", required=True, type=Path)
    build.add_argument("--fixture-nodes", type=Path)
    build.add_argument("--out", type=Path)
    build.add_argument("--created-at")
    return parser.parse_args(argv)


def main(argv: Optional[list] = None) -> int:
    """CLI 入口。"""
    args = parse_args(argv)
    try:
        if args.command != "build":
            raise SnapshotAdapterError(FH_OFFICECLI_SNAPSHOT_INVALID, f"未知命令：{args.command}")
        run_dir = args.run_dir.resolve()
        run_id = args.run_id or run_dir.name
        source_docx = (args.source_docx or (run_dir / "input" / "working.docx")).resolve()
        if args.fixture_nodes:
            raw_nodes = json.loads(args.fixture_nodes.read_text(encoding="utf-8"))
            if not isinstance(raw_nodes, list):
                raise SnapshotAdapterError(FH_OFFICECLI_SNAPSHOT_INVALID, "fixture-nodes 必须是数组")
            parts = []
            warnings = []
        else:
            collected = collect_snapshot_inputs_with_officecli(
                args.officecli_executable.resolve(),
                source_docx,
                args.capability_manifest.resolve(),
                run_dir,
                args.kind,
            )
            raw_nodes = collected["raw_nodes"]
            parts = collected["parts"]
            warnings = collected["warnings"]
        snapshot = build_snapshot(
            run_id=run_id,
            kind=args.kind,
            source_docx=source_docx,
            officecli_executable=args.officecli_executable.resolve(),
            capability_manifest=args.capability_manifest.resolve(),
            artifact_root=infer_workspace_root(run_dir),
            raw_nodes=raw_nodes,
            parts=parts,
            warnings=warnings,
            created_at=args.created_at,
        )
        out_path = args.out or default_snapshot_path(run_dir, args.kind)
        write_json_atomic(out_path, snapshot)
        result = {
            "ok": True,
            "snapshot": str(out_path),
            "snapshot_id": snapshot["snapshot_id"],
            "node_count": snapshot["document"]["node_count"],
        }
        sys.stdout.write(json.dumps(result, ensure_ascii=False, sort_keys=True) + "\n")
        return 0
    except SnapshotAdapterError as exc:
        sys.stdout.write(json.dumps(exc.to_json(), ensure_ascii=False, sort_keys=True) + "\n")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
