"""项目内受限 YAML 读写工具。

该模块只支持 format-helper 生成的安全子集：字典、列表、字符串、数字、
布尔值和 null。避免为 CODE-005 引入额外依赖。
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def dump_yaml(data: Any) -> str:
    """将受限结构序列化为稳定 YAML 文本。"""
    return "\n".join(_dump_node(data, 0)) + "\n"


def write_yaml(path: Path, data: Any) -> None:
    """写入 YAML 文件。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(dump_yaml(data), encoding="utf-8")


def load_yaml(path: Path) -> Any:
    """读取受限 YAML 文件。"""
    if isinstance(path, str):
        path = Path(path)
    lines = _prepare_lines(path.read_text(encoding="utf-8"))
    if not lines:
        return None
    result, index = _parse_block(lines, 0, lines[0][0])
    if index != len(lines):
        raise ValueError(f"无法解析 YAML 第 {index + 1} 行")
    return result


def _dump_node(value: Any, indent: int) -> list[str]:
    prefix = " " * indent
    if isinstance(value, dict):
        if not value:
            return [prefix + "{}"]
        lines: list[str] = []
        for key, item in value.items():
            if isinstance(item, (dict, list)):
                if not item:
                    lines.append(f"{prefix}{key}: {_format_scalar(item)}")
                else:
                    lines.append(f"{prefix}{key}:")
                    lines.extend(_dump_node(item, indent + 2))
            else:
                lines.append(f"{prefix}{key}: {_format_scalar(item)}")
        return lines
    if isinstance(value, list):
        if not value:
            return [prefix + "[]"]
        lines = []
        for item in value:
            if isinstance(item, (dict, list)):
                lines.append(prefix + "-")
                lines.extend(_dump_node(item, indent + 2))
            else:
                lines.append(f"{prefix}- {_format_scalar(item)}")
        return lines
    return [prefix + _format_scalar(value)]


def _format_scalar(value: Any) -> str:
    if value is None:
        return "null"
    if value is True:
        return "true"
    if value is False:
        return "false"
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    return json.dumps(str(value), ensure_ascii=False)


def _prepare_lines(text: str) -> list[tuple[int, str]]:
    lines: list[tuple[int, str]] = []
    for raw in text.splitlines():
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        indent = len(raw) - len(raw.lstrip(" "))
        if indent % 2:
            raise ValueError(f"YAML 缩进必须使用 2 的倍数：{raw}")
        lines.append((indent, raw.strip()))
    return lines


def _parse_block(lines: list[tuple[int, str]], index: int, indent: int) -> tuple[Any, int]:
    if index >= len(lines):
        return None, index
    current_indent, stripped = lines[index]
    if current_indent != indent:
        raise ValueError(f"YAML 缩进不匹配：期望 {indent}，实际 {current_indent}")
    if stripped.startswith("-"):
        return _parse_list(lines, index, indent)
    return _parse_dict(lines, index, indent)


def _parse_list(lines: list[tuple[int, str]], index: int, indent: int) -> tuple[list[Any], int]:
    result: list[Any] = []
    while index < len(lines):
        current_indent, stripped = lines[index]
        if current_indent != indent or not stripped.startswith("-"):
            break
        tail = stripped[1:].strip()
        index += 1
        if not tail:
            if index < len(lines) and lines[index][0] > indent:
                item, index = _parse_block(lines, index, lines[index][0])
            else:
                item = None
            result.append(item)
            continue
        if ":" in tail and not tail.startswith(('"', "'")):
            key, value = tail.split(":", 1)
            item = {key.strip(): _parse_scalar(value.strip()) if value.strip() else None}
            if index < len(lines) and lines[index][0] > indent:
                nested, index = _parse_dict(lines, index, lines[index][0])
                item.update(nested)
            result.append(item)
            continue
        result.append(_parse_scalar(tail))
    return result, index


def _parse_dict(lines: list[tuple[int, str]], index: int, indent: int) -> tuple[dict[str, Any], int]:
    result: dict[str, Any] = {}
    while index < len(lines):
        current_indent, stripped = lines[index]
        if current_indent != indent or stripped.startswith("-"):
            break
        if ":" not in stripped:
            raise ValueError(f"YAML 字典行缺少冒号：{stripped}")
        key, value = stripped.split(":", 1)
        key = key.strip()
        value = value.strip()
        index += 1
        if value:
            result[key] = _parse_scalar(value)
            continue
        if index < len(lines) and lines[index][0] > indent:
            result[key], index = _parse_block(lines, index, lines[index][0])
        else:
            result[key] = None
    return result, index


def _parse_scalar(value: str) -> Any:
    if value in {"null", "~"}:
        return None
    if value == "true":
        return True
    if value == "false":
        return False
    if value in {"[]", "{}"} or value.startswith(('"', "'", "[", "{")):
        try:
            return json.loads(value)
        except json.JSONDecodeError:
            return value.strip('"').strip("'")
    try:
        if "." in value:
            return float(value)
        return int(value)
    except ValueError:
        return value
