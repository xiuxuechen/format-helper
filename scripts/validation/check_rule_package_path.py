"""规则包路径校验工具。

根据 40-§5.2, 40-§3.4, 41-§11.2, 41-§11.3 实施路径约束：
- 规则包只能写入 format-rules/{rule_id}/
- 禁止写入 format_runs/*/rules
- 禁止写入 format_rules/（历史别名）
- rule_id 必须使用 [A-Za-z0-9_-]，长度 1-80
- 禁止路径穿越（..、绝对路径、symlink 跳出）
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional


# 错误码（对应 docx-rule-packager SKILL.md 失败模板）
RP_PATH_INVALID = "RP-PATH-INVALID"
RP_RULE_ID_INVALID = "RP-RULE-ID-INVALID"
RP_PATH_FORBIDDEN = "RP-PATH-FORBIDDEN"
RP_PATH_ESCAPE = "RP-PATH-ESCAPE"


# 规则 ID 正则（参考 41-§3.1 路径安全规则）
RULE_ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]{1,80}$")


class RulePackagePathError(Exception):
    """规则包路径校验错误。"""

    def __init__(self, code: str, message: str):
        self.code = code
        self.message = message
        super().__init__(f"[{code}] {message}")


def validate_rule_id(rule_id: str) -> None:
    """校验 rule_id 合法性（参考 40-§3.4, 41-§3.1）。

    Args:
        rule_id: 规则 ID

    Raises:
        RulePackagePathError: rule_id 不合法
    """
    if not rule_id:
        raise RulePackagePathError(
            RP_RULE_ID_INVALID,
            "rule_id 不能为空",
        )

    if not RULE_ID_PATTERN.match(rule_id):
        raise RulePackagePathError(
            RP_RULE_ID_INVALID,
            f"rule_id='{rule_id}' 不合法，只能使用 [A-Za-z0-9_-]，长度 1-80",
        )

    if rule_id in {".", ".."}:
        raise RulePackagePathError(
            RP_RULE_ID_INVALID,
            f"rule_id='{rule_id}' 不能是 . 或 ..",
        )


def validate_rule_package_path(
    target_path: str | Path,
    workspace_root: Optional[Path] = None,
) -> Path:
    """校验规则包目标路径是否符合 40-§5.2 约束。

    规则：
    - 必须位于 format-rules/{rule_id}/ 下
    - 禁止位于 format_runs/*/rules（参考 40-§5.2）
    - 禁止位于 format_rules/（历史别名）
    - 禁止路径穿越（.. 或绝对路径）

    Args:
        target_path: 目标路径（可以是字符串或 Path）
        workspace_root: 工作区根目录（用于检查路径是否在工作区内）

    Returns:
        规范化后的 Path 对象

    Raises:
        RulePackagePathError: 路径不符合约束
    """
    path = Path(target_path)
    path_str = str(path).replace("\\", "/")

    # 禁止 format_runs/*/rules（参考 40-§5.2）
    # 匹配模式：format_runs/ 开头或 /format_runs/ 包含，且路径片段中存在 "rules"
    path_parts = path.parts
    if len(path_parts) >= 3 and path_parts[0] == "format_runs" and "rules" in path_parts[2:]:
        raise RulePackagePathError(
            RP_PATH_FORBIDDEN,
            f"规则包不得写入 format_runs/*/rules：{path_str}。"
            f"规则包只能写入 format-rules/{{rule_id}}/",
        )
    # 兜底：检查字符串形式（处理绝对路径或其他格式）
    if "format_runs/" in path_str and "/rules" in path_str:
        # 进一步检查 rules 是否作为目录名（避免误伤 format_runs/abc/sub/rules_backup 这种）
        segments = path_str.split("/")
        for i, seg in enumerate(segments):
            if seg == "format_runs" and i + 2 < len(segments) and segments[i + 2] == "rules":
                raise RulePackagePathError(
                    RP_PATH_FORBIDDEN,
                    f"规则包不得写入 format_runs/*/rules：{path_str}。"
                    f"规则包只能写入 format-rules/{{rule_id}}/",
                )

    # 禁止 format_rules/（历史别名，v4 应使用 format-rules/）
    if path_str.startswith("format_rules/") or "/format_rules/" in path_str:
        raise RulePackagePathError(
            RP_PATH_FORBIDDEN,
            f"规则包路径使用了历史别名 format_rules/：{path_str}。"
            f"v4 规则库目录必须为 format-rules/（参考 40-§5.2）",
        )

    # 禁止路径穿越（参考 40-§3.4）
    if ".." in path.parts:
        raise RulePackagePathError(
            RP_PATH_ESCAPE,
            f"路径包含 ..，禁止路径穿越：{path_str}",
        )

    # 必须在 format-rules/ 下
    parts = path.parts
    if len(parts) < 2 or parts[0] != "format-rules":
        raise RulePackagePathError(
            RP_PATH_INVALID,
            f"规则包必须位于 format-rules/{{rule_id}}/ 下：{path_str}",
        )

    # 校验 rule_id
    rule_id = parts[1]
    validate_rule_id(rule_id)

    # 检查 workspace_root（可选）
    if workspace_root is not None:
        workspace_root = Path(workspace_root).resolve()
        try:
            resolved = (workspace_root / path).resolve()
            resolved.relative_to(workspace_root)
        except ValueError:
            raise RulePackagePathError(
                RP_PATH_ESCAPE,
                f"路径解析后不在工作区内：{path_str}",
            )

    return path


def get_rule_package_dir(rule_id: str) -> Path:
    """获取规则包目录路径。

    Args:
        rule_id: 规则 ID

    Returns:
        规则包目录路径（格式为 format-rules/{rule_id}/）

    Raises:
        RulePackagePathError: rule_id 不合法
    """
    validate_rule_id(rule_id)
    return Path("format-rules") / rule_id


def validate_rule_package_file(
    target_path: str | Path,
    rule_id: str,
    workspace_root: Optional[Path] = None,
) -> Path:
    """校验规则包内单个文件的路径。

    Args:
        target_path: 目标文件路径
        rule_id: 期望的规则 ID
        workspace_root: 工作区根目录（可选）

    Returns:
        规范化后的 Path 对象

    Raises:
        RulePackagePathError: 路径不符合约束
    """
    validate_rule_id(rule_id)
    path = validate_rule_package_path(target_path, workspace_root)

    # 验证路径的 rule_id 部分与期望一致
    parts = path.parts
    actual_rule_id = parts[1] if len(parts) > 1 else ""
    if actual_rule_id != rule_id:
        raise RulePackagePathError(
            RP_PATH_INVALID,
            f"路径中的 rule_id='{actual_rule_id}' 与期望的 rule_id='{rule_id}' 不一致",
        )

    return path
