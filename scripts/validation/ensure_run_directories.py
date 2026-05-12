"""运行目录预创建工具。

实现 40_DESIGN_FINAL.md §5.1 的目录预创建流程：
- 在确定 run_id 后立即执行
- 幂等创建（重复执行不得删除或覆盖已有产物）
- 覆盖必需子目录：input, snapshots, semantic, plans, output,
  output/_internal, reports, logs, logs/skill_results, review_results
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

# 错误码（参考 40-§3.4）
FH_RUN_ID_INVALID = "FH-RUN-ID-INVALID"
FH_DIR_CREATE_FAILED = "FH-DIR-CREATE-FAILED"
FH_PATH_ESCAPE = "FH-PATH-ESCAPE"


# run_id 正则（参考 40-§3.4, 41-§3.1）
RUN_ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]{1,80}$")


class RunDirError(Exception):
    """运行目录创建错误。"""

    def __init__(self, code: str, message: str):
        self.code = code
        self.message = message
        super().__init__(f"[{code}] {message}")


# 运行目录必需子目录清单（参考 40-§5.1）
REQUIRED_RUN_DIRS = [
    "input",
    "snapshots",
    "semantic",
    "plans",
    "output",
    "output/_internal",
    "reports",
    "logs",
    "logs/skill_results",
    "review_results",
]


def validate_run_id(run_id: str) -> None:
    """校验 run_id 合法性（参考 40-§3.4, 41-§3.1）。

    Args:
        run_id: 运行 ID

    Raises:
        RunDirError: run_id 不合法
    """
    if not run_id:
        raise RunDirError(
            FH_RUN_ID_INVALID,
            "run_id 不能为空",
        )

    if not RUN_ID_PATTERN.match(run_id):
        raise RunDirError(
            FH_RUN_ID_INVALID,
            f"run_id='{run_id}' 不合法，只能使用 [A-Za-z0-9_-]，长度 1-80",
        )

    if run_id in {".", ".."}:
        raise RunDirError(
            FH_RUN_ID_INVALID,
            f"run_id='{run_id}' 不能是 . 或 ..",
        )


def ensure_run_directories(
    run_id: str,
    workspace_root: Optional[Path] = None,
    dry_run: bool = False,
) -> dict:
    """幂等创建运行目录（参考 40-§5.1）。

    Args:
        run_id: 运行 ID
        workspace_root: 工作区根目录（默认为当前目录）
        dry_run: 只返回预期路径，不实际创建

    Returns:
        字典，包含：
        - run_id: 运行 ID
        - run_dir: 运行目录绝对路径
        - format_rules_dir: format-rules/ 绝对路径
        - created: 本次创建的目录列表
        - existed: 已存在的目录列表

    Raises:
        RunDirError: 创建失败或路径不合法
    """
    validate_run_id(run_id)

    if workspace_root is None:
        workspace_root = Path.cwd()
    else:
        workspace_root = Path(workspace_root).resolve()

    # 计算目录路径
    format_rules_dir = workspace_root / "format-rules"
    run_dir = workspace_root / "format_runs" / run_id

    created = []
    existed = []

    # 预先计算所有目录
    all_dirs = [format_rules_dir, run_dir]
    for sub in REQUIRED_RUN_DIRS:
        all_dirs.append(run_dir / sub)

    # 校验路径不逃逸
    for d in all_dirs:
        try:
            resolved = d.resolve() if d.exists() else (d.parent.resolve() / d.name)
            # 对于不存在的目录，需要检查其父目录的 resolved 路径
            relative_check = resolved if resolved.exists() else d.parent.resolve()
            try:
                relative_check.relative_to(workspace_root)
            except ValueError:
                raise RunDirError(
                    FH_PATH_ESCAPE,
                    f"目录路径逃逸工作区：{d}",
                )
        except (OSError, ValueError) as e:
            if isinstance(e, RunDirError):
                raise
            # 其他路径解析错误
            raise RunDirError(
                FH_PATH_ESCAPE,
                f"路径解析失败：{d}, 错误：{e}",
            )

    # 幂等创建
    for d in all_dirs:
        rel_path = d.relative_to(workspace_root) if workspace_root in d.parents else d.name
        if d.exists():
            existed.append(str(rel_path).replace("\\", "/"))
        else:
            if not dry_run:
                try:
                    d.mkdir(parents=True, exist_ok=True)
                except OSError as e:
                    raise RunDirError(
                        FH_DIR_CREATE_FAILED,
                        f"创建目录失败：{d}, 错误：{e}",
                    )
            created.append(str(rel_path).replace("\\", "/"))

    return {
        "run_id": run_id,
        "run_dir": str(run_dir),
        "format_rules_dir": str(format_rules_dir),
        "created": created,
        "existed": existed,
    }


def scan_legacy_paths(workspace_root: Optional[Path] = None) -> dict:
    """扫描历史路径（format_runs/*/rules、format_rules/）。

    实现 40-§5.2 和 50 §6 风险控制中的"规则路径残留旧目录"检查。

    Args:
        workspace_root: 工作区根目录（默认为当前目录）

    Returns:
        字典，包含：
        - format_runs_rules: 发现的 format_runs/*/rules 路径列表
        - format_rules_legacy: 发现的 format_rules/ 路径（历史别名）
        - total_issues: 总问题数
    """
    if workspace_root is None:
        workspace_root = Path.cwd()
    else:
        workspace_root = Path(workspace_root).resolve()

    issues = {
        "format_runs_rules": [],
        "format_rules_legacy": [],
        "total_issues": 0,
    }

    # 扫描 format_runs/*/rules
    format_runs_root = workspace_root / "format_runs"
    if format_runs_root.exists():
        for run_dir in format_runs_root.iterdir():
            if run_dir.is_dir():
                rules_dir = run_dir / "rules"
                if rules_dir.exists():
                    rel_path = rules_dir.relative_to(workspace_root)
                    issues["format_runs_rules"].append(
                        str(rel_path).replace("\\", "/")
                    )

    # 扫描 format_rules/（历史别名）
    format_rules_legacy = workspace_root / "format_rules"
    if format_rules_legacy.exists() and format_rules_legacy.is_dir():
        issues["format_rules_legacy"].append("format_rules")

    issues["total_issues"] = (
        len(issues["format_runs_rules"]) + len(issues["format_rules_legacy"])
    )

    return issues
