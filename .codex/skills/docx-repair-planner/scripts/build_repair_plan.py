#!/usr/bin/env python3
"""从 semantic_audit.json 生成 v5 repair_plan（draft 或 finalized）。"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[4]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.utils.simple_yaml import write_yaml
from scripts.validation.manual_review_repair import WHITELIST_ACTIONS, validate_repair_plan_v5


TZ = timezone(timedelta(hours=8))

EXECUTION_ORDER = [
    "normalize_styles",
    "apply_page_section_rules",
    "apply_heading_styles",
    "apply_body_styles",
    "apply_table_safe_fixes",
    "toc_content_audit",
    "replace_or_insert_auto_toc",
    "refresh_fields_or_mark_for_update",
    "save_repaired_docx",
]

FORMAT_STRATEGY = {
    "map_heading_native_style": "style-definition",
    "apply_body_style_definition": "style-definition",
    "apply_body_direct_format": "direct-format-override",
    "apply_table_cell_format": "direct-format-override",
    "apply_table_border": "direct-format-override",
    "toc_content_audit": "audit-only",
    "insert_or_replace_toc_field": "toc-field",
}

# V5-006: §9.1 业务动作 → OfficeCLI 后端动作映射
BACKEND_ACTION_MAP: dict[str, dict[str, Any]] = {
    "map_heading_native_style": {
        "command": "set",
        "element_type": "paragraph",
        "prop_keys": ["style", "outlineLvl"],
    },
    "apply_body_style_definition": {
        "command": "set",
        "element_type": "style",
        "prop_keys": [],
        "path_prefix": "/styles",
    },
    "apply_body_direct_format": {
        "command": "set",
        "element_type": "paragraph",
        "prop_keys": [],
    },
    "apply_table_cell_format": {
        "command": "set",
        "element_type": "cell",
        "prop_keys": ["vertical_alignment"],
    },
    "apply_table_border": {
        "command": "set",
        "element_type": "cell",
        "prop_keys": [],
        "border_action": True,
    },
    "toc_content_audit": {
        "command": "query",
        "element_type": "toc",
        "read_only": True,
    },
    "insert_or_replace_toc_field": {
        "command": "add",
        "element_type": "toc",
        "prop_keys": [],
    },
}

# V5-006: §9.3 已知禁止属性及其替代
FORBIDDEN_PROPERTIES: dict[str, str] = {
    "shd.fill": 'shd="clear;XXXXXX"',
    "ind.firstLine": "firstLineIndent",
}
FORBIDDEN_BORDER_KEYS = {"border.top", "border.bottom", "border.left", "border.right"}

# V5-006: §9.1 action_type → 风险分类
ACTION_RISK_CLASS: dict[str, str] = {
    "map_heading_native_style": "L1",
    "apply_body_style_definition": "L1",
    "apply_body_direct_format": "L1",
    "apply_table_cell_format": "L1",
    "apply_table_border": "L2",
    "toc_content_audit": "L1",
    "insert_or_replace_toc_field": "L1",
}


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def canonical_output_docx(source_docx: Path, requested_output: Path, now: datetime) -> Path:
    if requested_output.suffix.lower() == ".docx":
        return requested_output
    timestamp = now.strftime("%Y%m%d%H%M")
    return requested_output / f"{source_docx.stem}{timestamp}.docx"


def canonical_plan_output_path(requested_output: Path, plan_state: str, plan_revision: int) -> Path:
    """根据 CLI --output 和 v5 规范确定 repair_plan 落盘路径。"""
    if plan_state == "draft":
        return requested_output
    output_dir = requested_output if requested_output.suffix == "" else requested_output.parent
    return output_dir / f"repair_plan.finalized.r{plan_revision:03d}.yaml"


def normalize_path(path: Path | str) -> str:
    return str(path).replace("\\", "/")


def artifact_ref(path: Path, kind: str, schema_id: str | None = None, schema_version: str | None = None, artifact_id: str | None = None) -> dict[str, Any]:
    """构造 v5 ArtifactRef。"""
    file_hash = sha256_file(path)
    return {
        "artifact_id": artifact_id or f"{kind}-{file_hash[:12]}",
        "kind": kind,
        "relative_path": normalize_path(path),
        "sha256": file_hash,
        "size_bytes": path.stat().st_size,
        "schema_id": schema_id,
        "schema_version": schema_version,
    }


def load_snapshot_index(snapshot_path: Path) -> dict[str, dict[str, Any]]:
    """从 snapshot v2 构建 node_id → node 索引。"""
    snapshot = load_json(snapshot_path)
    index: dict[str, dict[str, Any]] = {}
    for node in snapshot.get("nodes", []):
        if isinstance(node, dict):
            node_id = node.get("node_id")
            if node_id:
                index[str(node_id)] = node
    return index


def build_target_binding(
    element_id: str | None,
    snapshot_index: dict[str, dict[str, Any]],
) -> dict[str, Any] | None:
    """V5-006: 从 v1 element_id / node_id 构建 target_binding。

    element_id 可能是 v1 格式（如 "p-00001"）或 v2 node_id（如 "N-..."）。
    优先按 node_id 查找，否则按 ordinal 近似匹配。
    """
    if not element_id:
        return None
    # v2 node_id 直接查找
    node = snapshot_index.get(str(element_id))
    if node is not None:
        fingerprint = "0" * 64
        sel = node.get("stable_selector")
        if isinstance(sel, dict):
            fp = sel.get("content_fingerprint")
            if isinstance(fp, str) and len(fp) == 64:
                fingerprint = fp
        return {
            "node_id": node.get("node_id", str(element_id)),
            "path": node.get("officecli_path", ""),
            "fingerprint": fingerprint,
        }
    # v1 element_id 格式已不再支持（V5-005 规范：禁止按数组位置猜测）
    if str(element_id).startswith("p-"):
        return None
    return None


def check_forbidden_attributes(after: dict[str, Any]) -> list[str]:
    """V5-006: §9.3 检查禁止属性。返回阻断原因列表。"""
    reasons: list[str] = []
    for key, alternative in FORBIDDEN_PROPERTIES.items():
        if key in after:
            reasons.append(f"禁止属性 {key}，请使用 {alternative}")
    for key in after:
        if key in FORBIDDEN_BORDER_KEYS:
            reasons.append(f"表格边框属性 {key} 只在 manifest 明确标记 set=true 时允许 L2；否则须走 L3_WRITE")
    return reasons


def assign_risk_class(
    action_type: str,
    after: dict[str, Any],
    backend_config: dict[str, Any] | None,
) -> str:
    """V5-006: §9.1 分配风险分类，§9.3 禁止属性升级到 L3_WRITE。"""
    base_risk = ACTION_RISK_CLASS.get(action_type, "L2")
    if check_forbidden_attributes(after):
        return "L3_WRITE"
    return base_risk


def build_backend_action(
    action_type: str,
    element_id: str | None,
    after: dict[str, Any],
    target_binding: dict[str, Any] | None,
) -> dict[str, Any] | None:
    """V5-006: §9.1 构建 OfficeCLI backend_action。"""
    config = BACKEND_ACTION_MAP.get(action_type)
    if config is None:
        return None
    path = target_binding.get("path", "") if target_binding else ""
    if config.get("path_prefix") and path:
        # 样式级操作：path 指向 /styles/{style_id}
        style_id = after.get("style") or after.get("styleId") or ""
        path = f"{config['path_prefix']}/{style_id}" if style_id else path

    prop_keys = config.get("prop_keys", [])
    if prop_keys:
        properties = {k: after[k] for k in prop_keys if k in after}
    else:
        # 未指定 prop_keys 时取全部 after 标量值（排除嵌套对象/数组）
        properties = {
            k: v for k, v in after.items()
            if not isinstance(v, (dict, list)) and k not in FORBIDDEN_PROPERTIES and k not in FORBIDDEN_BORDER_KEYS
        }

    raw_action = None
    has_forbidden = bool(check_forbidden_attributes(after))
    if config.get("border_action") and has_forbidden:
        # L3_WRITE: raw payload 必须在人工确认流程中填入，此处不生成
        return None

    return {
        "command": config["command"],
        "path": path,
        "element_type": config.get("element_type"),
        "properties": properties,
        "index": None,
        "destination_path": None,
        "raw": raw_action,
    }


def candidate_policy(item: dict[str, Any]) -> tuple[str, list[str]]:
    """判断候选动作是否允许自动修复。"""
    reasons: list[str] = []
    confidence = item.get("confidence")
    risk_level = item.get("risk_level")
    action = item.get("recommended_action") or {}
    action_type = action.get("action_type")
    requested_policy = action.get("auto_fix_policy")

    if requested_policy != "auto-fix":
        reasons.append(f"推荐策略为 {requested_policy or '空'}")
    if action_type not in WHITELIST_ACTIONS:
        reasons.append(f"动作 {action_type or '空'} 不在白名单内")
    if not isinstance(confidence, (int, float)) or confidence < 0.85:
        reasons.append("confidence 低于 0.85")
    if risk_level == "high":
        reasons.append("risk_level 为 high")
    if not item.get("evidence"):
        reasons.append("缺少语义证据")
    if not item.get("issue_id"):
        reasons.append("缺少 source issue")
    return ("manual-review" if reasons else "auto-fix", reasons)


def build_action(
    item: dict[str, Any],
    index: int,
    snapshot_index: dict[str, dict[str, Any]],
) -> tuple[dict[str, Any], dict[str, Any] | None]:
    """构造 v5 修复动作和可选人工确认项。"""
    recommended = item.get("recommended_action") or {}
    action_type = recommended.get("action_type") or "manual_review"
    policy, reasons = candidate_policy(item)
    after = recommended.get("after", item.get("after", {}))
    if not isinstance(after, dict):
        after = {}
    element_id = item.get("element_id")
    target_binding = build_target_binding(element_id, snapshot_index)
    if target_binding is None and element_id:
        # 无法绑定到 v2 snapshot → 阻塞，不生成 backend_action
        policy = "manual-review"
        backend_action = None
    else:
        backend_action = build_backend_action(action_type, element_id, after, target_binding)
    risk_class = assign_risk_class(action_type, after, BACKEND_ACTION_MAP.get(action_type))

    # L3_WRITE 禁止属性检测 — 强制转入 manual-review，不预填 raw/confirmation_ref
    manual_confirmation_ref = None
    forbidden = check_forbidden_attributes(after)
    if forbidden and risk_class == "L3_WRITE":
        policy = "manual-review"
        backend_action = None  # raw payload 由确认流程填入，不在此处伪造

    action: dict[str, Any] = {
        "action_id": f"A{index:03d}",
        "source_issue_ids": [item["issue_id"], *item.get("format_issue_ids", [])],
        "action_type": action_type,
        "format_write_strategy": FORMAT_STRATEGY.get(action_type, "manual-review"),
        "target": {
            "element_id": element_id,
            "expected_role": item.get("expected_role"),
        },
        "risk_class": risk_class,
        "confidence": item.get("confidence"),
        "semantic_evidence": item.get("evidence") or [],
        "before": recommended.get("before", item.get("before", {})),
        "after": after,
        "auto_fix_policy": policy,
        "risk_level": item.get("risk_level"),
        "status": "executable" if policy == "auto-fix" else "pending",
        "execution_status": "executable" if policy == "auto-fix" else "skipped",
        "allowed_by_policy": policy == "auto-fix",
    }
    # schema 要求这些字段必须是对象或 absent；禁止写 null
    if target_binding is not None:
        action["target_binding"] = target_binding
    if backend_action is not None:
        action["backend_action"] = backend_action
    if manual_confirmation_ref is not None:
        action["manual_confirmation_ref"] = manual_confirmation_ref
    if policy == "auto-fix":
        action["policy_match_ref"] = {
            "whitelist_id": f"WL-{action_type}",
            "source_kind": "action_whitelist",
        }

    if policy == "auto-fix":
        return action, None

    manual_item = {
        "item_id": f"M{index:03d}",
        "source_issue_ids": action["source_issue_ids"],
        "element_ref": {
            "element_id": element_id,
            "expected_role": item.get("expected_role"),
        },
        "reason": "；".join(reasons) if reasons else item.get("current_problem", "需要人工确认"),
        "required_decision": f"是否允许执行 {action_type}",
        "default_option": "manual-review",
    }
    return action, manual_item


def build_repair_plan(
    semantic_audit: dict[str, Any],
    source_docx: Path,
    working_docx: Path,
    output_docx: Path,
    snapshot_path: str,
    snapshot_path_abs: Path | None,
    capability_manifest_path: Path | None,
    risk_policy_path: str | None,
    run_id: str,
    plan_state: str,
    rule_id: str,
    rule_version: str,
    now: datetime,
    decision_snapshot: dict[str, Any] | None = None,
    finalized_from_plan_id: str | None = None,
) -> dict[str, Any]:
    """生成 v5 repair_plan 数据结构。"""
    plan_id = f"RP-{run_id}-{now.strftime('%Y%m%d-%H%M%S')}"
    plan_revision = 0 if plan_state == "draft" else _compute_revision(run_id, plan_id, now)

    snapshot_index: dict[str, dict[str, Any]] = {}
    if snapshot_path_abs and snapshot_path_abs.exists():
        snapshot_index = load_snapshot_index(snapshot_path_abs)

    actions: list[dict[str, Any]] = []
    manual_items: list[dict[str, Any]] = []
    for index, item in enumerate(semantic_audit.get("items", []), start=1):
        action, manual_item = build_action(item, index, snapshot_index)
        actions.append(action)
        if manual_item:
            manual_items.append(manual_item)

    # 构建 ArtifactRef
    snapshot_ref = artifact_ref(snapshot_path_abs or Path(snapshot_path), "snapshot",
                                "officecli-document-snapshot", "2.0.0",
                                "before-snapshot") if snapshot_path_abs and snapshot_path_abs.exists() else None
    capability_ref = artifact_ref(capability_manifest_path, "capability",
                                  "officecli-capability-manifest", "1.0.0",
                                  "officecli-capability") if capability_manifest_path and capability_manifest_path.exists() else None
    risk_ref = artifact_ref(Path(risk_policy_path), "evidence", "risk-policy", "1.0.0",
                            "risk-policy") if risk_policy_path and Path(risk_policy_path).exists() else None

    source_audit_paths = [normalize_path(Path(snapshot_path))]
    source_audit_refs = [snapshot_ref] if snapshot_ref else []

    manual_review_required = bool(manual_items)

    return {
        "schema_id": "repair-plan",
        "schema_version": "2.0.0",
        "contract_version": "v5",
        "run_id": run_id,
        "plan_id": plan_id,
        "plan_state": plan_state,
        "plan_revision": plan_revision,
        "execution_backend": "officecli",
        "backend_version": "1.0.113",
        "created_at": now.isoformat(),
        "extensions": {},
        "snapshot_ref": snapshot_ref,
        "capability_manifest_ref": capability_ref,
        "source_audit_paths": source_audit_paths,
        "source_audit_refs": source_audit_refs,
        "risk_policy_path": risk_policy_path or "",
        "risk_policy_ref": risk_ref,
        "manual_review_items_ref": {
            "ref_state": "draft" if plan_state == "draft" else "finalized",
            "path": "plans/manual_review_items.json",
            "pending_count": len(manual_items),
            "blocking_count": len(manual_items),
            "total_count": len(manual_items),
        },
        "decision_snapshot": decision_snapshot,
        "actions": actions,
        "manual_review_required": manual_review_required,
        "execution_order": EXECUTION_ORDER,
        "post_repair": {
            "generate_after_snapshot": True,
            "dispatch_second_round_review": True,
            "required_review_task_ids": ["T01", "T02", "T03", "T04", "T05", "T06"],
        },
        "generated_at": now.isoformat(),
        "finalized_from_plan_id": finalized_from_plan_id,
        "finalized_at": now.isoformat() if plan_state == "finalized" else None,
    }


def _compute_revision(run_id: str, plan_id: str, now: datetime) -> int:
    """确定性派生 plan_revision（1-999）。"""
    seed = f"{run_id}\n{plan_id}\n{now.isoformat()}"
    digest = int(hashlib.sha256(seed.encode()).hexdigest()[:8], 16)
    return (digest % 999) + 1


def main_from_args(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="生成 v5 repair_plan（draft 或 finalized）")
    parser.add_argument("--semantic-audit", required=True, type=Path)
    parser.add_argument("--format-audit", type=Path)
    parser.add_argument("--risk-policy", type=Path)
    parser.add_argument("--snapshot", required=True)
    parser.add_argument("--capability-manifest", type=Path)
    parser.add_argument("--run-id", required=True)
    parser.add_argument("--plan-state", required=True, choices=["draft", "finalized"])
    parser.add_argument("--rule-id", required=True)
    parser.add_argument("--rule-version", default="1.0.0")
    parser.add_argument("--source-docx", required=True, type=Path)
    parser.add_argument("--working-docx", required=True, type=Path)
    parser.add_argument("--output-docx", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--decision-snapshot", type=Path)
    args = parser.parse_args(argv)

    semantic_audit = load_json(args.semantic_audit)

    now = datetime.now(TZ)
    output_docx = canonical_output_docx(args.source_docx, args.output_docx, now)

    snapshot_abs = Path(args.snapshot).resolve() if args.snapshot else None
    capability_abs = args.capability_manifest.resolve() if args.capability_manifest else None
    risk_policy_str = normalize_path(args.risk_policy) if args.risk_policy else None

    decision_snapshot = None
    finalized_from = None
    if args.plan_state == "finalized":
        if args.decision_snapshot and args.decision_snapshot.exists():
            decision_snapshot = load_json(args.decision_snapshot)
        finalized_from = f"RP-{args.run_id}"

    plan = build_repair_plan(
        semantic_audit=semantic_audit,
        source_docx=args.source_docx,
        working_docx=args.working_docx,
        output_docx=output_docx,
        snapshot_path=args.snapshot,
        snapshot_path_abs=snapshot_abs,
        capability_manifest_path=capability_abs,
        risk_policy_path=risk_policy_str,
        run_id=args.run_id,
        plan_state=args.plan_state,
        rule_id=args.rule_id,
        rule_version=args.rule_version,
        now=now,
        decision_snapshot=decision_snapshot,
        finalized_from_plan_id=finalized_from,
    )

    # v5 校验 + 输出
    v5_result = validate_repair_plan_v5(plan)
    if not v5_result.valid:
        for error in v5_result.errors:
            print(error)
        return 1

    # 规范输出路径（§15.4）
    if args.plan_state == "draft":
        plan["plan_revision"] = 0
        out_path = canonical_plan_output_path(args.output, args.plan_state, plan["plan_revision"])
    else:
        revision = plan["plan_revision"]
        out_path = canonical_plan_output_path(args.output, args.plan_state, revision)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    write_yaml(out_path, plan)
    print(out_path)
    return 0


def main() -> int:
    return main_from_args()


if __name__ == "__main__":
    raise SystemExit(main())
