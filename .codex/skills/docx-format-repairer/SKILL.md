---
name: docx-format-repairer
description: 【OFFICECLI DEPRECATED】legacy DOCX 写回技能，仅供历史 run 只读说明。officecli 写回必须使用 scripts/officecli/request_builder.py 与 runtime_adapter.py。
---

# DOCX Format Repairer（已退役）

## 使用时机

仅在识别到历史 legacy 写回调用时用于返回退役阻断说明；不得执行修复。

## 定位

本技能不再是可执行生产能力。officecli 不允许调用本目录下的旧 Python OOXML 写回脚本。

## 固定路由

1. finalized repair plan 由 `scripts/officecli/request_builder.py` 转换为结构化 batch。
2. `scripts/officecli/runtime_adapter.py` 只修改工作副本。
3. `scripts/officecli/post_write_qa.py` 执行 validate、issues、snapshot v2 和 render。
4. 动态目录由 `scripts/officecli/toc_refresh_adapter.py` 完成 Word/WPS 原生刷新。

## 强制边界

- 禁止调用 `scripts/apply_repair_plan.py` 或 `scripts/optimize_table_pagination.py`。
- 禁止恢复 legacy/officecli 双后端开关。
- legacy 历史 run 仅可读取和生成阻塞说明，不得恢复执行。
- 任何调用本技能执行写回的请求必须返回 `FH-OFFICECLI-LEGACY-BACKEND-RETIRED`。

## 固定输出

任务清单
1. 已识别旧写回技能调用
2. 已拒绝旧 Python OOXML 后端

当前阶段
repair_execution

执行结果
阻塞

交付物
无

阻塞/人工确认
- 错误码：`FH-OFFICECLI-LEGACY-BACKEND-RETIRED`
- 恢复建议：从 finalized plan 重新构建 OfficeCLI execution request

下一步
- 进入 `scripts/officecli/request_builder.py` 与 `runtime_adapter.py`

验收自检
- [x] 未调用旧 OOXML 写回
- [x] 未覆盖原始 DOCX

## 失败分支

所有执行请求均以 `blocked` 结束，并返回 `FH-OFFICECLI-LEGACY-BACKEND-RETIRED`。
