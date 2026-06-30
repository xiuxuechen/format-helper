# V5-010 Round-06 最终复审结论 — PASS

- **评审对象**：`scripts/officecli/toc_refresh_adapter.py`
- **测试对象**：`tests/validation/test_toc_refresh_adapter_v5.py`
- **复审日期**：2026-06-30
- **规范基线**：`docs/v5/OFFICECLI_BACKEND_MIGRATION_SPEC.md` §21.7、§21.8；`docs/v5/schemas/toc-acceptance.schema.json`
- **复审来源**：主线程实现复核 + 子线程 Gibbs 只读评审
- **结论**：PASS，无未解决 P0/P1/P2

## P1/P2 核销

| 历史问题 | 当前结论 | 证据 |
|---|---|---|
| F-001 状态机阶段缺失 | PASS | `copy_refresh_input`、`open_hidden`、`update_all_fields`、`update_all_tocs`、`repaginate`、`save`、`close_document`、`quit_application`、`reopen_readonly`、`verify_visible_toc`、`officecli_revalidate` 均已进入实现路径。 |
| F-002 成功路径 `error.reason_code` 语义错误 | PASS | 成功路径使用 `code=NONE`、`reason_code=none`、`message=""`；schema 增加 passed/error 组合约束，final acceptance 同步拒绝阻断 reason。 |
| F-003 单阶段 120s 超时未落实 | PASS | CLI 从 probe 开始记录总预算；probe 由 `_run_probe_with_timeout()` 执行 `STAGE_TIMEOUT` watchdog，refresh worker 继承 probe 后剩余的 `TOTAL_TIMEOUT` 总预算。 |
| F-004 `readonly_recommended` 缺失 | PASS | reason code、schema enum 和异常分类均覆盖 `readonly_recommended`。 |
| F-005 transient reason 分类不足 | PASS | timeout 路径映射为可重试 `viewer_busy`，只读推荐等提示类异常已有分类；retryable 仍由 reason_code 纯函数决定。 |
| F-006 PID 追踪/清理机制缺失 | PASS | worker 记录 `worker_pid` 与 Office/WPS `application_pid`；probe 与刷新均使用 `DispatchEx`，只处理本次创建实例；PID 缺失时记录 `cleanup_failed=true`。 |

## 契约与证据

- worker 返回结果使用 `Draft202012Validator + FormatChecker` 校验完整 `toc-acceptance.schema.json`。
- timeout 清理证据写入 `error.message`，未向 `toc-acceptance` 顶层追加 schema 未定义字段。
- worker 状态文件采用同目录临时文件写入后 `replace()`，避免父进程读取半写入 JSON。
- `probe_viewer()` 与刷新路径均禁止 `win32com.client.Dispatch(`，只允许 `DispatchEx(...)` 创建独占 COM 实例。
- CLI `refresh` 入口通过 `_run_probe_with_timeout()` 执行 probe，probe 阶段同样受 `STAGE_TIMEOUT=120` watchdog 约束；随后通过 `_remaining_total_timeout()` 把 probe 已耗时从 refresh worker 的总预算中扣除。
- `Application.AutomationSecurity` 不可用时写入 `toc-refresh-warnings.json` evidence，并通过 `evidence_refs` 挂接到 TOC 验收结果。

## 测试

已通过：

```text
python -m unittest tests.validation.test_toc_refresh_adapter_v5 -v
Ran 27 tests ... OK

python -m unittest tests.validation.test_toc_refresh_adapter_v5 tests.validation.test_final_acceptance tests.validation.test_officecli_platform_evidence tests.validation.test_officecli_v5_release_gate -v
Ran 62 tests ... OK

python -m unittest discover -s tests\validation -p "test_*.py"
Ran 543 tests ... OK

python scripts/officecli/v5_release_gate.py static --root .
{"ok": true, "errors": []}
```

## 残余风险

- 当前本地测试以单元/静态约束为主；真实 Word/WPS COM 行为仍需 Windows dedicated runner 生成 native TOC evidence。
- V5 发布仍受 V5-014 八平台 platform evidence Gate 阻塞；该阻塞不属于 V5-010 代码正确性问题。

## 最终结论

V5-010 在当前复审范围内无未解决 P0/P1/P2，可进入 V5-014 平台证据闭环。
