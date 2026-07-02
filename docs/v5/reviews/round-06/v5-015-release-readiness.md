# V5-015 发布切换准备状态

- **日期**：2026-07-02
- **规范基线**：`docs/v5/OFFICECLI_BACKEND_MIGRATION_SPEC.md` §16、§18、§19
- **当前结论**：V5-015 可以继续做发布切换前置准备，但不得执行正式发布切换。

## 依赖关系

V5-015 在任务矩阵中定义为“发布切换”，依赖 V5-014，DoD 为“无运行时双后端”。最终验收 Gate 还要求 native TOC 未经 Word/WPS 精确刷新不能 accepted。

因此 V5-015 的正式放行必须同时满足：

1. V5-014 win/mac 必过平台证据齐全：`win-x64`、`osx-arm64`。
2. `aggregate-platform-gate` 通过。
3. `native-toc-evidence` 的 Word/WPS dedicated runner 证据通过。
4. 静态生产路径 Gate 通过，确认无运行时双后端。

## 当前阻塞

| Gate | 当前状态 | 说明 |
|---|---|---|
| V5-014 platform-contract-validation | passed | run `28558124509` 已通过。 |
| V5-014 `win-x64` evidence | passed | run `28558124509` 已通过，artifact digest `sha256:24d515d76346445d02344cdf645a41d5d968b186c930bd5638be7d595af7cf1f`。 |
| V5-014 `osx-arm64` evidence | passed | run `28558124509` 已通过，artifact digest `sha256:be53833531ec3ed3f3328c13131f4e05710129852a5875369241015dbb839ba3`。 |
| aggregate-platform-gate | passed | run `28558124509` 已成功校验两项 win/mac platform evidence。 |
| native-toc-evidence | blocked_waiting_runner | run `28558124509` 中 Word/WPS dedicated runner job 已进入 queued，等待 `officecli-windows-word` 与 `officecli-windows-wps` 接单。 |

## 已完成前置证据

本地静态 release gate 已通过：

```text
python scripts/officecli/v5_release_gate.py static --root .
{"ok": true, "errors": []}
```

该结果说明当前生产 Skill 与 `scripts/officecli` 路径未检测到直接 OOXML Python 生产调用，workflow 仍包含 win/mac 平台 Gate 和 Word/WPS native TOC dedicated runner 要求。

## 可以继续推进的 V5-015 前置项

| 项目 | 状态 | 说明 |
|---|---|---|
| 发布切换依赖清单 | ready | 本文件已列出 V5-015 的硬依赖和当前阻塞。 |
| 无双后端静态检查 | ready | `v5_release_gate.py static --root .` 已通过。 |
| 平台 Gate 收敛 | ready | V5-014 已收敛到 `win-x64`、`osx-arm64` 必过；`osx-x64` 为 best-effort。 |
| 平台聚合 Gate | ready | run `28558124509` 的 `aggregate-platform-gate` 已通过。 |
| 发布切换执行 | blocked | 必须等待 native-toc-evidence 全部通过。 |

## 放行条件

只有当以下状态全部满足时，才能将 V5-015 标为可发布切换：

- run `28558124509` 或后续同一 head_sha 候选 run 生成 `win-x64`、`osx-arm64` 两个平台证据。
- `aggregate-platform-gate` 成功校验两平台 evidence。
- Word 与 WPS native TOC evidence 均上传并通过。
- `python scripts/officecli/v5_release_gate.py static --root .` 保持通过。

## 决策

当前决策为：继续推进 V5-015 前置准备；不执行发布切换，不宣布 V5-015 通过。
