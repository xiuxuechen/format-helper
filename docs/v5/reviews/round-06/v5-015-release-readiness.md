# V5-015 发布切换准备状态

- **日期**：2026-07-06
- **规范基线**：`docs/v5/OFFICECLI_BACKEND_MIGRATION_SPEC.md` §16、§18、§19
- **当前结论**：V5-015 发布切换前复核已通过；可进入正式发布切换人工确认/合并环节，不再等待 dedicated runner 接单。

## 依赖关系

V5-015 在任务矩阵中定义为“发布切换”，依赖 V5-014，DoD 为“无运行时双后端”。最终验收 Gate 还要求 native TOC 未经 Word/WPS 精确刷新不能 accepted。

因此 V5-015 的正式放行必须同时满足：

1. V5-014 win/mac 必过平台证据齐全：`win-x64`、`osx-arm64`。
2. `aggregate-platform-gate` 通过。
3. Word/WPS native TOC evidence 通过；证据可来自 CI dedicated runner，或来自维护者本机并通过 `v5_release_gate.py native-toc` 校验。
4. 静态生产路径 Gate 通过，确认无运行时双后端。

## 当前 Gate 状态

| Gate | 当前状态 | 说明 |
|---|---|---|
| V5-014 platform-contract-validation | passed | run `28761474155` 已通过，head_sha `86e325998fded8bfc16019f815fb446487fc714a`。 |
| V5-014 `win-x64` evidence | passed | run `28761474155` 已通过，artifact digest `sha256:106089cae4d39df91c27f04851b4d90036c570a600a6e88c86fc25a72da7f3f0`。 |
| V5-014 `osx-arm64` evidence | passed | run `28761474155` 已通过，artifact digest `sha256:733bee290161b0c18c20226ed557374cacd44c6c8f2d6deda6dbc85d283523df`。 |
| aggregate-platform-gate | passed | run `28761474155` 已成功校验两项 win/mac platform evidence。 |
| native-toc-evidence 本机 Gate | passed | 本机 Word 12.0 与 WPS Writer 12.0 均已通过，发布候选证据目录 `artifacts/native-toc-evidence-release/` 已通过 `v5_release_gate.py native-toc`。 |
| native-toc-evidence CI dedicated runner | skipped_optional | run `28761474155` 为 PR 触发，Word/WPS dedicated runner job 按 `workflow_dispatch` 条件跳过；该路径为后补自动化，不再阻塞已通过本机 Gate 的发布候选。 |

## 已完成前置证据

本地静态 release gate 已通过：

```text
python scripts/officecli/v5_release_gate.py static --root .
{"ok": true, "errors": []}
```

该结果说明当前生产 Skill 与 `scripts/officecli` 路径未检测到直接 OOXML Python 生产调用，workflow 仍包含 win/mac 平台 Gate、native TOC 后补自动化入口，以及本机 native TOC evidence 的机器校验入口。

## 可以继续推进的 V5-015 前置项

| 项目 | 状态 | 说明 |
|---|---|---|
| 发布切换依赖清单 | ready | 本文件已列出 V5-015 的硬依赖和当前阻塞。 |
| 无双后端静态检查 | ready | `v5_release_gate.py static --root .` 已通过。 |
| 平台 Gate 收敛 | ready | V5-014 已收敛到 `win-x64`、`osx-arm64` 必过；`osx-x64` 为 best-effort。 |
| 平台聚合 Gate | ready | run `28761474155` 的 `aggregate-platform-gate` 已通过。 |
| 本机 Word/WPS native TOC evidence | ready | Word 与 WPS 均已在本机通过，且 `v5_release_gate.py native-toc` 已校验发布候选证据目录。 |
| 发布切换前复核 | passed | 已完成最终发布切换前复核，见 `docs/v5/reviews/round-06/v5-015-final-switch-review.md`。 |
| 发布切换执行 | ready_for_manual_release | 不再等待 runner 接单；下一步是人工确认后执行正式发布切换/合并。 |

## 放行条件

只有当以下状态全部满足时，才能将 V5-015 标为可发布切换：

- run `28761474155` 生成 `win-x64`、`osx-arm64` 两个平台证据，且两者绑定同一 head_sha `86e325998fded8bfc16019f815fb446487fc714a`。
- `aggregate-platform-gate` 成功校验两平台 evidence。
- Word 与 WPS native TOC evidence 均存在并通过 `python scripts/officecli/v5_release_gate.py native-toc --evidence-root artifacts/native-toc-evidence-release --lock tools/officecli/officecli.lock.json`。
- `python scripts/officecli/v5_release_gate.py static --root .` 保持通过。

## 决策

当前决策为：V5-015 发布切换前复核已通过；发布切换不再等待 GitHub self-hosted runner 接单，本机 Word/WPS native TOC evidence 已通过机器 Gate。正式发布切换仍应由维护者在合并/发布动作前进行人工确认。
