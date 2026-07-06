# V5-015 最终发布切换前复核

- **日期**：2026-07-06
- **规范基线**：`docs/v5/OFFICECLI_BACKEND_MIGRATION_SPEC.md` §16、§18、§19
- **复核对象**：PR 候选分支 `codex/v5-officecli-ci-evidence`
- **候选 head_sha**：`86e325998fded8bfc16019f815fb446487fc714a`
- **当前结论**：PASS。V5-015 已满足发布切换前 Gate；正式发布切换/合并前仍需维护者人工确认。

## Gate 复核结果

| Gate | 状态 | 证据 |
|---|---|---|
| win/mac 必过平台证据 | passed | GitHub Actions run `28761474155` 生成 `win-x64` 与 `osx-arm64` 两项 evidence，均绑定 head_sha `86e325998fded8bfc16019f815fb446487fc714a`。 |
| `win-x64` artifact | passed | `officecli-win-x64-evidence` digest `sha256:106089cae4d39df91c27f04851b4d90036c570a600a6e88c86fc25a72da7f3f0`。 |
| `osx-arm64` artifact | passed | `officecli-osx-arm64-evidence` digest `sha256:733bee290161b0c18c20226ed557374cacd44c6c8f2d6deda6dbc85d283523df`。 |
| `aggregate-platform-gate` | passed | run `28761474155` 中 `aggregate-platform-gate` 结论为 success。 |
| dedicated native TOC runner | optional | run `28761474155` 为 PR 触发，`native-toc-evidence` 与 `aggregate-native-toc-gate` 按 `workflow_dispatch` 条件 skipped；后续可通过 `workflow_dispatch` + `run_native_toc_dedicated=true` 补自动化证据。 |
| 本机 Word/WPS native TOC Gate | passed | `python scripts/officecli/v5_release_gate.py native-toc --evidence-root artifacts/native-toc-evidence-release --lock tools/officecli/officecli.lock.json` 输出 `{"ok": true, "errors": []}`。 |
| 无运行时双后端静态 Gate | passed | `python scripts/officecli/v5_release_gate.py static --root .` 输出 `{"ok": true, "errors": []}`。 |
| 回归单测 | passed | 2026-07-06 补跑 `python -m unittest tests.validation.test_officecli_v5_release_gate -v`，共 5 项通过，输出 `Ran 5 tests in 0.379s` / `OK`。 |

## 发布切换判断

V5-015 的发布切换依赖 V5-014。当前候选分支已满足：

- `win-x64` 与 `osx-arm64` 必过平台证据齐全，且来自同一候选 head_sha。
- `aggregate-platform-gate` 已通过，未混用不同 head_sha 的 PASS。
- Word 与 WPS native TOC 已由本机真实 Windows + Word/WPS evidence 覆盖，并通过 `native-toc` 机器 Gate。
- dedicated Windows runner 已保留为后补自动化路径，不再作为唯一放行来源。
- 静态生产路径 Gate 已确认不恢复 v4/v5 双后端，不引入运行时后端切换。

## 剩余事项

| 项目 | 状态 | 说明 |
|---|---|---|
| 正式发布切换/合并 | pending_manual_confirmation | 该动作会改变发布状态，应由维护者人工确认后执行。 |
| dedicated runner 后补证据 | optional | 不阻塞 V5-015；runner 在线后可手动触发 workflow_dispatch 补齐。 |
| 本地无关脏文件 | ignored_for_release | `.gitignore`、`scripts/validation/evidence_manifest.py`、`scripts/validation/final_acceptance.py`、`.claude/settings.local.json` 不属于本次 V5-015 复核提交范围。 |

## 决策

V5-015 最终发布切换前复核结论为 PASS。当前分支可进入人工确认后的正式发布切换/合并环节；若后续发现阻断缺陷，按规范只能回滚到上一完整发布版本，不恢复 v5 运行时双后端。
