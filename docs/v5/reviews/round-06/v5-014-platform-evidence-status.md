# V5-014 平台证据状态

- **日期**：2026-06-30
- **规范基线**：`docs/v5/OFFICECLI_BACKEND_MIGRATION_SPEC.md` §18、§19
- **Gate 脚本**：`scripts/officecli/v5_release_gate.py platform`
- **证据生成脚本**：`scripts/officecli/platform_evidence.py`

## 当前结论

V5-014 已具备平台证据生成脚本、GitHub Actions 八平台矩阵和聚合 Gate。当前本地已生成并验证 `win-x64` 平台证据；发布 Gate 仍因缺少 7 个远端/异构 runner 证据而阻塞。

## 已完成证据

| runtime_id | 状态 | 证据路径 | 说明 |
|---|---|---|---|
| `win-x64` | passed | `artifacts/platform-evidence/win-x64/win-x64.platform-evidence.json` | 本机 Windows x64 运行 `platform_evidence.py --offline` 生成，包含 version/create/add/get/set/validate/screenshot。 |

本机生成命令：

```powershell
python scripts/officecli/platform_evidence.py --workspace-root . --lock tools/officecli/officecli.lock.json --capability tools/officecli/officecli-capability-manifest.json --runtime-id win-x64 --output-dir artifacts/platform-evidence/win-x64 --offline
```

## 当前 Gate 结果

```text
python scripts/officecli/v5_release_gate.py platform --evidence-root artifacts/platform-evidence --lock tools/officecli/officecli.lock.json --capability tools/officecli/officecli-capability-manifest.json
{"ok": false, "errors": ["缺少平台证据：linux-arm64-gnu, linux-arm64-musl, linux-x64-gnu, linux-x64-musl, osx-arm64, osx-x64, win-arm64"]}
```

## 待补证据

| runtime_id | 推荐执行位置 | 备注 |
|---|---|---|
| `win-arm64` | Windows ARM64 自托管或可复现机器 | 需匹配 lock 中 `officecli-win-arm64.exe`。 |
| `linux-x64-gnu` | `ubuntu-latest` x64 | 可由 GitHub-hosted runner 生成。 |
| `linux-arm64-gnu` | Linux ARM64 runner/container | 需真实 ARM64 runner 或等效可复现环境。 |
| `linux-x64-musl` | Alpine x64 container/runner | runner 信息需显示 musl 与 Alpine distribution。 |
| `linux-arm64-musl` | Alpine ARM64 container/runner | runner 信息需显示 musl 与 Alpine distribution。 |
| `osx-x64` | `macos-13` x64 | 可由 GitHub-hosted runner 生成。 |
| `osx-arm64` | `macos-14` 或 `macos-15` arm64 | 可由 GitHub-hosted runner 生成。 |

## 下一步

1. 在 `.github/workflows/officecli-v5.yml` 触发 `workflow_dispatch`，或在对应 runner 上逐平台执行 `platform_evidence.py`。
2. 下载或汇总所有 `officecli-*-evidence` artifact 到同一 evidence root。
3. 执行聚合 Gate：

```powershell
python scripts/officecli/v5_release_gate.py platform --evidence-root artifacts/all-platform-evidence --lock tools/officecli/officecli.lock.json --capability tools/officecli/officecli-capability-manifest.json
```

4. 平台 Gate 通过后，再执行 native TOC dedicated runner 证据和最终发布 Gate 文档同步。

## 注意事项

- `artifacts/` 属于本地运行产物，当前 `.gitignore` 已忽略；本文件只记录证据状态，不把本地 smoke 产物纳入源码提交。
- 生成目录中若存在旧 smoke 文件或未被当前 evidence JSON 引用的命令产物，不影响 Gate；旧 `*.platform-evidence.json` 仍会被 Gate 递归读取并参与重复、未知、缺失 runtime 校验。删除历史文件需人工确认。
