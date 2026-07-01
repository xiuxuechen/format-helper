# V5-014 平台证据状态

- **日期**：2026-07-01
- **规范基线**：`docs/v5/OFFICECLI_BACKEND_MIGRATION_SPEC.md` §18、§19
- **Gate 脚本**：`scripts/officecli/v5_release_gate.py platform`
- **证据生成脚本**：`scripts/officecli/platform_evidence.py`

## 当前结论

V5-014 已具备平台证据生成脚本、GitHub Actions 八平台矩阵和聚合 Gate。最新候选 GitHub Actions run `28495182972` 已通过 `platform-contract-validation`，并成功生成 2 个平台证据 artifact；`win-x64` 在缓存替换阶段失败，5 个专用 runner 标签仍未被在线 runner 接单。

当前不得进入 V5-015。下一次候选证据必须基于修复后的新 run 重新生成，不能把失败候选 run 与历史 run 的 PASS 混合为发布证据。

## 已完成证据

### 历史可用证据

| runtime_id | 状态 | 证据路径 | 说明 |
|---|---|---|---|
| `win-x64` | passed | `officecli-win-x64-evidence` | GitHub Actions run `28485992493` 上传，artifact digest `sha256:4729621894fc38609ea1527ba2226b45f7a809454799309404195ba7d0ec7bf4`。 |
| `linux-x64-gnu` | passed | `officecli-linux-x64-gnu-evidence` | GitHub Actions run `28485992493` 上传，artifact digest `sha256:90a72cf504f621a451917575ba1eecc05280947b6c20c793302740b8c4f85e26`。 |
| `osx-arm64` | passed | `officecli-osx-arm64-evidence` | GitHub Actions run `28485992493` 上传，artifact digest `sha256:ab6dfbc4364fb54303d8ba373c20c28c034827c36829584c4eba6f1378256c03`。 |

上述 digest 来自 GitHub Actions artifact 元数据，均绑定到 head_sha `6c76cc0f0786f2b49aec1bb36bf925af6da61720`。

### 最新候选证据

| runtime_id | 状态 | 证据路径 | 说明 |
|---|---|---|---|
| `linux-x64-gnu` | passed | `officecli-linux-x64-gnu-evidence` | GitHub Actions run `28495182972` 上传，artifact digest `sha256:ac1aba1cbbc8eb4cf96afa5faee5ce86dc689f88afaf5525fb1f62afcc9af56e`。 |
| `osx-arm64` | passed | `officecli-osx-arm64-evidence` | GitHub Actions run `28495182972` 上传，artifact digest `sha256:6627d3903ff2eb53c8ada54cca84da7404bfa45bb24240925e688967fa82b5c9`。 |
| `win-x64` | failed | 无 | GitHub Actions job `84459888974` 在 `Generate locked platform evidence` 阶段失败，错误为 Windows 目标缓存 `officecli.exe` 被另一进程占用导致临时下载文件替换失败。 |

上述最新候选 digest 来自 GitHub Actions artifact 元数据，均绑定到 head_sha `8b68f8d1a79272843c0da3befe48904ddfe3ed65`。

## 本地修复状态

已在 `scripts/officecli/runtime_resolver.py` 增加下载后安装保护：当 `temp_path.replace(target)` 因 Windows 文件占用失败时，若目标缓存仍能通过锁文件 size 与 SHA-256 校验，则复用合法目标缓存；若目标缓存不存在或校验失败，仍以 `FH-OFFICECLI-DOWNLOAD-FAILED` 阻塞。

已在 `tests/validation/test_officecli_runtime_resolver.py` 增加回归测试，覆盖“目标文件被占用但缓存合法时复用”的场景。

本地验证：

```text
python -m unittest tests.validation.test_officecli_runtime_resolver -v
Ran 20 tests in 0.078s
OK

python -m unittest tests.validation.test_officecli_runtime_resolver tests.validation.test_officecli_platform_evidence tests.validation.test_officecli_native_toc_evidence tests.validation.test_officecli_v5_release_gate tests.validation.test_repair_plan_v5 -v
Ran 55 tests in 0.671s
OK
```

本机曾用于单平台 smoke 的离线生成命令：

```powershell
python scripts/officecli/platform_evidence.py --workspace-root . --lock tools/officecli/officecli.lock.json --capability tools/officecli/officecli-capability-manifest.json --runtime-id win-x64 --output-dir artifacts/platform-evidence/win-x64 --offline
```

## 当前 Gate 结果

以下输出表示历史证据根曾汇总 3 个成功 artifact 后仍缺 5 个平台；最新候选 run `28495182972` 因 `win-x64` 失败，不能复用为发布 Gate 输入。干净工作区若未下载 artifact，缺失数量会不同。

```text
python scripts/officecli/v5_release_gate.py platform --evidence-root artifacts/platform-evidence --lock tools/officecli/officecli.lock.json --capability tools/officecli/officecli-capability-manifest.json
{"ok": false, "errors": ["缺少平台证据：linux-arm64-gnu, linux-arm64-musl, linux-x64-musl, osx-x64, win-arm64"]}
```

GitHub Actions run `28495182972` 当前状态：

| job | 状态 | 结论 |
|---|---|---|
| `platform-contract-validation` | completed | success |
| `platform-evidence (linux-x64-gnu, ubuntu-latest)` | completed | success |
| `platform-evidence (osx-arm64, macos-14)` | completed | success |
| `platform-evidence (win-x64, windows-latest)` | completed | failure |
| `platform-evidence (linux-x64-musl, officecli-linux-x64-musl)` | queued | runner pending |
| `platform-evidence (win-arm64, officecli-win-arm64)` | queued | runner pending |
| `platform-evidence (linux-arm64-gnu, officecli-linux-arm64-gnu)` | queued | runner pending |
| `platform-evidence (osx-x64, macos-13)` | queued | runner pending |
| `platform-evidence (linux-arm64-musl, officecli-linux-arm64-musl)` | queued | runner pending |

## 待补证据

| runtime_id | 推荐执行位置 | 备注 |
|---|---|---|
| `win-arm64` | `officecli-win-arm64` | job 已 queued，等待自托管 runner 在线并匹配标签。 |
| `linux-arm64-gnu` | `officecli-linux-arm64-gnu` | job 已 queued，等待自托管 runner 在线并匹配标签。 |
| `linux-x64-musl` | `officecli-linux-x64-musl` | job 已 queued，等待自托管 runner 在线并匹配标签；runner 信息需显示 musl 与 Alpine distribution。 |
| `linux-arm64-musl` | `officecli-linux-arm64-musl` | job 已 queued，等待自托管 runner 在线并匹配标签；runner 信息需显示 musl 与 Alpine distribution。 |
| `osx-x64` | `macos-13` | job 已 queued，等待 GitHub-hosted macOS x64 runner 调度；若长期 queued，再检查配额、容量或并发限制。 |

## 下一步

1. 提交并推送 Windows 缓存替换热修，触发新的 `.github/workflows/officecli-v5.yml` 候选 run。
2. 待 8 个 `officecli-*-evidence` artifact 全部上传后，下载或汇总到同一 evidence root。
3. 执行聚合 Gate：

```powershell
python scripts/officecli/v5_release_gate.py platform --evidence-root artifacts/all-platform-evidence --lock tools/officecli/officecli.lock.json --capability tools/officecli/officecli-capability-manifest.json
```

4. 若 5 个专用 runner 标签仍保持 queued，则将其状态固化为 `blocked_waiting_runner`，由 runner 管理侧修正在线机器、标签、并发或配额后重新触发。
5. 平台 Gate 通过后，再执行 native TOC dedicated runner 证据和最终发布 Gate 文档同步。当前不得进入 V5-015 发布切换。

## 注意事项

- `artifacts/` 属于本地运行产物，当前 `.gitignore` 已忽略；本文件只记录证据状态，不把本地 smoke 产物纳入源码提交。
- 生成目录中若存在旧 smoke 文件或未被当前 evidence JSON 引用的命令产物，不影响 Gate；旧 `*.platform-evidence.json` 仍会被 Gate 递归读取并参与重复、未知、缺失 runtime 校验。删除历史文件需人工确认。
