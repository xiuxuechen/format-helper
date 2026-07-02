# V5-014 win/mac 平台证据状态

- **日期**：2026-07-02
- **规范基线**：`docs/v5/OFFICECLI_BACKEND_MIGRATION_SPEC.md` §18、§19
- **Gate 脚本**：`scripts/officecli/v5_release_gate.py platform`
- **证据生成脚本**：`scripts/officecli/platform_evidence.py`

## 当前结论

V5-014 已按 format-helper 作为 Codex CLI skill 的实际目标收敛为常见 win/mac 桌面平台 Gate。发布必过 runtime 为 `win-x64`、`osx-arm64`；`osx-x64`、`win-arm64` 与 Linux 系列降级为 best-effort/后补兼容，不再阻塞 V5-014/V5-015。

最新候选 GitHub Actions run `28558124509` 已通过 `platform-contract-validation`、`win-x64` evidence、`osx-arm64` evidence 与 `aggregate-platform-gate`。由于 Gate 已收敛，当前平台 evidence 不再阻塞 V5-014；不同 head_sha 的 PASS 仍不得混合作为同一发布候选。

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
| `win-x64` | passed | `officecli-win-x64-evidence` | GitHub Actions run `28558124509` 上传，artifact digest `sha256:24d515d76346445d02344cdf645a41d5d968b186c930bd5638be7d595af7cf1f`。 |
| `osx-arm64` | passed | `officecli-osx-arm64-evidence` | GitHub Actions run `28558124509` 上传，artifact digest `sha256:be53833531ec3ed3f3328c13131f4e05710129852a5875369241015dbb839ba3`。 |

### best-effort/后补证据

| runtime_id | 状态 | 说明 |
|---|---|---|
| `osx-x64` | best-effort | `macos-13` runner 长期 queued；Intel Mac 证据不再作为 V5-014 发布阻塞项。 |
| `linux-x64-gnu` | passed | GitHub Actions run `28502682372` 曾上传 artifact digest `sha256:49499cc5a87cc315a292e5397757bc9bccc98c356d6c208eb2d6ad2d393c0fed`；作为后补兼容证据记录，不参与发布阻塞。 |
| `win-arm64` | best-effort | 不再作为 V5-014 发布阻塞项。 |
| `linux-arm64-gnu` | best-effort | 不再作为 V5-014 发布阻塞项。 |
| `linux-x64-musl` | best-effort | 不再作为 V5-014 发布阻塞项。 |
| `linux-arm64-musl` | best-effort | 不再作为 V5-014 发布阻塞项。 |

最新候选证据表中的 `win-x64` 与 `osx-arm64` digest 来自 GitHub Actions artifact 元数据，均绑定到 head_sha `e1b0a641c8853bf5d4d81359e0f8c06c40b3350e`。

## 本地修复状态

已在 `scripts/officecli/runtime_resolver.py` 增加下载后安装保护：当 `temp_path.replace(target)` 因 Windows 文件占用失败时，若目标缓存仍能通过锁文件 size 与 SHA-256 校验，则复用合法目标缓存；若目标缓存不存在或校验失败，仍以 `FH-OFFICECLI-DOWNLOAD-FAILED` 阻塞。

该热修仅接受 `PermissionError` 且 `winerror == 32` 的 Windows 文件占用场景；非 WinError 32 的替换失败仍阻塞，避免掩盖权限、磁盘或路径问题。已在 `tests/validation/test_officecli_runtime_resolver.py` 增加回归测试，覆盖“目标文件被占用但缓存合法时复用”、“非文件占用替换失败必须阻塞”和“复用目标缓存后清理当前下载临时文件”。

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

以下旧输出来自八平台 Gate 语义，仅作为历史阻塞记录。收敛后发布 Gate 应改为只要求 `win-x64`、`osx-arm64` 两个平台证据。

```text
python scripts/officecli/v5_release_gate.py platform --evidence-root artifacts/platform-evidence --lock tools/officecli/officecli.lock.json --capability tools/officecli/officecli-capability-manifest.json
{"ok": false, "errors": ["缺少平台证据：linux-arm64-gnu, linux-arm64-musl, linux-x64-musl, osx-x64, win-arm64"]}
```

GitHub Actions run `28558124509` 当前状态：

| job | 状态 | 结论 |
|---|---|---|
| `platform-contract-validation` | completed | success |
| `platform-evidence (osx-arm64, macos-14)` | completed | success |
| `platform-evidence (win-x64, windows-latest)` | completed | success |
| `aggregate-platform-gate` | completed | success |
| `native-toc-evidence (word, officecli-windows-word)` | queued | waiting runner |
| `native-toc-evidence (wps, officecli-windows-wps)` | queued | waiting runner |

`aggregate-platform-gate` 已成功下载并校验 `win-x64`、`osx-arm64` 两个必过 platform evidence artifact；当前剩余阻塞已转移到 Word/WPS native TOC dedicated runner 证据。

## 待补证据

| runtime_id | 推荐执行位置 | 备注 |
|---|---|---|
| 无 | 无 | 必过平台证据已由 run `28558124509` 生成并通过聚合 Gate；后续等待 native-toc-evidence。 |

## 下一步

1. 等待 run `28558124509` 的 `native-toc-evidence (word, officecli-windows-word)` 接单并上传 Word native TOC evidence。
2. 等待 run `28558124509` 的 `native-toc-evidence (wps, officecli-windows-wps)` 接单并上传 WPS native TOC evidence。
3. Word/WPS native TOC Gate 均通过后，同步最终发布 Gate 文档。V5-015 仍需等待 native TOC Gate 通过后才能发布切换。

## 注意事项

- `artifacts/` 属于本地运行产物，当前 `.gitignore` 已忽略；本文件只记录证据状态，不把本地 smoke 产物纳入源码提交。
- 生成目录中若存在旧 smoke 文件或未被当前 evidence JSON 引用的命令产物，不影响 Gate；旧 `*.platform-evidence.json` 仍会被 Gate 递归读取并参与重复、未知、缺失 runtime 校验。删除历史文件需人工确认。
