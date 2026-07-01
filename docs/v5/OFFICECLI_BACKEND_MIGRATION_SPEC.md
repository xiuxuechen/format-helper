# v5 OfficeCLI 文档执行后端迁移规范

> 文档状态：Approved
>
> 规范版本：v1.0.1
>
> 基线日期：2026-06-15
>
> OfficeCLI 基线：v1.0.113
>
> OfficeCLI 固定提交：`8e5b17977493de1b46536561a50971799c4fc665`
>
> 上游许可证：Apache-2.0
>
> 本文档是 v5 OfficeCLI 接入、实现、迁移、测试和验收的单一事实源。

## 1. 文档权威与目标

### 1.1 权威关系

1. `docs/v4/` 是迁移输入，不是 v5 实现依据。
2. 本文档定义 v5 唯一目标架构。实现与 v4 冲突时，以本文档为准。
3. v5 实施完成后，必须把本文档定义的 Schema、Skill、测试和运行契约落到对应文件；不得只保留本文档而继续执行 v4 后端。
4. OfficeCLI 固定版本的机器帮助 Schema 是 OfficeCLI 自身能力的权威；本文档定义 format-helper 是否允许使用该能力以及如何纳入 Gate。
5. 未在本文档或锁定的 capability manifest 中登记的命令、元素、属性和枚举不得进入生产执行。

### 1.2 目标状态

- 保留 `format-helper` 的编排、语义判断、规则包、风险策略、人工确认、恢复、证据和报告能力。
- OfficeCLI 成为唯一 DOCX 事实读取、结构查询、写回、OpenXML 校验、问题扫描和预览后端。
- Python 只负责平台探测、供应链校验、进程调用、JSON 转换、Schema 校验、日志和错误映射。
- Python 生产路径禁止使用 `zipfile`、`xml.etree`、`lxml`、`python-docx` 或其他方式直接读取或写入 DOCX OOXML。
- 不保留 Python/OfficeCLI 运行时双后端，不提供后端切换开关。
- Windows、Linux、macOS 的 x64/ARM64 首版同时支持；Linux 区分 glibc 和 musl。
- OfficeCLI L1/L2 DOCX 能力全部登记并可被策略层引用；L3 读取默认允许，L3 写入必须人工确认。
- OfficeCLI 预览只承担结构和近似视觉验收；动态目录精确页码必须由 Word 或 WPS 原生刷新后验收。

### 1.3 非目标

- 不把 XLSX、PPTX 纳入 format-helper。
- 不通过 MCP 执行正式自动化流程。
- 不允许 Agent 或 Skill 自由拼接未校验的 OfficeCLI 命令。
- 不保证 OfficeCLI HTML/PNG 与 Microsoft Word 的分页像素一致。
- 不修改 OfficeCLI 上游源码，不维护私有 fork。

## 2. 固定架构

### 2.1 组件职责

| 组件 | 唯一职责 | 禁止事项 |
|---|---|---|
| `format-helper` | 运行编排、Gate、状态、人工确认、最终验收 | 直接拼接 OfficeCLI 命令；直接读写 OOXML |
| 语义与规则层 | 基于 snapshot v2 判断角色、规则和差异 | 调用 OfficeCLI 写回；修改原始 DOCX |
| repair planner | 生成 draft/finalized repair plan 和业务动作 | 生成 shell 字符串；直接执行命令 |
| OfficeCLI request builder | 把 finalized repair plan 转换为 execution request | 自行提高风险等级；绕过人工确认 |
| OfficeCLI runtime adapter | 获取固定二进制、执行 JSON 命令、收集结果 | 解释业务语义；修改计划 |
| OfficeCLI v1.0.113 | DOCX DOM、batch、raw、validate、view | 决定动作是否获准 |
| Word/WPS refresh adapter | 动态目录精确刷新和最终页码证据 | 执行普通格式修复 |
| reviewer/acceptance | after snapshot、差异、证据和终态判定 | 修改执行结果或 final acceptance |

### 2.2 唯一数据流

```text
external source.docx（只读）
  -> 原始文件 hash/size
  -> input/source.docx（只读副本，hash 必须相同）
  -> input/working.docx
  -> OfficeCLI capability/version preflight
  -> officecli-document-snapshot.before.json
  -> 语义/规则/审计
  -> repair_plan.draft.yaml
  -> manual_review_items.json
  -> repair_plan.finalized.rNNN.yaml
  -> officecli-execution-request.rNNN.json
  -> OfficeCLI batch 修改 working 副本
  -> output/_internal/executed.docx
  -> validate + issues + after snapshot + screenshot
  -> Word/WPS TOC refresh（仅需要时）
  -> 规范命名 output/*.docx
  -> review results
  -> evidence manifest
  -> final_acceptance.json
  -> reporting_result.json
```

### 2.3 一次性切换与回退

1. v5 发布包只能包含 OfficeCLI 后端。
2. 切换前必须在发布候选分支完成全部迁移任务和 Gate。
3. 切换后发现阻断缺陷时，回退到上一完整发布版本；不得在 v5 运行时恢复旧 Python 后端。
4. 回退不得继续写入已经由 v5 创建的运行目录。v5 run 必须保持只读证据，重新运行必须创建新 `run_id`。
5. 任何 OfficeCLI 获取、版本、能力或执行失败均进入 `blocked` 或可恢复 `retry`，不得回落到直接 OOXML。

## 3. OfficeCLI 供应链契约

### 3.1 锁文件

标准路径：

```text
tools/officecli/officecli.lock.json
```

锁文件字段：

| 字段 | 类型 | 必填 | 规则 |
|---|---|---:|---|
| `schema_id` | string | 是 | 固定 `officecli-lock` |
| `schema_version` | semver | 是 | 首版 `1.0.0` |
| `officecli_version` | string | 是 | 固定 `1.0.113`，不带 `v` |
| `release_tag` | string | 是 | 固定 `v1.0.113` |
| `source_commit` | sha40 | 是 | 固定 `8e5b17977493de1b46536561a50971799c4fc665` |
| `released_at` | RFC3339 | 是 | 固定 `2026-06-15T05:23:48Z` |
| `license` | string | 是 | 固定 `Apache-2.0` |
| `primary_base_url` | URI | 是 | 固定 GitHub versioned release URL |
| `mirror_base_url` | URI/null | 是 | 首版固定为 null；未经上游可审计证明不得启用 |
| `auto_update_disabled` | bool | 是 | 必须为 `true` |
| `assets` | array | 是 | 恰好包含下表 8 项 |

`assets[]` 字段：

| 字段 | 类型 | 必填 | 规则 |
|---|---|---:|---|
| `runtime_id` | enum | 是 | 下表定义值 |
| `os` | enum | 是 | `windows/linux/macos` |
| `arch` | enum | 是 | `x64/arm64` |
| `libc` | enum/null | 是 | Linux 为 `glibc/musl`，其他为 null |
| `asset_name` | string | 是 | 必须与官方资产精确一致 |
| `sha256` | sha64 | 是 | 必须与下表一致 |
| `size_bytes` | positive integer | 是 | 来自固定 release metadata |
| `primary_url` | URI | 是 | `primary_base_url/asset_name` |
| `mirror_url` | URI/null | 是 | 首版固定为 null |
| `executable_name` | string | 是 | Windows 为 `officecli.exe`，其他为 `officecli` |

### 3.2 固定资产清单

主下载地址统一为：

```text
https://github.com/iOfficeAI/OfficeCLI/releases/download/v1.0.113/{asset_name}
```

| runtime_id | 官方资产 | size_bytes | SHA-256 |
|---|---|---:|---|
| `win-x64` | `officecli-win-x64.exe` | 31997816 | `15d29f3a04e6ad00503de178f98dae872b47ef71f09fac3c614212b209c4d229` |
| `win-arm64` | `officecli-win-arm64.exe` | 32448388 | `94fa5101b94f2fe59c1458688bbc3ddcde4f244afe204143b7eac9bb5089f784` |
| `linux-x64-gnu` | `officecli-linux-x64` | 33950776 | `ffe09f5f8ec76240e44ff431b802b8a4466775afda328f1f7b606e3a79807311` |
| `linux-arm64-gnu` | `officecli-linux-arm64` | 33369562 | `893874471e6830d29580ba9cab0a5834eab80278092f77edb31292bffff1f9fd` |
| `linux-x64-musl` | `officecli-linux-alpine-x64` | 33968418 | `5579d760de781781c7a05e32774bea0bdd091ad3ba3d013129a35e2c837a09be` |
| `linux-arm64-musl` | `officecli-linux-alpine-arm64` | 33410098 | `a18f81e2a4f9cbc8bbec80fc305b20aec1352327094bdff1b48fdc13da3dddba` |
| `osx-x64` | `officecli-mac-x64` | 33330224 | `62ad1b63ec1b833efe01a51d3564238ce274b51a785b1a2fc91880c66381b0d2` |
| `osx-arm64` | `officecli-mac-arm64` | 32587584 | `35a733b598cb32a57d4edc1217a5edfcf63aa9c141916b0b4ef54aa37e4c30ba` |

### 3.3 平台探测

固定映射：

| 系统事实 | runtime_id |
|---|---|
| Windows AMD64/x86_64 | `win-x64` |
| Windows ARM64/aarch64 | `win-arm64` |
| macOS x86_64 | `osx-x64` |
| macOS arm64/aarch64 | `osx-arm64` |
| Linux x86_64 + glibc | `linux-x64-gnu` |
| Linux arm64/aarch64 + glibc | `linux-arm64-gnu` |
| Linux x86_64 + musl/Alpine | `linux-x64-musl` |
| Linux arm64/aarch64 + musl/Alpine | `linux-arm64-musl` |

Linux libc 判定顺序：

1. 存在 `/etc/alpine-release`，判定 musl。
2. `ldd --version` stdout/stderr 包含 `musl`，判定 musl。
3. 否则判定 glibc。

无法映射时返回 `FH-OFFICECLI-PLATFORM-UNSUPPORTED`，不得猜测。

### 3.4 下载、缓存和校验

缓存路径：

```text
.cache/officecli/v1.0.113/{runtime_id}/{executable_name}
```

`.cache/` 必须加入 `.gitignore`。固定算法：

1. 读取并校验锁文件 Schema。
2. 探测 `runtime_id`，精确选择一条资产。
3. 若缓存存在，先校验 size、SHA-256、执行权限和版本；全部通过则直接使用。
4. 获取 `{runtime_id}.lock` 排他文件锁，等待上限 60 秒。锁内容为 `pid/host/started_at/nonce`；持有进程不存在且锁龄超过 10 分钟才视为陈旧锁，删除前记录审计事件。
5. 缓存不存在或校验失败时，下载到同目录 `.download.{pid}.{nonce}`；不得复用固定临时文件名。
6. 首版只访问 `primary_url`。`mirror_url` 为 null；未来启用镜像必须有上游官方证明、固定 HTTPS host allowlist、禁止跨 host 重定向，并继续使用相同资产 hash。
7. 下载后校验 asset size 和 SHA-256。
8. Linux/macOS 设置所有者执行位；Windows 不修改 ACL。
9. 使用环境变量 `OFFICECLI_SKIP_UPDATE=1` 执行 `officecli --version`。
10. 解析版本必须精确等于 `1.0.113`；不接受更高、兼容或 latest。
11. 同目录原子重命名临时文件为正式文件，刷新目录元数据后释放锁。失败时删除当前 nonce 的临时文件；不得删除其他进程文件。

无网规则：

- 仅允许使用已通过本次进程重新计算 SHA-256 和版本检查的缓存。
- 缓存不存在或损坏时返回 `FH-OFFICECLI-OFFLINE-CACHE-MISS`。
- 不读取系统 PATH 中的 OfficeCLI 作为替代。

### 3.5 许可证

- 仓库新增 `THIRD_PARTY_NOTICES/OfficeCLI-Apache-2.0.txt`。
- 发布说明声明 OfficeCLI 为独立第三方可执行程序，版本为 v1.0.113。
- 下载器不得修改二进制。
- 升级 OfficeCLI 时必须重新核对 LICENSE、NOTICE、资产和哈希。

## 4. 能力清单契约

### 4.1 生成方式

标准路径：

```text
tools/officecli/officecli-capability-manifest.json
```

该文件必须由锁定二进制生成，不得手写属性列表。生成步骤：

1. 对 OfficeCLI v1.0.113 固定源码 `schemas/help/docx/*.json` 中的 40 个 help target 逐个执行：

```text
officecli help docx {help_target} --json
```

help target 与返回 JSON 中的 canonical element 不是一一同名：`fieldchar` 返回 `fieldChar`，`instrtext` 返回 `instrText`，`table-cell` 返回 `cell`，`table-column` 返回 `column`，`table-row` 返回 `row`。manifest 必须同时记录 `help_target` 与 `element`，后续业务契约引用 `element`，供应链与 drift 校验引用 `help_target`。

2. 从固定源码 root command 注册表生成全量命令清单，并对每个命令执行 `--help` 或对应 JSON help。首版至少必须覆盖：

```text
create import open close save get query set add remove move swap batch dump
raw raw-set add-part validate view merge refresh watch goto mark plugins mcp skills
```

其中 `add-part/open/close/save/watch/goto/mark/plugins/mcp/skills` 必须进入 deny 或非生产分类；是否出现在 root command 不等于 DOCX 生产可用。

3. 对 view 模式登记并分类：

```text
text annotated outline stats issues html screenshot pdf forms
```

`svg` 仍须登记为全局 view 模式，但固定标记 `docx_supported=false`；它在 v1.0.113 只支持 PPTX。`pdf`、`forms` 必须登记为 DOCX 能力。能力生成器不得把 `officecli help docx --json` 当作元素清单接口，因为 v1.0.113 根级调用输出人类可读列表；生成器以上述固定 40 个 help target 为迭代输入，并逐项校验 JSON。

4. 以 UTF-8、递归 key 排序、数组保持原序进行 canonical JSON 序列化。
5. 对每个原始 help JSON 计算 SHA-256。
6. manifest 记录聚合 hash。CI 每次运行重新生成并比较；任何差异必须阻塞。

### 4.2 manifest 字段

| 字段 | 类型 | 必填 | 规则 |
|---|---|---:|---|
| `schema_id` | string | 是 | `officecli-capability-manifest` |
| `schema_version` | semver | 是 | `1.0.0` |
| `officecli_version` | string | 是 | `1.0.113` |
| `source_commit` | sha40 | 是 | 固定提交 |
| `generated_at` | RFC3339 | 是 | 生成时间 |
| `generator_version` | semver | 是 | 适配器版本 |
| `global_commands` | array | 是 | 命令、参数和可用性 |
| `view_modes` | array | 是 | 全局模式、`docx_supported`、参数和输出类型 |
| `elements` | array | 是 | 恰好 40 项，且每项同时记录 `help_target` 和 canonical `element` |
| `raw_read_allowed` | bool | 是 | `true` |
| `raw_write_policy` | enum | 是 | `manual_confirmation_required` |
| `aggregate_sha256` | sha64 | 是 | canonical elements 内容 hash |

`elements[]` 必须原样保留 OfficeCLI help 中的：

- `element`
- `help_target`
- `elementAliases`
- `operations`
- `allowed_operations`，由 `operations` 中值为 `true` 的键按原序派生
- `properties`
- `children`
- `note`
- 原始 help JSON hash

### 4.3 固定元素覆盖基线

下表是 canonical element completeness Gate；manifest 少一项或多一项均失败。对应的 help target 由 §4.1 固定映射提供。

| 元素 | 允许操作 |
|---|---|
| `abstractNum` | add,set,get,query,remove |
| `body` | get,query |
| `bookmark` | add,set,get,query,remove |
| `chart` | add,set,get,query,remove |
| `chart-axis` | set,get |
| `chart-series` | add,set,get,remove |
| `comment` | add,set,get,query,remove |
| `document` | set,get,query |
| `endnote` | add,set,get,query,remove |
| `equation` | add,set,get,query,remove |
| `field` | add,set,get,query,remove |
| `fieldChar` | set,get,query,remove |
| `footer` | add,set,get,query,remove |
| `footnote` | add,set,get,query,remove |
| `formfield` | add,set,get,query,remove |
| `header` | add,set,get,query,remove |
| `hyperlink` | add,set,get,query,remove |
| `instrText` | set,get,query,remove |
| `level` | add,set,get,remove |
| `num` | add,set,get,query,remove |
| `numbering` | get,query |
| `ole` | add,set,get,query,remove |
| `pagebreak` | add,set,get,query,remove |
| `paragraph` | add,set,get,query,remove |
| `permStart` | add,get,remove |
| `picture` | add,set,get,query,remove |
| `ptab` | add,set,get,query,remove |
| `raw` | 元素级 operations 全部为 false；全局 `raw/raw-set` 命令必须在 `global_commands` 分类并受 L3 策略控制 |
| `revision` | set,get,query |
| `run` | add,set,get,remove |
| `sdt` | add,set,get,query,remove |
| `section` | add,set,get,query,remove |
| `style` | add,set,get,query,remove |
| `styles` | add,get,query |
| `table` | add,set,get,query,remove |
| `cell` | add,set,get,query,remove |
| `column` | add,remove |
| `row` | add,set,get,query,remove |
| `toc` | add,set,get,query,remove |
| `watermark` | add,set,get,query,remove |

“全部接入”的含义是全部进入 manifest、请求 Schema、风险分类和测试覆盖，不代表全部允许自动执行。

## 5. 通用 Schema 规则

### 5.1 版本

- v5 新 Schema 的 `schema_version` 首版统一为 `2.0.0`；供应链锁文件和 capability manifest 可独立使用 `1.0.0`。
- major 不一致：阻塞。
- same-major 且输入 minor 更高：默认阻塞；只有 Schema 明确声明 `extensions` 的位置可承载扩展字段，且扩展字段不得参与 Gate。
- enum、命令、元素类型、操作、风险级别出现未知值：始终阻塞。
- 输出只允许写当前实现支持的精确版本，不得写未来 minor。

### 5.2 路径

- 仓库级路径使用 repo-relative。
- 运行产物使用 run-relative。
- OfficeCLI DOM 路径单独使用 `officecli_path` 类型，不得当作文件路径解析。
- 文件路径在执行前必须 canonicalize 并验证未逃逸 workspace/run。
- OfficeCLI path 必须以 `/` 开头，禁止控制字符、换行和 shell 字符串拼接。

### 5.3 JSON 输出

- 适配器调用 OfficeCLI 时，要求 JSON 的命令必须携带 `--json`。
- stdout 必须只包含一个 UTF-8 JSON 值，允许末尾单个换行。
- stderr 只保存诊断文本，不参与业务成功解析。
- stdout 空、含 BOM 后仍不可解析、包含前后垃圾字符或 JSON envelope 不符合命令契约时返回 `FH-OFFICECLI-NONJSON-OUTPUT`。
- 日志中的 stdout/stderr 单项最大 1 MiB；超出部分写独立内部 artifact，日志记录路径、hash、size 和截断标志。

## 6. `officecli-document-snapshot` v2

标准路径：

```text
snapshots/officecli-document-snapshot.{standard|before|after}.json
```

### 6.1 顶层字段

| 字段 | 类型 | 必填 | 说明 |
|---|---|---:|---|
| `schema_id` | string | 是 | `officecli-document-snapshot` |
| `schema_version` | semver | 是 | `2.0.0` |
| `contract_version` | string | 是 | `v5` |
| `snapshot_id` | string | 是 | 内容派生 ID |
| `kind` | enum | 是 | `standard/before/after/post_toc` |
| `run_id` | string | 是 | 当前运行 |
| `officecli_version` | string | 是 | `1.0.113` |
| `capability_manifest_ref` | ArtifactRef | 是 | manifest path/hash/size |
| `source_docx_ref` | ArtifactRef | 是 | 当前被读取 DOCX |
| `created_at` | RFC3339 | 是 | 创建时间 |
| `document` | object | 是 | 根节点摘要 |
| `nodes` | array | 是 | 规范化 DOM 节点 |
| `parts` | array | 是 | OOXML 部件摘要 |
| `indexes` | object | 是 | 快速定位索引 |
| `warnings` | array | 是 | 非阻断读取警告 |
| `gate_check` | object | 是 | snapshot Gate |

### 6.2 `nodes[]`

| 字段 | 类型 | 必填 | 说明 |
|---|---|---:|---|
| `node_id` | string | 是 | `sha256(snapshot_source_hash + "\n" + officecli_path + "\n" + node_type)` 前 24 hex，加前缀 `N-` |
| `officecli_path` | string | 是 | OfficeCLI 返回的 canonical path |
| `node_type` | enum | 是 | 必须存在于 capability manifest |
| `parent_path` | string/null | 是 | 根节点为 null |
| `ordinal` | integer | 是 | 同父节点中的 0-based 顺序 |
| `part_name` | string | 是 | 例如 `document`、`styles`、`header1` |
| `text` | string/null | 是 | 无文本时 null |
| `text_sha256` | sha64/null | 是 | text 为 null 时 null |
| `attributes` | object | 是 | `get --json` 直接属性，key 保持 OfficeCLI canonical key |
| `effective_format` | object | 是 | 所有 `effective.*` 归一化后的值 |
| `effective_sources` | object | 是 | 每个 effective 属性对应 `.src` |
| `child_paths` | string[] | 是 | OfficeCLI 返回顺序 |
| `stable_selector` | object | 是 | 路径稳定性信息 |
| `raw_evidence_ref` | ArtifactRef/null | 是 | 只有 L2 无法表达且实际调用 raw 时存在 |

`stable_selector` 字段：

| 字段 | 类型 | 规则 |
|---|---|---|
| `kind` | enum | `native_id/semantic_key/positional` |
| `value` | string | 例如 paraId、styleId、bookmark name 或原路径 |
| `rebindable` | bool | positional 必须为 false |
| `content_fingerprint` | sha64 | type、文本、关键属性 canonical hash |

### 6.3 `parts[]`

只允许通过 OfficeCLI `raw` 读取部件元数据，不允许 Python 解包：

| 字段 | 类型 | 说明 |
|---|---|---|
| `part_name` | string | OfficeCLI semantic part name |
| `package_uri` | string/null | OfficeCLI 可返回时记录 |
| `sha256` | sha64 | raw 输出规范化前原始 UTF-8 bytes hash |
| `size_bytes` | integer | raw 输出 bytes |
| `required` | bool | document/styles/numbering/settings 等是否为本次流程必需 |

### 6.4 快照生成算法

1. 所有 OfficeCLI 子进程设置 `OFFICECLI_NO_AUTO_RESIDENT=1`；不得执行 `open`，也不得以 `close` 作为清理前置。
2. 执行 root、body、styles、numbering、header、footer 等 `get --json`。
3. 使用 `dump` 获取可重放结构，作为遍历补充，不直接作为业务快照。
4. 对 manifest 中可查询元素各执行一次完整 `query`，用于 completeness 交叉检查；遍历以 §21.3 BFS 为准。
5. 对每个节点执行深度受控 get；单次 `--depth` 最大 3，避免无界输出。
6. 归并同一 canonical path，检查类型冲突。
7. 计算 stable selector、fingerprint 和索引。
8. 只有缺失业务必需事实时才调用 `raw`；raw 内容写入 `output/_internal/officecli/raw/`，快照只引用。
9. Schema 校验通过后写原子文件。

### 6.5 path 稳定性与失效

- 优先使用 OfficeCLI 返回的 ID selector，如 `p[@paraId=...]`、`revision[@id=...]`。
- style、bookmark、comment 等使用其业务键。
- 纯位置路径标记 `positional`，不得跨结构写操作复用。
- execution request 构建时必须对每个目标重新 `get` 并比较 `node_type` 和 `content_fingerprint`。
- 不一致返回 `FH-OFFICECLI-TARGET-STALE`，整批不得开始。
- batch 内前序结构动作可能改变后序 positional path 时，request builder 必须：
  1. 使用稳定 selector；或
  2. 拆成不同 batch，并在批次间重新 snapshot/rebind；或
  3. 若无法重绑定则阻塞。

## 7. Repair plan v5 扩展

`repair-plan` 升级到 `2.0.0`。保留 v4 的 draft/finalized、decision snapshot、risk-policy hash 和 SelectedAction 约束，新增：

| 字段 | 位置 | 必填 | 说明 |
|---|---|---:|---|
| `execution_backend` | 顶层 | 是 | 固定 `officecli` |
| `backend_version` | 顶层 | 是 | 固定 `1.0.113` |
| `snapshot_ref` | 顶层 | 是 | before snapshot v2 |
| `capability_manifest_ref` | 顶层 | 是 | 固定 manifest |
| `backend_action` | `actions[]` | finalized executable 必填 | 结构化 OfficeCLI 动作 |
| `target_binding` | `actions[]` | 写动作必填 | node_id/path/fingerprint |
| `risk_class` | `actions[]` | 是 | `L1/L2/L3_READ/L3_WRITE` |
| `manual_confirmation_ref` | `actions[]` | L3_WRITE 必填 | 已批准 review item |

`backend_action` 字段：

| 字段 | 类型 | 说明 |
|---|---|---|
| `command` | enum | `set/add/remove/move/swap/raw-set` |
| `path` | string | canonical OfficeCLI path |
| `element_type` | string/null | add 时必填 |
| `properties` | object | 治理层有类型属性；只允许 string/number/boolean/null 标量 |
| `index` | integer/null | add/move 使用，0-based |
| `destination_path` | string/null | move 使用 |
| `raw` | object/null | L3_WRITE 专用 |

`raw` 必填字段：

- `part`
- `xpath`
- `action`：只允许 OfficeCLI 支持的 raw-set action
- `xml`
- `xml_sha256`
- `expected_match_count`：固定为 1
- `precondition_raw_sha256`
- `manual_review_id`
- `decision_snapshot_sha256`

任何 L3_WRITE 动作缺少以上字段必须 Schema 失败。

## 8. `officecli-execution-request`

标准路径：

```text
plans/officecli-execution-request.r{plan_revision}.json
```

### 8.1 顶层字段

本表是 `officecli-execution-request` 顶层完整字段表。

| 字段 | 类型 | 必填 |
|---|---|---:|
| `schema_id`=`officecli-execution-request` | string | 是 |
| `schema_version`=`2.0.0` | semver | 是 |
| `request_id` | string | 是 |
| `run_id` | string | 是 |
| `created_at` | RFC3339 | 是 |
| `extensions` | object | 是 |
| `plan_ref` | ArtifactRef | 是 |
| `plan_sha256` | sha64 | 是 |
| `plan_revision` | string | 是 |
| `working_docx_before_ref` | ArtifactRef | 是 |
| `snapshot_ref` | ArtifactRef | 是 |
| `lock_ref` | ArtifactRef | 是 |
| `capability_manifest_ref` | ArtifactRef | 是 |
| `runtime_id` | enum | 是 |
| `officecli_executable_ref` | ArtifactRef | 是 |
| `environment` | object | 是 |
| `batches` | array | 是 |
| `request_sha256` | sha64 | 是 |
| `gate_check` | GateCheck | 是 |

`environment` 固定包含：

```json
{
  "OFFICECLI_SKIP_UPDATE": "1",
  "OFFICECLI_NO_AUTO_RESIDENT": "1",
  "locale": "C.UTF-8",
  "timezone": "UTC"
}
```

Windows 不支持 `C.UTF-8` 时使用进程 UTF-8 编码设置，但序列化值仍记录 `C.UTF-8` 为逻辑规范。

### 8.2 `batches[]`

| 字段 | 类型 | 规则 |
|---|---|---|
| `batch_id` | string | 唯一 |
| `sequence` | integer | 从 1 连续递增 |
| `atomicity` | enum | 固定 `stop_on_first_error` |
| `timeout_seconds` | integer | 默认 120，范围 1..900 |
| `max_operations` | integer | 固定不超过 12 |
| `preconditions` | array | target rebind、hash、人工确认 |
| `operations` | array | 1..12 |
| `postconditions` | array | get/validate/expected property |

`operations[]` 必填：

- `operation_id`
- `source_action_id`
- `command`
- `path`
- `risk_class`
- `properties`
- `target_binding`
- `expected_result`
- `idempotency_key`

`operations[]` 是 format-helper 治理契约，不得直接作为 OfficeCLI 输入。request builder 必须为每个 batch 额外生成 `officecli_batch_ref` 指向原生 batch JSON。原生文件顶层只能是数组，每项只允许 v1.0.113 `BatchItem` 字段；最小字段映射如下：

| 治理字段 | OfficeCLI 原生字段 | 规则 |
|---|---|---|
| `command` | `command` | 生产 batch 仅允许 `set/add/remove/move/swap/raw-set` |
| `path` | `path` | 原样传递已重绑定 canonical path |
| `element_type` | `type` | `add` 时必填 |
| `properties` | `props` | 所有值先按 manifest 类型校验，再规范化为字符串；对象和数组禁止 |
| `destination_path` | `to` | `move` 时必填 |
| `index` | `index` | 十进制整数 |
| `raw.part` | `part` | `raw-set` 使用 |
| `raw.xpath` | `xpath` | `raw-set` 使用 |
| `raw.action` | `action` | `raw-set` 使用 |
| `raw.xml` | `xml` | `raw-set` 使用 |

治理字段 `operation_id/source_action_id/risk_class/target_binding/expected_result/idempotency_key` 只保留在 execution request，不得写入原生 batch 文件，否则 v1.0.113 会以 unknown field 拒绝。`null` 属性表示不下发该 key，不得序列化为字符串 `"null"`。

生产 batch 禁止 `get/query/view/raw/validate/open/close/save/create/import/refresh/add-part/watch/goto/mark/plugins/mcp/skills`。所有读命令、QA 命令、能力探测和渲染必须按工作流阶段作为独立进程执行。`add` 只允许 `path + type + props + index` 形态；`parent/from/after/before/selector/text/mode/depth` 等 v1.0.113 原生 BatchItem 字段在生产 batch 中全部禁止。

request builder 必须使用 JSON 文件传入：

```text
officecli batch {working_docx} --input {officecli_batch_json} --stop-on-error --json
```

禁止省略 `--stop-on-error`；v1.0.113 默认是继续执行并逐项报告。禁止通过 shell 拼接 `--prop`，禁止 stdin heredoc，禁止 resident 模式。`validate` 禁止放入 batch：v1.0.113 的 batch 内部 validate 输出文本且不能替代独立进程的验收判定，必须在写后阶段单独执行。

### 8.3 幂等

`idempotency_key`：

```text
sha256(
  repair_plan_sha256 + "\n" +
  working_docx_before_sha256 + "\n" +
  sequence + "\n" +
  canonical_operation_json
)
```

- 同 key 已成功执行且当前文件 hash 等于历史 after hash：返回 `already_applied`。
- 同 key 已成功执行但文件 hash 不同：`FH-OFFICECLI-IDEMPOTENCY-CONFLICT`。
- 部分 batch 失败后不得对同一工作副本原地重试；恢复时从 batch 前 checkpoint 副本重建。

## 9. 业务动作映射

### 9.1 当前动作的唯一映射

| action_type | OfficeCLI 映射 | 风险 | 验证 |
|---|---|---|---|
| `map_heading_native_style` | `set paragraph path`，`properties.style={style_id}`；有 outline 要求时附 `outlineLvl` | L1 | get paragraph，style/effective source 相符 |
| `apply_body_style_definition` | `set /styles/{style_id}`，只使用 manifest 允许属性 | L1/L2 | get style + 抽样段落 effective 属性 |
| `apply_body_direct_format` | `set paragraph` 应用段落属性；字符属性按需要再 `set` 子 run，禁止隐式覆盖未列出的 run | L1/L2 | before/after 属性逐项比较 |
| `apply_table_cell_format` | `set cell` 的 cell 属性；段落/run 属性分开操作 | L1/L2 | get cell/paragraph/run + validate |
| `apply_table_border` | 优先 manifest 中合法 cell/table border 属性；命中已知 schema-invalid 属性时禁止 L2，转 L3_WRITE 人工确认 | L2/L3_WRITE | validate + get/raw 证据 |
| `toc_content_audit` | 只读 `query/get/view outline`，不产生写操作 | L1 | 审计结果 |
| `insert_or_replace_toc_field` | 查询并删除旧 toc；`add /body --type toc`；按计划设置 levels/hyperlinks/index；禁止伪造缓存页码 | L1 | get toc 字段结构 + Word/WPS 刷新 |

### 9.2 通用 OfficeCLI 能力映射

新增业务动作不得直接复用 OfficeCLI command 名称。必须先在风险策略中登记业务动作，再映射：

| OfficeCLI command | 默认风险 | 自动执行条件 |
|---|---|---|
| `get/query/view/dump/raw` | L1 或 L3_READ | 只读路径合法、输出受限 |
| `set` | L1/L2 | 属性存在于 manifest 且 action whitelist 精确匹配 |
| `add/remove/move/swap` | L2 | 结构动作白名单、目标稳定、after snapshot 可验证 |
| `batch` | 容器命令 | 仅承载已逐项批准操作 |
| `raw-set` | L3_WRITE | 人工确认和 raw 全字段完整 |
| `add-part` | DOCX 禁止 | v1.0.113 batch 实现仅支持 PPTX part 类型，不得进入 DOCX capability allowlist |
| `validate` | L1 read | 每次写回后强制执行 |

### 9.3 已知 v1.0.113 禁止属性

request builder 必须静态拒绝：

- paragraph `shd.fill`
- paragraph `ind.firstLine`
- table cell `border.top/bottom/left/right`

对应替代：

- `shd="clear;XXXXXX"`
- `firstLineIndent`
- 单元格边框不得假定 `pbdr.*` 可写；只有 manifest 对目标元素和操作明确标记 `set=true` 时才允许 L2，否则 tcBorders/pBdr 一律走 L3_WRITE

其他 capability help 中标记 `add=false` 或 `set=false` 的属性不得用于对应操作。

## 10. L3 策略

### 10.1 L3_READ

- `raw` 默认允许，但必须指定 semantic part/path。
- 原始内容只写 `output/_internal/officecli/raw/`。
- 单次最大 10 MiB，超限阻塞。
- 禁止把 raw XML 原样放入用户报告。

### 10.2 L3_WRITE

必须同时满足：

1. L1/L2 无法表达，且记录 capability evidence。
2. finalized repair plan 中 `risk_class=L3_WRITE`。
3. `manual_review_items.json` 对应项为 `approved` 或 `modified`，`allows_continue=true`。
4. review item 展示 part、XPath、action、XML 摘要、完整 XML artifact hash、固定 expected match count 1 和风险。
5. 执行前 raw hash 等于 `precondition_raw_sha256`。
6. XPath 通过下述 `single_node_xpath_v1` 语法校验。
7. XML hash 等于 `xml_sha256`。
8. 执行后 validate 通过，并生成 raw before/after evidence。

任一不满足返回 `DFR-OFFICECLI-L3-NOT-AUTHORIZED`。

`single_node_xpath_v1` 是 v1.0.113 下唯一允许写入的 XPath 子集：

```text
^/[A-Za-z_][A-Za-z0-9_.-]*:[A-Za-z_][A-Za-z0-9_.-]*\[1\](?:/[A-Za-z_][A-Za-z0-9_.-]*:[A-Za-z_][A-Za-z0-9_.-]*\[[1-9][0-9]*\])*$
```

- 必须是绝对 XPath，根节点和每一级子节点都带正整数位置谓词。
- 禁止 `//`、`*`、属性谓词、函数、union、轴、相对路径和命名空间声明。
- 该语法最多匹配一个元素。OfficeCLI v1.0.113 对零匹配抛出 `raw-set: XPath matched no elements` 并返回失败，因此成功等价于恰好匹配一个目标。
- 每个 L3_WRITE action 单独占一个 batch。多个节点必须拆成多个 action、XPath 和确认引用。
- 适配器只校验 XPath 字符串语法和 OfficeCLI 结果，不解析 XML，不引入第二个 OOXML/XPath 后端。

## 11. `officecli-execution-result`

标准路径：

```text
logs/officecli-execution-result.r{plan_revision}.json
```

顶层完整字段：

| 字段 | 说明 |
|---|---|
| `schema_id/schema_version` | `officecli-execution-result` / `2.0.0` |
| `result_id/run_id/request_ref` | 身份与输入引用 |
| `created_at/extensions` | 公共字段 |
| `officecli_version/runtime_id/executable_sha256` | 执行环境 |
| `started_at/finished_at/duration_ms` | 时间 |
| `status` | `done/failed/blocked` |
| `working_docx_before_ref` | 执行前副本 |
| `working_docx_after_ref` | 成功时必填 |
| `batch_results` | 每批结果 |
| `failed_batch_id/failed_operation_id` | 失败时条件必填 |
| `retryable` | 是否允许从 checkpoint 重建后重试 |
| `error` | 统一错误 |
| `stdout_artifacts/stderr_artifacts` | 输出证据 |
| `gate_check` | 执行 Gate |

`batch_results[].operation_results[]`：

- `operation_id`
- `source_action_id`
- `status`：`executed/already_applied/failed/not_run`
- `native_success`
- `native_output`
- `native_error`
- `before_target_fingerprint`
- `after_target_fingerprint`
- `postconditions_passed`
- `duration_ms`

### 11.1 退出与部分执行

- exit code 0 但 JSON `success=false`：失败。
- exit code 非 0：失败。
- 适配器同时检查外层 `success` 与 `data.results[].success`，任一为 false 即失败。
- 使用 `--stop-on-error` 时，OfficeCLI 只返回已尝试的结果，不保证为后续项生成记录。适配器必须按 execution request 的 operation 顺序和 OfficeCLI `results[].index` 对齐；首个失败或最后一个返回索引之后的所有请求项由适配器合成 `not_run`，不得伪造 exit code、stdout 或 after fingerprint。
- 返回结果存在重复索引、越界索引、缺口位于首个失败之前或结果命令与原生 batch 项不一致时，返回 `FH-OFFICECLI-RESULT-MISMATCH` 并阻塞。
- OfficeCLI batch 不是文件级事务保证。每批前必须复制 checkpoint：

```text
output/_internal/checkpoints/{batch_id}.before.docx
```

- batch 失败时删除失败后的工作副本引用，从 checkpoint 重建下一次尝试。
- checkpoint 只可由当前 run 使用，不得作为用户交付物。

## 12. 完整工作流

### 12.1 preflight

1. 确认环境、文本编码和运行目录。
2. 校验 lock、license notice、capability manifest。
3. 探测平台、获取并验证二进制。
4. 执行 `--version`。
5. 随机抽取不少于 5 个 help target 重新取 JSON 并核对 hash；CI 使用全部 40 项。
6. 失败时 stage=`environment_preflight`，不得复制或修改 DOCX。

### 12.2 输入保护

1. 计算原始 DOCX hash/size。
2. 原子复制到 `input/source.docx`，校验 hash 等于原始 hash，并标记只读。
3. 从 `input/source.docx` 原子复制到 `input/working.docx`，校验 hash 等于 source hash；working 是唯一可写执行副本。
4. 后续 OfficeCLI 只能收到 `input/working.docx`、`output/_internal/checkpoints/*.docx`、`output/_internal/executed.docx`、`output/_internal/toc-refresh.docx` 或最终输出路径。
5. 命令参数出现外部原件路径或 `input/source.docx` 时立即阻塞。

### 12.3 快照与建规/审计

1. 生成 standard/before snapshot v2。
2. 语义层只消费 snapshot v2，不读取 DOCX。
3. rule slot 证据引用 `node_id + officecli_path + fingerprint`。
4. 规则包记录 `snapshot_schema_version=2.0.0`。
5. 未识别 node type 必须人工确认或阻塞，不得丢弃。

### 12.4 计划与执行

1. planner 生成 draft。
2. 主控汇总人工确认。
3. planner 生成 revisioned finalized plan。
4. request builder 校验 plan、policy、manifest 和目标 freshness。
5. 写 execution request。
6. adapter 设置 `OFFICECLI_NO_AUTO_RESIDENT=1`，按 batch 顺序执行，不调用 `open/save/close`。
7. 写 execution result 和升级后的 repair execution log。

### 12.5 写后校验

固定顺序：

1. 确认当前及所有前序 OfficeCLI 进程均设置 `OFFICECLI_NO_AUTO_RESIDENT=1`；不执行 `close`。若探测到工作副本已被 resident 持有，返回 `FH-OFFICECLI-RESIDENT-CONFLICT`，不得主动关闭不属于当前 run 的会话。
2. `officecli validate {docx} --json`。
3. `officecli view {docx} issues --json`。
4. 生成 after snapshot v2。
5. 按 §21.5 捕获 `officecli view {docx} html` 的非 JSON stdout 并原子写入 `{internal_html}`。
6. 执行 `officecli view {docx} stats --page-count --json` 获取 `data.pages`。页数缺失、非正整数或超过 500 时阻塞；不得猜测页数。v1.0.113 在 Windows 优先使用 Word repagination，其他环境使用 HTML DOM fallback；两者都不可用时命令失败。
7. 对 `N=1..data.pages` 逐页执行 `officecli view {docx} screenshot --page {N} --render html --out {render_dir}/page-{N:0000}.png`。每次必须 exit code 0、输出文件存在且 PNG hash 已记录。Windows 可额外生成 `--render native` 对照证据，但 OfficeCLI 自动验收基线固定为 `html`，避免平台差异。
8. `view pdf` 仅作为 capability 分类项：v1.0.113 依赖 exporter plugin，不纳入基础验收；未安装插件不得影响 screenshot Gate。`view forms --json` 对包含表单字段的 fixture 必须纳入能力测试。
9. 对 repair plan 每个 executable action 生成 review result。
10. 任一 blocking issue、validate error、动作未覆盖、页证据缺失或 after target 不匹配均阻塞。

### 12.6 动态目录

触发条件：

- repair plan 包含 TOC 动作；或
- 规则要求 native TOC；或
- 文档已有 TOC 且标题或分页相关结构被修改。

策略：

1. OfficeCLI 只创建/校验 TOC 字段结构。
2. 若 `toc_mode=native_toc`，必须调用 Word 或 WPS 原生刷新适配器。
3. 刷新必须在 `output/_internal` 副本上进行，成功后再复制为规范命名交付物。
4. 刷新后重新执行 OfficeCLI validate、snapshot、issues 和 screenshot。
5. 记录 viewer 名称、版本、平台、刷新时间、刷新前后 hash、可见目录文本和页码证据。
6. Word/WPS 均不可用、刷新失败、仍含占位文本或页码证据不一致时 `toc_acceptance=blocked`。
7. `equivalent_visible_toc` 只允许规则或人工确认明确选择；不能作为 native TOC 失败后的自动降级。
8. `not_required` 必须有规则/人工确认来源。

### 12.7 终态

- final acceptance 必须引用 lock、capability manifest、before/after snapshot、execution request/result、repair log、review、TOC acceptance 和原件保护证据。
- final acceptance 生成后不可变。
- 报告后置写 `reporting_result.json`。

## 13. v5 运行目录增量

```text
format_runs/{run_id}/
├── input/
│   ├── source.docx
│   └── working.docx
├── snapshots/
│   ├── officecli-document-snapshot.before.json
│   └── officecli-document-snapshot.after.json
├── plans/
│   ├── repair_plan.draft.yaml
│   ├── repair_plan.finalized.rNNN.yaml
│   └── officecli-execution-request.rNNN.json
├── logs/
│   ├── officecli-preflight.json
│   ├── officecli-execution-result.rNNN.json
│   ├── repair_execution_log.json
│   ├── toc_acceptance.json
│   └── final_acceptance.json
├── output/
│   ├── _internal/
│   │   ├── checkpoints/
│   │   ├── officecli/raw/
│   │   ├── preview/
│   │   └── executed.docx
│   └── {原文件名}{yyyyMMddHHmm}(_rNN)?.docx
└── review_results/
```

## 14. 错误码与重试

| 错误码 | 含义 | retryable |
|---|---|---:|
| `FH-OFFICECLI-LOCK-INVALID` | 锁文件无效 | 否 |
| `FH-OFFICECLI-PLATFORM-UNSUPPORTED` | 平台不支持 | 否 |
| `FH-OFFICECLI-DOWNLOAD-FAILED` | 固定主源不可用 | 是 |
| `FH-OFFICECLI-HASH-MISMATCH` | 二进制哈希不符 | 否 |
| `FH-OFFICECLI-VERSION-MISMATCH` | 版本不符 | 否 |
| `FH-OFFICECLI-OFFLINE-CACHE-MISS` | 离线缓存缺失 | 否 |
| `FH-OFFICECLI-CAPABILITY-DRIFT` | help 与 manifest 不一致 | 否 |
| `FH-OFFICECLI-NONJSON-OUTPUT` | stdout 非契约 JSON | 否 |
| `FH-OFFICECLI-RESULT-MISMATCH` | batch 返回索引或内容无法对齐 | 否 |
| `FH-OFFICECLI-RENDERER-UNAVAILABLE` | 无可用浏览器渲染后端 | 否 |
| `FH-OFFICECLI-SNAPSHOT-LIMIT` | snapshot 超过资源上限 | 否 |
| `FH-OFFICECLI-RESIDENT-CONFLICT` | 文件被 resident 持有 | 否 |
| `FH-OFFICECLI-TIMEOUT` | 进程超时 | 是，从 checkpoint |
| `FH-OFFICECLI-TARGET-STALE` | 目标 path/fingerprint 失效 | 是，需重新计划 |
| `FH-OFFICECLI-IDEMPOTENCY-CONFLICT` | 同 key 不同状态 | 否 |
| `DFR-OFFICECLI-BATCH-FAILED` | batch 失败 | 是，从 checkpoint |
| `DFR-OFFICECLI-L3-NOT-AUTHORIZED` | L3 写未获授权 | 否 |
| `DFR-OFFICECLI-VALIDATE-FAILED` | OpenXML 校验失败 | 否 |
| `DFR-OFFICECLI-POSTCONDITION-FAILED` | 写后属性不符 | 否 |
| `DFR-TOC-NATIVE-REFRESH-UNAVAILABLE` | Word/WPS 不可用 | 否 |
| `DFR-TOC-NATIVE-REFRESH-FAILED` | 原生刷新失败 | 由 §21.8 reason_code 固定决定 |

任何 retry 必须创建新的 skill result attempt，并保持原失败 result 不可变。

## 15. v4 迁移

### 15.1 历史运行

- v4 运行目录只读。
- v5 `resume` 遇到 v4 state/snapshot/repair plan 时返回 `manual_recover`。
- 不允许把 v4 positional `p-00001` 静默转换为 OfficeCLI path。
- 用户选择迁移时创建新 run，重新复制原始文件并生成 snapshot v2。
- v4 规则包可作为语义规则输入，但必须通过 v5 repackage，重新绑定 slot evidence。

### 15.2 退役代码

以下生产职责必须删除或改写：

- `scripts/ooxml/extract_docx_snapshot.py`：由 OfficeCLI snapshot adapter 替代。
- `.codex/skills/docx-format-repairer/scripts/apply_repair_plan.py`：由 request builder + runtime adapter 替代。
- `.codex/skills/docx-format-repairer/scripts/optimize_table_pagination.py`：可表达动作映射到 L1/L2，不能表达的进入 L3_WRITE。
- 所有 Python `zipfile`/XML 直接 DOCX 代码。

允许保留历史代码只读归档的前提：

- 不被任何 Skill、入口、测试运行路径 import 或调用。
- 文件头明确 `v4 historical reference; non-executable`。
- CI 静态扫描确认生产路径无直接 OOXML import。

### 15.3 必须更新的契约

- `format-helper` 与全部内部 `docx-*` Skill。
- snapshot、repair plan、repair execution log、review result、evidence manifest、final acceptance Schema。
- Schema examples、validator、Gate predicates、coverage matrix。
- v5 API spec、设计、开发计划、测试计划和进度日志。
- `.gitignore`、第三方许可证、下载器和平台 CI。

### 15.4 兼容别名

- v5 新写入只允许 revisioned finalized plan。
- `repair_plan.yaml`、`repair_plan.finalized.yaml` 只允许 v4 历史识别，不允许 v5 执行。
- 发现兼容别名时必须返回迁移提示，不得静默解析后执行。

### 15.5 v5 resume 状态机

`repair_execution_log.current_status` 枚举固定为：

| 状态 | 含义 | 允许 resume |
|---|---|---:|
| `preflight_pending` | 尚未复制输入 | 否 |
| `snapshot_ready` | before snapshot 已完成 | 是，继续计划 |
| `plan_finalized` | finalized plan 已存在 | 是，构建 request |
| `execution_in_progress` | batch 执行中断 | 是，仅从最后完整 checkpoint 新 attempt |
| `execution_failed_retryable` | 可恢复执行失败 | 是，从 checkpoint 新 attempt，最多一次 |
| `execution_failed_final` | 不可恢复执行失败 | 否 |
| `executed_ready` | `output/_internal/executed.docx` 已生成 | 是，进入写后 QA 或 TOC |
| `blocked_waiting_native_toc` | 非 Windows 已完成 executed，等待 Windows runner 原生 TOC 刷新 | 是，仅 Windows runner |
| `toc_refresh_in_progress` | Windows runner 正在刷新 | 是，前一 attempt 超时后按 TOC checkpoint 重试 |
| `toc_refresh_failed_final` | TOC 刷新不可恢复失败 | 否 |
| `review_ready` | after snapshot 和 QA 完成 | 是，生成 review/evidence |
| `accepted` | final acceptance 已 accepted | 否，幂等返回已完成 |
| `blocked` | Gate 阻塞 | 否，除非生成新 plan revision |

resume 命令固定入口：

```text
python -m scripts.officecli.runtime_adapter resume --run-dir {run_dir}
```

resume 前置校验：

1. `run_dir` 必须位于当前 workspace。
2. `logs/repair_execution_log.json`、最新 request/result、checkpoint、snapshot、plan 和 lock 均通过 ArtifactRef hash 校验。
3. `officecli_version`、lock hash、capability manifest hash 与原 run 一致。
4. 若状态为 `blocked_waiting_native_toc`，当前平台必须是 Windows，且 `output/_internal/executed.docx` hash 等于记录值。
5. resume 不得重新执行已成功 batch，不得修改 `input/source.docx`、历史 result、历史 checkpoint。
6. 新 attempt 必须追加到 `attempts[]`，使用递增 attempt_no 和新的 result_ref；旧失败 result 不可变。
7. 跨平台 TOC 接力只允许从 `executed_ready` 或 `blocked_waiting_native_toc` 进入，不允许在非 Windows 继续 native TOC，也不允许改变 run_id。

禁止事项：

- v4 run resume 不进入此状态机，仍返回 `manual_recover`。
- hash 不一致、lock 漂移、manifest 漂移、缺少 checkpoint、TOC 输入缺失时不得自动修复，返回 `FH-OFFICECLI-RESUME-INTEGRITY-FAILED`。
- 用户要求“继续”但当前状态不可 resume 时，只能生成阻塞报告或新 run 建议，不得猜测执行点。

## 16. 实施任务

| ID | 任务 | 依赖 | 输入 | 输出 | DoD |
|---|---|---|---|---|---|
| V5-001 | 供应链锁定 | 无 | 本文档 §3 | lock Schema、锁文件、license notice | 8 资产校验通过 |
| V5-002 | 跨平台 runtime resolver | V5-001 | lock | 下载器、缓存器、preflight | 8 runtime 测试通过 |
| V5-003 | capability manifest | V5-002 | 固定二进制 help | manifest 生成器和固定 manifest | 40 个 help target、canonical 元素、全局命令无漂移 |
| V5-004 | snapshot v2 Schema | V5-003 | OfficeCLI get/query/dump/raw | Schema、adapter、fixtures | 语义所需事实全覆盖 |
| V5-005 | 语义/规则迁移 | V5-004 | snapshot v2 | strategist、packager 更新 | 不再消费 v1 snapshot |
| V5-006 | repair plan v5 | V5-003,V5-005 | audit、policy | v5 plan Schema/planner | L1/L2/L3 风险闭合 |
| V5-007 | execution request builder | V5-006 | finalized plan | request Schema、builder | 映射确定且无 shell 拼接 |
| V5-008 | OfficeCLI runtime adapter | V5-002,V5-007 | request | result Schema、adapter、checkpoint | 失败可恢复，日志完整 |
| V5-009 | 写后 QA | V5-004,V5-008 | executed docx | validate/issues/render/review | blocking 缺陷阻断 |
| V5-010 | Word/WPS TOC adapter | V5-009 | TOC 文档 | toc evidence | 精确刷新分支全覆盖 |
| V5-011 | 验收与报告迁移 | V5-009,V5-010 | 全证据 | final acceptance/report | 终态引用完整 |
| V5-012 | Skill/playbook 更新 | V5-004..011 | 新契约 | `.codex/skills` | 用户流程无旧后端 |
| V5-013 | 旧后端退役 | V5-004,V5-008 | 替代实现 | 删除/归档、静态扫描 | 生产路径无直接 OOXML |
| V5-014 | win/mac 平台 CI 与全回归 | V5-001..013 | 全实现 | 测试矩阵、证据包 | L1 全绿，win/mac 必过平台证据齐全 |

失败回滚固定规则：

| 任务范围 | 回滚步骤 |
|---|---|
| V5-001..003 | 删除当前任务生成的未发布临时文件；保留失败证据；不得改 active lock/manifest |
| V5-004..007 | 回退当前 feature 分支的 v5 Schema/adapter 提交；v4 仅作迁移输入，不启用混合运行 |
| V5-008..011 | 保留 run、checkpoint 和失败 result 为只读；删除交付物引用；从上一成功任务重新实施 |
| V5-012..013 | 若任一入口仍引用旧后端，整次发布候选失败并回退到上一发布版本，不恢复 v5 双后端 |
| V5-014 | 不发布；保留平台证据；修复后重新生成完整 win/mac 候选证据，不复用失败平台的 PASS |
| V5-015 | 发布切换 | V5-014 | RC | v5 release | 无运行时双后端 |

每项失败回滚：撤销该任务尚未合并的提交，保留测试证据；不得启用临时生产 fallback。V5-015 后回滚只能部署上一完整版本。

## 17. 测试矩阵

### 17.1 供应链与平台

| 测试 | 场景 | 期望 |
|---|---|---|
| V5-T001..008 | 8 runtime 下载 | 资产、size、hash、version 全通过 |
| V5-T009 | 主源网络/5xx/超时失败 | `DOWNLOAD-FAILED`，不访问镜像 |
| V5-T010 | 主源 hash 不符 | 阻塞，供应链错误 |
| V5-T011 | 离线有效缓存 | 成功 |
| V5-T012 | 离线缓存缺失/损坏 | 阻塞 |
| V5-T013 | 系统 PATH 有其他版本 | 忽略 PATH |
| V5-T014 | 自动更新环境 | `OFFICECLI_SKIP_UPDATE=1` 生效 |

### 17.2 能力与快照

| 测试 | 场景 | 期望 |
|---|---|---|
| V5-T020 | 40 个 help target | manifest 完整 |
| V5-T021 | help hash 漂移 | CI 阻塞 |
| V5-T022 | L1 视图 | text/annotated/outline/stats/issues/html/screenshot 可用 |
| V5-T022A | DOCX forms 视图 | `view forms --json` 返回结构化结果 |
| V5-T022B | PDF exporter 未安装 | capability 已分类，基础 Gate 不误失败 |
| V5-T023 | L2 操作 | manifest 中每种允许 operation 至少一例 |
| V5-T024 | L3 raw read | evidence 可追溯 |
| V5-T025 | snapshot v2 正向 | Schema 通过 |
| V5-T026 | 缺 required | Schema 失败 |
| V5-T027 | 未知 node type | 阻塞 |
| V5-T028 | positional path 结构漂移 | stale |
| V5-T029 | 1000 页/大表/多图片 | 有界内存和超时策略生效 |

### 17.3 计划与执行

| 测试 | 场景 | 期望 |
|---|---|---|
| V5-T040..046 | 7 个现有白名单动作 | 映射、回读、review 通过 |
| V5-T047 | 非 manifest 属性 | request Schema/validator 拒绝 |
| V5-T048 | 已知 invalid 属性 | 静态拒绝 |
| V5-T049 | batch 第 3 项失败 | 后续 not_run、checkpoint 保留 |
| V5-T049A | 默认 batch 未传 stop 参数 | 静态测试拒绝；生产命令必须含 `--stop-on-error` |
| V5-T049B | batch 返回索引缺口/重复/越界 | `RESULT-MISMATCH` 阻塞 |
| V5-T050 | timeout | 终止进程树，从 checkpoint 可恢复 |
| V5-T051 | stdout 非 JSON | 阻塞 |
| V5-T052 | exit 0 + success false | 失败 |
| V5-T053 | 损坏 DOCX | validate 阻塞 |
| V5-T054 | 重复 idempotency key | already_applied |
| V5-T055 | key 相同但文件不同 | conflict |

### 17.4 L3、安全与原件

| 测试 | 场景 | 期望 |
|---|---|---|
| V5-T060 | L3 无确认 | 阻塞 |
| V5-T061 | XML hash 不符 | 阻塞 |
| V5-T062 | XPath match count 不符 | 阻塞 |
| V5-T062A | L3 XPath 含 `//`、属性谓词或无位置索引 | Schema 阻塞 |
| V5-T062B | single-node XPath 零匹配 | OfficeCLI 失败，工作副本从 checkpoint 恢复 |
| V5-T063 | precondition raw hash 不符 | stale |
| V5-T064 | L3 成功 | before/after raw + validate |
| V5-T065 | 命令目标为 source.docx | 阻塞 |
| V5-T066 | 完整修复后原始 hash | 保持不变 |
| V5-T067 | 路径逃逸/symlink/junction | 阻塞 |

### 17.5 渲染、TOC、恢复

| 测试 | 场景 | 期望 |
|---|---|---|
| V5-T080 | validate/issues/html/screenshot | 证据完整 |
| V5-T081 | 多页 screenshot | 每页可定位 |
| V5-T081A | 无浏览器后端 | preflight 阻塞 |
| V5-T081B | HTML stdout 空、超限或非 UTF-8 | 阻塞且不生成 artifact |
| V5-T082 | Word native TOC | 刷新后 accepted |
| V5-T083 | WPS native TOC | 刷新后 accepted |
| V5-T084 | 无 Word/WPS | blocked |
| V5-T085 | 刷新后占位文本 | blocked |
| V5-T086 | 页码/条目不一致 | blocked |
| V5-T087 | equivalent visible TOC 已批准 | 按该模式验收 |
| V5-T088 | native 失败自动降静态 | 必须拒绝 |
| V5-T089 | resume 成功 result | 不重复执行 |
| V5-T090 | resume 部分失败 | 从 checkpoint 新 attempt |
| V5-T091 | v4 run resume | manual_recover/new run |

### 17.6 全链路

- 标准文档建规。
- 规则确认与规则激活。
- audit-only。
- 自动修复。
- 包含 L3 人工确认的修复。
- 动态目录修复。
- 中断恢复。
- 中文报告。
- 原始文件未覆盖证明。
- 规范命名与 final acceptance 不可变。

## 18. CI 与证据

最低 CI：

| Runner | runtime_id |
|---|---|
| windows-latest x64 | `win-x64` |
| macos-13 x64 | `osx-x64` |
| macos-14/15 arm64 | `osx-arm64` |

V5-014 发布必过平台限定为 `win-x64`、`osx-x64`、`osx-arm64`，对应常见 Windows 与 macOS 桌面 Codex CLI 使用场景。锁文件仍保留 OfficeCLI v1.0.113 的 8 个官方资产，用于 runtime resolver 识别、离线缓存和未来扩展；`win-arm64` 与 Linux 系列证据属于 best-effort/后补兼容，不作为 V5-014/V5-015 发布阻塞项。

平台证据必须包含 runner 信息、资产 hash、`--version`、capability hash、最小 DOCX create/get/set/validate/screenshot 结果。

Word/WPS TOC 测试允许专用 Windows runner，但不得标记为永久不可自动化。没有该 runner 时 v5 不允许发布。

## 19. 最终验收 Gate

全部满足才允许 v5 发布：

- [ ] lock 中 8 个资产与官方 v1.0.113 一致。
- [ ] capability manifest 恰好覆盖 OfficeCLI v1.0.113 固定源码中的 40 个 DOCX help target、canonical 元素和规定全局命令。
- [ ] 所有生产 Skill 和脚本不再调用直接 OOXML Python。
- [ ] snapshot v2 是语义和规则层唯一事实输入。
- [ ] finalized plan 到 request 的每个动作映射可复算。
- [ ] L3_WRITE 无人工确认无法执行。
- [ ] OfficeCLI 只修改工作副本或内部输出。
- [ ] validate、issues、after snapshot、render 均为写后必经步骤。
- [ ] native TOC 未经 Word/WPS 精确刷新不能 accepted。
- [ ] win/mac 必过 runtime（`win-x64`、`osx-x64`、`osx-arm64`）均有通过证据；其他已知 runtime 证据为 best-effort，不阻塞发布。
- [ ] 所有 L1 Gate 自动化测试通过。
- [ ] v4 历史 run 不被静默执行。
- [ ] 文档、Schema、Skill、validator、fixture、测试矩阵和许可证同步完成。
- [ ] 双线程架构评审与实施评审均为 PASS，无未解决 P0/P1/P2。

## 20. 实施文件清单

实施者必须创建或更新：

```text
tools/officecli/officecli.lock.json
tools/officecli/officecli-capability-manifest.json
THIRD_PARTY_NOTICES/OfficeCLI-Apache-2.0.txt
scripts/officecli/runtime_resolver.py
scripts/officecli/capability_manifest.py
scripts/officecli/snapshot_adapter.py
scripts/officecli/request_builder.py
scripts/officecli/runtime_adapter.py
scripts/officecli/toc_refresh_adapter.py
docs/v5/schemas/officecli-lock.schema.json
docs/v5/schemas/officecli-capability-manifest.schema.json
docs/v5/schemas/officecli-document-snapshot.schema.json
docs/v5/schemas/officecli-execution-request.schema.json
docs/v5/schemas/officecli-execution-result.schema.json
```

并升级 v4 对应的 repair plan、repair execution、review、evidence、final acceptance 契约到 `docs/v5/schemas/`。Schema 文件名保持业务名，`schema_version` 升为 `2.0.0`。

Python 适配器公共入口固定为：

```text
python -m scripts.officecli.runtime_resolver ensure --lock tools/officecli/officecli.lock.json
python -m scripts.officecli.capability_manifest verify --lock ... --manifest ...
python -m scripts.officecli.snapshot_adapter build --run-dir ... --kind before
python -m scripts.officecli.request_builder build --run-dir ... --repair-plan ...
python -m scripts.officecli.runtime_adapter execute --run-dir ... --request ...
python -m scripts.officecli.toc_refresh_adapter refresh --run-dir ... --input ...
```

各模块内部可以拆分，但以上 CLI 是 Skill 和测试唯一允许调用的稳定入口。

## 21. 规范化公共契约与实施细则

本节对前文所有简写具有覆盖效力。Schema 实现必须逐字落实，不得另行设计。

### 21.1 JSON、Hash 与公共类型

所有 JSON Schema 使用 draft 2020-12。除明确的 `extensions` 外，根对象、嵌套对象和数组项均固定 `additionalProperties=false`。未知字段一律 Schema 失败；不存在“非关键未知字段兼容”。`extensions` 为可选 object，键必须匹配 `^[a-z0-9]+(?:[._-][a-z0-9]+)+$`，值只允许 JSON 标量或标量数组，不参与 Gate。

canonical JSON 算法固定为 RFC 8785 JCS，UTF-8 无 BOM、无尾随换行。对象自身 hash 字段计算时，先从对象中删除该 hash 字段，再执行 JCS 和 SHA-256 小写十六进制。YAML artifact hash 以原始 UTF-8 bytes 计算，不做 YAML 重排。

`ArtifactRef`：

| 字段 | 类型 | 规则 |
|---|---|---|
| `artifact_id` | string | run 内唯一 |
| `kind` | enum | `lock/capability/snapshot/plan/request/result/log/review/evidence/toc_acceptance/final_acceptance/docx/html/png/raw_xml/executable/license` |
| `relative_path` | string | 相对 run 根目录或仓库根目录；禁止 `..` 和绝对路径 |
| `sha256` | sha64 | 文件原始 bytes hash |
| `size_bytes` | integer | `>=0` |
| `schema_id` | string/null | 非结构化文件为 null |
| `schema_version` | semver/null | 非结构化文件为 null |

`Warning`：`code` string、`severity` enum `info/warning/blocking`、`message` string、`json_pointer` string/null、`source_command` string/null。`blocking` warning 必须使当前 Gate 失败。

`Error`：`code`、`reason_code`、`message`、`stage`、`retryable`、`failed_artifact_ref` nullable、`native_exit_code` nullable、`native_error` nullable、`stderr_artifact_ref` nullable。所有字段必填，nullable 字段必须显式为 null。

`GateCheck`：`gate_id`、`status` enum `passed/failed/blocked`、`checked_at`、`predicate_version`、`evidence_refs[]`、`failed_codes[]`。数组允许空但字段必填。

### 21.2 Schema 完整字段

下列顶层对象全部必填 `schema_id/schema_version/created_at/extensions`；`extensions` 缺省写 `{}`。

`officecli-document-snapshot` 2.0.0：

- `snapshot_id/run_id/kind`，kind=`standard/before/after/post_toc`
- `source_docx_ref/officecli_executable_ref/capability_manifest_ref`
- `snapshot_source_hash`：等于 `source_docx_ref.sha256`
- `document`：`format=docx`、`root_path`、`node_count`、`part_count`、`has_toc`、`has_forms`、`has_revisions`、`has_protection`
- `nodes[]/parts[]/indexes/warnings[]/gate_check`
- `indexes` 固定含 `by_type`、`by_style_id`、`by_native_id`、`by_logical_identity`，值均为 string 到 node_id 数组

`nodes[]` 除前文属性外新增 `logical_identity`、`parent_logical_identity` nullable、`native_identity` nullable。`node_id` 只在单个 snapshot 内稳定；跨快照禁止直接比较 node_id。

跨快照匹配顺序固定为：同类型 native identity；同类型业务键；同类型稳定 OfficeCLI selector；父 logical identity + 同类型 + content fingerprint；父 logical identity + 同类型 + 相邻锚点。出现 0 个候选为 `missing`，超过 1 个为 `ambiguous` 并阻塞，不允许按数组位置猜测。文本或格式被计划修改时，使用计划记录的 before fingerprint 与未修改锚点，不要求 after fingerprint 相同。

`officecli-execution-request` 2.0.0：

- `request_id/run_id/plan_ref/plan_sha256/plan_revision`
- `snapshot_ref/lock_ref/capability_manifest_ref/officecli_executable_ref`
- `working_docx_before_ref/environment/batches[]/request_sha256/gate_check/extensions`
- `environment` 固定含 `OFFICECLI_SKIP_UPDATE="1"`、`OFFICECLI_NO_AUTO_RESIDENT="1"`、`locale="C.UTF-8"`、`timezone="UTC"`
- `batches[]` 除前文字段新增 `officecli_batch_ref/checkpoint_ref`
- `preconditions[]/postconditions[]/expected_result` 项统一字段：`predicate_id/type` enum `hash_equals/target_fresh/manual_confirmation/property_equals/artifact_exists/validate_clean/issues_nonblocking`、`target_ref`、`json_pointer` nullable、`expected` JSON scalar、`failure_code`

`officecli-execution-result` 2.0.0：

- 顶层使用前文全部字段。
- `batch_results[]` 必填 `batch_id/sequence/started_at/finished_at/duration_ms/exit_code/native_success/status/stdout_artifact_ref/stderr_artifact_ref/working_before_sha256/working_after_sha256/operation_results[]`。
- `operation_results[]` 不含 exit code，必填 `operation_id/source_action_id/index/status/native_success/native_output/native_error/before_target_fingerprint/after_target_fingerprint/postconditions_passed/duration_ms`；`not_run` 时 native 和 after 字段为 null。

`repair_execution_log` 2.0.0：`run_id/plan_ref/request_ref/result_refs[]/attempts[]/current_status/resume_policy/gate_check`。`current_status` 必须使用 §15.5 枚举。`resume_policy` 含 `resume_allowed`、`allowed_platforms[]`、`required_artifact_refs[]`、`max_additional_attempts`、`blocked_reason_code` nullable。`attempts[]` 固定含 `attempt_no/started_at/finished_at/checkpoint_ref/request_ref/result_ref/outcome/retry_reason_code/source_status/target_status`。

`review_result` 2.0.0：`review_id/run_id/plan_ref/before_snapshot_ref/after_snapshot_ref/action_results[]/summary/gate_check`。每个 action result 含 `action_id/status` enum `passed/failed/not_executed/manual_required`、`before_node_ref/after_node_ref` nullable、`expected_changes[]/observed_changes[]/unexpected_changes[]/evidence_refs[]/failure_codes[]`。

`evidence_manifest` 2.0.0：`manifest_id/run_id/artifacts[]/relations[]/completeness/gate_check`。`relations[]` 含 `from_artifact_id/to_artifact_id/relation` enum `derived_from/executes/verifies/renders/refreshes/accepts`。`completeness` 含 `required_kinds[]/present_kinds[]/missing_kinds[]`。

`toc_acceptance` 2.0.0：`run_id/required/status` enum `not_required/passed/blocked`、`viewer` nullable、`viewer_version` nullable、`platform` nullable、`input_ref/output_ref` nullable、`before_sha256/after_sha256` nullable、`field_update_count` nullable、`toc_update_count` nullable、`page_count` nullable、`visible_entries[]/evidence_refs[]/error` nullable、`gate_check`。

`final_acceptance` 2.0.0：`acceptance_id/run_id/status` enum `accepted/blocked/rejected`、`source_docx_ref/final_docx_ref` nullable、`lock_ref/capability_ref/before_snapshot_ref/after_snapshot_ref/plan_ref/request_ref/result_refs[]/repair_log_ref/review_ref/evidence_manifest_ref/toc_acceptance_ref`、`source_hash_unchanged`、`all_actions_reviewed`、`all_gates_passed`、`accepted_at` nullable、`blocking_codes[]`、`gate_check`。`accepted` 时所有 nullable evidence 必须非 null、三个 bool 必须 true、blocking_codes 为空。

### 21.3 Snapshot 遍历算法

v1.0.113 `query` 没有分页参数，前文“分页收集”废止。固定算法：

1. 对 `/document`、`/body`、`/styles`、`/numbering`、可用 header/footer/footnote/endnote/comments 根路径执行 `get --depth 3 --json`。
2. 使用 FIFO BFS；队列键为 canonical path。每个响应中的直接 child path 未访问时入队，再对该 path 执行 depth 3。
3. 同一路径第二次出现时 node type 和 native identity 必须一致，否则阻塞。
4. 只对 manifest 中 `operations` 包含 `query` 的元素执行一次无分页 query，结果用于 queryable 子集 completeness 交叉检查，不作为遍历主来源。
5. 上限：200,000 节点、256 MiB 累计 stdout、单命令 120 秒、总构建 30 分钟；任一超限返回 `FH-OFFICECLI-SNAPSHOT-LIMIT`。
6. dump 的 skipped/unsupported warning 必须写入 warnings。涉及 body、styles、numbering、section、header/footer、TOC、field、table、form 的 warning 为 blocking；其他 warning 只有 capability manifest 明确分类 nonblocking 才可继续。
7. `document.node_count` 必须等于 BFS 去重节点数。query completeness 只比较 queryable 类型子集：对每个 queryable 类型，BFS 中该类型的 stable path 集合必须覆盖 query 返回的 stable path 集合；query 返回 BFS 未见路径时阻塞。非 queryable 类型只用 BFS/get/dump 校验，不参与 query 数量相等断言。

### 21.4 文件生命周期

1. 外部原件只读，记为 `external_source`，永不放入 OfficeCLI 参数。
2. 原子复制为 `input/source.docx`，校验 hash 相同；该文件此后只读。
3. 再复制为 `input/working.docx`，校验 hash 相同；所有 batch 只修改它。
4. 每批前从 working 复制 checkpoint；成功后 working 原地成为下一批输入。
5. 全部 batch 和写后 Gate 通过后，复制 working 为 `output/_internal/executed.docx`，从此 working 不再修改。
6. TOC 不需要刷新时，从 executed 复制规范命名交付物。
7. TOC 需要刷新时，从 executed 复制 `output/_internal/toc-refresh.docx`，Word/WPS 只修改该文件；复验通过后复制为交付物。
8. final acceptance 前重新计算 external_source、source、executed/toc-refresh、final hashes；external_source 与 source 必须保持初始 hash。

### 21.5 HTML 与 Renderer 依赖

v1.0.113 `view html` 忽略 `--out`。固定命令为 `officecli view {docx} html`，非 JSON stdout 必须是 UTF-8 HTML；适配器捕获最多 64 MiB，空输出或超限阻塞，再原子写入 `preview/document.html` 并记录 hash。stderr 只存诊断 artifact。

renderer preflight 按 OfficeCLI 固定源码顺序探测 Chromium family 后 Firefox，记录实际 executable path、`--version` 和 hash。允许名称和固定路径必须与 `HtmlScreenshot.cs` v1.0.113 清单一致。没有后端返回 `FH-OFFICECLI-RENDERER-UNAVAILABLE`。CI 镜像必须锁定浏览器包版本或容器 digest；产品不自行声明 OfficeCLI 未强制的最低浏览器版本。

### 21.6 完整命令分类

capability generator 必须从固定源码 root command 注册表生成所有命令，不限前文示例。每项含 `name/docx_supported/automation_surface/risk_class/production_policy/reason`。生产 allow：`get/query/set/add/remove/move/swap/batch/dump/raw/raw-set/validate/view`；条件 allow：`create/merge/refresh/import` 仅 fixture 或明确迁移任务；deny：`open/close/save/watch/goto/mark/plugins/mcp/skills/update/uninstall/add-part`。`refresh` 的 OfficeCLI fallback 永远不得满足 native TOC Gate。未分类命令使 manifest Gate 失败。

### 21.7 Word/WPS TOC 适配器

native TOC 首版执行平台固定 Windows。非 Windows run 在写完 executed 后进入 `blocked_waiting_native_toc`，允许把同一 run 目录交给 Windows runner resume；不得改 run_id 或重跑 batch。

探测顺序固定：

1. Microsoft Word COM ProgID `Word.Application`。
2. WPS Writer COM ProgID `kwps.Application`。
3. 均不可用则阻塞。

状态机固定：`probe -> copy_refresh_input -> open_hidden -> update_all_fields -> update_all_tocs -> repaginate -> save -> close_document -> quit_application -> reopen_readonly -> verify_visible_toc -> officecli_revalidate`。Word 使用 `Documents.Open`、遍历 `Fields.Update`、遍历 `TablesOfContents.Update`、`Repaginate`、`Save`、`Close`、`Quit`；WPS 使用同名兼容自动化成员，任一成员不存在即 `wps_api_incompatible`，不得猜测替代调用。单阶段 120 秒，总计 600 秒。只终止适配器自己启动且记录 PID 的进程，禁止清理用户 Office/WPS 进程。

COM 固定参数：

- 启动后立即设置 `Application.Visible=false`、`Application.DisplayAlerts=0`。
- 若存在 `Application.AutomationSecurity`，必须设置为 `3`，即强制禁用宏；属性不存在时记录 `automation_security_unavailable` warning，但继续。
- `Documents.Open` 必须使用命名参数：`FileName=toc-refresh.docx`、`ConfirmConversions=false`、`ReadOnly=false`、`AddToRecentFiles=false`、`PasswordDocument=""`、`PasswordTemplate=""`、`Revert=false`、`WritePasswordDocument=""`、`WritePasswordTemplate=""`、`Format=0`、`Encoding=65001`、`Visible=false`、`OpenAndRepair=false`、`NoEncodingDialog=true`、`UpdateLinks=0`。
- 若 `Documents.Open` 返回受保护视图、密码提示、只读推荐、修订保护、写保护、宏安全提示、链接更新提示或转换提示，适配器不得交互点击，必须在阶段超时或异常后返回固定 reason_code：`protected_view/password_required/readonly_recommended/revision_protected/write_protected/macro_prompt/link_update_prompt/conversion_prompt`。
- `Fields.Update` 前先检查 `Document.ProtectionType`，非 `-1` 或等效未保护值时返回 `document_protected`。
- 保存必须调用 `Document.Save()`，不得 `SaveAs` 覆盖其他路径。
- `Close` 固定 `SaveChanges=0`，因为保存已显式完成；若保存失败不得 close 后继续验收。
- 任一阶段超时后只关闭本适配器创建的 document/application；若无法关闭，记录 PID 和 `cleanup_failed`，但不得杀死非本适配器 PID。

成功条件：COM 调用无异常、输出 hash 改变或 TOC 已确认无需字节变化、TOC entry 非空、页码均为正整数或合法范围、二次 reopen 成功、OfficeCLI validate clean、页数与截图证据一致。Word/WPS 各自必须有专用 Windows runner 测试；WPS 不可用不能由 Word 测试替代发布要求。

### 21.8 错误重试固定映射

- `NONJSON-OUTPUT`、`RESULT-MISMATCH`、`L3-NOT-AUTHORIZED`、`VALIDATE-FAILED`、`POSTCONDITION-FAILED`：不可重试。
- `TIMEOUT/BATCH-FAILED`：仅从未变更 checkpoint 新建 attempt 可重试一次；再次失败不可重试。
- `TARGET-STALE`：原 run 不可重试执行；重新 snapshot 和 finalized plan 后新 revision 可执行。
- `TOC refresh failed`：`viewer_busy/temporary_file_lock/process_start_failed` 可在新 attempt 重试一次；`api_incompatible/document_corrupt/page_mismatch/field_update_failed` 不可重试。
- retryable 值必须由 reason_code 纯函数决定，不允许人工改写。

## 22. 评审闭环

评审产物固定保存：

```text
docs/v5/reviews/round-{NN}/
├── reviewer-a-architecture.md
├── reviewer-b-implementation.md
├── reviewer-a-cross-review.md
├── reviewer-b-cross-review.md
└── resolution.md
```

状态要求：

- `PASS`：无 P0、P1、未解决 P2，无冲突结论。
- `FAIL`：存在任一上述问题。
- P3 可作为后续建议，但必须确认不影响零猜测实施。
- 每项 finding 必须含 ID、严重级别、章节、问题、证据、要求修改、状态。
- 交叉评审必须检查对方 finding 的正确性、遗漏和冲突，并复审修订后全文。

## 23. 参考基线

- OfficeCLI Release：`https://github.com/iOfficeAI/OfficeCLI/releases/tag/v1.0.113`
- 固定源码：`https://github.com/iOfficeAI/OfficeCLI/tree/v1.0.113`
- 固定 DOCX help Schema：`schemas/help/docx/*.json`
- 固定 DOCX Skill：`skills/officecli-docx/SKILL.md`
- OfficeCLI 项目文件：`.NET 10`、`DocumentFormat.OpenXml 3.4.1`
- 上游许可证：Apache-2.0

## 24. 变更记录

| 版本 | 日期 | 作者 | 变更 |
|---|---|---|---|
| v1.0.0-rc1 | 2026-06-15 | Codex | 初版：锁定架构、供应链、Schema、迁移、测试和双评审闭环 |
| v1.0.1 | 2026-06-16 | Codex | 实施 V5-003 时按 OfficeCLI v1.0.113 固定源码校正 DOCX help target 数量为 40，并明确 help_target 与 canonical element 映射 |
| v1.0.0 | 2026-06-16 | Codex | 关闭 A/B 双评审 P0/P1/P2：补齐 L3、Schema、batch、renderer、TOC、resume、供应链与文件生命周期；A/B 交叉终审 PASS |
