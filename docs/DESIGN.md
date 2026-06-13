# AIOffice 设计文档（100% 自研版）

> AI-Native Office CLI —— 让 AI Agent 在 macOS / Windows / Linux 上可靠地创建、读取、编辑、渲染、校验真实的 `.docx` / `.xlsx` / `.pptx`。
>
> 本文是自研架构的唯一设计依据（single source of truth）。**此前"封装第三方引擎二进制"的方案已彻底废弃**：AIOffice 不下载、不捆绑、不调用任何第三方 Office 引擎进程，全部文档能力由我们自己的 C#/.NET 代码基于 DocumentFormat.OpenXml 与 ClosedXML 实现。

---

## 1. 定位与设计原则

### 1.1 定位

`aioffice` 不是"给人用的 Office 命令行"，而是"给 AI 用的 Office 执行面"。AI Agent 是第一受众，人类是第二受众（human-in-the-loop 审阅者）。这个定位决定所有取舍：**输出必须可解析、行为必须可预测、错误必须可自愈、变更必须可回滚。**

技术底座（已定，不再讨论）：

| 层 | 选型 | 理由 |
|---|---|---|
| 运行时 | .NET（Homebrew 安装的 SDK，TFM 取 `dotnet --version` 主版本对应的 `net<major>.0`） | 单一语言栈、AOT 可期、跨平台 |
| OOXML 底层 | `DocumentFormat.OpenXml` 3.x | 微软官方 SDK，无损读写 zip part，三格式通吃 |
| xlsx 高层 | `ClosedXML`（含公式计算） | 单元格/区域/公式求值的成熟抽象；不够用处回落 OpenXml SDK |
| MCP | `ModelContextProtocol`（官方 C# SDK） | stdio server 标准实现 |
| 测试 | xunit | 质量体系的承重墙（§4） |

刻意**不**引入：System.CommandLine（Core 内手写轻量 argv 解析）、ImageSharp、任何渲染引擎二进制。依赖最少化是确定性的一部分。

### 1.2 与 OfficeCLI 的关系（说清楚，只说一次）

OfficeCLI 的**能力清单**是我们长期对齐的北极星（台账见 `docs/PARITY.md`），它验证了"AI 操作真实 Office 文件"这个产品方向。但 AIOffice：

1. **不封装它**——不存在引擎下载、版本钉死、manifest 校验这些东西；我们的二进制就是全部。
2. **不模仿它的命令语法**——命令与参数刻意**不** 1:1 还原。原因：(a) 两套相似但不完全相同的语法会让 Agent 混淆（半兼容比不兼容更糟）；(b) 它的命令面是渐进生长出来的，动词语义有历史包袱（`set`/`add`/`view` 职责交叠）；(c) 我们要的是一个为批量原子编辑、rev 乐观锁、schema 自省而**从零设计**的最小正交动词集。
3. **不复制它的代码**——其源码只作为行为与边界 case 的学习材料，实现全部自写。

### 1.3 七条设计原则

| # | 原则 | 含义 | 落地机制 |
|---|------|------|----------|
| 1 | **Deterministic JSON** | 任何命令、任何结局（成功/失败/部分警告）输出同一个 JSON envelope，机器可无脑解析 | `{ ok, data, error, meta }` 信封（§2.2），非 TTY 默认 `--json` |
| 2 | **Errors that Teach** | 错误是下一步动作的输入，不是终点 | `AiofficeError` 构造器强制非空 `suggestion`；`invalid_path` 服务端自动近邻重查回填 `candidates[]` |
| 3 | **Stable Addressing** | AI 多步操作必须有一套稳定、简单、三格式统一的地址语法 | 自研寻址文法（§2.4）：1-based 路径，`query` 返回规范路径，`get`/`edit` 接受之 |
| 4 | **Atomic Batch Edits** | AI 天然产出操作序列；序列要么全成要么全不成 | `edit --ops` 整批一个保存周期、失败回滚、`--dry-run` 预演、`--expect-rev` 乐观锁 |
| 5 | **Render → Look → Fix** | AI 改完必须能"看见"结果再迭代，否则是盲改 XML | `render`（html/svg/text）+ `validate` lint；png 与实时 preview M1 落地，pdf（分页打印）M3 落地 |
| 6 | **Introspectable Schema** | "不确定就查 schema，而不是猜参数" | `schema [verb]` 输出整个命令面的机器可读 JSON；`help` 渐进式文档 |
| 7 | **Sandbox by Default** | 给 AI 的工具必须默认最小权限 | `--workspace`（默认 cwd）白名单，realpath + symlink 逃逸检查，越界即 `sandbox_denied`（exit 4） |

另一条贯穿性纪律：**未实现的能力 = typed `unsupported_feature` envelope（必须附 workaround 建议），绝不 crash、绝不静默 no-op。**

---

## 2. 命令面规范（v0）

CLI 动词与 MCP 工具 **1:1 镜像**（§2.7），一套心智模型。v0 起 13 个能力，M1 加 `preview`（14），M7 加 `audit`（15），M8 加 `diff`（16），M9 加 `convert`（17）。

### 2.1 全局约定

```
aioffice <verb> [<file>] [args...] [flags...]
```

**全局旗标**：

| 旗标 | 含义 |
|------|------|
| `--json` | 输出 JSON envelope（**stdout 非 TTY 时为默认**） |
| `--pretty` | 人类可读输出（TTY 默认）；与 `--json` 互斥，后写者胜 |
| `--workspace <dir>` | 沙箱根目录，默认 cwd；亦可用环境变量 `AIOFFICE_WORKSPACE` |
| `--quiet` | 抑制 stderr 上的提示性输出（envelope 不受影响） |

文件类型一律由扩展名推断（`.docx`/`.xlsx`/`.pptx`），无扩展名或未知扩展名 → `invalid_args`。

### 2.2 Envelope（每条命令、stdout、恰好一个 JSON 对象）

```json
{
  "ok": true,
  "data": { },
  "error": null,
  "meta": {
    "file": "report.docx",
    "rev": "9f8a2c1d4e6b",
    "elapsedMs": 42,
    "version": "0.1.0",
    "warnings": [ { "code": "formula_not_evaluated", "message": "Sheet1!B7 holds a cached value; formulas were not recalculated." } ]
  }
}
```

失败时：

```json
{
  "ok": false,
  "data": null,
  "error": {
    "code": "invalid_path",
    "message": "Path '/body/p[99]' not found: document has 12 paragraphs.",
    "suggestion": "Use one of the candidates, or run `aioffice query report.docx p` to list all paragraph paths.",
    "candidates": ["/body/p[12]", "/body/p[11]"]
  },
  "meta": { "file": "report.docx", "rev": "9f8a2c1d4e6b", "elapsedMs": 17, "version": "0.1.0" }
}
```

约束：

- `ok=true` ⇒ `error=null`；`ok=false` ⇒ `data=null` 且 `error.suggestion` **必须非空**（构造器层面强制，写不出建议的错误不允许 throw）。
- `meta.rev` = 文件字节 SHA256 的前 12 个 hex 字符；任何触碰文件的命令都返回。读 = 拿 rev，写 = 可带 `--expect-rev` 守卫。
- `meta.warnings[]` 承载非致命问题（如 `formula_not_evaluated`），不影响 `ok`。

**错误码（封闭枚举，新增需过评审）**：

| code | 触发 | 强制行为 |
|------|------|---------|
| `invalid_args` | 参数/旗标无法解析、缺必填、值非法 | suggestion 指出正确用法或让查 `schema <verb>` |
| `file_not_found` | 文件不存在 | suggestion 给出 workspace 内近名候选（若有） |
| `sandbox_denied` | 路径越出 workspace / symlink 逃逸 | suggestion 提示 `--workspace` 调整方式 |
| `invalid_path` | 地址/选择器解析失败或未命中 | **必须回填 `candidates[]`**：服务端自动做一次近邻匹配查询 |
| `unsupported_feature` | 已规划未实现的能力 | **suggestion 必须给出可行的替代做法**（workaround） |
| `format_corrupt` | zip/XML 损坏、非 OOXML | suggestion 建议 `validate` 或从快照恢复 |
| `stale_address` | `--expect-rev` 与当前 rev 不符（**写前检查，绝不部分落盘**） | suggestion 让重新 `read`/`get` 拿新 rev 重试 |
| `formula_not_evaluated` | （warning，不是 error）读到公式单元格的缓存值且未重算 | 提示数据可能陈旧 |
| `internal_error` | 我们的 bug | suggestion 让带 `--json` 输出提 issue |

### 2.3 动词逐一规范

#### `create` — 新建文档

```
aioffice create <file> [--from <source>] [--kind docx|xlsx|pptx] [--title <T>]
```

- `--kind` 默认由扩展名推断；显式给出且与扩展名冲突 → `invalid_args`。
- 产出最小合法文档：docx = 空 body + 默认样式部件；xlsx = 单 `Sheet1`；pptx = 单标题版式空白演示。`--title` 写入 core properties（及 pptx 首页标题占位符 / xlsx 首 sheet 名）。
- 文件已存在 → `invalid_args`（suggestion 提示换名或先删除；v0 刻意不提供 `--force`，破坏性默认是错误的）。
- **M5 导入桥** `--from <source>`：按**源扩展名**路由——`.md`/`.markdown` → docx 导入器（Markdig GFM：标题/列表/表格/链接/代码/图片）、`.csv`/`.tsv` → xlsx 导入器（RFC 4180，类型化单元格，前导零保文本）；源/目标不匹配 → `invalid_args` 且 suggestion 给全矩阵；源路径必经 workspace 沙箱。处理器侧钩子是 `IFormatHandler.CreateFrom(ctx, sourcePath)` 的**默认接口实现**（额外能力增量加入，无导入器格式默认 `unsupported_feature`），命令层路由在 `AIOffice.Mcp.Bridge`（CLI create 与 MCP office_create 共用）。

```bash
aioffice create report.docx --title "Q3 Review"
# => { "ok": true, "data": { "file": "report.docx", "kind": "docx" }, "meta": { "rev": "..." } }
aioffice create report.docx --from notes.md      # markdown -> docx
aioffice create orders.xlsx --from orders.csv    # csv -> xlsx
```

#### `read` — 整体读取

```
aioffice read <file> [--view outline|text|stats|structure|markdown|csv] [--range a..b] [--sheet NAME] [--max-bytes N]
```

- `--view outline`（默认）：标题树 / sheet 清单 / slide 标题清单。
- `--view text`：纯文本线性化（docx 段落、xlsx 单元格值 TSV 风格、pptx 形状文本）。
- `--view stats`：计数面板（段落/字数/sheet/行列已用区域/slide/shape 数、part 数、文件大小）。
- `--view structure`：浅层节点树（带规范路径），是 AI 后续 `get`/`edit` 的地址来源之一。
- `--view markdown`（M5，docx 专属）：body 导出为 GFM markdown，与 `create --from` 导入结构 round-trip。
- `--view csv`（M5，xlsx 专属）：单 sheet 以 RFC 4180 csv 导出，`--sheet` 选表（缺省首个）、`--range A1:C10` 限窗。
- 桥接视图用错格式（如对 xlsx 要 markdown）→ `unsupported_feature`，suggestion 列出该格式的全部有效视图（命令层 `Bridge.GuardBridgeView`，CLI 与 MCP 共用）。
- `--range a..b`：限定输出窗口（docx 段落序号、xlsx 行号、pptx slide 序号），1-based 闭区间。
- `--max-bytes N`：输出 data 超限即截断并在 `meta.warnings` 标注 `truncated`。

```bash
aioffice read data.xlsx --view stats
aioffice read deck.pptx --view outline --range 1..5
```

#### `query` — 声明式查找，返回稳定路径

```
aioffice query <file> <selector>
```

返回 `data.matches[]`，每项 `{ path, type, summary }`，`path` 为规范路径，可直接喂给 `get`/`edit`。选择器文法见 §2.5。无命中是 `ok=true` + 空数组（不是错误）；选择器语法错才是 `invalid_args`。

```bash
aioffice query report.docx "p[style=Heading1]"
aioffice query data.xlsx "cell[value>100]"
aioffice query deck.pptx "shape:contains('Q3')"
```

#### `get` — 取单个节点及属性

```
aioffice get <file> <path>
```

返回 `data.node = { path, type, text?, props: {...}, children?: [paths] }`。路径未命中 → `invalid_path` + `candidates[]`。

```bash
aioffice get report.docx "/body/p[3]"
aioffice get data.xlsx "/Sheet1/B2"
# xlsx 公式单元格: props 含 formula 与 cachedValue；若未重算，meta.warnings 标 formula_not_evaluated
```

#### `edit` — 原子批量修改（唯一的写动词）

```
aioffice edit <file> --ops <json|@file> [--dry-run] [--expect-rev <rev>]
# 单操作语法糖（与 --ops 等价，可与之混用但不建议）:
aioffice edit <file> --set <path> k=v [k=v...]
aioffice edit <file> --add <path> --type <T> [k=v...] [--position <n|start|end|before:<path>|after:<path>>]
aioffice edit <file> --remove <path>
```

`--ops` 是 JSON 数组（或 `@ops.json` 从文件读）：

```json
[
  { "op": "set",    "path": "/body/p[3]",  "props": { "text": "Revenue grew 25%", "style": "Heading2" } },
  { "op": "add",    "path": "/body",       "type": "paragraph", "props": { "text": "New summary." }, "position": "end" },
  { "op": "remove", "path": "/body/p[7]" },
  { "op": "move",   "path": "/body/p[2]",  "position": "after:/body/p[5]" }
]
```

语义（这是整个工具最重要的一条契约）：

1. **原子性**：整批操作在一次打开-保存周期内执行；任何一条失败 ⇒ 不落盘、文件零变化，envelope 报出第一条失败的 op 序号与原因。
2. **顺序内寻址**：op 按数组顺序执行，路径在**执行时**解析（前面的 add/remove 会影响后面的序号——envelope 的 `data.results[]` 回报每条 op 实际命中的规范路径，供 AI 复核）。
3. **乐观锁**：`--expect-rev` 不匹配 ⇒ `stale_address`，**发生在任何写动作之前**。
4. **自动快照**：每次成功落盘前，先把原文件压入快照环（§2.3 `snapshot`）。
5. **预演**：`--dry-run` 跑完全部解析与校验、返回将发生什么（`data.plan[]`），不写文件、不打快照。

成功返回：

```json
{ "ok": true,
  "data": { "applied": 4, "results": [ { "op": "set", "path": "/body/p[3]" }, ... ] },
  "meta": { "file": "report.docx", "rev": "新rev", "version": "0.1.0", "elapsedMs": 88 } }
```

各元素类型可用的 `type`/`props` 清单由 `aioffice help docx paragraph` 等渐进文档与 `schema edit` 给出，不在此罗列。

#### `render` — 看见结果

```
aioffice render <file> [--to html|svg|text] [--scope <path>] [-o <out>]
```

- MVP：docx/xlsx → html；pptx → 每页 svg（或 html）；text 三格式皆可。
- `--scope` 限定渲染子树（如 `/slide[2]`、`/Sheet1/A1:D20`）。
- `-o` 省略时渲染产物进 envelope `data.content`（受 `--max-bytes` 同款截断保护）；给出时写文件、data 只含路径。
- `--to png` 在 M1（无头浏览器探测）；v0 请求 png → `unsupported_feature`，suggestion: "Render to html and screenshot it with a browser, or wait for M1 png support."

#### `validate` — 校验与 lint

```
aioffice validate <file>
```

- 跑 `OpenXmlValidator` + 我们的 lint 规则（空样式引用、断裂的 rId、xlsx 公式悬空引用等）。
- 返回 `data.issues[] = { severity: error|warning, code, message, path?, suggestion }`。有 error 级 issue 时 `ok` 仍为 `true`（命令本身成功了），靠 `data.issues` 表达文档问题。

#### `audit` — 无障碍 + 质量审计（M7）

```
aioffice audit <file> [--category accessibility|quality|all] [--severity error|warning|info] [--fix]
```

- 与 `validate` 同理，findings 是**数据**：即使有 error 级 finding，`ok:true`（exit 0）。`validate` 看「文件合不合规」，`audit` 看「文件做得好不好、对辅助技术友不友好」。
- 返回 `data = { findings: [{ id: "code#path", severity, category, code, path?, message, suggestion, autofixable }], summary: { errors, warnings, infos } }`。`--category`/`--severity` 过滤（severity 是最低上报级）。
- `--fix` 只应用**安全、非破坏性**自动修复（占位 alt 文本、标记表头行、设文档/幻灯片标题、删孤儿书签），再重审，追加 `fixed:N` 与 `remaining:[id]`；修复前自动快照可撤销。
- Core 增量接口 `IAuditor { AuditResult Audit(ctx, opts); int Fix(ctx, findingIds); }`，三格式 handler 各实现；命令层共享 `AuditVerb`（CLI `audit` 与 MCP `office_audit` 同产物）。code 全集见 `aioffice help audit`。

#### `diff` — 语义对比 / 评审（M8）

```
aioffice diff <file> [<other>] [--snapshot N] [--view summary|detailed]
```

- 与 `audit`/`validate` 同理，变更是**数据**：哪怕两份文档差异巨大也 `ok:true`（exit 0）。语义对比 `file`（当前）与一份基线，返回**已排序、跨平台确定性**的变更清单。
- 基线 = `other`（另一**同格式**文件，跨格式 → `invalid_args`）**或** `--snapshot N`（`file` 自身编辑前快照环里的第 N 份），二者恰选其一（都缺/都给 → `invalid_args`）。`--snapshot` 把那份快照还原到一个临时基线再对比，原文件与快照环不动。
- 返回 `data = { changes: [{ kind: added|removed|modified|moved, path, before?, after?, detail? }], summary: { added, removed, modified, moved }, baseline, view }`。`path` 是当前文档的规范路径（`removed` 取基线路径）；`modified` 带简洁 before/after；`moved` 的 `detail` 是 `moved from <old path>`（段落/幻灯片/行用 LCS/内容哈希区分 moved 与 added+removed）。
- 变更**始终**按 `(path, kind)` 序贯字典序排序（`DiffResult.FromChanges`），故同样两份文档每平台、每次运行 diff 结果一致。`--view summary` 把每条裁剪为 `{kind, path}`。
- Core 增量接口 `IDiffer { DiffResult Diff(ctx, baselineFile); }`，三格式 handler 各实现；命令层共享 `DiffVerb`（CLI `diff` 与 MCP `office_diff` 同产物）。

#### `template` — 占位符合并

```
aioffice template <file> --data <json|@file> [-o <out>]
```

- 在 docx/xlsx/pptx 的文本 run 中做 `{{key}}` 替换（跨 run 拆分的占位符要正确拼接识别——这是已知脏活，测试覆盖）。
- `-o` 省略 = 原地写（先自动快照）；给出 = 写新文件、原文件不动。
- `data.unresolved[]` 列出文档里有但 JSON 没给的 key（warning，不失败）。

#### `snapshot` — 快照环与回滚

```
aioffice snapshot list <file>
aioffice snapshot restore <file> [n]
```

- 每次成功的 `edit`/`template`（原地写）自动把**前像**压入快照环，环大小 20。
- 存放：`<workspace>/.aioffice/snapshots/<path-hash>/<n>.snap`（path-hash 取规范化绝对路径的哈希，与内容无关，保证同一文件历史聚合）。
- `list` 返回 `{ n, takenAt, rev, byVerb }`；`restore` 不带 n 默认最近一份；restore 自身也先快照当前态（undo 可再 undo）。

#### `doctor` — 环境体检

```
aioffice doctor
```

诊断并返回结构化结果 + 修复建议：.NET 运行时版本、workspace 可写性、快照目录状态、（M1 起）png 渲染所需的无头浏览器探测。没有引擎要检查——我们就是引擎。

#### `schema` — 命令面自省

```
aioffice schema [verb]
```

输出整个命令面（或单动词）的机器可读 JSON：动词、参数、旗标、类型、枚举、必填性、示例、错误码。MCP 工具的 inputSchema 与它同源生成。**Agent 的第一课：不确定就 `schema`，不要猜。**

#### `help` — 渐进式文档

```
aioffice help [topic]
```

topic 包括：`addressing`（寻址文法全文）、`selectors`（选择器文法全文）、`docx|xlsx|pptx`（该格式元素类型清单）、`docx paragraph` 等（单元素的属性表 + 示例）。`--json` 下输出结构化。

#### `mcp` — 启动 MCP server

```
aioffice mcp
```

stdio 上的 MCP server，见 §2.7。

### 2.4 寻址文法（自研，稳定且简单）

```
路径   := '/' 段 ( '/' 段 )*
段     := 元素名 '[' 序号 ']'        # 序号 1-based
docx   :  /body/p[3]
          /body/p[3]/run[2]
          /body/table[1]/tr[2]/tc[1]
          /header[1]/p[1]            # 页眉; /footer[1]/... 同理
xlsx   :  /Sheet1/A1                 # 单元格
          /Sheet1/A1:C10             # 范围
          /Sheet1/row[3]             # 整行
          /'Q3 Data'/B2              # 含空格的 sheet 名用单引号
pptx   :  /slide[2]
          /slide[2]/shape[3]
          /slide[2]/shape[3]/p[1]
          # master/layout 寻址保留到 M1
```

规则：索引一律 1-based；`query` 返回的就是这套规范路径；shell 中路径建议引号包裹防 glob。序号会因增删漂移——多步编辑的正确姿势是 `edit` 一次原子批量 + `--expect-rev`，而不是多次单步盲改。

### 2.5 选择器文法（CSS-like）

```
选择器 := 类型 [ '[' 谓词 ']' ]* [ ':' 伪类 ]*
类型   := p | run | table | tr | tc | cell | row | sheet | slide | shape | ...   # 见 help <fmt>
谓词   := 属性 op 值        op := = | != | > | >= | < | <=
          p[style=Heading1]      cell[value>100]      run[bold=true]
伪类   := :contains('text')     # 大小写不敏感子串
          shape:contains('Q3')
组合   := 谓词可叠加: cell[value>100][value<500]
```

数值比较仅对可数值化的属性生效；类型或属性名拼错 → `invalid_args`，suggestion 指向 `help <fmt> <type>`。

### 2.6 退出码

| code | 含义 | 上层（shell/CI/Agent 框架）重试策略建议 |
|------|------|------|
| `0` | 成功（含带 warnings 的成功） | — |
| `2` | 用户/输入错误（invalid_args, file_not_found, invalid_path, stale_address） | 修正输入后重试，盲重试无意义 |
| `3` | 内部/格式错误（internal_error, format_corrupt） | 可 `validate`/`snapshot restore` 后重试一次 |
| `4` | 沙箱拒绝（sandbox_denied） | 不重试，须人工调整 workspace |
| `5` | 不支持的能力（unsupported_feature） | 按 suggestion 的 workaround 改道 |

任何退出码下 stdout 都有完整 envelope——退出码给 shell 分流，envelope 给 AI 决策。

### 2.7 MCP 工具面（M9 起 17 个，stdio，官方 C# SDK）

CLI 动词与 MCP 工具 1:1，同一条内部执行路径（杜绝"CLI 能做、MCP 不能做"）：

| MCP tool | = CLI verb | 备注 |
|---|---|---|
| `office_create` | `create` | |
| `office_read` | `read` | |
| `office_query` | `query` | |
| `office_get` | `get` | |
| `office_edit` | `edit` | ops 数组直接作为工具参数 |
| `office_render` | `render` | |
| `office_validate` | `validate` | |
| `office_audit` | `audit` | M7：无障碍 + 质量 lint，findings 是数据，`fix:true` 仅安全自动修复 |
| `office_diff` | `diff` | M8：语义对比文档与基线（`other` 文件 / `snapshot:N`），变更是数据，已排序确定性 |
| `office_template` | `template` | |
| `file_snapshot` | `snapshot` | list/restore 二合一，`action` 参数区分 |
| `office_status` | `doctor` | |
| `office_help` | `help` | |
| `office_schema` | `schema` | |
| `preview_open` | `preview open` | M1 转正：拉起 detached `aioffice preview open` 子进程，等 lockfile + health（≤10s），返回 url |
| `preview_selection` | `preview selection` | M1 转正：读回用户在浏览器中点选的规范路径 |

- 工具返回值即 §2.2 envelope（序列化为 tool result 文本）。
- 每个工具的 inputSchema 由 `schema` 注册表生成，与 CLI 旗标语义逐字段对应。
- server 同样受 `--workspace`/`AIOFFICE_WORKSPACE` 沙箱约束。

---

## 3. 架构

### 3.1 解决方案布局

```
AIOffice.sln
Directory.Build.props            # TFM net<major>.0、LangVersion latest、Nullable enable、
                                 # ImplicitUsings、TreatWarningsAsErrors、Version 0.1.0
src/AIOffice.Core/               # envelope、错误体系、寻址、选择器、沙箱、快照、命令抽象、argv 解析、schema 注册表
src/AIOffice.Word/               # docx handler（DocumentFormat.OpenXml wordprocessingml）
src/AIOffice.Excel/              # xlsx handler（ClosedXML 为主，必要处回落 OpenXml SDK）
src/AIOffice.Pptx/               # pptx handler（OpenXml SDK presentationml）
src/AIOffice.Mcp/                # MCP stdio server（官方 SDK），13 工具
src/AIOffice.Cli/                # Program.cs 入口，命令注册表接全格式；OutputType Exe；AssemblyName aioffice
tests/AIOffice.Core.Tests/       # + Word.Tests / Excel.Tests / Pptx.Tests / Mcp.Tests
fixtures/                        # 输入夹具（含 fixtures/manual-check/ 供人工用真 Office 打开验证）
goldens/                         # 金样输出（round-trip 与渲染对照）
docs/                            # 本文、PARITY.md、MCP.md
.github/workflows/ci.yml         # CI 矩阵（§4.4）
```

依赖方向（单向，禁止回环）：

```
Cli ──┐
      ├──▶ Core ◀── Word / Excel / Pptx（实现 Core 的 IFormatHandler）
Mcp ──┘
Word/Excel/Pptx ──▶ DocumentFormat.OpenXml / ClosedXML
```

### 3.2 Core 抽象

```csharp
// 每个动词一个命令对象；CLI argv 与 MCP 工具参数都先解析成它
public interface ICommand
{
    string Verb { get; }
    CommandSchema Schema { get; }                       // schema/help 的供数源
    Task<Envelope> ExecuteAsync(CommandContext ctx);    // ctx: workspace, 解析后参数, IFormatHandler 路由
}

// 每格式一个 handler，能力不全没关系——缺的能力返回 unsupported_feature
public interface IFormatHandler
{
    OfficeKind Kind { get; }                            // Docx | Xlsx | Pptx
    DocumentSession Open(SandboxedPath path);           // 读入内存、计算 rev
    Node Get(DocumentSession s, Address addr);
    IReadOnlyList<Match> Query(DocumentSession s, Selector sel);
    EditResult ApplyEdits(DocumentSession s, IReadOnlyList<EditOp> ops, bool dryRun);
    RenderResult Render(DocumentSession s, RenderTarget to, Address? scope);
    // ... create / readView / validate / template
}

// 一次命令生命周期内的文档会话：打开→操作→(快照→)保存
public sealed class DocumentSession : IDisposable
{
    public string Rev { get; }                          // 打开时文件字节 SHA256 前 12 hex
    public void EnsureRev(string? expectRev);           // 不符 throw AiofficeError(stale_address)——必须发生在任何写之前
    public void SaveWithSnapshot(SnapshotStore store);  // 前像入环 → 原子落盘（写临时文件再 rename）
}

// 寻址解析：路径字符串 → 强类型 Address；失败时自动近邻匹配产出 candidates
public static class AddressResolver
{
    public static Address Parse(OfficeKind kind, string path);              // 文法错 → invalid_args
    public static ResolveResult Resolve(DocumentSession s, Address a);      // 未命中 → invalid_path + candidates[]
}

public sealed class SnapshotStore   // <workspace>/.aioffice/snapshots/<path-hash>/，环 20
{
    public IReadOnlyList<SnapshotInfo> List(string canonicalPath);
    public void Push(string canonicalPath, byte[] preImage, string byVerb);
    public void Restore(string canonicalPath, int? n);   // restore 前先 Push 当前态
}

public sealed class Workspace        // 沙箱
{
    public SandboxedPath Resolve(string userPath);
    // realpath 规范化 + symlink 逃逸检查；越界 throw AiofficeError(sandbox_denied)
}

public sealed class AiofficeError : Exception
{
    public AiofficeError(ErrorCode code, string message, string suggestion,
                         IReadOnlyList<string>? candidates = null)
    {
        if (string.IsNullOrWhiteSpace(suggestion))
            throw new ArgumentException("Every AiofficeError MUST teach: suggestion is required.");
        // ...
    }
}
```

### 3.3 数据流

```
            ┌────────────┐        ┌─────────────┐
 argv ────▶ │ AIOffice.Cli│        │ AIOffice.Mcp │ ◀──── MCP stdio (官方 SDK)
            └──────┬─────┘        └──────┬──────┘
                   │  argv 解析            │  工具参数绑定
                   ▼                      ▼
            ┌──────────────────────────────────┐
            │  Core: CommandRegistry            │  ★ one source of truth
            │  · schema 校验  · Workspace 沙箱   │
            │  · rev/--expect-rev  · 快照        │
            │  · envelope 组装 · 错误码+suggestion│
            └──────────────┬───────────────────┘
                           │ 按扩展名路由
            ┌──────────────┼───────────────┐
            ▼              ▼               ▼
      AIOffice.Word   AIOffice.Excel  AIOffice.Pptx
      (OpenXml WP)    (ClosedXML/     (OpenXml PML)
                       OpenXml SS)
```

要点：

- **CLI 与 MCP 是同一命令层的两张皮**。两边都把输入归一成 `ICommand` + `CommandContext`，往下完全同路。
- **envelope 在 Core 统一组装**：handler 只 return 数据或 throw `AiofficeError`；`elapsedMs`/`rev`/`version`/退出码映射都在一处。
- **保存是原子的**：写临时文件 + rename，半成品永不覆盖原文件；快照发生在 rename 之前。
- xlsx 公式：优先用 ClosedXML 的求值能力刷新缓存值；求不动的（外部引用、不支持的函数）保留缓存值并打 `formula_not_evaluated` warning——**宁可显式说"这个数可能旧了"，不让 AI 把陈旧数字当真。**

---

## 4. 质量体系（anti-zero-tests 立场，不可谈判）

参照对象 OfficeCLI 约 20 万行 handler 而测试稀薄——我们反着来：**每个格式包与它的 xunit 测试包同一次提交落地，没有测试的能力等于不存在。**

### 4.1 Round-trip 定律

> 打开一个文件、不做任何编辑、保存——每个 zip part 必须逐字节相同；任何 part 合法地发生变化，必须在测试里写明白为什么。

- 测试在 `fixtures/` 全量夹具上执行：解包前后 zip，逐 part 比对字节。
- OpenXml SDK / ClosedXML 某些路径会重排属性或重写 content types——一旦出现，要么换成不破坏的 API 路径，要么在测试里用显式 allowlist 记录该 part + 原因（带链接到上游 issue）。allowlist 是耻辱柱，不是垃圾桶。

### 4.2 Validator oracle

每个**变更类**测试（edit/template/create）结束后必须追加断言：`OpenXmlValidator` 对产物报 **0 errors**。这是"文件还能被 Office 打开"的机器代理。

### 4.3 金样与人工核验

- `goldens/` 存金样产物（结构 dump、渲染 html/svg）；行为变化必须显式重录金样并 review diff。
- 真 Office 打开无法在 CI 自动化：测试把代表性产物生成到 `fixtures/manual-check/`，由人定期用 Word/Excel/PowerPoint 真开核验；OpenXmlValidator 作为日常代理。

### 4.4 CI 矩阵

`.github/workflows/ci.yml`：

- OS：`macos-latest` + `windows-latest` + `ubuntu-latest`（路径/大小写/zip 行为差异都在这暴露）。
- 步骤：`dotnet build -warnaserror` → `dotnet test`（全部测试包）→ round-trip 套件 → validator oracle 套件。
- 任何红 = 不可合入。没有"先合了再补测试"。

### 4.5 编码纪律

- `TreatWarningsAsErrors` + `Nullable enable` 全仓生效。
- 小而能跑的代码优于宏大的半成品：每个 PR 等价单元交付的是可编译、可测试、可回滚的最小完整闭环。
- 未实现 = typed `unsupported_feature`（带 workaround），handler 里不允许出现 `throw new NotImplementedException()`。

---

## 5. 路线图

诚实声明：以下 scope 是"做到什么程度敢叫完成"的承诺，不是营销清单。每项能力以测试 + validator oracle 为完成标准。

### M0 — Scaffold（本期）

- 解决方案六包 + 五测试包就位，CI 三 OS 矩阵全绿。
- 三格式 `create` / `read`（四种 view）/ `get` / `query`（基础选择器：类型 + `=`/数值比较 + `:contains`）/ `edit`（set/add/remove/move，原子批量，--dry-run，--expect-rev）。
- xlsx 公式求值（ClosedXML）+ 缓存值回退 + `formula_not_evaluated` warning。
- `validate`（OpenXmlValidator + 首批 lint 规则）、`template`（{{key}}，含跨 run 拼接）、`snapshot`（环 20 + restore）、`schema` / `help` / `doctor`。
- MCP server 13 工具（`preview_open` 返回 `unsupported_feature`）。
- 质量体系全部生效：round-trip 定律、validator oracle、金样、manual-check 夹具。

### M1 — Render 完整 + Preview（已交付，v0.2.0）

- ✅ `render --to png`：`AIOffice.Render` 无头浏览器（Chrome/Edge 探测，`doctor` 报告 browser 状态）截图 html/svg；docx/xlsx 整文档，pptx 单页（`--scope /slide[N]`，缺省首页 + meta warning）。探测不到浏览器 → `unsupported_feature` + 安装建议。
- ✅ 浏览器实时预览 + 选区读取：`AIOffice.Preview` 本地 server（`preview open/selection/close`，SSE 自动刷新），渲染产物携带 `data-aio-path` 规范路径标注；`preview_open` / `preview_selection` MCP 工具转正（13 → 14 工具）。
- ✅ pptx master/layout 寻址（`/master[1]`、`/master[1]/layout[2]`，get/query 只读）。
- ✅ 计划外提前落地：docx 页眉/页脚（`/header[1]/p[1]` 寻址 + 编辑）、xlsx 图表 bar/line/pie（自研 OpenXml ChartPart，`/Sheet1/chart[1]`）。
- 顺延：annotated 注记视图、query 高级选择器、稳定 ID 寻址、全文 find/replace、pptx 截图网格 —— 移入 M2 窗口（见 PARITY.md 🔜 M1 列）。

### M2 — 深水区能力（已交付，v0.3.0）

- ✅ docx 修订（track changes）：`edit --track --author`（MCP `office_edit {track, author}`）写文本级 `w:ins`/`w:del`；`read --view revisions`；`accept`/`reject` op 按 `/revision[@id=N]` 或范围裁决。作者解析：op props.author > `--author` > `AIOFFICE_AUTHOR` > "AIOffice"。格式/移动修订 → M3。
- ✅ docx 批注：`add type:comment`（锚到段落或 run）、`read --view comments`、`/comment[@id=N]` get/remove。
- ✅ docx 样式管理：`/styles` add、`/style[@id=X]` set/get/remove、`read --view styles`；套用即 `set p {style}`。
- ✅ docx / xlsx / pptx 图片：PNG/JPEG，src 必经 workspace 沙箱（逃逸 → `sandbox_denied`），缺省尺寸守纵横比。
- ✅ xlsx 数据透视表：rows/columns/filters + values（sum/average/count/min/max），targetSheet 自动建，`pivot[@name=X]` 寻址，refreshOnLoad（Excel 打开即重算）。
- ✅ xlsx 条件格式：cellIs / colorScale / dataBar / containsText 四类，`/Sheet1/conditionalFormat[i]` 寻址。
- ✅ pptx：真 `p:bg` 纯色背景、`/slide[i]/notes` 演讲者备注（set/add/remove/get）。
- ✅ 大文件守卫：超过 50MB（`AIOFFICE_MAX_FILE_MB` 可调）拒绝打开 → `file_too_large` + 建议；`doctor` 报告 `limits.maxFileMb`。
- ⏭ 大文件流式处理**没有**按 M2 交付：它需要一轮专门的基准驱动打磨（10 万行 xlsx / 千页 docx 实测），移入 M3；M2 以尺寸守卫诚实兜底。

### M3 — 功能第一（已交付，v0.4.0）

用户指令：**功能第一**——不让体积/性能的节俭挡住能力；二进制增长可接受。

- ✅ 大小上限翻转：`FileSizeGuard` 默认**不限大小**；`AIOFFICE_MAX_FILE_MB` 改为 opt-in 上限（设了且超限才 `file_too_large`）；`doctor` 缺省报告 `limits.maxFileMb: "unlimited"`。
- ✅ `render --to pdf`：`AIOffice.Render` PdfRenderer 经系统 Chrome `--headless=new --print-to-pdf --no-pdf-header-footer` 出分页 PDF；docx/xlsx html → A4 分页；pptx 整副 deck 一个 PDF、`@page` 钉死片尺寸、一页一片（`--scope` 可缩到单页）。无浏览器 → `unsupported_feature` + 安装/替代建议。CLI 与 MCP `office_render` 同步增加 `pdf` 目标。
- ✅ 跨文档 dataFrom（xlsx 数据 → pptx 图表）：命令层（CLI edit 与 MCP office_edit 共用 `CrossDocDataFrom`）把 `{"dataFrom":"metrics.xlsx!Sheet1/A1:B5"}` 展开为字面量 categories/series——首列 → 分类、表头行 → 系列名、其余列 → 系列值；工作簿必经沙箱解析，经 xlsx handler 读取；区域写错返回 candidates（表名/最近 usedRange）。
- ✅ docx：列表（编号/项目符号/嵌套/重启，text 视图标记 + HTML 真 `<ol>/<ul>`）、超链接（url/anchor）、书签、脚注、节属性（`/section[1]` 纸张/方向/边距）、格式修订 accept/reject（w:rPrChange/w:pPrChange）、批注线程回复（w15 commentsExtended）。
- ✅ xlsx：大工作簿流式读取（>20 MB 或 `stream:true` 走 SAX，stats/text/get 不加载 DOM）、scatter/area 图表、命名区域（`/name[@name=X]`，公式真求值）、冻结窗格、自动筛选、打印设置（方向/纸张/fitTo/printArea）。
- ✅ pptx：原生图表（bar/line/pie，字面量缓存 + `dataEditable:false` 诚实告知）、切换动画（fade/push/wipe + 时长）、预设几何（ellipse/triangle/diamond/arrow/roundRect + line 连接线、翻转）、z 序（front/back/forward/backward）。
- ⏭ 留给 M4 的种子：嵌入式图表工作簿（PowerPoint 内可编辑数据）、动画、尾注、docx 多节插入、大工作簿**写入**流式、跨 run find/replace、数据验证、连接线/组合。

### M4 — 查找替换 + 能力深化（已交付，v0.5.0）

- ✅ 三格式共享 find/replace 契约：`{"op":"replace","path":"<scope>","props":{find, replace, regex?, matchCase?, wholeWord?}}`；docx/pptx 对段落拼接文本匹配（跨 run 安全，重写保留首个受影响 run 的格式）；regex 走 .NET Regex + 2 秒匹配预算（超时 → `invalid_args`）；零命中 = ok + `find_no_match` warning；逐 op 回报 `{replacements, locations≤20}`；docx `track:true` 时逐命中生成 w:del+w:ins 修订对（body 作用域）。
- ✅ 文档级展开 + CLI 糖：replace op 的路径 `"/"` 在命令层（CLI edit 与 MCP office_edit 共用 `ReplaceSugar`，与 `CrossDocDataFrom` 同位）展开为缺省作用域——docx `/body`+全部 `/header[i]`+`/footer[i]`（track 时仅 body）、每个 sheet、每页 `/slide[i]`（含备注）——并把逐作用域结果聚合为文档级 `{replacements, locations}`、坍缩 `find_no_match` warning；CLI 糖 `edit f --find X --replace Y [--regex] [--match-case] [--whole-word]`。
- ✅ docx：目录 TOC（SdtBlock + TOC 域，标题扫描生成超链接条目，`toc_pages_unknown` 诚实告知页码需 Word 重算）、文本水印（页眉 VML，逐页眉写入、无页眉自动建）、尾注（EndnotesPart）、插入分节符（`add type:sectionBreak`，多节文档逐节页面设置）。
- ✅ xlsx：批量 2D 写入（锚点式范围推断 / 区域式精确匹配；>50k 单元格写空白 sheet 走 OpenXml SAX 流式，等值律测试钉死两条路径产物一致）、行列插删（公式引用自动重写）/行高列宽/隐藏（`col[C]` 字母寻址）、单元格批注（经单元格寻址的 add/get/remove）。
- ✅ pptx：嵌入式图表工作簿（新图表默认嵌入 + `c:numRef`/`c:strRef` 引用，PowerPoint「编辑数据」可用；旧图表 `embedData:true` 退化改造）、进入动画（appear/fade/flyIn/wipe 手写 `p:timing` 树，`/slide[i]/animation[k]` 寻址）、经典批注（SlideCommentsPart + 作者去重）。
- ✅ 发版自动化：`.github/workflows/release.yml` —— push `v*` tag → build -warnaserror → 全量 test → 6 rid 单文件 self-contained publish → `SHA256SUMS` → `gh release create`（说明由 `scripts/release-notes-template.md` 渲染）。
- ⏭ 留给 M5 的种子：pptx 批注回复、xlsx 现代线程批注、大工作簿**就地写入**流式、插件机制、SmartArt 读取、动画预设扩容（强调/退出/motion path）、数据验证、连接线/组合。

### M5 — markdown/csv 桥 + 能力深化（已交付，v0.6.0）

- ✅ **markdown/csv 桥**（AI 原生旗舰）：`IFormatHandler` 增量获得 `CreateFrom(ctx, sourcePath)` 默认接口实现（无导入器格式 = `unsupported_feature` + workaround）；命令层路由器 `AIOffice.Mcp.Bridge`（CLI `create --from` 与 MCP `office_create.from` 共用）按源扩展名路由并校验导入矩阵（`.md/.markdown → .docx，.csv/.tsv → .xlsx`），源路径先过沙箱再进 handler。
  - docx 导入：Markdig（BSD-2-Clause）GFM AST 走查——标题→Heading 样式、粗/斜/删/行内代码 run、嵌套列表（复用 M3 编号机制）、管道表格（表头加粗）、链接/图片（沙箱解析）/引用/代码块/分隔线；raw HTML 与缺图降级 warning。导出：`read --view markdown` 把 body 写回 GFM，结构与导入 round-trip（测试钉死）。
  - xlsx 导入：RFC 4180 解析（引号/嵌入逗号换行）、分隔符嗅探（`, ; tab |`）或 `delimiter` 强制；类型化与单格 set 一致，**前导零码保文本**；>50k 单元格复用 M4 SAX 流式写。导出：`read --view csv [--sheet][--range]`。
  - 桥接视图用错格式 → `unsupported_feature` 并列出该格式有效视图（`Bridge.GuardBridgeView`）。
- ✅ docx：深表格（表级 borders/shading/headerRow/width/columnWidths/alignment/cellPadding + 单元格 mergeRight/mergeDown/valign，HTML 渲染真 colspan/rowspan）、域（pageNumber/numPages/date/docTitle + leadingText，「Page X of Y」页脚）、首页/奇偶页眉页脚变体（`/header[firstPage]`、`/header[even]`，自动接线 w:titlePg / w:evenAndOddHeaders）。
- ✅ xlsx：数据验证（list 下拉——字面值或 sourceRange；wholeNumber/decimal/date/textLength + 算子；allowBlank/提示语/errorStyle）、迷你图（line/column/winLoss + color/markers，x14 extLst）、线程批注（真 xl/threadedComments + persons part，`reply` op，legacy note 退化镜像）、单元格超链接（外链/`#Sheet!A1` 内链 + tooltip）。
- ✅ pptx：原生表格（a:tbl 手写——rows×cols/headerRow/columnWidths + light/medium/dark 直绘样式 + mergeRight/mergeDown，SVG 渲染真网格）、强调动画（pulse/grow/spin/colorPulse）与退出动画（fadeOut/flyOut/wipeOut，尾帧 hide set）、批注回复线程（p15 threadingInfo/parentCm）、SmartArt 只读（structure/get 输出连接序节点树）。
- ⏭ 留给 M6 的种子：大工作簿**就地写入**流式、插件机制（外部格式 handler）、公式（OMML）、RTL 深化、现代 pptx 线程批注 part、动画时间轴编辑（效果链/motion path）、连接线/组合。

### M6 — 深水区攻坚（已交付，v0.7.0）

- ✅ **docx 公式**（AI 原生旗舰）：自研 LaTeX → OOXML Math（OMML）转换器（`AIOffice.Word.Equations`，**无 LaTeX 依赖**）——词法器 + 递归下降解析器 + OMML 发射器，覆盖分式（frac/dfrac/tfrac/cfrac）、根式、上下标、大型算符（sum/prod/int/lim 带 `_`/`^`）、矩阵环境（pmatrix/bmatrix/Bmatrix/vmatrix/Vmatrix/matrix，`&` 分列 `\\` 分行）、`\left…\right` 定界、重音（bar/overline）、文本 run（text/mathrm/mathbf/mathit/operatorname）、希腊字母与算符关系符号。`add type:equation`（latex + display）：行内追加 `m:oMath`（`/body/p[i]/omath[j]`），display 出居中 `m:oMathPara` 块。**未知命令降级为字面 run + `equation_partial` warning**（文件仍过校验）。原始 LaTeX 以 `mc:Ignorable` 厂商属性（`urn:aioffice:equation`）存档，`get` 忠实回读、round-trip 字节一致（命名空间在根 `w:document` 一次声明，按 SDK 规范化序，保证首存即匹配重存形态）。`read --view text` 出 `$…$`/`$$…$$` 标记。
- ✅ **docx RTL/双向**：段落 `w:bidi`（同时右对齐）、run `w:rtl`、表格 `w:bidiVisual`（镜像列序），`rtl` bool prop，get 回报。
- ✅ **docx 多栏分节**：`/section[1]` 的 `columns`/`columnGap`（等宽多栏），`add type:columnBreak`（`w:br w:type="column"`）。
- ✅ **xlsx 就地流式写入**（旗舰）：`ExcelStreamingWrite.cs`——`stream:true` 或 >20 MB 文件触发；`TryPlan` 判定整批均为可流式 op（`set value` / `set values` 批量 / 公式串）时，经 SAX writer **原地改写**既有大工作簿目标 sheet part，否则整批回退 ClosedXML DOM；公式写入无缓存值（Excel 打开重算）；50 MB+ 工作簿就地改写实测秒级、内存有界。
- ✅ **xlsx Excel 表（ListObject）**：`IXLRange.CreateTable` / `IXLTable`（ExcelTables.cs）——name + 内置样式（XLTableTheme，短名 medium2 映射）+ 汇总行（XLTotalsRowFunction）+ bandedRows/Columns；结构化引用 `=SUM(Sales[Amount])` 求值（表先于保存入模型）；`/Sheet1/table[@name=X]` get；remove 撤销表对象但**保留单元格数据**。
- ✅ **xlsx 大纲分组**：`IXLRows.Group`/`Collapse`/`Ungroup`（ExcelGroups.cs）——`add type:group` 套 `row[a]:row[b]`/`col[a]:col[b]` 跨段、`collapsed`、嵌套抬升大纲级；remove 反分组一级；行/列 get 与 structure 回报 outlineLevel/collapsed。
- ✅ **pptx 母版/版式编辑**：`PptxMasters.cs`——`set /master[m]`（background + accent1..6 主题强调色）、`set /master[m]/layout[l]`（background）、`add type:layout`（克隆既有版式，props.basedOn/name）、母版/版式形状复用 slide 形状 op；克隆版式经 `add type:slide props:{layout:N}` 套用。
- ✅ **pptx 幻灯片分节**：`p14:sectionLst`（presentation.xml extLst，`PptxSections.cs`）——`add type:section`（`/` 根，name + afterSlide 0-based）、`set`/`remove`（片留存）、按 sldId 跟踪抗重排；`read --view outline` 按节分组。
- ✅ **pptx 幻灯片尺寸**：`PptxSlideSize.cs`——`set /` `{slideSize}` 命名预设（16:9/4:3/16:10/A4/letter）或显式 `{width,height}`，改写 `p:sldSz`；`get /` 回报尺寸 + slideCount/sectionCount。
- ✅ **pptx 动画时间轴**：`move /slide[i]/animation[k] before/after …` 重排、`set` 重调既有动画 props。
- ✅ **Core 寻址扩展**（接到真实命令面的关键一步）：`DocPath` 新增 `/`（零段根路径，`IsRoot`）与元素跨段 `ElementSpan`（`row[a]:row[b]`/`col[a]:col[b]`）。M6 能力在格式层已实现并测试，但 `EditOp.ParseBatch` → `DocPath.Parse` 网关此前拒收这两种新形式（仅直构 EditOp 的格式层测试可达）；本次让它们通过网关，CLI/MCP 命令面直达，且 docx/xlsx 对 `/` 根 op 诚实返回 `unsupported_feature` 而非崩溃。
- ⏭ 留给 M7 的种子：pptx/xlsx 公式（幻灯片/单元格内 OMML，复用本转换器）、插件机制（外部格式 handler）、现代 xlsx 批注打磨、动画预设++（完整强调/退出全集、效果链、motion path）、OLE 对象、无障碍/alt-text 审计。

### M7 — 能力深化（规划）

- 以 `docs/PARITY.md` 为账本继续清零（M1/M2 余项合并进 M7 窗口），或显式标记"不做 + 理由"。
- pptx/xlsx 公式、插件机制、动画预设扩容、OLE 对象、无障碍审计。

---

## 附：术语对照

| 术语 | 含义 |
|---|---|
| envelope | §2.2 的统一 JSON 输出对象 |
| rev | 文件字节 SHA256 前 12 hex，乐观锁令牌 |
| op | `edit --ops` 数组中的一条操作 |
| 规范路径 | `query`/`get` 返回的、可直接复用的寻址字符串 |
| 快照环 | 每文件保留最近 20 份前像的回滚池 |
| workaround 建议 | `unsupported_feature` 必须携带的替代路径说明 |
