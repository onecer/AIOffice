# AIOffice MCP Server 规格（`aioffice mcp`）

> 状态：1.3.0 规格（第三个 1.0 之后功能版本，纯**增量**，`surfaceVersion` 保持 **`1.0`**）：**工具数维持 17**，op **kinds 不变**。`office_edit` 新增 `add type` 值：`shape`/`textBox`/`formField`（docx）、`model3d`（pptx 3D 模型，`.glb`/`.gltf` 嵌为 3DModel 媒体部件 + 海报回退，`src`/`poster` 沙箱）；新增 prop keys：xlsx+pptx 图表润色 7 族（`dataLabels`/`legend`/`axisTitles`/`trendline`/`errorBars`/`gridlines`/`secondaryAxis`，`add` 与 `set /…/chart[k]` 双入口，`get` 报 `polish`）、xlsx `conditionalFormat` 新增 `formula`/`topBottom`/`aboveBelowAverage` 三类、xlsx 透视 `calculatedFields`、pptx 动画 `motionPath` 效果（`path` line/arc/circle/custom + custom `points`）；新增寻址 `/theme`（docx，`set`/`get` 主题色彩+字体方案）、`/body/shape[i]`/`/body/textBox[i]`/`/formField[@name=…]`（docx）、`/slide[i]/model3d[@id=N]`（pptx）。新增警告码 `model3d_as_media`；新增 `office_help` 主题 `chart-polish`/`conditional-format`/`themes`/`3d-models`/`form-fields`/`animations`。这些都是枚举/prop/寻址增量，新表面详情入 `office_help`、工具 schema 不膨胀，schema 仍在 ≤ 3500 token 预算内（§2.1），由 `SchemaConsistencyTests` 与 token-budget 测试守卫。此前 1.2.0 规格（第二个 1.0 之后功能版本，纯**增量**，`surfaceVersion` 保持 **`1.0`**）：**工具数维持 17**，op **kinds 不变**（`group`/`ungroup` 是 `add` **type** 值而非新 op kind）。`office_edit` 新增 `add type` 值：`smartart`/`connector`/`group`/`ungroup`（pptx）、`tableOfFigures`/`indexEntry`/`index`/`mergeField`（docx）、`formControl`（xlsx）；新增 `set` props：cell `locked` + sheet/workbook 保护标志（`protected`/`protectStructure`/`protectWindows`/`password`/`allow*`）+ `numberFormat` 命名预设值（`accounting-usd`/`currency-*`/`percent`/`scientific`/`date-iso`…）。`office_read` 的 `structure` 视图（docx）增列 `tablesOfFigures`/`indexes`/`mergeFields`、`fields` 视图增列合并域——复用既有视图名，不新增视图。新增警告码 `figures_cached`/`index_cached`；新增 `office_help` 主题 `smartart`/`connectors`/`number-formats`/`structural-fields`。这些都是枚举/prop 增量，schema 仍在 ≤ 3500 token 预算内（§2.1，实测 ~2970 token），由 `SchemaConsistencyTests` 与 token-budget 测试守卫。此前 1.1.0 规格（首个 1.0 之后功能版本，纯**增量**，`surfaceVersion` 保持 **`1.0`**）：**工具数维持 17**。`office_edit` 新增 `add type` 值 `source`/`citation`/`bibliography`（docx 引用与参考文献）与 `media`（pptx 音视频）；xlsx `conditionalFormat` 新增 `iconSet` 类型；xlsx 与 pptx `chart` 的 `kind` 新增 `doughnut`/`radar`/`bubble`/`stackedBar`/`percentStackedBar`/`stackedArea`/`combo`；docx run 与 pptx 形状 `set` 新增 `shadow`/`glow`/`reflection`/`outline` 效果 prop；pptx 切换新增 `split`/`reveal`/`cut`/`zoom`。`office_read` 新增 `sources` 视图（docx 文献库）；新增警告码 `bibliography_cached`。这些都是枚举/prop 增量，schema 仍在 ≤ 3500 token 预算内（§2.1）；新增主题 `docx/citation`/`docx/effect`/`pptx/media`/`pptx/effect` 在 `office_help`。此前 M10 规格（v0.11.0，1.0 前最后一个功能里程碑）：**工具数维持 17**——嵌入对象与 pptx 公式搭载在既有 `office_edit`/`office_read`/`office_get` 上，不新增工具。`office_edit` 新增 `extract` op kind（把嵌入对象的字节导出到沙箱目标，产出型——不改源文档）与 `add type:embed`（任意文件作 OLE/包对象：docx `/body`、xlsx `/Sheet1`、pptx `/slide[i]`）/`add type:equation`（docx 已有，M10 起 **pptx** 亦可——同一份共享 LaTeX→OMML 转换器，`/slide[i]/shape[@id=N]/omath[k]`；xlsx 返回 `unsupported_feature`）；`office_read` 新增 `embeds` 视图；`properties` 视图三格式**统一**为 `data.properties.{core,custom}`；新寻址 `/embed[i]`、`/Sheet1/embed[i]`、`/slide[i]/embed[@id=N]`、`/slide[i]/shape[@id=N]/omath[k]`。`office_schema` 与 `office_status` capabilities 新增稳定的 **`surfaceVersion`**（`1.0-rc`，面向 AI 的契约版本，见 [CONTRACT.md](../CONTRACT.md)）；schema 仍在 ≤ 3500 token 预算内（§2.1）。此前 M9 规格（v0.10.0，pre-1.0 capstone）：**工具数 16 → 17**——新增 `office_convert`（跨格式互转 `{src, dest}`：docx/xlsx/pptx ↔ 互转经内容中立模型 `INeutralConvertible`，docx↔md / xlsx↔csv 复用文本桥，任意↔md 经命令层 `NeutralMarkdown` 序列化器，any→pdf/png/svg/html/txt 经 render 层；跨格式有损，丢失项汇入 `data.dropped` + 单条 `convert_lossy` 警告；目标新建，同扩展名 → `invalid_args`，未知目标 → `unsupported_feature`），以及 `office_status` 的 `capabilities` 自省块（动词与工具数、格式、convert 源/目标、render 目标、audit 类别）；加入 office_convert 后 schema 仍在 ≤ 3500 token 预算内（§2.1）。此前 M8 规格（v0.9.0）：**工具数 15 → 16**——新增 `office_diff`（语义对比文档与基线——另一同格式文件 `other` 或文档自身快照 `snapshot:N`，返回已排序、确定性的 added/removed/modified/moved 变更清单，变更是数据，`view:summary|detailed`），以及 `office_edit` 的 docx `add type:caption`（题注，Figure/Table/Equation + SEQ）/`add type:crossRef`（交叉引用 REF/PAGEREF）、xlsx `add type:slicer`（表列 / 透视字段切片器）、pptx shape `set {hyperlink}`（外链 url / `#slide:N` 跳转 / `#first…#end` 放映动作）/`{linkText}`；新寻址 `/caption[@label=Figure][i]`、`/Sheet1/slicer[i]`；加入 office_diff 后 schema 仍在 ≤ 3500 token 预算内（§2.1）。此前 M7（v0.8.0）：**工具数 14 → 15**——新增 `office_audit`（无障碍 + 质量 lint，findings 是数据，`--fix` 仅安全自动修复），以及 `office_read` 的 `properties`/`fields`/`styles` 视图、`office_edit` 的 `set /properties`（文档属性 core + 类型化 custom）、`add type:contentControl`（docx 内容控件）、`add type:cellStyle`（xlsx 命名单元格样式）、shape `altText`/`altTitle`（pptx）；新寻址 `/properties`、`/sdt[@tag=X]`、`/style[@name=X]`。此前 M6（v0.7.0）：工具数维持 14，深水区只扩 `office_edit` 的 add type 词表与寻址（`equation`/`columnBreak`/`section`/`layout`/`group`/`table`，`/`、`/body/p[i]/omath[j]`、`row[a]:row[b]`、可编辑 `/master[1]/layout[i]`）；M5：`office_create` 的 `from` + `office_read` 的 `markdown`/`csv` 视图；M2：`track`/`author` + accept/reject op + revisions/comments/styles 视图 + `file_too_large` 错误码。
> 实现位于 `src/AIOffice.Mcp/`，基于官方 C# MCP SDK（NuGet `ModelContextProtocol`），stdio 传输，100% 自研——OOXML 读写由 `DocumentFormat.OpenXml` / `ClosedXML` 完成，**无外部引擎、无网络、无二进制下载**。
> MCP 工具与 CLI 动词 1:1 镜像同一内部命令层（`AIOffice.Core`，one source of truth）：MCP 工具 = 参数校验 → 内部命令 → JSON envelope。一个心智模型，两个入口。
> 说明文字为中文，所有 schema / 字段名 / 错误码为英文。

---

## 0. 全局约定

### 0.1 启动

```bash
aioffice mcp [--workspace <dir>]
```

- stdio MCP server；`--workspace`（或环境变量 `AIOFFICE_WORKSPACE`，默认 cwd）定义沙箱根。
- 所有 `file` / `output` 参数在每次调用时做 realpath + symlink 逃逸检查，越界即 `sandbox_denied`。
- 17 个工具：`office_create` `office_read` `office_query` `office_get` `office_edit` `office_render` `office_validate` `office_audit` `office_diff` `office_template` `office_convert` `file_snapshot` `office_status` `office_help` `office_schema` `preview_open` `preview_selection`（`office_convert` M9 加入，见 §1.7c；`office_diff` M8 加入，见 §1.7b；`office_audit` M7 加入，见 §1.7a；preview 二件套 M1 转正注册，见 §1.13）。

### 0.2 统一返回 envelope

每个工具的 result 都是一段 JSON 文本（MCP `content[0].type="text"`），与 CLI stdout 完全同构：

```jsonc
{
  "ok": true,
  "data": { /* per-tool, see below */ },
  "error": {                          // only when ok=false; otherwise null
    "code": "invalid_path",
    "message": "No element at /body/p[99]",
    "suggestion": "Document has 12 paragraphs; call office_read view=outline or office_query \"p\" to list them.",
    "candidates": ["/body/p[11]", "/body/p[12]"]   // optional, from automatic server-side nearest-match query
  },
  "meta": {
    "file": "report.docx",            // when a file was involved
    "rev": "a3f9c12be01d",            // first 12 hex of SHA256 of file bytes, AFTER the call
    "elapsedMs": 42,
    "version": "0.3.0",               // aioffice version
    "warnings": [                     // optional, non-fatal
      { "code": "formula_not_evaluated", "message": "B7 has no cached value; returned formula text" }
    ]
  }
}
```

规则：

- `error.suggestion` 永不为空（`AiofficeError` 构造函数强制非空）。
- `rev` = 文件字节 SHA256 的前 12 个 hex 字符；**所有触文件的命令**都在 `meta.rev` 返回当前值。
- 变更类调用（`office_edit`、`office_template` 原地合并、`file_snapshot restore`）可传 `expect_rev` 做乐观锁，失配 → `stale_address`，**在任何写入发生之前**拒绝。
- 每次成功的变更前自动写快照（环形 20 份，见 §1.9），随时可回滚。
- 非致命问题（如 xlsx 公式无缓存值）走 `meta.warnings`，不打断调用。

### 0.3 寻址语法（自有设计，详情见 `office_help {topic:"addressing"}`）

| 格式 | 示例 |
|---|---|
| docx | `/body/p[3]`、`/body/table[1]/tr[2]/tc[1]`、`/body/p[3]/run[2]`、`/header[1]/p[1]`、`/footer[1]/p[1]`；M2 起 `/revision[@id=3]`、`/comment[@id=2]`、`/styles`、`/style[@id=Callout]` |
| xlsx | `/Sheet1/A1`、`/Sheet1/A1:C10`、`/Sheet1/row[3]`、`/Sheet1/chart[1]`，含空格的表名加引号：`/'Q3 Data'/B2`；M2 起 `/Pivot/pivot[@name=X]`、`/Sheet1/conditionalFormat[1]`、`/Sheet1/image[1]` |
| pptx | `/slide[2]`、`/slide[2]/shape[3]`、`/slide[2]/shape[3]/p[1]`；M1 起 `/master[1]`、`/master[1]/layout[2]`（get/query 只读）；M2 起 `/slide[2]/notes` |

索引一律 **1-based**。`office_query` 返回规范路径（canonical paths），`office_get` / `office_edit` 原样接受。位置索引在增删后会漂移——多步编辑中，删改之后**重新 query**，不要复用旧索引。

### 0.4 错误码总表（全集 10 个 + 1 个 warning）

| code | 含义 | 自愈线索 | CLI exit |
|---|---|---|---:|
| `invalid_args` | 参数/输入错误（坏 selector 语法、未知 help 主题、不存在的快照编号、模板 data 非法、目标已存在且未 `overwrite`…） | `suggestion` 给出正确形态；未知主题/编号附 `candidates[]` | 2 |
| `file_not_found` | 文件不存在 | `suggestion` 提示 workspace 根与相对路径写法 | 2 |
| `sandbox_denied` | 路径越出 workspace（含 symlink 逃逸） | `suggestion` 提示 `--workspace` / `AIOFFICE_WORKSPACE` | 4 |
| `invalid_path` | 元素路径不存在 | **必附 `candidates[]`**：服务端自动用路径末段元素名跑一次近似 query 回填 | 2 |
| `stale_address` | `expect_rev` 与当前 rev 失配（文件被外部或并行修改） | 重新 `office_get` 拿新 rev 后重试；写入未发生 | 2 |
| `unsupported_feature` | 能力在当前里程碑不可用（如 master/layout 编辑、scatter 图表） | **`suggestion` 必须给出 workaround** | 5 |
| `file_too_large` | 文件超过 `AIOFFICE_MAX_FILE_MB` **opt-in** 上限（M3 起默认不限大小；M2 时默认 50MB） | `suggestion` 提示调高/取消环境变量或拆分文件 | 2 |
| `format_corrupt` | 文件不是合法 OOXML / zip 损坏 | `suggestion` 提示 `office_validate` 与 `file_snapshot restore` | 3 |
| `internal_error` | 未预期异常（我们的 bug） | `suggestion` 提示带 `office_status` 输出报 issue | 3 |
| `formula_not_evaluated` | **warning（`meta.warnings`），非 error**：xlsx 公式无缓存值且本次未计算 | 读到的是公式文本而非值；data 中相应字段标注 | 0 |

> `office_validate` 发现文档有 schema 错误时仍是 `ok:true`（校验本身跑成功了）——错误码只描述「工具没跑成」，不描述「文档质量差」。

---

## 1. 工具规格

> 每个工具：用途 → inputSchema → result（`data` 形状）→ 示例 → 映射 CLI → 可能错误码。
> 示例 result 一律省略 `meta`（结构同 §0.2）。

---

### 1.1 `office_create`

**用途**：新建空白 .docx / .xlsx / .pptx（类型由扩展名推断），或经 `from`（M5）从 markdown / csv 导入。

```json
{
  "type": "object",
  "properties": {
    "file": { "type": "string", "description": "Target path ending in .docx/.xlsx/.pptx, inside workspace; kind inferred from extension" },
    "from": { "type": "string", "description": "Import source: .md/.markdown -> .docx (headings/lists/tables/links/code), .csv/.tsv -> .xlsx (typed cells; leading-zero codes stay text). Mismatched pairs fail with the valid matrix" },
    "kind": { "type": "string", "enum": ["docx", "xlsx", "pptx"], "description": "Override when extension is non-standard" },
    "title": { "type": "string", "description": "Document title written to core properties" },
    "overwrite": { "type": "boolean", "default": false }
  },
  "required": ["file"]
}
```

**result**：`data: { created: string /* absolute path */, kind: "docx"|"xlsx"|"pptx" }`；带 `from` 时附 `source`（绝对路径）与逐格式导入统计（docx：`format:"markdown"`；xlsx：`sheet/delimiter/rows/columns/cells/streamed`）。

**示例**

```json
// call
{ "file": "out/q4.pptx", "title": "Q4 Review" }
// result
{ "ok": true, "data": { "created": "/ws/out/q4.pptx", "kind": "pptx" } }
// call（M5 导入桥）
{ "file": "report.docx", "from": "notes.md" }
// result
{ "ok": true, "data": { "created": "/ws/report.docx", "kind": "docx", "source": "/ws/notes.md", "format": "markdown" } }
```

**映射 CLI**：`aioffice create <file> [--from notes.md|data.csv] [--kind docx|xlsx|pptx] [--title T]`

**错误码**：`invalid_args`（扩展名无法推断且未传 `kind`；目标已存在且未 `overwrite`；`from` 源/目标不匹配——suggestion 给全矩阵 `.md/.markdown → .docx, .csv/.tsv → .xlsx`）、`file_not_found`（`from` 源不存在）、`sandbox_denied`（目标或 `from` 源越界）、`internal_error`。

---

### 1.2 `office_read`

**用途**：读文档——大纲 / 纯文本 / 统计 / 完整结构树。理解一个文件的第一步。

```json
{
  "type": "object",
  "properties": {
    "file": { "type": "string" },
    "view": { "type": "string", "enum": ["outline", "text", "stats", "structure", "properties", "embeds", "revisions", "comments", "styles", "markdown", "csv"], "default": "outline",
      "description": "outline: headings/slides/sheets skeleton with paths; text: plain text; stats: counters; structure: full element tree with paths+types; revisions/comments/styles: docx only (M2); markdown: docx body as GFM (M5, round-trips office_create from); csv: one xlsx sheet as RFC 4180 csv (M5)" },
    "range": { "type": "string", "description": "Scope limit 'a..b' (1-based): paragraphs for docx, slides for pptx, rows for xlsx; csv view also takes 'A1:C10'" },
    "sheet": { "type": "string", "description": "csv view: sheet name (default: first sheet)" },
    "max_bytes": { "type": "integer", "description": "Cap payload size; truncation reported in meta.warnings and data.truncated" }
  },
  "required": ["file"]
}
```

**result**：随 view 变化——`outline`: `{ outline: [{ path, kind, text?, children? }] }`；`text`: `{ text: string, truncated: boolean }`；`stats`: `{ paragraphs?, words?, tables?, sheets?, usedRange?, slides?, shapes? }`；`structure`: `{ tree: Node, truncated: boolean }`；`markdown`（M5）: `{ view:"markdown", markdown: string }`；`csv`（M5）: `{ view:"csv", sheet, range?, content, truncated }`。桥接视图用错格式（如对 xlsx 要 markdown）→ `unsupported_feature`，suggestion 列出该格式的全部有效视图。

**示例**

```json
// call
{ "file": "report.docx", "view": "outline" }
// result
{ "ok": true, "data": { "outline": [
  { "path": "/body/p[1]", "kind": "Heading1", "text": "2026 年度计划" },
  { "path": "/body/p[8]", "kind": "Heading2", "text": "预算" },
  { "path": "/body/table[1]", "kind": "table", "text": "3x4" }
] } }
```

**映射 CLI**：`aioffice read <file> [--view outline|text|stats|structure|revisions|comments|styles|markdown|csv] [--range a..b] [--sheet NAME] [--max-bytes N]`

**错误码**：`file_not_found`、`sandbox_denied`、`format_corrupt`、`invalid_args`（坏 range / 未知 view，candidates 列出全部）、`unsupported_feature`（桥接视图用错格式）。

---

### 1.3 `office_query`

**用途**：CSS-like 选择器全文档检索，返回匹配元素的规范路径——后续 `get` / `edit` 的「寻址」入口。

```json
{
  "type": "object",
  "properties": {
    "file": { "type": "string" },
    "selector": {
      "type": "string",
      "description": "CSS-like selector. Examples: \"p[style=Heading1]\", \"cell[value>100]\", \"shape:contains('Q3')\", \"run[bold=true]\", \"table tr[1]\". Full grammar: office_help {topic:\"selectors\"}. Returns canonical paths usable in office_get/office_edit."
    }
  },
  "required": ["file", "selector"]
}
```

**result**：`data: { results: [{ path, type, text?, props? }], total: number, truncated: boolean }`（最多返回 50 条；`total` 是全量命中数，超出时 `truncated:true` 并在 suggestion 提示收窄 selector）。

**示例**

```json
// call
{ "file": "deck.pptx", "selector": "shape:contains('Q3')" }
// result
{ "ok": true, "data": { "results": [
  { "path": "/slide[3]/shape[2]", "type": "shape", "text": "Q3 risk summary" }
], "total": 1, "truncated": false } }
```

**映射 CLI**：`aioffice query <file> <selector>`

**错误码**：`invalid_args`（selector 语法错误，suggestion 给出语法提示并指向 `office_help {topic:"selectors"}`）、`file_not_found`、`sandbox_denied`、`format_corrupt`。Warning：`formula_not_evaluated`（xlsx 值谓词遇到无缓存值的公式单元格）。

---

### 1.4 `office_get`

**用途**：按路径精读单个节点及其属性（编辑前「看清楚再下手」）。

```json
{
  "type": "object",
  "properties": {
    "file": { "type": "string" },
    "path": {
      "type": "string",
      "description": "Canonical 1-based path, e.g. \"/body/p[3]\", \"/Sheet1/B2\", \"/'Q3 Data'/A1:C10\", \"/slide[2]/shape[3]\"; \"/\" for document-level properties. Grammar: office_help {topic:\"addressing\"}"
    }
  },
  "required": ["file", "path"]
}
```

**result**：`data: { node: { path, type, text?, props: object, children: string[] /* child canonical paths, one level */ } }`

**示例**

```json
// call
{ "file": "data.xlsx", "path": "/Sheet1/B2" }
// result
{ "ok": true, "data": { "node": { "path": "/Sheet1/B2", "type": "cell",
  "props": { "value": "42", "valueType": "number", "formula": "=SUM(A1:A10)", "numberFormat": "0.00" },
  "children": [] } } }
```

**映射 CLI**：`aioffice get <file> <path>`

**错误码**：`invalid_path`（**自动 candidates**：服务端用路径末段元素名跑一次近似 query，把最接近的规范路径塞进 `candidates[]`）、`file_not_found`、`sandbox_denied`、`format_corrupt`。Warning：`formula_not_evaluated`。

---

### 1.5 `office_edit`（fat tool — 所有变更走这里）

**用途**：增 / 改 / 删 / 移 / 查找替换 / 修订裁决，一次调用 = 一个**原子**保存周期：全部 ops 依序成功才落盘，任何一条失败则整批不写。自动快照 + rev 守卫。M2 起支持 docx 修订（`track`/`author` + `accept`/`reject` op）；M4 起支持 `replace` op（三格式共享契约）。

```json
{
  "type": "object",
  "properties": {
    "file": { "type": "string" },
    "ops": {
      "type": "array",
      "minItems": 1,
      "description": "Atomic batch, applied in order; all-or-nothing single save. Use office_help {topic:\"<fmt>/<element>\"} for exact prop names and element types — do NOT guess.",
      "items": {
        "type": "object",
        "properties": {
          "op": { "type": "string", "enum": ["set", "add", "remove", "move", "replace", "accept", "reject", "extract"],
            "description": "accept/reject resolve docx tracked revisions (path: /revision[@id=N] or a scope like /body). replace = find/replace in scope: props {find,replace,regex?,matchCase?,wholeWord?}; path \"/\" = whole document (docx body+headers+footers, every sheet, every slide+notes); 0 matches -> ok + find_no_match warning. extract writes an embedded object's bytes out: {op:extract, path:/embed[1] (or /Sheet1/embed[1], /slide[2]/embed[1]), props:{to:WORKSPACE_DEST}} -- a producing op, does NOT modify the source" },
          "path": { "type": "string", "description": "set/remove/move: target element. add: PARENT element, e.g. \"/body\", \"/slide[2]\", \"/Sheet1\". replace: container scope or \"/\". extract: the embed path" },
          "type": { "type": "string", "description": "add only: element type, e.g. paragraph, run, table (docx/pptx/xlsx ListObject), row, col, cell, slide, shape, image, comment, reply, note, style, header, footer, chart, pivot, conditionalFormat, toc, watermark, footnote, endnote, sectionBreak, equation (LaTeX->OMML; docx + pptx; xlsx -> unsupported_feature), embed (any file as an OLE/package object; props.src; docx /body, xlsx /Sheet1, pptx /slide[i]), columnBreak, animation, section (pptx), layout (pptx clone), group (xlsx outline), field, dataValidation, sparkline, source/citation/bibliography (docx, 1.1), media (pptx audio/video, 1.1), smartart/connector/group/ungroup (pptx, 1.2), tableOfFigures/indexEntry/index/mergeField (docx, 1.2), formControl (xlsx, 1.2), shape/textBox/formField (docx, 1.3), model3d (pptx 3D model, 1.3)" },
          "props": { "type": "object", "additionalProperties": { "type": "string" },
            "description": "String-valued props, e.g. {\"text\":\"Hi\",\"bold\":\"true\",\"size\":\"12pt\",\"fill\":\"FF0000\"}. Sizes unit-qualified (12pt, 2cm); colors hex/named. Table cells merge via {\"mergeRight\":\"2\"}/{\"mergeDown\":\"2\"}. pptx add chart: {\"dataFrom\":\"book.xlsx!Sheet1/A1:B5\"} pulls categories+series from a workbook (first col = categories, header row = series names) instead of literals" },
          "position": { "type": ["integer", "string"],
            "description": "add/move: 1-based index within parent, or \"before:<path>\" / \"after:<path>\"; omit = append" }
        },
        "required": ["op", "path"]
      }
    },
    "track": { "type": "boolean", "default": false, "description": "docx: record text set/add/remove ops as tracked revisions (w:ins/w:del); resolve later with accept/reject" },
    "author": { "type": "string", "description": "Author stamped on revisions/comments (op props.author wins; default: AIOFFICE_AUTHOR env, then \"AIOffice\")" },
    "expect_rev": { "type": "string", "description": "Optimistic lock: 12-hex rev from a previous meta.rev; mismatch fails with stale_address BEFORE any write" },
    "dry_run": { "type": "boolean", "default": false, "description": "Validate the whole batch without writing" }
  },
  "required": ["file", "ops"]
}
```

**result**：`data: { applied: number, results: [{ op, path, ok, createdPath? }], snapshot: number|null, dryRun: boolean }`
（`snapshot` = 本次自动写入的预改动快照编号，可直接喂给 `file_snapshot restore`；`dry_run` 时为 `null`；`add` 成功返回新元素的 `createdPath`。）

**示例**

```json
// call — 改标题样式 + 插入新段落 + 删空段，一次原子保存
{ "file": "report.docx", "expect_rev": "a3f9c12be01d", "ops": [
  { "op": "set", "path": "/body/p[1]", "props": { "style": "Heading1", "text": "2026 年度计划" } },
  { "op": "add", "path": "/body", "type": "paragraph", "position": "after:/body/p[1]",
    "props": { "text": "本文档为最终版。", "italic": "true" } },
  { "op": "remove", "path": "/body/p[7]" }
] }
// result
{ "ok": true, "data": { "applied": 3, "results": [
  { "op": "set", "path": "/body/p[1]", "ok": true },
  { "op": "add", "path": "/body", "ok": true, "createdPath": "/body/p[2]" },
  { "op": "remove", "path": "/body/p[8]", "ok": true }
], "snapshot": 7, "dryRun": false } }
```

> 注意示例最后一条：前面插入了一段，后续 ops 中的位置索引按**执行时点**解析（`/body/p[7]` 在插入后实际命中原第 7 段、现第 8 段）。同批内 ops 之间有索引依赖时，把目标排在插入/删除**之前**，或拆成两次 edit 用 query 重新寻址。

**跨文档 dataFrom（M3）**：目标是 pptx 时，`add chart` 的 props 可用 `{"dataFrom":"metrics.xlsx!Sheet1/A1:B5"}` 取代字面量 categories/series——命令层（CLI 与 MCP 共用）经 workspace 沙箱解析工作簿、用 xlsx handler 读取区域：首列 → 分类，表头行 → 系列名，其余列 → 系列值（空格 = 数据缺口）。区域写错 → 类型化错误 + candidates（表名 / 最近 usedRange）；展开发生在 rev 守卫与快照**之前**，坏数据源不会写盘。带空格的表名用引号：`book.xlsx!'Q3 Data'/A1:C5`。

**文档级查找替换（M4）**：`{"op":"replace","path":"/","props":{"find":"2025","replace":"2026"}}` 在命令层（CLI 与 MCP 共用 `ReplaceSugar`）展开为该格式的缺省作用域——docx `/body` + 全部 `/header[i]`/`/footer[i]`（`track:true` 时仅 body，tracked replace 是 body 契约）、xlsx 每个 sheet、pptx 每页 `/slide[i]`（slide 作用域缺省含演讲者备注）。逐作用域结果聚合为顶层 `data.replacements` + `data.locations`（≤20 条规范路径）；任一处命中则逐作用域的 `find_no_match` warning 全部吞掉，全文档零命中则坍缩为**一条**文档级 warning。展开同样发生在 rev 守卫与快照之前。CLI 等价糖：`aioffice edit f --find X --replace Y [--regex] [--match-case] [--whole-word]`。

**映射 CLI**：`aioffice edit <file> --ops <json|@file> [--track] [--author NAME] [--dry-run] [--expect-rev R]`；单 op 糖：`--set <path> k=v...`、`--add <path> --type T k=v...`、`--remove <path>`。

**错误码**：`stale_address`、`invalid_path`（带 candidates）、`invalid_args`（坏 op / 未知 type / 坏 position）、`unsupported_feature`（该元素类型或属性尚未实现，suggestion 给替代做法）、`file_not_found`、`sandbox_denied`、`format_corrupt`。

---

### 1.6 `office_render`

**用途**：把文档（或子树）渲染为可检视产物——「render → look → fix」循环的 look 步骤。docx/xlsx → html，pptx → 每页 svg（或 html）；`text` 全格式可用。**M1 起支持 `to=png`**：handler 产物（html/svg）→ headless 浏览器（Chrome/Edge，`office_status` 报告探测结果）截图；pptx 一次渲染一页（传 `scope`，缺省 `/slide[1]` + meta warning）。**M3 起支持 `to=pdf`**：同一浏览器管线 `--print-to-pdf --no-pdf-header-footer` 出分页 PDF——docx/xlsx A4 分页；pptx 整副 deck 一个 PDF、一页一片（`scope` 可缩到单页）。探测不到浏览器 → `unsupported_feature` + 安装建议。

```json
{
  "type": "object",
  "properties": {
    "file": { "type": "string" },
    "to": { "type": "string", "enum": ["html", "svg", "text", "png", "pdf"], "default": "html",
      "description": "html: docx/xlsx/pptx; svg: pptx, one file per slide; text: plain text; png: browser screenshot, written next to source (pptx: one slide, default /slide[1] — pass scope); pdf: paged print via local Chromium, written next to source (pptx: whole deck, one page per slide)" },
    "scope": { "type": "string", "description": "Render only this subtree, e.g. \"/slide[3]\", \"/Sheet1/A1:F20\", \"/body/table[1]\"" },
    "output": { "type": "string", "description": "Output file or directory inside workspace (default: alongside source)" }
  },
  "required": ["file"]
}
```

**result**：
- html/svg/text：`data: { outputs: string[] /* absolute paths */, content?: string /* inlined when single text-format output ≤ 256 KB */ }`
（单一 html/svg/text 产物且不超过 256 KB 时直接内联在 `content`，agent 无需再开文件；超限时只给路径并在 `meta.warnings` 标注。）
- png：`data: { format: "png", scope?: string, written: string /* absolute path */, sizeBytes: number }`，**且** MCP result 的 `content` 附第二个 block：`{type:"image", mimeType:"image/png", data:<base64>}`——文件字节原样内联（不降采样），模型直接看图。
- pdf：`data: { format: "pdf", scope?: string, written: string /* absolute path */, sizeBytes: number, pages?: number /* pptx：渲染的幻灯片页数 */ }`——PDF 是二进制且可能很大，**不**内联 image block，envelope 只带落盘路径。

**示例**

```json
// call — 只看第 3 页
{ "file": "deck.pptx", "to": "svg", "scope": "/slide[3]" }
// result
{ "ok": true, "data": { "outputs": ["/ws/deck.slide3.svg"], "content": "<svg width=\"1280\" height=\"720\">…</svg>" } }
```

```json
// call — png 截图（pptx 必须给 scope，否则默认 /slide[1] 并在 meta.warnings 提示）
{ "file": "deck.pptx", "to": "png", "scope": "/slide[2]" }
// result（content[0] 是这段 envelope 文本，content[1] 是 image/png block）
{ "ok": true, "data": { "format": "png", "scope": "/slide[2]", "written": "/ws/deck.png", "sizeBytes": 48213 } }
```

**映射 CLI**：`aioffice render <file> [--to html|svg|text|png|pdf] [--scope <path>] [-o out]`

**错误码**：`unsupported_feature`（`to=png|pdf` 且探测不到 Chrome/Edge → suggestion 给安装路径与 html/svg 替代）、`invalid_path`（坏 scope，带 candidates）、`file_not_found`、`sandbox_denied`、`format_corrupt`。

---

### 1.7 `office_validate`

**用途**：OpenXmlValidator schema 校验 + 自有 lint（空段落、悬空引用、混杂字体等），每条 issue 带 suggestion。每轮编辑后的廉价体检。

```json
{
  "type": "object",
  "properties": {
    "file": { "type": "string" }
  },
  "required": ["file"]
}
```

**result**：`data: { valid: boolean, schemaErrors: [{ part, path?, message }], issues: [{ code, path?, message, severity: "info"|"warn"|"error", suggestion }] }`

**示例**

```json
// call
{ "file": "data.xlsx" }
// result
{ "ok": true, "data": { "valid": true, "schemaErrors": [], "issues": [
  { "code": "empty_sheet", "path": "/Sheet3", "message": "Sheet3 has no used cells",
    "severity": "info", "suggestion": "Remove it via office_edit {op:\"remove\", path:\"/Sheet3\"} if unintended" }
] } }
```

**映射 CLI**：`aioffice validate <file>`

**错误码**：`file_not_found`、`sandbox_denied`、`format_corrupt`（zip 都打不开时；能打开但 schema 有错 → `ok:true, valid:false`）。

---

### 1.7a `office_audit`（M7）

**用途**：无障碍 + 质量 lint。findings 是**数据**，不是错误——即使有 error 级 finding，`ok:true`（exit 0），与 `office_validate` 同理（错误码只描述「工具没跑成」）。`fix:true` 只应用**安全、非破坏性**自动修复，再重审。

```json
{
  "type": "object",
  "properties": {
    "file": { "type": "string" },
    "category": { "type": "string", "enum": ["accessibility", "quality", "all"], "default": "all" },
    "severity": { "type": "string", "enum": ["error", "warning", "info"], "default": "info",
      "description": "Minimum severity to report" },
    "fix": { "type": "boolean", "default": false,
      "description": "Apply safe autofixes, then re-audit; result adds {fixed:N, remaining:[ids]}" }
  },
  "required": ["file"]
}
```

**result**：`data: { findings: [{ id: "code#path", severity: "error"|"warning"|"info", category: "accessibility"|"quality", code, path?, message, suggestion, autofixable }], summary: { errors, warnings, infos } }`；`fix:true` 时追加 `fixed: N` 与 `remaining: [id]`（重审后仍存在的 finding id）。

**code 全集**（详见 `office_help {topic:"audit"}`）：
- accessibility：`a11y_no_alt_text`(err,fix) · `a11y_no_table_header`(err,fix) · `a11y_no_doc_title`(warn,fix) · `a11y_no_slide_title`(warn,fix) · `a11y_heading_skip`(warn) · `a11y_low_contrast`(warn) · `a11y_tiny_font`(warn) · `a11y_merged_data_cells`(warn) · `a11y_reading_order`(info)。
- quality：`quality_broken_ref`(err) · `quality_formula_error`(err) · `quality_broken_link`(err) · `quality_empty_heading`(warn) · `quality_off_canvas`(warn) · `quality_empty_placeholder`(warn) · `quality_orphan_bookmark`(info,fix) · `quality_duplicate_id`(warn)。

安全自动修复：占位 alt 文本 `(describe this image)`、标记表头行、设文档/幻灯片标题（首个 Heading1 > 文件名 > 占位）、删孤儿书签。其余一律仅报告；`fix:true` 前自动建快照（可 `file_snapshot restore`）。

**示例**

```json
// call —— 故意做坏的 report.docx（图片无 alt + 无表头表格 + 无文档标题 + H1→H3 + 空标题）
{ "file": "report.docx" }
// result
{ "ok": true, "data": {
  "findings": [
    { "id": "a11y_no_alt_text#/body/p[5]", "severity": "error", "category": "accessibility",
      "code": "a11y_no_alt_text", "path": "/body/p[5]", "message": "An image has no alternative text (descr).",
      "suggestion": "Add a description … or --fix to insert a placeholder.", "autofixable": true }
    /* … */
  ],
  "summary": { "errors": 2, "warnings": 3, "infos": 0 } } }

// call —— fix:true
{ "file": "report.docx", "fix": true }
// result
{ "ok": true, "data": {
  "findings": [ /* 仅剩 report-only */ ],
  "summary": { "errors": 0, "warnings": 2, "infos": 0 },
  "fixed": 3,
  "remaining": ["a11y_heading_skip#/body/p[3]", "quality_empty_heading#/body/p[4]"] } }
```

**映射 CLI**：`aioffice audit <file> [--category …] [--severity …] [--fix]`

**错误码**：`invalid_args`（category/severity 非法）、`file_not_found`、`sandbox_denied`、`format_corrupt`、`unsupported_feature`（格式 handler 未实现审计——三格式均已实现）。

---

### 1.7b `office_diff`（M8）

**用途**：语义对比当前文档（`file`）与一份基线，返回**已排序、跨平台确定性**的变更清单。变更是**数据**，不是错误——哪怕两份文档差异巨大也 `ok:true`（exit 0），与 `office_validate`/`office_audit` 同理。基线 = `other`（另一**同格式**文件，跨格式 → `invalid_args`）**或** `snapshot`（`file` 自身的某个编辑前快照号），二者恰选其一（都缺/都给 → `invalid_args`）。读取-only，不改动 `file` 与基线。

```json
{
  "type": "object",
  "properties": {
    "file": { "type": "string", "description": "The current/new document" },
    "other": { "type": "string",
      "description": "Baseline: the OLD same-format document. A format mismatch is invalid_args. Exactly one of other/snapshot required" },
    "snapshot": { "type": "integer",
      "description": "Baseline: snapshot N of file from its auto pre-edit ring; a missing index is invalid_args naming the available numbers. Exactly one of other/snapshot required" },
    "view": { "type": "string", "enum": ["summary", "detailed"], "default": "detailed",
      "description": "detailed: full before/after per change; summary: counts + path+kind only" }
  },
  "required": ["file"]
}
```

**result**：`data: { changes: [{ kind: "added"|"removed"|"modified"|"moved", path, before?, after?, detail? }], summary: { added, removed, modified, moved }, baseline: "<file|snapshot N>", view }`。
- `path` 是**当前**文档的规范路径——`removed` 取**基线**路径（当前已不存在）。
- `modified` 的 `before`/`after` 给简洁的旧/新值（文本、`Normal`→`Heading1`、单元格值…）；`detail` 说明改了什么。
- `moved`：内容匹配但位置变了，`detail` = `moved from <old path>`（段落/幻灯片/行用 LCS/内容哈希区分 moved 与 added+removed）。
- `changes` 始终按 `(path, kind)` 序贯字典序排序，故同样两份文档在每个平台、每次运行结果一致。
- `view:summary` 把每条变更裁剪为 `{kind, path}`（`summary` 计数不变）。

**快照 diff 工作流**：每次成功 `office_edit` 都先自动快照前像，故 `office_diff {file, snapshot:1}` = 上次编辑造成的变更集——提交/分享前的快速评审。

**示例**

```json
// call —— 两份文件
{ "file": "new.docx", "other": "old.docx" }
// result
{ "ok": true, "data": {
  "changes": [
    { "kind": "modified", "path": "/body/p[1]", "before": "Heading1", "after": "Heading2", "detail": "style" },
    { "kind": "modified", "path": "/body/p[2]", "before": "First body line.", "after": "Edited body line.", "detail": "text" },
    { "kind": "added", "path": "/body/p[4]", "detail": "paragraph" }
  ],
  "summary": { "added": 1, "removed": 0, "modified": 2, "moved": 0 },
  "baseline": "old.docx", "view": "detailed" } }

// call —— 对比自身快照
{ "file": "report.docx", "snapshot": 1 }
// result（baseline: "snapshot 1"，仅那次编辑的变更）
```

**映射 CLI**：`aioffice diff <file> [<other>] [--snapshot N] [--view summary|detailed]`

**错误码**：`invalid_args`（基线缺失/二者都给/跨格式/`view` 非法/快照号不存在——附可用编号 candidates）、`file_not_found`、`sandbox_denied`、`format_corrupt`、`unsupported_feature`（格式 handler 未实现 diff——三格式均已实现）。

---

### 1.7c `office_convert`（M9）

**用途**：把文档从一种格式**互转**到另一种。按 (源扩展名, 目标扩展名) 对路由：同族文本桥 `docx↔md`（markdown 桥）/ `xlsx↔csv`（csv 桥）复用既有代码；跨格式对经一份**内容中立模型**（`INeutralConvertible`：源 handler `ExportNeutral` → 目标 handler `ImportNeutral`）覆盖 `docx↔pptx`/`docx↔xlsx`/`pptx↔xlsx`/`任意↔md`（md 经命令层 `NeutralMarkdown` 序列化器，故任意 office 格式可转 md）；`any → pdf/png/svg/html/txt` 经 render 层。`src` 只读；`dest` **新建**（覆盖既有 → `convert_overwrite` 警告）。

```json
{
  "type": "object",
  "properties": {
    "src":  { "type": "string", "description": "Source document (.docx/.xlsx/.pptx/.md/.csv); opened read-only" },
    "dest": { "type": "string",
      "description": "Destination, created fresh: .docx/.xlsx/.pptx/.md/.csv (content) or .pdf/.png/.svg/.html/.txt (render). Same ext as src -> invalid_args (use office_edit); unknown ext -> unsupported_feature" }
  },
  "required": ["src", "dest"]
}
```

**result**：`data: { from, to, blocksWritten, dropped: [<feature notes>], written }`。转换是内容转移、跨格式本质有损：每条丢失（导出侧的动画/图表/SmartArt + 导入侧目标无法承载的样式/公式取值/markdown 颜色）汇入 `data.dropped`，并由 verb 层折成单条 `meta.warnings` 的 `convert_lossy`。

**示例**

```json
// call —— docx → pptx（每个标题一页，要点作为正文）
{ "src": "report.docx", "dest": "deck.pptx" }
// result
{ "ok": true, "data": { "from": "docx", "to": "pptx", "blocksWritten": 10, "dropped": [], "written": ".../deck.pptx" } }

// call —— 含图表的 deck → docx（图表无法跨格式，诚实报告）
{ "src": "charted.pptx", "dest": "charted.docx" }
// result（meta.warnings 含 convert_lossy）
{ "ok": true, "data": { "from": "pptx", "to": "docx", "blocksWritten": 1,
  "dropped": ["charts (transferred as data is not supported; the chart is not converted)"], "written": ".../charted.docx" } }
```

**映射 CLI**：`aioffice convert <src> <dest>`

**错误码**：`invalid_args`（缺 src/dest、源与目标同扩展名 → 「use edit」）、`unsupported_feature`（未知目标扩展名——列出受支持目标；csv 仅与 xlsx 直转）、`file_not_found`、`sandbox_denied`、`format_corrupt`。

---

### 1.8 `office_template`

**用途**：模板合并——把 `data` 里的键值对填入文档文本 run 中的 `{{key}}` 占位符（docx/xlsx/pptx 通用，跨 run 拆分的占位符也能命中）。

```json
{
  "type": "object",
  "properties": {
    "file": { "type": "string", "description": "Template document containing {{key}} placeholders in text runs" },
    "data": { "type": "object", "additionalProperties": { "type": "string" },
      "description": "Merge map, e.g. {\"client\":\"ACME Corp\",\"date\":\"2026-06-12\"} fills {{client}} and {{date}}" },
    "output": { "type": "string", "description": "Result path (recommended). Omit = merge in place; pre-image auto-snapshotted" },
    "overwrite": { "type": "boolean", "default": false }
  },
  "required": ["file", "data"]
}
```

**result**：`data: { created: string, replaced: { "<key>": count }, unmatchedKeys: string[] /* keys with 0 hits */, leftoverPlaceholders: string[] /* {{x}} still in doc */ }`

**示例**

```json
// call
{ "file": "templates/contract.docx", "output": "out/acme-contract.docx",
  "data": { "client": "ACME Corp", "amount": "$150,000" } }
// result
{ "ok": true, "data": { "created": "/ws/out/acme-contract.docx",
  "replaced": { "client": 6, "amount": 2 }, "unmatchedKeys": [], "leftoverPlaceholders": ["{{signDate}}"] } }
```

**映射 CLI**：`aioffice template <file> --data <json|@file> [-o out]`

**错误码**：`invalid_args`（data 非法 / 目标已存在未 `overwrite`）、`file_not_found`、`sandbox_denied`、`format_corrupt`。
（`leftoverPlaceholders` 非空不是错误——填一半是合法用法；agent 应检查该字段。）

---

### 1.9 `file_snapshot`

**用途**：快照与撤销。每次成功的 `office_edit` / 原地 `office_template` 前自动建快照（环形保留 20 份，存于 `~/.aioffice/snapshots/<path-hash>/`，与 workspace 解耦）；本工具列出与回滚。

```json
{
  "type": "object",
  "properties": {
    "file": { "type": "string" },
    "action": { "type": "string", "enum": ["list", "restore"], "default": "list" },
    "n": { "type": "integer", "description": "restore: snapshot number from list (omit = latest)" }
  },
  "required": ["file"]
}
```

**result**：`list` → `data: { snapshots: [{ n, at, rev, bytes, trigger }] }`；`restore` → `data: { restored: n, rev /* new current rev */ }`（restore 自身也先快照当前态——undo 可再 undo）。

**示例**

```json
// call
{ "file": "report.docx", "action": "restore", "n": 7 }
// result
{ "ok": true, "data": { "restored": 7, "rev": "9be01dd4c2a7" } }
```

**映射 CLI**：`aioffice snapshot <list|restore> <file> [n]`

**错误码**：`invalid_args`（编号不存在——suggestion 提示先 `list`，`candidates[]` 回填现有编号）、`file_not_found`、`sandbox_denied`。

---

### 1.10 `office_status`

**用途**：健康检查（doctor）。.NET 运行时、平台、沙箱根、快照存储一次性体检。会话开始或连续报错时调用。

```json
{
  "type": "object",
  "properties": {}
}
```

**result**：`data: { healthy: boolean, version: string, runtime: { dotnet, os, arch }, workspace: string, snapshotStore: { path, count, bytes }, checks: [{ name, ok, detail }] }`

**示例**

```json
// call
{}
// result
{ "ok": true, "data": { "healthy": true, "version": "0.2.0",
  "runtime": { "dotnet": "10.0.300", "os": "macos", "arch": "arm64" },
  "workspace": "/ws",
  "snapshotStore": { "path": "~/.aioffice/snapshots", "count": 14, "bytes": 41943040 },
  "checks": [
    { "name": "workspace_writable", "ok": true, "detail": "/ws" },
    { "name": "snapshot_store_writable", "ok": true, "detail": "~/.aioffice/snapshots" }
  ] } }
```

**映射 CLI**：`aioffice doctor`

**错误码**：`internal_error`（仅自身崩溃；环境有问题时尽量返回 `ok:true, data.healthy:false` + checks，让 agent 能自诊断）。

---

### 1.11 `office_help`

**用途**：渐进式文档。属性名 / 值格式 / 元素类型 / selector 语法**不进 tool schema**，全部藏在这里——不确定就查，别猜。

```json
{
  "type": "object",
  "properties": {
    "topic": { "type": "string",
      "description": "Omit → topic index. Examples: \"addressing\", \"selectors\", \"errors\", \"envelope\", \"docx/paragraph\", \"xlsx/cell\", \"pptx/shape\", \"docx/paragraph#set\" (props usable with one verb)" }
  }
}
```

**result**：`data: { topic: string, doc: string /* markdown */, related: string[] /* neighbouring topics */ }`

**示例**

```json
// call — set 能改 paragraph 的哪些属性？
{ "topic": "docx/paragraph#set" }
// result
{ "ok": true, "data": { "topic": "docx/paragraph#set",
  "doc": "## paragraph — settable props\n- text\n- style (Heading1, Normal, ...)\n- bold/italic/underline: \"true\"|\"false\"\n- size: unit-qualified, e.g. \"12pt\"\n- color: hex \"FF0000\" or named\n- align: left|center|right|justify",
  "related": ["docx/paragraph#add", "docx/run", "addressing"] } }
```

**映射 CLI**：`aioffice help [topic]`

**错误码**：`invalid_args`（未知主题，**附近似主题 `candidates[]`**）。

---

### 1.12 `office_schema`

**用途**：整个命令面的机器可读 JSON——agent 自省全部动词、参数、错误码、示例，而不是猜。17 工具 / 17 动词的单一事实来源（与 `aioffice schema` 字节一致）。

```json
{
  "type": "object",
  "properties": {
    "verb": { "type": "string",
      "description": "Omit → full surface. One of: create|read|query|get|edit|render|validate|template|snapshot|doctor|schema|help|preview|mcp" }
  }
}
```

**result**：`data: { version: string, verbs: [{ name, summary, args: object, errors: string[], examples: [{ call, note }] }] }`（传 `verb` 时 `verbs` 仅含该项）。

**示例**

```json
// call
{ "verb": "edit" }
// result
{ "ok": true, "data": { "version": "0.2.0", "verbs": [
  { "name": "edit", "summary": "Atomic batch mutation",
    "args": { "ops": "...", "expect_rev": "...", "dry_run": "..." },
    "errors": ["stale_address", "invalid_path", "invalid_args", "unsupported_feature", "file_not_found", "sandbox_denied", "format_corrupt"],
    "examples": [{ "call": "aioffice edit r.docx --set /body/p[1] style=Heading1", "note": "single-op sugar" }] }
] } }
```

**映射 CLI**：`aioffice schema [verb]`

**错误码**：`invalid_args`（未知 verb，附 `candidates[]`）。

---

### 1.13 `preview_open`（M1 转正）

**用途**：打开本地实时预览——人在浏览器里**点选元素**（点击高亮并记录其 `data-aio-path` 规范路径），AI 用 `preview_selection` 读回，「人指哪、AI 打哪」。文件被编辑后浏览器经 SSE 自动刷新。

实现要点：阻塞型 server 不能住在 MCP 进程里（stdio 属于 JSON-RPC），所以本工具**拉起 detached `aioffice preview open` 子进程**（自身可执行文件定位：`AIOFFICE_EXE` env → 当前进程 → 同目录 `aioffice`），轮询 lockfile + HTTP health 最多 10s，成功返回 url。幂等：同文件已有活预览时直接返回它（`meta.warnings: already_running`）。

```json
{
  "type": "object",
  "properties": {
    "file": { "type": "string" },
    "port": { "type": "integer", "description": "Fixed port (default: auto-pick in 26500-26600)" }
  },
  "required": ["file"]
}
```

**result**：`data: { url: string, port: number, pid: number }`

**示例**

```json
// call
{ "file": "report.docx" }
// result
{ "ok": true, "data": { "url": "http://127.0.0.1:26500/", "port": 26500, "pid": 4242 } }
```

**映射 CLI**：`aioffice preview open <file> [--port N]`（CLI 形态是前台阻塞 server，envelope 先打印再阻塞）

**错误码**：`file_not_found`、`sandbox_denied`、`invalid_args`（坏 port）、`unsupported_feature`（定位不到 aioffice 可执行文件，suggestion 给 `AIOFFICE_EXE` / 手动启动）、`internal_error`（子进程 10s 未就绪或启动即死，relay 子进程 envelope）。

---

### 1.14 `preview_selection`（M1 转正）

**用途**：读回用户在预览浏览器里点选的元素的规范路径，直接喂给 `office_get` / `office_edit`。

```json
{
  "type": "object",
  "properties": { "file": { "type": "string" } },
  "required": ["file"]
}
```

**result**：`data: { paths: string[], rev: string | null, updatedAt: string | null }`（`rev` = 选择发生时的文件 rev；编辑前可与当前 `meta.rev` 对比判断选区是否已过期）

**示例**

```json
// call
{ "file": "report.docx" }
// result — 用户点了一个段落和一个表格
{ "ok": true, "data": { "paths": ["/body/p[3]", "/body/table[1]"], "rev": "a3f9c12be01d", "updatedAt": "2026-06-12T08:30:00Z" } }
```

**映射 CLI**：`aioffice preview selection <file>`（`preview close <file>` 关停 server 并清理 lockfile）

**错误码**：`preview_not_running`（该文件没有运行中的预览，suggestion 指向 `preview_open`）、`sandbox_denied`。

---

## 2. 设计原则

### 2.1 Schema token 预算（总额 ≤ 3500 tokens）

工具列表进入每个 agent 的上下文，schema 即税。预算按「name + description + inputSchema 序列化后」估算，CI 中用 tokenizer 实测并 fail 超额（`tests/AIOffice.Mcp.Tests` 的 schema-budget 测试）：

| 工具 | 预算 (tokens) | 理由 |
|---|---:|---|
| `office_edit` | 620 | 唯一 fat tool，承载全部变更动词与 ops 语法 |
| `office_render` | 320 | 5 种目标格式 + scope + png/pdf 行为说明 |
| `office_read` | 240 | 4 种 view + range |
| `office_query` | 230 | selector 示例占大头 |
| `office_template` | 230 | merge 语义 + 占位符示例 |
| `office_get` | 180 | 路径寻址示例 |
| `office_help` | 170 | topic 示例清单 |
| `file_snapshot` | 150 | |
| `office_create` | 150 | |
| `office_schema` | 120 | |
| `office_validate` | 100 | |
| `office_audit` | 140 | M7 加入；描述刻意精简，code 全集外置到 `office_help {topic:"audit"}` |
| `office_diff` | 160 | M8 加入；描述刻意精简，diff 语义/变更种类外置到 `office_help {topic:"diff"}` |
| `office_convert` | 164 | M9 加入；描述刻意精简，(src→dest) 矩阵 + 有损说明外置到 `office_help {topic:"convert"}` |
| `office_status` | 90 | |
| `preview_open` | 300 | M1 转正，吃掉当年预留额度的大头 |
| `preview_selection` | 180 | M1 转正 |
| **M9 小计（17 工具）** | **3484** | M8 小计 3320（M7 小计 3160 + `office_diff` 160）+ `office_convert` 164 |
| 措辞浮动预留 | 16 | description 迭代余量（历代新 type 词表/视图/提示均从这里支出，实测全表仍 ≤ 3500，CI schema-budget 测试把关） |
| **总额** | **≤ 3500** | CI 的 schema-budget 测试实测把关；加入 `office_convert`（第 17 个工具）后仍未超额，故沿用 3500 上限、不上调天花板 |

预算纪律：示例写进字段 description（一行内），不写长篇；枚举值自解释的不加 description；**属性名表 / selector 全语法 / 寻址细则一律外置到 `office_help`**——这是预算能压住的根本原因。

### 2.2 Few-fat vs many-thin：为什么是 17 个中粒度能力

- **Many-thin 的失败模式**：按「动词 × 格式」切会得到 50+ 个工具（docx_add_paragraph、xlsx_set_cell…），schema 总量超预算一个数量级，且把路由难题推给模型——工具选择错误率随工具数超线性上涨。
- **One-mega 的失败模式**：单个 `run_aioffice(argv: string)` 看似零预算，实际把 CLI 语法学习成本转嫁给模型：丢失参数级校验、丢失结构化错误与 `candidates[]`、丢失 `expect_rev` / 自动快照等横切机制的注入点。
- **我们的切法：按意图分层**——读三档粒度（`read` 全文档 → `query` 检索 → `get` 单点）、写**一档**（唯一的 fat `office_edit`：所有变更共享快照 / rev 守卫 / 原子保存这套横切机制，合并后机制只实现一次）、看（`render`）、体检（`validate`）、模板（`template`）、安全网（`snapshot` / `status`）、自描述（`help` / `schema`）、人机协作（`preview_open` / `preview_selection`）。
- CLI 动词与 MCP 工具 **1:1 镜像**（见附表）：agent 在两个入口间切换零学习成本，文档、测试、schema 只维护一份。

### 2.3 render → look → fix 循环

视觉类产物（尤其 pptx）的核心工作模式——`office_validate` 通过 ≠ 长得对：

```
office_edit(改)
  → office_render(to=svg, scope=/slide[3])      // pptx 单页 svg；docx/xlsx 用 to=html
  → [检视 data.content 里的标记与几何]            // 文本溢出 = text 宽度 > shape 宽度；
                                                 // 撞色 = 相邻元素 fill 相同；位置 = x/y/w/h 坐标
  → office_query 定位 → office_get 确认 → office_edit(修)
  → office_render 只重渲染受影响的 scope
  → 收敛后 office_validate 收尾
```

规则：**每完成一个视觉里程碑必须 render 看一次**，不要盲编 10 步再看；只渲染受影响的 `scope`，省 token 也省时间；非视觉问题（schema 违规、空段落、悬空引用）用 `office_validate` 的 issues，比渲染便宜。读 html/svg 标记可精确判断几何与文本；要像素级观感用 `to=png`——result 附带 MCP image content block，模型直接看图。

### 2.4 渐进式文档漏斗（`office_help` / `office_schema`）

OOXML 属性面有几千个属性名，全塞 schema 是预算自杀。三层漏斗：

1. **第 0 层（常驻）**：17 个 tool schema 里只有动词、路径语法和一行示例（≈3484 tokens）；
2. **第 1 层（按需）**：动手前 `office_help {topic:"<fmt>/<element>#<verb>"}` 拿该元素该动词的准确属性表——**一次 help 查询胜过 guess-fail-retry 三轮**；`office_schema` 给整个命令面的机器可读自省；
3. **第 2 层（自愈）**：错误路径也接进漏斗——`invalid_path` 自动回填 `candidates[]`；`invalid_args`（坏 selector / 未知主题）的 suggestion 直接指向对应的 help 主题；`unsupported_feature` 的 suggestion 必须给出当前可用的 workaround。

---

## 3. Agent 系统提示词片段（建议直接内嵌到调用方 agent 的 system prompt）

```text
You have aioffice MCP tools for real .docx/.xlsx/.pptx files. Rules:

1. Every result is one JSON envelope {ok, data, error, meta}. On ok=false, READ
   error.suggestion and error.candidates before retrying — never repeat the same
   call unchanged. Also check meta.warnings (e.g. formula_not_evaluated on xlsx).
2. Never guess property names, element types, or selector syntax. Call
   office_help first ({topic:"docx/paragraph#set"}, {topic:"selectors"},
   {topic:"addressing"}). office_schema returns the whole machine-readable surface.
3. Addressing: discover canonical paths with office_query (CSS-like selectors),
   inspect with office_get. Paths are 1-based and positional — after inserts or
   deletes, re-query instead of reusing old indices.
4. ALL mutations go through office_edit. Group related changes into one atomic
   ops[] batch (single save, all-or-nothing). Pass expect_rev (the last meta.rev)
   when editing the same file across turns; on stale_address, re-read, then
   re-apply. Use dry_run:true to pre-flight risky batches.
5. Work loop: edit → office_render (to=html for docx/xlsx, to=svg per slide for
   pptx, scope=affected subtree) → inspect the returned markup/geometry → fix →
   re-render only that scope. PNG rendering is M1; render to=png today returns
   unsupported_feature with a workaround. Finish with office_validate and treat
   schemaErrors as blockers.
6. Every successful edit auto-snapshots the pre-image — no extra safety step
   needed. To undo: file_snapshot{action:"list"} then {action:"restore", n}.
7. If tools start failing oddly, call office_status once and report data.checks
   instead of retrying blindly.
```

> 该片段约 300 tokens，英文以保证跨模型稳健；可按宿主 agent 语言习惯翻译，但工具名、字段名、错误码保持英文原样。

---

## 附：MCP 工具 ↔ CLI 动词映射总表（1:1，同一内部命令层）

| MCP 工具 | aioffice CLI | 备注 |
|---|---|---|
| `office_create` | `create <file> [--kind] [--title]` | |
| `office_read` | `read <file> [--view] [--range] [--max-bytes]` | view: outline/text/stats/structure |
| `office_query` | `query <file> <selector>` | 返回规范路径 |
| `office_get` | `get <file> <path>` | 单节点 + 属性 |
| `office_edit` | `edit <file> --ops <json|@file> [--dry-run] [--expect-rev]` | 糖：`--set/--add/--remove` |
| `office_render` | `render <file> [--to] [--scope] [-o]` | to: html/svg/text/png/pdf |
| `office_validate` | `validate <file>` | OpenXmlValidator + lint |
| `office_audit` | `audit <file> [--category] [--severity] [--fix]` | M7：无障碍 + 质量 lint，findings 是数据 |
| `office_diff` | `diff <file> [<other>] [--snapshot N] [--view]` | M8：语义对比，变更是数据 |
| `office_template` | `template <file> --data <json|@file> [-o]` | `{{key}}` 合并 |
| `office_convert` | `convert <src> <dest>` | M9：跨格式互转，丢失项汇入 convert_lossy |
| `file_snapshot` | `snapshot <list|restore> <file> [n]` | 环形 20 份 |
| `office_status` | `doctor` | |
| `office_help` | `help [topic]` | 渐进式文档 |
| `office_schema` | `schema [verb]` | 命令面自省 |
| `preview_open` | `preview open <file> [--port N]` | MCP 侧 detached 子进程；CLI 侧前台阻塞 |
| `preview_selection` | `preview selection <file>` | `preview close <file>` 关停 |

> CLI 全局旗标：`--json`（非 TTY 默认）| `--pretty` | `--workspace <dir>`（或 `AIOFFICE_WORKSPACE`）| `--quiet`。
> CLI exit codes：`0` ok | `2` user/input error | `3` internal/format error | `4` sandbox_denied | `5` unsupported_feature（与 §0.4 错误码表对应）。
