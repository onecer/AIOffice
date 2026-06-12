# AIOffice

**100% 自研的 AI-Native Office CLI / MCP** — 让 AI agent 像调用函数一样创建、查询、
编辑、渲染真实的 Office 文件（.docx / .xlsx / .pptx）。纯 C#/.NET 实现，直接
读写 OOXML，不包装任何第三方引擎二进制。

A fully self-built, AI-native command line + MCP server for real Office files.
Pure C#/.NET on top of the OOXML SDK — no wrapped third-party engine binaries.

> **设计声明 / Design statement** — AIOffice 的命令面与 OfficeCLI **不兼容**，
> 这是有意为之：能力上对齐业界水准，但动词、参数、寻址语法全部自研，为 agent
> 设计一套更小、更可预测的心智模型。CLI 动词与 MCP 工具 1:1 对应，一套心智模型
> 两种接入方式。/ The command surface is deliberately **not** compatible with
> any existing office CLI: capability parity is the north star, but the verbs,
> flags and addressing grammar are our own, designed for agents.

## 快速开始 / Quickstart

```bash
# 构建（.NET 10 SDK）/ build
dotnet build AIOffice.sln

# 运行 / run（或发布单文件二进制，见下）
alias aioffice='dotnet run --project src/AIOffice.Cli --'

# 体检 / diagnose
aioffice doctor
```

每条命令向 stdout 输出**恰好一个 JSON 信封** / every command prints exactly
one JSON envelope:

```json
{"ok":true,"data":{"name":"aioffice","version":"0.1.0","runtime":".NET 10.0.8"},
 "meta":{"elapsedMs":2,"version":"0.1.0"}}
```

### 创建 → 编辑 → 读取 / create → edit → read

```bash
aioffice create report.docx --title "Q3 报告"
# {"ok":true,"data":{...},"meta":{"file":"report.docx","rev":"817e02e9fe9f","elapsedMs":13,"version":"0.1.0"}}

# 原子批量编辑（自动快照前像，最多保留 20 份）
aioffice edit report.docx --ops '[
  {"op":"add","path":"/body","type":"p","position":"inside",
   "props":{"text":"季度报告","style":"Heading1"}}]'

# 语法糖：--set / --add / --remove
aioffice edit report.docx --set '/body/p[1]' text='营收增长 12%'

# 乐观并发：rev 不匹配则在任何写入前失败
aioffice edit report.docx --expect-rev 817e02e9fe9f --set '/body/p[1]' text=Hi
# {"ok":false,"error":{"code":"stale_address",
#   "message":"File rev is 3baf0725580f but --expect-rev was 817e02e9fe9f; ...",
#   "suggestion":"Re-run 'aioffice read' or 'aioffice query' to refresh paths, then retry with the new rev."},
#  "meta":{"elapsedMs":23,"version":"0.1.0"}}

aioffice read report.docx --view text
aioffice query report.docx "p[style=Heading1]"   # 返回稳定的规范路径
aioffice snapshot list report.docx               # 自动快照环
aioffice snapshot restore report.docx 1          # 回滚（回滚本身也可撤销）
```

错误信封永远带可执行的 `suggestion`；路径解析失败（`invalid_path`）还会附上
最近匹配的 `candidates`。Agent 不用猜 — `aioffice schema` 返回整个命令面的
机器可读 JSON。

### MCP（给 Claude 等 agent）

```bash
aioffice mcp        # stdio MCP server，12 个工具与 CLI 动词 1:1（preview_* 预留 M1）
```

Claude Desktop / Claude Code 配置 / config snippet:

```json
{
  "mcpServers": {
    "aioffice": {
      "command": "aioffice",
      "args": ["mcp", "--workspace", "/path/to/your/documents"]
    }
  }
}
```

工具对应关系：`office_create`=create · `office_read`=read · `office_query`=query ·
`office_get`=get · `office_edit`=edit · `office_render`=render ·
`office_validate`=validate · `office_template`=template · `file_snapshot`=snapshot ·
`office_status`=doctor · `office_help`=help · `office_schema`=schema ·
`preview_*` 预留 M1。

### 单文件发布 / Single-file publish

```bash
dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release \
  -p:PublishSingleFile=true --self-contained
```

## 命令面 / Command surface (v0)

| 命令 / verb | 说明 / summary |
|---|---|
| `create <file> [--kind] [--title]` | 新建文档，类型按扩展名推断 |
| `read <file> [--view outline\|text\|stats\|structure]` | 读取投影 |
| `query <file> <selector>` | CSS 风格选择器 → 规范路径，如 `p[style=Heading1]`、`cell[value>100]` |
| `get <file> <path>` | 取单个节点及属性 |
| `edit <file> --ops <json\|@file>` | **原子**批量 set/add/remove/move；`--dry-run`、`--expect-rev` |
| `render <file> [--to html\|svg\|text]` | 渲染（pptx 可 `--scope /slide[2]`）；png 属 M1 |
| `validate <file>` | OOXML 校验 + lint，带修复建议 |
| `template <file> --data <json>` | `{{key}}` 模板合并 |
| `snapshot <list\|restore> <file> [n]` | 编辑前自动快照环（20 份） |
| `doctor` | 环境/运行时/处理器诊断 |
| `schema [verb]` | 整个命令面的机器可读 JSON |
| `help [topic]` | addressing / selectors / properties-* / errors |
| `mcp` | stdio MCP server |

全局：`--json`（非 TTY 默认）· `--pretty` · `--workspace <dir>`（沙箱根，默认
cwd，亦可 `AIOFFICE_WORKSPACE`）· `--quiet`。退出码：0 ok · 2 用户错误 ·
3 内部/格式错误 · 4 sandbox_denied · 5 unsupported_feature。

寻址一律 1 起始 / addressing is 1-based:
`/body/p[3]` · `/body/table[1]/tr[2]/tc[1]` · `/Sheet1/A1:C10` · `/'Q3 Data'/B2` ·
`/slide[2]/shape[3]`。

## 架构 / Architecture

```
                 ┌─────────────────────────────────────────────┐
   agent/human → │  src/AIOffice.Cli   (aioffice, 14 verbs)    │
   MCP client  → │  src/AIOffice.Mcp   (stdio, 13 tools, 1:1)  │
                 └──────────────┬──────────────────────────────┘
                                │ envelope · addressing · selectors
                 ┌──────────────▼──────────────────────────────┐
                 │  src/AIOffice.Core                          │
                 │  envelope/errors · DocPath · Selector       │
                 │  Workspace sandbox · SnapshotStore · rev    │
                 └───────┬──────────────┬──────────────┬───────┘
                         │              │              │
                 ┌───────▼─────┐ ┌──────▼──────┐ ┌─────▼───────┐
                 │AIOffice.Word│ │AIOffice.Excel│ │AIOffice.Pptx│
                 │ OpenXml SDK │ │  ClosedXML  │ │ OpenXml SDK │
                 └───────┬─────┘ └──────┬──────┘ └─────┬───────┘
                         ▼              ▼              ▼
                       .docx          .xlsx          .pptx     (lossless OOXML)
```

质量门 / quality bar：每个格式包同 PR 配 xunit 测试；**往返定律**（开→存
不编辑，所有 zip part 字节不变）；每次变更后 OpenXmlValidator 0 错误；
`fixtures/manual-check/` 供真人用真实 Office 打开核验。

## 文档 / Docs

- `docs/DESIGN.md` — 设计与权衡
- `docs/MCP.md` — MCP 工具面
- `docs/PARITY.md` — 能力对齐路线（north star）
- `aioffice help addressing|selectors|properties-docx|properties-xlsx|properties-pptx|errors`
- `fixtures/README.md` — 黄金脚本与夹具

## 路线图 / Roadmap

- **M0（当前 / now）** — 14 个动词全量接线；docx/xlsx/pptx 创建·读取·查询·编辑·
  校验·模板；html/svg/text 渲染；快照环；MCP stdio；CI（macOS + Windows）+
  单文件发布。
- **M1** — png 渲染与 `preview_*` 工具；pptx 母版/版式寻址；docx 页眉页脚创建；
  xlsx 图表；更多形状类型。
- **M2** — 修订（tracked changes）、批注、样式管理；xlsx 数据透视；性能
  （大文件流式读）。
- **M3** — 跨文档工作流（xlsx 数据 → pptx 图表）、批处理、插件式能力扩展。

## License

Apache-2.0（见 `LICENSE` 与 `THIRD-PARTY-NOTICES.md`）。
