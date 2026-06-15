# AIOffice

[English](README.md) | **简体中文**

[![CI](https://github.com/onecer/AIOffice/actions/workflows/ci.yml/badge.svg)](https://github.com/onecer/AIOffice/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
![.NET](https://img.shields.io/badge/.NET-10-512BD4)
![platforms](https://img.shields.io/badge/platforms-macOS%20%7C%20Windows%20%7C%20Linux-lightgrey)

**为 AI agent 而生的 Office CLI + MCP 服务器。** AIOffice 让 agent 像调用函数一样创建、查询、编辑、渲染、预览、校验真实的 `.docx` / `.xlsx` / `.pptx`：一条命令进，恰好一个 JSON 信封出。

100% 自研，纯 C#/.NET 直接无损读写 OOXML（DocumentFormat.OpenXml + ClosedXML）。**一个约 36 MB 的单文件二进制，无需安装 Office，零运行时依赖，不包装任何第三方引擎。**

```bash
aioffice create report.docx --title "Q3 报告"
aioffice edit   report.docx --set '/body/p[1]' text="营收增长 12%"
aioffice read   report.docx --view outline
aioffice diff   report.docx old.docx        # 语义对比：一份已排序的变更清单
aioffice convert report.docx deck.pptx      # 跨格式：每个标题一页幻灯片，要点作为正文
aioffice mcp    # 同样的 17 项能力，以 MCP 工具形式走 stdio
```

## 眼见为实

下面每个文件都仅由 `aioffice` 创建、编辑、校验并截图——没装 Office，没有模板，没有手工修饰。完整脚本折叠在下方；[deck-1.svg](assets/demo/deck-1.svg) 是同一页标题幻灯片的 SVG 版本，其中每个形状都带有指回规范文档路径的 `data-aio-path` 属性。

| | |
|---|---|
| ![deck.pptx 第 1 页，由 aioffice 渲染为 PNG](assets/demo/deck-1.png)<br><sub>`aioffice render deck.pptx --to png --scope '/slide[1]' -o deck-1.png`</sub> | ![deck.pptx 第 2 页，由定位形状拼出的数据卡片](assets/demo/deck-2.png)<br><sub>`aioffice query deck.pptx 'shape:contains("14")'` → `/slide[2]/shape[@id=9]`</sub> |
| ![report.docx：页眉、标题与表格](assets/demo/report.png)<br><sub>`aioffice get report.docx '/header[1]/p[1]'` → `"text": "AIOffice Demo"`</sub> | ![metrics.xlsx：公式求值与数字格式](assets/demo/metrics.png)<br><sub>`aioffice get metrics.xlsx /Sheet1/B7` → `"cachedValue": 286900`</sub> |

<details>
<summary>完整脚本（逐条原样命令，含 render → look → fix 闭环）</summary>

```bash
# 在一个空目录中，aioffice 已加入 PATH

# ---- deck.pptx —— 3 页幻灯片，深色背景 + 强调色形状 ----
aioffice create deck.pptx
aioffice edit deck.pptx --ops '[
  {"op":"add","path":"/slide[1]","type":"shape","props":{"name":"bg","x":0,"y":0,"w":"33.87cm","h":"19.05cm","fill":"0F172A"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"name":"deco-1","x":26.2,"y":10.8,"w":12,"h":12,"fill":"1E293B"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"name":"deco-2","x":30.4,"y":15,"w":8,"h":8,"fill":"38BDF8"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"name":"accent","x":2.6,"y":6.1,"w":5.2,"h":0.16,"fill":"38BDF8"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"text":"AIOffice","x":2.5,"y":6.6,"w":24,"h":3.6,"fontSize":60,"bold":true,"color":"FFFFFF"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"text":"An AI-native CLI + MCP server for real Office files","x":2.5,"y":10.4,"w":26,"h":1.8,"fontSize":20,"color":"94A3B8"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"text":"This deck was built entirely by aioffice edit — no Office installed","x":2.5,"y":16.9,"w":26,"h":1.2,"fontSize":12,"color":"64748B"}}]'
aioffice edit deck.pptx --ops '[
  {"op":"add","path":"/slide[1]","type":"slide","position":"after"},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"name":"bg","x":0,"y":0,"w":"33.87cm","h":"19.05cm","fill":"0F172A"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"name":"accent","x":2.6,"y":2.0,"w":3.6,"h":0.16,"fill":"38BDF8"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"M1 in numbers","x":2.5,"y":2.5,"w":20,"h":2.2,"fontSize":34,"bold":true,"color":"FFFFFF"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"name":"card-1","x":2.5,"y":6.4,"w":8.6,"h":8.2,"fill":"1E293B"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"3","x":3.4,"y":7.4,"w":6.8,"h":2.6,"fontSize":48,"bold":true,"color":"38BDF8"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"file formats — docx, xlsx, pptx — one 36 MB binary","x":3.4,"y":10.6,"w":6.8,"h":3.4,"fontSize":13,"color":"94A3B8"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"name":"card-2","x":12.6,"y":6.4,"w":8.6,"h":8.2,"fill":"1E293B"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"14","x":13.5,"y":7.4,"w":6.8,"h":2.6,"fontSize":48,"bold":true,"color":"38BDF8"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"MCP tools, 1:1 with the CLI verbs","x":13.5,"y":10.6,"w":6.8,"h":3.4,"fontSize":13,"color":"94A3B8"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"name":"card-3","x":22.7,"y":6.4,"w":8.6,"h":8.2,"fill":"1E293B"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"0","x":23.6,"y":7.4,"w":6.8,"h":2.6,"fontSize":48,"bold":true,"color":"38BDF8"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"Office installs required — render, preview, validate built in","x":23.6,"y":10.6,"w":6.8,"h":3.4,"fontSize":13,"color":"94A3B8"}}]'
aioffice edit deck.pptx --ops '[
  {"op":"add","path":"/slide[2]","type":"slide","position":"after"},
  {"op":"add","path":"/slide[3]","type":"shape","props":{"name":"bg","x":0,"y":0,"w":"33.87cm","h":"19.05cm","fill":"0F172A"}},
  {"op":"add","path":"/slide[3]","type":"shape","props":{"name":"deco","x":-3,"y":13.5,"w":10,"h":10,"fill":"1E293B"}},
  {"op":"add","path":"/slide[3]","type":"shape","props":{"name":"accent","x":2.6,"y":7.0,"w":5.2,"h":0.16,"fill":"38BDF8"}},
  {"op":"add","path":"/slide[3]","type":"shape","props":{"text":"render → look → fix","x":2.5,"y":7.5,"w":28,"h":3.2,"fontSize":44,"bold":true,"color":"FFFFFF"}},
  {"op":"add","path":"/slide[3]","type":"shape","props":{"text":"One JSON envelope at a time. github.com/onecer/AIOffice","x":2.5,"y":11,"w":26,"h":1.6,"fontSize":18,"color":"94A3B8"}}]'
aioffice validate deck.pptx

# ---- report.docx —— 页眉、Heading1/2、导语、表格 ----
aioffice create report.docx
aioffice edit report.docx --ops '[
  {"op":"add","path":"/header[1]","type":"header","props":{"text":"AIOffice Demo"}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"AIOffice M1 Report","style":"Heading1"}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"This document was created entirely from the command line by aioffice — headings, header, table and styles included. Every block below is addressable (/body/p[2], /body/table[1]) and was written through the same atomic edit batches an AI agent would use."}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"Milestone snapshot","style":"Heading2"}},
  {"op":"add","path":"/body","type":"table","position":"inside","props":{"rows":4,"cols":3}},
  {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"Milestone"}},
  {"op":"set","path":"/body/table[1]/tr[1]/tc[2]","props":{"text":"Highlights"}},
  {"op":"set","path":"/body/table[1]/tr[1]/tc[3]","props":{"text":"Status"}},
  {"op":"set","path":"/body/table[1]/tr[2]/tc[1]","props":{"text":"M0"}},
  {"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"text":"create / query / edit / render / validate"}},
  {"op":"set","path":"/body/table[1]/tr[2]/tc[3]","props":{"text":"shipped"}},
  {"op":"set","path":"/body/table[1]/tr[3]/tc[1]","props":{"text":"M1"}},
  {"op":"set","path":"/body/table[1]/tr[3]/tc[2]","props":{"text":"png render, live preview, headers/footers, xlsx charts"}},
  {"op":"set","path":"/body/table[1]/tr[3]/tc[3]","props":{"text":"shipped"}},
  {"op":"set","path":"/body/table[1]/tr[4]/tc[1]","props":{"text":"M2"}},
  {"op":"set","path":"/body/table[1]/tr[4]/tc[2]","props":{"text":"tracked changes, comments, pivot tables"}},
  {"op":"set","path":"/body/table[1]/tr[4]/tc[3]","props":{"text":"next"}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"How it was made","style":"Heading2"}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"aioffice edit report.docx --ops … applied all of the above in one atomic batch; aioffice validate confirms the OOXML is clean, and aioffice render --to png produced the image you are looking at."}}]'
aioffice read report.docx --view outline        # 先看大纲：标题 + 规范路径
aioffice edit report.docx --remove '/body/p[1]' # 删掉 create 留下的空段落
aioffice validate report.docx
aioffice get report.docx '/header[1]/p[1]'      # → "text": "AIOffice Demo"

# ---- metrics.xlsx —— 销售表、SUM/AVERAGE、数字格式、柱状图 ----
aioffice create metrics.xlsx
aioffice edit metrics.xlsx --ops '[
  {"op":"set","path":"/Sheet1/A1","props":{"value":"Month"}},
  {"op":"set","path":"/Sheet1/B1","props":{"value":"Revenue"}},
  {"op":"set","path":"/Sheet1/C1","props":{"value":"Units"}},
  {"op":"set","path":"/Sheet1/A1:C1","props":{"bold":true,"fill":"DBEAFE"}},
  {"op":"set","path":"/Sheet1/A2","props":{"value":"Jan"}},
  {"op":"set","path":"/Sheet1/B2","props":{"value":48800}},
  {"op":"set","path":"/Sheet1/C2","props":{"value":305}},
  {"op":"set","path":"/Sheet1/A3","props":{"value":"Feb"}},
  {"op":"set","path":"/Sheet1/B3","props":{"value":52400}},
  {"op":"set","path":"/Sheet1/C3","props":{"value":327}},
  {"op":"set","path":"/Sheet1/A4","props":{"value":"Mar"}},
  {"op":"set","path":"/Sheet1/B4","props":{"value":61200}},
  {"op":"set","path":"/Sheet1/C4","props":{"value":382}},
  {"op":"set","path":"/Sheet1/A5","props":{"value":"Apr"}},
  {"op":"set","path":"/Sheet1/B5","props":{"value":57600}},
  {"op":"set","path":"/Sheet1/C5","props":{"value":360}},
  {"op":"set","path":"/Sheet1/A6","props":{"value":"May"}},
  {"op":"set","path":"/Sheet1/B6","props":{"value":66900}},
  {"op":"set","path":"/Sheet1/C6","props":{"value":418}},
  {"op":"set","path":"/Sheet1/A7","props":{"value":"Total","bold":true}},
  {"op":"set","path":"/Sheet1/B7","props":{"value":"=SUM(B2:B6)","bold":true}},
  {"op":"set","path":"/Sheet1/C7","props":{"value":"=SUM(C2:C6)","bold":true}},
  {"op":"set","path":"/Sheet1/A8","props":{"value":"Average"}},
  {"op":"set","path":"/Sheet1/B8","props":{"value":"=AVERAGE(B2:B6)"}},
  {"op":"set","path":"/Sheet1/C8","props":{"value":"=AVERAGE(C2:C6)","numberFormat":"0.0"}},
  {"op":"set","path":"/Sheet1/B2:B8","props":{"numberFormat":"$#,##0"}},
  {"op":"add","path":"/Sheet1","type":"chart","props":{"kind":"bar","dataRange":"A1:B6","anchor":"E2","title":"Revenue by month"}}]'
aioffice get metrics.xlsx /Sheet1/B7            # → "formula": "=SUM(B2:B6)", "cachedValue": 286900
aioffice get metrics.xlsx /Sheet1/B8            # → "formula": "=AVERAGE(B2:B6)", "cachedValue": 57380
aioffice get metrics.xlsx '/Sheet1/chart[1]'    # → 柱状图 "Revenue by month"，A1:B6 @ E2
aioffice validate metrics.xlsx

# ---- render：「看」的那一步 ----
aioffice render deck.pptx    --to png --scope '/slide[1]' -o deck-1.png
aioffice render deck.pptx    --to png --scope '/slide[2]' -o deck-2.png
aioffice render deck.pptx    --to svg --scope '/slide[1]' -o deck-1.svg
aioffice render report.docx  --to png -o report.png
aioffice render metrics.xlsx --to png -o metrics.png

# ---- look → fix：第 2 页卡片文字溢出了卡片边界 ----
aioffice query deck.pptx 'shape:contains("formats")'   # 找到标签 → /slide[2]/shape[6]
aioffice edit deck.pptx --ops '[
  {"op":"set","path":"/slide[2]/shape[6]","props":{"text":"file formats — one binary"}},
  {"op":"set","path":"/slide[2]/shape[9]","props":{"text":"MCP tools, 1:1 with the CLI"}},
  {"op":"set","path":"/slide[2]/shape[12]","props":{"text":"Office installs required"}}]'
aioffice validate deck.pptx
aioffice render deck.pptx --to png --scope '/slide[2]' -o deck-2.png
```

PNG 入库 `assets/demo/` 前仅做了等比缩放（≤900 px 宽，sips/Pillow），像素未做其他处理。xlsx 的 PNG 展示的是 HTML 渲染器当前画出的内容：单元格、格式与公式缓存结果；柱状图真实存在于文件中（见 `get '/Sheet1/chart[1]'`），用 Excel 打开即可看到。

</details>

## 为什么是 "AI-Native"？

大多数 office 库为程序员而写，大多数 office CLI 为人类而写。AIOffice 为 **agent** 而写——每个设计决策都在优化 LLM 真实运行的循环：*执行 → 观察 → 恢复 → 验证*。

| 特性 | 对 agent 意味着什么 |
|---|---|
| **每条命令恰好一个 JSON 信封** | stdout 永远是 `{ok, data, error, meta}`，无需解析猜测 |
| **会教学的错误** | 每个错误强制携带 `suggestion`；`invalid_path` 还附带服务端算好的 `candidates`（最近的合法路径）——一次失败，零轮浪费的恢复对话 |
| **稳定寻址** | `/body/p[3]`、`/Sheet1/A1:C10`、`/slide[2]/shape[3]` —— 1 起始、规范化、由 `query` 返回，编辑永远不用猜序号 |
| **原子批量编辑** | `edit --ops '[...]'` 全部成功或全部不落盘；支持 `--dry-run`；乐观并发守卫 `--expect-rev`——文件被外部改动时在**任何写入之前**以 `stale_address` 失败 |
| **自动可撤销** | 每次变更前自动快照前像（20 份环形保留），`snapshot restore` 一次调用即回滚，回滚本身也可撤销 |
| **公式写入即求值** | Excel 公式写入时计算并**把结果缓存进文件**（`=SUM(A1:A2)` → 重新打开直接是 `42`）；引擎算不了的函数显式给 `formula_not_evaluated` 警告——绝不静默留旧值 |
| **render → look → fix 闭环** | docx/xlsx 渲染 HTML、pptx 逐页渲染 SVG，还能经系统浏览器一键出 **PNG**，全程无需 Office——agent 能"看见"自己做的东西再改 |
| **人在环中的实时预览** | `preview open` 在 localhost 起一个实时视图；渲染节点带 `data-aio-path`，人类**点一下**，`preview selection` 就把规范路径还给 agent |
| **默认沙箱** | 所有文件参数在工作区白名单内解析（`--workspace`，含符号链接逃逸检查）；越界即 `sandbox_denied`，退出码 4 |
| **可自省的命令面** | `aioffice schema` 返回整个命令面的机器可读 JSON——agent 读规范，而不是幻觉规范 |
| **CLI = MCP 同一套心智模型** | 18 个 CLI 动词（其中 17 个与 MCP 工具 1:1 对应；第 18 个 `version` 仅 CLI），学一次，shell 和 stdio 两种接入 |

### "会教学的错误" —— 真实输出

```bash
$ aioffice get report.docx '/body/paragraph[1]'
```
```json
{
  "ok": false,
  "error": {
    "code": "invalid_path",
    "message": "'paragraph' cannot appear under /body (body contains: p, table).",
    "suggestion": "Use a candidate path, or run 'aioffice query <file> \"*\"' to list addressable nodes.",
    "candidates": ["/body/p[1]"]
  },
  "meta": { "file": "report.docx", "rev": "c73500e407fc", "elapsedMs": 143, "version": "0.1.0" }
}
```

### 诚实的公式 —— 真实输出

```bash
$ aioffice edit data.xlsx --ops '[{"op":"set","path":"/Sheet1/B1","props":{"value":"=SEQUENCE(3)"}}]'
```
```json
{
  "ok": true,
  "data": { "applied": 1, "ops": [{ "op": "set", "path": "/Sheet1/B1", "applied": ["formula"] }] },
  "meta": {
    "warnings": [{
      "code": "formula_not_evaluated",
      "message": "1 formula cell(s) use functions the built-in engine cannot evaluate: /Sheet1/B1. The formula text is saved without a cached value; Excel computes it when the file opens."
    }]
  }
}
```

## 安装

AIOffice 是**一个自包含的原生二进制**——无需 .NET 运行时、无需安装 Office、无额外文件。任选一种方式；每种方式都会下载与你平台匹配的二进制，并用发行版的 `SHA256SUMS` 做完整性校验。完整矩阵与平台注意事项见 **[docs/INSTALL.md](docs/INSTALL.md)**。

```bash
# 1) npm（Node ≥ 18）—— 全局安装，或用 npx 即用即走
npm install -g aioffice          # 把 `aioffice` 放到 PATH
npx aioffice doctor              # 一次性，无需安装（适合 CI / MCP 宿主）

# 2) Homebrew（macOS / Linux）
brew install onecer/tap/aioffice

# 3) 一行脚本（macOS / Linux；Windows 用 PowerShell——见下）
curl -fsSL https://raw.githubusercontent.com/onecer/AIOffice/main/dist/install.sh | sh

# 4) 直接下载 —— 从 release 页选你平台的 asset，校验 SHA256，chmod +x，放上 PATH：
#    https://github.com/onecer/AIOffice/releases/latest
```

Windows（PowerShell）：`irm https://raw.githubusercontent.com/onecer/AIOffice/main/dist/install.ps1 | iex`。
macOS 上直接下载的二进制首次运行可能被 Gatekeeper 拦截，执行一次 `xattr -d com.apple.quarantine <路径>`（脚本与 Homebrew 已为你处理；二进制尚未签名公证——见 [docs/SIGNING.md](docs/SIGNING.md)）。安装后 `aioffice version` 应输出 `{"ok":true,"data":{"name":"aioffice","version":"1.7.0",…}}`。

### 在你的 agent 里用它（MCP）

AIOffice 把**同一命令面暴露两次**——一个 CLI 和一个 stdio **MCP server**，17 个工具与动词一一对应。运行 `aioffice mcp --workspace <目录>`，把 agent 指过去：

```bash
aioffice mcp --workspace /path/to/your/documents     # stdio JSON-RPC，沙箱限定在 <目录>
```

各宿主（Claude Desktop / Claude Code / Cursor / 通用 stdio / TonoBraid）的配置见 **[docs/MCP-SETUP.md](docs/MCP-SETUP.md)**。把 agent 的系统提示指向 **[SKILL.md](SKILL.md)**——AI 上手指南（信封结构、寻址语法、读后写闭环）。可直接复制粘贴的任务配方见 **[docs/COOKBOOK.md](docs/COOKBOOK.md)**。

### 从源码构建

```bash
# 构建（需 .NET 10 SDK）
dotnet build AIOffice.sln

# 源码运行
alias aioffice='dotnet run --project src/AIOffice.Cli --'

# 或发布单文件二进制（约 36 MB，自包含）
dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release \
  -p:PublishSingleFile=true --self-contained
# rid 可选：osx-arm64 | osx-x64 | win-x64 | win-arm64 | linux-x64 | linux-arm64

aioffice doctor          # 环境 / 处理器 / 工作区体检
```

### 60 秒上手

```bash
# Word
aioffice create report.docx --title "Q3 报告"
aioffice edit   report.docx --ops '[
  {"op":"add","path":"/body","type":"p","position":"inside",
   "props":{"text":"季度业绩","style":"Heading1"}}]'
aioffice query  report.docx 'p[style=Heading1]'      # → 规范路径 + 文本片段
aioffice render report.docx --to html -o report.html

# Excel —— 公式写入即求值
aioffice create data.xlsx
aioffice edit   data.xlsx --ops '[
  {"op":"set","path":"/Sheet1/A1","props":{"value":21}},
  {"op":"set","path":"/Sheet1/A2","props":{"value":21}},
  {"op":"set","path":"/Sheet1/A3","props":{"value":"=SUM(A1:A2)"}}]'
aioffice get data.xlsx /Sheet1/A3                    # → "cachedValue": 42

# PowerPoint —— 看见自己做的东西
aioffice create deck.pptx          # 新建即自带 1 页空白幻灯片
aioffice edit   deck.pptx --ops '[{"op":"add","path":"/slide[1]","type":"slide",
                                   "position":"after","props":{"title":"Hello"}}]'
aioffice edit   deck.pptx --add '/slide[1]' --type shape text="AIOffice" x=2cm y=3cm w=10cm h=2cm
aioffice render deck.pptx --to svg --scope '/slide[1]' -o slide1.svg

# 安全网
aioffice snapshot list report.docx                   # 编辑前自动快照
aioffice snapshot restore report.docx 1              # 一次调用回滚
aioffice validate report.docx                        # OOXML 校验 + lint
```

## Markdown 进，Office 出

Agent 用 markdown 和 csv 思考。M5（v0.6.0）把它们变成正门——下面每条命令都原样来自发版冒烟：

```bash
# markdown -> 真 docx（标题、嵌套列表、管道表格、链接、加粗、代码）
aioffice create report.docx --from notes.md
aioffice read   report.docx --view outline     # 标题 + 规范路径
aioffice read   report.docx --view markdown    # ……再以 GFM 导出——结构 round-trip
aioffice validate report.docx                  # "valid": true，0 错误

# csv -> 类型化 xlsx（带引号的逗号值不丢，日期成日期，"007" 保持文本）
aioffice create orders.xlsx --from orders.csv
aioffice get    orders.xlsx /Sheet1/A2         # → "value": "007", "type": "text"
aioffice read   orders.xlsx --view csv         # 单 sheet 以 RFC 4180 csv 导出
```

源/目标不匹配会快速失败，suggestion 给出矩阵：`.md/.markdown → .docx，.csv/.tsv → .xlsx`。MCP 同一套接线：`office_create {file, from}` 与 `office_read {view:"markdown"|"csv"}`。

## LaTeX 进，公式出

M6（v0.7.0）新增了一个手写的 LaTeX → OOXML Math 转换器（无 LaTeX 依赖）。你写 LaTeX，文档里得到**真正的 Office Math**，Word 当作公式渲染。下面每条命令都原样来自发版冒烟：

```bash
# 新建文档，加行内公式 + 显示块（二次求根公式）+ 矩阵
aioffice create report.docx --title "M6 Report"
aioffice edit report.docx --ops '[{"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"E = mc^2"}}]'
aioffice edit report.docx --ops '[{"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"\\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}","display":true}}]'
aioffice edit report.docx --ops '[{"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"\\begin{pmatrix}1&0\\\\0&1\\end{pmatrix}","display":true}}]'

aioffice read report.docx --view text     # → "M6 Report$E = mc^2$"、"$$\frac{-b \pm \sqrt{b^2-4ac}}{2a}$$"、…
aioffice get  report.docx /body/p[1]/omath[1]   # → "latex": "E = mc^2", "display": false
aioffice validate report.docx              # "valid": true，0 错误

# 未知命令降级——文件仍校验通过，只是给一个警告
aioffice edit report.docx --ops '[{"op":"add","path":"/body/p[4]","type":"equation","props":{"latex":"\\foobar{x} + \\alpha"}}]'
# → ok:true + meta.warnings: [{ "code": "equation_partial", "message": "…\\foobar… appear literally…" }]
```

原始 LaTeX 以 `mc:Ignorable` 厂商属性存档在每个公式上，所以 `get` 原样返回你的源、round-trip 字节忠实。支持的子集——分式、根式、上下标、大型算符、矩阵（`pmatrix`/`bmatrix`/…）、`\left…\right`、重音、`\text`、希腊字母与算符——见 `aioffice help equations`。渲染出来公式是真的：

```bash
aioffice render report.docx --to png -o quadratic.png   # 下图正是这样产出的
```

![从 LaTeX 经真 Office Math 渲染出的二次求根公式](assets/demo/equation-quadratic.png)

## 交付前先审计

`aioffice audit` 是 Office 文件的无障碍 + 质量 lint。findings 是**数据**，不是错误——即使报出 error 级 finding，命令也返回 `0`，与 `validate` 完全一致。`--fix` 只应用**安全、非破坏性**的自动修复，并重审，让你看清还剩什么。

```bash
# 一个故意做坏的报告：图片无 alt、H1 后直接 H3、表格没有表头、空标题、没有文档标题。
$ aioffice audit report.docx
{ "ok": true, "data": {
  "findings": [
    { "code": "a11y_no_doc_title",   "severity": "warning", "path": "/properties",      "autofixable": true  },
    { "code": "a11y_no_alt_text",     "severity": "error",   "path": "/body/p[5]",        "autofixable": true  },
    { "code": "a11y_heading_skip",    "severity": "warning", "path": "/body/p[3]",        "autofixable": false },
    { "code": "quality_empty_heading","severity": "warning", "path": "/body/p[4]",        "autofixable": false },
    { "code": "a11y_no_table_header", "severity": "error",   "path": "/body/table[1]",    "autofixable": true  }
  ],
  "summary": { "errors": 2, "warnings": 3, "infos": 0 } } }

# --fix 应用安全自动修复（alt 文本、表头行、文档标题），再重审。标题跳级与空标题是
# 「仅报告」类 finding，会留在 remaining 里等你手工处理。
$ aioffice audit report.docx --fix
{ "ok": true, "data": {
  "findings": [
    { "code": "a11y_heading_skip",     "path": "/body/p[3]", "autofixable": false },
    { "code": "quality_empty_heading", "path": "/body/p[4]", "autofixable": false }
  ],
  "summary": { "errors": 0, "warnings": 2, "infos": 0 },
  "fixed": 3,
  "remaining": ["a11y_heading_skip#/body/p[3]", "quality_empty_heading#/body/p[4]"] } }

# 修复后仍是合法 OOXML。
$ aioffice validate report.docx
{ "ok": true, "data": { "valid": true, "count": 0, "issues": [] } }
```

同一个动词、同一套 findings、同一份 `--fix` 语义也覆盖 **.xlsx**（`#DIV/0!` 等公式错误、合并数据单元格、缺 alt 文本/标题）与 **.pptx**（图片 alt 文本、幻灯片标题、画布外形状、过小字号、阅读顺序）。用 `--category accessibility|quality|all` 与 `--severity error|warning|info` 限定范围；完整 code 清单与哪些可自动修复见 `aioffice help audit`。走 MCP 即 `office_audit {file, category?, severity?, fix?}`——第 15 个工具。

## 对比与审阅

`aioffice diff` **语义**对比文档与一份基线，返回**已排序、确定性**的变更清单——`added` / `removed` / `modified`（带 `before`/`after`）/ `moved`。变更是**数据**：哪怕两份文档差异巨大也 exit 0。基线既可以是**另一份同格式文件**，也可以是文档**自身的某个编辑前快照**。

```console
# 两份文件 —— 这版草稿相比上一版改了什么
$ aioffice diff new.docx old.docx
{
  "ok": true,
  "data": {
    "changes": [
      { "kind": "modified", "path": "/body/p[1]",
        "before": "Heading1", "after": "Heading2", "detail": "style" },
      { "kind": "modified", "path": "/body/p[2]",
        "before": "Second paragraph.", "after": "Second paragraph EDITED.", "detail": "text" },
      { "kind": "added", "path": "/body/p[4]", "detail": "paragraph" }
    ],
    "summary": { "added": 1, "removed": 0, "modified": 2, "moved": 0 },
    "baseline": "old.docx",
    "view": "detailed"
  }
}

# 对比快照 —— "我刚才到底改了什么"（每次编辑都会先自动快照）
$ aioffice edit report.docx --set '/body/p[2]' text='Revised'
$ aioffice diff report.docx --snapshot 1
#   → 那次编辑造成的唯一变更，带 before/after，确定性

# 简评模式 —— 仅计数 + 每条变更一行 path+kind
$ aioffice diff metrics.xlsx baseline.xlsx --view summary
```

变更清单始终按 `(path, kind)` 排序，因此同样两份文档在每个平台、每次运行 diff 结果完全一致。跨格式基线（`.docx` 对 `.xlsx`）返回 `invalid_args` 并指明不匹配。走 MCP 即 `office_diff {file, other? | snapshot?, view?}`——第 16 个工具。

## 格式互转

`convert` 把文档从一种格式转到另一种，按 `(源, 目标)` 扩展名对路由：同族文本桥复用既有代码（`docx↔md`、`xlsx↔csv`），跨格式对经一份**内容中立模型**（标题、段落、带格式 run、列表、表格、图片），`任意 → pdf/png/svg/html` 走 render 层。转换跨格式本质有损——目标无法承载的内容（动画、图表、SmartArt、精确样式、公式取值、markdown 不支持的颜色）honest 列入 `data.dropped` 并汇成单条 `convert_lossy` 警告。

```bash
# docx → pptx —— 每个标题一页，标题下的要点成为该页正文，表格成为原生 pptx 表格
$ aioffice convert report.docx deck.pptx
# → {"from":"docx","to":"pptx","blocksWritten":10,"dropped":[],"written":".../deck.pptx"}

# xlsx → docx —— 每个工作表一个标题 + 表格；公式以缓存的显示值跨格式
$ aioffice convert data.xlsx data.docx
# → docx 里 Sheet1/Totals 表格保存显示值（例如 =SUM(...) → 30）

# 任意格式 → markdown —— 经中立模型，故 xlsx 与 pptx 也能转 md
$ aioffice convert report.docx report.md      # 标题、要点、管道表格
$ aioffice convert report.md roundtrip.docx   # 再转回去 —— 大纲 round-trip

# 任意格式 → PDF —— 经系统浏览器分页打印
$ aioffice convert report.docx report.pdf

# 护栏
$ aioffice convert a.docx b.docx              # invalid_args：「use edit」（同格式）
$ aioffice convert a.docx a.xyz               # unsupported_feature，列出受支持目标
```

从含图表的工作簿转出的 deck 会**报告损失**而非隐藏：`meta.warnings: [{ "code": "convert_lossy", "message": "Some content did not survive the conversion: charts (…the chart is not converted)." }]`。走 MCP 即 `office_convert {src, dest}`——第 17 个工具。完整 `(src → dest)` 矩阵见 `aioffice help convert`。

## 1.7 新增

1.7.0 是第六个 **1.0 之后的功能版本** —— 纯**增量**，故 `surfaceVersion` 保持 **`1.0`**（以下全部位于冻结的 [CONTRACT.md](CONTRACT.md) 契约线内，只增不删、不改名；op **kinds** 不变；**无新动词 / 无新 MCP 工具**）。它**补齐打印就绪**、新增**相机工具**与**计算设置**、并**深化公式引擎**。18 动词 / 17 MCP 工具不变。

- **xlsx 打印完整性** —— `set` 于 `/SheetN`：`printTitleRows`/`printTitleCols`（每页重复标题带）、`pageBreaks`（`{rows,cols}`）、`fitToPage`（`{fitToWidth,fitToHeight}` 或 `{scale}`）、`centerHorizontally`/`centerVertically`、`printGridlines`/`printHeadings`、`printHeader`/`printFooter`（`{left,center,right}` 字段码串：`&P` `&N` `&D` `&T` `&F` `&A`）。`get /SheetN` 回报。详见 `aioffice help print-setup`。
- **xlsx 链接图片 / 相机工具** —— `add type:linkedPicture` 把单元格区域镜像为图片；诚实的**静态快照**（活链接对 validator 脆弱），带 `linked_picture_static` 警告。寻址 `/SheetN/linkedPicture[i]`。
- **xlsx 计算设置** —— `set /`（工作簿根）：`calculationMode`（`auto`/`manual`/`autoExceptTables`）、`iterativeCalc`、`maxIterations`、`maxChange`、`fullPrecision`。`get /` 现回报 `{calculation, workbookProtection}`（此前 `unsupported_feature`，处理过该错误的 agent 仍可用）。
- **docx 首字下沉、图片水印、新域种类、带编号公式** —— 段落 `set {dropCap,dropCapLines,dropCapFont}`；`add type:watermark {image,washout}` 图片水印；`add type:field` 新增 `styleRef`（STYLEREF）/`symbol`（SYMBOL）/`quote`（QUOTE）；`add type:equation {display:true,number:…}` 带编号显示公式，寻址 `/equation[@num=…]`。
- **更深的公式（docx + pptx 共享 Core 转换器）** —— `\begin{aligned|gathered|cases}` 公式数组、`\binom`/`\dbinom`/`\tbinom`、`\overbrace`/`\underbrace`、多重积分 `\iint`/`\iiint`、更多重音与关系/箭头。**向后兼容** —— 更多 LaTeX 渲染（`equation_partial` 警告更少）；pptx 公式因共享转换器自动获得。详见 `aioffice help equations`。
- **pptx 备注/讲义母版、动画时间、表格单元格对齐** —— `set /notesMaster`（`{background,bodyFont}`）与 `set /handoutMaster`（`{background,headerFooter,slidesPerPage}`）；动画 `repeat`/`rewind`/`autoReverse`（`set /slide[i]/animation[k]`）；表格单元格 `valign`/`marginLeft|Right|Top|Bottom`/`textDirection`。详见 `aioffice help masters`。

这些都暴露在 `schema` / `office_schema`，在 `aioffice help` 下有文档（新主题 `print-setup`、`masters`；深化 `equations`/`animations`/`properties-*`/`addressing`），并由 `SchemaConsistencyTests` 守卫两面一致（MCP 工具面仍在 3,500 token 预算内）。新增 `add` type（`linkedPicture`）、新 set 路径（`/notesMaster`、`/handoutMaster`）、新寻址（`/SheetN/linkedPicture[i]`、`/equation[@num=…]`）、新 prop keys、两个新警告码（`linked_picture_static`、`equation_numbers_cached`）—— 全部增量。详见 [CONTRACT.md §7g](CONTRACT.md) 与 [CHANGELOG.md](CHANGELOG.md)。**2331 测试**跨 7 项目全绿，macOS + Windows 双平台（Core 180 · Word 681 · Excel 619 · Pptx 709 · MCP 87 · Preview 24 · Render 31）。

## 1.6 新增

1.6.0 是一个**分发与上手版本——能力零变化。** 原生二进制与 1.5.0 字节级同面：**18 个 CLI 动词 / 17 个 MCP 工具 / `surfaceVersion` `1.0`**，冻结的 [CONTRACT.md](CONTRACT.md) 契约线、每个信封、错误码、寻址形式全部不变。1.6 新增的是二进制*周边*的一切，让 agent 与人都能一行装好、一键接入：

- **npm 包** —— `npm i -g aioffice`（或 `npx aioffice …`）。一个极薄的 wrapper，安装时从 GitHub release 下载与平台匹配的原生二进制并做 SHA256 校验；二进制本身不打进 tarball。见 [`npm/README.md`](npm/README.md) 与 `AIOFFICE_DOWNLOAD_VERSION` / `AIOFFICE_DOWNLOAD_BASEURL` 覆盖项。
- **Homebrew formula** —— `brew install onecer/tap/aioffice`，安装预编译的逐平台二进制（[`dist/Formula/aioffice.rb`](dist/Formula/aioffice.rb)）。
- **一行安装脚本** —— `dist/install.sh`（macOS/Linux，POSIX sh）与 `dist/install.ps1`（Windows PowerShell）：检测平台 → 下载 → SHA256 校验 → 安装 → PATH 提示；macOS quarantine 自动剥除。
- **上手文档** —— [SKILL.md](SKILL.md)（AI 上手指南）、[docs/COOKBOOK.md](docs/COOKBOOK.md)（10 个复制即用配方）、[docs/INSTALL.md](docs/INSTALL.md)（四种安装方式）、[docs/MCP-SETUP.md](docs/MCP-SETUP.md)（Claude Desktop / Claude Code / Cursor / 通用 stdio / TonoBraid 配置）、[docs/SIGNING.md](docs/SIGNING.md)（签名 / 公证路线图）。

整个版本都在 C#/.NET 源码树**之外**——构建与测试不受影响。**2125 测试**跨 7 项目全绿（Core 124 · Word 612 · Excel 572 · Pptx 675 · MCP 87 · Preview 24 · Render 31）。发布方式：`cd npm && npm publish --access public`（先 `npm login`）；创建 `onecer/homebrew-tap`，把 `dist/Formula/aioffice.rb` 作为 `Formula/aioffice.rb`（用 v1.6.0 的 `SHA256SUMS` 填四个 `sha256` 值）；安装脚本直接从 `main` 提供。详见 [CHANGELOG.md](CHANGELOG.md)。

## 1.5 新增

1.5.0 是第五个 **1.0 之后的功能版本** —— 纯**增量**，故 `surfaceVersion` 保持 **`1.0`**（以下全部位于冻结的 1.0 契约线内，只增不删、不改名；op **kinds** 不变）。**它填平了现代标量函数缺口**（`XLOOKUP`/`IFS`/`SWITCH`/`LET`/… 现在会求值），并补全表格**模拟运算工具箱**（方案管理器 + 单变量求解）；同时 Word 新增表格单元格公式、构建基块与行号，PowerPoint 新增嵌入字体、动作按钮与自定义版式。18 动词 / 17 MCP 工具不变。

- **标量函数求值（xlsx）** —— `=XLOOKUP`、`=IFS`、`=SWITCH`、`=LET`、`=MAXIFS`、`=MINIFS`、`=AVERAGEIFS`（以及会溢出的 `=TEXTSPLIT`）现在写入时**求值**并写入缓存值，无头读取者看到真实结果，不再触发 `formula_not_evaluated`。`TEXTJOIN`/`CONCAT`/`IFERROR`/`SUMIFS`/`COUNTIFS` 此前已求值，照常工作。
- **方案管理器（xlsx）** —— `add type:scenario`（`{name, cells:{addr:value,…}, comment?}`）保存命名的可变单元格集合；`set {applyScenario:"name"}` 写入这些值并**重算**依赖项。`/Sheet1/scenario[@name=…]` 寻址。
- **单变量求解（xlsx）** —— 在可变单元格上 `set {goalSeek:{targetCell, targetValue}}` 求解使目标公式达标的输入值（牛顿法 + 二分回退），设置该值并重算；无解 → 新警告 `goal_seek_no_solution`（单元格保持不变）。
- **表格单元格公式（docx）** —— 表格单元格上的 `formula` prop（`=SUM(ABOVE)`、`=AVERAGE(LEFT)` 或形如 `=A1*B2` 的单元格引用算式）会变成 `w:fldSimple` 域，值无头计算并缓存。可选 `numberFormat` 修饰；输入为域时触发 `table_formula_cached` 警告。
- **构建基块（docx）** —— `add type:buildingBlock` 把可复用的自动图文集 / 文档部件内容存入词汇表部件；`add type:buildingBlockRef` 把已存基块的内容插入正文。`/buildingBlock[@name=…]` 寻址。
- **行号（docx）** —— `/section[i]` 的 `lineNumbers` `set` prop（`{start, increment, restart, distance?}`，或 `"none"`）。
- **嵌入字体（pptx）** —— `add type:font` 在 `/fonts` 上把经沙箱解析的 `.ttf`/`.otf` 嵌为字体部件并注册 `p:embeddedFont`（默认 regular 槽，`embedAll` 填满四槽）。`src` 必填且须在工作区内。`/fonts/font[@name=…]` 寻址。
- **动作按钮（pptx）** —— 幻灯片上 `add type:actionButton`（`{action, target?, …}`；`action` ∈ `first|last|next|prev|home|end|slide|url`）—— 基于 M8 形状超链接的导航按钮。
- **自定义版式（pptx）** —— M6 的 `add type:layout` 新增 `placeholders` prop（`[{type, x, y, w, h}]`），在母版上构建全新的 `slideLayout` 部件；幻灯片按名绑定（`add type:slide {layoutName:"Hero"}`）。

这些都暴露在 `schema` / `office_schema`，在 `aioffice help` 下有文档（新主题 `scenarios`、`goal-seek`、`table-formulas`、`building-blocks`、`embedded-fonts`、`action-buttons`、`layouts`、`line-numbers`；`formulas` 补标量函数），并由 `SchemaConsistencyTests` 守卫两面一致（MCP 工具面仍在 3,500 token 预算内——新表面详情入 `office_help`，工具 schema 不膨胀）。新增 `add` type（`scenario`/`buildingBlock`/`buildingBlockRef`/`font`/`actionButton`）、新 prop keys（`applyScenario`/`goalSeek`/单元格 `formula`/`lineNumbers`/`embedAll`/版式 `placeholders`）、新寻址（`/Sheet1/scenario[@name=…]`、`/buildingBlock[@name=…]`、`/fonts`）、两个新警告码（`goal_seek_no_solution`、`table_formula_cached`）—— 全部增量。标量函数求值**向后兼容**：曾带 `formula_not_evaluated` 的单元格现带缓存值，处理过该警告的 agent 仍可用。详见 [CONTRACT.md §7e](CONTRACT.md) 与 [CHANGELOG.md](CHANGELOG.md)。**2125 测试**跨 7 项目全绿，macOS + Windows 双平台。

## 1.4 新增

1.4.0 是第四个 **1.0 之后的功能版本** —— 纯**增量**，故 `surfaceVersion` 保持 **`1.0`**（以下全部位于冻结的 1.0 契约线内，只增不删、不改名；op **kinds** 不变）。**它填平了长期存在的动态数组缺口**：`FILTER`/`UNIQUE`/`SORT`（及同类）现在会求值并溢出，而不再只是被识别。18 动词 / 17 MCP 工具不变。

- **动态数组求值与溢出（xlsx）** —— 把单元格 `value` 设为 `=FILTER`/`=UNIQUE`/`=SORT`/`=SORTBY`/`=SEQUENCE`/`=RANDARRAY`/`=TRANSPOSE` 现在会**求值**并把结果数组**溢出**到以该单元格为锚点的矩形（锚点保留数组公式，每个溢出单元格带缓存值）。这些不再触发 `formula_not_evaluated`；`get` 锚点回报 `spillRange`。`RANDARRAY` 按参数确定性播种以保往返稳定。
- **财务函数（xlsx）** —— `RATE`/`IRR`/`XIRR`/`NPV`/`PV`/`FV`/`PMT`/`NPER` 写入时求值（迭代者用牛顿法 + 二分回退）并写入缓存数值 —— 不再触发 `formula_not_evaluated`。
- **模拟运算数据表（xlsx）** —— `add type:dataTable` 在区域上建单/双变量数据表（`{rowInput?, colInput?}`）；角公式沿输入轴重算入带 Excel `{=TABLE(…)}` 构造的缓存正文。`/Sheet1/dataTable[i]` 寻址。溢出受阻触发新错误码 `spill_blocked`（exit 2）。
- **邮件合并执行（docx）** —— `template` 动词在 `--data` 为 JSON **数组**记录时执行合并：用新 **`--output`** 模式每记录一份（`{n}` = 记录序号、`{Field}` = 记录值；每条展开路径经沙箱），或省略时一份合并节文档。`office_template` 增可选 `output` 参 —— 仍 17 工具。单对象 `--data` 保持单文档填充不变。
- **IF 域（docx）** —— `add type:ifField` 插入 Word «IF» 域（`{field, operator, value, trueText, falseText}`），合并时按记录求值。
- **页面边框（docx）** —— `/section[i]` 的 `pageBorder` `set` prop（`{style, color?, widthPt?, sides?}`，或 `"none"`）。
- **缩放导航（pptx）** —— `add type:zoom` 在幻灯片上加 slide/section/summary 缩放跳转对象（`{kind, target?, x?, y?, w?, h?}`），`/slide[i]/zoom[k]` 寻址。
- **点击触发动画（pptx）** —— `add type:animation` 的 `triggerOn:"@N"` prop，点击另一形状（稳定 id `N`）时播放效果。
- **表格样式（pptx）** —— `add type:table` 接受内建 `style`（`none|light1|light2|medium1|medium2|medium3|dark1|dark2`）+ 条带/边缘标志 `firstRow`、`lastRow`、`bandRow`、`firstCol`。

这些都暴露在 `schema` / `office_schema`，在 `aioffice help` 下有文档（新主题 `formulas`、`data-tables`、`mail-merge`、`page-borders`、`zoom`、`table-styles`；`animations` 补 `triggerOn`），并由 `SchemaConsistencyTests` 守卫两面一致（MCP 工具面仍在 3,500 token 预算内——新表面详情入 `office_help`，工具 schema 不膨胀）。新增 `add` type（`dataTable`/`ifField`/`zoom`）、新 prop keys（`pageBorder`/`triggerOn`/表格 `style`）、新寻址（`/Sheet1/dataTable[i]`、`/slide[i]/zoom[k]`）、扩展的 `template` 行为、一个新错误码（`spill_blocked`）—— 全部增量。动态数组/财务求值是**向后兼容**的：曾带 `formula_not_evaluated` 的单元格现带缓存值，处理过该警告的 agent 仍可用。详见 [CONTRACT.md §7d](CONTRACT.md) 与 [CHANGELOG.md](CHANGELOG.md)。**2016 测试**跨 7 项目全绿，macOS + Windows 双平台。

## 1.3 新增

1.3.0 是第三个 **1.0 之后的功能版本** —— 纯**增量**，故 `surfaceVersion` 保持 **`1.0`**（以下全部位于冻结的 1.0 契约线内，只增不删、不改名；op **kinds** 不变）。18 动词 / 17 MCP 工具不变。

- **图表润色（xlsx + pptx）** —— 既可在 `add` 图表时带上，也可经 `set` 作用于既有图表路径（`/Sheet1/chart[i]`、`/slide[i]/chart[k]`）：`dataLabels`（`true` 或 `{show, position?}`）、`legend`（`none|right|left|top|bottom`）、`axisTitles`（`{category?, value?}`）、`trendline`（`none|linear|exponential|movingAverage`）、`errorBars`（`none|stdErr|stdDev|percent`）、`gridlines`（`{major?, minor?}`）、`secondaryAxis`（命名系列移到次值轴）。`get` 在 `polish` 子对象回报。
- **高级条件格式（xlsx）** —— 三类新 `conditionalFormat`：`formula`（`=$B1>100` 表达式规则）、`topBottom`（top/bottom N 或 N%）、`aboveBelowAverage`（高于/低于区域均值，可选 `stdDev`）。
- **透视表计算字段（xlsx）** —— `add type:pivot` 新增 `calculatedFields`：`[{name, formula}]` 由源列名计算（如 `Margin = Revenue - Cost`），add 时校验、`get` 回报。
- **正文形状与文本框（docx）** —— `add type:shape`（浮动 DrawingML 形状：`rect|roundRect|ellipse|line|arrow`，`/body/shape[i]`）与 `add type:textBox`（`/body/textBox[i]`），各带 fill/line/内联文本。
- **传统表单域（docx）** —— `add type:formField`（`text|checkbox|dropdown`），`/formField[@name=…]` 寻址、`set` 改值、`read --view fields` 列出。
- **主题编辑（docx）** —— `set /theme` 改主题色彩方案（`dk1/lt1/dk2/lt2`、`accent1`…`accent6`、`hlink`、`folHlink`）与字体（`majorFont`、`minorFont`）；`get /theme` 回报。
- **3D 模型（pptx）** —— `add type:model3d` 把 `.glb`/`.gltf` 嵌入为真实 3DModel 媒体部件，置于海报图片回退之后（PowerPoint 2019+ 渲染）；`src`/`poster` 沙箱解析，随附 `model3d_as_media` 警告。`/slide[i]/model3d[@id=N]` 寻址。
- **动作路径动画（pptx）** —— 新增 `motionPath` 动画效果，`path` 取 `line|arc|circle|custom`（custom 取归一化 `points` 列表）；`read --view structure` 列出。

这些都暴露在 `schema` / `office_schema`，在 `aioffice help` 下有文档（新主题 `chart-polish`、`conditional-format`、`themes`、`3d-models`、`form-fields`、`animations`；扩展了 `properties-docx` / `properties-xlsx` / `properties-pptx`），并由 `SchemaConsistencyTests` 守卫两面一致（MCP 工具面仍在 3,500 token 预算内——新表面详情入 `office_help`，工具 schema 不膨胀）。新增 `add` type（`shape`/`textBox`/`formField`/`model3d`）、新 prop keys、`/theme` 寻址形式、一个新警告码（`model3d_as_media`）—— 全部增量。详见 [CONTRACT.md §7c](CONTRACT.md) 与 [CHANGELOG.md](CHANGELOG.md)。**1924 测试**跨 7 项目全绿，macOS + Windows 双平台。

## 1.2 新增

1.2.0 是第二个 **1.0 之后的功能版本** —— 纯**增量**，故 `surfaceVersion` 保持 **`1.0`**（以下全部位于冻结的 1.0 契约线内，只增不删、不改名；op **kinds** 不变——`group`/`ungroup` 是 `add` **type** 值而非新 op kind）。18 动词 / 17 MCP 工具不变。

- **SmartArt 创建（pptx，创建 + 只读）** —— `add type:smartart` 写出真正的图：`layout` 取 `list`/`process`/`hierarchy`/`orgChart`/`cycle`，各映射到内置 PowerPoint 版式，打开时按 layout+data 重生成。`nodes` 为扁平 `{text, level}` 列表（0-based level 构建层级）。`office_get /slide[i]/smartart[k]` 回报 layout 与节点树；就地改节点报 `unsupported_feature`（重建即可）。
- **连接符（pptx）** —— `add type:connector` 在两个形状间（`@id` 或名）布 `p:cxnSp`：`kind` straight/elbow/curved、`startArrow`/`endArrow` none/arrow/triangle、`color`/`width`/`name`。
- **形状组合/分组（pptx）** —— `add type:group` 把 ≥2 个形状包成组（`/slide[i]/group[@id=N]`，子形状寻址 `…/group[@id=N]/shape[@id=M]`）；`add type:ungroup` 溶解组、子形状带绝对坐标回到幻灯片。
- **图表目录（docx）** —— `add type:tableOfFigures` 列出 Figure/Table/Equation 题注；条目来自缓存题注，随附 `figures_cached` 警告。
- **索引（docx）** —— `add type:indexEntry` 标记 `XE` 域；`add type:index` 按字母构建索引（`columns`），页码缓存随附 `index_cached` 警告。
- **邮件合并域（docx）** —— `add type:mergeField` 插入 `MERGEFIELD`，由 `template` 动词**用与 `{{key}}` 同一份 `--data`** 按名填充。
- **表单控件（xlsx）** —— `add type:formControl` 写出 `checkbox`/`optionButton`/`spinner`/`comboBox`/`listBox`/`button`，带 `linkedCell`、`items`/`listFillRange`、`min`/`max`/`increment`。
- **保护（xlsx）** —— 单元格 `locked`；sheet 路径 `protected`（+ `password`、`allow*` 标志）；工作簿根 `protectStructure`（+ `protectWindows`）。Excel 轻量 UI 保护（非加密），AIOffice 始终拥有并可解除。
- **numberFormat 命名预设（xlsx）** —— `numberFormat` 接受命名预设：`accounting-usd`、`currency-usd/eur/gbp/jpy`、`percent`、`scientific`、`date-iso`、`datetime-iso`、`duration`… 预设解析为对应 Excel 格式码；非预设串按字面保留。

这些都暴露在 `schema` / `office_schema`，在 `aioffice help` 下有文档（新主题 `smartart`、`connectors`、`number-formats`、`structural-fields`；扩展了 `properties-pptx` / `properties-xlsx`），并由 `SchemaConsistencyTests` 守卫两面一致（MCP 工具面仍在 3,500 token 预算内）。详见 [CONTRACT.md §7b](CONTRACT.md) 与 [CHANGELOG.md](CHANGELOG.md)。

## 1.1 新增

1.1.0 是首个 **1.0 之后的功能版本** —— 纯**增量**，故 `surfaceVersion` 保持 **`1.0`**（以下全部位于冻结的 1.0 契约线内，只增不删、不改名）。18 动词 / 17 MCP 工具不变。

- **图表新增 7 类，xlsx 与 pptx 都支持** —— 两个图表工厂各自新增 `doughnut`、`radar`、`bubble`、`stackedBar`、`percentStackedBar`、`stackedArea`、`combo`，在既有 `bar`/`line`/`pie`/`scatter`/`area` 之上。`bubble` 每点一个 x/y/size 三元组；`combo` 首系列画柱、其余画线（≥2 系列）。未支持类型仍返回 `unsupported_feature` 并列出扩展后的集合。
- **iconSet 条件格式（xlsx）** —— 新的 `conditionalFormat` 类型，按排名画 3/4/5 图标集（`3TrafficLights1`、`3Arrows`、`4Rating`、`5Quarters`…），可选 `reverse`/`showValue`。
- **引用与参考文献（docx）** —— `add type:source` 写入文献库，`add type:citation` 落 `CITATION` 域，`add type:bibliography` 按 `APA`/`MLA`/`Chicago` 渲染所有**已引用**来源。`read --view sources` 列出文献库；参考文献为缓存（Word 按 `F9` 重建），随附 `bibliography_cached` 警告。
- **嵌入媒体（pptx）** —— `add type:media` 把音视频（mp4/mov/m4a/mp3/wav）嵌为 `p:pic` + `a:videoFile`/`a:audioFile`；`src` 与可选 `poster` 图片均经沙箱解析（越界路径 `sandbox_denied`）。
- **文字与外观效果** —— docx run（Word 2010 `w14:` 效果）或 pptx 形状（`a:effectLst`）上 `set` `shadow`/`glow`/`reflection`/`outline`；每项取 `true`（默认）或颜色/对象，`false` 清除。
- **新切换（pptx）** —— `split`、`reveal`、`cut`、`zoom`，在 `none`/`fade`/`push`/`wipe` 之上。

这些都暴露在 `schema` / `office_schema`，在 `aioffice help` 下有文档（新主题 `docx/citation`、`docx/effect`、`pptx/media`、`pptx/effect`；扩展了 `charts` / `conditional-format` / `transition`），并由 `SchemaConsistencyTests` 守卫两面一致。详见 [CONTRACT.md §7a](CONTRACT.md) 与 [CHANGELOG.md](CHANGELOG.md)。

## MCP（接入 Claude 等 agent）

```bash
aioffice mcp     # stdio MCP 服务器 —— 17 个工具，与 CLI 动词 1:1
```

Claude Desktop / Claude Code 配置：

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

| MCP 工具 | CLI 动词 | MCP 工具 | CLI 动词 |
|---|---|---|---|
| `office_create` | create | `office_template` | template |
| `office_read` | read | `file_snapshot` | snapshot |
| `office_query` | query | `office_status` | doctor |
| `office_get` | get | `office_help` | help |
| `office_edit` | edit | `office_schema` | schema |
| `office_render` | render | `preview_open` | preview open |
| `office_validate` | validate | `preview_selection` | preview selection |
| `office_audit` | audit | `office_diff` | diff |
| `office_convert` | convert | | |

`preview_open` / `preview_selection`（实时预览 + 人工点选回读）已于 M1（v0.2.0）注册；`office_audit`（无障碍 + 质量 lint）是第 15 个工具，M7（v0.8.0）加入；`office_diff`（语义对比）是第 16 个工具，M8（v0.9.0）加入；`office_convert`（跨格式互转）是第 17 个工具，M9（v0.10.0）加入。全部工具 schema 的 token 预算上限 3,500——由测试强制，加入 `office_convert` 后仍在预算内。

## 命令面（v1.7.0）

| 动词 | 说明 |
|---|---|
| `create <file> [--from notes.md\|data.csv] [--kind] [--title]` | 新建文档（类型按扩展名推断）——或**导入**：`.md` → `.docx`、`.csv` → `.xlsx` |
| `read <file> [--view outline\|text\|stats\|structure\|properties\|embeds\|markdown\|csv]` | 低成本巡检投影，可分页；`properties` 返回 core+custom 文档属性于 `data.properties.{core,custom}`（三格式统一），`embeds` 列出嵌入的 OLE/包对象（三格式），`markdown` 把 docx body 导出为 GFM，`csv` 导出单个 xlsx sheet（`--sheet`、`--range`） |
| `query <file> <selector>` | CSS 风格选择器 → 规范路径（`p[style=Heading1]`、`cell[value>100]`、`shape:contains('Q3')`） |
| `get <file> <path>` | 取单个节点及属性 |
| `edit <file> --ops <json\|@file>` | **原子**批量 set/add/remove/move/replace/accept/reject/**extract** · `--dry-run` · `--expect-rev` · 语法糖 `--set/--add/--remove` · add 类型含 **`embed`**（任意文件作为 OLE/包对象，三格式）与 **`equation`**（LaTeX → OMML 公式，docx+pptx）；**`extract`** 把嵌入对象的字节导出 · **文档级查找替换糖** `--find X --replace Y [--regex] [--match-case] [--whole-word]`（docx 正文+全部页眉页脚、每个 sheet、每页 slide 含备注；聚合返回 `{replacements, locations}`；docx 配 `--track` 时逐命中生成修订对） |
| `render <file> [--to html\|svg\|text\|png\|pdf] [--scope]` | "看"的那一步——docx/xlsx 出 html，pptx 逐页出 svg，**png/pdf** 经系统浏览器输出（pptx pdf：整副 deck，一页一张幻灯片） |
| `validate <file>` | OOXML 校验 + lint，带修复建议 |
| `audit <file> [--category accessibility\|quality\|all] [--severity error\|warning\|info] [--fix]` | **无障碍 + 质量 lint**——findings 是数据（`ok:true`，exit 0）；`--fix` 只做安全自动修复（alt 文本、表头行、文档/幻灯片标题、孤儿书签）并返回 `{fixed, remaining}` |
| `diff <file> [<other>] [--snapshot N] [--view summary\|detailed]` | **语义对比**——对比同格式文件*或*文档自身的某个快照，返回已排序、确定性的 `{changes:[added/removed/modified/moved], summary, baseline}`（变更是数据，exit 0）；`--view summary` 裁剪为计数 + path+kind |
| `convert <src> <dest>` | **跨格式互转**——docx/xlsx/pptx ↔ 互转（内容中立模型）、docx↔md、xlsx↔csv、任意→pdf/png/svg/html。跨格式有损：`data.dropped` + `convert_lossy` 警告列出未存活内容；同扩展名 → `invalid_args`，未知目标 → `unsupported_feature` |
| `template <file> --data <json\|@file>` | `{{key}}` 模板合并（跨 run 拆分安全） |
| `snapshot <list\|restore> <file> [n]` | 编辑前快照环（20 份） |
| `preview <open\|selection\|close> <file> [--port N]` | localhost 实时预览；人点选 → `selection` 返回规范路径 |
| `doctor` | 环境 / 运行时 / 处理器诊断 |
| `schema [verb]` | 整个命令面的机器可读 JSON |
| `help [topic]` | addressing · selectors · properties-docx/xlsx/pptx · errors · **equations**（docx+pptx）· **embeds** · **rtl** · **sections** · **audit** · **diff** · **convert** · **docx/citation** · **docx/effect** · **pptx/media** · **pptx/effect** |
| `mcp` | stdio MCP 服务器 |
| `version` | 版本信息 |

全局参数：`--json`（非 TTY 默认）· `--pretty` · `--workspace <dir>`（沙箱根，默认 cwd，亦可 `AIOFFICE_WORKSPACE`）· `--quiet`。
退出码：`0` 成功 · `2` 用户错误 · `3` 内部/格式错误 · `4` sandbox_denied · `5` unsupported_feature。

**寻址语法**（1 起始）：`/body/p[3]` · `/body/table[1]/tr[2]/tc[1]` · `/body/p[3]/omath[1]` · `/Sheet1/A1:C10` · `/Sheet1/table[@name=Sales]` · `/Sheet1/row[2]:row[6]` · `/'Q3 Data'/B2` · `/slide[2]/shape[3]` · `/section[1]` · `/master[1]/layout[2]` · `/`（pptx 幻灯片尺寸 + 分节）· **M7**：`/properties`（文档 core + custom 属性）· `/sdt[@tag=status]`（docx 内容控件）· `/style[@name=Currency-Red]`（xlsx 命名单元格样式）· **M8**：`/caption[@label=Figure][1]`（docx 题注）· `/Sheet1/slicer[1]`（xlsx 切片器）· **M10**：`/embed[1]` · `/Sheet1/embed[1]` · `/slide[2]/embed[@id=7]`（嵌入对象）· `/slide[2]/shape[@id=7]/omath[1]`（pptx 公式）。

**面向 AI 的契约** —— 稳定的 v1.0 命令面（信封结构、冻结的错误码、寻址语法、退出码、op/view/工具词表，以及[已知限制](CONTRACT.md#10-known-limitations-what-aioffice-does-not-do)）记录于 **[CONTRACT.md](CONTRACT.md)**，并通过 `schema` 与 `doctor` capabilities 中的 `surfaceVersion` 字段（`1.0`）声明。

## 当前能力（M0 + M1 + M2 + M3 + M4 + M5 + M6 + M7 + M8 + M9 + M10 + 1.1）

| 格式 | M0（v0.1.0） | + M1（v0.2.0） | + M2（v0.3.0） | + M3（v0.4.0） | + M4（v0.5.0） | + M5（v0.6.0） | + M6（v0.7.0） | + M7（v0.8.0） | + M8（v0.9.0） | + M9（v0.10.0） | + M10（v0.11.0） | + 1.1（v1.1.0） |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| **.docx** | 创建 · 段落/标题/样式 · 表格 · 文本与格式编辑（加粗/斜体/颜色/对齐/字号）· query/get · outline/text/stats/structure 视图 · HTML 渲染 · `{{key}}` 模板 · 校验 | **页眉/页脚**（创建 + 编辑，`/header[1]/p[1]`）· PNG 渲染 · 实时预览 | **修订**（`--track --author`，`read --view revisions`，按 `/revision[@id=N]` 或范围 accept/reject）· **批注**（增/读/删，`/comment[@id=N]`）· **自定义样式**（`/styles` add，`/style[@id=X]` set/get/remove）· **图片**（PNG/JPEG，src 必经沙箱，缺省守纵横比） | **列表**（编号/项目符号，多级嵌套与重新编号；text 视图带 `1.`/`•` 标记，HTML 渲染输出真 `<ol>/<ul>`）· **超链接**（外链 + 书签锚点）· **书签** · **脚注** · **页面设置**（`/section[1]`：纸张/方向/边距）· **格式修订 accept/reject**（w:rPrChange/w:pPrChange）· **批注线程回复**（在 `/comment[@id=N]` 上 `add type:reply`） | **目录 TOC**（`add type:toc`，levels/title/position；`/toc[1]` get 回报 entryCount）· **文本水印**（`add type:watermark`，写入每个页眉，无页眉自动建）· **尾注**（`/endnote[@id=N]`）· **分节符**（`add type:sectionBreak`，逐节独立页面设置——同一文件横竖混排）· **查找替换**（跨 run 匹配；`--track` 时逐命中生成 w:del+w:ins 修订对） | **markdown 桥**（`create --from notes.md` 导入 GFM——标题/列表/表格/链接/代码；`read --view markdown` 导出回去，结构 round-trip）· **深表格**（`mergeRight`/`mergeDown`、边框 all/outer/none、底纹、`headerRow` 跨页重复、`columnWidths`、valign；HTML 渲染出真 colspan/rowspan）· **域**（PAGE/NUMPAGES/DATE/TITLE + `leadingText`——「Page X of Y」页脚）· **首页/奇偶页眉页脚变体**（`/header[firstPage]`，自动接线 `w:titlePg`/`w:evenAndOddHeaders`） | **公式**（`add type:equation`——LaTeX → 真 Office Math；行内 `/body/p[i]/omath[j]` 或显示块；text 视图出 `$…$`/`$$…$$`；未知命令降级 + `equation_partial` 警告；原始 LaTeX 存档可忠实回读）· **右到左/双向**（段落 `w:bidi` / run `w:rtl` / 表格 `w:bidiVisual` 的 `rtl` prop）· **多栏分节**（`/section[1]` 的 `columns`/`columnGap` + `add type:columnBreak`） | **审计**（`audit report.docx [--fix]`——无 alt 文本、标题跳级、表头行缺失、低对比度、无文档标题、空/坏标题、坏链接、孤儿书签；alt/表头/标题/书签可安全自动修复）· **文档属性**（`set /properties` core + 类型化 custom；`read --view properties`）· **内容控件**（`add type:contentControl` text/dropdown/date/checkbox；`/sdt[@tag=X]`；`read --view fields`）· **图片 alt 文本**（`set {alt}`） | **diff**（`diff new.docx old.docx`——段落/表格/页眉页脚/属性变更归为 added/removed/modified/moved，确定性）· **题注**（`add type:caption` Figure/Table/Equation + SEQ；`/caption[@label=Figure][i]`）· **交叉引用**（`add type:crossRef` REF/PAGEREF；`labelAndNumber`/`numberOnly`/`text`/`page`） | **convert**（`convert report.docx deck.pptx` / `report.xlsx` / `report.md` / `report.pdf`——经内容中立模型跨格式互转，docx↔md 文本桥，任意→pdf/png/svg/html；`convert_lossy` 列出丢失项）· **能力自省**（`doctor` 携带动词/工具数、格式、convert + render 目标、audit 类别） | **嵌入对象**（`add type:embed` 把任意文件作为 OLE/包对象放进 `/body`/tc/页眉/页脚；`read --view embeds`；`extract` 把字节按位导出；`/embed[i]`）· **统一属性结构**（`read --view properties` → `data.properties.{core,custom}`，三格式一致） | **引用与参考文献**（`add type:source`/`citation`/`bibliography`，`APA`/`MLA`/`Chicago`；`read --view sources`；`bibliography_cached` 警告）· **文字效果**（run 级 `set {shadow,glow,reflection,outline}`——Word 2010 `w14:` 效果） |
| **.xlsx** | 创建 · 类型化写入（数字/布尔/字符串/日期）· **公式求值并缓存结果** + 诚实警告 · 数字格式 · 合并单元格 · 表格/工作表 · 区域读取 · 按值/公式查询 · HTML 渲染 · 模板 · 校验 | **图表**（bar/line/pie，`add type:chart`）· PNG 渲染 · 实时预览 | **数据透视表**（rows/columns/filters + sum/average/count/min/max，`pivot[@name=X]` 寻址）· **条件格式**（cellIs/colorScale/dataBar/containsText）· **图片**（anchor 锚定，PNG/JPEG） | **大工作簿流式读取**（SAX 扫原始 XML——`read --view stats/text` 与单元格/区域 `get` 不加载 DOM；41 MB / 33 万行约 2 秒出 stats）· **散点图与面积图** · **命名区域**（`/name[@name=X]`，公式可用——`=SUM(SalesData)` 真求值）· **冻结窗格** · **自动筛选** · **打印设置**（方向/纸张/fitTo/打印区域） | **批量 2D 写入**（锚点式 `set /Sheet1/A2 values:[[…]]` 或精确区域式；公式随写随求值；>50k 单元格写空白 sheet 走 SAX 流式）· **行列操作**（插删自动重写公式引用、行高列宽、隐藏、`col[C]` 字母寻址）· **单元格批注**（增/读/删 + 作者）· **查找替换**（文本单元格；`inFormulas:true` 才进公式文本） | **csv 桥**（`create --from orders.csv`：RFC 4180、分隔符嗅探、类型化单元格——`007` 保持文本，>50k 单元格走流式；`read --view csv [--sheet] [--range]` 反向导出）· **数据验证**（list 下拉，字面值或引用区域；wholeNumber/decimal/date/textLength 规则 + 算子；错误样式）· **迷你图**（line/column/winLoss，颜色、markers）· **线程批注**（真 `xl/threadedComments` + 按 `/Sheet1/comment[@id=GUID]` 回复，旧客户端 note 退化）· **单元格超链接**（`https://…` 外链 + `#Sheet!A1` 内链，tooltip） | **就地流式写入**（`stream:true` 或任意 >20 MB 文件经 SAX writer 原地改写工作簿——50 MB+ 大表里写深单元格、批量写区域秒级完成；结果标 `streamed:true`）· **Excel 表 / ListObject**（`add type:table` 套区域——命名、内置样式、汇总行 sum/avg/…、结构化引用 `=SUM(Sales[Amount])` 求值；`/Sheet1/table[@name=X]`）· **大纲分组**（`add type:group` 套 `row[a]:row[b]` / `col[a]:col[b]` 跨段、`collapsed`、嵌套级别） | **审计**（`audit metrics.xlsx [--fix]`——公式错误 `#DIV/0!`/`#REF!`、合并数据单元格、图片无 alt、无文档标题；alt/标题可安全自动修复）· **命名单元格样式**（`add type:cellStyle` numberFormat/bold/fill/color/border 一次定义，`set {cellStyle:"X"}` 套区域，`read --view styles`，`/style[@name=X]`）· **文档属性**（`set /properties`；`read --view properties`）· **图片 alt 文本** | **diff**（`diff new.xlsx old.xlsx`——单元格改动、增删工作表、命名区域、表格归为 modified/added/removed）· **切片器**（`add type:slicer` 套表列或透视字段，原始 OpenXml 自研；`/Sheet1/slicer[i]`，`get`/`remove`/structure） | **convert**（`convert data.xlsx data.docx` / `.pptx` / `.csv` / `.md` / `.pdf`——经中立模型工作表→表格，xlsx↔csv 文本桥；图表/透视/图片汇入 `data.dropped` + `convert_lossy`） | **嵌入对象**（`add type:embed` 把任意文件锚定到工作表；`read --view embeds`；`extract` 按位导出；`/Sheet1/embed[i]`）· **统一属性结构**（`data.properties.{core,custom}`）· 公式 N/A（Excel 只有单元格公式，无 OMML 公式对象 → `add type:equation` 返回 `unsupported_feature` 并给出替代方案） | **图表扩展**（`doughnut`/`radar`/`bubble`/`stackedBar`/`percentStackedBar`/`stackedArea`/`combo`）· **iconSet 条件格式**（3/4/5 图标集，`reverse`/`showValue`） |
| **.pptx** | 创建（validator 零错误，PowerPoint/Keynote 可直接打开）· 增/删/重排幻灯片 · 定位文本形状（cm/EMU）· 稳定 shape id 的 query/get · **逐页 SVG 渲染** · 模板 · 校验 | 形状**填充/字体/颜色/对齐属性** · **母版/版式只读寻址** · 逐页 PNG 渲染 · 实时预览 | **幻灯片背景**（真 `p:bg` 纯色填充）· **演讲者备注**（`/slide[i]/notes` set/add/remove/get）· **图片**（PNG/JPEG，稳定 `shape[@id=N]` 路径） | **原生图表**（bar/line/pie，字面量数据缓存，`/slide[i]/chart[k]`）· **跨文档 `dataFrom`**（图表系列直接取自工作簿）· **切换动画**（fade/push/wipe + 时长）· **预设几何形状**（椭圆/三角/菱形/箭头/圆角矩形 + line 连接线，支持翻转）· **z 序**（`move` 到 front/back/forward/backward） | **可编辑图表数据**（新建图表嵌入真实工作簿——PowerPoint 右键「编辑数据」直接可用；旧图表 `set {embedData:true}` 一键改造）· **进入动画**（appear/fade/flyIn/wipe，方向与 click/with/after 触发，`/slide[i]/animation[k]`）· **幻灯片批注**（`add type:comment`，`/slide[i]/comment[@id=N]`）· **查找替换**（slide 作用域含演讲者备注） | **原生表格**（`add type:table` rows×cols、`headerRow`、light/medium/dark 三档样式、单元格 `mergeRight`/`mergeDown`、`/slide[i]/table[k]/tr[r]/tc[c]` 寻址、SVG 画出真网格）· **强调与退出动画**（pulse/grow/spin/colorPulse · fadeOut/flyOut/wipeOut，structure 视图按播放顺序列出）· **批注回复**（`add type:reply`——p15 线程，PowerPoint 2013+ 显示）· **SmartArt 读取**（`/slide[i]/smartart[k]` 嵌套节点树；编辑保持类型化 `unsupported_feature`） | **母版与版式编辑**（`set /master[m]` 背景 + 主题强调色、`add type:layout` 克隆版式、编辑母版/版式形状、克隆版式经 `add type:slide props:{layout:N}` 套用）· **幻灯片分节**（`/` 上 `add type:section`，`afterSlide` 区间；`read --view outline` 按节分组；抗重排）· **幻灯片尺寸/宽高比**（`set / {slideSize:"4:3"}` 或显式 `width`/`height`）· **动画时间轴重排**（`move /slide[i]/animation[2] before …`） | **审计**（`audit deck.pptx [--fix]`——图片无 alt、幻灯片无标题、画布外形状、过小字号（<12pt 警告 / <8pt 错误）、阅读顺序；alt/标题可安全自动修复）· **显式 alt 文本/标题**（shape `set {altText}`/`{altTitle}`；SVG 渲染出 `<title>`）· **文档属性**（`set /properties`） | **diff**（`diff new.pptx old.pptx`——重排幻灯片归为 moved，编辑形状/尺寸/分节/背景归为 modified）· **形状超链接/动作**（`set {hyperlink}` 外链 url / `#slide:N` 跳转 / `#first…#end` 放映动作、`{linkText}`；`get` 回读规范形式） | **convert**（`convert deck.pptx outline.docx` / `.xlsx` / `.md` / `.pdf`——经中立模型逐标题落每页内容；动画/切换/图表/SmartArt/嵌入汇入 `data.dropped` + `convert_lossy`） | **嵌入对象**（`add type:embed` 把任意文件作为 OLE 对象放到幻灯片；`read --view embeds`；`extract` 按位导出；`/slide[i]/embed[@id=N]`）· **公式**（`add type:equation`——与 docx *同一个* LaTeX → OMML 转换器，原生渲染在幻灯片文本框；`/slide[i]/shape[@id=N]/omath[k]`；`equation_partial` 警告；LaTeX 存档可回读）· **统一属性结构**（`data.properties.{core,custom}`） | **图表扩展**（`doughnut`/`radar`/`bubble`/`stackedBar`/`percentStackedBar`/`stackedArea`/`combo`）· **嵌入媒体**（`add type:media` 音视频，`src`/`poster` 经沙箱）· **形状效果**（`set {shadow,glow,reflection,outline}` → `a:effectLst`）· **切换扩展**（`split`/`reveal`/`cut`/`zoom`） |

M3 跨格式新增（功能第一）：

- **`render --to pdf`** —— docx/xlsx 打印为分页 PDF；pptx 整副 deck 出一个 PDF、**一页一张幻灯片**，复用与 PNG 相同的系统浏览器管线（无浏览器时返回类型化 `unsupported_feature` + workaround）。
- **跨文档 `dataFrom`** —— `{"op":"add","type":"chart","props":{"dataFrom":"metrics.xlsx!Sheet1/A1:B5"}}` 直接用工作簿活数据建 pptx 图表：首列 → 分类，表头行 → 系列名，其余列 → 系列值。必经沙箱解析，区域写错时返回 candidates，CLI 与 MCP 行为一致。
- **大小上限改为 opt-in** —— M2 的 50 MB 默认值取消；默认任意大小都能打开（超大 xlsx 读取走流式路径）。设 `AIOFFICE_MAX_FILE_MB` 可恢复硬性 `file_too_large` 上限；未设时 `doctor` 报告 `limits.maxFileMb: "unlimited"`。

M4 跨格式新增 —— **三格式共用一份查找替换契约**：

- `{"op":"replace","path":"<scope>","props":{"find","replace","regex?","matchCase?","wholeWord?"}}` 作用于任意容器路径；路径 `"/"` 表示整份文档（docx 正文 + 全部页眉页脚、每个 sheet、每页 slide 含备注），逐作用域结果聚合为一份 `{replacements, locations}`。
- 跨格式 run 拆分安全（docx/pptx 中被格式切碎的文本照样命中，重写保留首个受影响 run 的格式）；regex 为 .NET 语法 + 2 秒匹配预算（超时 → 类型化 `invalid_args`）；零命中是 `ok:true` + `find_no_match` warning——绝不是错误。
- CLI 糖：`aioffice edit report.docx --find 2025 --replace 2026 [--regex] [--match-case] [--whole-word]`；MCP `office_edit` 行为一致。docx 配 `--track` 时每处替换记成 `w:del`+`w:ins` 修订对，可后续 accept/reject。

M9 跨格式新增 —— **三格式共用一个 `convert` 动词**（参见上文 [格式互转](#格式互转)）：

- **内容中立模型** —— 跨格式对（`docx↔pptx`、`docx↔xlsx`、`pptx↔xlsx`，以及任意格式 `↔ md`）经一份共享的 `NeutralDoc` 模型（标题、带格式 run、列表、表格、图片）：源 handler 的 `ExportNeutral` 投影、目标 handler 的 `ImportNeutral` 从中重建新文件。`INeutralConvertible` 接口位于 Core（与 M7 `IAuditor`、M8 `IDiffer` 同模式），三 handler 各实现。
- **复用文本桥** —— `docx↔md`（M5 markdown 桥）与 `xlsx↔csv`（M5 csv 桥）经同一 `convert` 入口；`xlsx→md` / `pptx→md` 走中立模型 + 命令层 `NeutralMarkdown` 序列化器。
- **render 目标** —— `任意 → pdf/png/svg/html/txt` 打开源文件，路由到 render 层。
- **诚实有损** —— 转换是内容转移、跨格式本质有损：`data.dropped` 与单条 `convert_lossy` 警告列出未存活内容（动画、图表、SmartArt、精确样式、公式取值、markdown 颜色）。目标**新建**（覆盖 → `convert_overwrite`）；同扩展名 → `invalid_args`，未知目标 → `unsupported_feature`。
- **自省** —— `doctor`（与 `office_status`）新增 `capabilities` 块：动词数、MCP 工具数、支持格式、convert 源/目标、render 目标、audit 类别，一次调用即可让 agent 学到全表面。

M10 跨格式新增 —— **嵌入对象全覆盖 + 公式扩展到 pptx + pre-1.0 契约冻结**：

- **嵌入对象（`embed` / `extract`）** —— 把任意文件（报告里夹带的源 `.xlsx`、`.pdf`、`.zip`…）作为 OLE/包对象嵌入三种格式：`add type:embed src=…`（docx `/body`、xlsx `/Sheet1`、pptx `/slide[i]`），`read --view embeds` 列出 `{path,name,mediaType,size,container}`，新增 **`extract`** op 把负载**按位**导出（一个产出型 op，绝不改动源文档）。媒体类型从源文件嗅探；`src` 与 `extract` 目标都经沙箱解析。
- **共享 LaTeX → OMML 转换器** —— M6 公式引擎迁入 Core（`AIOffice.Core.Equations`）：一个纯 `System.Xml.Linq` 的 OMML 生产器，无 `DocumentFormat.OpenXml` 依赖。Word（载入 SDK 数学模型）与 **pptx**（新增——公式原生渲染在幻灯片文本框）共用这一个转换器，同一段 LaTeX 在两种格式渲染一致。xlsx 按设计 N/A（单元格公式，非公式对象）。未知命令仍以 `equation_partial` 警告降级。
- **统一属性结构** —— `read --view properties` 与 `get /properties` 现于三格式统一返回 `data.properties.{core,custom}`（此前 docx 嵌套、xlsx/pptx 扁平）——一次 pre-1.0 一致性修正。
- **契约冻结** —— 稳定的 v1.0 命令面记录于 **[CONTRACT.md](CONTRACT.md)**，并通过 `schema` 与 `doctor` capabilities 中的 `surfaceVersion`（`1.0`）声明：信封结构、冻结的错误码集合、寻址语法、退出码、op/view/工具词表。

长期能力对齐台账（对标业界最强 CLI）见 [docs/PARITY.md](docs/PARITY.md)——能力对齐是北极星，命令面刻意自研。稳定性承诺见 [CONTRACT.md](CONTRACT.md)。

## 架构

```
                 ┌─────────────────────────────────────────────┐
   agent/human → │  src/AIOffice.Cli   （aioffice，18 动词）    │
   MCP client  → │  src/AIOffice.Mcp   （stdio，17 工具，1:1）  │
                 ├─────────────────────────────────────────────┤
                 │  src/AIOffice.Render （经浏览器出 png/pdf）  │
                 │  src/AIOffice.Preview  （实时点选预览）      │
                 └──────────────┬──────────────────────────────┘
                                │ 信封 · 寻址 · 选择器
                 ┌──────────────▼──────────────────────────────┐
                 │  src/AIOffice.Core                          │
                 │  envelope/errors · DocPath · Selector       │
                 │  Workspace 沙箱 · SnapshotStore · rev       │
                 └───────┬──────────────┬──────────────┬───────┘
                         │              │              │
                 ┌───────▼─────┐ ┌──────▼───────┐ ┌────▼────────┐
                 │AIOffice.Word│ │AIOffice.Excel│ │AIOffice.Pptx│
                 │ OpenXml SDK │ │  ClosedXML   │ │ OpenXml SDK │
                 └───────┬─────┘ └──────┬───────┘ └────┬────────┘
                         ▼              ▼              ▼
                       .docx          .xlsx          .pptx    （无损 OOXML）
```

## 质量

本项目的起点之一，是研究过一个**零自动化测试**却日更发布的优秀 office CLI——AIOffice 选择完全相反的立场：

- **2331 个测试**横跨 7 个项目（Core 180 · Word 681 · Excel 619 · Pptx 709 · MCP 87 · Preview 24 · Render 31），每次提交全绿。
- **往返定律**：打开 → 不编辑直接保存，所有 zip part 必须字节级一致；已文档化的例外被逐一精确断言。
- **独立裁判**：每个变更测试后 OpenXmlValidator 必须报 0 错误——工具不给自己的作业打分。
- **CI 矩阵**：macOS 14 + Windows，warnings-as-errors 构建、金样脚本、单文件发布并冒烟。
- **发版自动化**：推送 `v*` tag 即构建、全量测试、发布 6 平台单文件二进制 + `SHA256SUMS`，并用 `scripts/release-notes-template.md` 渲染的说明创建 GitHub release（打 tag 前先改模板）——见 `.github/workflows/release.yml`。
- **真人核验**：`fixtures/manual-check/` 中的生成文件供真实 Office 打开验证。

## 路线图

- **M0** —— 以上全部；单文件发布；macOS + Windows CI。
- **M1（已交付，v0.2.0）** —— PNG 渲染（探测系统浏览器）· `preview_open`/`preview_selection`（实时预览 + 人工点选回读）· docx 页眉页脚 · pptx 母版/版式只读寻址 · xlsx 图表（bar/line/pie）。
- **M2（已交付，v0.3.0）** —— 修订（`--track`/`--author`，accept/reject）· 批注 · 样式管理 · 数据透视表 · 条件格式 · 图片（三格式）· pptx 背景与演讲者备注 · 大文件守卫（`file_too_large`，`AIOFFICE_MAX_FILE_MB`）。大文件**流式**没有按 M2 交付：它需要一轮专门的基准驱动打磨，M2 以尺寸守卫兜底——移入 M3。
- **M3（已交付，v0.4.0）** —— 功能第一：docx 列表/超链接/书签/脚注/页面设置/格式修订处理/批注回复 · xlsx 流式读取（SAX）/散点+面积图/命名区域/冻结窗格/自动筛选/打印设置 · pptx 原生图表/切换动画/预设几何/z 序 · 跨文档 `dataFrom`（xlsx 数据 → pptx 图表，CLI 与 MCP 同路径）· `render --to pdf`（docx/xlsx 分页；pptx 一页一片）· 大小上限改 opt-in（默认不限）。
- **M4（已交付，v0.5.0）** —— 三格式共用一份**查找替换**契约（跨 run 安全、regex 带超时、文档级 `"/"` 作用域、CLI `--find/--replace` 糖、docx 修订对）· docx **目录/水印/尾注/分节符** · xlsx **批量 2D 写入**（空白 sheet SAX 流式）/**行列操作**/**单元格批注** · pptx **可编辑图表数据**（嵌入工作簿，PowerPoint「编辑数据」可用，`embedData:true` 改造）/**进入动画**/**幻灯片批注** · tag 触发的**发版自动化**。
- **M5（已交付，v0.6.0）** —— **markdown/csv 桥**：`create --from`（`.md` → `.docx` 经 Markdig，`.csv` → `.xlsx` 类型化导入）+ `read --view markdown|csv` 导出 round-trip · docx **深表格**（合并/边框/底纹/columnWidths/headerRow）/**域**（PAGE/NUMPAGES/DATE/TITLE）/**首页与奇偶页眉页脚变体** · xlsx **数据验证**（下拉 + 规则）/**迷你图**/**线程批注 + 回复**/**单元格超链接** · pptx **原生表格**（合并 + 样式）/**强调与退出动画**/**批注回复**/**SmartArt 读取** · `IFormatHandler.CreateFrom` 导入钩子（增量式，默认 `unsupported_feature`）。
- **M6（已交付，v0.7.0）** —— **深水区攻坚**：docx **公式**（手写 LaTeX → Office Math 转换器，行内/显示、矩阵、部分降级警告、原始 LaTeX 存档可忠实回读）/**右到左与双向**（段/run/表）/**多栏分节**（含分栏符）· xlsx **就地流式写入**（经 SAX writer 原地改写 50 MB+ 大工作簿）/**Excel 表**（ListObject + 汇总 + 结构化引用）/**大纲分组**（行/列跨段）· pptx **母版与版式编辑**（背景、主题强调色、克隆版式）/**幻灯片分节**/**幻灯片尺寸与宽高比**/**动画时间轴重排** · 新增寻址形式 `/`（演示文稿根）、`/body/p[i]/omath[j]`、`/section[i]`、`row[a]:row[b]` 跨段、可编辑 `/master[1]/layout[i]`。
- **M7（已交付，v0.8.0）** —— **交付前先审计**：三格式共享 `audit` 动词 + `office_audit` MCP 工具（第 15 个工具）—— 无障碍 + 质量 lint，findings 是数据（`ok:true`，exit 0），`--fix` 只做安全自动修复（占位 alt 文本、表头行、文档/幻灯片标题、孤儿书签）并返回 `{fixed, remaining}` · docx **文档属性**（core + 类型化 custom）/**内容控件**（text/dropdown/date/checkbox，`read --view fields`）/图片 alt 文本 · xlsx **命名单元格样式**（`add type:cellStyle`，`read --view styles`）/文档属性/公式错误 + 合并数据单元格 + alt 文本审计 · pptx alt-text/标题/阅读顺序审计 + 显式 `altText`/`altTitle`/文档属性 · 新增寻址形式 `/properties`、`/sdt[@tag=X]`、`/style[@name=X]`；`office_read` view 枚举新增 `properties`/`fields`/`styles`。
- **M8（已交付，v0.9.0）** —— **对比与审阅**：三格式共享 `diff` 动词 + `office_diff` MCP 工具（第 16 个工具）—— 语义对比文档与基线（另一同格式文件 `other` *或*文档自身的某个编辑前快照 `--snapshot N`），返回已排序、确定性的 `{changes:[added/removed/modified/moved], summary, baseline}`（变更是数据，exit 0；LCS/内容哈希区分 moved；`--view summary` 裁剪为 path+kind）· docx **题注**（Figure/Table/Equation + SEQ）/**交叉引用**（REF/PAGEREF，`labelAndNumber`/`numberOnly`/`text`/`page`）· xlsx **切片器**（表列 / 透视字段，原始 OpenXml 自研）· pptx **形状超链接/动作**（url / `#slide:N` 跳转 / `#first…#end` 放映动作 / `linkText`）· 新增寻址形式 `/caption[@label=Figure][i]`、`/Sheet1/slicer[i]`。
- **M9（已交付，v0.10.0 —— pre-1.0 capstone）** —— **跨格式互转**：三格式共享 `convert` 动词 + `office_convert` MCP 工具（第 17 个工具）—— docx/xlsx/pptx ↔ 互转经一份内容中立模型（`NeutralDoc` + `INeutralConvertible`，与 `IAuditor`/`IDiffer` 同一 Core 接口模式），`docx↔md` 与 `xlsx↔csv` 走 M5 文本桥，任意格式 ↔ markdown 走 `NeutralMarkdown` 序列化器，`任意 → pdf/png/svg/html` 走 render 层；跨格式本质有损，故 `data.dropped` + `convert_lossy` 警告列出未存活内容 · **1.0 加固**：每个 CLI 动词都有 help 主题并出现在 `schema`（新增 `convert` help 主题含 `(src→dest)` 矩阵），`doctor`/`office_status` 新增 `capabilities` 自省块（动词与工具数、格式、convert 源/目标、render 目标、audit 类别）。
- **M10（已交付，v0.11.0 —— 1.0 前最后一个功能里程碑）** —— **文档智能 II**：三格式 **嵌入对象**（`add type:embed` 把任意文件作为 OLE/包对象，`read --view embeds`，新增 **`extract`** op 把负载按位导出；`/embed[i]`、`/Sheet1/embed[i]`、`/slide[i]/embed[@id=N]`；src + 目标均经沙箱）· **pptx 公式** 经迁入 Core 的**共享** LaTeX → OMML 转换器（`AIOffice.Core.Equations`，纯 `System.Xml.Linq` 生产器；Word 与 Pptx 共用；`/slide[i]/shape[@id=N]/omath[k]`；xlsx 按设计 N/A）· **1.0 契约准备**：`read --view properties` / `get /properties` 信封统一为三格式一致的 `data.properties.{core,custom}`，`schema` 与 `doctor` capabilities 声明稳定的 **`surfaceVersion`**（`1.0-rc`），冻结的面向 AI 契约写入 **[CONTRACT.md](CONTRACT.md)**。嵌入与公式都搭载在 `office_edit`/`office_read` 上——仍是 **17 个 MCP 工具**。
- **1.0.0（已交付 —— API 稳定性版本）** —— **稳定化，而非新功能**：`surfaceVersion` 升为 **`1.0`**；[CONTRACT.md](CONTRACT.md) 定稿为冻结的 v1.0 命令面，附**稳定性承诺**与**已知限制**两节；CLI 与 MCP 自省面在共享词表上锁定一致（`SchemaConsistencyTests` 守卫，漂移即红）；`read --view` 枚举 + `edit` op 列表在**两个**面上都暴露 `embeds` / `extract`；xlsx 读视图回显 `data.view`；xlsx convert 现经 `data.dropped` + `convert_lossy` 报告图表/透视/图片丢失（不再静默）；完整**警告码词表**集中、文档化并经 `schema` 的 `data.warningCodes` 暴露；寻址 help 文档补齐 M8–M10 的 `embed` / `omath` 形式。18 动词 / 17 MCP 工具，**1590 个测试**在 macOS + Windows 全绿。
- **1.1.0（已交付 —— 首个 1.0 之后功能版本，纯增量）** —— `surfaceVersion` **保持 `1.0`**（所有变更都在冻结的 1.0 契约线内、只增不删）：xlsx 与 pptx **图表扩展**（`doughnut`/`radar`/`bubble`/`stackedBar`/`percentStackedBar`/`stackedArea`/`combo`，在 `bar`/`line`/`pie`/`scatter`/`area` 之上）· xlsx **`iconSet` 条件格式**（3/4/5 图标集）· docx **引用与参考文献**（`add type:source`/`citation`/`bibliography`，`read --view sources`，`bibliography_cached` 警告）· pptx **嵌入媒体**（`add type:media` 音视频，`src`/`poster` 经沙箱）· docx run 与 pptx 形状 **文字/外观效果**（`shadow`/`glow`/`reflection`/`outline`）· pptx **新切换**（`split`/`reveal`/`cut`/`zoom`）。新增 `add` type、一个新 `read --view`（`sources`）、一个新警告码——全部增量，由 `SchemaConsistencyTests` 守卫。18 动词 / 17 MCP 工具不变，**1692 个测试**在 macOS + Windows 全绿。
- **1.2.0（已交付 —— 第二个 1.0 之后功能版本，纯增量）** —— `surfaceVersion` **保持 `1.0`**（op kinds 不变——`group`/`ungroup` 是 `add` **type** 值）：pptx **SmartArt 创建**（`add type:smartart`，`list`/`process`/`hierarchy`/`orgChart`/`cycle`）、**连接符**（`add type:connector`，straight/elbow/curved + 箭头）、**形状组合/分组**（`add type:group`/`ungroup`）· docx **图表目录**（`tableOfFigures` + `figures_cached`）、**索引**（`indexEntry`/`index` + `index_cached`）、**邮件合并域**（`mergeField`，由 `template` 与 `{{key}}` 同填）· xlsx **表单控件**（`formControl`：checkbox/optionButton/spinner/comboBox/listBox/button）、**轻量保护**（cell `locked` + sheet `protected` + workbook `protectStructure`）、**numberFormat 命名预设**（`accounting-usd`/`percent`/`date-iso`/…）。新增 `add` type 与 prop 键、两个新警告码——全部增量，由 `SchemaConsistencyTests` 与 token-budget 测试守卫。18 动词 / 17 MCP 工具不变，**1807 个测试**在 macOS + Windows 全绿。
- **1.5.0（已交付 —— 第五个 1.0 之后功能版本，纯增量）** —— `surfaceVersion` **保持 `1.0`**（op kinds 不变）；**填平现代标量函数缺口**：xlsx **标量求值**（`=XLOOKUP`/`IFS`/`SWITCH`/`LET`/`MAXIFS`/`MINIFS`/`AVERAGEIFS` 求值并缓存——不再 `formula_not_evaluated`；`=TEXTSPLIT` 溢出）、**方案管理器**（`add type:scenario` + `set {applyScenario}`，已重算）、**单变量求解**（`set {goalSeek:{targetCell,targetValue}}`，`goal_seek_no_solution` 警告）· docx **表格单元格公式**（单元格 `set {formula}` → 缓存 `w:fldSimple`，`table_formula_cached` 警告）、**构建基块**（`add type:buildingBlock`/`buildingBlockRef`）、**行号**（`set /section[i] {lineNumbers}`）· pptx **嵌入字体**（`/fonts` 上 `add type:font`，`embedAll`）、**动作按钮**（`add type:actionButton`）、**自定义版式**（`add type:layout {placeholders}` + 幻灯片 `layoutName`）。新增 `add` type（`scenario`/`buildingBlock`/`buildingBlockRef`/`font`/`actionButton`）、新 prop 键（`applyScenario`/`goalSeek`/`formula`/`lineNumbers`/`embedAll`/`placeholders`）、新寻址（`/Sheet1/scenario[@name=…]`、`/buildingBlock[@name=…]`、`/fonts`）、两个新警告码（`goal_seek_no_solution`、`table_formula_cached`）——全部增量且向后兼容，由 `SchemaConsistencyTests` 与 token-budget 测试守卫。18 动词 / 17 MCP 工具不变，**2125 个测试**在 macOS + Windows 全绿。
- **1.6.0（已交付 —— 分发与上手，能力零变化）** —— `surfaceVersion` **保持 `1.0`**；二进制与 1.5.0 字节级同面。新增二进制*周边*的打包与上手层（npm 包、Homebrew formula、一行安装脚本、`SKILL.md` + cookbook/install/MCP-setup/signing 文档），几乎全在 C#/.NET 源码树之外。18 动词 / 17 MCP 工具不变，**2125 个测试**全绿。
- **1.7.0（已交付 —— 第六个 1.0 之后功能版本，纯增量）** —— `surfaceVersion` **保持 `1.0`**（op kinds 不变；无新动词/工具）；**补齐打印就绪、相机工具、计算模式与公式深度**：xlsx **打印完整性**（`printTitleRows`/`printTitleCols`/`pageBreaks`/`fitToPage`/居中/网格/标题/`printHeader`/`printFooter`）、**链接图片/相机工具**（`add type:linkedPicture`，`linked_picture_static` 警告）、**计算设置**（`set /` `calculationMode`/`iterativeCalc`/…）· docx **首字下沉**（`set {dropCap,…}`）、**图片水印**（`add type:watermark {image}`）、**STYLEREF/SYMBOL/QUOTE 域**、**带编号显示公式**（`number:…`，`/equation[@num=…]`）· **更深的 LaTeX→OMML**（docx+pptx 共享：`\begin{aligned|gathered|cases}`、`\binom`、`\overbrace`/`\underbrace`、`\iint`/`\iiint`、更多重音/关系——`equation_partial` 警告更少）· pptx **备注/讲义母版**（`/notesMaster`、`/handoutMaster`）、**动画时间**（`repeat`/`rewind`/`autoReverse`）、**表格单元格对齐**（`valign`/边距/`textDirection`）。新增 `add` type（`linkedPicture`）、新 set 路径（`/notesMaster`、`/handoutMaster`）、新寻址（`/SheetN/linkedPicture[i]`、`/equation[@num=…]`）、新 prop 键、两个新警告码（`linked_picture_static`、`equation_numbers_cached`）——全部增量且向后兼容，由 `SchemaConsistencyTests` 与 token-budget 测试守卫（新 `office_help` 主题 `print-setup`/`masters`）。18 动词 / 17 MCP 工具不变，**2331 个测试**在 macOS + Windows 全绿。

### 1.0 之后（候选想法，未承诺）

1.0 是 API 稳定线；以下均为**增量**，受契约「只增不删」承诺约束：

- **插件机制** —— 运行时发现的外部格式 handler，第三方无需 fork 即可加格式。
- **更高 convert 保真** —— 更细的 `data.dropped` 粒度；可选地把更多样式带过中立模型。
- **xlsx 打磨** —— sheetView RTL、现代线程批注细化。
- **pptx 动画深化** —— 更多预设、多效果链 / motion path。

## 设计声明

AIOffice 的命令面与现有 office CLI **刻意不兼容**：能力对齐是目标，但动词、参数、寻址语法全部为 agent 从零设计——更小、更可预测、可自省。一套心智模型，两种接入方式（CLI 与 MCP）。

## 文档

**上手：** [docs/INSTALL.md](docs/INSTALL.md) —— 四种安装方式 · [docs/MCP-SETUP.md](docs/MCP-SETUP.md) —— 接入 Claude / Cursor / 任意 MCP 宿主 · [SKILL.md](SKILL.md) —— AI 上手指南 · [docs/COOKBOOK.md](docs/COOKBOOK.md) —— 10 个复制即用配方 · [docs/SIGNING.md](docs/SIGNING.md) —— 代码签名 / 公证路线图

**参考：** [docs/DESIGN.md](docs/DESIGN.md) —— 架构与命令面规范 · [docs/MCP.md](docs/MCP.md) —— MCP 工具规范 · [docs/PARITY.md](docs/PARITY.md) —— 能力对齐台账 · [CONTRACT.md](CONTRACT.md) —— 冻结的 v1.0 AI 命令面 · `aioffice help <topic>` —— 内置渐进式文档 · [SMOKE_REPORT.md](SMOKE_REPORT.md) —— 端到端真实验证记录

## License

[Apache-2.0](LICENSE)。见 [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md)——全部依赖均为宽松许可（MIT：DocumentFormat.OpenXml、ClosedXML、ModelContextProtocol C# SDK；BSD-2-Clause：Markdig），无捆绑第三方二进制。
