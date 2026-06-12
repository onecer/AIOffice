# AIOffice

**English** | [简体中文](README.zh-CN.md)

[![CI](https://github.com/onecer/AIOffice/actions/workflows/ci.yml/badge.svg)](https://github.com/onecer/AIOffice/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
![.NET](https://img.shields.io/badge/.NET-10-512BD4)
![platforms](https://img.shields.io/badge/platforms-macOS%20%7C%20Windows%20%7C%20Linux-lightgrey)

**An AI-native CLI + MCP server for real Office files.** AIOffice lets AI agents create, query, edit, render, preview and validate `.docx` / `.xlsx` / `.pptx` the way they call functions: one command in, exactly one JSON envelope out.

100% self-built on pure C#/.NET — direct lossless OOXML via DocumentFormat.OpenXml + ClosedXML. **One ~36 MB single-file binary. No Microsoft Office, no runtime dependencies, no wrapped third-party engines.**

```bash
aioffice create report.docx --title "Q3 Report"
aioffice edit   report.docx --set '/body/p[1]' text="Revenue grew 12%"
aioffice read   report.docx --view outline
aioffice mcp    # the same 14 capabilities, as MCP tools over stdio
```

## Show, don't tell

Every file below was created, edited, validated and screenshotted by `aioffice` alone — no Office installed, no templates, no manual touch-ups. The exact script is in the fold below; [deck-1.svg](assets/demo/deck-1.svg) is the same title slide as SVG, where every shape carries a `data-aio-path` attribute pointing back to its canonical document path.

| | |
|---|---|
| ![deck.pptx slide 1, rendered to PNG by aioffice](assets/demo/deck-1.png)<br><sub>`aioffice render deck.pptx --to png --scope '/slide[1]' -o deck-1.png`</sub> | ![deck.pptx slide 2, stat cards built from positioned shapes](assets/demo/deck-2.png)<br><sub>`aioffice query deck.pptx 'shape:contains("14")'` → `/slide[2]/shape[@id=9]`</sub> |
| ![report.docx with running header, headings and table](assets/demo/report.png)<br><sub>`aioffice get report.docx '/header[1]/p[1]'` → `"text": "AIOffice Demo"`</sub> | ![metrics.xlsx with evaluated formulas and number formats](assets/demo/metrics.png)<br><sub>`aioffice get metrics.xlsx /Sheet1/B7` → `"cachedValue": 286900`</sub> |

<details>
<summary>The full script (every command verbatim, including the render → look → fix loop)</summary>

```bash
# in an empty working directory, with aioffice on PATH

# ---- deck.pptx — 3 slides, dark background + accent shapes ----
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

# ---- report.docx — running header, Heading1/2, intro, table ----
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
aioffice read report.docx --view outline        # orient: headings + canonical paths
aioffice edit report.docx --remove '/body/p[1]' # drop the empty paragraph create left behind
aioffice validate report.docx
aioffice get report.docx '/header[1]/p[1]'      # → "text": "AIOffice Demo"

# ---- metrics.xlsx — sales table, SUM/AVERAGE, formats, bar chart ----
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
aioffice get metrics.xlsx '/Sheet1/chart[1]'    # → bar chart "Revenue by month", A1:B6 @ E2
aioffice validate metrics.xlsx

# ---- render: the "look" step ----
aioffice render deck.pptx    --to png --scope '/slide[1]' -o deck-1.png
aioffice render deck.pptx    --to png --scope '/slide[2]' -o deck-2.png
aioffice render deck.pptx    --to svg --scope '/slide[1]' -o deck-1.svg
aioffice render report.docx  --to png -o report.png
aioffice render metrics.xlsx --to png -o metrics.png

# ---- look → fix: the slide-2 card labels overflowed their cards ----
aioffice query deck.pptx 'shape:contains("formats")'   # find the label → /slide[2]/shape[6]
aioffice edit deck.pptx --ops '[
  {"op":"set","path":"/slide[2]/shape[6]","props":{"text":"file formats — one binary"}},
  {"op":"set","path":"/slide[2]/shape[9]","props":{"text":"MCP tools, 1:1 with the CLI"}},
  {"op":"set","path":"/slide[2]/shape[12]","props":{"text":"Office installs required"}}]'
aioffice validate deck.pptx
aioffice render deck.pptx --to png --scope '/slide[2]' -o deck-2.png
```

The PNGs were then downscaled to ≤900 px wide (plain `sips`/Pillow) before landing in `assets/demo/` — pixels untouched otherwise. The xlsx PNG shows the sheet as the HTML renderer draws it today: cells, formats and cached formula results; the bar chart lives in the file (see `get '/Sheet1/chart[1]'`) and shows up when Excel opens it.

</details>

## Why AI-native?

Most office libraries are built for programmers. Most office CLIs are built for humans. AIOffice is built for **agents** — every design decision optimizes the loop an LLM actually runs: *act → observe → recover → verify*.

| Feature | What it means for an agent |
|---|---|
| **One JSON envelope per command** | `{ok, data, error, meta}` on stdout, always. Nothing to scrape, nothing to guess. |
| **Errors that teach** | Every error carries a mandatory `suggestion`. `invalid_path` even ships `candidates` — the nearest valid paths, computed server-side. One failed call, zero wasted recovery turns. |
| **Stable addressing** | `/body/p[3]`, `/Sheet1/A1:C10`, `/slide[2]/shape[3]` — 1-based, canonical, returned by `query` so edits never aim at guessed indices. |
| **Atomic batch edits** | `edit --ops '[...]'` applies all-or-nothing, supports `--dry-run`, and guards with optimistic concurrency (`--expect-rev`) — a stale file fails *before* any write, with `stale_address`. |
| **Automatic undo** | Every mutation snapshots the pre-image into a 20-deep ring. `snapshot restore` is one call — and is itself undoable. |
| **Write-time formula evaluation** | Excel formulas are computed and **cached into the file** (`=SUM(A1:A2)` → reopen shows `42` instantly). Functions the engine can't evaluate produce an explicit `formula_not_evaluated` warning — never a silently stale value. |
| **render → look → fix** | Render docx/xlsx to HTML, pptx slides to SVG — and any of them to **PNG** via the system browser, no Office installed. The agent *sees* what it made and fixes it. |
| **Human-in-the-loop preview** | `preview open` serves a live view on localhost; rendered nodes carry `data-aio-path`, so a human **click** comes back to the agent as a canonical path via `preview selection`. |
| **Sandboxed by default** | All file args resolve inside a workspace allowlist (`--workspace`, symlink-escape checked). Out-of-bounds access → `sandbox_denied`, exit 4. |
| **Introspectable surface** | `aioffice schema` returns the entire command surface as machine-readable JSON. Agents read the spec instead of hallucinating it. |
| **CLI = MCP, one mental model** | 15 CLI verbs and 14 MCP tools map 1:1. Learn it once, drive it from a shell or over stdio. |

### Errors that teach — real output

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

### Honest formulas — real output

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

## Quickstart

```bash
# Build (requires .NET 10 SDK)
dotnet build AIOffice.sln

# Run from source
alias aioffice='dotnet run --project src/AIOffice.Cli --'

# Or publish a single-file binary (~36 MB, self-contained)
dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release \
  -p:PublishSingleFile=true --self-contained
# rids: osx-arm64 | osx-x64 | win-x64 | win-arm64 | linux-x64 | linux-arm64

aioffice doctor          # environment / handlers / workspace diagnosis
```

### A 60-second tour

```bash
# Word
aioffice create report.docx --title "Q3 Report"
aioffice edit   report.docx --ops '[
  {"op":"add","path":"/body","type":"p","position":"inside",
   "props":{"text":"Quarterly results","style":"Heading1"}}]'
aioffice query  report.docx 'p[style=Heading1]'      # → canonical paths + snippets
aioffice render report.docx --to html -o report.html

# Excel — formulas evaluated at write time
aioffice create data.xlsx
aioffice edit   data.xlsx --ops '[
  {"op":"set","path":"/Sheet1/A1","props":{"value":21}},
  {"op":"set","path":"/Sheet1/A2","props":{"value":21}},
  {"op":"set","path":"/Sheet1/A3","props":{"value":"=SUM(A1:A2)"}}]'
aioffice get data.xlsx /Sheet1/A3                    # → "cachedValue": 42

# PowerPoint — see what you made
aioffice create deck.pptx          # starts with one blank slide
aioffice edit   deck.pptx --ops '[{"op":"add","path":"/slide[1]","type":"slide",
                                   "position":"after","props":{"title":"Hello"}}]'
aioffice edit   deck.pptx --add '/slide[1]' --type shape text="AIOffice" x=2cm y=3cm w=10cm h=2cm
aioffice render deck.pptx --to svg --scope '/slide[1]' -o slide1.svg

# Safety nets
aioffice snapshot list report.docx                   # automatic pre-edit snapshots
aioffice snapshot restore report.docx 1              # one-call rollback
aioffice validate report.docx                        # OOXML validation + lint
```

## MCP (for Claude and other agents)

```bash
aioffice mcp     # stdio MCP server — 14 tools, 1:1 with the CLI verbs
```

Claude Desktop / Claude Code config:

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

| MCP tool | CLI verb | MCP tool | CLI verb |
|---|---|---|---|
| `office_create` | create | `office_validate` | validate |
| `office_read` | read | `office_template` | template |
| `office_query` | query | `file_snapshot` | snapshot |
| `office_get` | get | `office_status` | doctor |
| `office_edit` | edit | `office_help` | help |
| `office_render` | render | `office_schema` | schema |
| `preview_open` | preview open | `preview_selection` | preview selection |

`preview_open` / `preview_selection` (live preview with human click-to-select) registered in M1 (v0.2.0). Total tool-schema budget is capped at 3,500 tokens — enforced by a test.

## Command surface (v0.2.0)

| Verb | Summary |
|---|---|
| `create <file> [--kind] [--title]` | New document; kind inferred from extension |
| `read <file> [--view outline\|text\|stats\|structure]` | Cheap inspection projections, paged |
| `query <file> <selector>` | CSS-like selectors → canonical paths (`p[style=Heading1]`, `cell[value>100]`, `shape:contains('Q3')`) |
| `get <file> <path>` | One node + its properties |
| `edit <file> --ops <json\|@file>` | **Atomic** batch set/add/remove/move · `--dry-run` · `--expect-rev` · sugar `--set/--add/--remove` |
| `render <file> [--to html\|svg\|text\|png] [--scope]` | The *look* step — html for docx/xlsx, svg per pptx slide, **png** via the system browser |
| `validate <file>` | OOXML validation + lint with fix suggestions |
| `template <file> --data <json\|@file>` | `{{key}}` merge across docx/xlsx/pptx (split-run safe) |
| `snapshot <list\|restore> <file> [n]` | Pre-edit snapshot ring (20) |
| `preview <open\|selection\|close> <file> [--port N]` | Live localhost preview; human clicks → canonical paths via `selection` |
| `doctor` | Environment / runtime / handler diagnosis |
| `schema [verb]` | Machine-readable JSON of the whole surface |
| `help [topic]` | addressing · selectors · properties-docx/xlsx/pptx · errors |
| `mcp` | stdio MCP server |
| `version` | Version info |

Global flags: `--json` (default when not a TTY) · `--pretty` · `--workspace <dir>` (sandbox root, default cwd, or `AIOFFICE_WORKSPACE`) · `--quiet`.
Exit codes: `0` ok · `2` user error · `3` internal/format error · `4` sandbox_denied · `5` unsupported_feature.

**Addressing** (1-based): `/body/p[3]` · `/body/table[1]/tr[2]/tc[1]` · `/Sheet1/A1:C10` · `/'Q3 Data'/B2` · `/slide[2]/shape[3]`.

## What works today (M0 + M1 + M2)

| Format | M0 (v0.1.0) | + M1 (v0.2.0) | + M2 (v0.3.0) |
|---|---|---|---|
| **.docx** | create · paragraphs/headings/styles · tables · text & formatting edits (bold/italic/color/alignment/size) · query/get · outline/text/stats/structure views · HTML render · `{{key}}` templates · validate | **headers/footers** (create + edit, `/header[1]/p[1]`) · PNG render · live preview | **tracked changes** (`--track --author`, `read --view revisions`, accept/reject by `/revision[@id=N]` or scope) · **comments** (add/read/remove, `/comment[@id=N]`) · **custom styles** (`/styles` add, `/style[@id=X]` set/get/remove) · **images** (PNG/JPEG, sandboxed `src`, aspect-keeping) |
| **.xlsx** | create · typed cell writes (number/bool/string/date) · **formula evaluation with cached values** + honest warnings · number formats · merge · tables/sheets · range reads · query by value/formula · HTML render · templates · validate | **charts** (bar/line/pie, `add type:chart`) · PNG render · live preview | **pivot tables** (rows/columns/filters + sum/average/count/min/max values, `pivot[@name=X]`) · **conditional formatting** (cellIs/colorScale/dataBar/containsText) · **images** (anchored, PNG/JPEG) |
| **.pptx** | create (validator-clean, opens in PowerPoint/Keynote) · add/reorder/remove slides · positioned text shapes (cm/EMU) · query/get with stable shape ids · **SVG render per slide** · templates · validate | shape **fill/font/color/align props** · **master/layout read addressing** · PNG render per slide · live preview | **slide backgrounds** (real `p:bg` solid fill) · **speaker notes** (`/slide[i]/notes` set/add/remove/get) · **images** (PNG/JPEG, stable `shape[@id=N]` paths) |

Cross-format in M2: a **file-size guard** — opening anything over 50 MB (env `AIOFFICE_MAX_FILE_MB`) fails fast with `file_too_large` and an actionable suggestion; `doctor` reports `limits.maxFileMb`.

The long-term capability ledger (vs. the strongest CLI in the field) lives in [docs/PARITY.md](docs/PARITY.md) — capability parity is the north star; the command surface is deliberately our own.

## Architecture

```
                 ┌─────────────────────────────────────────────┐
   agent/human → │  src/AIOffice.Cli   (aioffice, 15 verbs)    │
   MCP client  → │  src/AIOffice.Mcp   (stdio, 14 tools, 1:1)  │
                 ├─────────────────────────────────────────────┤
                 │  src/AIOffice.Render   (png via browser)    │
                 │  src/AIOffice.Preview  (live click-select)  │
                 └──────────────┬──────────────────────────────┘
                                │ envelope · addressing · selectors
                 ┌──────────────▼──────────────────────────────┐
                 │  src/AIOffice.Core                          │
                 │  envelope/errors · DocPath · Selector       │
                 │  Workspace sandbox · SnapshotStore · rev    │
                 └───────┬──────────────┬──────────────┬───────┘
                         │              │              │
                 ┌───────▼─────┐ ┌──────▼───────┐ ┌────▼────────┐
                 │AIOffice.Word│ │AIOffice.Excel│ │AIOffice.Pptx│
                 │ OpenXml SDK │ │  ClosedXML   │ │ OpenXml SDK │
                 └───────┬─────┘ └──────┬───────┘ └────┬────────┘
                         ▼              ▼              ▼
                       .docx          .xlsx          .pptx    (lossless OOXML)
```

## Quality

Born from studying an excellent office CLI that ships **zero automated tests** — AIOffice takes the opposite stance:

- **383 tests** across 7 projects (Core 104 · Word 73 · Pptx 72 · Excel 59 · MCP 30 · Preview 24 · Render 21), green on every commit.
- **Round-trip law**: open → save with no edits must leave every zip part byte-identical; documented exceptions are asserted exactly.
- **Independent oracle**: OpenXmlValidator must report 0 errors after every mutating test — the tool never grades its own homework.
- **CI matrix**: macOS 14 + Windows, builds with warnings-as-errors, runs golden scripts, publishes and smokes the single-file binary.
- **Human check**: generated files in `fixtures/manual-check/` for opening in real Office.

## Roadmap

- **M0** — everything above; single-file publish; CI on macOS + Windows.
- **M1 (shipped, v0.2.0)** — PNG render (system browser detection) · `preview_open`/`preview_selection` (live preview, human click-to-select) · docx headers/footers · pptx master/layout read addressing · xlsx charts (bar/line/pie).
- **M2 (shipped, v0.3.0)** — tracked changes (`--track`/`--author`, accept/reject) · comments · style management · pivot tables · conditional formatting · images (all three formats) · pptx backgrounds & speaker notes · file-size guard (`file_too_large`, `AIOFFICE_MAX_FILE_MB`). Large-file *streaming* did **not** ship: it needs a dedicated benchmark-driven pass; M2 ships a size guard instead — moved to M3.
- **M3** — large-file streaming (benchmarked) · cross-document workflows (xlsx data → pptx charts) · batch pipelines · capability plugins · full parity ledger burn-down.

## Design statement

AIOffice's surface is **deliberately incompatible** with existing office CLIs: capability parity is the goal, but the verbs, flags and addressing grammar are designed from scratch for agents — smaller, more predictable, introspectable. One mental model, two transports (CLI & MCP).

## Docs

[docs/DESIGN.md](docs/DESIGN.md) — architecture & surface spec · [docs/MCP.md](docs/MCP.md) — MCP tool spec · [docs/PARITY.md](docs/PARITY.md) — capability ledger · `aioffice help <topic>` — built-in progressive docs · [SMOKE_REPORT.md](SMOKE_REPORT.md) — real end-to-end verification log

## License

[Apache-2.0](LICENSE). See [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) — all dependencies are MIT (DocumentFormat.OpenXml, ClosedXML, ModelContextProtocol C# SDK); no bundled third-party binaries.
