# AIOffice

**English** | [简体中文](README.zh-CN.md)

[![CI](https://github.com/onecer/AIOffice/actions/workflows/ci.yml/badge.svg)](https://github.com/onecer/AIOffice/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
![.NET](https://img.shields.io/badge/.NET-10-512BD4)
![platforms](https://img.shields.io/badge/platforms-macOS%20%7C%20Windows%20%7C%20Linux-lightgrey)

**An AI-native CLI + MCP server for real Office files.** AIOffice lets AI agents create, query, edit, render and validate `.docx` / `.xlsx` / `.pptx` the way they call functions: one command in, exactly one JSON envelope out.

100% self-built on pure C#/.NET — direct lossless OOXML via DocumentFormat.OpenXml + ClosedXML. **One ~36 MB single-file binary. No Microsoft Office, no runtime dependencies, no wrapped third-party engines.**

```bash
aioffice create report.docx --title "Q3 Report"
aioffice edit   report.docx --set '/body/p[1]' text="Revenue grew 12%"
aioffice read   report.docx --view outline
aioffice mcp    # the same 12 capabilities, as MCP tools over stdio
```

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
| **render → look → fix** | Render docx/xlsx to HTML and pptx slides to SVG, no Office installed — the agent can *see* what it made and fix it. |
| **Sandboxed by default** | All file args resolve inside a workspace allowlist (`--workspace`, symlink-escape checked). Out-of-bounds access → `sandbox_denied`, exit 4. |
| **Introspectable surface** | `aioffice schema` returns the entire command surface as machine-readable JSON. Agents read the spec instead of hallucinating it. |
| **CLI = MCP, one mental model** | 14 CLI verbs and 12 MCP tools map 1:1. Learn it once, drive it from a shell or over stdio. |

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
aioffice mcp     # stdio MCP server — 12 tools, 1:1 with the CLI verbs
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

`preview_open` / `preview_selection` (live preview with human click-to-select) are reserved for M1. Total tool-schema budget is capped at 3,500 tokens — enforced by a test.

## Command surface (v0)

| Verb | Summary |
|---|---|
| `create <file> [--kind] [--title]` | New document; kind inferred from extension |
| `read <file> [--view outline\|text\|stats\|structure]` | Cheap inspection projections, paged |
| `query <file> <selector>` | CSS-like selectors → canonical paths (`p[style=Heading1]`, `cell[value>100]`, `shape:contains('Q3')`) |
| `get <file> <path>` | One node + its properties |
| `edit <file> --ops <json\|@file>` | **Atomic** batch set/add/remove/move · `--dry-run` · `--expect-rev` · sugar `--set/--add/--remove` |
| `render <file> [--to html\|svg\|text] [--scope]` | The *look* step; png lands in M1 |
| `validate <file>` | OOXML validation + lint with fix suggestions |
| `template <file> --data <json\|@file>` | `{{key}}` merge across docx/xlsx/pptx (split-run safe) |
| `snapshot <list\|restore> <file> [n]` | Pre-edit snapshot ring (20) |
| `doctor` | Environment / runtime / handler diagnosis |
| `schema [verb]` | Machine-readable JSON of the whole surface |
| `help [topic]` | addressing · selectors · properties-docx/xlsx/pptx · errors |
| `mcp` | stdio MCP server |
| `version` | Version info |

Global flags: `--json` (default when not a TTY) · `--pretty` · `--workspace <dir>` (sandbox root, default cwd, or `AIOFFICE_WORKSPACE`) · `--quiet`.
Exit codes: `0` ok · `2` user error · `3` internal/format error · `4` sandbox_denied · `5` unsupported_feature.

**Addressing** (1-based): `/body/p[3]` · `/body/table[1]/tr[2]/tc[1]` · `/Sheet1/A1:C10` · `/'Q3 Data'/B2` · `/slide[2]/shape[3]`.

## What works today (M0)

| Format | Capabilities (v0.1.0) |
|---|---|
| **.docx** | create · paragraphs/headings/styles · tables · text & formatting edits (bold/italic/color/alignment/size) · query/get · outline/text/stats/structure views · HTML render · `{{key}}` templates · validate |
| **.xlsx** | create · typed cell writes (number/bool/string/date) · **formula evaluation with cached values** + honest warnings · number formats · merge · tables/sheets · range reads · query by value/formula · HTML render · templates · validate |
| **.pptx** | create (validator-clean, opens in PowerPoint/Keynote) · add/reorder/remove slides · positioned text shapes (cm/EMU) · query/get with stable shape ids · **SVG render per slide** · templates · validate |

The long-term capability ledger (vs. the strongest CLI in the field) lives in [docs/PARITY.md](docs/PARITY.md) — capability parity is the north star; the command surface is deliberately our own.

## Architecture

```
                 ┌─────────────────────────────────────────────┐
   agent/human → │  src/AIOffice.Cli   (aioffice, 14 verbs)    │
   MCP client  → │  src/AIOffice.Mcp   (stdio, 12 tools, 1:1)  │
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

- **275 tests** across 5 projects (Core 104 · Word 52 · Pptx 48 · Excel 41 · MCP 30), green on every commit.
- **Round-trip law**: open → save with no edits must leave every zip part byte-identical; documented exceptions are asserted exactly.
- **Independent oracle**: OpenXmlValidator must report 0 errors after every mutating test — the tool never grades its own homework.
- **CI matrix**: macOS 14 + Windows, builds with warnings-as-errors, runs golden scripts, publishes and smokes the single-file binary.
- **Human check**: generated files in `fixtures/manual-check/` for opening in real Office.

## Roadmap

- **M0 (now)** — everything above; single-file publish; CI on macOS + Windows.
- **M1** — PNG render (system browser detection) · `preview_open`/`preview_selection` (live preview, human click-to-select) · docx headers/footers · pptx master/layout addressing · xlsx charts.
- **M2** — tracked changes · comments · style management · pivot tables · conditional formatting · large-file streaming.
- **M3** — cross-document workflows (xlsx data → pptx charts) · batch pipelines · capability plugins · full parity ledger burn-down.

## Design statement

AIOffice's surface is **deliberately incompatible** with existing office CLIs: capability parity is the goal, but the verbs, flags and addressing grammar are designed from scratch for agents — smaller, more predictable, introspectable. One mental model, two transports (CLI & MCP).

## Docs

[docs/DESIGN.md](docs/DESIGN.md) — architecture & surface spec · [docs/MCP.md](docs/MCP.md) — MCP tool spec · [docs/PARITY.md](docs/PARITY.md) — capability ledger · `aioffice help <topic>` — built-in progressive docs · [SMOKE_REPORT.md](SMOKE_REPORT.md) — real end-to-end verification log

## License

[Apache-2.0](LICENSE). See [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) — all dependencies are MIT (DocumentFormat.OpenXml, ClosedXML, ModelContextProtocol C# SDK); no bundled third-party binaries.
