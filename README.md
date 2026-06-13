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
aioffice diff   report.docx old.docx        # semantic compare: a sorted change list
aioffice convert report.docx deck.pptx      # cross-format: a slide per heading, bullets below
aioffice mcp    # the same 17 capabilities, as MCP tools over stdio
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
| **CLI = MCP, one mental model** | 17 CLI verbs and 17 MCP tools map 1:1. Learn it once, drive it from a shell or over stdio. |

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

## Markdown in, Office out

Agents think in markdown and csv. M5 (v0.6.0) makes those the front door — every command below is verbatim from the release smoke run:

```bash
# markdown -> real docx (headings, nested lists, pipe tables, links, bold, code)
aioffice create report.docx --from notes.md
aioffice read   report.docx --view outline     # headings + canonical paths
aioffice read   report.docx --view markdown    # …and back out as GFM — structure round-trips
aioffice validate report.docx                  # "valid": true, 0 errors

# csv -> typed xlsx (quoted commas survive, dates type, "007" STAYS text)
aioffice create orders.xlsx --from orders.csv
aioffice get    orders.xlsx /Sheet1/A2         # → "value": "007", "type": "text"
aioffice read   orders.xlsx --view csv         # one sheet back out as RFC 4180 csv
```

Mismatched pairs fail fast with the matrix in the suggestion: `.md/.markdown → .docx, .csv/.tsv → .xlsx`. Same wiring over MCP: `office_create {file, from}` and `office_read {view:"markdown"|"csv"}`.

## LaTeX in, equations out

M6 (v0.7.0) adds a hand-rolled LaTeX → OOXML Math converter (no LaTeX dependency). You write LaTeX; the document gets **real Office Math** that Word renders as an equation. The commands below are verbatim from the release smoke run:

```bash
# create a doc and add an inline equation + a display block (the quadratic formula) + a matrix
aioffice create report.docx --title "M6 Report"
aioffice edit report.docx --ops '[{"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"E = mc^2"}}]'
aioffice edit report.docx --ops '[{"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"\\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}","display":true}}]'
aioffice edit report.docx --ops '[{"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"\\begin{pmatrix}1&0\\\\0&1\\end{pmatrix}","display":true}}]'

aioffice read report.docx --view text     # → "M6 Report$E = mc^2$", "$$\frac{-b \pm \sqrt{b^2-4ac}}{2a}$$", …
aioffice get  report.docx /body/p[1]/omath[1]   # → "latex": "E = mc^2", "display": false
aioffice validate report.docx              # "valid": true, 0 errors

# unknown commands degrade — the file still validates, you just get a warning
aioffice edit report.docx --ops '[{"op":"add","path":"/body/p[4]","type":"equation","props":{"latex":"\\foobar{x} + \\alpha"}}]'
# → ok:true + meta.warnings: [{ "code": "equation_partial", "message": "…\\foobar… appear literally…" }]
```

The original LaTeX is stored on each equation (as an `mc:Ignorable` vendor attribute), so `get` returns your source verbatim and round-trips stay byte-faithful. The supported subset — fractions, radicals, super/subscripts, big operators, matrices (`pmatrix`/`bmatrix`/…), `\left…\right`, accents, `\text`, Greek letters and operators — is documented at `aioffice help equations`. Render it and the math is real:

```bash
aioffice render report.docx --to png -o quadratic.png   # the formula below was produced exactly this way
```

![The quadratic formula, rendered from LaTeX through real Office Math](assets/demo/equation-quadratic.png)

## Audit before you ship

`aioffice audit` is accessibility + quality lint for Office files. Findings are
**data**, never errors — the command exits `0` even when it surfaces
error-severity findings, exactly like `validate`. `--fix` applies only the
**safe, non-destructive** autofixes and re-audits, so you see what is left.

```bash
# A deliberately-bad report: image with no alt text, H1 then H3, a header-less
# table, an empty heading, no document title.
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

# --fix applies the safe autofixes (alt text, table header, doc title), then
# re-audits. The heading-skip and empty-heading findings are report-only and
# stay in `remaining` for you to fix by hand.
$ aioffice audit report.docx --fix
{ "ok": true, "data": {
  "findings": [
    { "code": "a11y_heading_skip",     "path": "/body/p[3]", "autofixable": false },
    { "code": "quality_empty_heading", "path": "/body/p[4]", "autofixable": false }
  ],
  "summary": { "errors": 0, "warnings": 2, "infos": 0 },
  "fixed": 3,
  "remaining": ["a11y_heading_skip#/body/p[3]", "quality_empty_heading#/body/p[4]"] } }

# Still valid OOXML afterwards.
$ aioffice validate report.docx
{ "ok": true, "data": { "valid": true, "count": 0, "issues": [] } }
```

Same verb, same findings, same `--fix` semantics for **.xlsx** (formula errors
like `#DIV/0!`, merged data cells, missing alt text/title) and **.pptx** (picture
alt text, slide titles, off-canvas shapes, tiny fonts, reading order). Scope with
`--category accessibility|quality|all` and `--severity error|warning|info`; the
full code list and which codes are autofixable is in `aioffice help audit`.
Over MCP it is `office_audit {file, category?, severity?, fix?}` — the 15th tool.

## Diff & review

`aioffice diff` **semantically** compares a document against a baseline and
returns a sorted, deterministic change list — `added` / `removed` / `modified`
(with `before`/`after`) / `moved`. Changes are *data*: exit 0 even when the two
documents differ a lot. The baseline is either **another same-format file** or
**one of the file's own pre-edit snapshots**.

```console
# two files — what changed between this draft and the last one
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

# against a snapshot — "what did I just do?" (every edit auto-snapshots first)
$ aioffice edit report.docx --set '/body/p[2]' text='Revised'
$ aioffice diff report.docx --snapshot 1
#   → the single change that edit made, before/after, deterministic

# terse review — counts plus one path+kind line per change
$ aioffice diff metrics.xlsx baseline.xlsx --view summary
```

The change list is always sorted by `(path, kind)`, so the same two documents
diff identically on every platform and every run. A cross-format baseline (a
`.docx` against a `.xlsx`) is `invalid_args` naming the mismatch. Over MCP it is
`office_diff {file, other? | snapshot?, view?}` — the 16th tool.

## Convert between formats

`convert` moves a document from one format to another. It routes by the
`(source, destination)` extension pair: same-family text bridges reuse the
existing code (`docx↔md`, `xlsx↔csv`), cross-format pairs go through a
**content-neutral model** (headings, paragraphs, formatted runs, lists, tables,
images), and `any → pdf/png/svg/html` routes to the render layer. Conversion is
inherently lossy across formats — whatever the target cannot represent is named
honestly in `data.dropped` and a single `convert_lossy` warning.

```bash
# docx → pptx — a slide per heading, its bullets as the slide body, tables as native pptx tables
$ aioffice convert report.docx deck.pptx
# → {"from":"docx","to":"pptx","blocksWritten":10,"dropped":[],"written":".../deck.pptx"}

# xlsx → docx — a heading + table per sheet; formulas cross as their cached values
$ aioffice convert data.xlsx data.docx
# → a docx whose Sheet1/Totals tables hold the display values (e.g. =SUM(...) → 30)

# anything → markdown — via the neutral model, so xlsx and pptx reach md too
$ aioffice convert report.docx report.md      # headings, bullets and a pipe table
$ aioffice convert report.md roundtrip.docx   # and back — the outline round-trips

# anything → PDF — paged print via the system browser
$ aioffice convert report.docx report.pdf

# guardrails
$ aioffice convert a.docx b.docx              # invalid_args: "use edit" (same format)
$ aioffice convert a.docx a.xyz               # unsupported_feature, naming the targets
```

A deck converted from a chart-bearing workbook reports the loss instead of hiding
it: `meta.warnings: [{ "code": "convert_lossy", "message": "Some content did not
survive the conversion: charts (…the chart is not converted)." }]`. Over MCP it is
`office_convert {src, dest}` — the 17th tool. See `aioffice help convert` for the
full `(src → dest)` matrix.

## MCP (for Claude and other agents)

```bash
aioffice mcp     # stdio MCP server — 17 tools, 1:1 with the CLI verbs
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
| `office_create` | create | `office_template` | template |
| `office_read` | read | `file_snapshot` | snapshot |
| `office_query` | query | `office_status` | doctor |
| `office_get` | get | `office_help` | help |
| `office_edit` | edit | `office_schema` | schema |
| `office_render` | render | `preview_open` | preview open |
| `office_validate` | validate | `preview_selection` | preview selection |
| `office_audit` | audit | `office_diff` | diff |
| `office_convert` | convert | | |

`preview_open` / `preview_selection` (live preview with human click-to-select) registered in M1 (v0.2.0); `office_audit` (accessibility + quality lint) is the 15th tool, added in M7 (v0.8.0); `office_diff` (semantic compare) is the 16th tool, added in M8 (v0.9.0); `office_convert` (cross-format conversion) is the 17th tool, added in M9 (v0.10.0). Total tool-schema budget is capped at 3,500 tokens — enforced by a test, and still under it with `office_convert` added.

## Command surface (v0.10.0)

| Verb | Summary |
|---|---|
| `create <file> [--from notes.md\|data.csv] [--kind] [--title]` | New document (kind inferred from extension) — or **import**: `.md` → `.docx`, `.csv` → `.xlsx` |
| `read <file> [--view outline\|text\|stats\|structure\|markdown\|csv]` | Cheap inspection projections, paged; `markdown` exports a docx body as GFM, `csv` exports one xlsx sheet (`--sheet`, `--range`) |
| `query <file> <selector>` | CSS-like selectors → canonical paths (`p[style=Heading1]`, `cell[value>100]`, `shape:contains('Q3')`) |
| `get <file> <path>` | One node + its properties |
| `edit <file> --ops <json\|@file>` | **Atomic** batch set/add/remove/move/replace · `--dry-run` · `--expect-rev` · sugar `--set/--add/--remove` · **document-wide find/replace sugar** `--find X --replace Y [--regex] [--match-case] [--whole-word]` (docx body+headers+footers, every sheet, every slide incl. notes; aggregate `{replacements, locations}`; with `--track` on docx every hit becomes a revision pair) |
| `render <file> [--to html\|svg\|text\|png\|pdf] [--scope]` | The *look* step — html for docx/xlsx, svg per pptx slide, **png/pdf** via the system browser (pptx pdf: whole deck, one page per slide) |
| `validate <file>` | OOXML validation + lint with fix suggestions |
| `audit <file> [--category accessibility\|quality\|all] [--severity error\|warning\|info] [--fix]` | **Accessibility + quality lint** — findings are data (`ok:true`, exit 0); `--fix` applies only safe autofixes (alt text, table header, doc/slide title, orphan bookmark) and reports `{fixed, remaining}` |
| `diff <file> [<other>] [--snapshot N] [--view summary\|detailed]` | **Semantic compare** against a same-format file *or* a snapshot of the file itself — a sorted, deterministic `{changes:[added/removed/modified/moved], summary, baseline}` (changes are data, exit 0); `--view summary` trims to counts + path+kind |
| `convert <src> <dest>` | **Cross-format conversion** — docx/xlsx/pptx ↔ each other (content-neutral model), docx↔md, xlsx↔csv, any→pdf/png/svg/html. Lossy across formats: `data.dropped` + a `convert_lossy` warning name what didn't survive; same ext → `invalid_args`, unknown target → `unsupported_feature` |
| `template <file> --data <json\|@file>` | `{{key}}` merge across docx/xlsx/pptx (split-run safe) |
| `snapshot <list\|restore> <file> [n]` | Pre-edit snapshot ring (20) |
| `preview <open\|selection\|close> <file> [--port N]` | Live localhost preview; human clicks → canonical paths via `selection` |
| `doctor` | Environment / runtime / handler diagnosis |
| `schema [verb]` | Machine-readable JSON of the whole surface |
| `help [topic]` | addressing · selectors · properties-docx/xlsx/pptx · errors · **equations** · **rtl** · **sections** · **audit** · **diff** · **convert** |
| `mcp` | stdio MCP server |
| `version` | Version info |

Global flags: `--json` (default when not a TTY) · `--pretty` · `--workspace <dir>` (sandbox root, default cwd, or `AIOFFICE_WORKSPACE`) · `--quiet`.
Exit codes: `0` ok · `2` user error · `3` internal/format error · `4` sandbox_denied · `5` unsupported_feature.

**Addressing** (1-based): `/body/p[3]` · `/body/table[1]/tr[2]/tc[1]` · `/body/p[3]/omath[1]` · `/Sheet1/A1:C10` · `/Sheet1/table[@name=Sales]` · `/Sheet1/row[2]:row[6]` · `/'Q3 Data'/B2` · `/slide[2]/shape[3]` · `/section[1]` · `/master[1]/layout[2]` · `/` (pptx slide size + sections) · **M7**: `/properties` (core + custom doc properties) · `/sdt[@tag=status]` (docx content controls) · `/style[@name=Currency-Red]` (xlsx named cell styles) · **M8**: `/caption[@label=Figure][1]` (docx captions) · `/Sheet1/slicer[1]` (xlsx slicers).

## What works today (M0 + M1 + M2 + M3 + M4 + M5 + M6 + M7 + M8 + M9)

| Format | M0 (v0.1.0) | + M1 (v0.2.0) | + M2 (v0.3.0) | + M3 (v0.4.0) | + M4 (v0.5.0) | + M5 (v0.6.0) | + M6 (v0.7.0) | + M7 (v0.8.0) | + M8 (v0.9.0) |
|---|---|---|---|---|---|---|---|---|---|
| **.docx** | create · paragraphs/headings/styles · tables · text & formatting edits (bold/italic/color/alignment/size) · query/get · outline/text/stats/structure views · HTML render · `{{key}}` templates · validate | **headers/footers** (create + edit, `/header[1]/p[1]`) · PNG render · live preview | **tracked changes** (`--track --author`, `read --view revisions`, accept/reject by `/revision[@id=N]` or scope) · **comments** (add/read/remove, `/comment[@id=N]`) · **custom styles** (`/styles` add, `/style[@id=X]` set/get/remove) · **images** (PNG/JPEG, sandboxed `src`, aspect-keeping) | **lists** (numbered/bulleted, nested levels, restart; `1.`/`•` markers in text view, real `<ol>/<ul>` in HTML) · **hyperlinks** (external url + bookmark anchors) · **bookmarks** · **footnotes** · **page setup** (`/section[1]`: pageSize/orientation/margins) · **formatting-revision accept/reject** (w:rPrChange/w:pPrChange) · **threaded comment replies** (`add type:reply` on `/comment[@id=N]`) | **table of contents** (`add type:toc`, levels/title/position; `/toc[1]` get with entryCount) · **text watermarks** (`add type:watermark`, every header, auto-creates one) · **endnotes** (`/endnote[@id=N]`) · **section breaks** (`add type:sectionBreak`, per-section page setup — portrait & landscape in one file) · **find/replace** (split-run safe; `--track` makes every hit a w:del+w:ins pair) | **markdown bridge** (`create --from notes.md` imports GFM — headings/lists/tables/links/code; `read --view markdown` exports it back, structure round-trips) · **deep tables** (`mergeRight`/`mergeDown`, borders all/outer/none, shading, `headerRow` repeat, `columnWidths`, valign; real colspan/rowspan in HTML) · **fields** (PAGE/NUMPAGES/DATE/TITLE + `leadingText` — 'Page X of Y' footers) · **firstPage/even header+footer variants** (`/header[firstPage]`, auto `w:titlePg`/`w:evenAndOddHeaders`) | **equations** (`add type:equation` — LaTeX → real Office Math; inline `/body/p[i]/omath[j]` or display block; `$…$`/`$$…$$` in text view; unknown commands degrade with an `equation_partial` warning; original LaTeX stored for faithful read-back) · **right-to-left / bidi** (`rtl` on paragraph `w:bidi` / run `w:rtl` / table `w:bidiVisual`) · **multi-column sections** (`columns`/`columnGap` on `/section[1]` + `add type:columnBreak`) | **audit** (`audit report.docx [--fix]` — no-alt-text, heading skips, table headers, low contrast, doc title, empty/broken headings, broken links, orphan bookmarks; safe autofixes for alt/header/title/bookmark) · **document properties** (`set /properties` core + typed custom; `read --view properties`) · **content controls** (`add type:contentControl` text/dropdown/date/checkbox; `/sdt[@tag=X]`; `read --view fields`) · **image alt text** (`set {alt}`) | **diff** (`diff new.docx old.docx` — paragraph/table/header/property changes as added/removed/modified/moved, deterministic) · **captions** (`add type:caption` Figure/Table/Equation + SEQ; `/caption[@label=Figure][i]`) · **cross-references** (`add type:crossRef` REF/PAGEREF; `labelAndNumber`/`numberOnly`/`text`/`page`) |
| **.xlsx** | create · typed cell writes (number/bool/string/date) · **formula evaluation with cached values** + honest warnings · number formats · merge · tables/sheets · range reads · query by value/formula · HTML render · templates · validate | **charts** (bar/line/pie, `add type:chart`) · PNG render · live preview | **pivot tables** (rows/columns/filters + sum/average/count/min/max values, `pivot[@name=X]`) · **conditional formatting** (cellIs/colorScale/dataBar/containsText) · **images** (anchored, PNG/JPEG) | **streaming reads** for huge workbooks (SAX over raw XML — `read --view stats/text` and cell/range `get` without loading the DOM; a 41 MB / 330k-row book answers stats in ~2 s) · **scatter & area charts** · **defined names** (`/name[@name=X]`, live in formulas — `=SUM(SalesData)` evaluates) · **freeze panes** · **autoFilter** · **print setup** (orientation/paperSize/fitTo/printArea) | **bulk 2D writes** (anchor `set /Sheet1/A2 values:[[…]]` or exact range; formulas ride along and evaluate; >50k cells into a blank sheet stream via SAX) · **rows & columns** (insert/delete with formula rewriting, height/width, hidden, `col[C]` letter addressing) · **cell notes** (add/read/remove + author) · **find/replace** (text cells; `inFormulas:true` opts into formula text) | **csv bridge** (`create --from orders.csv`: RFC 4180, sniffed delimiter, typed cells — `007` stays text, >50k cells stream; `read --view csv [--sheet] [--range]` exports back) · **data validation** (list dropdowns from values or a source range; wholeNumber/decimal/date/textLength rules with operators; error styles) · **sparklines** (line/column/winLoss, color, markers) · **threaded comments** (real `xl/threadedComments` + replies by `/Sheet1/comment[@id=GUID]`, legacy-note fallback) · **cell hyperlinks** (`https://…` + internal `#Sheet!A1`, tooltips) | **in-place streaming writes** (`stream:true` or any >20 MB file rewrites the workbook through the SAX writer — set deep cells & bulk-write ranges in a 50 MB+ book in seconds; `streamed:true` in the result) · **Excel Tables / ListObjects** (`add type:table` over a range — name, built-in style, totals row with sum/avg/…, structured references `=SUM(Sales[Amount])` evaluate; `/Sheet1/table[@name=X]`) · **outline grouping** (`add type:group` over a `row[a]:row[b]` / `col[a]:col[b]` span, `collapsed`, nested levels) | **audit** (`audit metrics.xlsx [--fix]` — formula errors `#DIV/0!`/`#REF!`, merged data cells, image alt text, doc title; safe autofixes for alt/title) · **named cell styles** (`add type:cellStyle` numberFormat/bold/fill/color/border once, `set {cellStyle:"X"}` to a range, `read --view styles`, `/style[@name=X]`) · **document properties** (`set /properties`; `read --view properties`) · **image alt text** | **diff** (`diff new.xlsx old.xlsx` — changed cells, added/removed sheets, defined names, tables as modified/added/removed) · **slicers** (`add type:slicer` on a table column or pivot field, raw-OpenXml authored; `/Sheet1/slicer[i]`, `get`/`remove`/structure) |
| **.pptx** | create (validator-clean, opens in PowerPoint/Keynote) · add/reorder/remove slides · positioned text shapes (cm/EMU) · query/get with stable shape ids · **SVG render per slide** · templates · validate | shape **fill/font/color/align props** · **master/layout read addressing** · PNG render per slide · live preview | **slide backgrounds** (real `p:bg` solid fill) · **speaker notes** (`/slide[i]/notes` set/add/remove/get) · **images** (PNG/JPEG, stable `shape[@id=N]` paths) | **native charts** (bar/line/pie with literal data caches, `/slide[i]/chart[k]`) · **`dataFrom` cross-doc data** (chart series pulled straight from a workbook) · **slide transitions** (fade/push/wipe + duration) · **preset geometries** (ellipse/triangle/diamond/arrow/roundRect + line connectors, flips) · **z-order** (`move` to front/back/forward/backward) | **editable chart data** (new charts embed a real workbook — right-click → *Edit Data* works in PowerPoint; retrofit old charts with `set {embedData:true}`) · **entrance animations** (appear/fade/flyIn/wipe, directions, click/with/after triggers, `/slide[i]/animation[k]`) · **slide comments** (`add type:comment`, `/slide[i]/comment[@id=N]`) · **find/replace** (slide scope includes speaker notes) | **native tables** (`add type:table` rows×cols, `headerRow`, light/medium/dark looks, cell `mergeRight`/`mergeDown`, `/slide[i]/table[k]/tr[r]/tc[c]` paths, real grid in SVG) · **emphasis & exit animations** (pulse/grow/spin/colorPulse · fadeOut/flyOut/wipeOut, ordered in structure view) · **comment replies** (`add type:reply` — p15 threads PowerPoint 2013+ shows) · **SmartArt read** (`/slide[i]/smartart[k]` nested node trees; editing stays a typed `unsupported_feature`) | **master & layout editing** (`set /master[m]` background + theme accents, `add type:layout` clones a layout, edit master/layout shapes, use a cloned layout via `add type:slide props:{layout:N}`) · **slide sections** (`add type:section` on `/`, `afterSlide` ranges; `read --view outline` groups slides; survive reordering) · **slide size / aspect ratio** (`set / {slideSize:"4:3"}` or explicit `width`/`height`) · **animation timeline reorder** (`move /slide[i]/animation[2] before …`) | **audit** (`audit deck.pptx [--fix]` — picture alt text, slide titles, off-canvas shapes, tiny fonts (<12pt warn / <8pt error), reading order; safe autofixes for alt/title) · **explicit alt text / title** (`set {altText}`/`{altTitle}` on a shape; SVG renders a `<title>`) · **document properties** (`set /properties`) | **diff** (`diff new.pptx old.pptx` — reordered slides as moved, edited shapes/size/sections/background as modified) · **shape hyperlinks / actions** (`set {hyperlink}` url / `#slide:N` jump / `#first…#end` show actions, `{linkText}`; `get` reads back the canonical form) |

Cross-format in M3 (功能第一 — features first):

- **`render --to pdf`** — docx/xlsx print to paged PDF, a pptx deck becomes one PDF with **one page per slide**, via the same system-browser pipeline as PNG (no browser → typed `unsupported_feature` with the workaround).
- **Cross-document `dataFrom`** — `{"op":"add","type":"chart","props":{"dataFrom":"metrics.xlsx!Sheet1/A1:B5"}}` builds a pptx chart from live workbook data: first column → categories, header row → series names, remaining columns → series. Sandbox-resolved, candidates on a wrong range, identical over CLI and MCP.
- **Size cap is now opt-in** — the M2 50 MB default is gone; files of any size open by default (huge xlsx reads go through the streaming path). Set `AIOFFICE_MAX_FILE_MB` to restore a hard `file_too_large` cap; `doctor` reports `limits.maxFileMb: "unlimited"` otherwise.

Cross-format in M4 — **one find/replace contract for all three formats**:

- `{"op":"replace","path":"<scope>","props":{"find","replace","regex?","matchCase?","wholeWord?"}}` works on any container path; path `"/"` means the whole document (docx body + every header/footer, every sheet, every slide including notes) and the per-scope results are aggregated into one `{replacements, locations}` pair.
- Matches split across formatting runs are found (docx/pptx); the rewritten text keeps the first affected run's formatting. Regex is .NET syntax with a 2 s match budget (timeout → typed `invalid_args`). Zero hits is `ok:true` plus a `find_no_match` warning — never an error.
- CLI sugar: `aioffice edit report.docx --find 2025 --replace 2026 [--regex] [--match-case] [--whole-word]`; identical over MCP `office_edit`. On docx, `--track` records every replacement as a `w:del`+`w:ins` revision pair for later accept/reject.

Cross-format in M9 — **one `convert` verb for all format transfers** (see [Convert between formats](#convert-between-formats) above):

- **Content-neutral model** — cross-format pairs (`docx↔pptx`, `docx↔xlsx`, `pptx↔xlsx`, and any format `↔ md`) go through a shared `NeutralDoc` model (headings, formatted runs, lists, tables, images): the source handler's `ExportNeutral` projects to it, the destination handler's `ImportNeutral` rebuilds a fresh file from it. The `INeutralConvertible` interface lives in Core (the same pattern as M7 `IAuditor` and M8 `IDiffer`); each of the three handlers implements it.
- **Text bridges reused** — `docx↔md` (the M5 markdown bridge) and `xlsx↔csv` (the M5 csv bridge) are reached through the same `convert` entrypoint, while `xlsx→md` / `pptx→md` ride the neutral model + a command-layer `NeutralMarkdown` serializer.
- **Render targets** — `any → pdf/png/svg/html/txt` opens the source and routes to the render layer.
- **Honest lossiness** — conversion is content-transfer and inherently lossy: `data.dropped` and a single `convert_lossy` warning name what didn't survive (animations, charts, SmartArt, exact styling, formulas-as-values, markdown colors). The destination is created fresh (overwrite → `convert_overwrite`); same-extension is `invalid_args`, an unknown target is `unsupported_feature`.
- **Introspection** — `doctor` (and `office_status`) now carry a `capabilities` block: verb count, MCP tool count, supported formats, convert sources/targets, render targets and audit categories, so an agent can learn the whole surface in one call.

The long-term capability ledger (vs. the strongest CLI in the field) lives in [docs/PARITY.md](docs/PARITY.md) — capability parity is the north star; the command surface is deliberately our own.

## Architecture

```
                 ┌─────────────────────────────────────────────┐
   agent/human → │  src/AIOffice.Cli   (aioffice, 17 verbs)    │
   MCP client  → │  src/AIOffice.Mcp   (stdio, 17 tools, 1:1)  │
                 ├─────────────────────────────────────────────┤
                 │  src/AIOffice.Render  (png/pdf via browser) │
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

- **1501 tests** across 7 projects (Core 124 · Word 459 · Excel 374 · Pptx 410 · MCP 79 · Preview 24 · Render 31), green on every commit.
- **Round-trip law**: open → save with no edits must leave every zip part byte-identical; documented exceptions are asserted exactly.
- **Independent oracle**: OpenXmlValidator must report 0 errors after every mutating test — the tool never grades its own homework.
- **CI matrix**: macOS 14 + Windows, builds with warnings-as-errors, runs golden scripts, publishes and smokes the single-file binary.
- **Release automation**: pushing a `v*` tag builds, tests, publishes all 6 single-file binaries + `SHA256SUMS`, and creates the GitHub release with notes rendered from `scripts/release-notes-template.md` (edit it before tagging) — see `.github/workflows/release.yml`.
- **Human check**: generated files in `fixtures/manual-check/` for opening in real Office.

## Roadmap

- **M0** — everything above; single-file publish; CI on macOS + Windows.
- **M1 (shipped, v0.2.0)** — PNG render (system browser detection) · `preview_open`/`preview_selection` (live preview, human click-to-select) · docx headers/footers · pptx master/layout read addressing · xlsx charts (bar/line/pie).
- **M2 (shipped, v0.3.0)** — tracked changes (`--track`/`--author`, accept/reject) · comments · style management · pivot tables · conditional formatting · images (all three formats) · pptx backgrounds & speaker notes · file-size guard (`file_too_large`, `AIOFFICE_MAX_FILE_MB`). Large-file *streaming* did **not** ship: it needs a dedicated benchmark-driven pass; M2 ships a size guard instead — moved to M3.
- **M3 (shipped, v0.4.0)** — 功能第一: docx lists/links/bookmarks/footnotes/page-setup/format-revision-resolve/comment-replies · xlsx streaming reads (SAX)/scatter+area charts/defined names/freeze/autoFilter/print setup · pptx native charts/transitions/preset geometries/z-order · cross-doc `dataFrom` (xlsx data → pptx charts, CLI & MCP) · `render --to pdf` (paged docx/xlsx; pptx one page per slide) · size cap flipped to opt-in (default unlimited).
- **M4 (shipped, v0.5.0)** — one **find/replace** contract for all three formats (split-run safe, regex with timeout, document-wide `"/"` scope, CLI `--find/--replace` sugar, tracked revision pairs on docx) · docx **TOC / watermarks / endnotes / section breaks** · xlsx **bulk 2D writes** (SAX streaming into blank sheets) / **row & column ops** / **cell notes** · pptx **editable chart data** (embedded workbooks, Edit-Data in PowerPoint, `embedData:true` retrofit) / **entrance animations** / **slide comments** · tag-driven **release automation**.
- **M5 (shipped, v0.6.0)** — the **markdown/csv bridge**: `create --from` (`.md` → `.docx` via Markdig, `.csv` → `.xlsx` typed import) + `read --view markdown|csv` exports that round-trip · docx **deep tables** (merges/borders/shading/columnWidths/headerRow) / **fields** (PAGE/NUMPAGES/DATE/TITLE) / **firstPage & even header/footer variants** · xlsx **data validation** (dropdowns + rules) / **sparklines** / **threaded comments + replies** / **cell hyperlinks** · pptx **native tables** (merges + looks) / **emphasis & exit animations** / **comment replies** / **SmartArt read** · `IFormatHandler.CreateFrom` import hook (additive, default `unsupported_feature`).
- **M6 (shipped, v0.7.0)** — the **deep-water pass**: docx **equations** (a hand-rolled LaTeX → Office Math converter, inline/display, matrices, partial-degrade warnings, LaTeX stored for faithful read-back) / **right-to-left & bidi** (paragraph/run/table) / **multi-column sections** (+ column breaks) · xlsx **in-place streaming writes** (rewrite a 50 MB+ workbook through the SAX writer) / **Excel Tables** (ListObjects + totals + structured references) / **outline grouping** (row/col spans) · pptx **master & layout editing** (backgrounds, theme accents, cloned layouts) / **slide sections** / **slide size & aspect ratio** / **animation timeline reorder** · new addressing forms `/` (presentation root), `/body/p[i]/omath[j]`, `/section[i]`, `row[a]:row[b]` spans, editable `/master[1]/layout[i]`.
- **M7 (shipped, v0.8.0)** — **audit before you ship**: a shared `audit` verb + `office_audit` MCP tool (the 15th tool) across all three formats — accessibility + quality lint where findings are *data* (`ok:true`, exit 0), with safe-only `--fix` (placeholder alt text, table header rows, doc/slide titles, orphan bookmarks) reporting `{fixed, remaining}` · docx **document properties** (core + typed custom) / **content controls** (text/dropdown/date/checkbox, `read --view fields`) / image alt text · xlsx **named cell styles** (`add type:cellStyle`, `read --view styles`) / document properties / formula-error + merged-data-cell + alt-text audit · pptx alt-text/title/reading-order audit + explicit `altText`/`altTitle` / document properties · new addressing forms `/properties`, `/sdt[@tag=X]`, `/style[@name=X]`; `office_read` view enum gains `properties`/`fields`/`styles`.
- **M8 (shipped, v0.9.0)** — **diff & review**: a shared `diff` verb + `office_diff` MCP tool (the 16th tool) across all three formats — semantic compare against another same-format file *or* one of the document's own pre-edit snapshots, returning a sorted, deterministic `{changes:[added/removed/modified/moved], summary, baseline}` (changes are *data*, exit 0; LCS/content-hash distinguishes moves; `--view summary` trims to path+kind) · docx **captions** (Figure/Table/Equation + SEQ) / **cross-references** (REF/PAGEREF, `labelAndNumber`/`numberOnly`/`text`/`page`) · xlsx **slicers** (table-column / pivot-field, authored on raw OpenXml) · pptx **shape hyperlinks / actions** (url / `#slide:N` jump / `#first…#end` show actions / `linkText`) · new addressing forms `/caption[@label=Figure][i]`, `/Sheet1/slicer[i]`.
- **M9 (shipped, v0.10.0 — the pre-1.0 capstone)** — **cross-format conversion**: a shared `convert` verb + `office_convert` MCP tool (the 17th tool) — docx/xlsx/pptx ↔ each other through a content-neutral model (`NeutralDoc` + `INeutralConvertible`, the same Core-interface pattern as `IAuditor`/`IDiffer`), `docx↔md` and `xlsx↔csv` via the M5 text bridges, any format ↔ markdown via a `NeutralMarkdown` serializer, and `any → pdf/png/svg/html` via the render layer; lossy by nature, so `data.dropped` + a `convert_lossy` warning name what didn't survive · **1.0 hardening**: every CLI verb has a help topic and appears in `schema` (new `convert` help topic with the `(src→dest)` matrix), and `doctor`/`office_status` gain a `capabilities` introspection block (verb + tool counts, formats, convert sources/targets, render targets, audit categories).

### Toward 1.0

The capstone shipped; what remains before a 1.0 API-stability commitment:

- **pptx/xlsx equations** — OMML in slides and cells, through the shared `AIOffice.Word.Equations` converter (today equations are docx-only).
- **OLE objects** — embedded spreadsheets/objects across formats.
- **Plugin mechanism** — external format handlers discovered at runtime, so third parties can add formats without forking.
- **API-stability commitment** — freeze the envelope, addressing grammar and verb surface under a versioned compatibility promise.

## Design statement

AIOffice's surface is **deliberately incompatible** with existing office CLIs: capability parity is the goal, but the verbs, flags and addressing grammar are designed from scratch for agents — smaller, more predictable, introspectable. One mental model, two transports (CLI & MCP).

## Docs

[docs/DESIGN.md](docs/DESIGN.md) — architecture & surface spec · [docs/MCP.md](docs/MCP.md) — MCP tool spec · [docs/PARITY.md](docs/PARITY.md) — capability ledger · `aioffice help <topic>` — built-in progressive docs · [SMOKE_REPORT.md](SMOKE_REPORT.md) — real end-to-end verification log

## License

[Apache-2.0](LICENSE). See [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) — all dependencies are permissive (MIT: DocumentFormat.OpenXml, ClosedXML, ModelContextProtocol C# SDK; BSD-2-Clause: Markdig); no bundled third-party binaries.
