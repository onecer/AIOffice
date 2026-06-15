---
name: aioffice
description: >-
  Create, read, and edit REAL Office files (.docx, .xlsx, .pptx) from the command line —
  no Microsoft Office, no cloud, no templates. Use AIOffice whenever a task involves a
  Word document, Excel workbook, or PowerPoint deck as input or output: generating reports,
  filling spreadsheets with evaluated formulas, building slide decks, mail-merging letters,
  auditing for accessibility, diffing drafts, or converting between formats. A single
  self-contained binary that speaks one JSON envelope per command and is built to be
  driven by an AI agent.
license: See repository LICENSE.
---

# AIOffice

AIOffice is an AI-native CLI (and matching MCP server) for real OOXML files. One self-contained
binary, `aioffice`, reads and writes `.docx`, `.xlsx`, and `.pptx` with no Office install, no
network, and no headless-Office engine — OOXML in, OOXML out. It is designed to be driven by a
model: every command prints exactly **one JSON envelope**, errors **teach you how to fix them**,
and elements have **stable 1-based addresses** you can read, query, and edit.

## When to use it

Use AIOffice when the deliverable (or the source) is a genuine Office file:

- Generate a report, memo, or letter as a real `.docx`.
- Fill a spreadsheet with data and **formulas that are evaluated and cached** at write time.
- Build a slide deck with positioned shapes, tables, and charts.
- Mail-merge personalized documents from a JSON dataset.
- Audit a document for accessibility/quality and auto-fix it.
- Diff a draft against a baseline, or convert between formats.

Do **not** reach for it when a plain `.txt`/`.md`/`.csv` is the real deliverable (though
AIOffice *can* import markdown→docx and csv→xlsx, and export back).

## Install

Download the single binary for your platform from the latest release and put it on your PATH.

```bash
# macOS (Apple Silicon). Pick the asset matching your platform:
#   darwin+arm64 -> aioffice-mac-arm64     linux+x64    -> aioffice-linux-x64
#   darwin+x64   -> aioffice-mac-x64        linux+arm64  -> aioffice-linux-arm64
#   win+x64      -> aioffice-win-x64.exe     win+arm64    -> aioffice-win-arm64.exe
mkdir -p ~/.aioffice/bin
curl -L -o ~/.aioffice/bin/aioffice \
  https://github.com/onecer/AIOffice/releases/latest/download/aioffice-mac-arm64
chmod +x ~/.aioffice/bin/aioffice
export PATH="$HOME/.aioffice/bin:$PATH"   # add to your shell profile
aioffice version
# {"ok":true,"data":{"name":"aioffice","version":"…","runtime":".NET …"},"meta":{…}}
```

On macOS you may need to clear the quarantine flag once: `xattr -d com.apple.quarantine ~/.aioffice/bin/aioffice`.
Verify your download against the release's `SHA256SUMS`. There are no other dependencies — the
binary is self-contained (~36 MB).

### Register as an MCP server

The same surface is available over stdio as 17 MCP tools (`office_create`, `office_read`,
`office_edit`, … — 1:1 with the CLI verbs). Run it with `aioffice mcp`. The `--workspace`
(or `AIOFFICE_WORKSPACE`) defines the sandbox root.

```jsonc
// Generic stdio MCP client config
{ "command": "/Users/you/.aioffice/bin/aioffice",
  "args": ["mcp", "--workspace", "/Users/you"],
  "env": { "AIOFFICE_WORKSPACE": "/Users/you" } }
```

For TonoBraid, add an entry to the **array** in `~/.tonoagent/mcp.json` and restart the app:

```jsonc
{ "mcpServers": [
  { "name": "aioffice", "transport": "stdio",
    "command": "/Users/you/.aioffice/bin/aioffice",
    "args": ["mcp", "--workspace", "/Users/you"],
    "env": { "AIOFFICE_WORKSPACE": "/Users/you" }, "enabled": true } ] }
```

This skill describes the **CLI**; the MCP tools take the same arguments and return the same
envelopes, so everything below applies to both entry points.

## Mental model

Internalize these five ideas and you will drive AIOffice correctly:

1. **One JSON envelope per command.** Every invocation prints exactly:
   ```json
   { "ok": true|false, "data": {…}|null,
     "error": { "code", "message", "suggestion", "candidates"? }|null,
     "meta": { "file"?, "rev"?, "elapsedMs", "version", "warnings"? } }
   ```
   Parse it. `ok:false` is a normal, recoverable outcome — not a crash.

2. **Errors teach.** When `ok` is false, `error.suggestion` tells you the fix and
   `error.candidates` lists the nearest valid alternatives. **Read them and retry** instead of
   guessing. Example: address `/body/p[99]` → `invalid_path` with
   `candidates: ["/body/p[1]","/body/p[2]"]`. A wrong property name returns the supported set.
   A wrong key returns the right key. You almost never need to consult docs at runtime — the
   binary corrects you.

3. **1-based stable addressing.** Paths are slash-separated, indices start at 1, and `query`
   returns canonical paths that `get` and `edit` accept unchanged: `/body/p[3]`,
   `/Sheet1/A1`, `/slide[2]/shape[3]`. Shapes also have stable `@id` forms
   (`/slide[2]/shape[@id=7]`) that survive reordering.

4. **The render → look → fix loop.** AIOffice can render docx/xlsx to HTML, any slide to SVG,
   and **any file to PNG or PDF** via the system browser — no Office. After a visual edit,
   `render --to png` and actually *look* at the result, then fix. For non-visual problems
   (schema violations, dangling references) use `validate` — it is cheaper than rendering.

5. **Sandbox + snapshots.** All file access is confined to a workspace (default: cwd, or
   `--workspace <dir>` / `AIOFFICE_WORKSPACE`). A path escaping it fails with `sandbox_denied`.
   Every `edit` auto-snapshots the pre-image first (last 20 per file); `snapshot restore` rolls
   back, and the restore is itself snapshotted.

The intended agent loop:

```
aioffice read  <file> --view outline    # orient
aioffice query <file> "<selector>"      # find -> canonical paths
aioffice get   <file> <path>            # inspect one node
aioffice edit  <file> --ops '[...]'     # atomic batch (auto-snapshots)
aioffice validate <file>                # prove the file is still sound
```

## Verb cheat-sheet (all 18)

| verb | one-liner |
|------|-----------|
| `create`   | New empty doc (kind from extension), or import `--from notes.md` / `data.csv`. |
| `read`     | Read as `--view outline\|text\|stats\|structure\|properties\|embeds\|revisions\|comments\|styles\|fields\|markdown\|csv`. |
| `query`    | Find nodes with a CSS-like selector → stable canonical paths. |
| `get`      | Get one node + its properties by path. |
| `edit`     | Apply an atomic batch of `--ops` (set/add/remove/move/replace/accept/reject/extract); or `--set/--add/--remove/--find/--replace` sugar. Auto-snapshots. |
| `render`   | Render whole doc or a `--scope` to `--to html\|svg\|text\|png\|pdf`. |
| `validate` | OOXML conformance + lint rules → issues with suggestions. |
| `template` | Merge `{{key}}` placeholders with JSON; a JSON **array** runs a mail merge. |
| `audit`    | Accessibility + quality findings (data, exit 0); `--fix` applies safe autofixes. |
| `diff`     | Semantic compare against another file or a `--snapshot N`. |
| `convert`  | Cross-format conversion (docx/xlsx/pptx ↔, ↔ md/csv, → pdf/png/svg/html); lossy parts named. |
| `snapshot` | `list` / `restore` the automatic pre-edit ring (last 20 per file). |
| `doctor`   | Diagnose runtime, workspace, handlers, dependencies, browser. |
| `schema`   | Machine-readable JSON of the whole surface — introspect instead of guessing. |
| `help`     | Progressive docs: `aioffice help [addressing\|selectors\|properties-docx\|…\|errors]`. |
| `preview`  | Live browser preview (`open`/`selection`/`close`); click in the page = select a path. |
| `mcp`      | Run the stdio MCP server (same 17 tools). |
| `version`  | Print the version. |

Global flags: `--pretty` (indented JSON), `--quiet` (suppress success envelopes),
`--workspace <dir>` (sandbox root). Exit codes: `0` ok · `2` user/input · `3` internal/format ·
`4` sandbox_denied · `5` unsupported_feature.

## Addressing grammar (compact)

A path always starts with `/`; `/` alone is the document root. Indices are 1-based. When a
path doesn't resolve you get `invalid_path` with `candidates`.

**docx** — `/body/p[3]` (paragraph) · `/body/p[3]/run[2]` (run) ·
`/body/table[1]/tr[2]/tc[1]` (table cell) · `/header[1]/p[1]` (header) ·
`/body/p[1]/omath[1]` (inline equation) · `/style[@id=Callout]` · `/section[1]` (page setup) ·
`/revision[@id=3]` · `/comment[@id=2]` · `/embed[1]` · `/properties` (metadata).

**xlsx** — `/Sheet1/A1` (cell) · `/Sheet1/A1:C10` (range) · `/Sheet1/row[3]` · `/Sheet1/col[C]` ·
`/Sheet1/table[@name=Sales]` (ListObject) · `/Sheet1/chart[1]` · `/Pivot/pivot[@name=X]` ·
`/Sheet1/conditionalFormat[1]` · `/'Q3 Data'/B2` (quote sheet names with spaces; double a quote
to escape: `'O''Brien'`) · `/properties`.

**pptx** — `/` (root: slide size + sections) · `/slide[2]` · `/slide[2]/shape[3]` (or
`/slide[2]/shape[@id=7]`) · `/slide[2]/shape[3]/p[1]` (paragraph in a shape) ·
`/slide[2]/notes` (speaker notes) · `/slide[2]/chart[1]` · `/slide[2]/table[1]` ·
`/section[1]` · `/master[1]/layout[2]` · `/properties`.

**Selectors** for `query`: `element[attr OP value]:contains('text')`. Operators `= != > >= < <=`
(numeric) and `*=` (substring). Examples: `p[style=Heading1]`, `p:contains('Q3')`,
`cell[value>100]`, `cell[formula*=SUM]`, `shape:contains('Revenue')`, `*[name*=Total]`.

Full grammar at runtime: `aioffice help addressing`, `aioffice help selectors`,
`aioffice help properties-docx|properties-xlsx|properties-pptx`.

## The edit op model

`edit` applies an **atomic batch** of operations to a file. Either every op succeeds and the
batch is written, or the first failure rolls back the **entire batch** (the file is untouched).
Ops:

```jsonc
{"op":"set",    "path":"<path>", "props":{…}}                  // change properties
{"op":"add",    "path":"<path>", "type":"p|table|chart|…",     // insert a node
                "position":"inside|before|after", "props":{…}}
{"op":"remove", "path":"<path>"}                               // delete a node
{"op":"move",   "path":"<path>", "position":"front|back|…"}    // reorder (z-order / timeline)
{"op":"replace","path":"<scope>|/", "props":{find,replace,…}}  // scoped find/replace
{"op":"accept"|"reject", "path":"/revision[@id=N]"}            // docx tracked changes
{"op":"extract","path":"/embed[1]", "props":{"to":"out.xlsx"}} // write an embed's bytes out
```

Control flags:

- `--dry-run` — validate + report what *would* change, write nothing.
- `--expect-rev <rev>` — optimistic concurrency. `meta.rev` is the first 12 hex of the file's
  SHA-256; pass back the rev you last read and the edit fails with `stale_address` if the file
  changed underneath you. Re-read and retry.
- `--track [--author NAME]` (docx) — record text edits as `w:ins`/`w:del` revisions.

Sugar for one-liners: `--set <path> k=v…`, `--add <path> --type T k=v…`, `--remove <path>`, and
document-wide `--find X --replace Y [--regex] [--match-case] [--whole-word] [--track]`.

## Golden examples (every command was run and verified)

### 1. Create, edit, and read a Word document

```bash
aioffice create note.docx
aioffice edit note.docx --ops '[
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"Status Report","style":"Heading1"}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"All systems nominal."}}]'
aioffice edit note.docx --remove '/body/p[1]'   # drop the empty para `create` left behind
aioffice read note.docx --view outline
# → {"ok":true,"data":{"view":"outline","headings":[{"path":"/body/p[1]","level":1,"text":"Status Report","children":[]}]},…}
```

### 2. Write an Excel formula and read its evaluated value

```bash
aioffice create sales.xlsx
aioffice edit sales.xlsx --ops '[
  {"op":"set","path":"/Sheet1/A1","props":{"value":"Item"}},
  {"op":"set","path":"/Sheet1/B1","props":{"value":"Amount"}},
  {"op":"set","path":"/Sheet1/B2","props":{"value":1200}},
  {"op":"set","path":"/Sheet1/B3","props":{"value":850}},
  {"op":"set","path":"/Sheet1/B4","props":{"value":"=SUM(B2:B3)","bold":true,"numberFormat":"#,##0"}}]'
aioffice get sales.xlsx /Sheet1/B4
# → "value":2050,"formula":"=SUM(B2:B3)","cachedValue":2050,"text":"2,050"
```

Formulas are evaluated and cached at write time. A leading `=` in a `value` makes it a formula
(use `valueType:"text"` to keep a literal `=x`).

### 3. Build a slide with a chart

```bash
aioffice create pitch.pptx
aioffice edit pitch.pptx --ops '[
  {"op":"set","path":"/slide[1]","props":{"background":"0F172A"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"text":"Q3 Review","x":2,"y":6,"w":24,"h":3,"fontSize":54,"bold":true,"color":"FFFFFF"}}]'
aioffice edit pitch.pptx --ops '[
  {"op":"add","path":"/slide[1]","type":"slide","position":"after","props":{"title":"Revenue by quarter"}}]'
aioffice edit pitch.pptx --ops '[
  {"op":"add","path":"/slide[2]","type":"chart","props":{"kind":"bar","categories":["Q1","Q2","Q3"],
    "series":[{"name":"Revenue","values":[120,150,210]}],"title":"Revenue ($k)","x":2,"y":5,"w":20,"h":11}}]'
aioffice get pitch.pptx '/slide[2]/chart[1]'   # → chartKind:"bar", categories/series echoed back
```

Note: `title` is a prop of **adding** a slide, not of `set`-ing an existing one — `set /slide[1]
{title}` is rejected (it tells you the valid slide props: background/transition). For an existing
slide, add a title shape or a new slide with a title.

### 4. The render → look → fix loop

```bash
aioffice render pitch.pptx --to png --scope '/slide[1]' -o slide1.png   # → 1280x720 PNG written
# look at slide1.png, notice an overflow, fix it, render again
aioffice render pitch.pptx --to svg --scope '/slide[2]' -o slide2.svg    # SVG: shapes carry data-aio-path
```

The PNG is a real screenshot rendered by the system browser. Over MCP, `render --to png` also
returns the image inline so the model can see it directly.

### 5. Audit for accessibility and auto-fix

```bash
aioffice audit bad.docx
# → findings are DATA, exit 0: a11y_no_doc_title (warning), a11y_no_alt_text (error), each "autofixable":true
aioffice audit bad.docx --fix
# → {"summary":{"errors":0,"warnings":0,"infos":0},"fixed":2,"remaining":[]}
```

Findings carry a stable `id` (`code#path`); `--fix` applies only the safe autofixes and never
destroys content. Report-only findings stay in `remaining`.

### 6. Diff against a snapshot to review an edit

```bash
aioffice edit q3-report.docx --find '+12%' --replace '+11.8%'   # every edit auto-snapshots first
aioffice diff q3-report.docx --snapshot 4 --view detailed
# → {"changes":[{"kind":"modified","path":"/body/table[1]/tr[4]/tc[3]","before":"+12%","after":"+11.8%","detail":"cell"}],…}
aioffice snapshot restore q3-report.docx   # roll it back; the restore is itself snapshotted
```

### 7. Convert between formats

```bash
aioffice convert note.docx note-deck.pptx
# → {"from":"docx","to":"pptx","blocksWritten":2,"dropped":[],"written":".../note-deck.pptx"}
aioffice convert pitch.pptx pitch-outline.docx
# → meta.warnings:[{"code":"convert_lossy","message":"…charts… is not converted"}]   ← lossy parts are named
```

Conversion is inherently lossy; whatever didn't survive is named in a `convert_lossy` warning.

### 8. Mail-merge personalized letters from JSON

```bash
# letter.docx body has {{name}}, {{balance}}, {{tier}} placeholders
cat recipients.json   # [{"name":"Ada","balance":"$1,200","tier":"Gold"}, {"name":"Grace",…}]
aioffice template letter.docx --data @recipients.json --output 'letter-{name}.docx'
# → produced: letter-Ada.docx, letter-Grace.docx
aioffice read letter-Ada.docx --view text   # → "Dear Ada,"  /  "Your balance is $1,200…Gold member."
```

A JSON **object** fills one document; a JSON **array** runs a mail merge. Template body uses
**double** braces `{{key}}`; the `--output` filename pattern uses **single** braces `{Field}`
and `{n}` (1-based index).

## Gotchas (the real ones, all verified)

- **Quote bracket paths in your shell.** `/body/p[3]`, `/slide[2]/chart[1]`, and selectors
  contain `[` `]` `*` `(` `)` — wrap every path/selector in single quotes so the shell doesn't
  glob or split them: `aioffice get f.docx '/body/p[3]'`.
- **Edit batches are atomic.** If any op in `--ops` fails, the *whole* batch rolls back and the
  file is untouched — including ops that came before the bad one. The error names the offending
  op as `ops[N] (…)`. Fix that op and re-send the batch.
- **`create` leaves an empty first paragraph in docx.** A fresh `.docx` has one empty `/body/p[1]`;
  your added content lands at `p[2]+`. Remove `p[1]` afterward (it then re-indexes), or add
  `position:"before"`/`"after"`.
- **`title` on a slide is an `add` prop, not a `set` prop.** `set /slide[1] {title}` errors;
  it's only honored when you `add type:"slide"`. The error lists the real settable slide props.
- **docx table styling uses geometry props, not a style name.** On a table use
  `{headerRow, borders:"all"|"outer"|"none", shading, columnWidths, width, alignment}` —
  `set {style}` on a table is `unsupported_feature` (the error names the valid props).
- **Newer functions are stored with the `_xlfn.` prefix.** `XLOOKUP`, `LET`, `IFS`, `TEXTJOIN`
  etc. read back as `=_xlfn.XLOOKUP(…)` in the `formula` field — that is correct OOXML; Excel
  shows the clean name. The `cachedValue`/`value` is the evaluated result.
- **Bubble/scatter charts need numeric data.** A bubble chart's X column (and scatter's X axis)
  must be numbers, not labels — the header row is the only text allowed.
- **`extract` takes `props.to`, not `props.dest`.** `{"op":"extract","path":"/embed[1]",
  "props":{"to":"out.xlsx"}}`. Extracted bytes are byte-identical to what was embedded.
- **Document properties nest under `data.properties.{core,custom}`.** Read the title at
  `data.properties.core.title` (identical shape across docx/xlsx/pptx). Write with
  `set /properties {title, author, custom:{…}}`.
- **pptx lengths default to centimeters.** A bare number means cm (`"x": 2.5`); strings take a
  unit suffix (`"36pt"`, `"1in"`, `"96px"`, `"5cm"`).
- **pptx chart data is cached literally.** Charts added with inline `categories`/`series` report
  `dataEditable:false` (no embedded workbook). Use `dataFrom:"book.xlsx!Sheet1/A1:B5"` to source
  from a workbook.
- **When stuck, ask the binary.** `aioffice schema [verb]` is the machine-readable surface and
  `aioffice help <topic>` is the prose. You rarely need anything else — the error suggestions
  carry you the rest of the way.

See also: `docs/COOKBOOK.md` for 10 task-oriented, copy-paste recipes.
