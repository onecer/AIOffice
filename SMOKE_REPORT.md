# AIOffice 0.1.0 — Integration Smoke Report

Date: 2026-06-12 · Machine: macOS 26.3.0 arm64 · dotnet 10.0.300 (TFM net10.0)
All commands below were actually executed; outputs are trimmed but real.

## 1. Build — PASS

```
$ dotnet build AIOffice.sln -warnaserror --no-incremental
已成功生成。 0 个警告 0 个错误
```

Integration fixes applied to get the surface coherent (build was already green; the drift was behavioral):

| Drift found | Fix |
| --- | --- |
| `WordHandler.Render`/`Template` read arg key `"out"`, while CLI, MCP, Excel and Pptx all use `"output"` → `render -o` silently wrote nothing for docx | Unified on `"output"` (docs/MCP.md is the source of truth); two Word tests updated to the canonical key |
| `AIOffice.Mcp.HandlerDiscovery` required a *parameterless* ctor (`GetConstructor(Type.EmptyTypes)`), but all three handlers have `(SnapshotStore? = null)` ctors → **zero format handlers registered over MCP**; `office_status` reported `healthy:false`, every `office_*` document tool was dead | Discovery now accepts ctors whose params are all optional-or-SnapshotStore (same convention as the CLI's discovery), injects one shared SnapshotStore, and reports which kinds self-snapshot |
| MCP `CommandService` always took its own pre-image snapshot even when the handler also snapshots → duplicate ring entries | `CommandService` now accepts `handlerManagedSnapshots` and skips its snapshot for those kinds (one snapshot per successful edit) |

## 2. Tests — PASS (272/272)

```
$ dotnet test AIOffice.sln
AIOffice.Core.Tests   104 passed
AIOffice.Word.Tests    52 passed
AIOffice.Excel.Tests   41 passed
AIOffice.Pptx.Tests    48 passed
AIOffice.Mcp.Tests     27 passed   (was 26; discovery tests strengthened — they now ASSERT all 3 handlers are found instead of tolerating an empty registry, which is how the MCP bug had slipped through)
```

Round-trip law is enforced by `tests/*/RoundTripTests.cs` / `RoundTripLawTests.cs`; validator oracle (OpenXmlValidator = 0 errors) runs after mutating tests.

## 3. End-to-end CLI smoke (temp workspace /tmp/aioffice-smoke) — PASS

### doctor — PASS
```
$ aioffice doctor --json
ok:true · handlers docx/xlsx/pptx all "ready", snapshotOwner:"handler"
deps: DocumentFormat.OpenXml 3.5.1 · ClosedXML 0.105.0 · ModelContextProtocol.Core 1.4.0
```

### docx — PASS
```
$ aioffice create demo.docx                          → ok, rev c174f124a87f
$ aioffice edit demo.docx --add /body --type p text="AIOffice 自研引擎首测" style=Heading1
                                                     → ok, applied:1, path /body/p[2], snapshot:1
$ aioffice read demo.docx --view outline             → headings:[{path:/body/p[2], level:1, text:"AIOffice 自研引擎首测"}]
$ aioffice render demo.docx --to html -o demo.html   → ok, written:/private/tmp/aioffice-smoke/demo.html
$ cat demo.html                                      → <p></p>\n<h1>AIOffice 自研引擎首测</h1>
```
(The `-o` write FAILED before integration — the `"out"`-vs-`"output"` drift above; passes now.)

### xlsx — PASS
```
$ aioffice edit data.xlsx --ops '[{"op":"set","path":"/Sheet1/A1","props":{"value":1}},
   {"op":"set","path":"/Sheet1/A2","props":{"value":2}},
   {"op":"set","path":"/Sheet1/A3","props":{"value":"=SUM(A1:A2)"}}]'
                                  → ok, applied:3 (A3 applied:["formula"])
$ aioffice get data.xlsx /Sheet1/A3
                                  → {value:3, formula:"=SUM(A1:A2)", cachedValue:3, text:"3"}
$ aioffice validate data.xlsx     → {valid:true, errors:0}
```
Note: there is no `formula` prop — formulas are set via `value` with a leading `=`
(the `invalid_args` envelope for a wrong prop name correctly listed the supported props as candidates).

### pptx — PASS
```
$ aioffice edit deck.pptx --ops '[{"op":"add","path":"/slide[2]","type":"slide","props":{"title":"Hello"}}]'   → ok, slides:2
$ aioffice edit deck.pptx --ops '[{"op":"add","path":"/slide[1]","type":"shape",
   "props":{"text":"AIOffice","x":"2cm","y":"3cm","w":"10cm","h":"2cm"}}]'                                     → ok, target /slide[1]/shape[@id=2]
$ aioffice render deck.pptx --to svg --scope '/slide[1]' -o slide1.svg → ok; slide1.svg contains "AIOffice"
$ aioffice validate deck.pptx                                          → {valid:true, count:0}
```

### sandbox — PASS
```
$ aioffice read ../outside.docx
→ {ok:false, error:{code:"sandbox_denied", message:"Path escapes the workspace sandbox: ../outside.docx", suggestion:"Use a path inside the workspace (/private/tmp/aioffice-smoke)…"}}
exit code = 4
```

### snapshots — PASS
```
$ aioffice snapshot list data.xlsx     → count:1, [{number:1, rev:f81101c77737}]  (auto pre-edit snapshot)
$ aioffice snapshot restore data.xlsx  → restored number:1; file rev 5ad012be52bd → f81101c77737 (verified with shasum)
```

### schema / help — PASS
```
$ aioffice schema          → JSON parses; verbs: create, read, query, get, edit, render, validate,
                             template, snapshot, doctor, schema, help, mcp, version
$ aioffice help addressing → addressing grammar text (docx /body/p[3]…, xlsx /Sheet1/A1:C10, pptx /slide[2]/shape[3])
```

### error-contract probes — PASS
```
edit --expect-rev 000000000000 → {code:"stale_address", suggestion:"Re-run 'aioffice read'…"}, exit 2, nothing written
render deck.pptx --to png      → {code:"unsupported_feature", suggestion:"Render --to svg … rsvg-convert"}, exit 5
unsupported formula (=SEQUENCE(3)) → ok:true + meta.warnings:[{code:"formula_not_evaluated", …"Excel computes it when the file opens"}]
```

## 4. MCP stdio server — PASS

Spawned `dotnet run --project src/AIOffice.Cli -- mcp --workspace /tmp/aioffice-smoke`, driven by a node JSON-RPC script:

```
initialize  → serverInfo {name:"aioffice", title:"AIOffice — AI-native Office document tools", version:"0.1.0"}
tools/list  → 12 tools: file_snapshot, office_create, office_edit, office_get, office_help,
              office_query, office_read, office_render, office_schema, office_status,
              office_template, office_validate          (preview_* reserved for M1, by design)
office_status → healthy:true; checks: workspace_writable ✓, snapshot_store_writable ✓,
                format_handlers ✓ ".docx, .pptx, .xlsx"
office_create mcp-made.docx → ok, rev 0eb98913d2ee
office_edit  add p "hello from MCP" → ok, applied:1, snapshot:1 (exactly one — no duplicate ring entries)
```
(Before integration: `office_status` → `healthy:false, "no format handlers registered yet (M0 in progress)"` — FAIL, root-caused and fixed, see §1.)

## 5. Published binary — PASS

```
$ dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release -p:PublishSingleFile=true --self-contained -o dist/osx-arm64
$ ls -l dist/osx-arm64/aioffice   → 37,523,497 bytes (35.8 MiB), self-contained single file
$ dist/osx-arm64/aioffice doctor                       → ok (handlers ready, .NET 10.0.8)
$ dist/osx-arm64/aioffice create/edit/validate out.docx → ok / ok / {valid:true, count:0}
$ dist/osx-arm64/aioffice read ../nope.docx             → exit 4
```

## 6. Manual-check fixtures (open these in real Office)

`fixtures/manual-check/`: `demo.docx` (Heading1 中文), `data.xlsx` (A3 =SUM(A1:A2) cached 3, B1 =SEQUENCE(3) uncached — Excel should compute on open), `deck.pptx` (2 slides, positioned text shape). All pass OpenXmlValidator with 0 errors (the in-repo proxy for "opens in real Office").

## Known gaps (honest list)

1. **xlsx round-trip is NOT byte-identical** — measured & pinned in `RoundTripLawTests`: a no-edit ClosedXML save rewrites `xl/worksheets/sheetN.xml` only (root-element attribute order; cached `<v>` dropped for formulas, recomputed by Excel on open). docx/pptx round-trips ARE byte-identical (tests enforce).
2. **Formula engine coverage** — SUM-class functions evaluate with cached values; dynamic-array functions (SEQUENCE, …) are saved as formula text without cache + `formula_not_evaluated` warning. Never silent.
3. **render --to png** → typed `unsupported_feature` (M1); svg/html available today. docx render is html/text only.
4. **MCP tool count is 12**, not 13 — `preview_*` is reserved M1 surface and intentionally not registered yet.
5. **pptx master/layout addressing** reserved M1.
6. Cosmetic envelope drift: `validate` data shape differs per format (docx/pptx `{valid,count,issues}` vs xlsx `{valid,errors,warnings,issues}`). Functional, but worth unifying in M1.
7. Snapshot ring lives at `~/.aioffice/snapshots` (user-level, outside the workspace by design); `--workspace` does not relocate it.

---

# AIOffice 0.2.0 (M1) — Integration Smoke Report

Date: 2026-06-12 · Machine: macOS 26.3.0 arm64 · dotnet 10.0.300 (TFM net10.0) · Chrome installed
All commands below were actually executed; outputs are trimmed but real.

## 1. Build — PASS

```
$ dotnet build AIOffice.sln -warnaserror
已成功生成。 0 个警告 0 个错误      (now 8 src + 7 test projects: + AIOffice.Render, AIOffice.Preview)
```

Integration drift found & fixed (each fix carries/updates a test):

| Drift found | Fix |
| --- | --- |
| `tests/AIOffice.Core.Tests/EnvelopeTests.cs` pinned `meta.version == "0.1.0"` → failed on the 0.2.0 bump | Asserts `Meta.ToolVersion` (the same source the envelope writes), like the MCP tests already do |
| `SchemaHelpStatusTests` expected a 13-verb surface → failed once `preview` joined `SurfaceSchema` | Updated to 14 (verb-name set equality was already asserted; only the count was stale) |
| `PptxEditor.PickLayout` accepted **only JSON-number** `props.layout`, but the CLI `--ops`/sugar and the MCP `office_edit` schema (`additionalProperties:{type:"string"}`) deliver props **string-valued** → `{"layout":"1"}` was rejected with `invalid_args` (caught live in the pptx smoke below, FAIL → root-caused) | `PickLayout` now also parses string-backed integers (invariant culture); non-numeric strings like `"Blank"` still fail typed. Regression test `AddSlide_WithStringLayoutIndex_BindsThatLayout` added |

## 2. Tests — PASS (383/383 across 7 projects)

```
$ dotnet test tests/<each>
AIOffice.Core.Tests     104 passed
AIOffice.Word.Tests      73 passed   (was 52: + header/footer, data-aio-path render contract)
AIOffice.Excel.Tests     59 passed   (was 41: + charts bar/line/pie, td data-aio-path)
AIOffice.Pptx.Tests      72 passed   (was 48: + master/layout read, g-wrap data-aio-path, layout-prop regression)
AIOffice.Mcp.Tests       30 passed   (was 27: + preview tools, png image block; token budget ≤3500 unchanged)
AIOffice.Preview.Tests   24 passed   (new)
AIOffice.Render.Tests    21 passed   (new; includes real-browser screenshot tests against local Chrome)
```

## 3. End-to-end CLI smoke (temp workspace /tmp/aio-m1-smoke, `dotnet run --project src/AIOffice.Cli --`) — PASS

### docx: header + structure + png — PASS
```
$ aioffice create report.docx --title "M1 Smoke"                       → ok, rev b06572e4c06f
$ aioffice edit report.docx --ops '[{"op":"add","path":"/header[1]","type":"header","props":{"text":"机密"}}]'
                                                                       → ok, applied:1, paragraph:/header[1]/p[1], snapshot:1
$ aioffice read report.docx --view structure                           → data.headers:[{path:/header[1], text:"机密", children:[/header[1]/p[1]…]}]
$ aioffice render report.docx --to png                                 → {format:"png", written:report.png, sizeBytes:8586}
$ file report.png                                                      → PNG image data, 1280 x 720, 8-bit/color RGB
```

### xlsx: values + bar chart + validate + get — PASS
```
$ aioffice edit sales.xlsx --ops '[…6 set ops A1:B3…,
   {"op":"add","path":"/Sheet1","type":"chart","props":{"kind":"bar","dataRange":"A1:B3","anchor":"D2","title":"Sales by Region"}}]'
                                       → ok, applied:7, chart path /Sheet1/chart[1], series:1
$ aioffice validate sales.xlsx         → {valid:true, errors:0, warnings:0}    (OpenXmlValidator over the hand-built ChartPart)
$ aioffice get sales.xlsx '/Sheet1/chart[1]'
                                       → {kind:"chart", chartKind:"bar", title:"Sales by Region", dataRange:"A1:B3", anchor:"D2", series:1}
```

### pptx: add slide (layout) + master + scoped png — PASS (after one real FAIL)
```
$ aioffice edit deck.pptx --ops '[{"op":"add","path":"/slide[2]","type":"slide","props":{"layout":"1","title":"第二页"}}]'
   1st attempt → FAIL: invalid_args "props.layout is not a valid layout index: \"1\""   ← string-prop drift, see §1; fixed
   after fix   → ok, slides:2
$ aioffice get deck.pptx '/master[1]'  → {kind:"master", theme:"AIOffice", layoutCount:1,
                                          layouts:[{path:/master[1]/layout[1], name:"Blank", type:"blank", usedBySlides:[1,2]}]}
$ aioffice render deck.pptx --to png --scope '/slide[1]'
                                       → {format:"png", scope:"/slide[1]", written:deck.png, sizeBytes:8662}
```
(Also probed: add slide path must be `/slide[N]` — `path:"/"` correctly returns `invalid_path` with grammar suggestion.)

### preview: open → click-map → selection → close — PASS
```
$ aioffice preview open report.docx &    → envelope printed BEFORE blocking:
   {url:"http://127.0.0.1:26500/", port:26500, pid:73942, lockfile:~/.aioffice/preview/cc461c558a01.json}
$ curl :26500/                           → page contains data-aio-path="/header[1]/p[1]", "/body/p[1]", "/body/p[2]" …
$ curl -X POST :26500/selection -d '{"paths":["/body/p[1]","/header[1]/p[1]"]}'
                                         → {paths:[…2…], rev:"0308e51174dc", updatedAt:"2026-06-12T07:50:04Z"}
$ aioffice preview selection report.docx → ok, same 2 paths + rev (matches the file's current rev)
$ aioffice preview close report.docx     → {closed:true}; lock dir empty afterwards (verified with ls)
```

### doctor (new M1 sections) — PASS
```
$ aioffice doctor → browser:{found:true, kind:"chrome", path:"/Applications/Google Chrome.app/…"}
                    preview:{lockDirectory:"~/.aioffice/preview", exists:true, lockfiles:0}
$ aioffice schema → verbs: …, preview, … (14 verbs + mcp/version CLI extras)
$ aioffice help addressing → covers /header[1]/p[1], /master[1]/layout[2], /Sheet1/chart[1]
```

## 4. MCP stdio server — PASS

Spawned `dotnet run --project src/AIOffice.Cli -- mcp --workspace /tmp/aio-m1-smoke`, driven by a python JSON-RPC script:

```
initialize         → serverInfo {name:"aioffice", version:"0.2.0"}
tools/list         → 14 tools (12 v0 + preview_open + preview_selection); schema budget test ≤3500 tokens still green
preview_open       → spawned detached 'aioffice preview open' child, waited lockfile+health
                     → {url:"http://127.0.0.1:26500/", port:26500, pid:85043}
POST /selection ["/body/p[2]"] (simulating the human click)
preview_selection  → {paths:["/body/p[2]"], rev:"0308e51174dc", updatedAt:…}
office_render to=png → content[0] = envelope text {format:"png", written:…}
                       content[1] = {type:"image", mimeType:"image/png"} — 8238 base64 chars, bytes verbatim (no downscale)
```

## 5. Published binary — PASS

```
$ dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release -p:PublishSingleFile=true --self-contained -o dist/osx-arm64
$ ls -l dist/osx-arm64/aioffice    → 37,605,465 bytes (35.9 MiB), self-contained single file (was 35.8 MiB in 0.1.0; +Render/Preview ≈ +0.1 MiB)
$ aioffice doctor                  → version 0.2.0; browser found:true kind:chrome; preview lock dir ok
$ aioffice create/edit(+header 机密)/render --to png bin.docx
                                   → ok / ok / {format:"png", sizeBytes:9437} — real 1280x720 PNG (file(1) verified)
$ aioffice preview open bin.docx --port 26555 &  → envelope first, then blocks; GET / has 10 data-aio-path attrs
$ aioffice preview close bin.docx  → {closed:true}, lockfile removed
```

## 6. M0 known-gaps status after M1

| M0 gap | Status |
| --- | --- |
| 3. `render --to png` unsupported | **CLOSED** — native png via headless Chrome/Edge (AIOffice.Render); pptx is per-slide (`--scope`, default slide 1 + warning) |
| 4. MCP tool count 12 (preview reserved) | **CLOSED** — 14 tools registered, inside the token budget reserved in docs/MCP.md §2.1 |
| 5. pptx master/layout addressing reserved | **CLOSED (read-only)** — get/query/structure on `/master[1]`, `/master[1]/layout[2]`; **editing them is still `unsupported_feature` (M2)** |
| 1. xlsx round-trip not byte-identical | unchanged (pinned in RoundTripLawTests) |
| 2. formula engine coverage | unchanged (`formula_not_evaluated` warning path) |
| 6. validate envelope shape drift | **still open** — punted again; carry to M2 |
| 7. snapshot ring location | unchanged (by design) |

New honest limits introduced in M1: png needs an installed Chrome/Edge (typed `unsupported_feature` + install suggestion otherwise — `doctor` shows what the probe found); pptx png renders exactly one slide per call (no grid yet); xlsx charts are bar/line/pie only (scatter/area → typed error naming the workaround); headers/footers are default-type only (first-page/even-odd variants → typed M2 error); preview serves one file per server process (one lockfile per file).

---

# AIOffice 0.3.0 — M2 Integration Smoke Report

Date: 2026-06-12 · Machine: macOS 26.3.0 arm64 · dotnet 10.0.300 (TFM net10.0)
All commands below were actually executed; outputs are trimmed but real.

## M2.1 Build — PASS

```
$ dotnet build AIOffice.sln -warnaserror
已成功生成。 0 个警告 0 个错误
```

## M2.2 Tests — PASS (540/540)

```
$ dotnet test AIOffice.sln
AIOffice.Core.Tests     111 passed   (was 104: +ArgParser --track flag, +FileSizeGuard logic/env tests)
AIOffice.Word.Tests     124 passed   (was 122: +file_too_large wiring on read/edit, sparse 51 MB fixture)
AIOffice.Excel.Tests    117 passed   (was 115: +file_too_large wiring)
AIOffice.Pptx.Tests     111 passed   (was 109: +file_too_large wiring)
AIOffice.Mcp.Tests       32 passed   (was 30: +office_edit track/author reach the handler ctx; token budget still green)
AIOffice.Preview.Tests   24 passed
AIOffice.Render.Tests    21 passed
```

Token budget after the office_edit schema grew `track`/`author` + accept/reject: ~1785 of 3500 tokens (chars/4) — no description trimming needed.

## M2.3 CLI smoke (temp workspace /tmp/aio-m2-smoke, `dotnet run --project src/AIOffice.Cli --`) — PASS

docx tracked changes + comments + styles + image:

```
$ aioffice create demo.docx --title "M2 Demo"        → ok, rev 7507474119de
$ aioffice edit demo.docx --track --author Reviewer --set /body/p[2] text='Tracked replacement text'
  → {applied:1, ops:[{op:set, path:/body/p[2], tracked:true, author:"Reviewer"}]}
$ aioffice read demo.docx --view revisions
  → {count:1, revisions:[{path:"/revision[@id=1]", kind:"insert", author:"Reviewer", date:"2026-06-12T09:32:05Z"}]}
$ aioffice edit demo.docx --ops '[{"op":"accept","path":"/revision[@id=1]"}]'   → {applied:1}
$ aioffice read demo.docx --view revisions            → {count:0}            ✓ empty after accept
$ aioffice validate demo.docx                         → {valid:true, count:0}

$ aioffice edit demo.docx --ops '[{"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"…","author":"Reviewer"}}]'
  → {path:"/comment[@id=1]", anchor:"/body/p[1]"}
$ aioffice read demo.docx --view comments             → {count:1, anchorText:"M2 Demo"}
$ aioffice edit demo.docx --ops '[{"op":"remove","path":"/comment[@id=1]"}]'   → ok; comments count:0

$ aioffice edit demo.docx --ops '[{"op":"add","path":"/styles","type":"style","props":{"id":"Callout","bold":true,"color":"1F4E79","fontSize":12,"alignment":"center"}}]'
$ aioffice edit demo.docx --set /body/p[2] style=Callout
$ aioffice get demo.docx /body/p[2]                   → properties.style:"Callout" ✓
$ aioffice get demo.docx '/style[@id=Callout]'        → {builtin:false, inUse:true, bold:true, color:"1F4E79"} ✓

$ aioffice render demo.docx --to png -o logo.png      → 10,197-byte real 1280x720 PNG (the image source, per the brief)
$ aioffice edit demo.docx --ops '[…add type:image src:logo.png width:6cm…]'    → {widthCm:6, heightCm:3.38}
$ aioffice validate demo.docx                         → {valid:true, count:0}
```

xlsx pivot + conditional formats + image:

```
$ aioffice create data.xlsx; edit … 4-col sales table A1:D7 (Region/Product/Quarter/Sales)
$ aioffice edit data.xlsx --ops '[…add type:pivot name:SalesPivot sourceRange:A1:D7 targetSheet:Pivot rows:[Region] values:[{field:Sales,agg:sum}]…]'
  → {path:"/Pivot/pivot[@name=SalesPivot]", targetSheet:"Pivot"}
$ aioffice get data.xlsx '/Pivot/pivot[1]'            → {name:"SalesPivot", rows:["Region"], values:[{field:"Sales",agg:"sum"}]} ✓ (addressed on its TARGET sheet)
$ aioffice edit data.xlsx --ops '[cellIs >150 fill C6EFCE on D2:D7, colorScale FFFFFF→63BE7B on D2:D7]'
  → conditionalFormat[1] + conditionalFormat[2]
$ aioffice get data.xlsx '/Sheet1/conditionalFormat[1]'  → {cfKind:"cellIs", operator:">", value:"150", fill:"C6EFCE"}
$ aioffice edit data.xlsx --ops '[…add type:image src:logo.png anchor:F2 widthPx:120…]'  → {anchor:"F2", widthPx:120, heightPx:68}
$ aioffice validate data.xlsx                         → {valid:true, errors:0, warnings:0}
```

pptx background + image + notes:

```
$ aioffice create deck.pptx --title "M2 Deck"
$ aioffice edit deck.pptx --ops '[set /slide[1] background:0F172A, add type:image src:logo.png w:8cm, set /slide[1]/notes text:"Opening line.\nMention the Q3 numbers."]'
  → {applied:3, image target:"/slide[1]/shape[@id=3]"}
$ aioffice read deck.pptx --view outline              → slide 1 carries notes:"Opening line. Mention the Q3 numbers." ✓
$ aioffice render deck.pptx --to png -o deck-slide1.png  → real 1280x720 PNG, 10,724 bytes (file(1) verified) ✓ still works
$ aioffice validate deck.pptx                         → {valid:true, count:0}
```

File-size guard + flags surface:

```
$ AIOFFICE_MAX_FILE_MB=0 aioffice read demo.docx
  → {ok:false, error:{code:"file_too_large", message:"File is 0.0 MB, over the 0 MB limit: demo.docx",
       suggestion:"Split the document into smaller files, or raise the limit with AIOFFICE_MAX_FILE_MB=<mb> …"}}  exit=2
$ aioffice doctor → data.limits {maxFileMb:50, maxFileMbDefault:50, maxFileMbEnv:"AIOFFICE_MAX_FILE_MB"}; =200 with the env set
$ aioffice schema edit → usage shows [--track] [--author NAME]; options list ['ops','set','add','type','remove','position','track','author','dry-run','expect-rev']
$ AIOFFICE_AUTHOR="Env Author" aioffice edit demo.docx --track --set /body/p[2] text='…'
  → ops[0].author:"Env Author" ✓ (resolution: props.author > --author > AIOFFICE_AUTHOR > "AIOffice")
```

## M2.4 MCP stdio server — PASS

Spawned `dotnet run --project src/AIOffice.Cli -- mcp --workspace /tmp/aio-m2-smoke`, driven by a python JSON-RPC script:

```
initialize       → serverInfo {name:"aioffice", version:"0.3.0"}
tools/list       → 14 tools (count unchanged from M1; office_edit schema gained track/author + accept/reject)
office_edit {track:true, author:"MCP Reviewer", ops:[set /body/p[2] text…]}
                 → {ops:[{tracked:true, author:"MCP Reviewer"}]}
office_read {view:"revisions"}
                 → count:2 — /revision[@id=1] delete + /revision[@id=2] insert, both author="MCP Reviewer" ✓
office_edit {ops:[{op:"reject", path:"/body"}]}  → ok (cleanup)
```

## M2.5 Manual-check fixtures regenerated (for a human to open in real Office)

- `demo.docx` — ONE PENDING tracked change (insert+delete pair by "Reviewer"), 1 comment, custom `Callout` style applied, embedded 6 cm PNG (generated via `aioffice render --to png`). validate: 0 errors.
- `data.xlsx` — SalesPivot pivot (rows Region × columns Quarter, sum of Sales, on sheet "Pivot", refreshOnLoad), bar chart, cellIs + colorScale conditional formats on D2:D7, anchored image. validate: 0 errors.
- `deck.pptx` — real `p:bg` dark background (0F172A), white title, embedded image, speaker notes. validate: 0 errors.
- `word-sample.docx` / `excel-sample.xlsx` / `pptx-showcase.pptx` — regenerated by the test suite, now also carrying the M2 surface (pending revision + comment + style + image / pivot + dataBar + colorScale + image / background + notes + image).

## M2.6 Published binary — PASS

```
$ dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release -p:PublishSingleFile=true --self-contained -o dist/osx-arm64
$ ls -l dist/osx-arm64/aioffice    → 37,667,545 bytes (35.9 MiB), self-contained single file (was 37.6 MB / 35.9 MiB in 0.2.0; M2 adds ≈60 KB)
$ aioffice doctor                  → version 0.3.0; limits.maxFileMb:50; handlers docx/xlsx/pptx all "ready"
$ aioffice create/edit --track --author Reviewer/read --view revisions/accept/validate (bin.docx)
                                   → tracked edit ok (author Reviewer) → 1 revision → accept → 0 revisions → valid:true
$ AIOFFICE_MAX_FILE_MB=0 aioffice read bin.docx   → file_too_large ✓
```

## M2.7 Honest limits introduced/kept in M2

- Tracked changes are **text-level only** (w:ins/w:del + paragraph-mark ins/del): tracked formatting, tracked moves and tracked find&replace are typed `invalid_args`/`unsupported_feature` with the workaround named → M3.
- Conditional formats cover 4 kinds (cellIs/colorScale/dataBar/containsText); the other 7 upstream kinds answer `unsupported_feature` listing the supported set.
- Pivot tables: rows/columns/filters + sum/average/count/min/max. layout/topN/calculatedField/showDataAs → M3.
- Images are PNG/JPEG only (header-sniffed); SVG answers a typed error.
- **Large-file streaming did NOT ship** — it needs a dedicated benchmark-driven pass; M2 ships the `file_too_large` size guard (default 50 MB, `AIOFFICE_MAX_FILE_MB`) instead. Moved to M3.
- M0 gap 6 (validate envelope shape drift) — **still open**, punted again; carry to M3.
