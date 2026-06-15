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

---

# AIOffice 0.4.0 (M3) — Integration Smoke Report

Date: 2026-06-12 · Machine: macOS 26.3.0 arm64 · dotnet 10.0.300 (TFM net10.0) · Browser: Google Chrome (auto-detected)
All commands below were actually executed via `dotnet run --project src/AIOffice.Cli --` in a temp workspace (`/tmp/aio-m3-smoke`); outputs are trimmed but real.

## M3.1 Build — PASS

```
$ dotnet build AIOffice.sln -warnaserror
已成功生成。 0 个警告 0 个错误
```

Integration fixes applied (drift found while integrating the M3 worker drops):

| Drift found | Fix |
| --- | --- |
| `WordHandler.OpenPackage` did not wrap `System.IO.FileFormatException` (System.IO.Packaging's "not a package" signal) — with the size cap now unlimited, a huge non-zip `.docx` crashed with an unwrapped exception instead of a typed envelope | Added `FileFormatException` to the `format_corrupt` catch list (xlsx/pptx already wrapped it via their broader catch) |
| `tests/AIOffice.{Word,Pptx}.Tests/FileSizeGuardTests` were written for the M2 50 MB *default* and went red after the cap flip | Rewritten to pin `AIOFFICE_MAX_FILE_MB=50` explicitly (comment cites the M3 功能第一 directive) + a new test asserting an oversized file PASSES the guard when the env var is unset (Excel's equivalent had already been updated by its worker) |
| `tests/AIOffice.Core.Tests/FileSizeGuardTests.Env_var_overrides_the_default_limit` asserted the old 50 MB fallback | Rewritten: unset/unparsable env → `MaxFileMb == null` (unlimited); added a 200 MB-sparse-file pass test |

## M3.2 Tests — PASS (724/724)

```
$ dotnet test AIOffice.sln
AIOffice.Core.Tests     112 passed   (was 111; size-cap flip tests rewritten + unlimited-default test)
AIOffice.Word.Tests     186 passed   (lists, links, bookmarks, footnotes, sections, format revisions, replies)
AIOffice.Excel.Tests    168 passed   (streaming, scatter/area, names, sheet props; ~39 MB fixture generated per run)
AIOffice.Pptx.Tests     159 passed   (charts, transitions, geometries, z-order)
AIOffice.Mcp.Tests       44 passed   (was 32; +12 CrossDocDataFrom incl. one over the real MCP wire; token budget still ≤3500)
AIOffice.Render.Tests    31 passed   (was 21; +10 PdfRenderer incl. one against real Chrome)
AIOffice.Preview.Tests   24 passed
```

## M3.3 docx — PASS

```
create report.docx → edit (5 list ops: number ×2 + nested level 1, bullet + nested)
$ read --view text
  /body/p[3] "1. Step one" · /body/p[4] "  1. Nested under one" · /body/p[5] "2. Step two"
  /body/p[6] "• Bullet point" · /body/p[7] "  • Nested bullet"                              ✓ markers
$ render --to html      → has <ol>: True · has <ul>: True (real nested lists)               ✓
edit: bookmark "Steps" on /body/p[3] + external link + anchor link + footnote on /body/p[8]
$ validate              → valid:true, count:0                                               ✓
$ edit --set /section[1] pageSize=A4 orientation=landscape; get /section[1]
  → {orientation:"landscape", pageSize:"A4", widthCm:29.7, heightCm:21}                     ✓
edit: comment (author Reviewer) on /body/p[3]; then add type:reply on /comment[@id=1]
$ read --view comments  → count:2, comment id=1 with replies:[{id:2, parentId:1,
                          author:"Author2", text:"Yes — verified."}]                        ✓ thread
$ render report.docx --to pdf → {format:"pdf", written:…/report.pdf, sizeBytes:24793}
$ head -c5 report.pdf   → 2550 4446 2d  (%PDF-)                                             ✓ magic bytes
```

## M3.4 xlsx — PASS

```
$ dotnet run --project /tmp/aio-bigbook -- large.xlsx 330000   # verbatim port of the Excel
  test helper BigWorkbookGenerator (StreamingTestSupport.cs)   → 43,350,767 bytes (41.3 MB)
$ read large.xlsx --view stats   (AIOFFICE_MAX_FILE_MB unset)
  → {streamed:true, usedRange:"A1:F330000", cells:1980000, formulas:330000} in 2211 ms      ✓ fast path
$ get large.xlsx /Sheet1/E310000 → {value:620000, formula:"=A310000*2", streamed:true}      ✓ deep cell
$ AIOFFICE_MAX_FILE_MB=1 read large.xlsx → ok:false, code:file_too_large,
  suggestion:"Raise or unset AIOFFICE_MAX_FILE_MB (it is an opt-in cap; default is unlimited)…" ✓ opt-in cap
metrics.xlsx: 5×3 quarter table + numeric X/Y block, then one atomic batch:
  add chart kind:scatter dataRange:E1:F4 · add type:name path:/Sheet1/B2:C5 name:SalesData
  set /Sheet1/D6 value:"=SUM(SalesData)" · freezeRows:1 · autoFilter on A1:C5 · printArea:A1:F10
$ get /Sheet1/D6        → {value:150, formula:"=SUM(SalesData)", cachedValue:150}           ✓ name evaluates
$ get /name[@name=SalesData] → {refersTo:"Sheet1!$B$2:$C$5", ranges:["/Sheet1/B2:C5"]}      ✓
$ get /Sheet1/chart[1]  → {chartKind:"scatter", title:"Growth", dataRange:"E1:F4"}           ✓
$ get /Sheet1           → freezeRows:1 · autoFilter:"A1:C5" · pageSetup.printArea:"A1:F10"  ✓
$ validate metrics.xlsx → valid:true, errors:0                                              ✓
(NOTE: first attempt used `add chart` with a RANGE path → typed invalid_args
 "add chart targets a sheet path like /Sheet1; the data location goes in props.dataRange"
 and the whole 6-op batch wrote NOTHING — atomicity held.)
```

## M3.5 pptx — PASS

```
deck.pptx: slide 2 added · literal bar chart (2 series) on /slide[1] · fade transition 0.5s
           ellipse + arrow + line shapes on /slide[2]
edit: {"op":"add","type":"chart","props":{kind:"bar", dataFrom:"metrics.xlsx!Sheet1/A1:C5"}}
$ get /slide[2]/chart[1] → title "From metrics.xlsx" · categories [Q1,Q2,Q3,Q4]
                           series [(North,[10,15,12,18]), (South,[20,25,22,28])]            ✓ dataFrom
z-order: move shape[@id=3] front + shape[@id=4] back → ok                                   ✓
$ get /slide[1]          → {transition:"fade", transitionDuration:"0.5s"}                   ✓
$ render --to svg --scope /slide[2]
  → 11 <rect> (mini chart bars), <ellipse>, <polygon> (arrow), <line>, 5× data-aio-path     ✓ mini-chart
$ render deck.pptx --to pdf → {format:"pdf", written:…/deck.pdf, sizeBytes:35874, pages:2}
  head -c5 → %PDF- · pdf page tree /Count 2 (one page per slide)                            ✓ multi-page
$ validate deck.pptx     → valid:true, count:0                                              ✓
Every chart add carries the honest warning: chart_data_not_editable (literal caches,
no embedded workbook — "Edit Data" in PowerPoint will prompt; planned M4).
```

## M3.6 MCP — PASS

```
$ aioffice mcp   (stdio JSON-RPC, driven by a python client)
tools/list → 14 tools: office_create office_read office_query office_get office_edit
  office_render office_validate office_template file_snapshot office_status office_help
  office_schema preview_open preview_selection                                             ✓ still 14
tools/call office_edit {file:"deck.pptx", ops:[{op:add, type:chart,
  props:{kind:"line", dataFrom:"metrics.xlsx!Sheet1/A1:B5"}}]} → ok:true, applied:1         ✓ dataFrom over MCP
tools/call office_get /slide[1]/chart[2] → title "MCP dataFrom", series [North], cats 4     ✓
tools/call office_render {to:"pdf"} → ok:true, pages:2, written:…/deck-mcp.pdf              ✓ pdf over MCP
Token budget test: tool surface ≈ 1,852 tokens of the 3,500 budget (pdf + dataFrom
descriptions added with room to spare).
```

## M3.7 Manual-check fixtures regenerated (for a human in real Office)

- `word-sample.docx` — + numbered/nested/bulleted lists, external+anchor links over a bookmark, a real footnote, a threaded comment reply, landscape A4 pages (worker-updated generator). validate: 0.
- `excel-sample.xlsx` — + "Metrics" sheet: scatter chart, defined name `GrowthYs` powering a live `=SUM` (cached 32.6 asserted in the raw part), frozen header row, AutoFilter, print area (generator extended this round). validate: 0.
- `pptx-showcase.pptx` — + native bar chart, fade transition 0.7s, ellipse/diamond/arrow geometries, flipped line, z-order exercise (worker-updated generator). validate: 0.

## M3.8 Published binary — PASS

```
$ dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release -p:PublishSingleFile=true --self-contained -o dist/osx-arm64
$ ls -l dist/osx-arm64/aioffice → 37,737,001 bytes (36.0 MiB; 0.3.0 was 37,667,545 — M3 adds ≈68 KB)
$ aioffice doctor               → version 0.4.0 · limits {maxFileMb:"unlimited", maxFileMbDefault:"unlimited"} · 3 handlers ready
$ create m.xlsx → edit 3×3 values → create d.pptx → edit add chart dataFrom:"m.xlsx!Sheet1/A1:C3"
  → get /slide[1]/chart[1] → series [(North,[10,15]),(South,[20,25])] → validate valid:true  ✓ chart-from-xlsx loop
$ render d.pptx --to pdf        → {pages:1, sizeBytes:16633}, magic %PDF-                    ✓
$ AIOFFICE_MAX_FILE_MB=1 read big.docx (2 MB sparse) → file_too_large                        ✓ opt-in cap
```

## M3.9 Honest limits introduced/kept in M3

- pptx chart data is a **literal cache** (c:strLit/c:numLit, no embedded workbook): PowerPoint renders it fine, but "Edit Data" prompts to create a workbook. Every projection says `dataEditable:false` and chart creation attaches a `chart_data_not_editable` warning. Embedded chart workbooks → M4 seed.
- `dataFrom` is one-shot: the chart holds a snapshot of the workbook values, not a live link.
- xlsx streaming is **read-only** (stats/text + cell/range get). Edits on huge workbooks still load the full ClosedXML DOM — they work, but slowly. Write-streaming → M4 seed.
- docx footnotes shipped; **endnotes** did not (typed `unsupported_feature`) → M4 seed.
- Writing **tracked formatting** is still rejected with the workaround named; M3 only resolves (accept/reject) existing w:rPrChange/w:pPrChange revisions.
- docx page setup edits the existing `/section[1..n]`; inserting NEW section breaks → M4 seed.
- PDF page count is reported for pptx (slides) only; docx/xlsx pagination is decided by the browser at print time and not echoed back.
- M0 gap 6 (validate envelope shape drift: docx/pptx say `count`, xlsx says `errors/warnings`) — **still open**, observed again in this very smoke; carry to M4.

---

# AIOffice 0.5.0 (M4) — Integration Smoke Report

Date: 2026-06-12 · Machine: macOS 26.3.0 arm64 · dotnet 10.0.300 (TFM net10.0)
All commands below were actually executed (`dotnet run --project src/AIOffice.Cli --`, temp workspace `/tmp/aioffice-m4-smoke/ws`); outputs are trimmed but real.

## M4.1 Build — PASS

```
$ dotnet build AIOffice.sln -warnaserror
0 个警告 0 个错误
```

No integration drift this round: the parallel M4 feature work (Word toc/watermark/endnotes/sections-break/replace, Excel bulk/rows-cols/notes/replace, Pptx animations/comments/chart-workbook/replace) built and tested green as handed over. Integrator additions: `ReplaceSugar` (shared CLI/MCP document-wide replace expansion + aggregation), the `--find/--replace` CLI sugar, the `"/"` path carve-out for replace ops in `EditOp.ParseBatch`, schema/help updates, and `.github/workflows/release.yml`.

## M4.2 Tests — PASS (871/871 across 7 projects)

```
$ dotnet test AIOffice.sln
AIOffice.Core.Tests     112 passed
AIOffice.Word.Tests     238 passed
AIOffice.Pptx.Tests     200 passed
AIOffice.Excel.Tests    216 passed
AIOffice.Mcp.Tests       50 passed   (+6 ReplaceSugar: per-format "/" expansion, aggregation, tracked body-only, no-match warning collapse, root-path rejected for non-replace ops)
AIOffice.Preview.Tests   24 passed
AIOffice.Render.Tests    31 passed
```

Token budget test still green: the 14-tool MCP surface (now documenting the replace op + 6 new add types) stays ≤ 3500 tokens.

## M4.3 docx smoke — PASS

```
$ aioffice create report.docx + 6 headings (Overview/Goals/Risks/Mitigations/Timeline/Appendix)
$ edit --add toc (levels 1-3, before /body/p[1])
  → {"op":"add","type":"toc","path":"/toc[1]","levels":"1-3","entries":6}
  → warning toc_pages_unknown (pagination needs Word; honest, not estimated)
$ get /toc[1] → {"levels":"1-3","entryCount":6,"title":"Contents"}                            ✓ 6 entries
$ read --view structure → TOC SDT + 9 body paragraphs with canonical paths                    ✓
$ edit --add watermark text=DRAFT → {"path":"/watermark[1]","headers":1}; validate → 0 issues ✓
$ edit --add endnote on /body/p[3] → {"path":"/endnote[@id=1]"}                               ✓
$ edit --add sectionBreak (nextPage) on /body/p[5] → sections:2
$ set /section[2] orientation=landscape
  → get /section[1] → "orientation":"portrait"   · get /section[2] → "orientation":"landscape","pageSize":"A4"  ✓ mixed-orientation document
$ edit report.docx --find 2025 --replace 2026 --track --author Reviewer
  → {"replacements":3,"locations":["/body/p[3]","/body/p[3]","/body/p[5]"]} (aggregate at data root)
$ read --view revisions → count:6 — three w:del("2025")+w:ins("2026") pairs, author Reviewer  ✓ tracked replace
$ edit --ops '[{"op":"accept","path":"/body"}]' → applied:6; read revisions → count:0         ✓ clean
$ edit --find '20(26)' --replace 'FY$1' --regex
  → expanded to /body + /header[1] (the watermark created a header): replacements:3           ✓ regex + group substitution
$ validate → {"valid":true,"count":0}                                                         ✓
```

## M4.4 xlsx smoke — PASS

```
$ create metrics.xlsx
$ edit --ops @bulk-ops.json   — ONE set op: anchor /Sheet1/A1, values 100×4 (header + 99 rows, col D = "=Bn*Cn" formulas)
  → applied:1 {"applied":["values"]}
$ get /Sheet1/D5  → {"value":45,"formula":"=B5*C5","cachedValue":45}   (B5=6 · C5=7.5)        ✓ cached eval
$ get /Sheet1/D100 → {"value":45,"formula":"=B100*C100","cachedValue":45}                     ✓
$ edit --ops '[{"op":"add","path":"/Sheet1/row[2]","type":"row"}]'
$ get /Sheet1/D6 → formula "=B6*C6" (was D5 "=B5*C5")                                         ✓ references shifted
$ set /Sheet1/col[C] width=18 · /Sheet1/col[E] hidden=true
  → get col[C] → "width":18 · get col[E] → "hidden":true                                      ✓
$ add note on /Sheet1/B2 (author Reviewer) → get B2 → note:{text,author}                      ✓
$ replace in range /Sheet1/A1:A20 find=Item replace=SKU
  → {"replacements":19,"locations":["/Sheet1/A1","/Sheet1/A3",… 19 cell paths]}               ✓ scoped replace
$ edit metrics.xlsx --find SKU --replace Part --match-case
  → "/" expanded to every sheet (/Sheet1): aggregate replacements:19                          ✓ sugar
$ validate → {"valid":true,"errors":0,"warnings":0}                                           ✓
```

Also observed (script bug, kept honest): the first attempt created the workbook with `--title 'M4 Metrics'`, which names the sheet — every `/Sheet1/...` op then failed with typed `invalid_path` + `candidates:["/'M4 Metrics'"]`, and the document-wide sugar correctly expanded to `/'M4 Metrics'` and returned `replacements:0` + ONE collapsed `find_no_match` warning ("matched nothing in any of the 1 document scope(s)").

## M4.5 pptx smoke — PASS

```
$ create deck.pptx + slide 2
$ add chart (bar, 3 categories × 2 series literals) on /slide[1]
$ get /slide[1]/chart[1] → "dataEditable":true                                                ✓ new charts embed a workbook
$ python: strip c:externalData + embedded part + rel (simulate an M3-era cached-only chart)
$ get → "dataEditable":false
$ edit --ops '[{"op":"set","path":"/slide[1]/chart[1]","props":{"embedData":true}}]'
$ get → "dataEditable":true, categories ["Jan","Feb","Mar"] intact                            ✓ retrofit
$ two shapes on slide 2 → add animation fade(0.5s) on shape[@id=3], flyIn(left) on shape[@id=4]
$ read --view structure → slide[2].animations:
  [{"path":"/slide[2]/animation[1]","effect":"fade","duration":"0.5s","target":"…shape[@id=3]"},
   {"path":"/slide[2]/animation[2]","effect":"flyIn","direction":"left","target":"…shape[@id=4]"}]  ✓ in play order
$ add comment on /slide[1] (author Dana) → "/slide[1]/comment[@id=1]"                         ✓
$ set /slide[2]/notes "Q3 talking points: emphasise Q3 growth"
$ edit deck.pptx --find Q3 --replace Q4
  → expanded to /slide[1] + /slide[2] (notes included by default):
    {"replacements":3,"locations":["/slide[2]/shape[@id=3]/p[1]","/slide[2]/notes"]}          ✓ deck-wide incl. notes
$ get /slide[2]/notes → "Q4 talking points: emphasise Q4 growth"                              ✓
$ unzip -l deck.pptx | grep embeddings → ppt/slides/charts/embeddings/package.bin             ✓ embedded xlsx part present
$ validate deck.pptx / report.docx / metrics.xlsx → all valid, 0 errors                       ✓
```

## M4.6 MCP wire — PASS

```
$ aioffice mcp --workspace /tmp/aioffice-m4-smoke/ws   (JSON-RPC over stdio, python driver)
initialize → serverInfo {"name":"aioffice","version":"0.5.0"}
tools/list → 14 tools (unchanged set)                                                          ✓
tools/call office_edit {"file":"report.docx","ops":[{"op":"replace","path":"/","props":{"find":"FY26","replace":"FY27"}}]}
  → envelope ok:true, aggregate {"replacements":3,"locations":["/body/p[3]","/body/p[3]","/body/p[5]"]}  ✓ replace over the wire
```

## M4.7 Manual-check fixtures regenerated (for a human in real Office)

- `word-sample.docx` — + auto-generated **Contents** (TOC, 8 entries, expect Word to fill page numbers on open/F9), **DRAFT watermark** in the header, an **endnote**, and a second **landscape** section (section 1 portrait / section 2 landscape — check the page turn). validate: 0.
- `excel-sample.xlsx` — + "Bulk" sheet written by ONE bulk op (Region/Units/Price/Revenue + formula column + `=SUM` total, cached 163.9), a **cell note** on B2, **column E hidden**, column D widened. validate: 0.
- `pptx-showcase.pptx` — slide 4 chart is now **Edit-Data-able** (right-click → Edit Data must open the embedded workbook — please verify in PowerPoint); slide 5 has a **fade** box, a **flyIn-after-previous** box, and a slide **comment** pointing at the chart check. validate: 0.

## M4.8 Published binary — PASS

```
$ dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release -p:PublishSingleFile=true --self-contained -o dist/osx-arm64
$ ls -l dist/osx-arm64/aioffice → 37,792,009 bytes (36.0 MiB; 0.4.0 was 37,737,001 — M4 adds ≈54 KB)
$ aioffice doctor → version 0.5.0 · 3 handlers ready
$ TOC loop: create → 2 headings → add toc → get /toc[1] → entryCount:2                        ✓
$ chart loop: create d.pptx → add bar chart → get → dataEditable:true → validate valid:true   ✓
$ replace loop: edit r.docx --find One --replace Uno → replacements:2 → validate valid:true   ✓
```

Release automation landed with this milestone: pushing a `v*` tag runs `.github/workflows/release.yml` (build -warnaserror → full test → publish all 6 rids single-file self-contained → `SHA256SUMS` → `gh release create` with notes rendered from `scripts/release-notes-template.md`). Not exercised end-to-end in this smoke (needs a pushed tag); the steps mirror the manual v0.4.0 commands verbatim.

## M4.9 Honest limits introduced/kept in M4

- **Tracked replace is body-only** (handler contract: headers/footers cannot carry these revision marks here). The document-wide sugar under `--track` therefore expands to `/body` only — header/footer text is silently out of scope for a *tracked* replace; run untracked to cover them.
- Replace locations are paragraph-level canonical paths, one entry **per match** (capped at 20). Matches inside TOC/SDT content replace correctly but may report the scope path (`/body`) instead of a paragraph path — observed in the binary smoke when "One" also matched its own TOC entry.
- TOC page numbers are unknown at write time (`toc_pages_unknown` warning); Word computes them on open/F9. TOC styling is minimal (hyperlinked entries, no leader dots).
- xlsx replace matches **text cells only** (numbers/booleans/dates never match; display text is a formatting concern); formula text requires opting in with `inFormulas:true`, and a replacement that destroys the leading `=` turns the cell into literal text.
- pptx animations are **entrance presets only** (appear/fade/flyIn/wipe); emphasis/exit/motion-paths remain deep water (M5 seed: preset expansion). Comments are **legacy** SlideCommentsPart (PowerPoint shows them fine); modern p188 threaded comments and replies → M5 seeds.
- The document-wide `"/"` scope is a replace-only carve-out in `EditOp.ParseBatch`; `"/"` still does not resolve for get/set ops (DocPath requires ≥1 segment — the office_get schema's "/" mention predates this and remains aspirational).
- xlsx in-place streaming **writes** for big existing sheets still load the ClosedXML DOM (M5 seed); only bulk writes into blank sheets stream.
- M0 gap 6 (validate envelope drift: docx/pptx say `count`, xlsx says `errors/warnings`) — **still open**, observed again in M4.5; carry to M5.

---

# AIOffice 0.6.0 (M5) — Integration Smoke Report

Date: 2026-06-12 · Machine: macOS 26.3.0 arm64 · dotnet 10.0.300 (TFM net10.0)
All commands below were actually executed; outputs are trimmed but real.
Theme: the **markdown/csv bridge** — `create --from` + `read --view markdown|csv` — plus deep tables, fields, header variants, data validation, sparklines, threaded comments, hyperlinks, pptx tables, emphasis/exit animations, comment replies, SmartArt read.

## M5.1 Build — PASS

```
$ dotnet build AIOffice.sln -warnaserror
已成功生成。 0 个警告 0 个错误
```

No integration drift this round: the parallel M5 feature work (Word markdown import/export + deep tables + fields + header variants, Excel csv + dataValidation + sparklines + threaded comments + hyperlinks, Pptx tables + animation expansion + replies + SmartArt) built and tested green as handed over. Integrator additions: `IFormatHandler.CreateFrom` default interface hook (additive; non-importing formats refuse with `unsupported_feature`), `AIOffice.Mcp.Bridge` (shared CLI/MCP `--from` routing + import matrix + bridge-view guard), CLI `--from`/`--sheet` plumbing, MCP `office_create.from` + `office_read` view/sheet schema, schema/help refresh, Markdig in THIRD-PARTY-NOTICES.

One honest correction along the way: Markdig's license is **BSD-2-Clause**, not MIT as the milestone brief said — the notices entry records the real license (verified from the nuspec).

## M5.2 Tests — PASS (1053/1053 across 7 projects)

```
$ dotnet test AIOffice.sln
AIOffice.Core.Tests     112 passed
AIOffice.Word.Tests     300 passed   (+62: markdown bridge round-trips, deep tables, fields, header/footer variants)
AIOffice.Excel.Tests    260 passed   (+44: csv bridge, dataValidation, sparklines, threaded comments, hyperlinks)
AIOffice.Pptx.Tests     264 passed   (+64: tables, emphasis/exit animations, comment replies, SmartArt)
AIOffice.Mcp.Tests       62 passed   (+12 BridgeTests: --from routing + matrix errors, sandbox/file_not_found on source,
                                      default-hook unsupported_feature, bridge-view guard, create-from + dataValidation
                                      + csv view over a REAL in-process MCP wire; tools/list still 14)
AIOffice.Preview.Tests   24 passed
AIOffice.Render.Tests    31 passed
```

Token budget test still green: the 14-tool surface (now with `from`, markdown/csv views, `sheet`, and the field/dataValidation/sparkline/reply type vocabulary) measures ≈ 2082 tokens of the 3500 budget — no assertion change needed.

## M5.3 docx smoke (temp workspace /tmp/aioffice-m5-smoke) — PASS

```
$ write notes.md  (h1 + bold/italic/inline-code/link + nested bullets + two-level numbered list + pipe table + code block)
$ aioffice create report.docx --from notes.md
  → {"created":"…/report.docx","kind":"docx","source":"…/notes.md","format":"markdown"}        ✓
$ read --view outline  → "Q2 Operations Report" + Highlights/Action items/Numbers headings      ✓
$ read --view markdown → full GFM back: # h1, ## h2, **bold**, [docs link](…), "- one",
  4-space nested sub-bullets, "1. Ship M5" + nested "1. word-sample.docx", | Region | Q1 | Q2 |,
  fenced code block, final paragraph                                                            ✓ structure round-trips
$ validate → {"valid":true,"count":0}                                                           ✓
$ edit: add 4x4 table → set table[2] borders=all headerRow=true columnWidths=[4cm,3cm,3cm,3cm]
        set tr[1]/tc[1] mergeRight:2 · tr[2]/tc[1] mergeDown:2
$ render --to html → <td … colspan="2">Spanning header</td> · <td … rowspan="2">Tall cell</td>  ✓ real spans
$ validate → valid:true                                                                         ✓
$ edit: add /header[firstPage] "Q2 Operations — cover" + footer fields pageNumber/numPages
$ get /header[firstPage] → {"variant":"firstPage","text":"Q2 Operations — cover"}               ✓
$ get /footer[1]/p[1]    → {"text":"Page 1 of 1","fields":["pageNumber","numPages"]}            ✓ Page X of Y
$ read --view structure → headers/footers/tables all listed with canonical paths · validate → 0 ✓
```

## M5.4 xlsx smoke — PASS

```
$ orders.csv: header + 007,"Acme, Inc.",2026-06-12,1234.5 + 042,Globex,2026-06-13,99
$ aioffice create orders.xlsx --from orders.csv
  → {"kind":"xlsx","sheet":"Sheet1","delimiter":",","rows":3,"columns":4,"cells":12}            ✓
$ get /Sheet1/A2 → {"value":"007","type":"text"}                                                ✓ leading zero kept
$ read --view csv → code,customer,date,total\r\n007,"Acme, Inc.",2026-06-12,1234.5\r\n042,…     ✓ semantically equal (RFC 4180 CRLF, quotes back)
$ edit: dataValidation list ["Open","Paid","Refunded"] on E2:E10 + wholeNumber between 0..100000 on D2:D10
$ read --view structure → dataValidations: [{path:/Sheet1/dataValidation[1],dvKind:list,values:[…]},
                                            {…dataValidation[2],dvKind:wholeNumber,operator:between}]  ✓ 2 rules
$ validate → valid:true                                                                         ✓
$ add comment on /Sheet1/B2 (author Reviewer) → path /Sheet1/comment[@id=B1B71570-…]            ✓ thread root (GUID id)
$ add reply on that path (author Author) → {"replies":1}
$ read --view comments → one thread: root "Is this the right customer?" + reply
  "Yes — confirmed with sales." nested under replies                                            ✓ thread shape
$ add sparkline /Sheet1/F2 (column, D2:D3) → /Sheet1/sparkline[1] · validate → valid:true       ✓
$ edit: set G1 hyperlink=https://example.com/docs (+tooltip) · G2 hyperlink=#Notes!A1 (internal, sheet added in same batch)
$ get G1/G2 → both hyperlinks read back · validate → valid:true                                 ✓
```

## M5.5 pptx smoke — PASS

```
$ create deck.pptx --title "M5 deck"
$ add table 3x4 headerRow style=medium on /slide[1] → /slide[1]/table[1]
  set tr[1]/tc[1] {"text":"Merged header","mergeRight":2}   (pptx merge counts = cells to ABSORB)
$ render --to svg --scope /slide[1] → 10 elements carry data-aio-path="/slide[1]/table[1]…"     ✓ grid drawn
$ add animation pulse + fadeOut(0.5s) on the title shape
$ read --view structure → animations in play order:
  [{path:/slide[1]/animation[1],class:"emphasis",effect:"pulse"},
   {path:/slide[1]/animation[2],class:"exit",effect:"fadeOut",duration:"0.5s"}]                 ✓
$ add comment (Reviewer) → /slide[1]/comment[@id=1]; add reply on it (Author) → comment[@id=2]
$ read --view comments → thread: "Tighten the title" + nested reply "Done." with parentId:1     ✓ p15 threading
$ validate → valid:true                                                                         ✓
```

## M5.6 MCP wire — PASS

```
$ aioffice mcp --workspace /tmp/aioffice-m5-smoke   (JSON-RPC over stdio, python driver)
initialize → serverInfo {"name":"aioffice","version":"0.6.0"}
tools/list → 14 tools (unchanged count); office_create.inputSchema now has "from"                ✓
tools/call office_create {"file":"wire.xlsx","from":"orders.csv"}
  → ok:true {"kind":"xlsx","sheet":"Sheet1","delimiter":",","rows":3,"columns":4,"cells":12}     ✓ import over the wire
tools/call office_edit {"file":"wire.xlsx","ops":[{"op":"add","path":"/Sheet1/D2:D10",
  "type":"dataValidation","props":{"kind":"list","values":["Open","Paid"]}}]}
  → ok:true [{"path":"/Sheet1/dataValidation[1]","dvKind":"list","range":"D2:D10"}]              ✓ dataValidation over the wire
tools/call office_read {"file":"wire.xlsx","view":"csv"} → first line "sku,qty"                  ✓ csv view over the wire
```

## M5.7 Manual-check fixtures regenerated (for a human in real Office)

- `report.docx` (+ its source `notes.md`) — created **from markdown** (`create --from notes.md`), then a second deep table (merged 4-col header, mergeDown cell with shading, full borders, column widths), a **first-page-only header**, an all-pages header, and a **"Page X of Y"** field footer. validate: 0. Check: heading styles from markdown, the table spans, and that page 2 shows the regular header while page 1 shows the cover header.
- `orders.xlsx` (+ its source `orders.csv`) — imported **from csv** (leading-zero `007`/`042`/`113` codes must show as text, left-aligned), **status dropdown** on H2:H10 (please check the dropdown arrows appear in Excel), wholeNumber rule on the quarters, a **line/column/winLoss sparkline** per row in column I, a **threaded comment + reply** on B2, and a hyperlink in J1. validate: 0.
- `deck.pptx` — **native 3x4 table** with a merged 3-column header and the *medium* style (please check the table looks right in PowerPoint — header fill + banded rows), a **pulse** emphasis shape and a **fadeOut** exit shape (run the slideshow to see both), and a **comment thread** (root + reply). validate: 0.

## M5.8 Published binary — PASS

```
$ dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release -p:PublishSingleFile=true --self-contained -o dist/osx-arm64
$ ls -l dist/osx-arm64/aioffice → 38,067,433 bytes (36.3 MiB; 0.5.0 was 37,792,009 — M5 adds ≈269 KB, mostly Markdig)
$ aioffice doctor → version 0.6.0 · handlers docx/xlsx/pptx all ready
$ md loop:  notes.md → create loop.docx --from notes.md → read --view markdown
  → "# Binary Smoke" + "## Loop" + "**one**" + "| a | b |" all round-trip → validate valid:true ✓
$ csv loop: orders.csv ("007","A, B",12.5) → create loop.xlsx --from orders.csv
  → get A2 {"value":"007","type":"text"} → read --view csv → 007,"A, B",12.5 → validate true    ✓
```

## M5.9 Honest limits introduced/kept in M5

- **Merge-count semantics differ by format** (each side pinned by its own tests): docx `mergeRight/mergeDown` is the resulting TOTAL span (1 = unmerge); pptx counts the cells to ABSORB (`mergeRight:1` → span 2). `office_help docx/table` and `pptx/table` both spell this out.
- The markdown bridge round-trips **structure**, not bytes: code-block language hints are dropped on export (` ``` ` without the language), inline code in the HTML render is plain text, and raw HTML blocks degrade to warnings on import.
- csv import routes only `.csv`/`.tsv` at the command layer (the documented matrix); the Excel handler itself also accepts `.txt` when called directly — kept off the public matrix on purpose.
- `read --view csv` emits CRLF line endings (RFC 4180) regardless of the source file's endings — semantically equal, not byte-equal.
- pptx animation expansion covers 4 emphasis + 3 exit presets; the full upstream preset set (15 emphasis + 16 exit), effect chains and motion paths stay deep water (M6 seed: timeline editing).
- SmartArt is **read-only**; edits answer `unsupported_feature` naming the diagram part. Modern pptx threaded-comment *part* (p188) still unshipped — M5 replies thread classic comments via p15.
- xlsx in-place streaming **writes** for big existing sheets still load the ClosedXML DOM (M6 seed); the csv import reuses the M4 blank-sheet SAX path.
- MCP `office_create {overwrite:true}` on an existing file still hits the handler's "File already exists" guard (pre-existing M0 behavior, applies to `from` too) — delete first or pick a new name; carry to M6.
- M0 gap 6 (validate envelope drift: docx/pptx say `count`, xlsx says `errors/warnings`) — **still open**, observed again in M5.3/M5.4; carry to M6.

---

# AIOffice 0.7.0 (M6) — Integration Smoke Report

Date: 2026-06-13 · Machine: macOS 26.3.0 (Darwin 25.3.0) arm64 · dotnet 10.0.300 (TFM net10.0)
All commands below were actually executed; outputs are trimmed but real. M6 = the deep-water pass:
docx equations (LaTeX → Office Math) / RTL / multi-column · xlsx in-place streaming writes / Excel
Tables / outline grouping · pptx master+layout editing / slide sections / slide size / animation timeline.

## M6.1 Build — PASS

```
$ dotnet build AIOffice.sln -warnaserror
已成功生成。  0 个警告  0 个错误
```

## M6.2 Tests — PASS (1253/1253 across 7 projects)

```
$ dotnet test AIOffice.sln --no-build
Core 124 · MCP 63 · Word 385 · Pptx 327 · Preview 24 · Render 31 · Excel 299   →  1253 passed, 0 failed
```

New M6 tests landed WITH the integration (baseline before M6 was 1237): Core +8 addressing
(ElementSpan + root "/"), Core +1 ParseBatch theory (M6 span/root paths survive the gate), Excel +2
(group span through ParseBatch), Pptx +1 (root section/slideSize through ParseBatch), MCP +1
(equation over the wire).

One assertion changed honestly (justified inline): `ReplaceSugarTests.Root_path_is_rejected_for_non_replace_ops`
now expects `unsupported_feature` instead of `invalid_path` — "/" is a real path form in M6, so a
non-replace docx root op parses and is rejected by the handler (still a hard rejection, nothing written).

## M6.3 docx equations / RTL / multi-column — PASS

```
$ aioffice create report.docx --title "M6 Report"
  → {"created":".../report.docx","kind":"docx"}                                                    ✓
$ aioffice edit report.docx --ops '[{op:add,path:/body/p[1],type:equation,props:{latex:"E = mc^2"}}]'
  → ops[0] {type:equation, path:/body/p[1]/omath[1], display:false, latex:"E = mc^2"}              ✓ inline
$ aioffice edit report.docx --ops '[{...latex:"\frac{-b \pm \sqrt{b^2-4ac}}{2a}", display:true}]'
  → ops[0] {type:equation, path:/body/p[2]/omath[1], display:true}                                 ✓ display block
$ aioffice edit report.docx --ops '[{...latex:"\begin{pmatrix}1&0\\0&1\end{pmatrix}", display:true}]'
  → ops[0] {type:equation, display:true, latex:"\begin{pmatrix}1&0\\0&1\end{pmatrix}"}             ✓ matrix
$ aioffice read report.docx --view text
  → p[1]: "M6 Report$E = mc^2$"   p[2]: "$$\begin{pmatrix}1&0\\0&1\end{pmatrix}$$"
    p[3]: "$$\frac{-b \pm \sqrt{b^2-4ac}}{2a}$$"                                                    ✓ $…$ / $$…$$ markers
$ aioffice get report.docx /body/p[1]/omath[1]
  → {type:equation, properties:{latex:"E = mc^2", display:false}}                                  ✓ stored LaTeX read back
$ aioffice edit report.docx --ops '[{...type:equation,props:{latex:"\foobar{x} + \alpha"}}]'
  → ok:true + meta.warnings:[{code:"equation_partial",
      message:"…not recognized and appear literally: \foobar…"}]                                   ✓ partial degrade, file still valid
$ aioffice edit report.docx --ops '[{op:add,path:/body,type:p,props:{text:"مرحبا بالعالم"}},
                                     {op:set,path:/body/p[5],props:{rtl:true}}]'  → applied:2        ✓
$ aioffice get report.docx /body/p[5]   → {rtl:true, alignment:"right"}                             ✓ RTL paragraph
$ aioffice edit report.docx --ops '[{op:add,path:/body,type:table,props:{rows:2,columns:2}},
                                     {op:set,path:/body/table[1],props:{rtl:true}}]'  → applied:2    ✓
$ aioffice get report.docx /body/table[1]   → properties.rtl = true                                 ✓ RTL table
$ aioffice edit report.docx --ops '[{op:set,path:/section[1],props:{columns:2}}]'  → applied:1       ✓
$ aioffice get report.docx /section[1]   → {columns:2, columnGapCm:1.27}                            ✓ two-column section
$ aioffice validate report.docx   → {valid:true, count:0, issues:[]}                                ✓ 0 errors after all the above
```

## M6.4 xlsx in-place streaming write / Excel Table / outline grouping — PASS

```
# ~54 MB CSV → big.xlsx via the streaming SAX writer
$ aioffice create big.xlsx --from big.csv
  → {rows:330001, columns:6, cells:1980006, streamed:true}  (elapsed ~4.9 s; file 42,181,943 bytes)  ✓
# in-place streaming write into the existing 42 MB workbook (stream=true)
$ aioffice edit big.xlsx --ops '[{op:set,path:/Sheet1/H200000,props:{value:"STREAMED-DEEP"}},
                                  {op:set,path:/Sheet1/J5,props:{values:[["a","b"],["c","d"]]}}]' stream=true
  → {applied:2, ops:[{path:/Sheet1/H200000, streamed:true},{path:/Sheet1/J5, streamed:true}],
     streamed:true}  (elapsed ~6.5 s)                                                              ✓ in-place
$ aioffice get big.xlsx /Sheet1/H200000 stream=true   → {value:"STREAMED-DEEP", streamed:true}      ✓ reopen-verified deep cell
$ aioffice get big.xlsx /Sheet1/K6 stream=true        → {value:"d", streamed:true}                  ✓ reopen-verified bulk cell
# Excel Table "Sales" medium2 with a totals row, then a structured-reference SUM
$ aioffice edit metrics.xlsx --ops '[{op:add,path:/Sheet1/A1:B5,type:table,
      props:{name:"Sales",style:"medium2",totals:{"Amount":"sum"}}}]'
  → {name:"Sales", style:"TableStyleMedium2", totalsRow:true, totals:{Amount:"sum"}}               ✓
$ aioffice edit metrics.xlsx --ops '[{op:set,path:/Sheet1/D1,props:{value:"=SUM(Sales[Amount])"}}]'
$ aioffice get metrics.xlsx /Sheet1/D1
  → {value:825, type:number, formula:"=SUM(Sales[Amount])", cachedValue:825}                        ✓ 100+250+175+300=825 evaluated
# outline grouping over a row span — through the REAL CLI (was blocked at the Core gate before M6)
$ aioffice edit metrics.xlsx --ops '[{op:add,path:/Sheet1/row[2]:row[6],type:group,props:{collapsed:true}}]'
  → {type:group, axis:row, from:"2", to:"6", collapsed:true, outlineLevel:1}                        ✓
$ aioffice get metrics.xlsx /Sheet1/row[3]   → {hidden:true, outlineLevel:1}                        ✓ structure shows outline
$ aioffice validate metrics.xlsx   → {valid:true, errors:0, warnings:0}                             ✓ 0 errors
```

## M6.5 pptx master/layout / sections / slide size / animation timeline — PASS

```
$ aioffice edit deck.pptx --ops '[{op:set,path:/master[1],props:{background:"0F172A",accent1:"38BDF8"}}]'  ✓ master bg + accent
$ aioffice edit deck.pptx --ops '[{op:add,path:/master[1],type:layout,props:{name:"M6 Custom",basedOn:1}}]'
  → target:/master[1]/layout[2]                                                                     ✓ cloned layout
$ aioffice get deck.pptx /master[1]
  → {layoutCount:2, layouts:[{name:"Blank"},{name:"M6 Custom", usedBySlides:[]}]}                   ✓
$ aioffice edit deck.pptx --ops '[{op:add,path:/slide[3],type:slide,props:{title:"Uses M6 Custom",layout:2}}]'  ✓ slide uses cloned layout
$ aioffice edit deck.pptx --ops '[{op:add,path:/,type:section,props:{name:"Intro",afterSlide:0}},
                                   {op:add,path:/,type:section,props:{name:"Body",afterSlide:1}}]'  → applied:2   ✓ (path "/" — was blocked pre-M6)
$ aioffice read deck.pptx --view outline   → sections:[{name:"Intro",slides:[/slide[1]]},
                                                        {name:"Body",slides:[/slide[2],/slide[3],/slide[4]]}]    ✓ outline groups by section
$ aioffice edit deck.pptx --ops '[{op:set,path:/,props:{slideSize:"4:3"}}]'   → applied:1            ✓
$ aioffice get deck.pptx /   → {slideSize:"4:3", widthCm:25.4, heightCm:19.05, slideCount:4, sectionCount:2}     ✓ dims
$ aioffice edit deck.pptx --ops '[{...animation fade click},{...animation pulse afterPrevious}]'  → applied:2
  structure animations → ["fade","pulse"]
$ aioffice edit deck.pptx --ops '[{op:move,path:/slide[1]/animation[2],position:"before /slide[1]/animation[1]"}]'
  structure animations → ["pulse","fade"]                                                          ✓ timeline reorder
$ aioffice validate deck.pptx   → {valid:true, count:0, issues:[]}                                  ✓ 0 errors
```

## M6.6 MCP over the wire — PASS

```
$ tools/list  → 14 tools (unchanged; M6 added op types + addressing forms, never tools)            ✓
$ office_edit {file:eq.docx, ops:[{op:add,path:/body/p[1],type:equation,props:{latex:"E = mc^2"}}]}
  → ops[0].path = /body/p[1]/omath[1]                                                              ✓
$ office_get {file:eq.docx, path:/body/p[1]/omath[1]}  → properties.latex = "E = mc^2"             ✓
```

Schema token budget after the M6 additions: ~2214 tokens (budget 3500) — the `office_edit` type
enum gained equation/columnBreak/section/layout/group, the `office_get` path hint gained the new
forms, all within the reserved wording headroom; the budget test stays green untouched.

## M6.7 Rendered equation (dogfood) — PASS

```
$ aioffice create quad.docx --title "Quadratic Formula"
$ aioffice edit quad.docx --ops '[{op:set,path:/body/p[1],props:{text:"The quadratic formula:"}},
      {op:add,path:/body/p[1],type:equation,props:{latex:"x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}",display:true}}]'
$ aioffice render quad.docx --to png -o quad.png   → {format:png, sizeBytes:12312}                  ✓
  Visually verified: x = (-b ± √(b²-4ac)) / 2a renders with a real fraction bar, radical,
  superscript and ± glyph. Embedded at assets/demo/equation-quadratic.png (README, both languages).
```

## M6.8 Published binary — PASS

```
$ dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release -p:PublishSingleFile=true -p:SelfContained=true -o dist/osx-arm64
  → dist/osx-arm64/aioffice  38,131,657 bytes (36 MB)
$ aioffice doctor   → {version:"0.7.0", runtime:".NET 10.0.8", ok:true}                             ✓ 0.7.0
$ aioffice edit b.docx … type:equation latex:"E = mc^2" display:true → get omath → "E = mc^2"; validate 0   ✓ equation loop
$ aioffice edit b.xlsx … type:table name:"T" style:medium2 totals:{Amount:sum};
  set D1 "=SUM(T[Amount])" → get D1 = 30                                                            ✓ Excel Table loop
$ aioffice edit b.pptx … set /master[1] {background,accent1}; add type:layout "Bin Layout";
  get /master[1] → layoutCount:2 ["Blank","Bin Layout"]; validate 0                                ✓ master-edit loop
```

## M6.9 Manual-check fixtures refreshed (open these in real Office)

`fixtures/manual-check/` regenerated with M6 content — open in real Word/Excel/PowerPoint to confirm
equations render, sections appear, tables/grouping show:
- `report.docx` — inline `E=mc^2`, the display quadratic formula, a 2×2 matrix, an RTL Arabic
  paragraph + RTL table, a two-column section.
- `metrics.xlsx` — an Excel Table "Sales" (medium2) with a totals row, `=SUM(Sales[Amount])`, and
  rows 2–6 collapsed into an outline group.
- `deck.pptx` — an edited master (dark background + accent), a cloned layout, two sections
  ("Intro"/"Body"), 4:3 slide size, and a reordered animation timeline.

## M6.10 Honest limits introduced/kept in M6

- **Equations are docx-only in M6.** pptx/xlsx equations (OMML in slides/cells) are M7 seeds — the
  same `AIOffice.Word.Equations` converter will be reused. Tracked-changes does NOT record equation
  adds: `add type:equation` with `--track` answers `unsupported_feature` (run it without `--track`).
- **The LaTeX subset is curated, not complete.** Unknown commands degrade to literal runs + an
  `equation_partial` warning (the file always validates); see `aioffice help equations` for the
  supported set. This is deliberate — an honest partial render over a silent failure.
- **RTL is docx-only in M6** (paragraph/run/table). xlsx `sheetView rightToLeft` and pptx shape RTL
  are M7. RTL never reshapes glyphs — it sets the OOXML direction; Word/LibreOffice lay out the text.
- **In-place streaming writes require an all-streamable batch.** Any non-streamable op
  (bold/fill/numberFormat/merge/…) in the batch falls the whole batch back to the ClosedXML DOM path.
  Streamed formula cells carry no cached value (Excel recomputes on open). This is the M5 seed paid off.
- **Excel Table `remove` keeps the cell data** (ClosedXML's `Clear` would erase values) — the
  ListObject is unregistered, the values survive. Documented in `office_help xlsx/table`.
- **The "/" root and `row[a]:row[b]` span are new Core addressing forms** added so M6 features reach
  the real CLI/MCP surface (they were blocked at `EditOp.ParseBatch` → `DocPath.Parse` before).
  docx/xlsx have no document-root edit/get target, so they answer `unsupported_feature` for "/"
  (naming the body/section/sheet workaround) rather than crashing.
- M0 gap 6 (validate envelope drift: docx/pptx say `count`, xlsx says `errors/warnings`) — **still
  open**, observed again in M6.3/M6.4; carry to M7.

---

# M7 (v0.8.0) — audit before you ship, document properties, content controls, named cell styles

## M7.1 Build — PASS

`dotnet build AIOffice.sln -warnaserror` → **0 warnings, 0 errors** across all 9 src + 7 test projects (net10.0, Nullable enable, TreatWarningsAsErrors).
Reconciled Core: a single canonical `IAuditor` + `AuditFinding`/`AuditResult`/`AuditSummary`/`AuditOptions` in `src/AIOffice.Core/Abstractions.cs`; all three handlers implement it. New 15th MCP tool `office_audit`; new CLI `audit` verb; shared `AIOffice.Mcp.AuditVerb` so CLI and MCP emit a byte-identical payload.

## M7.2 Tests — PASS (1366/1366 across 7 projects)

- Core 124 · Word 424 · Excel 335 · Pptx 365 · MCP 63 · Preview 24 · Render 31 = **1366** (up from 1253; the format owners landed audit/properties/content-control/named-style tests WITH the features).
- MCP tool-count expectations updated 14 → 15 (`ServerBootTests`, `SchemaHelpStatusTests`, `BridgeTests`). Token-budget test (`TokenBudgetTests`, ≤ 3500) **still green** with `office_audit` added — the terse description kept the surface under budget, so the 3500 ceiling was **not** raised.
- OpenXmlValidator reports 0 errors after every mutating test (independent oracle, unchanged).

## M7.3 docx audit → --fix → re-audit → validate — PASS

Built a deliberately-bad `report.docx` through the real CLI (`dotnet run --project src/AIOffice.Cli -- --workspace …`): image with no alt, H1 then H3 (skip), an empty H2, a header-less 2×2 table, no document title.

```
$ aioffice audit report.docx
  findings: a11y_no_doc_title(warn,/properties) · a11y_no_alt_text(error,/body/p[5]) ·
            a11y_heading_skip(warn,/body/p[3]) · quality_empty_heading(warn,/body/p[4]) ·
            a11y_no_table_header(error,/body/table[1])
  summary: {errors:2, warnings:3, infos:0}   exit=0   (findings are data, not errors)

$ aioffice audit report.docx --fix
  fixed: 3   remaining: [a11y_heading_skip#/body/p[3], quality_empty_heading#/body/p[4]]
  summary: {errors:0, warnings:2, infos:0}   exit=0
  (alt text, table header row, and doc title set from the first Heading1 "Overview";
   heading-skip + empty-heading are report-only and correctly REMAIN)

$ aioffice validate report.docx → {valid:true, count:0, issues:[]}   exit=0
```

## M7.4 docx document properties (core + typed custom) — PASS

```
$ aioffice edit report.docx --ops '[{"op":"set","path":"/properties","props":{"title":"Q3 Report","author":"Onecer","custom":{"Project":"Aurora","Reviewed":true}}}]'
  → core:["title","author"] custom:["Project","Reviewed"]
$ aioffice read report.docx --view properties
  → core.title="Q3 Report", core.author="Onecer";
    custom.Project="Aurora" (string), custom.Reviewed=true (REAL boolean — typed round-trip, not "true")
```

## M7.5 docx content controls (dropdown) — PASS

```
$ aioffice edit report.docx --ops '[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"dropdown","tag":"status","title":"Status","items":["Draft","Final"]}}]'
  → path "/sdt[@tag=status]" kind dropdown
$ aioffice edit report.docx --ops '[{"op":"set","path":"/sdt[@tag=status]","props":{"text":"Final"}}]'  → value "Final"
$ aioffice read report.docx --view fields
  → [{ path:"/sdt[@tag=status]", kind:"dropdown", tag:"status", title:"Status", value:"Final", items:["Draft","Final"] }]
$ aioffice validate report.docx → {valid:true, count:0}   (w14:checkbox-style sdt parts validate clean)
```

## M7.6 xlsx audit / --fix / named cell style — PASS

Built a bad `metrics.xlsx`: `=B2/0` (→ `#DIV/0!`), a merged data range A2:A3 inside the used region A1:C3, an anchored image with no alt, no doc title.

```
$ aioffice audit metrics.xlsx
  findings: quality_formula_error(error,/Sheet1/C2 → #DIV/0!) · a11y_merged_data_cells(warn,/Sheet1/A2:A3) ·
            a11y_no_alt_text(warn,/Sheet1/image[1]) · a11y_no_doc_title(warn,/properties)
  summary: {errors:1, warnings:3, infos:0}   exit=0

$ aioffice audit metrics.xlsx --fix
  fixed: 2 (alt text + title derived from filename "metrics")
  remaining: [a11y_merged_data_cells#/Sheet1/A2:A3]   (formula-error + merged-cells are report-only)

$ aioffice edit metrics.xlsx --ops '[{"op":"add","path":"/styles","type":"cellStyle","props":{"name":"Currency-Red","numberFormat":"$#,##0.00","color":"FF0000","bold":true}}]'
$ aioffice edit metrics.xlsx --ops '[{"op":"set","path":"/Sheet1/B2:B3","props":{"cellStyle":"Currency-Red"}}]'  → applied:["cellStyle"]
$ aioffice read metrics.xlsx --view styles
  → [{ path:"/style[@name=Currency-Red]", kind:"cellStyle", name:"Currency-Red", numberFormat:"$#,##0.00", bold:true, color:"FF0000" }]
$ aioffice validate metrics.xlsx → {valid:true, errors:0, warnings:1}
  (the 1 warning is the #DIV/0! formula lint — a quality signal, NOT a schema error; intentional)
```

## M7.7 pptx audit / --fix / explicit altText — PASS

Built a bad `deck.pptx`: an off-canvas shape (x=60cm), a 6pt shape, a picture with no alt, no slide title.

```
$ aioffice audit deck.pptx
  findings: a11y_no_slide_title(warn,/slide[1]) · a11y_reading_order(info,/slide[1]) ·
            quality_off_canvas(warn,/slide[1]/shape[@id=2]) · a11y_tiny_font(error,/slide[1]/shape[@id=3], 6pt<8pt) ·
            a11y_no_alt_text(warn,/slide[1]/shape[@id=4])
  summary: {errors:1, warnings:3, infos:1}   exit=0

$ aioffice audit deck.pptx --fix
  fixed: 2 (picture alt text + slide title placeholder)
  remaining: [a11y_reading_order#/slide[1], quality_off_canvas#…shape[@id=2], a11y_tiny_font#…shape[@id=3]]

$ aioffice edit deck.pptx --ops '[{"op":"set","path":"/slide[1]/shape[@id=4]","props":{"altText":"Company logo, red square"}}]'
$ aioffice audit deck.pptx --category accessibility  → the picture alt finding is GONE (off-canvas/tiny-font/reading-order remain)
$ aioffice validate deck.pptx → {valid:true, count:0}
```

## M7.8 Clean file — zero findings — PASS

```
$ aioffice create clean.docx --title "Clean Report"
$ aioffice edit clean.docx --ops '[{"op":"set","path":"/body/p[1]","props":{"text":"Summary","style":"Heading1"}},{"op":"add","path":"/body","type":"p","props":{"text":"Body text here."}}]'
$ aioffice edit clean.docx --ops '[{"op":"set","path":"/properties","props":{"title":"Clean Report"}}]'
$ aioffice audit clean.docx → {findings:[], summary:{errors:0,warnings:0,infos:0}}   exit=0
```

(Honest note: `create --title` writes the title as a Heading1 paragraph, NOT the core property Title — pre-existing M0 docx behavior — so a freshly-created file still trips `a11y_no_doc_title` until you `set /properties {title}`. This is exactly the finding the auditor is supposed to raise; `--fix` would set it from the first heading.)

## M7.9 MCP over the wire — PASS (tools/list == 15, office_audit, office_audit fix:true)

Drove the real `aioffice mcp` stdio server (JSON-RPC: initialize → notifications/initialized → tools/list → tools/call) against a fresh bad docx:

```
TOOL COUNT: 15        HAS office_audit: True
office_audit          → ok:true, isError:None (findings are DATA, not protocol errors),
                        summary {errors:2, warnings:1}, codes [a11y_no_doc_title, a11y_no_alt_text, a11y_no_table_header]
office_audit fix:true → ok:true, fixed:3, remaining:[]
```

## M7.10 Schema / help / version — PASS

```
$ aioffice version → version "0.8.0"
$ aioffice doctor  → version "0.8.0"
$ aioffice help audit → the audit topic (codes + --fix semantics + examples) loads from the embedded HelpTopics/audit.md
$ aioffice schema  → 15 verbs incl. audit (mcpTool office_audit); office_read view enum includes properties/fields/styles
```

## M7.11 Published binary (dist/osx-arm64) — PASS

```
$ dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release -p:PublishSingleFile=true --self-contained -o dist/osx-arm64
  → dist/osx-arm64/aioffice  = 38,187,209 bytes (~36 MB, single self-contained file, no runtime dependency)
$ ./dist/osx-arm64/aioffice doctor   → version "0.8.0"; handlers docx/xlsx/pptx all "ready"
  audit loop (binary): audit → {errors:2,warnings:1} · audit --fix → {fixed:3, remaining:[]} · validate → valid:true
  MCP (binary stdio): tools/list == 15, office_audit over the wire → ok:true
```

## M7.12 Token budget — under ceiling, NOT raised

The MCP tool surface (15 tools incl. `office_audit`) stays well under the 3500-token budget
(`TokenBudgetTests` green; the terse `office_audit` description + the `office_help {topic:"audit"}`
externalization of the code list kept it in budget). The 3500 ceiling was therefore **kept**, not raised.

## M7 Honest notes / known gaps

- **`audit` findings are data, never errors.** Like `validate`, a successful audit is `ok:true` / exit 0
  even with error-severity findings. The only non-zero path is the usual `invalid_args`/`file_not_found`/
  `sandbox_denied`/`format_corrupt` — i.e. "the audit could not run", not "the document has problems".
- **`--fix` is safe-only and never destructive.** It applies placeholder alt text, a marked table header
  row, a doc/slide title (first heading > filename > placeholder), and orphan-bookmark removal. Heading
  skips, low contrast, empty headings, off-canvas shapes, formula errors, broken links/refs, tiny fonts
  and merged data cells are **report-only** and remain in `remaining` after `--fix`. Every `--fix` pass
  snapshots the pre-image first (undoable via `file_snapshot`).
- **`create --title` (docx) still writes a Heading1 paragraph, not the core property Title** (pre-existing
  M0 behavior). So a freshly created docx trips `a11y_no_doc_title` until you `set /properties {title}` —
  which is exactly the finding the auditor should raise. `--fix` derives the title from the first heading.
- **xlsx `validate` reports `#DIV/0!` as a lint warning, not a schema error.** A file with a formula error
  is still valid OOXML (`valid:true`); the formula-error signal surfaces through `audit`
  (`quality_formula_error`, severity error) and as a `validate` lint warning. Both are intentional.
- **The deck.pptx good fixture keeps one `a11y_reading_order` (info) finding** — a non-defect heuristic
  noting the picture sits visually above the text shape. Info-level findings are signal, not failure.
- M0 gap 6 (validate envelope drift: docx/pptx say `count`, xlsx says `errors/warnings`) — **still open**,
  observed again in M7.3/M7.6; carry to M8.

---

# M8 (v0.9.0) — diff & review, captions/cross-references, slicers, shape hyperlinks

All commands below were run through the real CLI (`dotnet run --project src/AIOffice.Cli --`,
then the published binary) against a fresh temp workspace. Outputs are verbatim; failures stay FAIL.

## M8.1 Build — PASS

```
$ dotnet build AIOffice.sln -warnaserror   → 0 warnings, 0 errors (8 projects + 7 test projects)
```

## M8.2 Tests — PASS (1449/1449 across 7 projects)

```
Core 124 · Word 448 · Excel 363 · Pptx 396 · MCP 63 · Preview 24 · Render 31   = 1449, 0 failures
```

## M8.3 docx two-file diff (detailed) — PASS

base.docx = 3 paras (p1 Heading1). new.docx = copy, then: change p2 text + append a para + restyle p1 → Heading2.

```
$ aioffice diff new.docx base.docx --json
{ "ok": true, "data": {
  "changes": [
    { "kind":"modified","path":"/body/p[1]","before":"Heading1","after":"Heading2","detail":"style" },
    { "kind":"modified","path":"/body/p[2]","before":"Second paragraph.","after":"Second paragraph EDITED.","detail":"text" },
    { "kind":"added","path":"/body/p[4]","detail":"paragraph" }
  ],
  "summary": { "added":1,"removed":0,"modified":2,"moved":0 },
  "baseline": "base.docx", "view": "detailed" } }
```

1 modified text + 1 added + 1 style modified, sorted by (path, kind), exit 0. Deterministic across runs.

## M8.4 diff --view summary + identical files — PASS

```
$ aioffice diff new.docx base.docx --view summary
  changes: [{kind:modified,path:/body/p[1]},{kind:modified,path:/body/p[2]},{kind:added,path:/body/p[4]}]  (before/after trimmed)
  summary: {added:1,removed:0,modified:2,moved:0}

$ aioffice diff identical.docx base.docx   (byte-copy of base)
  ok:true · changes:[] · summary:{0,0,0,0} · exit 0
```

## M8.5 snapshot diff + bad index → candidates — PASS

report.docx edited twice (each edit auto-snapshots the pre-image): snapshot 1 = before edit1, snapshot 2 = before edit2.

```
$ aioffice snapshot list report.docx   → count:2  numbers:[1,2]
$ aioffice diff report.docx --snapshot 2 --json
  ok:true · baseline:"snapshot 2"
  changes: [{kind:modified,path:/body/p[1],before:"Original report line",after:"CHANGED report line",detail:text}]

$ aioffice diff report.docx --snapshot 99   → invalid_args · candidates: ["1","2"]
```

## M8.6 docx captions + cross-references — PASS

```
$ aioffice edit cap.docx --ops '[{op:add,path:/body/p[1],type:caption,props:{label:Figure,text:"Quarterly trend",position:after}}]'  → ok
$ aioffice edit cap.docx --ops '[{op:add,path:/body/p[1],type:crossRef,props:{to:"/caption[@label=Figure][1]",show:labelAndNumber}}]' → ok
$ aioffice get cap.docx '/caption[@label=Figure][1]' --json
  → { type:"caption", properties:{ label:"Figure", number:1, text:"Quarterly trend", bookmark:"_Ref100000000", anchorPath:"/body/p[2]" } }
$ aioffice read cap.docx --view structure   → lists the caption
$ aioffice validate cap.docx   → valid:true, count:0
```

(Integrator fix: the CLI `get` verb pre-parsed the path through `DocPath.Parse`, which rejected the
two-bracket virtual `/caption[@label=…][i]` form BEFORE the handler's own interception — so `get` on a
caption was `invalid_path` on the CLI while `office_get` worked over MCP. Fixed `FileVerbs.Get` to skip
the fail-fast parse for `/caption[`/`/crossRef[` paths, mirroring `office_get` which never pre-parses.
Caught only by this smoke step — a real CLI↔MCP parity bug.)

## M8.7 xlsx two-workbook diff (changed cell + added sheet) — PASS

```
$ aioffice diff wb-new.xlsx wb-base.xlsx --json
  changes: [{kind:modified,path:/Sheet1/B2,before:"100",after:"250"}, {kind:added,path:/Summary,detail:sheet}]
  summary: {added:1,removed:0,modified:1,moved:0}
```

## M8.8 xlsx slicer on a table → validate 0 — PASS

```
$ aioffice edit dash.xlsx --ops '[…seed A1:B3, add type:table name=Sales…]'  → ok
$ aioffice edit dash.xlsx --ops '[{op:add,path:"/Sheet1/table[@name=Sales]",type:slicer,props:{column:"Region",x:"E2"}}]' → ok
$ aioffice get dash.xlsx '/Sheet1/slicer[1]' --json
  → { kind:"slicer", sheet:"Sheet1", name:"Slicer_Region", source:"table", sourceName:"Sales", column:"Region", caption:"Region" }
$ aioffice validate dash.xlsx   → valid:true, errors:0   (slicer parts authored on raw OpenXml validate clean)
```

## M8.9 pptx diff (reordered slide + edited shape) — PASS

3-slide deck; slide 1 (carrying a text box) moved to position 3, then its shape text edited.

```
$ aioffice diff deck-new.pptx deck-base.pptx --json
  changes: [
    {kind:moved,path:/slide[3],detail:"slide reordered 1 -> 3"},
    {kind:modified,path:"/slide[3]/shape[@id=2]",before:"Alpha",after:"Box EDITED",detail:text}
  ]
  summary: {added:0,removed:0,modified:1,moved:1}
```

## M8.10 pptx shape hyperlink (jump-to-slide + url) → get + validate 0 — PASS

```
$ aioffice edit links.pptx --ops '[{op:set,path:"/slide[1]/shape[@id=3]",props:{hyperlink:"#slide:2"}},{op:set,path:"/slide[1]/shape[@id=4]",props:{hyperlink:"https://example.com"}}]' → ok
$ aioffice get links.pptx '/slide[1]/shape[@id=3]'   → properties.hyperlink == "#slide:2"     (canonical jump form)
$ aioffice get links.pptx '/slide[1]/shape[@id=4]'   → properties.hyperlink == "https://example.com"
$ aioffice validate links.pptx   → valid:true, errors:0
```

## M8.11 Guard rails — PASS

```
$ aioffice diff a.docx b.xlsx                 → invalid_args "Cannot diff a .docx against a '.xlsx' file."
$ aioffice diff new.docx base.docx --snapshot 1  → invalid_args "diff takes EITHER a baseline file OR --snapshot N, not both."
$ aioffice diff new.docx                       → invalid_args (no baseline)
$ aioffice diff new.docx base.docx --view fancy  → invalid_args · candidates: ["summary","detailed"]
```

## M8.12 New addressing through the real ParseBatch gate — PASS

Not just unit tests — through `EditOp.ParseBatch` → `DocPath.Parse` via the real CLI `edit --ops`:

```
caption  via --ops  → ok    crossRef via --ops (to "/caption[@label=Table][1]") → ok
slicer   get /Sheet1/slicer[1] via CLI gate → ok    shape hyperlink "#last" via --ops → ok
```

## M8.13 MCP over the wire — PASS (tools/list == 16, office_diff two-file + snapshot)

```
$ initialize → serverInfo.name "aioffice"
$ tools/list → 16 tools, office_diff present
$ office_diff {file:"new.docx", other:"base.docx"}      → ok:true · summary {added:1,modified:2} · baseline "base.docx"
$ office_diff {file:"report.docx", snapshot:2}          → ok:true · baseline "snapshot 2" · the one change
$ office_diff {file:"base.docx", other:"book.xlsx"}     → isError:true · code invalid_args
$ office_diff {file:"new.docx"}  (no baseline)          → code invalid_args
```

## M8.14 schema / help / version — PASS

```
$ aioffice schema diff   → name:"diff", mcpTool:"office_diff", usage "aioffice diff <file> [<otherFile>] [--snapshot N] [--view summary|detailed]"
$ aioffice help diff     → the diff topic (two modes + result shape + snapshot-diff workflow) from embedded HelpTopics/diff.md
$ aioffice schema        → 16 verbs incl. diff; aioffice doctor → version "0.9.0", unaffected
```

## M8.15 Token budget — under ceiling, NOT raised

The MCP tool surface (16 tools incl. `office_diff`) stays under the 3500-token budget
(`TokenBudgetTests` green; terse `office_diff` description + `office_help {topic:"diff"}` externalization
of the diff semantics kept it in budget). The 3500 ceiling was therefore **kept**, not raised.

## M8.16 Published binary (dist/osx-arm64) — PASS

```
$ dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release -p:PublishSingleFile=true --self-contained -o dist/osx-arm64
  → dist/osx-arm64/aioffice = 38,238,969 bytes (~36.5 MB, single self-contained file, no runtime dependency)
$ ./dist/osx-arm64/aioffice doctor   → version "0.9.0"; handlers docx/xlsx/pptx all "ready"
  two-file diff (binary): new.docx vs base.docx → ok:true, summary {modified:1}, baseline "base.docx"
  snapshot diff (binary): rep.docx --snapshot 2 → ok:true, baseline "snapshot 2", the one change
  MCP (binary stdio): tools/list == 16, office_diff over the wire → ok:true
```

## M8 Honest notes / known gaps

- **`diff` changes are data, never errors.** Like `validate`/`audit`, a successful diff is `ok:true` / exit
  0 no matter how large the change set. The only non-zero path is `invalid_args` (no/both baselines, format
  mismatch, bad `--view`, missing snapshot index)/`file_not_found`/`sandbox_denied`/`format_corrupt`.
- **Diff order is deterministic by construction.** `DiffResult.FromChanges` sorts by `(path, kind)` ordinal,
  then `detail`, so the same two documents diff byte-identically on every platform and every run — verified
  by the dedicated determinism test in each format's `DiffTests`. No `Environment.NewLine`, byte-size or
  iteration-order assertions in the new tests (CI-hygiene grep clean).
- **Snapshot diff never touches the file or its ring.** `--snapshot N` copies the snapshot bytes to a
  throwaway temp baseline (same extension, inside the workspace so it passes the sandbox), diffs against it,
  and deletes the temp — `SnapshotStore.Restore` (which would overwrite the original) is deliberately NOT used.
- **CLI↔MCP parity fix (caption get).** The CLI `get` verb's fail-fast `DocPath.Parse` rejected the virtual
  `/caption[@label=…][i]` path before the handler could intercept it; `office_get` over MCP never pre-parses,
  so the two surfaces diverged. Fixed in `FileVerbs.Get` (skip the pre-parse for `/caption[`/`/crossRef[`).
  Found by smoke step M8.6, not by any unit test.
- M0 gap 6 (validate envelope drift: docx/pptx say `count`, xlsx says `errors/warnings`) — **still open**,
  observed again in M8.8 (xlsx `errors:0`) vs M8.6 (docx `count:0`); carry to M9.

---

# AIOffice 0.10.0 (M9) — Integration Smoke Report (the pre-1.0 capstone)

> Scope: the `convert` verb + `office_convert` MCP tool (16 → 17 tools), the
> reconciled `NeutralDoc`/`INeutralConvertible` Core contract (three handlers
> implement it), the command-layer `NeutralMarkdown` serializer, and the 1.0
> hardening pass (`convert` help topic + `doctor`/`office_status` capabilities
> block). All outputs below are real, captured from the binary or `dotnet run`.

## M9.1 Build — PASS

```
$ dotnet build AIOffice.sln -warnaserror   → 0 warnings, 0 errors (all 13 projects)
```

The three owners each added `NeutralDoc`/`INeutralConvertible` to `Core/Abstractions.cs`;
the integrator reconciled them to ONE canonical definition (exactly per the shared contract)
— it merged cleanly because all three wrote the same records, so the working tree already
held the single definition. Clean build with `TreatWarningsAsErrors`.

## M9.2 Tests — PASS (1501 / 1501 across 7 projects)

```
Core 124 · Word 459 · Excel 374 · Pptx 410 · MCP 79 · Preview 24 · Render 31  = 1501
```

+52 over M8's 1449: the three handler `ConvertTests` (Word/Excel/Pptx), the MCP
`ConvertTests` (end-to-end through the real handlers + over the wire) and
`NeutralMarkdownTests`. CI-hygiene grep over the NEW tests is clean — no
`Environment.NewLine`, no exact-byte-size, no unsorted-order assertions (the one
`\r\n` reference is a CRLF→LF normalization before a csv text assert).

## M9.3 The shared Core contract — reconciled, one definition

`Core/Abstractions.cs` carries exactly the shared contract: `NeutralBlockKind`,
`NeutralRun`, `NeutralBlock`, `NeutralDoc`, `ImportResult`, `INeutralConvertible`
(NOT `System.IConvertible`). Each format handler implements it —
`WordHandler.Neutral.cs`, `ExcelHandler.Convert.cs`, `PptxConvert.cs` — the same
"Core interface + per-handler implementation" pattern as M7 `IAuditor` and M8 `IDiffer`.

## M9.4 docx → pptx (a slide per heading, bullets below, a table) — PASS

```
$ aioffice create report.docx --title "Quarterly Report"  → ok
$ aioffice edit report.docx --ops @… (3 Heading1 + bullets under each + a 2×2 table) → applied 12
$ aioffice convert report.docx deck.pptx
  → {"from":"docx","to":"pptx","blocksWritten":10,"dropped":[],"written":".../deck.pptx"}
$ aioffice read deck.pptx --view text   → "Sales / North up 12% / South flat / Costs / … / Outlook / Grow EU"
$ aioffice query deck.pptx "shape:contains('Sales')" → /slide[2]/shape[@id=2]   (real title placeholder)
$ aioffice validate deck.pptx           → valid:true, 0 issues
```

The `convert_lossy` warning is empty here (no charts/animations in the source). The
docx title `Quarterly Report` is itself a Heading1, so it opens its own slide —
4 headings → 4 slides, every bullet present, the table a native pptx table.

## M9.5 docx ↔ md round-trip — PASS

```
$ aioffice convert report.docx report.md
  → report.md: "# Quarterly Report / # Sales / - North up 12% / - South flat / … / | Region | Total | / | North | 100 |"
$ aioffice convert report.md roundtrip.docx
  → read roundtrip.docx --view outline: Heading1 "Quarterly Report" / "Sales" / "Costs" / "Outlook"  (matches)
```

docx↔md reuses the M5 markdown bridge verbatim (the `convert` verb just unifies the
entrypoint); the outline round-trips.

## M9.6 xlsx → docx (a table per sheet, cached values) — PASS

```
$ aioffice create data.xlsx; edit (+sheet Totals, Sheet1 A1:B3 data, Totals/A1 =SUM(Sheet1!B2:B3))
$ aioffice convert data.xlsx data.docx
  → {"from":"xlsx","to":"docx","blocksWritten":4,"dropped":[],"written":".../data.docx"}
$ aioffice read data.docx --view markdown:
    ## Sheet1   | Region | Units | / | North | 10 | / | South | 20 |
    ## Totals   | 30 |          ← the formula crossed as its cached display value
$ aioffice validate data.docx  → valid:true, 0 issues
```

`xlsx → md` and `pptx → md` go through the neutral model + `NeutralMarkdown`
(verified by `Xlsx_to_md_uses_the_neutral_markdown_serializer`): a sheet becomes a
level-2 heading, its used range a pipe table.

## M9.7 pptx → docx (titles + bullets) — PASS

```
$ aioffice convert deck.pptx outline.docx
  → {"from":"pptx","to":"docx","blocksWritten":9,…}
$ read outline.docx --view markdown: "# Quarterly Report / # Sales / - North up 12% / … / | Region | Total |"
```

The pptx → docx trip even recovers the table cells (native pptx table → docx table).

## M9.8 Honest lossiness — convert_lossy names the drop — PASS

```
$ create charted.pptx; edit (+ a bar chart on slide 1)
$ aioffice convert charted.pptx charted.docx
  → ok:true, data.dropped: ["charts (transferred as data is not supported; the chart is not converted)"]
  → meta.warnings: [{ "code":"convert_lossy",
       "message":"Some content did not survive the conversion: charts (…the chart is not converted)." }]
```

The pptx exporter's `ExportNeutral(ctx, out dropped)` overload names export-side
losses (animations/charts/SmartArt/transitions); they fold into the same
`convert_lossy` warning as the import-side `ImportResult.Dropped`.

## M9.9 render route — docx → pdf — PASS

```
$ aioffice convert report.docx report.pdf
  → {"from":"docx","to":"pdf","blocksWritten":0,"written":".../report.pdf"}
$ file report.pdf  → "PDF document, version 1.4, 1 pages" (22,890 bytes, via headless Chrome)
```

`any → pdf/png/svg/html/txt` routes to the render layer (`PdfRenderVerb`/`PngRenderVerb`
for pdf/png, `handler.Render` + write-to-file for svg/html/txt).

## M9.10 Error gates — PASS

```
$ aioffice convert a.docx b.docx   → ok:false, invalid_args, suggestion mentions "use edit"
$ aioffice convert a.docx a.xyz    → ok:false, unsupported_feature
   suggestion: "Supported destinations: office .docx/.xlsx/.pptx, text .md/.markdown/.csv/.tsv, render .pdf…"
$ aioffice convert report.docx deck.pptx (dest already existed) → ok:true + meta.warnings convert_overwrite
```

## M9.11 1.0 hardening — schema sanity + capabilities block — PASS

```
$ aioffice schema   → 18 verbs, every one has summary+usage; convert→office_convert mapped;
                       all op kinds listed in edit usage; errorCodes + exitCodes present
$ aioffice schema convert → name "convert", usage "aioffice convert <src> <dest> [--json]", positionals [src,dest]
$ aioffice help convert   → the convert topic (the (src→dest) matrix + lossy note + examples) from embedded HelpTopics/convert.md
$ aioffice doctor --json  → capabilities { verbs:18, mcpTools:17, formats:[docx,xlsx,pptx],
     convert{sources, contentTargets, renderTargets}, renderTargets, auditCategories } — one-call introspection
```

## M9.12 MCP over the wire — PASS (tools/list == 17, office_convert)

```
$ initialize → serverInfo.name "aioffice"
$ tools/list → 17 tools, office_convert present
$ office_convert {src:"report.docx", dest:"wire.pptx"}  → ok:true · from "docx" → to "pptx" · blocksWritten 10
```

## M9.13 Token budget — under ceiling, NOT raised

The MCP tool surface (17 tools incl. `office_convert`, ~164 tokens) stays under the
3500-token budget (`TokenBudgetTests` green; terse `office_convert` description +
`office_help {topic:"convert"}` externalization of the (src→dest) matrix kept it in
budget — estimated whole-surface total ~3484). The 3500 ceiling was therefore **kept**, not raised.

## M9.14 Published binary (dist/osx-arm64) — PASS

```
$ dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release -p:PublishSingleFile=true --self-contained -o dist/osx-arm64
  → dist/osx-arm64/aioffice = 38,266,089 bytes (~36.5 MB, single self-contained file, no runtime dependency)
$ ./dist/osx-arm64/aioffice doctor → version "0.10.0"; handlers docx/xlsx/pptx all "ready";
     capabilities block present (mcpTools 17, verbs 18, convert targets, render targets)
$ binary convert r.docx d.pptx     → ok:true, from "docx" → to "pptx", blocks 4; validate d.pptx → valid:true
$ binary MCP (stdio): tools/list == 17, office_convert over the wire → ok:true (to "pptx", blocks 4)
```

## M9.15 Fixtures refreshed (manual-check) — see §6

`fixtures/manual-check/` gains a `report.docx` and the `deck.pptx` CONVERTED from it
(open both in real Office to compare the document against the generated deck side by
side), plus an `xlsx-to-docx-sample.docx` converted from a 2-sheet workbook.

## M9 Honest notes / known gaps

- **`convert` is content-transfer, inherently lossy.** Every drop (export-side
  animations/charts/SmartArt/transitions, import-side styling/formulas-as-values/markdown
  colors) is named in `data.dropped` and a single `convert_lossy` meta warning. Nothing is
  silently lost — the smoke confirmed the chart case (M9.8).
- **docx↔md and xlsx↔csv reuse the M5 bridges verbatim.** `convert` only unifies the
  entrypoint for them; every other md pair (pptx↔md, xlsx↔md) and every office↔office pair
  goes through the neutral model + (for md) `NeutralMarkdown`. csv is xlsx-family only:
  `csv ↔ a non-xlsx office format` is `unsupported_feature` ("go via a workbook").
- **The pptx `outline` view shows `title:null` for convert-built decks.** The title placeholder
  IS created (`PlaceholderValues.Title`, confirmed by `query shape:contains('Sales')` →
  `/slide[2]/shape[@id=2]`) and the `text` view shows every title; only the `outline` view's
  title projection doesn't pick it up. Pre-existing outline-view detail, NOT a convert bug —
  content is complete and validates clean. Carry to a future outline-view polish.
- **M0 gap 6 (validate envelope drift: docx/pptx `count` vs xlsx `errors/warnings`)** — still
  open, unchanged by M9; carry to 1.0 hardening proper.

---

# M9 Hardening Fix — Document Core Properties → standard `docProps/core.xml` (VERIFIER pass)

Date: 2026-06-13 · Machine: macOS 26.3.0 (Darwin 25.3.0) arm64 · dotnet 10.0.300 (TFM net10.0)
All commands below were actually executed via `dotnet src/AIOffice.Cli/bin/Debug/net10.0/aioffice.dll`;
outputs are real (trimmed only for width).

## Summary — PASS

The bug: core document properties (title/author/…) were written through `System.IO.Packaging`
`PackageProperties`, which materialized a NON-standard `package/services/metadata/core-properties/{GUID}.psmdcp`
part instead of the conventional OOXML `docProps/core.xml` (the SDK `CoreFilePropertiesPart`). Consequences:
unzip showed nothing at `docProps/core.xml`, handlers were inconsistent, and `convert` dropped the Title.

The fix (all three formats now read/write the standard `docProps/core.xml` part, migrate-on-read from the
legacy façade, and never store via `PackageProperties`):
- docx — `src/AIOffice.Word/WordHandler.CoreProperties.cs` (hand-rolled `cp:coreProperties` reader/writer;
  legacy `.psmdcp` migrated to `docProps/core.xml` on first write).
- pptx — `src/AIOffice.Pptx/PptxCoreProps.cs` (standard part; drops a legacy `.psmdcp` part on write).
- xlsx — `src/AIOffice.Excel/ExcelCoreProperties.cs` (`NormalizeAfterSave` relocates ClosedXML's core part
  from the legacy URI to `docProps/core.xml`, bytes preserved verbatim to keep the round-trip law).

Verifier-added change (requirement #3 — the NeutralMarkdown serializer dropped the title):
- `src/AIOffice.Mcp/NeutralMarkdown.cs` `Write` now emits `NeutralDoc.Title` as a leading `# Title`
  (de-duplicated when the title is already the first H1; `Parse` reads it back) → fixes `xlsx→md` / `pptx→md`.
- `src/AIOffice.Mcp/ConvertVerb.cs` `docx→md` bridge prepends the document Title as a leading `# Title`
  (read from `--view properties`, skipped when the body already opens with that exact heading). The
  general-purpose `read --view markdown` view stays title-free, so its existing contract is untouched.
- Tests landed with the fix: 3 new `NeutralMarkdownTests` (title emitted / not duplicated / absent-when-null)
  plus an updated `Docx_to_md_and_back_preserves_…` assertion (`# Quarterly Field Report` leads the file).

## Build — PASS

```
$ dotnet build AIOffice.sln -warnaserror
已成功生成。  0 个警告  0 个错误
```

## Tests — PASS (1514 across 7 projects, 0 failures)

```
$ dotnet test AIOffice.sln
AIOffice.Core.Tests     124 passed
AIOffice.Word.Tests     463 passed
AIOffice.Pptx.Tests     413 passed
AIOffice.Excel.Tests    377 passed
AIOffice.Mcp.Tests       82 passed   (+3 new NeutralMarkdown title tests)
AIOffice.Render.Tests    31 passed
AIOffice.Preview.Tests   24 passed
-------------------------------------------------
TOTAL                  1514 passed, 0 failed, 0 skipped
```

## Real smoke — the exact scenario that exposed the bug — PASS

### 1) report.docx: set /properties title + author; Heading1 A + 2 bullets + Heading1 B + 1 bullet + table

```
$ unzip -p report.docx docProps/core.xml
<?xml version="1.0" encoding="utf-8" standalone="yes"?><cp:coreProperties
  xmlns:cp="…/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="…/terms/" xmlns:xsi="…/XMLSchema-instance">
  <dc:title>Quarterly Field Report</dc:title><dc:creator>Field Team</dc:creator></cp:coreProperties>

$ unzip -l report.docx | grep -i psmdcp
(none — standard docProps/core.xml part only)
```

The standard part now EXISTS with the `dc:title` element; no legacy psmdcp part.

### 2) convert report.docx → deck.pptx

```
$ aioffice convert report.docx deck.pptx
{"ok":true,"data":{"from":"docx","to":"pptx","blocksWritten":6,"dropped":[],"written":".../deck.pptx"}}

$ aioffice read deck.pptx --view properties
… "core":{"title":"Quarterly Field Report", …}             ← title crossed via docProps/core.xml

$ aioffice read deck.pptx --view outline
slide[1]: Title "A" + Content "A-one A-two"                ← FIRST slide is "A" (NO empty leading slide)
slide[2]: Title "B" + Content "B-one" + Table

$ aioffice validate deck.pptx
{"valid":true,"count":0}
```

### 3) convert report.docx → report.md (includes the title + headings/bullets/table)

```
$ aioffice convert report.docx report.md
$ cat report.md
# Quarterly Field Report      ← document Title, now carried into the markdown

# A

A-one

A-two

# B

B-one

| H1 | H2 |
| --- | --- |
| c1 | c2 |
```

### 4) convert report.md → back.docx (outline headings match)

```
$ aioffice convert report.md back.docx
$ aioffice read back.docx --view outline
headings: [ "Quarterly Field Report" (h1), "A" (h1), "B" (h1) ]   ← title-as-H1 + body headings A,B
$ aioffice validate back.docx → {"valid":true,"count":0}
```

### 5) Properties round-trip on all three formats (set title+author+custom, reopen, read back, core.xml holds title)

```
props.docx  read --view properties → title "Round Trip docx", author "QA Bot", custom {Project:Acme, Reviewed:true}
            docProps/core.xml: <dc:title>Round Trip docx</dc:title><dc:creator>QA Bot</dc:creator>   (no psmdcp)  validate 0
props.xlsx  read --view properties → title "Round Trip xlsx", author "QA Bot", custom {Project:Acme, Reviewed:true}
            docProps/core.xml: <dc:title>Round Trip xlsx</dc:title><dc:creator>QA Bot</dc:creator>   (no psmdcp)  validate 0
            (xlsx core.xml uses the default-namespace coreProperties form — ClosedXML's verbatim
             serialization relocated to the standard URI; schema-valid and Office-readable.)
props.pptx  read --view properties → title "Round Trip pptx", author "QA Bot", custom {Project:Acme, Reviewed:true}
            docProps/core.xml: <dc:title>Round Trip pptx</dc:title><dc:creator>QA Bot</dc:creator>   (no psmdcp)  validate 0
```

### 6) NeutralMarkdown title (the convert→md paths that go through the serializer)

```
$ aioffice convert deck.pptx deckmd.md   → leads with "# Quarterly Field Report"   (pptx→md, NeutralMarkdown.Write)
$ aioffice convert props.xlsx xlsxmd.md  → leads with "# Round Trip xlsx"           (xlsx→md, NeutralMarkdown.Write)
```

All three format→md paths (docx via the bridge, xlsx/pptx via NeutralMarkdown) now carry the document Title.

## Invariants — held
- One JSON envelope; `--view properties` / `set /properties` / `get /properties` shape & behavior UNCHANGED.
- OpenXmlValidator: 0 errors on every mutated/converted file above.
- Round-trip law preserved (xlsx core part relocated bytes-verbatim; no exact-byte-size or
  collection-order asserts added; tests normalize line endings).
- No .sln edits, no git commit/push.

---

# AIOffice 0.11.0 (M10) — Integration Smoke Report

M10 = embedded objects (embed/extract) across all three formats · pptx equations through the shared LaTeX→OMML converter extracted into Core · 1.0 contract prep (unified `data.properties.{core,custom}`, `surfaceVersion`, CONTRACT.md). Tool count unchanged (17).

## 1. Build — PASS

`dotnet build AIOffice.sln -warnaserror` → 0 warnings, 0 errors. Net10.0, Nullable enable, TreatWarningsAsErrors. The Core abstractions (`IEmbedHost` + `EmbeddedObject`) reconciled to one canonical definition; all three handlers implement it.

## 2. Tests — PASS (1585/1585 across 7 projects)

```
AIOffice.Core.Tests     124   AIOffice.Mcp.Tests       82
AIOffice.Word.Tests     481   AIOffice.Pptx.Tests     443
AIOffice.Preview.Tests   24   AIOffice.Render.Tests    31
AIOffice.Excel.Tests    400
```

Word +1 (embed media-type-from-src regression), Pptx +10 (new pptx equation suite: validator-clean, latex round-trip, equation_partial warning, placement+fontSize, remove, errors). All deterministic — new tests carry no Environment.NewLine / exact-byte-size / unsorted-order asserts; no real-browser test runs on CI.

## 3. End-to-end CLI smoke (temp workspace, real `EditOp.ParseBatch` / `DocPath` gate) — PASS

```
[1] DOCX embed report.docx <- data.xlsx
  add: /embed[1] application/vnd.openxmlformats-officedocument.spreadsheetml.sheet 5804 bytes
  list --view embeds: [('/embed[1]', 'Q3 model')]
  extract byte-identical (shasum src == out): PASS
  validate after remove: True
[2] XLSX embed book.xlsx <- data.xlsx (sheet /Data)
  add: /Data/embed[1]
  extract byte-identical: PASS
  validate: True
[3] PPTX embed deck.pptx <- data.xlsx
  add: /slide[1]/embed[@id=3]
  extract byte-identical: PASS
  validate: True
[4] PPTX equation
  add: /slide[1]/shape[@id=5]/omath[1]
  get latex: x = \frac{1}{2}
  "\foobar x" partial warning: equation_partial   (file still valid)
  validate: True
[5] Unified properties shape (data.properties.core.title) on all 3 formats
  report.docx: T1   book.xlsx: T2   deck.pptx: T3
[6] surfaceVersion
  schema.surfaceVersion: 1.0-rc
  doctor.capabilities.surfaceVersion: 1.0-rc | version 0.11.0 | mcpTools 17
[7] Sandbox denial
  embed src escape (../escape.xlsx)  -> sandbox_denied
  extract dest escape (../escape.bin) -> sandbox_denied
[8] XLSX equation N/A
  add type:equation on .xlsx -> unsupported_feature ("Excel has no equation object — spreadsheets use cell formulas, not OMML math")
```

The extracted bytes equal the embedded source byte-for-byte (verified by `shasum -a 256`), even across an open+save cycle — the round-trip law holds for embedded payloads. The pptx equation OMML is native (`m:oMath` → `m:f` fraction inside `mc:AlternateContent`/`a14:m`/`m:oMathPara`), validator-clean.

## Invariants — held
- One JSON envelope; every error carries a non-empty suggestion; `unsupported_feature` names the workaround.
- `--view properties` is now `data.properties.{core,custom}` on docx AND xlsx AND pptx (deliberate pre-1.0 consistency fix; per-format tests updated with a comment).
- OpenXmlValidator: 0 errors on every mutated file above (embed add/remove, equation add, partial equation).
- 1-based addressing; new canonical forms `/embed[i]`, `/Sheet1/embed[i]`, `/slide[i]/embed[@id=N]`, `/slide[i]/shape[@id=N]/omath[k]` flow through the real ParseBatch/DocPath/PptxAddress gate.
- No .sln edits, no git commit/push.

## 4. Published binary — PASS (osx-arm64, 0.11.0)

`dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release -p:PublishSingleFile=true --self-contained -o dist/osx-arm64`

```
binary: dist/osx-arm64/aioffice — 38,313,113 bytes (~36 MB), single self-contained file
doctor: version 0.11.0 | surfaceVersion 1.0-rc | mcpTools 17
embed+extract loop: report.docx <- data.xlsx, extract -> byte-identical (shasum match): PASS
pptx equation: add /slide[1]/shape[@id=3]/omath[1], get latex "x = \frac{1}{2}", validate: True
help embeds: resolves; help equations: now documents pptx; schema.surfaceVersion: 1.0-rc
```

## 5. Manual-check fixtures (open these in real Office)

- `fixtures/manual-check/embed-demo.docx` — a report carrying an embedded `.xlsx`; double-click the OLE icon in Word to open the source workbook.
- `fixtures/manual-check/equation-demo.pptx` — a slide with the quadratic formula and E=mc² as native PowerPoint math (Insert → Equation shows them as editable math).

## Invariants — held (binary + fixtures)
- The published binary returns the SAME envelopes as `dotnet run` (one JSON object, 17 tools, surfaceVersion 1.0-rc).
- Both fixtures pass `validate` (OpenXmlValidator 0 errors); the embed payload round-trips byte-identical; the equations carry their LaTeX for read-back.
- No .sln edits, no git commit/push.

---

# 1.0.0 release — stabilization sign-off (osx-arm64) — PASS

The 1.0.0 stabilization pass: `Version` bumped to 1.0.0, `surfaceVersion` promoted
`1.0-rc` → `1.0`, the CONTRACT finalized, the surface/contract/doc drift the six
auditors found fixed, and the surface re-verified end-to-end against a freshly
published single-file binary.

## Build & tests — PASS (1590/1590 across 7 projects)

`dotnet build AIOffice.sln -c Release -warnaserror` → **0 warnings, 0 errors**.
`dotnet test AIOffice.sln -c Release`:

| project | tests |
|---|---|
| AIOffice.Core.Tests | 124 |
| AIOffice.Word.Tests | 481 |
| AIOffice.Excel.Tests | 400 |
| AIOffice.Pptx.Tests | 443 |
| AIOffice.Mcp.Tests | 87 |
| AIOffice.Preview.Tests | 24 |
| AIOffice.Render.Tests | 31 |
| **total** | **1590** |

0 failures, 0 skipped (browser tests skip only on CI; locally Chromium present).
The +5 over the 0.11.0 baseline (1585) are 4 new `SchemaConsistencyTests` (the CLI
vs MCP surface guard) and 1 new xlsx `convert_lossy` test.

## Publish — PASS

```
$ dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release -p:PublishSingleFile=true --self-contained -o dist/osx-arm64
$ ls -l dist/osx-arm64/aioffice
-rwxr-xr-x  38315481  dist/osx-arm64/aioffice   # 36.5 MB
$ dist/osx-arm64/aioffice version
{"ok":true,"data":{"name":"aioffice","version":"1.0.0","runtime":".NET 10.0.8"},...}
```

## End-to-end CLI smoke (fresh temp workspace, published binary) — PASS

One command per verb, every one `ok:true` / exit 0:

```
create r.docx · create b.xlsx · create d.pptx · read r.docx outline ·
edit b.xlsx set range · query b.xlsx cells · get b.xlsx /Sheet1/A1 ·
render r.docx html · validate b.xlsx · template t.docx -o o.docx ·
audit r.docx · diff r.docx --snapshot 1 · convert b.xlsx b.docx ·
snapshot list r.docx · doctor · schema · help addressing · version   (all exit 0)

doctor: version 1.0.0 | surfaceVersion 1.0 | verbs 18 | mcpTools 17
schema.warningCodes: 18 codes   (new in 1.0: warning vocabulary surfaced)
```

Stabilization fixes verified live on the published binary:

- **schema consistency** — CLI `schema` and MCP `office_schema` now agree on the
  shared vocabularies: `read --view` enum includes `embeds` on both; `edit` op list
  includes `extract` on both.
- **xlsx read view echo** — `read b.xlsx --view outline` → `view=outline kind=xlsx`
  (previously `kind` only).
- **xlsx convert_lossy** — `convert b.xlsx b.docx` (workbook with a chart) →
  `data.dropped=1` + a `convert_lossy` warning (previously silent; the loss only
  showed as a body `[dropped]` note).
- **edit snapshot parity** — xlsx and pptx edits now surface `snapshot=<n>` like docx.
- **query/validate parity** — xlsx `query` carries `count` (+ `total`); xlsx
  `validate` carries `count` (+ `errors`/`warnings`).
- **comments view** — works on docx/xlsx/pptx (CONTRACT §6 corrected to "universal").
- **addressing help** — documents the M8–M10 `/embed[…]` and pptx `…/omath[…]` forms.

## Sandbox & security — PASS (every escape DENIED, exit 4, nothing leaked)

```
read /etc/hosts                  -> exit 4 (sandbox_denied)
read ../../../../etc/hosts       -> exit 4 (sandbox_denied)
read <canary outside ws>         -> exit 4 (sandbox_denied)
convert <ws>.docx /tmp/esc.md    -> exit 4; /tmp/esc.md NOT created
template --data @/etc/hosts      -> exit 4 (sandbox_denied)
canary file intact (never read/written)
```

## MCP — PASS

`tools/list` over stdio returns exactly **17** tools, including `file_snapshot`
(the snapshot tool's real name; CONTRACT §7 corrected from 16 → 17 names).

## Invariants — held (1.0.0)
- Published binary == `dotnet run` envelopes (one JSON object per call); **17 MCP
  tools**, **surfaceVersion `1.0`**, package version **1.0.0**.
- Every error carries a non-empty `suggestion`; exit-code map (0/2/3/4/5) unchanged.
- CONTRACT.md §§1–7 match the code exactly (verified by smoke + `SchemaConsistencyTests`).
- No git commit/push, no tags (left for the human release engineer).

---

# 1.1.0 release — first post-1.0 feature release, additive only (osx-arm64) — PASS

The 1.1.0 pass: `Version` bumped 1.0.0 → 1.1.0, **`surfaceVersion` stays `1.0`** (all
changes are additive within the frozen 1.0 contract line — new chart kinds, an
`iconSet` conditional-format kind, new `add` types `source`/`citation`/`bibliography`
(docx) and `media` (pptx), text/shape effects, new transitions, a `sources` read
view, and one new warning `bibliography_cached`). CONTRACT.md got an additive §7a; the
CLI/MCP schema + `office_help` surface the new vocabulary; verified end-to-end against
a freshly published single-file binary.

## Build & tests — PASS (1692/1692 across 7 projects)

`dotnet build AIOffice.sln -c Debug -warnaserror` → **0 warnings, 0 errors**.
`dotnet test AIOffice.sln`:

| project | tests |
|---|---|
| AIOffice.Core.Tests | 124 |
| AIOffice.Word.Tests | 507 |
| AIOffice.Excel.Tests | 430 |
| AIOffice.Pptx.Tests | 489 |
| AIOffice.Mcp.Tests | 87 |
| AIOffice.Preview.Tests | 24 |
| AIOffice.Render.Tests | 31 |
| **total** | **1692** |

0 failures, 0 skipped. The +102 over the 1.0.0 baseline (1590) are the new 1.1
feature tests: Word +26 (`CitationTests`, `TextEffectTests`), Excel +30
(`ExpandedChartKindsTests`, `IconSetConditionalFormatTests`), Pptx +46
(`EffectTests`, `MediaTests`, expanded `ChartTests`/`TransitionTests`). The 1.0
`SchemaConsistencyTests` / `TokenBudgetTests` still pass — the additive enum/prop
growth stays inside the ≤3500-token MCP budget (the descriptions are mostly enum
additions; the rich prose lives in `office_help`, not the always-resident surface).

## Publish — PASS

```
$ dotnet publish src/AIOffice.Cli -r osx-arm64 -c Release -p:PublishSingleFile=true --self-contained -o dist/osx-arm64
$ ls -l dist/osx-arm64/aioffice
-rwxr-xr-x  38352937  dist/osx-arm64/aioffice   # 36.6 MB
$ dist/osx-arm64/aioffice doctor
... version 1.1.0 | surfaceVersion 1.0 | verbs 18 | mcpTools 17 ...
```

## End-to-end smoke (fresh temp workspaces, published binary) — PASS

**xlsx** — sales sheet + four new charts + an icon-set conditional format:

```
add chart kind=doughnut     -> ok; validate valid=true issues=0
add chart kind=stackedBar   -> ok; validate valid=true issues=0
add chart kind=combo        -> ok (first series columns + rest a line, >=2 series)
add chart kind=bubble       -> ok (X col + Y/size pair per series); validate issues=0
add conditionalFormat iconSet set=3TrafficLights1 -> ok; validate issues=0
add chart kind=funnel       -> unsupported_feature, candidates list the expanded set
```

**pptx** — radar + stacked-area charts, embedded media, an effect, a transition:

```
add chart kind=radar        -> ok; validate valid=true issues=0
add chart kind=stackedArea  -> ok; validate valid=true issues=0
add media src=clip.mp4      -> ok; media/mediadatamp4 part present
add media src=tone.wav      -> ok; media/mediadatawav part present
add media src=../../../etc/hosts -> sandbox_denied (escaping src refused before any read)
set shape glow=FFAA00       -> ok (a:effectLst)
set slide transition=zoom   -> ok; validate valid=true issues=0
```

**docx** — two sources, both cited, a bibliography, a shadow text effect:

```
add source Smith2020 + Jones2019 -> ok
cite both in /body/p[1], /body/p[2] -> ok
add bibliography style=APA  -> ok + bibliography_cached warning present
read --view sources         -> count=2 (tags Smith2020, Jones2019)
set run shadow=true         -> ok; w14:shadow present; validate valid=true issues=0
```

**doctor / schema / MCP**:

```
doctor: version 1.1.0 | surfaceVersion 1.0 | verbs 18 | mcpTools 17
schema edit: type enum includes source/citation/bibliography/media
tools/list over stdio: exactly 17 tools (file_snapshot included)
SchemaConsistencyTests / TokenBudgetTests / SchemaHelpStatusTests: 14/14 PASS
```

## Manual-check fixtures (1.1) — added

- `fixtures/manual-check/deck-1.1-media.pptx` — a slide with a **doughnut** chart, an
  embedded **media** placeholder, and a **zoom** transition (validator-clean).
- `fixtures/manual-check/doc-1.1-bibliography.docx` — two cited sources + an **APA
  bibliography** + a **shadow** text effect (validator-clean).

## Invariants — held (1.1.0)
- Published binary == `dotnet run` envelopes; **17 MCP tools**, **surfaceVersion
  `1.0`** (unchanged), package version **1.1.0**.
- All 1.1 changes are additive: nothing in CONTRACT §§1–7 removed or renamed; §7a
  records the additions; `bibliography_cached` added to the frozen warning list.
- Every error carries a non-empty `suggestion`; exit-code map (0/2/3/4/5) unchanged.
- Binary size 38,352,937 bytes (~36.6 MB).
- No git commit/push, no tags (left for the human release engineer).

---

# AIOffice 1.2.0 — Integration Smoke Report

Date: 2026-06-15 · Machine: macOS 26.3.0 (Darwin 25.3.0) arm64 · dotnet 10.0.300 (TFM net10.0)
All commands below were actually executed against the published `dist/osx-arm64/aioffice`; outputs are trimmed but real.

## Build & tests — PASS

```
$ dotnet build AIOffice.sln -warnaserror   → 0 warnings, 0 errors
$ dotnet test  AIOffice.sln --no-build      → 1807 passed, 0 failed across 7 projects
   Core 124 · Word 529 · Excel 477 · Pptx 535 · MCP 87 · Preview 24 · Render 31
```

The 1.0 guards stay green: `SchemaConsistencyTests` (verb set, `read --view` enum,
`edit` op list, 17-tool catalog) and `TokenBudgetTests` (whole tool surface ~2970
tokens ≤ 3500). The new 95 tests landed WITH the features (FormControl 15 ·
NumberFormatPreset 8 · Protection 13 · Connector 11 · Group 16 · SmartArtCreate 11 ·
Index 8 · MergeField 7 · TableOfFigures 6). CI-hygiene scan of the new tests: no
`Environment.NewLine`, no exact-byte-size asserts, no unsorted-order asserts.

## Surface wiring (additive within frozen 1.0)

- New `office_edit` `add` types: pptx `smartart`/`connector`/`group`/`ungroup`,
  docx `tableOfFigures`/`indexEntry`/`index`/`mergeField`, xlsx `formControl` —
  wired into CLI `--type` doc, MCP `office_edit` schema, CONTRACT §7b.
  `group`/`ungroup` are `add` **types**, NOT new op kinds (op kinds stay 8).
- New props: cell `locked`, sheet/workbook protection (`protected`,
  `protectStructure`, `protectWindows`, `password`, `allow*`), `numberFormat`
  named presets — surfaced in `properties-xlsx` help.
- Reused views (no new view names): docx `read --view structure` now lists
  `tablesOfFigures`/`indexes`/`mergeFields`; `read --view fields` lists merge fields.
- New warnings `figures_cached`/`index_cached` (Core `WarningCodes`, CONTRACT §1).
- New help topics: `smartart`, `connectors`, `number-formats`, `structural-fields`
  (CLI markdown + MCP `HelpTopics` dictionary), index topic extended.

## Real end-to-end smoke (published binary)

**pptx** — create deck; `add smartart` (process, 4 nodes) → validate 0 → `get`
reads back `Plan/Build/Ship/Review`; `add` two shapes then a `connector` (elbow,
endArrow) → validate 0; `group` the two shapes → `get /slide[1]/group[@id=N]` lists
2 children → `ungroup` → validate 0.

**docx** — add 2 Figure captions; `add tableOfFigures` → structure shows it +
`figures_cached` warning + validate 0; `add indexEntry` then `add index` →
`index_cached` + structure entryCount 1 + validate 0; `add mergeField "Name"`;
`template --data {"Name":"Acme"}` filled **both** the MERGEFIELD «Name» and the
`{{Name}}` placeholder (Acme ×2, no markers left), validate 0.

**xlsx** — `add formControl` checkbox→F2 and a comboBox with items → validate 0 →
structure lists 2 controls (checkbox, comboBox); unlock `A1:B2` then `protect` the
sheet → `get /Sheet1` shows `protection.protected:true`, `get /Sheet1/A1` shows
`locked:false` → validate 0; `set B2 numberFormat:"accounting-usd"` → `get` shows
the resolved code `_("$"* #,##0.00…)` and cached display ` $ 1,234.50 `.

```
doctor: version 1.2.0 | surfaceVersion 1.0 | verbs 18 | mcpTools 17
schema edit: type doc includes smartart/connector/group/tableOfFigures/formControl
tools/list over stdio: exactly 17 tools (file_snapshot included)
SchemaConsistencyTests / TokenBudgetTests: PASS
```

## Published-binary smoke loop (dist/osx-arm64/aioffice)

- pptx SmartArt (cycle, 3 nodes) → validate 0 → reads back `A/B/C`.
- xlsx formControl (checkbox→C2) → validate 0.
- docx caption + tableOfFigures → validate 0 + `figures_cached` + structure
  entryCount 1.

## Manual-check fixtures (1.2) — added

- `fixtures/manual-check/deck-1.2-smartart.pptx` — a **SmartArt** process diagram +
  two rounded shapes joined by an **elbow connector** (validator-clean).
- `fixtures/manual-check/doc-1.2-figures.docx` — two Figure captions, a **table of
  figures** ("List of Figures") and a 2-column **index** (`figures_cached` +
  `index_cached`; validator-clean).
- `fixtures/manual-check/workbook-1.2-protected.xlsx` — a **checkbox** form control,
  an `accounting-usd` numberFormat cell, an unlocked `A1:B2` range and a
  **protected** sheet (validator-clean).

## Invariants — held (1.2.0)
- Published binary == `dotnet run` envelopes; **17 MCP tools**, **surfaceVersion
  `1.0`** (unchanged), package version **1.2.0**.
- All 1.2 changes are additive: nothing in CONTRACT §§1–7 removed or renamed; §7b
  records the additions; op kinds unchanged (`group`/`ungroup` are `add` types);
  `figures_cached`/`index_cached` added to the frozen warning list.
- Every error carries a non-empty `suggestion`; exit-code map (0/2/3/4/5) unchanged.
- Binary size 38,408,217 bytes (~36.6 MB).
- No git commit/push, no tags (left for the human release engineer).

---

# AIOffice 1.3.0 — Integration Smoke Report

Date: 2026-06-15 · Machine: macOS 26.3.0 arm64 · dotnet 10.0.300 (TFM net10.0)
Third post-1.0 feature release — purely additive, `surfaceVersion` stays `1.0`.

## Build & tests — PASS (1924/1924 across 7 projects)

```
$ dotnet build AIOffice.sln -warnaserror           → 0 warnings, 0 errors
$ dotnet test AIOffice.sln --no-build              → all green
  Core 124 · Word 556 · Excel 522 · Pptx 580 · MCP 87 · Preview 24 · Render 31
SchemaConsistencyTests / TokenBudgetTests: PASS (17 MCP tools; surface ≤ 3500 tokens)
```

New dedicated feature tests landed WITH the features: ChartPolishTests (Excel 17 /
Pptx 16), AdvancedConditionalFormatTests 16, PivotCalculatedFieldTests 9,
BodyShapeTests 10, FormFieldTests 10, ThemeTests 6, Model3DTests 13,
MotionPathTests 14.

## Surface wiring (additive within frozen 1.0)

- `office_edit` add types: docx `shape`/`textBox`/`formField`, pptx `model3d`.
- Chart-polish props (xlsx + pptx) on `add` and `set /…/chart[k]`: `dataLabels`,
  `legend`, `axisTitles`, `trendline`, `errorBars`, `gridlines`, `secondaryAxis`;
  `get` reports them under `polish`.
- xlsx `conditionalFormat` kinds `formula`/`topBottom`/`aboveBelowAverage`; pivot
  `calculatedFields`; pptx animation `motionPath` effect; docx `set /theme`.
- New `office_help` topics: `chart-polish`, `conditional-format`, `themes`,
  `3d-models`, `form-fields`, `animations` (CLI help array + MCP HelpTopics index).
  Details live in `office_help`, NOT in the tool schemas — token budget unchanged.
- New warning `model3d_as_media`. surfaceVersion stays `1.0`; CONTRACT §7c records
  the additions; 18 verbs / 17 MCP tools unchanged.

## Real end-to-end smoke (`dotnet run`, fresh temp workspace) — PASS

**xlsx** — sales sheet → bar chart with `dataLabels` + `trendline:linear` +
`axisTitles` + `legend:bottom` → validate 0 → `get` shows the polish; `set` adds a
secondary axis for one series → validate 0, `get.polish.secondaryAxis=["Target"]`;
`formula` CF (`=$B2>100`) + `topBottom` (top 3) CF → validate 0; pivot with a
`calculatedField` `Margin=Revenue-Target` → validate 0 → `get.calculatedFields`
reports it.

**docx** — roundRect body shape (fill/line/text "Reviewed") + a text box →
validate 0 → `get` shows geometry (xCm/yCm/wCm/hCm/fill/line/text); `set /theme
accent1=38BDF8 + minorFont=Calibri` → `get /theme` reflects both; dropdown form
field `status` (items) → `read --view fields` lists it (`fieldKind:dropdown`) →
validate 0.

**pptx** — line chart with `dataLabels` + `legend:right` → validate 0 → `get` shows
polish; tiny generated `.glb` 3D model embedded → media part `media/mediadataglb`
present + `model3d_as_media` warning + sandbox denial on `../outside.glb`
(`sandbox_denied`) → validate 0; `motionPath` (arc) animation on the model shape →
`read --view structure` lists it (`effect:motionPath`, `class:path`,
`direction:arc`) → validate 0.

## Published-binary smoke loop (dist/osx-arm64/aioffice) — PASS

```
doctor: version 1.3.0 | surfaceVersion 1.0 | verbs 18 | mcpTools 17
help index lists chart-polish / conditional-format / themes / 3d-models /
  form-fields / animations
```

- xlsx polished bar chart (dataLabels + legend:bottom + linear trendline) →
  validate 0 → `get.polish` reports the settings.
- pptx 3D model embed → `model3d_as_media` warning → validate 0.
- docx theme edit loop → `set /theme {accent1, minorFont}` → `get /theme` reflects
  accent1=38BDF8, minorFont=Calibri → validate 0.

## Manual-check fixtures (1.3) — added

- `fixtures/manual-check/workbook-1.3-polished.xlsx` — a bar chart with data labels,
  a legend, axis titles and a linear trendline, plus a pivot with a `Margin`
  calculated field (validator-clean).
- `fixtures/manual-check/deck-1.3-model3d.pptx` — a placeholder `.glb` 3D model
  (poster fallback, `model3d_as_media`) and an arc `motionPath` animation
  (validator-clean).
- `fixtures/manual-check/doc-1.3-theme.docx` — an edited theme (accent1 + minorFont),
  a rounded-rect body shape, a text box and a dropdown form field (validator-clean).

## Invariants — held (1.3.0)
- Published binary == `dotnet run` envelopes; **17 MCP tools**, **surfaceVersion
  `1.0`** (unchanged), package version **1.3.0**.
- All 1.3 changes are additive: nothing in CONTRACT §§1–7 removed or renamed; §7c
  records the additions; op kinds unchanged; `model3d_as_media` added to the frozen
  warning list.
- Every error carries a non-empty `suggestion`; exit-code map (0/2/3/4/5) unchanged.
- Binary size 38,479,545 bytes (~36.7 MB).
- No git commit/push, no tags (left for the human release engineer).

# AIOffice 1.4.0 — Integration Smoke Report

Date: 2026-06-15 · Machine: macOS 26.3.0 arm64 · dotnet 10.0.300 (TFM net10.0)
Fourth post-1.0 feature release — purely additive, `surfaceVersion` stays `1.0`.
Closes the long-standing dynamic-array gap (FILTER/UNIQUE/SORT now evaluate + spill).

## Build & tests — PASS (2016/2016 across 7 projects)

```
$ dotnet build AIOffice.sln -warnaserror           → 0 warnings, 0 errors
$ dotnet test AIOffice.sln --no-build              → all green
  Core 124 · Word 580 · Excel 545 · Pptx 625 · MCP 87 · Preview 24 · Render 31
SchemaConsistencyTests / TokenBudgetTests: PASS (18 verbs / 17 MCP tools; surface ~3150 ≤ 3500 tokens)
```

New dedicated feature tests landed WITH the features (engine workers): DynamicArrayTests,
FinancialFunctionTests, DataTableTests (Excel), MailMergeTests, IfFieldTests,
PageBorderTests (Word), ZoomTests, AnimationTriggerTests, TableStyleTests (Pptx).

## Surface wiring (additive within frozen 1.0)

- `office_edit` `type` enum (CLI + MCP) gains `dataTable` (xlsx), `ifField` (docx),
  `zoom` (pptx); a dynamic-array formula on `set value` spills automatically.
- `office_template` (the verb is unchanged): `--data` accepts a record ARRAY (mail
  merge); the `template` verb gains `--output PATTERN`; `office_template` gains an
  optional `output` param — **still 17 tools / 18 verbs** (SchemaConsistencyTests green).
- New addressing forms: `/Sheet1/dataTable[i]`, `/slide[i]/zoom[k]`.
- New props surfaced: `pageBorder` (docx section), `triggerOn` (pptx animation),
  pptx table `style`/`firstRow`/`lastRow`/`bandRow`/`firstCol`.
- New error code `spill_blocked`. New `office_help` topics: `formulas`, `data-tables`,
  `mail-merge`, `page-borders`, `zoom`, `table-styles`; `animations` extended with `triggerOn`.
- `doctor` reports version 1.4.0 / surfaceVersion 1.0 / verbs 18 / tools 17.

## Real end-to-end smoke (`dotnet run`, fresh temp workspace) — PASS

xlsx:
- `=UNIQUE(A1:A6)` on a column with duplicates → anchor D1 spills 4 distinct values
  over `spillRange D1:D4`, **no `formula_not_evaluated` warning**.
- `=SORT(A1:A6)` spills sorted; `=FILTER(B1:B6,C1:C6)` (boolean flag column) spills
  50/80/90; `=SEQUENCE(2,3)` spills 1..6 over J1:L2.
- `=IRR(N1:N5)` → `get` shows a cached numeric value (0.1803), no warning.
- what-if `dataTable` (colInput) → validate 0.
- spill into an occupied range (`=UNIQUE` over a non-empty P1:P4) → `spill_blocked`
  with a suggestion naming the range to clear.

docx:
- template with «Name» MERGEFIELD + `{{city}}` + an `ifField` + a `pageBorder` →
  `--data` of 3 records + `--output "letters/letter-{n}.docx"` → 3 docs, each merged
  correctly (Ada/London, Grace/New York, Linus/Helsinki); the IF field resolves per
  record (US → "Domestic", FI → "International"); all 3 validate clean.
- `--output "../escape-{n}.docx"` → `sandbox_denied`.
- single-object `--data` still returns `{replaced, written}` (no `records` key) and
  fills one document unchanged.

pptx:
- slide zoom on slide 2 → slide 3 → validate 0; `get /slide[2]/zoom[1]` reports
  `{kind:"slide", target:"slide 3"}`.
- animation on shape A (id 2) `triggerOn:"@3"` → structure / `get` report
  `triggerOn: /slide[1]/shape[@id=3]`, `triggerShapeId: 3` → validate 0.
- table with `style:"medium2"` + `bandRow` → validate 0; `get` reports the style + flags.

## Published-binary smoke loop (dist/osx-arm64/aioffice) — PASS

```
$ aioffice doctor        → version 1.4.0 · surfaceVersion 1.0 · verbs 18 · tools 17
$ =UNIQUE(A1:A3)         → spills x,y over spillRange C1:C2 (no warning)
$ template … --output    → 2 records → m-1.docx "Hi Ada", m-2.docx "Hi Grace"
$ add type:zoom (slide)  → /slide[1]/zoom[1] → validate 0
```

## Manual-check fixtures (1.4) — added

- `fixtures/manual-check/workbook-1.4-dynamic-arrays.xlsx` — a `=UNIQUE` spill and a
  `=SORT` spill over a region column, plus an `=IRR` cashflow cell carrying a cached
  value (validator-clean).
- `fixtures/manual-check/doc-1.4-mailmerge.docx` — a mail-merge template: a «Name»
  MERGEFIELD, a `{{city}}` marker, an «IF Country = "US"» field and a page border
  (validator-clean). Drive it with `template --data '[…]' --output "letter-{n}.docx"`.
- `fixtures/manual-check/deck-1.4-zoom.pptx` — a slide zoom (slide 1 → slide 2) and a
  `medium2` banded-row table (validator-clean).

## Invariants — held (1.4.0)
- Published binary == `dotnet run` envelopes; **18 verbs / 17 MCP tools**,
  **surfaceVersion `1.0`** (unchanged), package version **1.4.0**.
- All 1.4 changes are additive: nothing in CONTRACT §§1–7 removed or renamed; §7d
  records the additions; op kinds unchanged; `spill_blocked` added to the frozen error
  list. The dynamic-array / financial evaluation is backward-compatible — cells that
  used to carry `formula_not_evaluated` now carry a cached value; the warning code is
  unchanged and still fires for other unevaluated formulas.
- Every error carries a non-empty `suggestion`; exit-code map (0/2/3/4/5) unchanged.
- MCP tool surface ~3150 tokens, within the 3500 ceiling (TokenBudgetTests green).
- Binary size 38,528,457 bytes (~36.7 MB).
- No git commit/push, no tags (left for the human release engineer).

# AIOffice 1.5.0 — Integration Smoke Report

Date: 2026-06-15 · Machine: macOS 26.3.0 arm64 · dotnet 10.0.300 (TFM net10.0)
Fifth post-1.0 feature release — purely additive, `surfaceVersion` stays `1.0`.
Closes the modern scalar-function gap (XLOOKUP/IFS/SWITCH/LET/TEXTJOIN now evaluate)
and adds the what-if toolkit (Scenario Manager + Goal Seek); Word table-cell formulas,
building blocks, line numbering; PowerPoint embedded fonts, action buttons, custom layouts.

## Build & tests — PASS (2125/2125 across 7 projects)

```
$ dotnet build AIOffice.sln -warnaserror           → 0 warnings, 0 errors
$ dotnet test AIOffice.sln --no-build              → all green
  Core 124 · Word 612 · Excel 572 · Pptx 675 · MCP 87 · Preview 24 · Render 31
SchemaConsistencyTests / TokenBudgetTests: PASS (18 verbs / 17 MCP tools; surface ~3235 ≤ 3500 tokens)
```

New dedicated feature tests landed WITH the features: ScalarFunctionTests, ScenarioTests,
GoalSeekTests (Excel), TableFormulaTests, BuildingBlockTests, LineNumberTests (Word),
FontEmbedTests, ActionButtonTests, CustomLayoutTests (Pptx). CI-hygiene scan of the new
test files: no Environment.NewLine, no exact-byte-size asserts, no unsorted-order asserts.

## Surface wiring (additive within frozen 1.0)

- `office_edit` `type` enum (CLI + MCP) gains `scenario` (xlsx), `buildingBlock` /
  `buildingBlockRef` (docx), `font` (pptx), `actionButton` (pptx); the M6 `layout` add
  gains a `placeholders` prop. Scalar XLOOKUP/IFS/SWITCH/LET/TEXTJOIN on `set value`
  now evaluate to cached values.
- New `set` props surfaced: `applyScenario` + `goalSeek` (xlsx), table-cell `formula` +
  `lineNumbers` (docx), `embedAll` (pptx font).
- New addressing forms: `/Sheet1/scenario[@name=…]`, `/buildingBlock[@name=…]`, `/fonts`,
  `/fonts/font[@name=…]`.
- New warning codes `goal_seek_no_solution`, `table_formula_cached`. New `office_help`
  topics: `scenarios`, `goal-seek`, `table-formulas`, `building-blocks`, `embedded-fonts`,
  `action-buttons`, `layouts`, `line-numbers`; `formulas` extended with the scalar functions
  (both the MCP HelpTopics registry and the CLI HelpTopics/*.md embedded resources).
- `doctor` reports version 1.5.0 / surfaceVersion 1.0 / verbs 18 / tools 17.

## Real end-to-end smoke (`dotnet run`, fresh temp workspace) — PASS

- **xlsx**: `=XLOOKUP("Banana",A1:A3,B1:B3)` → cached **20** (no `formula_not_evaluated`);
  `=IFS(B1>15,…)` → "small" (B1=10); `=TEXTJOIN(",",TRUE,A1:A3)` → "Apple,Banana,Cherry";
  `=LET(x,5,x*2)` → **10**. Added scenario "High" (B1=120, B2=80) + `applyScenario` →
  B1 became 120, B2 80, and B5(`=B1*2`) recalculated to **240**. `goalSeek` B1 so B5=100 →
  B1 solved to ~**50** (converged, achievedTarget ~100). validate → 0 errors.
- **docx**: 3-row numeric table, bottom cell `=SUM(ABOVE)` → `get` shows formula
  `=SUM(ABOVE)` + cached **30** (raw `w:fldSimple` renders 30). Building block "Note" added
  + inserted into body — its content appears at `/body/p[3]` (`read --view text`). Section
  line numbering start 1 → `get /section[1]` reflects `{start:1, increment:1,
  restart:continuous}`. validate → 0 errors.
- **pptx**: embedded a tiny generated `.ttf` → `ppt/fonts/font.ttf` part present +
  `embeddedFontLst`/`embeddedFont`/`p:font` registered + `get /fonts` lists it; an escaping
  `src` (`../../etc/hosts`) → `sandbox_denied` (never read). Custom layout "Hero" with
  title+body placeholders added on master[1] (type `cust`, shapeCount 2); a new slide bound
  via `layoutName:"Hero"` references `/master[1]/layout[2]`. A "next" action button added.
  validate → 0 errors.

## Published-binary smoke loop (dist/osx-arm64/aioffice) — PASS

```
$ aioffice doctor      → version 1.5.0 · surfaceVersion 1.0 · verbs 18 · tools 17
$ XLOOKUP loop         → D1 = 20 (no warning), validate ok
$ table-formula loop   → =SUM(ABOVE) cached 30, validate ok
$ embedded-font loop   → embed .ttf, get /fonts count 1, validate ok
```

## Manual-check fixtures (1.5) — added

- `fixtures/manual-check/workbook-1.5-scenarios.xlsx` — a `=XLOOKUP("South",…)` lookup, a
  `=B2*2` dependent, and a saved "High" scenario (validator-clean).
- `fixtures/manual-check/doc-1.5-table-formula.docx` — a 3-row table with a
  `=SUM(ABOVE)` formula cell (integer-formatted, cached 350), a "Footer Note" building block
  inserted into the body, and section line numbering (validator-clean).
- `fixtures/manual-check/deck-1.5-fonts.pptx` — an embedded "Brand Sans" font, a custom
  "Hero" layout (title+body placeholders) used by slide 1, and a "next" action button
  (validator-clean).

## Invariants — held (1.5.0)
- Published binary == `dotnet run` envelopes; **18 verbs / 17 MCP tools**,
  **surfaceVersion `1.0`** (unchanged), package version **1.5.0**.
- All 1.5 changes are additive: nothing in CONTRACT §§1–7 removed or renamed; §7e
  records the additions; op kinds unchanged; the frozen error/exit-code lists unchanged
  (the two new codes are `meta.warnings`, not errors). The scalar-function evaluation is
  backward-compatible — cells that used to carry `formula_not_evaluated` for XLOOKUP/IFS/
  SWITCH/LET/TEXTJOIN now carry a cached value; the warning code is unchanged and still
  fires for LAMBDA and the lambda-helpers.
- Every error carries a non-empty `suggestion`; exit-code map (0/2/3/4/5) unchanged.
- MCP tool surface ~3235 tokens, within the 3500 ceiling (TokenBudgetTests green).
- Binary size 38,578,633 bytes (~36.8 MB).
- No git commit/push, no tags (left for the human release engineer).

# AIOffice 1.6.0 — Distribution & Onboarding Smoke Report

**Distribution release — no capability change.** The native binary keeps the same surface as
1.5.0 (18 verbs / 17 MCP tools / `surfaceVersion 1.0`). This section records the real local
verification of the packaging + onboarding paths (npm, install.sh, cookbook, MCP handshake)
done by the integrator. **No publish, no signing, no git** was performed.

## Build & tests — PASS (2125/2125 across 7 projects, at version 1.6.0)

`dotnet build AIOffice.sln -c Release -warnaserror` → **0 warnings, 0 errors**.
`dotnet test AIOffice.sln -c Release --no-build` → all green:

| Project | Passed | Failed | Skipped |
| --- | --- | --- | --- |
| AIOffice.Core.Tests | 124 | 0 | 0 |
| AIOffice.Word.Tests | 612 | 0 | 0 |
| AIOffice.Excel.Tests | 572 | 0 | 0 |
| AIOffice.Pptx.Tests | 675 | 0 | 0 |
| AIOffice.Mcp.Tests | 87 | 0 | 0 |
| AIOffice.Preview.Tests | 24 | 0 | 0 |
| AIOffice.Render.Tests | 31 | 0 | 0 |
| **Total** | **2125** | **0** | **0** |

(SchemaConsistency + token-budget guards green. The version bump to 1.6.0 broke nothing.)

## Version bump — applied & consistent

- `Directory.Build.props` `<Version>` → **1.6.0**; the freshly published `osx-arm64` binary
  reports `{"version":"1.6.0"}` on `version` and `doctor` (surfaceVersion `1.0`, verbs 18,
  mcpTools 17).
- `npm/package.json` `version` = **1.6.0**; `dist/Formula/aioffice.rb` `version "1.6.0"`;
  `dist/install.sh` / `install.ps1` `FALLBACK_VERSION = v1.6.0`; `docs/INSTALL.md` and
  `npm/README.md` examples pin **v1.6.0**. Verified no stray version drift across all
  builder outputs (the only `1.5.0` strings remaining are the Formula's clearly-labeled
  "example from v1.5.0" sha256 placeholders and the dist/README test note).
- Homebrew formula `license` corrected **MIT → Apache-2.0** to match the repo `LICENSE`.

## npm path — PASS (tarball + bin shim + doctor + MCP handshake)

The repo is **private**, so anonymous `https.get` to its release assets 404s
(`releases/download/v1.5.0/SHA256SUMS → HTTP 404`); the real binary was fetched with
authenticated `gh release download` into a local mirror to drive the verification.

- `npm pack` (in `npm/`) → `aioffice-1.6.0.tgz` (6,255 bytes). Tarball contents exactly:
  `package/{README.md, bin/aioffice.js, install.js, package.json, platform.js}` — **no
  binary shipped inside the tarball** (downloaded on install, as designed).
- `npm install --ignore-scripts <tarball>` then seeded the **real v1.5.0** `aioffice-mac-arm64`
  into `bin/`; the postinstall idempotency path correctly logged
  *"could not fetch SHA256SUMS (Protocol http: not supported); a binary is already
  installed, keeping it"* over the plain-HTTP mirror (the production fetch uses HTTPS to
  GitHub) and kept the binary.
- **bin shim → native binary**: `node_modules/.bin/aioffice doctor` →
  `ok:true | version:1.5.0 | surfaceVersion:1.0 | verbs:18 | mcpTools:17` (the pinned
  download is 1.5.0, proving the shim spawns the right binary with full passthrough).
- **MCP handshake through the shim**: `initialize` + `notifications/initialized` +
  `tools/list` over stdio → `initialize` returned `serverInfo {name:"aioffice"}`,
  `protocolVersion:"2024-11-05"`; `tools/list` returned **exactly 17 tools**:
  `file_snapshot, office_audit, office_convert, office_create, office_diff, office_edit,
  office_get, office_help, office_query, office_read, office_render, office_schema,
  office_status, office_template, office_validate, preview_open, preview_selection`.

## install.sh path — PASS (download → SHA256 verify → install → run)

`VERSION=v1.5.0 AIOFFICE_BIN=/tmp/aio-bin sh install.sh` run end-to-end against the local
v1.5.0 mirror (the only edit was the base URL → mirror, because the private repo 404s the
public download URL — identical to the dist/README test note):

```
aioffice-install: v1.5.0 aioffice-mac-arm64 -> /tmp/aio-bin/aioffice
aioffice-install: sha256 verified (63a5d987013ed77561ee5e50d4a8e4f66ae81ec668cb0a5e99cc2b5e0c6aee6c)
aioffice-install: NOTE /tmp/aio-bin is not on your PATH. ... export PATH="/tmp/aio-bin:$PATH"
aioffice-install: installed. {"ok":true,"data":{"name":"aioffice","version":"1.5.0",...}}
```

- Installed binary: `aioffice version` → valid JSON envelope; `aioffice doctor` →
  `ok:true | version:1.5.0 | surfaceVersion:1.0 | verbs:18 | mcpTools:17 | handlers ready:true`.
- The four `sha256` example values in `dist/Formula/aioffice.rb` were confirmed to **exactly
  match** the real v1.5.0 `SHA256SUMS` (mac-arm64 `63a5d98…`, mac-x64 `b8bc8d2…`, linux-arm64
  `c73b06d…`, linux-x64 `842a55b…`) — the human swaps them for v1.6.0 values after tagging.

## Cookbook recipes — PASS (run verbatim on the 1.6.0 binary)

| Recipe | Result |
| --- | --- |
| 1. Quarterly report (.docx) | `validate` → `{valid:true, count:0}` |
| 3. Budget (.xlsx) SUM + XLOOKUP | `B4` → value 17000 / `=SUM(B2:B3)` / cached 17000; `F2` → value 12500 / `=_xlfn.XLOOKUP(...)` / cached 12500 |
| 4. Mail-merge from JSON | `{records:2, produced:[letter-Ada.docx, letter-Grace.docx], unresolved:[]}`; "Dear Ada," / "Your balance is $1,200…Gold member." |
| 7. Convert docx → pptx | `{from:docx, to:pptx, blocksWritten:2, dropped:[]}`; slide 1 Title "Status Report" |

## SKILL.md golden examples — PASS (spot-checked 3 on the 1.6.0 binary)

- **Example 1** (create/edit/read docx): `read --view outline` →
  `{path:/body/p[1], level:1, text:"Status Report"}` ✓
- **Example 2** (xlsx formula): `get /Sheet1/B4` →
  `{value:2050, formula:=SUM(B2:B3), cachedValue:2050, text:"2,050"}` ✓
- **Example 6** (diff vs snapshot): `diff --snapshot N --view detailed` →
  `{kind:modified, path:/body/table[1]/tr[4]/tc[3], before:+12%, after:+11.8%, detail:cell}` ✓

## Human-only publish steps (NOT done here — need credentials)

1. **Cut the v1.6.0 release** (tag `v1.6.0`; the existing `release.yml` builds + uploads the
   6 binaries + `SHA256SUMS`). Make the repo/releases **public** so the curl/npm/brew download
   URLs resolve for end users (today they 404 — private repo).
2. **npm publish**: `cd npm && npm publish --access public` (after `npm login`). To automate on
   tag, add the **`NPM_TOKEN`** repo secret — `.github/workflows/npm-publish.yml` then waits for
   `SHA256SUMS` and publishes (it skips cleanly with a notice while the secret is absent).
3. **Homebrew tap**: create public repo **`onecer/homebrew-tap`**, copy
   `dist/Formula/aioffice.rb` in as `Formula/aioffice.rb`, and replace the four `sha256`
   placeholders with the v1.6.0 values:
   `gh release download v1.6.0 -R onecer/AIOffice -p SHA256SUMS -O - | sort`.
4. **Install scripts** (`dist/install.sh` / `install.ps1`) are served straight from `main` —
   no publish step.
5. **Signing/notarization**: not done; roadmap in `docs/SIGNING.md`. Binaries remain unsigned
   (`xattr -d com.apple.quarantine` for direct macOS downloads).

## Invariants — held (1.6.0)
- Published `osx-arm64` binary reports version **1.6.0**, **18 verbs / 17 MCP tools**,
  **surfaceVersion `1.0`** (unchanged). The 1.6.0 binary is the same surface as 1.5.0 — no
  capability change; only the version string moved.
- No surface change in CONTRACT §§1–7; §7f records that 1.6.0 is a distribution release with
  no contract change. The frozen error/exit-code lists, op kinds, and addressing forms are
  unchanged.
- Every builder output (npm/, dist/, SKILL.md, docs/*) reconciled: product name `aioffice`,
  version `1.6.0`, asset names, `onecer/AIOffice` URLs, Apache-2.0 license — no drift.
- dist `osx-arm64/aioffice` rebuilt at 1.6.0, size 38,578,649 bytes (~36.8 MB); `doctor`
  smoke green.
- No npm publish, no homebrew push, no signing/notarization, no git commit/push/tag —
  all left for the human release engineer.

# AIOffice 1.7.0 — Print Readiness, Camera Tool, Calc Mode & Deeper Equations Smoke Report

**Sixth post-1.0 feature release — purely additive.** `surfaceVersion` stays `1.0`; **18 CLI
verbs / 17 MCP tools** unchanged; no new verb or tool. New `set` props on existing paths, one
new `add` type (`linkedPicture`), two new `set`-paths (`/notesMaster`, `/handoutMaster`), one
new addressing form (`/equation[@num=…]`), two new warnings (`linked_picture_static`,
`equation_numbers_cached`), and a deeper LaTeX→OMML converter shared by docx + pptx.

## Build & tests — PASS (2331/2331 across 7 projects, at version 1.7.0)

`dotnet build AIOffice.sln -c Debug -warnaserror` → **0 warnings, 0 errors**.
`dotnet test AIOffice.sln` → **2331 passed, 0 failed** across 7 projects:
Core **180** · Word **681** · Excel **619** · Pptx **709** · MCP **87** · Preview **24** ·
Render **31**. The 1.0 guards stay green: `SchemaConsistencyTests` (18 verbs / 17 tools, the
`read --view` enum, the 8 `edit` op kinds, the 17-tool catalog) and `TokenBudgetTests` (the
MCP tool surface still ≤ 3500 tokens after the additive `office_edit` description + help
topics). CI-hygiene grep over the new test files: no `Environment.NewLine`, no exact-byte-size
asserts, no unsorted-order asserts.

## Surface wiring (additive within frozen 1.0)

- `office_edit` schema (`ToolCatalog`) + `aioffice schema` (`SurfaceSchema`/`CommandSurface`)
  extended: `linkedPicture` add type, the new xlsx print/calc props, docx dropCap/picture
  watermark/STYLEREF·SYMBOL·QUOTE/equation-number, pptx notes/handout master set-paths +
  animation timing + table-cell alignment, and the deeper LaTeX constructs — all as additive
  description text.
- `office_help` / `aioffice help`: new topics `print-setup` and `masters`; `equations`,
  `animations`, `properties-docx`, `properties-xlsx`, `properties-pptx`, `addressing`
  deepened (both the CLI embedded `.md` topics and the MCP `HelpTopics` dictionary).
- `doctor` reports **version 1.7.0 / surfaceVersion 1.0 / verbs 18 / tools 17**.

## Real end-to-end smoke (`dotnet run`, fresh temp workspace) — PASS

- **docx**: `\begin{cases}` display equation with `number:true` → validate 0, fires
  `equation_numbers_cached`; `get /equation[@num=1]` reports `number:"(1)"`. `\begin{aligned}`
  display equation → validate 0. Drop cap (`dropCap=drop dropCapLines=3`) on a paragraph →
  validate 0. Picture watermark from a generated PNG → validate 0, `get /watermark[1]` →
  `{kind:picture, washout:true}`; an escaping picture path → `sandbox_denied`. STYLEREF (in a
  header) + SYMBOL (©) + QUOTE fields → validate 0.
- **xlsx**: print titles `1:1` + manual page break (row 3) + `fitToPage.fitToWidth 1` + print
  footer `{center:"Page &P of &N"}` → `get /Sheet1` reflects them under `pageSetup` → validate
  0. `add linkedPicture` of `A1:C5` at `G2` → validate 0, fires `linked_picture_static`,
  `get /Sheet1/linkedPicture[1]` → `{sourceRange:"A1:C5", anchor:"G2", …}`. `calculationMode
  manual` + `iterativeCalc` on `/` → `get /` reflects `{calculationMode:"manual",
  iterativeCalc:true, …}`.
- **pptx**: handout master header/footer + `slidesPerPage 3` → validate 0, `get /handoutMaster`
  reflects them. Animation `repeat:"untilClick"` + `autoReverse:true` via
  `set /slide[1]/animation[1]` → validate 0, `read --view structure` / `get` report
  `repeat:"indefinite", autoReverse:true`. Table cell `valign:middle` + four margins →
  validate 0. A pptx equation with `\begin{cases}` → validate 0 (confirms the deepened shared
  Core converter reaches pptx with no `equation_partial`).

> Note: two smoke steps invoke an existing convention rather than the literal phrasing —
> a new slide is added via `/slide[i]` (not `/`), and animation timing props are applied via
> `set /slide[i]/animation[k]` after the `add` (they are retime props, not add props). The help
> docs and CONTRACT describe both accurately.

## Published-binary smoke loop (dist/osx-arm64/aioffice) — PASS

`dotnet publish -c Release -r osx-arm64 --self-contained -p:PublishSingleFile=true` → the
single-file binary copied to `dist/osx-arm64/aioffice`. **doctor** → version **1.7.0**,
surfaceVersion **1.0**, **18 verbs / 17 tools**. One-loop check: docx `\begin{cases}` numbered
equation (validate 0, `equation_numbers_cached`); xlsx print titles `1:1` + print footer
(validate 0); pptx handout master header + `slidesPerPage 3` (validate 0). Binary size
**38,636,745 bytes (~36.85 MB)**.

## Fixtures — refreshed (1.7)

The manual-check fixtures (`fixtures/manual-check/word-sample.docx`,
`excel-sample.xlsx`, `pptx-showcase.pptx`) were refreshed by the per-format owners to exercise
the 1.7 surface (a cases/aligned equation + drop cap doc, a print-ready workbook, a deck with a
handout master + table-cell alignment).

## Invariants — held (1.7.0)

- One JSON envelope per call; errors carry a non-empty `suggestion`; unsupported →
  `unsupported_feature` naming the workaround. Exit codes 0/2/3/4/5 unchanged.
- `OpenXmlValidator` 0 errors after every mutating smoke. 1-based addressing throughout. The
  picture-watermark `image` and linked-picture source are sandbox-resolved (escaping →
  `sandbox_denied`).
- CONTRACT §§1–7 unchanged; the additive 1.7.0 surface is recorded in **§7g** (new add type,
  new set-paths, new props, new field kinds, the two new warnings, and the
  backward-compatible deeper-equation note: more LaTeX now renders, fewer `equation_partial`
  warnings). `surfaceVersion` stays `1.0`; 18 verbs / 17 tools unchanged.
- Version bumped to **1.7.0** in `Directory.Build.props`, `npm/package.json`,
  `dist/Formula/aioffice.rb`, and the `install.sh` / `install.ps1` fallback versions — all
  consistent.
- No npm publish, no homebrew push, no signing/notarization, no git commit/push/tag — all left
  for the human release engineer.
