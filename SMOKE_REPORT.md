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
