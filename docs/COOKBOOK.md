# AIOffice Cookbook

Ten task-oriented recipes for driving `aioffice`. Every command below was **run and verified**
against the binary; the `# →` lines show the actual (abridged) JSON that came back. Each command
prints one JSON envelope `{ok, data, error, meta}`; on `ok:false`, read `error.suggestion` and
`error.candidates` and retry.

Conventions used throughout:
- Quote every path/selector in your shell (`'/body/p[3]'`) so brackets aren't globbed.
- All file access is sandboxed to the workspace (default: cwd; override with `--workspace <dir>`).
- Every `edit` auto-snapshots the pre-image first, so you can always `snapshot restore`.

---

## 1. Generate a quarterly report (.docx) — title, headings, table

```bash
aioffice create q3-report.docx
aioffice edit q3-report.docx --ops '[
  {"op":"set","path":"/properties","props":{"title":"Q3 2026 Business Review","author":"AIOffice"}},
  {"op":"add","path":"/header[1]","type":"header","props":{"text":"Q3 2026 Business Review"}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"Q3 2026 Business Review","style":"Title"}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"Executive Summary","style":"Heading1"}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"Revenue grew 12% quarter-over-quarter, led by the Enterprise segment."}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"Results by Region","style":"Heading1"}},
  {"op":"add","path":"/body","type":"table","position":"inside","props":{"rows":4,"cols":3}},
  {"op":"set","path":"/body/table[1]","props":{"headerRow":true,"borders":"all"}},
  {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"Region"}},
  {"op":"set","path":"/body/table[1]/tr[1]/tc[2]","props":{"text":"Revenue"}},
  {"op":"set","path":"/body/table[1]/tr[1]/tc[3]","props":{"text":"YoY"}},
  {"op":"set","path":"/body/table[1]/tr[2]/tc[1]","props":{"text":"North"}},
  {"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"text":"$1.2M"}},
  {"op":"set","path":"/body/table[1]/tr[2]/tc[3]","props":{"text":"+14%"}},
  {"op":"set","path":"/body/table[1]/tr[3]/tc[1]","props":{"text":"South"}},
  {"op":"set","path":"/body/table[1]/tr[3]/tc[2]","props":{"text":"$0.9M"}},
  {"op":"set","path":"/body/table[1]/tr[3]/tc[3]","props":{"text":"+8%"}},
  {"op":"set","path":"/body/table[1]/tr[4]/tc[1]","props":{"text":"Total"}},
  {"op":"set","path":"/body/table[1]/tr[4]/tc[2]","props":{"text":"$2.1M"}},
  {"op":"set","path":"/body/table[1]/tr[4]/tc[3]","props":{"text":"+12%"}}]'
aioffice edit q3-report.docx --remove '/body/p[1]'   # drop the empty para `create` left behind
aioffice validate q3-report.docx
# → {"valid":true,"count":0,"issues":[]}
```

**Check:** `aioffice read q3-report.docx --view outline` lists the two Heading1s; `validate`
reports `valid:true`. (Style a table with `headerRow`/`borders`/`shading` — *not* a `style`
name; a table `set {style}` is `unsupported_feature`.)

---

## 2. Build a pitch deck (.pptx) — title slide + content slide + a chart

```bash
aioffice create deck.pptx
aioffice edit deck.pptx --ops '[
  {"op":"set","path":"/slide[1]","props":{"background":"0F172A"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"text":"Acme Q3 Review","x":2,"y":6,"w":28,"h":3,"fontSize":54,"bold":true,"color":"FFFFFF"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"text":"Prepared by AIOffice","x":2,"y":9.5,"w":20,"h":1.5,"fontSize":18,"color":"94A3B8"}}]'
aioffice edit deck.pptx --ops '[{"op":"add","path":"/slide[1]","type":"slide","position":"after","props":{"title":"Highlights"}}]'
aioffice edit deck.pptx --ops '[{"op":"add","path":"/slide[2]","type":"shape","props":{"text":"Revenue up 12% QoQ\nEnterprise segment led growth\nChurn down to 1.8%","x":2,"y":5,"w":28,"h":8,"fontSize":24,"color":"E2E8F0"}}]'
aioffice edit deck.pptx --ops '[{"op":"add","path":"/slide[2]","type":"slide","position":"after","props":{"title":"Revenue by quarter"}}]'
aioffice edit deck.pptx --ops '[{"op":"add","path":"/slide[3]","type":"chart","props":{"kind":"bar","categories":["Q1","Q2","Q3"],"series":[{"name":"Revenue ($k)","values":[120,150,210]}],"title":"Quarterly revenue","x":2,"y":4,"w":22,"h":12}}]'
aioffice validate deck.pptx
# → {"valid":true,"count":0,"issues":[]}
```

**Check:** `aioffice read deck.pptx --view outline` shows 3 slides. Bullets on a slide are
**newline-separated text** in one shape — you cannot `add type:"p"` to a pptx shape (the error
lists the addable types). `title` is honored only when **adding** a slide, never on `set`.
Then render and look (recipe 8) — in this layout the chart slightly overlaps the title box, a
typical look→fix nudge (`set /slide[3]/chart[1] {y}`).

---

## 3. Fill a budget (.xlsx) — table + SUM / XLOOKUP, verify cached values

```bash
aioffice create budget.xlsx
aioffice edit budget.xlsx --ops '[
  {"op":"set","path":"/Sheet1/A1","props":{"value":"Category"}},
  {"op":"set","path":"/Sheet1/B1","props":{"value":"Budget"}},
  {"op":"set","path":"/Sheet1/C1","props":{"value":"Actual"}},
  {"op":"set","path":"/Sheet1/A1:C1","props":{"bold":true,"fill":"E2E8F0"}},
  {"op":"set","path":"/Sheet1/A2","props":{"value":"Marketing"}},
  {"op":"set","path":"/Sheet1/B2","props":{"value":5000}},
  {"op":"set","path":"/Sheet1/C2","props":{"value":4200}},
  {"op":"set","path":"/Sheet1/A3","props":{"value":"Engineering"}},
  {"op":"set","path":"/Sheet1/B3","props":{"value":12000}},
  {"op":"set","path":"/Sheet1/C3","props":{"value":12500}},
  {"op":"set","path":"/Sheet1/A4","props":{"value":"Total"}},
  {"op":"set","path":"/Sheet1/B4","props":{"value":"=SUM(B2:B3)","bold":true}},
  {"op":"set","path":"/Sheet1/C4","props":{"value":"=SUM(C2:C3)","bold":true}},
  {"op":"set","path":"/Sheet1/B2:C4","props":{"numberFormat":"accounting-usd"}},
  {"op":"set","path":"/Sheet1/F1","props":{"value":"Engineering"}},
  {"op":"set","path":"/Sheet1/F2","props":{"value":"=XLOOKUP(F1,A2:A3,C2:C3)","numberFormat":"accounting-usd"}}]'
aioffice get budget.xlsx /Sheet1/B4
# → "value":17000,"formula":"=SUM(B2:B3)","cachedValue":17000,"text":" $ 17,000.00 "
aioffice get budget.xlsx /Sheet1/F2
# → "value":12500,"formula":"=_xlfn.XLOOKUP(F1,A2:A3,C2:C3)","cachedValue":12500
```

**Check:** the `cachedValue` of every formula is the evaluated number (`B4`=17000, the
`XLOOKUP` in `F2`=12500) — proof the formula engine ran at write time. Newer functions store
as `_xlfn.XLOOKUP` (correct OOXML; Excel shows the clean name). `accounting-usd` is a named
preset that resolves to a real Excel format code.

---

## 4. Mail-merge personalized letters from a JSON dataset

```bash
# Build the template once; {{key}} placeholders live in the body text.
aioffice create letter.docx
aioffice edit letter.docx --ops '[
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"Dear {{name}},"}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"Your balance is {{balance}}. Thank you for being a {{tier}} member."}}]'
aioffice edit letter.docx --remove '/body/p[1]'

cat > recipients.json <<'JSON'
[ {"name":"Ada","balance":"$1,200","tier":"Gold"},
  {"name":"Grace","balance":"$980","tier":"Silver"} ]
JSON

aioffice template letter.docx --data @recipients.json --output 'letter-{name}.docx'
# → {"records":2,"produced":["…/letter-Ada.docx","…/letter-Grace.docx"],"unresolved":[]}
aioffice read letter-Ada.docx --view text
# → "Dear Ada,"  /  "Your balance is $1,200. Thank you for being a Gold member."
```

**Check:** one output file per record; `unresolved` is empty (every placeholder was filled). The
template body uses **double** braces `{{name}}`; the `--output` filename pattern uses **single**
braces `{name}` (and `{n}` for the 1-based index). A JSON object instead of an array fills a
single document.

---

## 5. Audit a document for accessibility and `--fix`

```bash
# A document with a missing title and an image lacking alt text.
aioffice audit bad.docx
# → findings (DATA, exit 0):
#   a11y_no_doc_title  (warning, /properties,  "autofixable":true)
#   a11y_no_alt_text   (error,   /body/p[3],   "autofixable":true)
aioffice audit bad.docx --fix
# → {"summary":{"errors":0,"warnings":0,"infos":0},"fixed":2,"remaining":[]}
```

**Check:** the first run exits 0 — findings are data, not failures. `--fix` repairs only the
safe, autofixable findings (sets a title, inserts placeholder alt text) and re-audits; anything
report-only (e.g. a heading-level skip) stays in `remaining` for you to fix by hand. Scope with
`--category accessibility|quality|all` and `--severity error|warning|info`.

---

## 6. Diff a document against its snapshot to review an edit

```bash
# Make a change — every edit auto-snapshots the pre-image first.
aioffice edit q3-report.docx --find '+12%' --replace '+11.8%'
aioffice snapshot list q3-report.docx        # see the ring (last 20 per file)
aioffice diff q3-report.docx --snapshot 4 --view detailed
# → {"changes":[{"kind":"modified","path":"/body/table[1]/tr[4]/tc[3]",
#                "before":"+12%","after":"+11.8%","detail":"cell"}],
#     "summary":{"added":0,"removed":0,"modified":1,"moved":0}}
aioffice snapshot restore q3-report.docx     # undo it (the restore is itself snapshotted)
```

**Check:** the diff is a sorted, deterministic change list showing exactly the cell that changed,
before/after. `--view summary` collapses it to counts plus one path per change. `diff a.docx
b.docx` compares two separate files the same way.

---

## 7. Convert a docx report into a pptx outline

```bash
aioffice convert note.docx note-deck.pptx
# → {"from":"docx","to":"pptx","blocksWritten":2,"dropped":[],"written":"…/note-deck.pptx"}
aioffice read note-deck.pptx --view outline
# → slide 1: Title "Status Report" + Content "All systems nominal."
```

**Check:** headings become slide titles and the following body becomes the slide content.
Conversion is inherently lossy — anything that couldn't cross is named in a `convert_lossy`
warning (e.g. converting a deck *back* to docx drops charts:
`{"code":"convert_lossy","message":"…charts… is not converted"}`). Other pairs work too:
`docx↔md`, `xlsx↔csv`, and any source → `pdf/png/svg/html`.

---

## 8. Render a slide to PNG for the look → fix loop

```bash
aioffice render deck.pptx --to png --scope '/slide[1]' -o deck-1.png
# → {"format":"png","scope":"/slide[1]","written":"…/deck-1.png","sizeBytes":21149}   (1280x720)
aioffice render deck.pptx --to png --scope '/slide[3]' -o deck-3.png   # the chart slide
# now OPEN the PNG and look; fix overlaps/overflow, then render again
aioffice render deck.pptx --to svg --scope '/slide[3]' -o deck-3.svg    # SVG: each shape has data-aio-path
```

**Check:** the PNG is a real browser screenshot (1280×720 for 16:9). Over MCP, `render --to png`
also returns the image inline so the model sees it directly. SVG/HTML carry `data-aio-path`
attributes mapping every rendered element back to its canonical path — render once per visual
edit, look, fix; use `validate` for non-visual problems (it's cheaper than rendering).

---

## 9. Add a LaTeX equation to a document

```bash
aioffice create math.docx
aioffice edit math.docx --ops '[
  {"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"E = mc^2"}},
  {"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"\\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}","display":true}}]'
aioffice read math.docx --view text
# → "$E = mc^2$"   /   "$$\frac{-b \pm \sqrt{b^2-4ac}}{2a}$$"
aioffice get math.docx '/body/p[1]/omath[1]'
# → {"latex":"E = mc^2","display":false}
```

**Check:** LaTeX becomes real Office Math (OMML) and round-trips — `get` returns the stored
`latex` and `read --view text` shows `$…$` (inline) / `$$…$$` (display) markers. `display:true`
makes a centered block. Unrecognized commands degrade to literal text with an `equation_partial`
warning; the file still validates.

---

## 10. Embed a spreadsheet into a report, then extract it back

```bash
aioffice create embed-report.docx
aioffice edit embed-report.docx --ops '[
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"Attached workbook","style":"Heading1"}},
  {"op":"add","path":"/body","type":"embed","props":{"src":"sales.xlsx"}}]'
aioffice read embed-report.docx --view embeds
# → [{"path":"/embed[1]","name":"sales.xlsx","mediaType":"…spreadsheetml.sheet","size":6356}]
aioffice edit embed-report.docx --ops '[{"op":"extract","path":"/embed[1]","props":{"to":"extracted.xlsx"}}]'
shasum -a 256 sales.xlsx extracted.xlsx        # identical hashes -> byte-for-byte round-trip
aioffice get extracted.xlsx /Sheet1/B4          # the extracted workbook opens and evaluates
```

**Check:** the embed appears in `read --view embeds` with its real media type; `extract` writes
the original bytes back out — the SHA-256 of `extracted.xlsx` matches the source exactly, and it
re-opens as a working workbook. The extract op key is **`to`** (a workspace-relative path), not
`dest` — using `dest` returns an `invalid_args` error naming the right key.

---

### Quick reference

- Orient: `aioffice read <file> --view outline`
- Find: `aioffice query <file> '<selector>'` (e.g. `'p[style=Heading1]'`, `'cell[value>100]'`)
- Inspect: `aioffice get <file> '<path>'`
- Edit: `aioffice edit <file> --ops '[…]'` (add `--dry-run` to preview, `--expect-rev <rev>` to guard)
- Prove: `aioffice validate <file>`
- Ask the binary: `aioffice schema [verb]` (machine-readable) · `aioffice help <topic>` (prose)
