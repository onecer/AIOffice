# xlsx properties

## cell (`/Sheet1/A1`)

| prop        | type                 | notes                                          |
|-------------|----------------------|------------------------------------------------|
| value       | scalar               | auto-typed, see below                          |
| valueType   | string               | auto · text · number · boolean · dateTime      |
| numberFormat| string               | Excel format code, e.g. `0.00%`, `yyyy-mm-dd`  |
| bold        | bool                 |                                                |
| italic      | bool                 |                                                |
| color       | string               | font color, hex RGB                            |
| fill        | string               | background color, hex RGB                      |
| hyperlink   | string               | M5: `https://…` external, `#Sheet!A1` internal |
| hyperlinkTooltip | string          | M5: hover text for the hyperlink               |

Value auto-typing: a string with a leading `=` is a **formula**
(`"value": "=SUM(A1:A2)"`), ISO dates (`2024-05-01`, `2024-05-01T08:30`)
become DateTime, `true`/`false` booleans, parseable numbers numbers, the rest
text. Pass `valueType: "text"` to keep a literal (e.g. the string `=x`).

`get` on a formula cell returns both `formula` (with the leading `=`) and the
evaluated value. When a formula cannot be evaluated, the cached value is
returned and the envelope carries a `formula_not_evaluated` warning in
`meta.warnings`.

## range (`/Sheet1/A1:C10`)

`set` applies the given props to every cell in the range. `get` returns a 2-D
`values` array (and `formulas` where present).

## row (`/Sheet1/row[3]`)

| prop   | type   | notes               |
|--------|--------|---------------------|
| height | number | points              |
| hidden | bool   |                     |

`get` also reports M6 outline state: `outlineLevel` and `collapsed` (see
**outline grouping** below). Columns are addressed `/Sheet1/col[C]`.

## sheet (add)

`{op:"add", path:"/'New Sheet'", type:"sheet"}` creates a sheet. Sheet names
with spaces are quoted in paths: `/'Q3 Data'/B2`.

## chart (M1+M3, 1.1 expanded, `/Sheet1/chart[1]`)

`{op:"add", path:"/Sheet1", type:"chart", props:{kind:"bar|line|pie|scatter|area",
dataRange:"A1:C6", anchor:"E3", title?}}`. scatter (M3) plots numbers against
numbers: the first dataRange column is the numeric X axis, so every X cell
must be a number (header row excepted).

1.1 adds the kinds `doughnut`, `radar`, `bubble`, `stackedBar`,
`percentStackedBar`, `stackedArea`, `combo`. `bubble` lays out dataRange as an
X column then a Y+size column pair per series (e.g. `A1:C5` = one bubble
series). `combo` draws the first series as columns and the rest as a line, so
it needs at least two series. An unsupported `kind` returns
`unsupported_feature` listing the full supported set.

## name (M3, `/name[@name=SalesData]`)

`{op:"add", path:"/Sheet1/B2:C5", type:"name", props:{name:"SalesData",
scope?:"workbook|sheet"}}` — the op path is the range the name refers to.
Formulas can use it right away: `{op:"set", path:"/Sheet1/D6",
props:{value:"=SUM(SalesData)"}}` evaluates and caches a real value.
`get /name[@name=X]` (workbook scope) or `/Sheet1/name[@name=X]` (sheet scope).

## sheet props (M3, set on `/Sheet1` or a range)

| prop                    | target      | notes                              |
|-------------------------|-------------|------------------------------------|
| freezeRows, freezeCols  | sheet       | 0 clears that axis                 |
| autoFilter (bool)       | range       | one filter per sheet; false clears |
| orientation, paperSize  | sheet       | landscape/portrait; A4/Letter/…    |
| fitToWidth, fitToHeight | sheet       | print scaling, pages               |
| printArea               | sheet       | e.g. "A1:F40"                      |

`get /Sheet1` reflects freeze/autoFilter/pageSetup. Streaming reads (M3): files
over 20 MB (or `stream:true`) answer `read --view stats|text` and cell/range
`get` via a SAX scan without loading the workbook DOM.

## in-place streaming write (M6 flagship)

`edit … stream=true` (or any file over 20 MB) rewrites a LARGE existing
workbook **in place**, streaming the target sheet part through the SAX writer
instead of loading the whole DOM — set deep cells / bulk-write ranges in a
50 MB+ workbook in seconds with bounded memory.

Streamable ops (the whole batch must be streamable, else it falls back to the
DOM path): `set value` on a cell, `set value` with a formula (`"=…"`, no cached
value), `set values` (a 2-D bulk write) on a cell/range. Any other prop
(bold/fill/numberFormat/merge/…) makes the batch non-streamable. The response
marks `streamed:true`.

    aioffice edit big.xlsx --ops '[{"op":"set","path":"/Sheet1/H200000","props":{"value":"x"}}]' stream=true

## table (M6 Excel Table / ListObject, `/Sheet1/table[@name=Sales]`)

`{op:"add", path:"/Sheet1/A1:D20", type:"table", props:{...}}` turns a range
(first row = headers) into a structured table.

| prop          | type   | notes                                                  |
|---------------|--------|--------------------------------------------------------|
| name          | string | table name (used in the path + structured references)  |
| style         | string | none · light1-21 · medium1-28 · dark1-11 (e.g. `medium2`), or full `TableStyleMedium2` |
| headerRow     | bool   | default true                                           |
| totalsRow     | bool   | show/hide the totals row                               |
| totals        | object | column → function map, e.g. `{"Amount":"sum"}` (turns the totals row on) |
| bandedRows / bandedColumns | bool | row/column stripes                       |

Totals functions: sum · average · count · countNumbers · min · max · stdDev ·
var. Structured references evaluate because the table exists in the model
before save: `{op:"set", path:"/Sheet1/D1", props:{value:"=SUM(Sales[Amount])"}}`.
`get /Sheet1/table[@name=Sales]` describes range/style/columns/totals; `remove`
drops the ListObject but **keeps the cell data**.

## outline grouping (M6)

Group a contiguous row or column span into an outline level. The span is one
path segment: `row[a]:row[b]` or `col[a]:col[b]`.

    aioffice edit m.xlsx --ops '[{"op":"add","path":"/Sheet1/row[2]:row[6]","type":"group","props":{"collapsed":true}}]'
    aioffice edit m.xlsx --ops '[{"op":"add","path":"/Sheet1/col[B]:col[E]","type":"group"}]'

`{collapsed:true}` collapses the new group. `remove` over the same span
ungroups one outline level. Nested groups raise the level. A row/column `get`
and `read --view structure` report `outlineLevel`/`collapsed`; Excel draws the
outline symbols on reopen.

## pivot (M2, `/Pivot/pivot[@name=SalesPivot]`)

| prop        | type   | notes                                              |
|-------------|--------|----------------------------------------------------|
| sourceRange | string | required; header row + data rows, e.g. `A1:E7`     |
| targetSheet | string | required; created when absent — pivots land THERE  |
| name        | string | default auto-generated; used in the canonical path |
| targetAnchor| string | top-left cell on the target sheet (default A1/A3)  |
| rows / columns / filters | string[] | source column header names           |
| values      | array  | `[{field:"Sales", agg:"sum|average|count|min|max"}]` |

`{op:"add", path:"/Sheet1", type:"pivot", props:{...}}` — the op path is the
SOURCE sheet. `get`/`remove` address the pivot on its TARGET sheet:
`/Pivot/pivot[1]` or `/Pivot/pivot[@name=SalesPivot]`. Excel recomputes the
pivot on open (refreshOnLoad).

## conditionalFormat (M2, `/Sheet1/conditionalFormat[1]`)

`{op:"add", path:"/Sheet1/A1:C10", type:"conditionalFormat", props:{...}}` —
the op path is the range the rule covers.

| kind         | props                                                        |
|--------------|--------------------------------------------------------------|
| cellIs       | operator (`>` `>=` `<` `<=` `==` `!=` `between`), value, value2 (between only), fill?, color?, bold? |
| colorScale   | minColor, maxColor, midColor?                                |
| dataBar      | color                                                        |
| containsText | text, fill?, color?, bold?                                   |
| iconSet      | set (e.g. `3TrafficLights1`, `3Arrows`, `4Rating`, `5Quarters`), reverse?, showValue? (1.1) |

`iconSet` (1.1) ranks each cell against the others in the range and paints a
3/4/5-icon glyph; pick the icon family with `set` (3-icon families start with
`3`, etc.), `reverse:true` flips the order, `showValue:false` hides the number.

`get`/`remove` by `/Sheet1/conditionalFormat[i]` (later indices shift down
after a remove).

## image (M2, `/Sheet1/image[1]`)

| prop              | type   | notes                                      |
|-------------------|--------|--------------------------------------------|
| src               | string | required; PNG/JPEG path inside the sandbox |
| anchor            | string | required; top-left cell, e.g. `E2`         |
| name              | string | optional picture name                      |
| widthPx, heightPx | number | omit either to keep the aspect ratio       |

`{op:"add", path:"/Sheet1", type:"image", props:{src:"logo.png", anchor:"E2"}}`.

## dataValidation (M5, `/Sheet1/dataValidation[1]`)

`{op:"add", path:"/Sheet1/B2:B20", type:"dataValidation", props:{...}}` — the
op path is the range the rule covers.

| kind        | props                                                          |
|-------------|----------------------------------------------------------------|
| list        | `values:["Open","Closed"]` OR `sourceRange:"/Lists/A1:A3"` (not both) — Excel shows a dropdown |
| wholeNumber | operator (`between` `> >= < <= == !=`), value, value2 (between) |
| decimal     | same operator/value props                                       |
| date        | same; dates as ISO strings, e.g. `2026-01-01`                   |
| textLength  | same operator/value props                                       |

Extras on any kind: `allowBlank`, `inputTitle`, `inputMessage`, `errorTitle`,
`errorMessage`, `errorStyle: stop|warning|information`. `get`/`remove` by
`/Sheet1/dataValidation[i]`; `read --view structure` lists per-sheet rules.

## sparkline (M5, `/Sheet1/sparkline[1]`)

`{op:"add", path:"/Sheet1/E2", type:"sparkline", props:{dataRange:"A2:D2",
kind:"line|column|winLoss", color?:"376092", markers?:true}}` — the op path is
the host CELL; `markers` applies to line sparklines only. One sparkline per op.
`get /Sheet1/sparkline[i]`; structure lists `sparklineGroups`; remove by path.

## threaded comments (M5, `/Sheet1/comment[@id=GUID]`)

`{op:"add", path:"/Sheet1/B2", type:"comment", props:{text, author?}}` anchors
a thread root; `{op:"add", path:"/Sheet1/comment[@id=GUID]", type:"reply",
props:{text, author?}}` appends a reply (the GUID comes back from the add,
`read --view comments` or structure). Real `xl/threadedComments` parts plus a
legacy note fallback for old clients. Remove the root to drop the whole thread.

## csv bridge (M5)

`aioffice create orders.xlsx --from orders.csv` imports RFC 4180 csv (quoted
commas/newlines, sniffed `, ; tab |` delimiter, `--title` names the sheet):
numbers/dates/booleans are typed, leading-zero codes like `007` stay text,
>50k cells stream through the SAX writer. `aioffice read orders.xlsx --view csv
[--sheet NAME] [--range A1:C10]` exports one sheet back as csv.

## named cell styles (M7, `add type:cellStyle`, `read --view styles`)

Reusable, named bundles of formatting: define once, apply by name to any cell or
range. Addressed by `/style[@name=X]`; listed by `read --view styles`.

| prop         | type   | notes                                            |
|--------------|--------|--------------------------------------------------|
| name         | string | required; the style's unique name                |
| numberFormat | string | e.g. `"$#,##0.00"`, `"0.0%"`, `"yyyy-mm-dd"`     |
| bold, italic | bool   |                                                  |
| fill         | string | background hex, e.g. `"FFF2CC"`                   |
| color        | string | font hex, e.g. `"C00000"`                         |
| border       | string | none · thin · medium · thick · dashed · dotted · double · hair |

    {op:"add", path:"/styles", type:"cellStyle",
      props:{name:"Currency-Red", numberFormat:"$#,##0.00", color:"C00000", bold:true}}
    {op:"set", path:"/Sheet1/B2:B10", props:{cellStyle:"Currency-Red"}}   # apply to a range

Applying writes the concrete formatting onto the target cells (cached display
reflects the number format immediately) and also registers a real `cellStyle`
entry so Excel surfaces it in its gallery; the file stays validator-clean.
`get /style[@name=X]` reads a definition back; `remove /style[@name=X]` drops it.

## document properties (M7, `/properties`)

Core + custom workbook metadata on a virtual `/properties` node (same contract
as docx/pptx). `set /properties` writes; `get /properties` or
`read --view properties` returns `data.properties.{core:{title,…}, custom:{…}}`
(M10: the nested shape is now identical across all three formats — read the title
at `data.properties.core.title`). `title` is what `audit`'s `a11y_no_doc_title`
checks; `--fix` derives it from the filename.

    {op:"set", path:"/properties", props:{title:"Sales Metrics", author:"AIOffice",
      custom:{Quarter:"Q3", Final:true}}}

## slicers (M8)

A **slicer** (`type:slicer`) is a dashboard filter wired to a table column or a
pivot field. Authored on raw OpenXml (ClosedXML cannot create slicer parts), so
it validates clean and round-trips byte-identical.

| prop          | type   | meaning                                                       |
|---------------|--------|---------------------------------------------------------------|
| `column`      | string | table slicer: the column to slice (path is `/Sheet/table[@name=…]`) |
| `field`       | string | pivot slicer: the field to slice (path is `/Sheet/pivot[i]`)   |
| `caption`     | string | optional header text (default: the column/field name)         |
| `x`           | string | top-left anchor cell, e.g. `H2` (default: 2 cols right of data)|
| `widthCells`  | int    | box width in cells (default 2)                                |
| `heightCells` | int    | box height in cells (default 7)                               |

    # table-column slicer
    {op:"add", path:"/Sheet1/table[@name=Sales]", type:"slicer",
      props:{column:"Region", x:"H2"}}
    # pivot-field slicer
    {op:"add", path:"/Sheet1/pivot[1]", type:"slicer", props:{field:"Region"}}

`column`/`field` match case-insensitively; a wrong name is `invalid_args` with
the available names as candidates. Address a slicer `/Sheet1/slicer[1]` (1-based
per sheet) or `/Sheet1/slicer[@name=X]`. `get` reports `kind` (`table`/`pivot`),
`column`, `caption`; `read --view structure` lists per-sheet slicers; `remove`
by either path.

## audit (M7)

`aioffice audit metrics.xlsx [--fix]` lints accessibility + quality: cells whose
cached value is an Excel error (`quality_formula_error`, e.g. `#DIV/0!`/`#REF!`),
merged ranges inside a data region (`a11y_merged_data_cells`), images/charts with
no alt text (`a11y_no_alt_text`), and a missing workbook Title
(`a11y_no_doc_title`). `--fix` safely sets a placeholder alt text and a title;
formula errors and merged data cells are report-only. See `aioffice help audit`.

## Protection (v1.2)

Per-cell lock plus light sheet/workbook protection — Excel's UI guards (which
actions Excel allows), NOT encryption. aioffice always owns and can lift this
protection. `get` reflects the state.

- Per-cell `locked` (round-trips natively):

      {op:"set", path:"/Sheet1/A1:B2", props:{locked:false}}   # leave an editable window

- Sheet protection on a sheet path (`/Sheet1`):

      {op:"set", path:"/Sheet1", props:{protected:true}}
      {op:"set", path:"/Sheet1", props:{protected:true, password:"pw", allowSort:true}}

  Optional `allow*` flags relax a single action: `allowFormatCells`,
  `allowFormatColumns`, `allowFormatRows`, `allowInsertColumns`,
  `allowInsertRows`, `allowDeleteColumns`, `allowDeleteRows`, `allowSort`,
  `allowAutoFilter`, `allowPivotTables`, `allowSelectLockedCells`,
  `allowSelectUnlockedCells`. Cells default to locked, but the lock only bites
  once the sheet is protected — unlock a range first to leave it editable.

- Workbook structure protection on the root path (`/`):

      {op:"set", path:"/", props:{protectStructure:true}}        # also protectWindows, password

`get /Sheet1` shows `protected` and the relaxed actions; `get` on a cell shows
its `locked` state.

## numberFormat presets (v1.2)

The `numberFormat` cell prop accepts a named preset (e.g. `accounting-usd`,
`percent`, `date-iso`) in addition to a literal Excel format code. A preset
resolves to its code; any non-preset string is preserved verbatim. Full table:
`aioffice help number-formats`.

      {op:"set", path:"/Sheet1/B2", props:{numberFormat:"accounting-usd"}}

## Form controls (v1.2)

Interactive legacy form controls on a sheet path; the anchor cell goes in
`props.cell`. Kinds: `checkbox`, `optionButton`, `spinner`, `comboBox`,
`listBox`, `button`.

      {op:"add", path:"/Sheet1", type:"formControl", props:{kind:"checkbox", cell:"E2", linkedCell:"F2"}}
      {op:"add", path:"/Sheet1", type:"formControl", props:{kind:"comboBox", cell:"E4", linkedCell:"F4", items:["Red","Green","Blue"]}}
      {op:"add", path:"/Sheet1", type:"formControl", props:{kind:"spinner", cell:"E6", linkedCell:"F6", min:0, max:10, increment:1}}

`linkedCell` is where the control writes its value; `comboBox`/`listBox` need
`items` (a list) or `listFillRange` (a worksheet range). `read --view structure`
lists per-sheet form controls (1-based); `remove` by `/Sheet1/formControl[N]`.
