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

## chart (M1+M3, `/Sheet1/chart[1]`)

`{op:"add", path:"/Sheet1", type:"chart", props:{kind:"bar|line|pie|scatter|area",
dataRange:"A1:C6", anchor:"E3", title?}}`. scatter (M3) plots numbers against
numbers: the first dataRange column is the numeric X axis, so every X cell
must be a number (header row excepted).

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
