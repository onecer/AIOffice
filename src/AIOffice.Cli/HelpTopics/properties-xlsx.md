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
| gradientFill| object               | 1.8: gradient fill, see **gradient fill** below|
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

1.3 adds **chart-polish** props accepted both on `add` and on `set
/Sheet1/chart[i]`: `dataLabels`, `legend`, `axisTitles`, `trendline`,
`errorBars`, `gridlines`, `secondaryAxis`. `get` reports them. Full grammar +
examples: `aioffice help chart-polish`.

1.8 adds **`seriesColors`** — a brand palette accepted on both `add` and `set
/Sheet1/chart[i]`. Pass an array of 6-hex RGB strings (a leading `#` is
optional), one per series in dataRange order; a short list cycles across the
series:

      {op:"add", path:"/Sheet1", type:"chart",
       props:{kind:"bar", dataRange:"A1:C5", anchor:"E2",
              seriesColors:["2E5AAC", "#E8743B"]}}

Bar/area series get a solid fill in the color; line/scatter series get a tinted
line; a single pie/doughnut series colors each *slice* (the palette cycles across
the data points). `get` on the chart reports the applied colors back under
`polish.seriesColors`. Because chart colors live in the chart part (which
ClosedXML preserves byte-identical), they survive later edits.

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
| tabColor                | sheet       | hex tab color (**1.24**); `''` clears |
| margins                 | sheet       | print margins in **inches** (**1.25**): `{top?,bottom?,left?,right?,header?,footer?}`; any subset |
| autoFilter              | range       | bool toggles the dropdowns; an object/array applies real criteria (**1.12**, see **autofilter criteria**) |
| orientation, paperSize  | sheet       | landscape/portrait; A4/Letter/…    |
| fitToWidth, fitToHeight | sheet       | print scaling, pages               |
| printArea               | sheet       | e.g. "A1:F40"                      |
| printTitleRows/Cols     | sheet       | **1.7**: repeat bands, e.g. "1:1" / "A:A" |
| pageBreaks              | sheet       | **1.7**: `{rows:[20],cols:["F"]}`  |
| fitToPage               | sheet       | **1.7**: `{fitToWidth,fitToHeight}` or `{scale}` |
| centerHorizontally/Vertically | sheet | **1.7**: bools                     |
| printGridlines/printHeadings  | sheet | **1.7**: bools                     |
| printHeader/printFooter | sheet       | **1.7**: `{left,center,right}` field-code strings |
| calculationMode, iterativeCalc, maxIterations, maxChange, fullPrecision | **`/` (workbook root)** | **1.7**: `manual`/`auto`/`autoExceptTables` + iteration settings |

Full grammar + examples: `aioffice help print-setup`.

`get /Sheet1` reflects freeze/autoFilter/pageSetup; `get /` reflects the
workbook calculation settings. Streaming reads (M3): files
over 20 MB (or `stream:true`) answer `read --view stats|text` and cell/range
`get` via a SAX scan without loading the workbook DOM.

## autofilter criteria (1.12, on a range path)

The `autoFilter` sheet-prop accepts a **criteria** object/array as well as the M3
bool. A criteria filter is *applied*: the non-matching rows are hidden headlessly
(a real filtered view a headless reader sees) and Excel re-applies it on open.

| form                                   | meaning                                      |
|----------------------------------------|----------------------------------------------|
| `{autoFilter:true\|false}`             | enable / clear the filter (unchanged)        |
| `{autoFilter:{column, values:[…]}}`    | **values** filter: keep rows whose column equals one of the listed values |
| `{autoFilter:{column, criteria:"…"}}`  | **comparison/text** filter (see grammar)     |
| `{autoFilter:[{column,…}, …]}`         | several columns, ANDed together              |

`column` resolves against the filter range as a **header name** (the first row of
the range), a **column letter** inside the range, or a **1-based index** within
the range; an unresolved name is `invalid_args` with the headers as candidates.

`criteria` grammar — a leading operator then a value, or wildcard text:

| criteria      | keeps rows where the column …                       |
|---------------|-----------------------------------------------------|
| `">100"`      | is greater than 100 (`>= <= < >` and `=` likewise)  |
| `"<=0"`       | is at most 0                                         |
| `"<>0"` / `"!=0"` | is not equal to 0                               |
| `"*text*"`    | contains `text` (wildcards: `*pre`, `suf*`, `*mid*`)|
| `"East"`      | equals `East` (a bare value with no operator)       |

      # values filter on a header column
      {op:set, path:/Sheet1/A1:D20, props:{autoFilter:{column:"Region", values:["East","West"]}}}
      # comparison filter on another column (index/letter also work)
      {op:set, path:/Sheet1/A1:D20, props:{autoFilter:{column:"Amount", criteria:">100"}}}
      # two columns at once (ANDed)
      {op:set, path:/Sheet1/A1:D20, props:{autoFilter:[
        {column:"Region", values:["East"]}, {column:"Amount", criteria:">=200"}]}}

`get /Sheet1` reports the active filters under **`autoFilterColumns`** (per column:
`kind` `values`/`custom`, the `values`, or the `criteria` operator + value); a
plain bool-enabled filter omits it. `{autoFilter:false}` clears the criteria and
un-hides the rows.

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
| values      | array  | `[{field:"Sales", agg:"sum|average|count|min|max", showAs?, baseField?, baseItem?}]` |
| calculatedFields | array | (1.3) `[{name:"Margin", formula:"=Revenue-Cost"}]` — formula fields computed from source headers |

`{op:"add", path:"/Sheet1", type:"pivot", props:{...}}` — the op path is the
SOURCE sheet. `get`/`remove` address the pivot on its TARGET sheet:
`/Pivot/pivot[1]` or `/Pivot/pivot[@name=SalesPivot]`. Excel recomputes the
pivot on open (refreshOnLoad).

`calculatedFields` (1.3) add formula columns to the pivot: each `formula`
references other source-column header names (e.g. `=Revenue-Cost`,
`=Profit/Revenue`). They are validated against the headers at add time and
surface under `get`'s `calculatedFields`; Excel recomputes them on open.

`showAs` (1.13) sets a value field's **show-values-as** (the dataField's
`showDataAs` is written so Excel computes it on open; `get` reports it back).
Supported modes:

| showAs                  | base needed              | meaning                                  |
|-------------------------|--------------------------|------------------------------------------|
| `normal` (default)      | —                        | the raw aggregate                        |
| `percentOfTotal`        | —                        | % of the grand total                     |
| `percentOfColumn`       | —                        | % of the column total                    |
| `percentOfRow`          | —                        | % of the row total                       |
| `index`                 | —                        | Excel's index ((cell×grand)/(row×col))   |
| `runningTotal`          | `baseField`              | running total along the base field       |
| `percentOf`             | `baseField` + `baseItem` | % of a base item's value                 |
| `differenceFrom`        | `baseField` + `baseItem` | difference from a base item              |
| `percentDifferenceFrom` | `baseField` + `baseItem` | % difference from a base item            |

`baseField` is a source header; `baseItem` is a literal value of that field, or
the `"(previous)"` / `"(next)"` sentinel. The same field may appear as several
value fields differing only by `showAs` (e.g. a raw sum *and* its `percentOfTotal`).
An unknown `showAs` is `invalid_args` (with the modes as candidates); a mode
OOXML cannot carry — `percentOfParentTotal`, `rankAscending`, `rankDescending` —
is `unsupported_feature` (never silently ignored).

    # show Sales as % of grand total, plus a running total down the Month axis
    {op:add, path:/Sheet1, type:pivot, props:{sourceRange:"A1:C7", targetSheet:"Pivot",
      rows:["Region"], columns:["Month"], values:[
        {field:"Sales", agg:"sum", showAs:"percentOfTotal"},
        {field:"Sales", agg:"sum", showAs:"runningTotal", baseField:"Month"},
        {field:"Sales", agg:"sum", showAs:"differenceFrom", baseField:"Month", baseItem:"(previous)"}]}}

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
| formula      | formula (`=$B1>100`), fill?, color?, bold? (1.3) |
| topBottom    | mode (`top`/`bottom`), rank (N), percent?, fill?, color?, bold? (1.3) |
| aboveBelowAverage | mode (`above`/`below`/`aboveOrEqual`/`belowOrEqual`), stdDev?, fill?, color?, bold? (1.3) |

`iconSet` (1.1) ranks each cell against the others in the range and paints a
3/4/5-icon glyph; pick the icon family with `set` (3-icon families start with
`3`, etc.), `reverse:true` flips the order, `showValue:false` hides the number.

The 1.3 kinds: `formula` evaluates an `=expression` (relative to the range's
top-left cell) per cell; `topBottom` paints the highest/lowest N (or N% with
`percent:true`); `aboveBelowAverage` paints cells above/below the range mean
(`stdDev` shifts the threshold). Full grammar + examples: `aioffice help
conditional-format`.

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

## linkedPicture (1.7, `/Sheet1/linkedPicture[1]` — the camera tool)

Mirrors a cell range as a picture (Excel's camera tool). aioffice renders a
**static snapshot** of the range to a PNG and embeds it, raising a
`linked_picture_static` warning (the picture is the values as of this edit, not
a live mirror). The source range/anchor are recorded in a workbook registry so
`get`/`remove` of `/Sheet1/linkedPicture[i]` distinguish it from a plain image.

| prop        | type   | notes                                                  |
|-------------|--------|--------------------------------------------------------|
| sourceRange | string | required; the range to snapshot, e.g. `A1:C5`          |
| anchor      | string | required; top-left cell where the picture lands, e.g. `G2` |
| sheet       | string | optional; sheet the source range lives on (default: the target sheet) |
| name        | string | optional picture name (also the ordering key)          |

`{op:"add", path:"/Sheet1", type:"linkedPicture", props:{sourceRange:"A1:C5", anchor:"G2"}}`.

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
commas/newlines, sniffed `, ; tab` delimiter, `--title` names the sheet):
numbers/dates/booleans are typed, leading-zero codes like `007` stay text,
>50k cells stream through the SAX writer. `aioffice read orders.xlsx --view csv
[--sheet NAME] [--range A1:C10] [--delimiter tab|;|,]` exports one sheet back as
csv (`--delimiter` **1.23** emits TSV/semicolon; default comma).

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

1.8 adds scaled-money / scaled-number presets (`usd-millions`, `usd-thousands`,
`millions`, `thousands-k`), `accounting-eur` / `accounting-gbp`, and explicit
`percent-0` / `percent-1` / `percent-2`.

## gradient fill (1.8, cell / range)

The `gradientFill` prop fills a cell or range with a color gradient — for KPI
bands, header strips, and brand surfaces. It is an object:

      {op:"set", path:"/Sheet1/A1", props:{gradientFill:{
        type:"linear", angle:90,
        stops:[{color:"#2E5AAC", pos:0}, {color:"#E8743B", pos:1}]}}}

- `type` (optional, default `linear`): `linear`, `radial`, or `path`. `radial`
  and `path` author a centered gradient and take no `angle`.
- `angle` (linear only): degrees — `0` runs left→right, `90` top→bottom.
- `stops` (required, ≥ 2): each `{color, pos}` — `color` is a 6-hex RGB (a
  leading `#` is optional), `pos` is the fractional position in `0..1`.

`get` on the cell reports the gradient back under `gradientFill`. The cell also
keeps a solid fill of the first stop's color as a fallback, so its number
format / font / borders are preserved and a brand color always shows.

Round-trip caveat (honest): a gradient survives the edit that authors it, but a
*later* edit that re-saves the workbook through ClosedXML drops the cell back to
that solid fallback color (ClosedXML cannot re-emit a gradient it does not
model). Re-apply the gradient in the same batch as any later change to keep it.

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
