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

## sheet (add)

`{op:"add", path:"/'New Sheet'", type:"sheet"}` creates a sheet. Sheet names
with spaces are quoted in paths: `/'Q3 Data'/B2`.

## chart (M1, `/Sheet1/chart[1]`)

`{op:"add", path:"/Sheet1", type:"chart", props:{kind:"bar|line|pie",
dataRange:"A1:C6", anchor:"E3", title?}}`.

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
