# print-setup (1.7 — xlsx print completeness + calculation settings)

v1.7 rounds out the xlsx page-setup surface so a workbook is print-ready from the
CLI/MCP alone. All print props are set on a **sheet path** (`/Sheet1`); the
calculation settings are set on the **workbook root** (`/`). The M3/v1.4 print
props (`orientation`, `paperSize`, `margins`, `fitToWidth`, `printArea`,
`scale`) still apply — these are additive.

## Repeat titles (rows/cols on every page)

    aioffice edit book.xlsx --ops '[{"op":"set","path":"/Sheet1","props":{"printTitleRows":"1:1"}}]'
    aioffice edit book.xlsx --ops '[{"op":"set","path":"/Sheet1","props":{"printTitleCols":"A:A"}}]'

- `printTitleRows` is a 1-based row band (`"1:1"`, `"1:3"`); `printTitleCols` is
  a column-letter band (`"A:A"`, `"A:C"`).
- An empty string (`""`) clears that band.

## Manual page breaks

    aioffice edit book.xlsx --ops '[{"op":"set","path":"/Sheet1","props":{"pageBreaks":{"rows":[20,40],"cols":["F"]}}}]'

- `rows` = 1-based row numbers a break sits **above**; `cols` = column letters a
  break sits **left of**. An empty array on an axis clears that axis's breaks.

## Fit to page

    aioffice edit book.xlsx --ops '[{"op":"set","path":"/Sheet1","props":{"fitToPage":{"fitToWidth":1,"fitToHeight":0}}}]'
    aioffice edit book.xlsx --ops '[{"op":"set","path":"/Sheet1","props":{"fitToPage":{"scale":80}}}]'

- `fitToPage` takes EITHER `{fitToWidth,fitToHeight}` (page counts; `0` = that
  axis automatic) OR `{scale:N}` (10–400 percent) — Excel keeps only one.
- The flat `fitToWidth` / `fitToHeight` / `scale` props (M3) still work too.

## Centering on the page

    {"centerHorizontally":true,"centerVertically":true}

## Print headers & footers

    aioffice edit book.xlsx --ops '[{"op":"set","path":"/Sheet1","props":{"printFooter":{"center":"Page &P of &N"}}}]'
    aioffice edit book.xlsx --ops '[{"op":"set","path":"/Sheet1","props":{"printHeader":{"left":"&F","center":"&A","right":"&D"}}}]'

- `printHeader` / `printFooter` each take a `{left,center,right}` object; any
  subset is allowed. A section sent as `null`/`""` clears it; an omitted section
  is left untouched.
- Excel field codes pass through verbatim: `&P` page, `&N` pages, `&D` date,
  `&T` time, `&F` file, `&A` sheet.
- `printGridlines` / `printHeadings` (booleans) toggle printing the gridlines
  and the row/column headings.

## Workbook calculation settings (root path)

    aioffice edit book.xlsx --ops '[{"op":"set","path":"/","props":{"calculationMode":"manual","iterativeCalc":true,"maxIterations":100,"maxChange":0.001}}]'

- `calculationMode`: `auto` | `manual` | `autoExceptTables`.
- `iterativeCalc` (bool) enables iterative (circular-reference) calculation;
  `maxIterations` (default 100) and `maxChange` (default 0.001) bound it.
- `fullPrecision` (bool) toggles precision-as-displayed off/on.
- These ride the same workbook-root set as `protectStructure`; one root op can
  carry both. `get /` reflects them back.

## Read it back

`get /Sheet1` reports the page-setup block (titles, breaks, fit, centering,
header/footer). `get /` reports the workbook calculation settings.
