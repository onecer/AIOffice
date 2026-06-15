# formulas — dynamic arrays + scalar + financial functions (xlsx, 1.4/1.5)

Set a cell's `value` to a formula and aioffice evaluates three families that the
underlying engine cannot, writing the cached result(s) back into the saved file.

## dynamic arrays (spill)

`=FILTER`, `=UNIQUE`, `=SORT`, `=SORTBY`, `=SEQUENCE`, `=RANDARRAY` and
`=TRANSPOSE` are **evaluated and spilled** (1.4). Before 1.4 they were recognised
but left a `formula_not_evaluated` warning; now the anchor cell carries the
formula and every cell of the result rectangle carries its cached value.

```jsonc
{op:"set", path:"/Sheet1/D1", props:{value:"=UNIQUE(A1:A20)"}}   // spills distinct values down
{op:"set", path:"/Sheet1/D1", props:{value:"=SORT(A1:A20)"}}     // spills sorted (=SORT(A1:A20,1,-1) descending)
{op:"set", path:"/Sheet1/D1", props:{value:"=FILTER(A1:B20,C1:C20)"}} // keeps the rows whose flag column C is TRUE
{op:"set", path:"/Sheet1/D1", props:{value:"=SEQUENCE(2,3)"}}    // spills 1..6 over D1:F2
```

The `include` argument of `FILTER` is a **boolean range** (a TRUE/FALSE column of the
same height), not an inline inequality — compute the flags into a helper range first
(e.g. set `C1` to `=B1>100`, fill down), then `=FILTER(A1:B20,C1:C20)`. Range
references must be explicit (`A1:A20` or `Sheet!A1:A20`); whole-column refs like `A:A`
are not supported.

- `get /Sheet1/D1` on the anchor shows the formula and the spill range; the spilled
  cells read their cached value.
- The anchor is written as a compatible array formula, so Excel re-spills on open.
- **`spill_blocked`** (1.4 error): if the result rectangle would overwrite a
  non-empty cell, nothing is written — clear the target range first, then set the
  formula. The suggestion names the exact range to clear.

`=TEXTSPLIT(text, col_delim, [row_delim], …)` joins this family in 1.5: it splits
text into a spilled array (1-D by the column delimiter, 2-D with a row delimiter);
the text and delimiters may be string literals or cell references.

## scalar functions (1.5)

`=XLOOKUP`, `=IFS`, `=SWITCH`, `=LET`, `=MAXIFS`, `=MINIFS` and `=AVERAGEIFS` —
which the base engine returns `#NAME?` for — are now **evaluated** at write time
and the cached scalar value is written, so headless readers see a real result with
no `formula_not_evaluated` warning.

```jsonc
{op:"set", path:"/Sheet1/D2", props:{value:"=XLOOKUP(\"Widget\",A1:A20,B1:B20)"}} // exact match (default)
{op:"set", path:"/Sheet1/D3", props:{value:"=IFS(A1>90,\"A\",A1>80,\"B\",TRUE,\"C\")"}}
{op:"set", path:"/Sheet1/D4", props:{value:"=LET(x,5,x*2)"}}                        // binds x then evaluates -> 10
{op:"set", path:"/Sheet1/D5", props:{value:"=TEXTJOIN(\",\",TRUE,A1:A5)"}}
```

`XLOOKUP` does exact (default) or approximate (`match_mode` -1/1) match and returns
`if_not_found` on a miss; `LET` binds names left-to-right then evaluates the
calculation; the conditional aggregates honour numeric/text criteria operators.
`TEXTJOIN`, `CONCAT`, `CONCATENATE`, `IFERROR`, `SUMIFS` and `COUNTIFS` were already
evaluated by the base engine and keep working.

Still **not** evaluated (an honest `formula_not_evaluated` warning, never a wrong
value): stored `LAMBDA` and the lambda-helpers `MAP`/`REDUCE`/`SCAN`/`BYROW`/`BYCOL`
— Excel computes them on open.

## financial functions

`RATE`, `IRR`, `XIRR`, `NPV`, `PV`, `FV`, `PMT`, `NPER` are **evaluated** at save
time (iterative ones converge numerically). A `=IRR(A1:A6)` cell carries a cached
numeric value instead of a `formula_not_evaluated` warning.

```jsonc
{op:"set", path:"/Sheet1/B10", props:{value:"=IRR(B1:B6)"}}      // cached rate of return
{op:"set", path:"/Sheet1/B11", props:{value:"=PMT(B2/12,B3,-B1)"}}
```

## backward compatibility

This is a behaviour **improvement**, not a contract change: cells that used to
carry a `formula_not_evaluated` warning now carry a cached value, and the warning
no longer fires for these functions. An agent that handled the warning still
works — the warning simply stops appearing for the now-evaluated functions.

See also: `aioffice help data-tables` (what-if data tables),
`aioffice help scenarios` (scenario manager) and `aioffice help goal-seek`.
