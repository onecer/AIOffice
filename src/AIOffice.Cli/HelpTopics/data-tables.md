# data-tables — what-if data tables (xlsx, 1.4)

A what-if **data table** recomputes a corner formula across a row and/or column of
input values. Add it with `type:"dataTable"` over a range whose top-left cell holds
the formula to analyze.

```jsonc
// One-variable column table: B1 holds =PMT(rate/12, nper, -principal);
// A2:A6 hold candidate rates; colInput points at the formula's rate cell.
{op:"add", path:"/Sheet1/A1:B6", type:"dataTable", props:{colInput:"$C$1"}}

// Two-variable table: the corner formula varies with both a row and a column input.
{op:"add", path:"/Sheet1/A1:E6", type:"dataTable", props:{rowInput:"$C$1", colInput:"$C$2"}}
```

## props

| prop | meaning |
|------|---------|
| `rowInput` | the cell the **row** of inputs feeds (two-variable, or a row table) |
| `colInput` | the cell the **column** of inputs feeds (two-variable, or a column table) |

- The **corner cell** (top-left of the range) must already hold the formula to
  analyze; otherwise the op fails with a suggestion to set it first.
- The body cells are written with cached results (a full recalc per probe), so a
  headless reader sees the analyzed values; Excel keeps the live `{=TABLE(...)}`
  array on open.
- `validate` is clean. `read --view structure` lists the data tables on a sheet;
  `remove /Sheet1/dataTable[k]` drops one.

See also: `aioffice help formulas` (dynamic arrays + financial functions).
