# table-formulas — formula fields in a Word table (docx, 1.5)

A docx table cell can carry a **`formula`** prop. aioffice computes the value
headlessly *now*, caches it as the field result, and stores a Word table-formula
field (`w:fldSimple` whose instruction is the `=` expression), so
`read --view text`, `get`, and Word (before pressing F9) all show the same number.

## set a cell formula

```jsonc
{op:"set", path:"/body/table[1]/row[3]/cell[2]", props:{formula:"=SUM(ABOVE)"}}
{op:"set", path:"/body/table[1]/row[3]/cell[2]", props:{formula:"=A1*B2"}}
```

- **Directional aggregates** — `=SUM`, `=AVERAGE`, `=PRODUCT`, `=COUNT`, `=MIN`,
  `=MAX` over `(ABOVE | BELOW | LEFT | RIGHT)`.
- **Cell-ref arithmetic** — `=A1*B2`, `=(A1+A2)/2`, using the table's A1
  addressing (column letters across, 1-based rows down).

## number format

An optional `numberFormat` (valid **only** alongside `formula` on the same cell)
shapes the cached text and the field's stored picture:

```jsonc
{op:"set", path:"/body/table[1]/row[4]/cell[3]",
 props:{formula:"=SUM(ABOVE)", numberFormat:"currency"}}
```

Presets: `integer | number | percent | currency`, or a raw Word `\#` picture.
`numberFormat` on its own (no `formula`) → `invalid_args`.

## cached-value note

When an input cell is itself a field, the value is still cached but a
**`table_formula_cached`** warning flags that Word may refresh to a different
number on F9. Static cell values compute exactly.

See also: `aioffice help properties-docx` (docx/table), `aioffice help formulas`.
