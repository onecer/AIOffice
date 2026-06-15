# scenarios — what-if scenario manager (xlsx, 1.5)

A **scenario** is a named set of *changing cells* plus the constant values they
take. It is stored in the worksheet's scenarios part (validator-clean) so Excel's
Scenario Manager shows it, and aioffice can *apply* it to write those values in and
recalculate dependents.

## add a scenario

```jsonc
{op:"add", path:"/Sheet1", type:"scenario",
 props:{name:"Best Case", cells:{B1:120, B2:0.15}, comment:"optimistic"}}
```

- `name` — required, unique within the sheet.
- `cells` — a map of `address -> value`; values are **constants** (numbers, text,
  booleans), not formulas.
- `comment` — optional note.

## apply a scenario

```jsonc
{op:"set", path:"/Sheet1", props:{applyScenario:"Best Case"}}
```

Writes the stored values into the changing cells and **recalculates** dependents,
so the saved file carries real cached values (not just the inputs).

## read / get / remove

```jsonc
{op:"remove", path:"/Sheet1/scenario[@name=Best Case]"}
```

- `get /Sheet1/scenario[@name=Best Case]` → `{name, comment, cells}`.
- `read --view structure` lists each sheet's scenarios.
- Quote names with specials: `scenario[@name='Q3 Plan']`.

This is additive behaviour on the frozen `add`/`set` verbs — no new verb or tool.

See also: `aioffice help goal-seek`, `aioffice help formulas`.
