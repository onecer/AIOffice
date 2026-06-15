# goal-seek — solve for an input (xlsx, 1.5)

Goal Seek finds the value of a *changing cell* that makes a dependent *formula
cell* equal a target. aioffice solves it headlessly (Newton's method with a
bisection fallback), then **sets** the changing cell to the found value and
recalculates.

## set props.goalSeek

```jsonc
{op:"set", path:"/Sheet1/B1",
 props:{goalSeek:{targetCell:"B5", targetValue:1000}}}
```

- `path` (`/Sheet1/B1`) — the **changing cell** (the input to solve for).
- `targetCell` (`B5`) — a cell holding a formula that **depends on** the changing
  cell.
- `targetValue` — the number you want `targetCell` to reach.

On success the op result reports `{converged:true, input, achievedTarget}` and the
changing cell is left holding the found value.

## no solution

If the solver cannot converge it raises a **`goal_seek_no_solution`** warning (not
a hard error) and the changing cell is left **unchanged**. Check that `targetCell`
actually holds a formula depending on the changing cell and that a solution exists.

This is a compute action that persists the found input — additive behaviour on the
frozen `set` verb, no new verb.

See also: `aioffice help scenarios`, `aioffice help formulas`.
