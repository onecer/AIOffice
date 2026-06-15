# line-numbers — section line numbering (docx, 1.5)

Word can print line numbers down the margin, configured per section
(`w:lnNumType`). Set the **`lineNumbers`** prop on a section.

## turn line numbering on

```jsonc
{op:"set", path:"/section[1]",
 props:{lineNumbers:{start:1, increment:1, restart:"continuous"}}}
```

- `start` — the first printed number (`>= 0`, default `1`).
- `increment` — print every Nth line (`>= 1`, default `1`).
- `restart` — `continuous` (default) | `newPage` | `newSection`.
- `distance` — optional gap from the text to the numbers (a length).

## turn it off

```jsonc
{op:"set", path:"/section[1]", props:{lineNumbers:"none"}}
```

`get /section[1]` reads the current setting back. Validator-clean. Additive on the
frozen `set` verb.

See also: `aioffice help sections`, `aioffice help page-borders`.
