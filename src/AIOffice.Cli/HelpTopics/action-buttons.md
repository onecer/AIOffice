# action-buttons — navigation buttons (pptx, 1.5)

An **action button** is a navigation shape (building on M8 shape hyperlinks): a
`p:sp` with a preset action-button geometry plus a `ppaction://` click verb on its
`cNvPr`. The SVG render draws the button glyph.

## add a button

```jsonc
{op:"add", path:"/slide[1]", type:"actionButton", props:{action:"next", x:"2cm", y:"2cm"}}
```

- `action` — one of:
  - `first | last | next | prev | home | end` — a show-jump (no relationship);
  - `slide` with `target:"slide N"` — a slide-jump (a real relationship, so
    PowerPoint navigates);
  - `url` with `target:"https://…"` — an external link.
- `x` / `y` / `w` / `h` — optional position and size (lengths).
- `label` — optional caption text.
- `fill` — optional hex fill.
- `name` — optional shape name.

## get / remove

- `get` on the button's shape path reports the resolved `action` + `target`.
- `remove` deletes the button.

Validator-clean for every action. Additive on the frozen `add`/`get`/`remove`
verbs.

See also: `aioffice help layouts`, `aioffice help properties-pptx`.
