# embedded-fonts — embed a font file in a deck (pptx, 1.5)

A deck can carry its own font files so it renders the same on a machine that does
not have them installed. Each embedded font is a font part referenced from a
`p:font` slot inside a `p:embeddedFont` in the presentation's `p:embeddedFontLst`.

## embed a font

```jsonc
{op:"add", path:"/fonts", type:"font", props:{src:"MyFont.ttf", name:"My Font"}}
```

- `src` — **required**, a `.ttf`/`.otf` file **inside the workspace**. aioffice
  cannot pull a system font — you must supply the file. The media type is sniffed
  (ttf/otf). An `src` that escapes the workspace is `sandbox_denied` and is never
  read.
- `name` — the typeface name (defaults to the font file name).

By default only the **regular** slot is embedded.

## all four styles

```jsonc
{op:"add", path:"/fonts", type:"font",
 props:{src:"reg.ttf", embedAll:"true", bold:"b.ttf", italic:"i.ttf", boldItalic:"bi.ttf"}}
```

`embedAll` plus the per-style files fills the `regular`/`bold`/`italic`/`boldItalic`
slots.

## list / get / remove

- `get /fonts` lists every embedded font; `get /fonts/font[@name=My Font]` reports
  one.
- `remove /fonts/font[@name=My Font]` drops the registration and its font parts.

Validator-clean. Additive on the frozen `add`/`get`/`remove` verbs.

See also: `aioffice help action-buttons`, `aioffice help properties-pptx`.
