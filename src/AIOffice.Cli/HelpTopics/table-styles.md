# table-styles — built-in table looks + banding (pptx, 1.4)

When you add a `table` to a slide, give it a built-in **style** and banding/edge
options so it picks up a polished look from the theme instead of flat cells.

```jsonc
{op:"add", path:"/slide[2]", type:"table",
 props:{rows:"4", cols:"3", style:"medium2", firstRow:"true", bandRow:"true"}}
```

## style

`style` is a stock table-style preset (mapped to PowerPoint's built-in
`a:tableStyleId` GUIDs):

| group | presets |
|-------|---------|
| none | `none` (no built-in style) |
| light | `light1` · `light2` |
| medium | `medium1` · `medium2` · `medium3` |
| dark | `dark1` · `dark2` |

## style options

Banding and edge emphasis are independent boolean flags:

| flag | effect |
|------|--------|
| `firstRow` | emphasize the header row |
| `lastRow` | emphasize the total row |
| `bandRow` | alternating row shading (banded rows) |
| `firstCol` | emphasize the first column |

`headerRow:"true"` is a shortcut that turns on `firstRow`. Every style + option
combination validates OpenXmlValidator-clean. `read --view structure` and
`get /slide[i]/table[k]` report the applied style and flags.
