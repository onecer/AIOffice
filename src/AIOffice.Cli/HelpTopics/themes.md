# themes (`/theme`) — v1.3

Every docx/pptx carries one theme part (`theme1.xml`) with a colour scheme
(`a:clrScheme`) and a font scheme (`a:fontScheme`) that styles reference by
name. Edit those slots with `set /theme`; read them back with `get /theme`.

## docx — `set /theme`

```jsonc
{op:"set", path:"/theme", props:{
  accent1:"38BDF8",        // a 6-hex RGB (with or without leading #)
  dk1:"1E293B", lt1:"FFFFFF",
  majorFont:"Calibri Light",
  minorFont:"Calibri"}}
```

Colour slots: `dk1`, `lt1`, `dk2`, `lt2`, `accent1`…`accent6`, `hlink`,
`folHlink` — each a 6-digit hex. Font slots: `majorFont` (headings), `minorFont`
(body). Any other key returns `unsupported_feature` naming the valid slots.

`get /theme` returns the colour scheme as a hex map plus `majorFont`/`minorFont`.
Changing `accent1` recolours every style and shape that binds to that theme slot.

```jsonc
// read it back
{op:"get", path:"/theme"}
// -> {colors:{accent1:"38BDF8", …}, majorFont:"Calibri Light", minorFont:"Calibri"}
```

(pptx theme editing lives on the slide master — see `aioffice help
properties-pptx`, the master/theme section.)
