# sections

"Section" means two different things across formats. Both are M-series features
and both are now fully editable.

## docx sections — page setup + columns (`/section[1]`, M3 + M6)

A docx section owns page geometry and (M6) newspaper-style columns. The
implicit single section is materialized on demand at `/section[1]`.

| prop                        | notes                                        |
|-----------------------------|----------------------------------------------|
| pageSize                    | A4 · Letter · Legal · A3 · A5                |
| orientation                 | portrait · landscape (width/height swap)     |
| marginTop/Bottom/Left/Right | unit-qualified, e.g. "2cm"                   |
| columns                     | M6: column count (1 clears multi-column)     |
| columnGap                   | M6: gap between columns, e.g. "1.25cm"       |

    aioffice edit r.docx --ops '[{"op":"set","path":"/section[1]","props":{"columns":2}}]'
    aioffice get r.docx /section[1]        # -> "columns": 2, "columnGapCm": 1.27

Add a new section break with `{op:"add", type:"sectionBreak", props:{kind}}`;
push following text into the next column with `{op:"add",
type:"columnBreak"}`.

## pptx slide sections — outline grouping (`/section[i]`, M6)

PowerPoint slide sections group slides in the outline / section view. They are
managed on the presentation root (`/`).

    aioffice edit deck.pptx --ops '[{"op":"add","path":"/","type":"section","props":{"name":"Intro","afterSlide":0}}]'
    aioffice edit deck.pptx --ops '[{"op":"add","path":"/","type":"section","props":{"name":"Body","afterSlide":1}}]'

- `name` (required) is the display name.
- `afterSlide` is 0-based: `0` starts the section before slide 1, `N` starts it
  after slide N. The new section claims the still-unsectioned slides from there
  up to the next section's first slide.

`set /section[i] {name}` renames; `remove /section[i]` drops the section (its
slides survive, just unsectioned). `read --view outline` groups slides under
their sections; `get /section[i]` returns the name and the slides it owns.
Sections track slides by `p:sldId`, so they survive slide reordering.
