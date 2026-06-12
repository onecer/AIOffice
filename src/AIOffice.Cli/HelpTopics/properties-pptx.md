# pptx properties

## slide (`/slide[2]`)

| prop       | type   | notes                                              |
|------------|--------|----------------------------------------------------|
| title      | string | `add`: creates the title shape with this text      |
| background | string | M2: hex RGB solid fill, a real `p:bg` (`set`/`add`)|

Add a slide after slide 1 (`position`: `at`/`before` = the new slide takes the
path's index, `after` = one past it):

    {op:"add", path:"/slide[1]", type:"slide", position:"after",
     props:{"title":"Q3 Results"}}

## shape (`/slide[2]/shape[3]`)

`add` with `type:"shape"` (or `"textbox"`) targets the **slide** path and
appends a textbox; more shape kinds are M1.

| prop     | type   | notes                                    |
|----------|--------|-------------------------------------------|
| text     | string | full text of the shape                    |
| title    | string | alias for text                            |
| x, y     | length | position from the top-left                |
| w, h     | length | size                                      |
| fill     | string | background color, hex RGB                 |
| fontSize | number | points                                    |
| bold     | bool   |                                           |
| color    | string | font color, hex RGB                       |
| align    | string | left · center · right · justify           |
| name     | string | shape name                                |

Lengths: a bare number means **centimeters** (`"x": 2.5`); strings take a unit
suffix: `"5cm"`, `"36pt"`, `"1in"`, `"96px"`, `"1800000emu"`.

## p (paragraph inside a shape, `/slide[2]/shape[3]/p[1]`)

| prop     | type   | notes                          |
|----------|--------|--------------------------------|
| text     | string |                                |
| fontSize | number | points                         |
| bold     | bool   |                                |
| color    | string | hex RGB                        |
| align    | string | left · center · right · justify |

## notes (M2, `/slide[2]/notes`)

The path addresses the whole speaker-notes body (no `/p[j]` beneath it).

    {op:"set",    path:"/slide[2]/notes", props:{text:"line1\nline2"}}   # replace (\n = new paragraph)
    {op:"add",    path:"/slide[2]/notes", props:{text:"follow-up"}}      # append a paragraph
    {op:"remove", path:"/slide[2]/notes"}                                # clear

`read` (outline) and `get /slide[2]/notes` read them back.

## image (M2)

| prop       | type   | notes                                            |
|------------|--------|--------------------------------------------------|
| src        | string | required; PNG/JPEG path inside the sandbox       |
| x, y, w, h | length | omit w/h to keep natural size/aspect             |
| name       | string | optional picture name                            |

`{op:"add", path:"/slide[1]", type:"image", props:{src:"logo.png", w:"6cm"}}` —
the result path is `/slide[1]/shape[@id=N]`; `remove` it by that path.

## chart (M3, `/slide[1]/chart[1]`)

`{op:"add", path:"/slide[1]", type:"chart", props:{kind:"bar|line|pie",
categories:["Q1","Q2"], series:[{name:"Sales", values:[10,20]}],
title?, x?, y?, w?, h?}}`. Data is cached literally in the chart XML (no
embedded workbook): every projection reports `dataEditable:false` and the add
attaches a warning. Cross-document data (M3): replace categories/series with
`{"dataFrom":"book.xlsx!Sheet1/A1:B5"}` — first column = categories, header
row = series names, remaining columns = series values (quote sheet names with
spaces: `book.xlsx!'Q3 Data'/A1:C5`). `set` on `/slide[i]/chart[k]` retitles;
`remove` deletes frame + part.

## transition (M3, slide-level set)

`{op:"set", path:"/slide[2]", props:{transition:"fade|push|wipe|none",
transitionDuration:"0.5s"}}` — `get /slide[2]` and the outline view read it
back.

## geometry & z-order (M3)

`add type:shape` accepts `shape:"rect|roundRect|ellipse|triangle|diamond|
arrow|line"` plus `flipH`/`flipV` ("line" builds a connector; `fill` sets its
stroke). Z-order: `{op:"move", path:"/slide[1]/shape[@id=5]",
position:"front|back|forward|backward"}` reorders the paint order (front =
topmost).

## table (M5, `/slide[1]/table[1]`)

`{op:"add", path:"/slide[1]", type:"table", props:{rows:3, cols:4,
headerRow?:true, style?:"light|medium|dark", x?, y?, w?, columnWidths?:[…]}}` —
styles are direct paint (header fill + banded rows), no theme dependency.
Cells: `{op:"set", path:"/slide[1]/table[1]/tr[1]/tc[1]", props:{text,
mergeRight?:2, mergeDown?:2}}` — pptx merge counts are the cells to **absorb**
(`mergeRight:1` → span 2; docx counts the total span instead). `get` the table
for rows/cols/headerRow plus per-cell paths and merge shape; `render --to svg`
draws the real grid; `remove` by the table path.

## animation (M4 entrance, M5 emphasis + exit; `add` on the shape path)

`{op:"add", path:"/slide[1]/shape[@id=4]", type:"animation",
props:{effect, trigger?, duration?, delay?, direction?, color?}}`

| class    | effects                          | extras                            |
|----------|----------------------------------|-----------------------------------|
| entrance | appear · fade · flyIn · wipe     | direction: flyIn/wipe             |
| emphasis | pulse · grow · spin · colorPulse | color: colorPulse only            |
| exit     | fadeOut · flyOut · wipeOut       | direction: flyOut/wipeOut         |

Triggers: `click` (default) · `withPrevious` · `afterPrevious`; durations like
`"0.5s"`. `read --view structure` lists per-slide animation order; remove
`/slide[i]/animation[k]` to drop one.

## comment + reply (M2 comments, M5 threads)

`{op:"add", path:"/slide[2]", type:"comment", props:{text, author?}}` adds a
comment; `{op:"add", path:"/slide[2]/comment[@id=1]", type:"reply",
props:{text, author?}}` threads a reply under it (p15 threadingInfo —
PowerPoint 2013+ shows real threads). `read --view comments` lists threads;
removing the root removes its replies too.

## SmartArt (M5, read-only)

SmartArt diagrams surface in `read --view structure` (per-slide `smartArt`
rows) and `get /slide[i]/smartart[k]` as nested node trees in connection
order. Editing SmartArt is `unsupported_feature` with the workaround named.

## Rendering

`aioffice render deck.pptx --to svg --scope /slide[2] -o slide2.svg` renders
one slide; `--to html` renders a simple HTML projection; `--to png` screenshots
one slide via a local Chrome/Edge (default `/slide[1]` — pass `--scope`);
`--to pdf` (M3) prints the WHOLE deck to one PDF, one page per slide
(`--scope /slide[N]` narrows it to a single page).

Master/layout editing is still `unsupported_feature`.
