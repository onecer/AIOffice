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
| altText  | string | M7: accessibility description (`a:cNvPr/@descr`); any shape kind; `""` clears |
| altTitle | string | M7: accessibility title (`a:cNvPr/@title`); any shape kind; `""` clears |

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

## chart (M3, 1.1 expanded, `/slide[1]/chart[1]`)

`{op:"add", path:"/slide[1]", type:"chart", props:{kind:"bar|line|pie",
categories:["Q1","Q2"], series:[{name:"Sales", values:[10,20]}],
title?, x?, y?, w?, h?}}`. Data is cached literally in the chart XML (no
embedded workbook): every projection reports `dataEditable:false` and the add
attaches a warning. Cross-document data (M3): replace categories/series with
`{"dataFrom":"book.xlsx!Sheet1/A1:B5"}` — first column = categories, header
row = series names, remaining columns = series values (quote sheet names with
spaces: `book.xlsx!'Q3 Data'/A1:C5`). `set` on `/slide[i]/chart[k]` retitles;
`remove` deletes frame + part.

1.1 adds the kinds `doughnut`, `radar`, `bubble`, `stackedBar`,
`percentStackedBar`, `stackedArea`, `combo`. `doughnut` (like `pie`) needs
exactly one series; `bubble` series carry `{name, values, x?, size?}` triples
(one entry per category); `combo` draws the first series as columns plus the
rest as a line, so it needs at least two series. An unsupported `kind` returns
`unsupported_feature` listing the full supported set.

1.3 adds **chart-polish** props accepted both on `add` and on `set
/slide[i]/chart[k]`: `dataLabels`, `legend`, `axisTitles`, `trendline`,
`errorBars`, `gridlines`, `secondaryAxis`. `get` reports them. Full grammar:
`aioffice help chart-polish`.

## 3D model (1.3, `/slide[1]` add type "model3d")

`{op:"add", path:"/slide[1]", type:"model3d", props:{src:"chair.glb", poster?,
x?, y?, w?, h?, name?}}` embeds a `.glb`/`.gltf` 3D model as a real 3DModel media
part behind a poster picture fallback (PowerPoint 2019+ renders the model). `src`
and `poster` are sandbox-resolved; the add carries a `model3d_as_media` warning.
Addressed by `/slide[i]/model3d[@id=N]`. Full topic: `aioffice help 3d-models`.

## media (1.1, `/slide[1]` add type "media")

`{op:"add", path:"/slide[1]", type:"media", props:{src:"clip.mp4", poster?,
x?, y?, w?, h?, name?, autoplay?}}` embeds an audio/video file (mp4/mov/m4a/
mp3/wav) as a `p:pic` with an `a:videoFile`/`a:audioFile`. `src` (and an
optional `poster` image) is sandbox-resolved — a path outside the workspace is
`sandbox_denied` before any byte is read. The op returns the canonical embed
path; `read {view:"embeds"}` lists the media part.

## text & shape effects (1.1, shape-level set)

`{op:"set", path:"/slide[1]/shape[@id=5]", props:{shadow?, glow?,
reflection?, outline?}}` writes into the shape's `a:effectLst`. Each effect
takes `true` (default accent) or a hex/named color (`shadow`/`glow`/`outline`
tint; `reflection` is colorless); `false` clears it. `outline` adds a line
border.

## transition (M3, 1.1 expanded, slide-level set)

`{op:"set", path:"/slide[2]", props:{transition:"none|fade|push|wipe|split|
reveal|cut|zoom", transitionDuration:"0.5s"}}` — `get /slide[2]` and the
outline view read it back. 1.1 adds `split`, `reveal`, `cut`, `zoom`.

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

Cell alignment & spacing (**1.7**), set on a cell path:

| prop          | type   | notes                                              |
|---------------|--------|----------------------------------------------------|
| valign        | string | `top` · `middle` · `bottom` (the `a:tcPr` anchor)  |
| marginLeft    | length | internal cell padding, e.g. `"0.2cm"` / `"6pt"`    |
| marginRight   | length | "                                                  |
| marginTop     | length | "                                                  |
| marginBottom  | length | "                                                  |
| textDirection | string | `horizontal` · `vertical` · `vertical270` · `stacked` |

## animation (M4 entrance, M5 emphasis + exit; `add` on the shape path)

`{op:"add", path:"/slide[1]/shape[@id=4]", type:"animation",
props:{effect, trigger?, duration?, delay?, direction?, color?}}`

| class    | effects                          | extras                            |
|----------|----------------------------------|-----------------------------------|
| entrance | appear · fade · flyIn · wipe     | direction: flyIn/wipe             |
| emphasis | pulse · grow · spin · colorPulse | color: colorPulse only            |
| exit     | fadeOut · flyOut · wipeOut       | direction: flyOut/wipeOut         |
| motion (1.3) | motionPath                   | path: line/arc/circle/custom; custom takes points |

The 1.3 `motionPath` effect moves the shape along a `path` (`line`/`arc`/
`circle`/`custom`); `custom` traces a `points` array of normalised `[x,y]`
pairs. Full grammar: `aioffice help animations`.

Triggers: `click` (default) · `withPrevious` · `afterPrevious`; durations like
`"0.5s"`. `read --view structure` lists per-slide animation order. M6 timeline
editing: `{op:"move", path:"/slide[i]/animation[2]", position:"before
/slide[i]/animation[1]"}` reorders the timeline; `set /slide[i]/animation[k]`
retunes its props; remove `/slide[i]/animation[k]` to drop one.

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

## slide size (M6, `/` root)

`{op:"set", path:"/", props:{slideSize:"16:9|4:3|16:10|A4|letter"}}` or explicit
`{width, height}` (cm/in/emu) rewrites `p:sldSz`; existing shapes keep their
coordinates. `get /` reports `slideSize`, `widthCm`, `heightCm`, `slideCount`,
`sectionCount`.

## sections (M6, `/section[i]`)

Standard PowerPoint sections group slides in the outline. Add on the `/` root:

    {op:"add", path:"/", type:"section", props:{name:"Intro", afterSlide:0}}
    {op:"add", path:"/", type:"section", props:{name:"Body",  afterSlide:1}}

`afterSlide` is 0-based (0 = before slide 1, N = after slide N); a new section
claims the still-unsectioned slides from there up to the next section. `set
/section[i] {name}` renames; `remove /section[i]` drops the section (its slides
survive, just unsectioned). `read --view outline` groups slides under their
sections; `get /section[i]` reports its name and slide list.

## master / layout editing (M6)

The M1 read-only debt is paid: slide masters and layouts are now editable.

| target                    | ops                                                    |
|---------------------------|--------------------------------------------------------|
| `/master[m]`              | `set` background, accent1..accent6 (theme color scheme) |
| `/master[m]/layout[l]`    | `set` background                                       |
| `/master[m]` (add layout) | `{op:"add", type:"layout", props:{name, basedOn?}}` clones an existing layout |
| `/master[m]/shape[i]`, `/master[m]/layout[l]/shape[i]` | the slide shape ops (set/add/remove) |

    aioffice edit deck.pptx --ops '[{"op":"set","path":"/master[1]","props":{"background":"0F172A","accent1":"38BDF8"}}]'
    aioffice edit deck.pptx --ops '[{"op":"add","path":"/master[1]","type":"layout","props":{"name":"Section","basedOn":1}}]'

Use a cloned layout on a new slide via `props:{layout:N}` (1-based) on `add
type:slide`. `remove` a layout only when no slide references it. `get
/master[m]` lists layouts (name, type, `usedBySlides`).

## document properties (M7, `/properties`)

Core + custom package metadata on a virtual `/properties` node (same contract as
docx/xlsx). Read with `get /properties` or `read --view properties`
(`data.properties.{core:{…}, custom:{…}}`; M10: this nested shape is now
identical across all three formats — read the title at
`data.properties.core.title`); write with `set`:

    {op:"set", path:"/properties", props:{
      title:"Q3 Deck", author:"Dana", subject:"Quarterly review",
      keywords:"q3;sales", category:"Reports", comments:"draft",
      created:"2026-06-01", modified:"2026-06-13",
      custom:{ Project:"Acme", Reviewed:true, Budget:1000, Due:"2026-07-01" }}}

Core props go to the package CoreFilePropertiesPart; `custom` becomes typed
`vt:*` entries (bool/number/date/string) in the CustomFilePropertiesPart. Clear
a core prop with `""`, a custom prop with `null`. Settable core props: title,
subject, author, keywords, category, comments, lastModifiedBy, created,
modified, revision.

## alt text + reading order (M7)

`set altText`/`altTitle` on any shape path writes `a:cNvPr/@descr` and `@title`
(the screen-reader description); `get` reports them; `render --to svg` emits an
SVG `<title>` from the alt text so assistive tech announces it.

    {op:"set", path:"/slide[2]/shape[@id=5]", props:{altText:"Sales chart Q3"}}

Reading order is **document order** of the slide's shapes (the same lever as
z-order: narration order = tab order = paint order). `read --view structure`
lists shapes per slide in that order; reorder with

    {op:"move", path:"/slide[2]/shape[@id=5]", position:"readingOrder 1"}

(1 = narrated first). `front`/`back`/`forward`/`backward` move relatively.

## audit (M7, `aioffice audit`)

`aioffice audit deck.pptx [--category accessibility|quality|all]
[--severity error|warning|info] [--fix] [--json]` lints the deck. Findings
carry a stable `id` (`code#path`) so `--fix` can target specific ones; default
`--fix` repairs every autofixable finding and never destroys content.

| code (accessibility)   | severity      | autofix                                  |
|------------------------|---------------|------------------------------------------|
| a11y_no_alt_text       | warning       | yes — placeholder `(describe this image)`|
| a11y_no_slide_title    | warning       | yes — sets/adds a `(slide title)`        |
| a11y_reading_order     | info          | no  — reorder shapes (readingOrder N)    |
| a11y_tiny_font         | warn <12pt · error <8pt | no — raise the font size       |
| a11y_low_contrast      | warning       | no  — fix the text/fill colors (WCAG 4.5:1)|

| code (quality)         | severity      | autofix                                  |
|------------------------|---------------|------------------------------------------|
| quality_off_canvas     | warning       | no  — move the shape onto the slide      |
| quality_empty_placeholder | warning    | yes — removes the empty placeholder      |
| quality_duplicate_id   | error         | no  — give each shape a unique id        |

## shape hyperlinks & actions (M8)

A shape's click action lives on its `set` props. Set `hyperlink` to one of:

| target form              | effect                                                       |
|--------------------------|--------------------------------------------------------------|
| `https://…`              | open an external URL                                         |
| `#slide:4`               | jump to slide 4                                              |
| `#first` · `#last`       | jump to the first / last slide                              |
| `#next` · `#prev`        | go to the next / previous slide                             |
| `#end`                   | end the show                                                |
| `""` (empty)             | clear the shape's click action                              |

    {op:"set", path:"/slide[1]/shape[@id=5]", props:{hyperlink:"https://example.com"}}
    {op:"set", path:"/slide[1]/shape[@id=5]", props:{hyperlink:"#slide:4"}}   # jump-to-slide
    {op:"set", path:"/slide[1]/shape[@id=5]", props:{hyperlink:"#next"}}      # show action

`linkText` wraps the shape's text runs in a text hyperlink (same target
grammar). `get /slide[1]/shape[@id=5]` reports the canonical hyperlink form
(`url` / `#slide:N` / `#first…`) back.

## Rendering

`aioffice render deck.pptx --to svg --scope /slide[2] -o slide2.svg` renders
one slide; `--to html` renders a simple HTML projection; `--to png` screenshots
one slide via a local Chrome/Edge (default `/slide[1]` — pass `--scope`);
`--to pdf` (M3) prints the WHOLE deck to one PDF, one page per slide
(`--scope /slide[N]` narrows it to a single page).

## SmartArt, connectors, grouping (v1.2)

Three structural shape operations on a slide (all `add` ops, all validate clean).
See `aioffice help smartart` and `aioffice help connectors` for full prop tables.

- **SmartArt** — a real diagram (regenerates in PowerPoint from layout + data):

      {op:"add", path:"/slide[1]", type:"smartart",
        props:{layout:"process", nodes:[{text:"Plan",level:0},{text:"Build",level:0}]}}

  `layout` is one of `list`, `process`, `hierarchy`, `orgChart`, `cycle`;
  `nodes` is a flat `{text, level}` list (0-based level builds the hierarchy).
  Create + read only — `get /slide[1]/smartart[1]` returns the layout and node
  tree; editing nodes in place is `unsupported_feature` (rebuild instead).

- **Connector** — a line between two shapes (`@id` or name):

      {op:"add", path:"/slide[1]", type:"connector",
        props:{from:"@2", to:"@3", kind:"elbow", endArrow:"arrow"}}

  `kind` straight|elbow|curved; `startArrow`/`endArrow` none|arrow|triangle;
  `color`, `width`, `name`. Returns `/slide[i]/shape[@id=N]`.

- **Group / ungroup** — wrap or dissolve a set of shapes:

      {op:"add", path:"/slide[1]", type:"group", props:{shapes:["@2","@3"]}}
      {op:"add", path:"/slide[1]/group[@id=5]", type:"ungroup"}

  `group` needs two or more distinct shapes and returns
  `/slide[i]/group[@id=N]`; a group child is
  `/slide[i]/group[@id=N]/shape[@id=M]`. `ungroup` (on the group path, no props)
  promotes children back onto the slide with absolute coordinates.
