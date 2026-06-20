# masters (pptx slide / notes / handout masters)

A pptx has three master parts. M6 added slide-master + layout editing
(`/master[i]`, `/master[i]/layout[j]`); v1.7 adds the two remaining ones — the
**notes master** and the **handout master** — addressed as the set-paths
`/notesMaster` and `/handoutMaster` (singletons, no index). Both parts are
created with a minimal valid body on first edit and reported by `get` and by
`read --view structure`.

## Slide master / layout (M6, recap)

    aioffice edit deck.pptx --ops '[{"op":"set","path":"/master[1]","props":{"background":"1F2937"}}]'
    aioffice edit deck.pptx --ops '[{"op":"add","path":"/master[1]","type":"layout"}]'  # clone a layout

`set /master[m]` props: `background` (solid hex), `accent1..accent6` (theme
colors), and — added in **1.8** — `majorFont` / `minorFont` (the theme font
scheme's headings/body faces) plus `gradient` / `image` backgrounds (the same
specs as a slide; see `aioffice help properties-pptx`). `set
/master[m]/layout[l]` takes `background`, `gradient`, `image`.

    # 1.8: brand the deck's theme fonts (one master edit restyles every slide)
    aioffice edit deck.pptx --ops '[{"op":"set","path":"/master[1]","props":{"majorFont":"Montserrat","minorFont":"Inter"}}]'
    # 1.8: a gradient master background
    aioffice edit deck.pptx --ops '[{"op":"set","path":"/master[1]","props":{"gradient":{"type":"linear","angle":90,"stops":[{"color":"0F172A","at":0},{"color":"1E293B","at":100}]}}}]'

`get /master[1]` reports `majorFont` and `minorFont` back.

### Deck-wide footer / slide number / date (1.13)

`set /` (the deck root) or `set /master[m]` also accepts `footer`,
`slideNumber`, `date` and `skipTitle` — these stamp the footer / slide-number /
date placeholders onto the master, its layouts **and every slide**, and wire the
master/layout `p:hf` visibility, so an agent can number a whole deck in one op:

    aioffice edit deck.pptx --ops '[{"op":"set","path":"/","props":{"footer":"Q3 Review","slideNumber":true,"date":true}}]'
    aioffice edit deck.pptx --ops '[{"op":"set","path":"/master[1]","props":{"slideNumber":true}}]'

`skipTitle` (default `true`) hides the footer on title-layout slides. Per-slide
control and the full grammar live in `aioffice help properties-pptx`.

## Notes master (1.7)

The shared look of every slide's notes page.

    aioffice edit deck.pptx --ops '[{"op":"set","path":"/notesMaster","props":{"background":"FFFFFF","bodyFont":"Calibri"}}]'

- `background` — a fill (hex, named, or a gradient/picture spec as elsewhere).
- `bodyFont` — the notes body font (written to the part's theme).

## Handout master (1.7)

The print-handout layout (the page Word/PowerPoint prints N slides onto).

    aioffice edit deck.pptx --ops '[{"op":"set","path":"/handoutMaster","props":{
      "headerFooter":{"header":"Acme Confidential","footer":"Q4 deck","date":true,"pageNumber":true},
      "slidesPerPage":3}}]'

- `background` — a fill, as above.
- `headerFooter` — an object: `header` / `footer` are text strings (create or
  rewrite the matching placeholder); `date` / `pageNumber` are booleans toggling
  those placeholders' visibility. Any subset is allowed.
- `slidesPerPage` — one of `1, 2, 3, 4, 6, 9`.

## Read them back

    aioffice get deck.pptx /notesMaster
    aioffice get deck.pptx /handoutMaster
    aioffice read deck.pptx --view structure   # both masters appear in the tree

Every master edit keeps the package validator-clean (`validate` → 0 errors).
