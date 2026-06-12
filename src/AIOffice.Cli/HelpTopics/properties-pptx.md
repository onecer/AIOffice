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

## Rendering

`aioffice render deck.pptx --to svg --scope /slide[2] -o slide2.svg` renders
one slide; `--to html` renders a simple HTML projection; `--to png` screenshots
one slide via a local Chrome/Edge (default `/slide[1]` — pass `--scope`).

Master/layout editing is still `unsupported_feature` (M3).
