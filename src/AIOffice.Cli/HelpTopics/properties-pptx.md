# pptx properties (v0)

## slide (`/slide[2]`)

| prop   | type   | notes                                            |
|--------|--------|--------------------------------------------------|
| title  | string | `add`: creates the title shape with this text    |

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

## Rendering

`aioffice render deck.pptx --to svg --scope /slide[2] -o slide2.svg` renders
one slide; `--to html` renders a simple HTML projection. PNG export is M1 and
answers `unsupported_feature` until then (workaround: render svg).

Master/layout editing and speaker notes are reserved for M1.
