# zoom — slide/section/summary navigation (pptx, 1.4)

A **zoom** is a navigation object that jumps to a slide, a section, or summarizes
several. Add it with `type:"zoom"` on a slide path; it lands at
`/slide[i]/zoom[k]`.

```jsonc
// Slide zoom: a thumbnail on slide 1 that jumps to slide 3.
{op:"add", path:"/slide[1]", type:"zoom", props:{kind:"slide", target:"slide3", x:"2cm", y:"2cm"}}

// Section zoom: jumps to a named section.
{op:"add", path:"/slide[1]", type:"zoom", props:{kind:"section", target:"Results"}}

// Summary zoom: one frame per section.
{op:"add", path:"/slide[1]", type:"zoom", props:{kind:"summary"}}
```

| prop | meaning |
|------|---------|
| `kind` | `slide` · `section` · `summary` |
| `target` | slide zoom: `slide3` / a slide index; section zoom: the section name |
| `x` `y` `w` `h` | position/size of the zoom frame (lengths, e.g. `2cm`); sensible defaults |
| `name` | optional graphic-frame name |

- The target slide(s) are referenced by a real slide relationship, so PowerPoint
  navigates correctly.
- `get /slide[i]/zoom[k]` reports the kind and target; `read --view structure`
  lists the zooms on each slide.
- A zoom's kind/target are immutable in place — `remove` it and add a new one to
  retarget. `validate` is clean for every kind.
