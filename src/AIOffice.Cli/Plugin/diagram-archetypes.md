# aioffice diagram archetypes → native primitives

When the idea has a *shape* (a process, a cycle, a 2×2, a funnel), build the real diagram — don't fake it with a row of cards. aioffice has the native primitives; they're just not in the terse schema. Run `office_help smartart connectors properties-pptx` first.

**Primitive inventory (what actually exists):**
- **SmartArt** `add type:smartart {layout, nodes:[{text,level}]}` — `layout` ∈ `list process hierarchy orgChart cycle` ONLY. Create+read (no in-place node edit — rebuild to change). Native, regenerates in PowerPoint — but **some LibreOffice versions render it blank**, so if your `--engine soffice` render shows nothing, use the connected-cards form (shapes + connectors) below, which renders everywhere.
- **Connector** `add type:connector {from:"@id", to:"@id", kind:"straight|elbow|curved", startArrow, endArrow:"none|arrow|triangle", color, width}`.
- **Shapes** `add type:shape {shape:"rect|roundRect|ellipse|triangle|diamond|arrow|line", x, y, w, h, fill, flipH, flipV}` — **only those 7 presets** (no trapezoid/chevron/pentagon/star). Z-order via `{op:"move", position:"front|back|forward|backward"}`.
- **Text** `add type:text {text, x, y, w, fontSize, color, bold, align}` for labels on top of geometry.
- **Equation** `add type:equation {latex, display}` (pptx + docx).
- **Charts** `add type:chart {kind, …, dataLabels, legend, axisTitles}` for quantitative data — a real chart beats a hand-drawn one. (`seriesColors` is **xlsx-only**; pptx chart series take their color from the master theme accents — set `accent1..3` on `/master[1]`.)

Canvas = 16:9, **centimeters**, 33.87 w × 19.05 h. Coordinates below are illustrative — adapt to your layout.

## process / pipeline / steps
The fastest path is one native SmartArt:
```json
[{"op":"add","path":"/slide[1]","type":"smartart","props":{"layout":"process","nodes":[{"text":"Plan","level":0},{"text":"Build","level":0},{"text":"Ship","level":0},{"text":"Review","level":0}]}}]
```
Want full control of styling? Lay `roundRect` cards and wire them:
```json
[{"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"roundRect","name":"s1","x":1.5,"y":8,"w":6,"h":3,"fill":"{panel}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"roundRect","name":"s2","x":9,"y":8,"w":6,"h":3,"fill":"{panel}"}},
 {"op":"add","path":"/slide[1]","type":"connector","props":{"from":"s1","to":"s2","kind":"elbow","endArrow":"triangle","color":"{accent1}","width":"2pt"}}]
```
(Then `add type:text` for each card's label.)

## cycle / loop / flywheel (PDCA, growth loop)
```json
[{"op":"add","path":"/slide[1]","type":"smartart","props":{"layout":"cycle","nodes":[{"text":"Acquire","level":0},{"text":"Activate","level":0},{"text":"Retain","level":0},{"text":"Refer","level":0}]}}]
```

## hierarchy / org chart
`level` is 0-based; a node hangs under the nearest shallower node:
```json
[{"op":"add","path":"/slide[1]","type":"smartart","props":{"layout":"orgChart","nodes":[{"text":"CEO","level":0},{"text":"CTO","level":1},{"text":"CFO","level":1},{"text":"VP Eng","level":2}]}}]
```

## framework / hub + satellites
No SmartArt fits — build a center + spokes:
```json
[{"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"ellipse","name":"hub","x":13.9,"y":7.5,"w":6,"h":4,"fill":"{accent1}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"roundRect","name":"n1","x":2,"y":2,"w":6,"h":2.4,"fill":"{panel}"}},
 {"op":"add","path":"/slide[1]","type":"connector","props":{"from":"hub","to":"n1","kind":"straight","color":"{muted}"}}]
```
Repeat satellites around the hub (4–6), each its own `connector` back to `hub`.

## matrix / 2×2 (SWOT, BCG, Eisenhower)
Four quadrants + two axis lines + labels:
```json
[{"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"roundRect","x":4,"y":3,"w":12.7,"h":6,"fill":"{panel}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"roundRect","x":17.2,"y":3,"w":12.7,"h":6,"fill":"{panel}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"roundRect","x":4,"y":9.5,"w":12.7,"h":6,"fill":"{panel}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"roundRect","x":17.2,"y":9.5,"w":12.7,"h":6,"fill":"{panel}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"line","x":16.9,"y":3,"w":0,"h":12.5,"fill":"{muted}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"line","x":4,"y":9.2,"w":25.9,"h":0,"fill":"{muted}"}}]
```
Add axis-title text boxes outside the grid; color one quadrant's fill with an accent to mark the "winning" cell.

## funnel (stage drop-off)
No trapezoid preset — center `rect` bands of decreasing width, graduated fill:
```json
[{"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"rect","x":7,"y":4,"w":20,"h":2.4,"fill":"{accent1}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"rect","x":9.5,"y":6.7,"w":15,"h":2.4,"fill":"{accent2}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"rect","x":12,"y":9.4,"w":10,"h":2.4,"fill":"{accent3}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"rect","x":14,"y":12.1,"w":6,"h":2.4,"fill":"{ink}"}}]
```
Label each band with its stage + count/percent (`add type:text`, white on the colored band).

## pyramid (value / Maslow stack)
A `triangle` with horizontal divider `line`s and side labels (no SmartArt pyramid layout):
```json
[{"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"triangle","x":9,"y":3,"w":15,"h":12,"fill":"{accent1}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"line","x":11.5,"y":7,"w":10,"h":0,"fill":"{bg}","width":"2pt"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"line","x":10.2,"y":11,"w":12.6,"h":0,"fill":"{bg}","width":"2pt"}}]
```
Put tier labels as text boxes beside each band.

## timeline / roadmap
A horizontal axis `line` + milestone `ellipse` markers + dated labels:
```json
[{"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"line","x":2,"y":9.5,"w":29.8,"h":0,"fill":"{muted}","width":"2pt"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"ellipse","x":4,"y":9,"w":1,"h":1,"fill":"{accent1}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"ellipse","x":12,"y":9,"w":1,"h":1,"fill":"{accent1}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"ellipse","x":20,"y":9,"w":1,"h":1,"fill":"{accent2}"}}]
```
Alternate labels above/below the axis; a `breathing`-rhythm timeline reads better than a dense one.

## comparison / before-after
Symmetric split — two panels + a divider, each side its own visual anchor:
```json
[{"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"roundRect","x":1.5,"y":4,"w":14.5,"h":12,"fill":"{panel}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"roundRect","x":17.9,"y":4,"w":14.5,"h":12,"fill":"{ink}"}},
 {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"line","x":16.9,"y":3,"w":0,"h":13,"fill":"{muted}"}}]
```
Use color contrast (light vs dark panel) to carry "old vs new"; keep the two columns structurally identical.

## equation / formula
```json
[{"op":"add","path":"/slide[1]","type":"equation","props":{"latex":"\\frac{\\partial L}{\\partial w} = \\sum_i (y_i-\\hat y_i)\\,x_i","display":true}}]
```

## quantitative data → a real chart, not a drawn one
If the "diagram" is actually numbers (trend, share, comparison of values), use `add type:chart` with `dataLabels`+`legend`+`axisTitles` (and `seriesColors` on **xlsx**; pptx series follow the master theme accents) — see `office_help chart-polish`. A native chart is editable, accurate, and re-renders on open; hand-drawn bars are none of those.
