# smartart (v1.2 — author a SmartArt diagram in pptx)

`add type:smartart` builds a real SmartArt graphic on a slide: the four diagram
parts PowerPoint needs (a `dgm:dataModel` node tree, a reference to a built-in
layout, plus the quick-style and colors parts). PowerPoint and LibreOffice
regenerate the diagram visual from the layout + data when the file opens.
Reading `/slide[i]/smartart[k]` back returns the same layout name and node texts.

## Add a SmartArt

    aioffice edit deck.pptx --ops '[{"op":"add","path":"/slide[1]","type":"smartart","props":{"layout":"process","nodes":[{"text":"Plan","level":0},{"text":"Build","level":0},{"text":"Ship","level":0},{"text":"Review","level":0}]}}]'

- `props.layout` (required): one of `list`, `process`, `hierarchy`, `orgChart`,
  `cycle`. Each maps to the standard built-in PowerPoint layout of that shape, so
  the diagram regenerates natively on open.
- `props.nodes` (required): a flat list of `{text, level}`. `level` is 0-based;
  a node attaches under the most recent node one level shallower (level 0 hangs
  off the diagram root). This flat-with-level form expresses hierarchies and
  org charts without nesting.
- Optional placement/box props: `x`, `y`, `w`, `h` (EMU or unit-qualified) and
  `name`, `colorStyle`.

## Read it back

- `get /slide[1]/smartart[1]` → `{ layout, nodeCount, texts:[{text, level, children?}, …], … }`
  (the node tree; nested children appear when a node has descendants).
- `read --view structure` lists each SmartArt as a `smartart` row with its path,
  layout name, node count and indented texts.

## Notes

- SmartArt is **create + read** in v1.2. Editing an existing diagram's nodes
  in place is `unsupported_feature` — rebuild it with a fresh `add` (remove the
  old one first) to change the node set.
- docx/xlsx are N/A by design (SmartArt is a DrawingML diagram inside a slide).
- The file passes `validate` with 0 OpenXML errors after the add.
