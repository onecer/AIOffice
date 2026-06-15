# connectors (v1.2 — link shapes, group and ungroup in pptx)

v1.2 adds three structural shape operations to slides: **connector** (a line
between two shapes), **group** (wrap shapes in a group) and **ungroup**
(dissolve a group). All three are `add` ops on a slide and pass `validate` with
0 OpenXML errors.

## Connector — `add type:connector`

Wires a `p:cxnSp` between two existing shapes (by `@id` or name):

    aioffice edit deck.pptx --ops '[{"op":"add","path":"/slide[1]","type":"connector","props":{"from":"@2","to":"@3","kind":"elbow"}}]'

- `props.from`, `props.to` (required): the two shapes to link, each `"@id"` or a
  shape name; they must be different shapes.
- `props.kind`: `straight` (default), `elbow` (right-angle) or `curved`.
- `props.startArrow`, `props.endArrow`: `none` (default), `arrow` or `triangle`.
- `props.color` (hex/named, default black), `props.width` (e.g. `2pt`),
  `props.name`.
- Returns the connector's path `/slide[i]/shape[@id=N]` (it lives in the shape
  tree as a connection shape). `get` on it reports the endpoint shape ids.

## Group — `add type:group`

Wraps two or more shapes in a `p:grpSp`:

    aioffice edit deck.pptx --ops '[{"op":"add","path":"/slide[1]","type":"group","props":{"shapes":["@2","@3"]}}]'

- `props.shapes` (required): at least two distinct shape references (`"@id"` or
  name). Duplicates are ignored; paint order is preserved.
- Returns the group's path `/slide[i]/group[@id=N]`.
- `get /slide[i]/group[@id=N]` lists the group's children. A group child is
  addressed `/slide[i]/group[@id=N]/shape[@id=M]`; `set` on a child takes normal
  shape props, while `set` on the group itself takes `name`, `altText`,
  `altTitle` (to resize, ungroup first, then move/resize the shapes).

## Ungroup — `add type:ungroup`

Dissolves a group, promoting its children back onto the slide with absolute
coordinates:

    aioffice edit deck.pptx --ops '[{"op":"add","path":"/slide[1]/group[@id=5]","type":"ungroup"}]'

- The path is the group path; `ungroup` takes no props.
- The children land on the slide at their effective positions (the group's
  transform is baked into each child).

## Notes

- `group`/`ungroup` are `add` **type** values, not new `op` kinds — the op set
  stays set/add/remove/move/replace/accept/reject/extract.
- docx/xlsx are N/A (these are slide DrawingML operations). xlsx has its own
  `add type:group` for **row/column outline grouping** — a different feature.
