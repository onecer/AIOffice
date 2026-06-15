# layouts — custom slide layouts with placeholders (pptx, M6/1.5)

A slide layout lives on a master. M6 could **clone** an existing layout; 1.5 adds
building a **fresh** layout from a list of placeholder shapes, then binding a slide
to it by name.

## clone an existing layout (M6)

```jsonc
{op:"add", path:"/master[1]", type:"layout", props:{basedOn:2, name:"Hero"}}
```

`basedOn` is a 1-based layout index to copy. `name` is the new layout's display
name.

## a fresh layout with placeholders (1.5)

```jsonc
{op:"add", path:"/master[1]", type:"layout",
 props:{name:"Hero",
        placeholders:[
          {type:"title", x:"2cm", y:"1cm",  w:"28cm", h:"3cm"},
          {type:"body",  x:"2cm", y:"5cm",  w:"13cm", h:"10cm"}
        ]}}
```

Each placeholder is `{type, x, y, w, h}` with `type` one of `title | body | pic |
chart | table | …`; aioffice assigns the correct placeholder type/idx. A new
`slideLayout` part is created on the master.

`basedOn` and `placeholders` are mutually exclusive — clone **or** build fresh.

## use the layout

```jsonc
{op:"add", path:"/slide[2]", type:"slide", props:{layoutName:"Hero"}}
```

Binds a new slide to the layout by **name** with `layoutName`; bind by **index**
with `layout` (a 1-based layout index into the master).

## get / remove

- `get /master[1]/layout[@name=Hero]` reports the layout's placeholders.
- `remove` drops a layout when no slide references it.

Validator-clean. Additive on the frozen `add`/`get`/`remove` verbs.

See also: `aioffice help action-buttons`, `aioffice help properties-pptx`.
