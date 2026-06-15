# building-blocks — reusable AutoText / Quick Parts (docx, 1.5)

A **building block** is reusable content stored in the document's glossary part
(`w:docPart`): a named block in a gallery holding one or more paragraphs. Store it
once, then insert its content into the body by name.

## store a block

```jsonc
{op:"add", path:"/buildingBlocks", type:"buildingBlock",
 props:{name:"Disclaimer", gallery:"quickParts", content:"All rights reserved."}}
```

- `name` — required, unique. Re-adding the same name replaces the block in place.
- `gallery` — `quickParts` (default) or `autoText`.
- `category` — optional grouping.
- `content` — plain text (`\n` splits further paragraphs), or a JSON array of
  paragraph strings.

## insert a block into the body

```jsonc
{op:"add", path:"/body", type:"buildingBlockRef", props:{name:"Disclaimer"}}
```

Clones the stored block's content paragraphs into the body. An optional `position`
(`before:<path>` / `after:<path>` / a 1-based index) places it; the default is the
end of the body.

## read / get / remove

- `get /buildingBlock[@name=Disclaimer]` → `{name, gallery, category, content}`.
- `remove /buildingBlock[@name=Disclaimer]` drops the stored block.
- `read --view structure` lists every stored block.

Validator-clean. Additive on the frozen `add`/`remove` verbs.

See also: `aioffice help structural-fields`, `aioffice help properties-docx`.
