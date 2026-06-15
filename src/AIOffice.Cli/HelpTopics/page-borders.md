# page-borders + IF fields (docx, 1.4)

## page borders

Set a `pageBorder` on a section to draw a border around the page. The path targets
a section: `/section[i]`.

```jsonc
{op:"set", path:"/section[1]", props:{pageBorder:{style:"single", color:"38BDF8", widthPt:1.5, sides:"all"}}}
{op:"set", path:"/section[1]", props:{pageBorder:"none"}}   // remove the border
```

| prop | values |
|------|--------|
| `style` | `single` · `double` · `thick` · `dashed` · `dotted` · `wave` |
| `color` | hex (e.g. `38BDF8`); default `auto` |
| `widthPt` | positive number of points (default 0.5) |
| `sides` | `all` · `top` · `bottom` · `left` · `right` (default `all`) |

`validate` is clean. The whole `pageBorder` object is read back by `get /section[i]`.

## IF fields

`type:"ifField"` adds a Word «IF» field that resolves during a `template` merge:
the field compares a merge field against a literal and emits one of two texts.

```jsonc
{op:"add", path:"/body/p[2]", type:"ifField",
 props:{field:"Country", operator:"=", value:"US",
        trueText:"Domestic shipping", falseText:"International shipping"}}
```

| prop | meaning |
|------|---------|
| `field` | the merge field to compare (e.g. `Country`) |
| `operator` | `=` `<>` `>` `<` `>=` `<=` (default `=`) |
| `value` | the literal compared against the field's merged value |
| `trueText` | text emitted when the comparison holds |
| `falseText` | text emitted otherwise |

During `template`/mail-merge, each record's value for `field` drives the choice, so
the same letter says "Domestic" for a US record and "International" for an FR one.
`get` reads the parsed field/operator/value/trueText/falseText back.

See also: `aioffice help mail-merge`.
