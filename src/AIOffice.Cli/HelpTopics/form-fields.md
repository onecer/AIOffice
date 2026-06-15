# legacy form fields (docx) — v1.3

Classic Word form fields (the `fldChar`/`ffData` kind the protected-form
templates use): a named text input, checkbox, or dropdown. Distinct from M-era
*content controls* (`read --view fields` lists both; a form field shows
`kind:"formField"`).

## add

```jsonc
// text input
{op:"add", path:"/body/p[2]", type:"formField",
 props:{kind:"text", name:"clientName", default?:"", maxLength?:40}}

// checkbox
{op:"add", path:"/body/p[3]", type:"formField",
 props:{kind:"checkbox", name:"agree", checked?:true}}

// dropdown
{op:"add", path:"/body/p[4]", type:"formField",
 props:{kind:"dropdown", name:"status", items:["Open","Closed","Pending"]}}
```

- `kind`: `text` (default) · `checkbox` · `dropdown`.
- `name` is **required** — the field identifier you address and set by.
- `dropdown` needs a non-empty `items` array; `text` accepts `default` +
  `maxLength`; `checkbox` accepts `checked`.

## read / set / remove

`read --view fields` lists every form field with its `name`, `kind` and current
value. Update a value by name:

```jsonc
{op:"set", path:"/formField[@name=clientName]", props:{value:"Acme"}}   // text/dropdown
{op:"set", path:"/formField[@name=agree]",      props:{checked:true}}    // checkbox
{op:"remove", path:"/formField[@name=status]"}
```

Form fields are structural, so adding/removing one inside a `--track` session is
`unsupported_feature` (set the value instead). The doc validates clean.
