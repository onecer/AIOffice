# mail-merge — batch the template verb (docx, 1.4)

The `template` verb fills `{{key}}` placeholders, docx `MERGEFIELD` complex fields
and «IF» fields from a JSON merge map. In 1.4 the same verb runs a **mail merge**
when `--data` is a JSON **array** of record objects — additive new behaviour, no
new verb.

## one document per record

Pass `--output PATTERN` to write one merged document per record:

```bash
aioffice template letter.docx \
  --data '[{"Name":"Ada","city":"London"},{"Name":"Grace","city":"NYC"}]' \
  --output "letters/letter-{n}.docx"
```

- `{n}` is the **1-based record index**; any other `{Field}` is that record's value
  (e.g. `--output "letters/{Name}.docx"`).
- Every expanded path is **sandbox-resolved** — an escaping pattern like
  `../escape-{n}.docx` is denied (`sandbox_denied`). A pattern that produces the
  same path twice is `invalid_args`.
- Result: `{records:N, produced:[paths], unresolved:[...]}`.

## one combined document

Omit `--output` for a single combined document — the source body repeated once per
record, separated by next-page section breaks:

```bash
aioffice template letter.docx --data @people.json
```

The combined document is written back to the source file with an auto-snapshot
(undoable).

## fields filled per record

Each record fills, by name: `{{key}}` placeholders, `MERGEFIELD Name` complex
fields, and «IF» fields (`aioffice help page-borders` covers `ifField`). Any field
left without data raises a single `template_unresolved` warning naming the fields,
and the chevrons are left in place.

## single-object fill (unchanged)

A JSON **object** `--data` still fills exactly one document and returns
`{replaced, keys, unresolved, written}` — the original behaviour is untouched.
