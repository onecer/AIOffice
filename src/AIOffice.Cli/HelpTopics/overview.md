# aioffice — AI-native Office CLI

Every command prints exactly ONE JSON envelope to stdout:

    { "ok": true|false, "data": {...}|null,
      "error": { "code", "message", "suggestion", "candidates"? }|null,
      "meta": { "file"?, "rev"?, "elapsedMs", "version", "warnings"? } }

Errors always carry an actionable `suggestion`. `meta.rev` is the first 12 hex
chars of the SHA-256 of the file bytes — pass it back with `edit --expect-rev`
for optimistic concurrency.

The intended agent loop:

    1. aioffice read  <file> --view outline     # orient
    2. aioffice query <file> "<selector>"       # find → canonical paths
    3. aioffice get   <file> <path>             # inspect one node
    4. aioffice edit  <file> --ops '[...]'      # atomic batch (auto-snapshots)
    5. aioffice validate <file>                 # prove the file is still sound

Document-wide find/replace sugar (M4):

    aioffice edit <file> --find X --replace Y [--regex] [--match-case] [--whole-word]

covers the whole document (docx body + headers/footers, every sheet, every
slide incl. notes) and reports aggregate {replacements, locations}. On docx,
add --track to record every hit as a w:del+w:ins revision pair. The same op is
available in --ops batches as {"op":"replace","path":"<scope>|/","props":{...}}.

Markdown/csv bridge (M5):

    aioffice create report.docx --from notes.md     # GFM markdown -> real docx
    aioffice read   report.docx --view markdown     # docx body -> GFM markdown
    aioffice create orders.xlsx --from orders.csv   # typed cells; 007 stays text
    aioffice read   orders.xlsx --view csv          # one sheet -> RFC 4180 csv

Matrix: .md/.markdown -> .docx, .csv/.tsv -> .xlsx; any other pair fails with
the matrix in the suggestion.

Global flags: `--json` (compact; default when stdout is not a TTY), `--pretty`,
`--workspace <dir>` (sandbox root, default cwd, also AIOFFICE_WORKSPACE),
`--quiet` (suppress success envelopes).

Exit codes: 0 ok · 2 user/input error · 3 internal/format error ·
4 sandbox_denied · 5 unsupported_feature.

Topics: `aioffice help addressing | selectors | properties-docx |
properties-xlsx | properties-pptx | errors`. Machine-readable surface:
`aioffice schema [verb]`.
