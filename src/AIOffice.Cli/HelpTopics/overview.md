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

Global flags: `--json` (compact; default when stdout is not a TTY), `--pretty`,
`--workspace <dir>` (sandbox root, default cwd, also AIOFFICE_WORKSPACE),
`--quiet` (suppress success envelopes).

Exit codes: 0 ok · 2 user/input error · 3 internal/format error ·
4 sandbox_denied · 5 unsupported_feature.

Topics: `aioffice help addressing | selectors | properties-docx |
properties-xlsx | properties-pptx | errors`. Machine-readable surface:
`aioffice schema [verb]`.
