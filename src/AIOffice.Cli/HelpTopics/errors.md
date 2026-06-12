# Error codes

Every failure envelope has `error.code`, a human `message`, an actionable
`suggestion`, and sometimes `candidates` (ranked alternatives). Exit codes in
parentheses.

| code                  | exit | meaning / what to do                                      |
|-----------------------|------|-----------------------------------------------------------|
| invalid_args          | 2    | Bad flags or arguments; the suggestion shows the shape.    |
| file_not_found        | 2    | Check the path or `aioffice create` it.                    |
| sandbox_denied        | 4    | Path escapes `--workspace`; widen the sandbox or move the file. |
| invalid_path          | 2    | Address didn't resolve; `candidates` lists nearest existing paths. |
| stale_address         | 2    | The file changed since you read it (`--expect-rev` mismatch); re-run read/query and retry. |
| unsupported_feature   | 5    | Capability not built yet; the suggestion names the workaround. |
| file_too_large        | 2    | File exceeds the opt-in `AIOFFICE_MAX_FILE_MB` cap (default: unlimited); raise/unset the cap or split the file. |
| format_corrupt        | 3    | The file is not valid OOXML; try `aioffice validate` for details. |
| internal_error        | 3    | A bug in aioffice; re-run and report it.                   |

Warning-level (in `meta.warnings`, command still succeeds):

| code                   | meaning                                            |
|------------------------|----------------------------------------------------|
| formula_not_evaluated  | A formula could not be evaluated; cached value used. |

Recovery tools: `aioffice snapshot list <file>` shows the automatic pre-edit
snapshot ring (last 20); `aioffice snapshot restore <file> [n]` rolls back —
and the restore itself is snapshotted, so it is undoable too.
