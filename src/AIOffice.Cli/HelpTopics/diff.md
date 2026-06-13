# diff

`aioffice diff <file>` **semantically** compares a document against a baseline
and reports a sorted, deterministic change list. Changes are *data*, never
errors: a successful diff always exits `0`, even when the two documents differ a
lot — the same way `aioffice validate` and `aioffice audit` return `ok:true`
while reporting issues.

    aioffice diff new.docx old.docx
    aioffice diff report.docx --snapshot 1
    aioffice diff metrics.xlsx baseline.xlsx --view summary

## the two baseline modes

Exactly one baseline source is required:

| mode            | form                              | baseline is…                                            |
|-----------------|-----------------------------------|---------------------------------------------------------|
| two files       | `aioffice diff <file> <other>`    | `<other>`, the OLD document (sandbox-resolved, same format) |
| snapshot        | `aioffice diff <file> --snapshot N` | snapshot `N` of `<file>` from its automatic pre-edit ring |

`<file>` is always the **current/new** document; the baseline is the **old**
one. A baseline of a different format (e.g. diffing a `.docx` against a `.xlsx`)
fails with `invalid_args` naming the mismatch. Passing both a file and
`--snapshot` — or neither — is also `invalid_args`.

For the snapshot mode, snapshot `N` is restored into a throwaway temp file used
only as the baseline; your file and its snapshot ring are never touched. A
snapshot index that does not exist is `invalid_args` listing the available
numbers as candidates (`aioffice snapshot list <file>` shows the ring).

## options

| option       | values              | default    | meaning                                            |
|--------------|---------------------|------------|----------------------------------------------------|
| `--snapshot` | `N`                 | —          | baseline = snapshot N of the file's own ring       |
| `--view`     | `summary`·`detailed`| `detailed` | `summary` trims each change to `{kind, path}`       |

## result shape

    {
      "changes": [
        { "kind": "modified", "path": "/body/p[2]",
          "before": "First body line.", "after": "Edited body line.", "detail": "text" },
        { "kind": "added", "path": "/body/p[3]", "detail": "paragraph" }
      ],
      "summary": { "added": 1, "removed": 0, "modified": 1, "moved": 0 },
      "baseline": "old.docx",
      "view": "detailed"
    }

- `kind` is `added` · `removed` · `modified` · `moved`.
- `path` is the canonical path in the **current** file — except for `removed`,
  which reports the **baseline** path (it no longer exists in the current file).
- `modified` carries concise `before`/`after` values (old vs new text,
  `Normal` → `Heading1`, a cell value, …); `detail` names what changed
  (`text`, `style`, `cell`, a property name, …).
- `moved` content matched but changed position; `detail` is
  `moved from <old path>`. Moves are recovered with LCS / content-hash matching
  over paragraphs, slides and rows; when in doubt the differ prefers
  modified/added/removed over guessing a move.

`--view summary` keeps the same `summary` counts but trims every entry in
`changes` to just `{kind, path}` — one terse line per change.

## deterministic order

`changes` is always sorted by `(path, kind)` with an ordinal comparer, so the
same two documents diff **identically on every platform and every run**.

## snapshot-diff workflow ("what did I just do?")

Every successful `aioffice edit` snapshots the pre-image first. So after an edit
you can ask exactly what it changed:

    aioffice edit report.docx --set /body/p[2] text='Revised'
    aioffice diff report.docx --snapshot 1

That diff is the last edit's change set — a fast review step before you commit
or share the document.
