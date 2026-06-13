# equations (M6 — LaTeX in, Office Math out)

aioffice ships a hand-rolled LaTeX → OOXML Math (OMML) converter (no LaTeX
dependency). You write LaTeX; the document gets real `m:oMath` that Word,
PowerPoint Online and LibreOffice render as an equation. The original LaTeX is
stored on the equation (as an `mc:Ignorable` vendor attribute) so `get` returns
your source verbatim and round-trips are byte-faithful.

## Add an equation

    aioffice edit r.docx --ops '[{"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"E = mc^2"}}]'
    aioffice edit r.docx --ops '[{"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"\\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}","display":true}}]'

- `display:false` (default) appends an **inline** equation to the target
  paragraph; the result path is `/body/p[i]/omath[j]`.
- `display:true` emits a **centered display block** (`m:oMathPara`) placed
  `before`/`after`/`inside` like any other block add.

## Read it back

- `get /body/p[i]/omath[j]` → `{ latex, display }` (your stored source).
- `read --view text` shows `$…$` (inline) and `$$…$$` (display) markers inline
  with the surrounding text.

## Supported LaTeX subset

| feature              | examples                                                  |
|----------------------|-----------------------------------------------------------|
| super/subscripts     | `x^2`, `a_i`, `x_i^2`, `\sum_{i=1}^n`                      |
| fractions            | `\frac{a}{b}`, `\dfrac`, `\tfrac`, `\cfrac`               |
| radicals             | `\sqrt{x}`, `\sqrt 3` (unbraced single token)             |
| big operators        | `\sum`, `\prod`, `\int`, `\oint`, `\lim` with `_`/`^`     |
| matrices/environments| `\begin{pmatrix}1&0\\0&1\end{pmatrix}` (also bmatrix, Bmatrix, vmatrix, Vmatrix, plain matrix; `&` = column, `\\` = row) |
| delimiters           | `\left( … \right)`, `\left[ … \right]`                    |
| accents              | `\bar{x}`, `\overline{x}`                                 |
| text runs            | `\text{…}`, `\mathrm`, `\mathbf`, `\mathit`, `\operatorname` |
| Greek letters        | `\alpha \beta \gamma … \omega` (lower + upper)            |
| operators/relations  | `\pm \times \cdot \leq \geq \neq \approx \infty \partial \nabla \rightarrow \cdot` and more |
| spacing              | `\,` `\;` `\quad` `\qquad` (consumed; no visible glyph)   |

## Partial equations (honesty)

An unrecognized command (e.g. `\foobar`) does **not** fail the edit: it degrades
to a literal run of its token text, the rest of the equation still renders, and
the envelope carries an `equation_partial` warning naming the unknown tokens.
The file still passes `validate` with 0 errors. To force literal text, wrap the
fragment in `\text{…}`.

## Notes

- Equations are a **docx** feature in M6 (pptx/xlsx equations are M7 seeds).
- Tracked-changes (`--track`) does not record equation adds — run the op
  without `--track`.
- Remove with `{op:"remove", path:"/body/p[i]/omath[j]"}` (a display equation's
  whole paragraph goes with it when it becomes empty).
