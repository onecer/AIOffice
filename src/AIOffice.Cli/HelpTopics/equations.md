# equations (M6/M10 — LaTeX in, Office Math out; docx + pptx)

aioffice ships a hand-rolled LaTeX → OOXML Math (OMML) converter (no LaTeX
dependency). You write LaTeX; the document gets real `m:oMath` that Word,
PowerPoint and LibreOffice render as an equation. The original LaTeX is stored
on the equation (as an `mc:Ignorable` vendor attribute) so `get` returns your
source verbatim and round-trips are byte-faithful. M10 moved the converter into
Core (`AIOffice.Core.Equations`, a pure System.Xml.Linq OMML producer), so the
**same engine drives docx and pptx** — a given LaTeX string renders identically
in both. Excel is N/A by design (cell formulas, not math objects): `add
type:equation` on an `.xlsx` returns `unsupported_feature` with the workaround.

## Add an equation (docx)

    aioffice edit r.docx --ops '[{"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"E = mc^2"}}]'
    aioffice edit r.docx --ops '[{"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"\\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}","display":true}}]'

- `display:false` (default) appends an **inline** equation to the target
  paragraph; the result path is `/body/p[i]/omath[j]`.
- `display:true` emits a **centered display block** (`m:oMathPara`) placed
  `before`/`after`/`inside` like any other block add.

## Add an equation (pptx, M10)

    aioffice edit deck.pptx --ops '[{"op":"add","path":"/slide[1]","type":"equation","props":{"latex":"x = \\frac{1}{2}"}}]'

- Target `/slide[i]` to create a text box holding the equation, or
  `/slide[i]/shape[@id=N]` to append the math to an existing text box.
- Optional props `x`/`y`/`w`/`h` (place a new box) and `fontSize` (a bare point
  number, e.g. `32`) for the fallback run.
- The result path is `/slide[i]/shape[@id=N]/omath[k]`. The OMML lands natively
  inside the text box (`mc:AlternateContent`/`a14:m`/`m:oMathPara`).

## Numbered display equations (1.7, docx)

A display equation can carry an equation number on the right margin:

    aioffice edit r.docx --ops '[{"op":"add","path":"/body","type":"equation","props":{"latex":"E = mc^2","display":true,"number":true}}]'

- `number:true` auto-increments the document's equation counter (1, 2, 3, …).
- `number:"(1.1)"` (or any string) sets the label verbatim.
- A number on an **inline** equation is `invalid_args` — pass `display:true`.
- Address a numbered equation by its number:
  `get r.docx "/equation[@num=1.1]"` (the numeric core matches a `"(1.1)"`
  label; a whole-number label answers to `/equation[@num=1]`). `get` on the omath
  path also reports the `number`.
- pptx/xlsx do not number equations (display equations there are plain).

## Read it back

- `get <omath path>` → `{ latex, … }` (your stored source).
- docx `read --view text` shows `$…$` (inline) and `$$…$$` (display) markers
  inline with the surrounding text.

## Supported LaTeX subset

| feature              | examples                                                  |
|----------------------|-----------------------------------------------------------|
| super/subscripts     | `x^2`, `a_i`, `x_i^2`, `\sum_{i=1}^n`                      |
| fractions            | `\frac{a}{b}`, `\dfrac`, `\tfrac`, `\cfrac`               |
| radicals             | `\sqrt{x}`, `\sqrt 3` (unbraced single token)             |
| big operators        | `\sum`, `\prod`, `\int`, `\oint`, `\lim` with `_`/`^`     |
| matrices/environments| `\begin{pmatrix}1&0\\0&1\end{pmatrix}` (also bmatrix, Bmatrix, vmatrix, Vmatrix, plain matrix; `&` = column, `\\` = row) |
| equation arrays (1.7)| `\begin{aligned} … \end{aligned}`, `\begin{gathered}`, `\begin{cases} x & x\geq 0 \\ -x & x<0 \end{cases}` (emit `m:eqArr`; cases wraps in a `{` brace) |
| binomials (1.7)      | `\binom{n}{k}`, `\dbinom`, `\tbinom` (no-bar fraction in parentheses) |
| horizontal braces (1.7)| `\overbrace{a+b}^{n}`, `\underbrace{x+y}_{m}` (`m:groupChr`) |
| multi-integrals (1.7)| `\iint`, `\iiint` (double/triple integral) alongside `\int \oint` |
| delimiters           | `\left( … \right)`, `\left[ … \right]`                    |
| accents              | `\bar{x}`, `\overline{x}`, `\hat{x}`, `\vec{x}`, `\tilde{x}` |
| text runs            | `\text{…}`, `\mathrm`, `\mathbf`, `\mathit`, `\operatorname` |
| Greek letters        | `\alpha \beta \gamma … \omega` (lower + upper)            |
| operators/relations  | `\pm \times \cdot \leq \geq \neq \approx \infty \partial \nabla \rightarrow` plus (1.7) `\prec \succ \preceq \succeq \lesssim \gtrsim \leqslant \geqslant \subsetneq \supsetneq \longrightarrow \Longrightarrow \hookrightarrow \rightleftharpoons \nrightarrow \triangleq \doteq` and many more arrows/relations |
| spacing              | `\,` `\;` `\quad` `\qquad` (consumed; no visible glyph)   |

## Partial equations (honesty)

An unrecognized command (e.g. `\foobar`) does **not** fail the edit: it degrades
to a literal run of its token text, the rest of the equation still renders, and
the envelope carries an `equation_partial` warning naming the unknown tokens.
The file still passes `validate` with 0 errors. To force literal text, wrap the
fragment in `\text{…}`.

## Notes

- Equations work in **docx** (M6) and **pptx** (M10) through the same shared
  converter. xlsx is N/A by design (`unsupported_feature` — put math in a cell
  formula or embed a rendered image).
- Tracked-changes (`--track`) does not record equation adds — run the op
  without `--track`.
- Remove with `{op:"remove", path:"<omath path>"}` — docx `/body/p[i]/omath[j]`
  (a display equation's whole paragraph goes with it when it becomes empty) or
  pptx `/slide[i]/shape[@id=N]/omath[k]`.
