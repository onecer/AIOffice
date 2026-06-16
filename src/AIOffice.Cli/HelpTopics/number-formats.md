# number-formats (v1.2 — named numberFormat presets, xlsx)

The xlsx `numberFormat` cell prop accepts a literal Excel format **code**
(e.g. `0.00%`, `#,##0`, `yyyy-mm-dd`) — that behavior is unchanged. v1.2 adds a
small library of **named presets**: a preset name resolves to its format code,
so an agent can write `accounting-usd` instead of memorizing the custom-format
string. Anything that is not a known preset is preserved verbatim as a literal
code (Excel shows it as you wrote it), so existing custom formats keep working.

## Use a preset

    aioffice edit data.xlsx --ops '[{"op":"set","path":"/Sheet1/B2","props":{"numberFormat":"accounting-usd"}}]'

`get /Sheet1/B2` then reports the resolved format code and the cached display
string.

## Presets

| preset           | format code                                |
|------------------|--------------------------------------------|
| accounting-usd   | `_("$"* #,##0.00_);_("$"* \(#,##0.00\);…)` |
| accounting-eur   | `_("€"* #,##0.00_);_("€"* \(#,##0.00\);…)` |
| accounting-gbp   | `_("£"* #,##0.00_);_("£"* \(#,##0.00\);…)` |
| currency-usd     | `"$"#,##0.00`                              |
| currency-eur     | `"€"#,##0.00`                              |
| currency-gbp     | `"£"#,##0.00`                              |
| currency-jpy     | `"¥"#,##0`                                 |
| usd-millions     | `"$"#,##0.0,,"M"`  (1,234,567 → $1.2M)     |
| usd-thousands    | `"$"#,##0,"K"`  (1,234,567 → $1,235K)      |
| millions         | `#,##0.0,,"M"`  (1,234,567 → 1.2M)         |
| thousands-k      | `#,##0,"K"`  (1,234,567 → 1,235K)          |
| percent          | `0%`                                       |
| percent-0        | `0%`                                       |
| percent-1        | `0.0%`                                     |
| percent-2        | `0.00%`                                    |
| percent2         | `0.00%`                                    |
| scientific       | `0.00E+00`                                 |
| fraction         | `# ?/?`                                    |
| thousands        | `#,##0`                                    |
| thousands2       | `#,##0.00`                                 |
| integer          | `0`                                        |
| number2          | `0.00`                                     |
| date-iso         | `yyyy-mm-dd`                               |
| datetime-iso     | `yyyy-mm-dd hh:mm:ss`                       |
| time             | `hh:mm:ss`                                 |
| duration         | `[h]:mm:ss`                                |
| text             | `@` (display content as entered)           |

The scaled presets (`usd-millions`, `usd-thousands`, `millions`, `thousands-k`)
use OOXML trailing commas to divide by 1,000 each, then append a literal unit
suffix — so a raw cell value of `1234567` *displays* as `$1.2M` while the stored
number is unchanged. `percent-0` / `percent-1` / `percent-2` are explicit
decimal-precision percent aliases (`percent-2` equals the older `percent2`).

## Notes

- Preset names are case-insensitive. A literal code like `#,##0.00` or
  `yyyy-mm-dd` is never treated as a preset — it is stored as-is.
- A misspelled preset (e.g. `acounting-usd`) is stored as a literal format;
  it is not an error. Run `aioffice help number-formats` to see the exact names.
- Presets apply to cells and to named cell styles
  (`add type:cellStyle props:{numberFormat:"currency-eur"}`).
