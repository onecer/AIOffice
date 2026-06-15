# conditional formatting (xlsx)

`{op:"add", path:"/Sheet1/A1:C10", type:"conditionalFormat", props:{kind:…, …}}`
— the op path is the range the rule covers. `get`/`remove` address the rule by
`/Sheet1/conditionalFormat[i]` (later indices shift down after a remove).

## kinds

| kind | since | props |
|------|-------|-------|
| `cellIs` | M2 | `operator` (`>` `>=` `<` `<=` `==` `!=` `between`), `value`, `value2` (between), `fill?`, `color?`, `bold?` |
| `colorScale` | M2 | `minColor`, `maxColor`, `midColor?` |
| `dataBar` | M2 | `color` |
| `containsText` | M2 | `text`, `fill?`, `color?`, `bold?` |
| `iconSet` | 1.1 | `set` (e.g. `3TrafficLights1`, `3Arrows`, `4Rating`, `5Quarters`), `reverse?`, `showValue?` |
| `formula` | **1.3** | `formula` (an `=expression` using the top-left cell, e.g. `=$B1>100`), `fill?`, `color?`, `bold?` |
| `topBottom` | **1.3** | `mode` (`top`/`bottom`), `rank` (the N), `percent?` (N is a %), `fill?`, `color?`, `bold?` |
| `aboveBelowAverage` | **1.3** | `mode` (`above`/`below`/`aboveOrEqual`/`belowOrEqual`), `stdDev?` (whole number of standard deviations), `fill?`, `color?`, `bold?` |

## the 1.3 kinds

- **`formula`** — an arbitrary boolean Excel formula evaluated per cell; write it
  relative to the range's top-left cell. `=$B1>100` highlights any row whose
  column B exceeds 100. Maps to an `expression` cfRule.
- **`topBottom`** — the highest- (`top`) or lowest- (`bottom`) ranked N values in
  the range. `{mode:"top", rank:3}` paints the top 3; `{mode:"bottom", rank:10,
  percent:true}` paints the bottom 10%.
- **`aboveBelowAverage`** — cells above/below the range average. `stdDev:1` shifts
  the threshold by one standard deviation; the `*OrEqual` modes include the mean.

```jsonc
{op:"add", path:"/Sheet1/B2:B100", type:"conditionalFormat",
 props:{kind:"formula", formula:"=$B2>100", fill:"FFC7CE"}}

{op:"add", path:"/Sheet1/C2:C100", type:"conditionalFormat",
 props:{kind:"topBottom", mode:"top", rank:3, fill:"C6EFCE", bold:true}}

{op:"add", path:"/Sheet1/D2:D100", type:"conditionalFormat",
 props:{kind:"aboveBelowAverage", mode:"above", stdDev:1, color:"006100"}}
```

An unsupported `kind` or sub-value returns `unsupported_feature` listing the
supported set.
