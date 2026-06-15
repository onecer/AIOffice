# chart polish (v1.3)

Additive presentation props for native charts in **both** xlsx and pptx. They are
accepted two ways:

- when you **add** a chart (`type:"chart"`, alongside `kind`/`dataRange`/…), and
- when you **set** them on an existing chart by its path
  (`/Sheet1/chart[1]` in xlsx, `/slide[i]/chart[k]` in pptx).

`get` on the chart reports the polish settings back. Every combination stays
OpenXmlValidator-clean and opens in real Excel / PowerPoint. An unsupported
sub-value returns `unsupported_feature` listing the supported set.

## the props

| prop | value | effect |
|------|-------|--------|
| `dataLabels` | `true` or `{show, position?}` | show data labels. `show`: `value` (default) · `percent` · `category` · `seriesName`. `position`: `outEnd` · `inEnd` · `center` · `bestFit` |
| `legend` | `none` · `right` · `left` · `top` · `bottom` | place or hide the legend |
| `axisTitles` | `{category?, value?}` | axis titles, e.g. `{category:"Month", value:"Sales"}`; either key optional |
| `trendline` | `none` · `linear` · `exponential` · `movingAverage` | fit a trendline (per series or chart-wide). `movingAverage` period defaults to 2 |
| `errorBars` | `none` · `stdErr` · `stdDev` · `percent` | add Y error bars |
| `gridlines` | `{major?, minor?}` | toggle value-axis major / minor gridlines |
| `secondaryAxis` | `["Series A", …]` | move the named series onto a secondary value axis |

## examples

```jsonc
// xlsx — add a bar chart already polished
{op:"add", path:"/Sheet1", type:"chart", props:{
  kind:"bar", dataRange:"A1:C6", anchor:"E2",
  dataLabels:{show:"value", position:"outEnd"},
  legend:"bottom",
  axisTitles:{category:"Month", value:"Revenue"},
  trendline:"linear",
  gridlines:{major:true, minor:false}}}

// xlsx — push one series to a secondary axis on an existing chart
{op:"set", path:"/Sheet1/chart[1]", props:{secondaryAxis:["Target"]}}

// pptx — same props on a slide chart
{op:"set", path:"/slide[2]/chart[1]", props:{dataLabels:true, legend:"right"}}
```

`trendline`, `errorBars` and `secondaryAxis` operate on series; pie/doughnut
charts (no value axis) reject axis-only props with `unsupported_feature`.
