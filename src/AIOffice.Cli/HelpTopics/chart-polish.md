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

## seriesColors (1.8 — brand palette, xlsx)

A chart prop accepted on both `add type:chart` and `set /Sheet1/chart[i]`. Pass an
array of 6-hex RGB strings (a leading `#` is optional), one per series in dataRange
order; a short list cycles across the series:

```
aioffice edit book.xlsx --ops '[{"op":"set","path":"/Sheet1/chart[1]","props":{"seriesColors":["2E5AAC","#E8743B"]}}]'
```

Bar/area series get a solid fill in the color; line/scatter series get a tinted line;
a single pie/doughnut series colors each slice (the palette cycles across the data
points). `get` on the chart reports the applied colors under `polish.seriesColors`.
Chart colors live in the chart part (preserved byte-identical), so they survive later
edits. This lets a dashboard's charts match the chosen brand direction instead of
Excel's default office palette.

## edit chart data in place (1.12 — pptx)

`set /slide[i]/chart[k]` now edits an **existing** native pptx chart's data without
remove-and-re-add. Three additive props, any combination in one op (data props apply
first, then any polish props in the same op):

| prop | value | effect |
|------|-------|--------|
| `title` | string · `false` | set/replace the chart title, or `false` to remove it (sets `c:autoTitleDeleted`). `true` is rejected — pass the text |
| `categories` | array of scalar labels | relabel the category axis for every series (rewrites each `c:cat` cache) |
| `series` | array of `{name?, values:[…]}` | replace each series — `values` updates the `c:val` number cache, `name` (optional) the `c:tx`. Omit `name` to keep it |

```jsonc
// pptx — retitle + reshape an existing chart's data
{op:"set", path:"/slide[2]/chart[1]", props:{
  title:"Q4 Actuals",
  categories:["Jan","Feb","Mar"],
  series:[{name:"Plan", values:[10,20,30]},
          {name:"Actual", values:[11,19,34]}]}}

// remove the title only
{op:"set", path:"/slide[2]/chart[1]", props:{title:false}}
```

Both the chart-XML caches **and** the chart's embedded "Edit Data" workbook are
rewritten, so `get`/`render` and PowerPoint's *Edit Data* sheet all show the new
values (a foreign cached-only chart with no embedded workbook still updates its
caches; add one with `{embedData:true}`).

Series are matched **by index**: pass one object per series in order. Passing fewer
series than the chart has leaves the trailing series untouched; passing more updates
the overlapping ones and ignores the surplus (adding/removing whole series restructures
the plot area — remove and re-add the chart for that). Each replacement series' `values`
length must equal the category count; to change the number of categories, pass new
`categories` **and** re-supply every series in the same op. Bubble charts (x/y/size
triples) reject in-place data edits — remove and re-add them.
