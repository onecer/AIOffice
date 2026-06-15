using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace AIOffice.Pptx;

/// <summary>
/// The shared "chart polish" props (v1.3.0, additive): data labels, legend, axis
/// titles, trendlines, error bars, gridlines and a secondary value axis. These are
/// accepted both when a chart is created (alongside kind/categories/series) and as
/// an in-place edit on an existing chart (<c>set /slide[i]/chart[k]</c>). Every
/// element written stays OpenXmlValidator-clean and opens in real PowerPoint.
/// </summary>
internal static class PptxChartPolish
{
    /// <summary>The polish prop keys this module owns (everything else is a chart-data/geometry prop).</summary>
    public static readonly IReadOnlyList<string> PropKeys =
        ["dataLabels", "legend", "axisTitles", "trendline", "errorBars", "gridlines", "secondaryAxis"];

    private static readonly IReadOnlyList<string> LegendPositions = ["none", "right", "left", "top", "bottom"];
    private static readonly IReadOnlyList<string> DataLabelShows = ["value", "percent", "category", "seriesName"];
    private static readonly IReadOnlyList<string> DataLabelPositions = ["outEnd", "inEnd", "center", "bestFit"];
    private static readonly IReadOnlyList<string> Trendlines = ["none", "linear", "exponential", "movingAverage"];
    private static readonly IReadOnlyList<string> ErrorBars = ["none", "stdErr", "stdDev", "percent"];

    private const uint SecondaryCategoryAxisId = 100003u;
    private const uint SecondaryValueAxisId = 100004u;

    /// <summary>True when the props object carries any chart-polish key.</summary>
    public static bool Handles(JsonObject props) =>
        props.Any(kv => PropKeys.Contains(kv.Key, StringComparer.Ordinal));

    /// <summary>Splits the polish keys out of a props object, returning them and (via out) the rest unchanged.</summary>
    public static JsonObject Split(JsonObject props, out JsonObject rest)
    {
        var polish = new JsonObject();
        rest = new JsonObject();
        foreach (var (key, value) in props)
        {
            if (PropKeys.Contains(key, StringComparer.Ordinal))
            {
                polish[key] = value?.DeepClone();
            }
            else
            {
                rest[key] = value?.DeepClone();
            }
        }

        return polish;
    }

    // ----- apply ---------------------------------------------------------------

    /// <summary>
    /// Applies every polish prop present in <paramref name="props"/> to the chart's
    /// chartSpace. Validates each value (unsupported sub-values throw
    /// <c>unsupported_feature</c> listing the supported set) before mutating, so a
    /// bad value leaves the chart untouched.
    /// </summary>
    public static void Apply(C.ChartSpace chartSpace, JsonObject props)
    {
        var chart = chartSpace.GetFirstChild<C.Chart>()
            ?? throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                "The chart part has no c:chart element.",
                "Remove the chart and add it again, or restore a snapshot.");
        var plotArea = chart.GetFirstChild<C.PlotArea>()
            ?? throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                "The chart has no plot area.",
                "Remove the chart and add it again, or restore a snapshot.");

        foreach (var (key, value) in props)
        {
            switch (key)
            {
                case "dataLabels":
                    ApplyDataLabels(plotArea, value);
                    break;
                case "legend":
                    ApplyLegend(chart, value);
                    break;
                case "axisTitles":
                    ApplyAxisTitles(plotArea, value);
                    break;
                case "trendline":
                    ApplyTrendline(plotArea, value);
                    break;
                case "errorBars":
                    ApplyErrorBars(plotArea, value);
                    break;
                case "gridlines":
                    ApplyGridlines(plotArea, value);
                    break;
                case "secondaryAxis":
                    ApplySecondaryAxis(plotArea, value);
                    break;
                default:
                    break; // never reached: Split filters non-polish keys out
            }
        }
    }

    // ----- dataLabels ----------------------------------------------------------

    /// <summary>dataLabels: true (show values) or {show, position?} placed on every chart group.</summary>
    private static void ApplyDataLabels(C.PlotArea plotArea, JsonNode? value)
    {
        string show = "value";
        string? position = null;

        if (value is JsonValue scalar && scalar.TryGetValue<bool>(out var flag))
        {
            if (!flag)
            {
                foreach (var group in ChartGroups(plotArea))
                {
                    group.Elements<C.DataLabels>().ToList().ForEach(d => d.Remove());
                }

                return;
            }
        }
        else if (value is JsonObject obj)
        {
            foreach (var (k, _) in obj)
            {
                if (k is not ("show" or "position"))
                {
                    throw Unsupported($"dataLabels.{k}", "show, position");
                }
            }

            if (obj.TryGetPropertyValue("show", out var showNode) && showNode is not null)
            {
                show = Canonical(J.ScalarText(showNode).Trim(), DataLabelShows, "dataLabels.show");
            }

            if (obj.TryGetPropertyValue("position", out var posNode) && posNode is not null)
            {
                position = Canonical(J.ScalarText(posNode).Trim(), DataLabelPositions, "dataLabels.position");
            }
        }
        else
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"dataLabels must be true or an object {{show, position?}}: {value?.ToJsonString() ?? "null"}",
                "Use {\"dataLabels\":true} or {\"dataLabels\":{\"show\":\"percent\",\"position\":\"center\"}}.");
        }

        foreach (var group in ChartGroups(plotArea))
        {
            group.Elements<C.DataLabels>().ToList().ForEach(d => d.Remove());
            InsertGroupChild(group, BuildDataLabels(show, position));
        }
    }

    private static C.DataLabels BuildDataLabels(string show, string? position)
    {
        var labels = new C.DataLabels();
        if (position is not null)
        {
            labels.Append(new C.DataLabelPosition { Val = position switch
            {
                "inEnd" => C.DataLabelPositionValues.InsideEnd,
                "center" => C.DataLabelPositionValues.Center,
                "bestFit" => C.DataLabelPositionValues.BestFit,
                _ => C.DataLabelPositionValues.OutsideEnd,
            } });
        }

        // The five show-flags must appear in schema order (legendKey, value, catName,
        // serName, percent, bubbleSize). Only the requested channel is turned on.
        labels.Append(new C.ShowLegendKey { Val = false });
        labels.Append(new C.ShowValue { Val = show == "value" });
        labels.Append(new C.ShowCategoryName { Val = show == "category" });
        labels.Append(new C.ShowSeriesName { Val = show == "seriesName" });
        labels.Append(new C.ShowPercent { Val = show == "percent" });
        labels.Append(new C.ShowBubbleSize { Val = false });
        return labels;
    }

    // ----- legend --------------------------------------------------------------

    private static void ApplyLegend(C.Chart chart, JsonNode? value)
    {
        var token = Canonical(J.ScalarText(value ?? string.Empty).Trim(), LegendPositions, "legend");
        chart.Elements<C.Legend>().ToList().ForEach(l => l.Remove());

        if (token == "none")
        {
            return;
        }

        var legend = new C.Legend(
            new C.LegendPosition { Val = token switch
            {
                "left" => C.LegendPositionValues.Left,
                "top" => C.LegendPositionValues.Top,
                "bottom" => C.LegendPositionValues.Bottom,
                _ => C.LegendPositionValues.Right,
            } },
            new C.Overlay { Val = false });

        // c:legend follows c:plotArea and precedes c:plotVisOnly in the c:chart child order.
        if (chart.GetFirstChild<C.PlotVisibleOnly>() is { } plotVisible)
        {
            chart.InsertBefore(legend, plotVisible);
        }
        else if (chart.GetFirstChild<C.PlotArea>() is { } plotArea)
        {
            chart.InsertAfter(legend, plotArea);
        }
        else
        {
            chart.Append(legend);
        }
    }

    // ----- axisTitles ----------------------------------------------------------

    private static void ApplyAxisTitles(C.PlotArea plotArea, JsonNode? value)
    {
        if (value is not JsonObject obj)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"axisTitles must be an object {{category?, value?}}: {value?.ToJsonString() ?? "null"}",
                "Use {\"axisTitles\":{\"category\":\"Quarter\",\"value\":\"Revenue\"}}.");
        }

        foreach (var (k, v) in obj)
        {
            if (k is not ("category" or "value"))
            {
                throw Unsupported($"axisTitles.{k}", "category, value");
            }

            var axis = k == "category"
                ? plotArea.GetFirstChild<C.CategoryAxis>() as OpenXmlCompositeElement
                : plotArea.GetFirstChild<C.ValueAxis>();
            if (axis is null)
            {
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"This chart has no {k} axis to title.",
                    "Axis titles apply to bar/line/area/combo/radar charts; pie and doughnut have no axes.");
            }

            SetAxisTitle(axis, v is null ? string.Empty : J.ScalarText(v));
        }
    }

    private static void SetAxisTitle(OpenXmlCompositeElement axis, string text)
    {
        axis.Elements<C.Title>().ToList().ForEach(t => t.Remove());

        if (text.Length == 0)
        {
            return;
        }

        var title = new C.Title(
            new C.ChartText(new C.RichText(
                new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                new DocumentFormat.OpenXml.Drawing.ListStyle(),
                new DocumentFormat.OpenXml.Drawing.Paragraph(
                    new DocumentFormat.OpenXml.Drawing.Run(
                        new DocumentFormat.OpenXml.Drawing.Text(text))))),
            new C.Overlay { Val = false });

        // In CT_CatAx/CT_ValAx, c:title follows axPos/gridlines and precedes numFmt/
        // tick marks/crossAx. Insert it right before the first numFmt/tick/cross element,
        // after any gridlines (placing it just before c:crossAx is always schema-valid).
        InsertAxisChild(axis, title, before: ["numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt"]);
    }

    /// <summary>Inserts an axis child before the first of the named (later-in-schema) elements.</summary>
    private static void InsertAxisChild(OpenXmlCompositeElement axis, OpenXmlElement child, string[] before)
    {
        var anchor = axis.ChildElements.FirstOrDefault(c => before.Contains(c.LocalName, StringComparer.Ordinal));
        if (anchor is not null)
        {
            axis.InsertBefore(child, anchor);
        }
        else
        {
            axis.Append(child);
        }
    }

    // ----- trendline -----------------------------------------------------------

    private static void ApplyTrendline(C.PlotArea plotArea, JsonNode? value)
    {
        var token = Canonical(J.ScalarText(value ?? string.Empty).Trim(), Trendlines, "trendline");

        foreach (var ser in AllSeries(plotArea))
        {
            ser.Elements<C.Trendline>().ToList().ForEach(t => t.Remove());
        }

        if (token == "none")
        {
            return;
        }

        // Trendlines only make sense for category/value series that carry a c:val.
        foreach (var ser in AllSeries(plotArea))
        {
            if (ser.Elements<C.Values>().Any())
            {
                InsertTrendline(ser, BuildTrendline(token));
            }
        }
    }

    private static C.Trendline BuildTrendline(string kind)
    {
        var trendline = new C.Trendline(new C.TrendlineType { Val = kind switch
        {
            "exponential" => C.TrendlineValues.Exponential,
            "movingAverage" => C.TrendlineValues.MovingAverage,
            _ => C.TrendlineValues.Linear,
        } });

        if (kind == "movingAverage")
        {
            trendline.Append(new C.Period { Val = 2 });
        }

        trendline.Append(new C.DisplayRSquaredValue { Val = false });
        trendline.Append(new C.DisplayEquation { Val = false });
        return trendline;
    }

    /// <summary>c:trendline sits before c:errBars and c:cat/c:val in a ser (CT_*Ser child order).</summary>
    private static void InsertTrendline(OpenXmlCompositeElement ser, C.Trendline trendline)
    {
        var anchor = ser.ChildElements.FirstOrDefault(c => c.LocalName is "errBars" or "cat" or "xVal" or "val" or "yVal");
        if (anchor is not null)
        {
            ser.InsertBefore(trendline, anchor);
        }
        else
        {
            ser.Append(trendline);
        }
    }

    // ----- errorBars -----------------------------------------------------------

    private static void ApplyErrorBars(C.PlotArea plotArea, JsonNode? value)
    {
        var token = Canonical(J.ScalarText(value ?? string.Empty).Trim(), ErrorBars, "errorBars");

        foreach (var ser in AllSeries(plotArea))
        {
            ser.Elements<C.ErrorBars>().ToList().ForEach(e => e.Remove());
        }

        if (token == "none")
        {
            return;
        }

        foreach (var ser in AllSeries(plotArea))
        {
            if (ser.Elements<C.Values>().Any())
            {
                InsertErrorBars(ser, BuildErrorBars(token));
            }
        }
    }

    private static C.ErrorBars BuildErrorBars(string kind)
    {
        // stdErr/stdDev/percent map to the OOXML error-bar value type; the value
        // (1 std-dev, 5%) is a sensible PowerPoint-style default.
        var (valueType, val) = kind switch
        {
            "stdDev" => (C.ErrorValues.StandardDeviation, 1.0),
            "percent" => (C.ErrorValues.Percentage, 5.0),
            _ => (C.ErrorValues.StandardError, 1.0),
        };

        return new C.ErrorBars(
            new C.ErrorDirection { Val = C.ErrorBarDirectionValues.Y },
            new C.ErrorBarType { Val = C.ErrorBarValues.Both },
            new C.ErrorBarValueType { Val = valueType },
            new C.NoEndCap { Val = false },
            new C.ErrorBarValue { Val = val });
    }

    /// <summary>c:errBars sits after c:trendline and before c:cat/c:val (CT_*Ser child order).</summary>
    private static void InsertErrorBars(OpenXmlCompositeElement ser, C.ErrorBars errorBars)
    {
        var anchor = ser.ChildElements.FirstOrDefault(c => c.LocalName is "cat" or "xVal" or "val" or "yVal");
        if (anchor is not null)
        {
            ser.InsertBefore(errorBars, anchor);
        }
        else
        {
            ser.Append(errorBars);
        }
    }

    // ----- gridlines -----------------------------------------------------------

    private static void ApplyGridlines(C.PlotArea plotArea, JsonNode? value)
    {
        if (value is not JsonObject obj)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"gridlines must be an object {{major?, minor?}}: {value?.ToJsonString() ?? "null"}",
                "Use {\"gridlines\":{\"major\":true,\"minor\":false}}.");
        }

        var valueAxis = plotArea.GetFirstChild<C.ValueAxis>()
            ?? throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "This chart has no value axis to put gridlines on.",
                "Gridlines apply to bar/line/area/combo charts; pie and doughnut have no axes.");

        foreach (var (k, v) in obj)
        {
            if (k is not ("major" or "minor"))
            {
                throw Unsupported($"gridlines.{k}", "major, minor");
            }

            var on = v is not null && AsBool($"gridlines.{k}", v);
            if (k == "major")
            {
                valueAxis.Elements<C.MajorGridlines>().ToList().ForEach(g => g.Remove());
                if (on)
                {
                    // c:majorGridlines precedes minorGridlines/title/numFmt/.../crossAx.
                    InsertAxisChild(valueAxis, new C.MajorGridlines(),
                        before: ["minorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt"]);
                }
            }
            else
            {
                valueAxis.Elements<C.MinorGridlines>().ToList().ForEach(g => g.Remove());
                if (on)
                {
                    // c:minorGridlines precedes title/numFmt/.../crossAx (but follows majorGridlines).
                    InsertAxisChild(valueAxis, new C.MinorGridlines(),
                        before: ["title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt"]);
                }
            }
        }
    }

    // ----- secondaryAxis -------------------------------------------------------

    /// <summary>
    /// secondaryAxis: ["Series B", ...] moves each named line/bar series onto a
    /// secondary value axis (a new line group sharing a hidden category axis). The
    /// series is detached from its primary group and re-homed on the secondary pair.
    /// </summary>
    private static void ApplySecondaryAxis(C.PlotArea plotArea, JsonNode? value)
    {
        if (value is not JsonArray array)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"secondaryAxis must be an array of series names: {value?.ToJsonString() ?? "null"}",
                "Use {\"secondaryAxis\":[\"Series B\"]} — each named series moves to a secondary value axis.");
        }

        var wanted = array.Select(n => n is null ? string.Empty : J.ScalarText(n)).ToList();
        if (wanted.Count == 0)
        {
            return;
        }

        // Find the named series in the primary group(s); collect (ser, sourceGroup).
        var moving = new List<(OpenXmlCompositeElement Ser, OpenXmlCompositeElement Group)>();
        foreach (var group in ChartGroups(plotArea))
        {
            foreach (var ser in group.ChildElements.Where(e => e.LocalName == "ser").Cast<OpenXmlCompositeElement>().ToList())
            {
                if (wanted.Contains(SeriesName(ser), StringComparer.Ordinal) && ser.Elements<C.Values>().Any())
                {
                    moving.Add((ser, group));
                }
            }
        }

        if (moving.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"None of the named series were found: {string.Join(", ", wanted)}.",
                "secondaryAxis names must match the chart's series; run 'aioffice get <file> " +
                "<chartPath>' to list them.");
        }

        EnsureSecondaryAxes(plotArea);

        // A line group hung on the secondary axis pair carries the moved series.
        var secondaryGroup = new C.LineChart(
            new C.Grouping { Val = C.GroupingValues.Standard },
            new C.VaryColors { Val = false });
        foreach (var (ser, group) in moving)
        {
            ser.Remove();
            // A moved bar/area series becomes a line series under the secondary group;
            // the cloned children carry the same idx/order/tx/cat/val caches.
            secondaryGroup.Append(ConvertToLineSeries(ser));

            // Drop an emptied primary group (no series left) to keep the plot valid.
            if (!group.ChildElements.Any(e => e.LocalName == "ser"))
            {
                group.Remove();
            }
        }

        secondaryGroup.Append(new C.AxisId { Val = SecondaryCategoryAxisId });
        secondaryGroup.Append(new C.AxisId { Val = SecondaryValueAxisId });

        // Chart groups precede the axes in a plotArea; insert before the first axis.
        var firstAxis = plotArea.ChildElements.FirstOrDefault(e => e.LocalName.EndsWith("Ax", StringComparison.Ordinal));
        if (firstAxis is not null)
        {
            plotArea.InsertBefore(secondaryGroup, firstAxis);
        }
        else
        {
            plotArea.Append(secondaryGroup);
        }
    }

    /// <summary>A ser re-homed as a line series: same idx/order/tx/cat/val, no bar-only children.</summary>
    private static C.LineChartSeries ConvertToLineSeries(OpenXmlCompositeElement source)
    {
        var line = new C.LineChartSeries();
        foreach (var child in source.ChildElements)
        {
            if (child.LocalName is "idx" or "order" or "tx" or "cat" or "val" or "spPr")
            {
                line.Append(child.CloneNode(true));
            }
        }

        return line;
    }

    /// <summary>Adds a secondary category axis (deleted/hidden) + value axis (right) crossing it, once.</summary>
    private static void EnsureSecondaryAxes(C.PlotArea plotArea)
    {
        if (plotArea.Elements<C.ValueAxis>().Any(a => a.AxisId?.Val?.Value == SecondaryValueAxisId))
        {
            return;
        }

        plotArea.Append(new C.ValueAxis(
            new C.AxisId { Val = SecondaryValueAxisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = C.AxisPositionValues.Right },
            new C.CrossingAxis { Val = SecondaryCategoryAxisId },
            new C.Crosses { Val = C.CrossesValues.Maximum }));

        plotArea.Append(new C.CategoryAxis(
            new C.AxisId { Val = SecondaryCategoryAxisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = true },
            new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
            new C.CrossingAxis { Val = SecondaryValueAxisId }));
    }

    // ----- read-back (get projection) ------------------------------------------

    /// <summary>The polish settings a chart currently carries, for the `get` projection (null sub-fields omitted on the wire).</summary>
    public static object ReadSettings(C.ChartSpace chartSpace)
    {
        var chart = chartSpace.GetFirstChild<C.Chart>();
        var plotArea = chart?.GetFirstChild<C.PlotArea>();

        return new
        {
            DataLabels = ReadDataLabels(plotArea),
            Legend = ReadLegend(chart),
            AxisTitles = ReadAxisTitles(plotArea),
            Trendline = ReadTrendline(plotArea),
            ErrorBars = ReadErrorBars(plotArea),
            Gridlines = ReadGridlines(plotArea),
            SecondaryAxis = ReadSecondaryAxis(plotArea),
        };
    }

    private static object? ReadDataLabels(C.PlotArea? plotArea)
    {
        var labels = plotArea?.Descendants<C.DataLabels>().FirstOrDefault(d => d.Parent is not null && d.Parent.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        if (labels is null)
        {
            return null;
        }

        var show = labels.GetFirstChild<C.ShowValue>()?.Val?.Value == true ? "value"
            : labels.GetFirstChild<C.ShowPercent>()?.Val?.Value == true ? "percent"
            : labels.GetFirstChild<C.ShowCategoryName>()?.Val?.Value == true ? "category"
            : labels.GetFirstChild<C.ShowSeriesName>()?.Val?.Value == true ? "seriesName"
            : "value";
        var position = labels.GetFirstChild<C.DataLabelPosition>()?.Val?.Value switch
        {
            { } p when p == C.DataLabelPositionValues.InsideEnd => "inEnd",
            { } p when p == C.DataLabelPositionValues.Center => "center",
            { } p when p == C.DataLabelPositionValues.BestFit => "bestFit",
            { } p when p == C.DataLabelPositionValues.OutsideEnd => "outEnd",
            _ => null,
        };

        return new { Show = show, Position = position };
    }

    private static string? ReadLegend(C.Chart? chart)
    {
        var legend = chart?.GetFirstChild<C.Legend>();
        if (legend is null)
        {
            return null;
        }

        return legend.GetFirstChild<C.LegendPosition>()?.Val?.Value switch
        {
            { } p when p == C.LegendPositionValues.Left => "left",
            { } p when p == C.LegendPositionValues.Top => "top",
            { } p when p == C.LegendPositionValues.Bottom => "bottom",
            _ => "right",
        };
    }

    private static object? ReadAxisTitles(C.PlotArea? plotArea)
    {
        var category = AxisTitleText(plotArea?.GetFirstChild<C.CategoryAxis>());
        var value = AxisTitleText(plotArea?.GetFirstChild<C.ValueAxis>());
        return category is null && value is null ? null : new { Category = category, Value = value };
    }

    private static string? AxisTitleText(OpenXmlCompositeElement? axis)
    {
        var text = axis?.GetFirstChild<C.Title>()?.InnerText;
        return string.IsNullOrEmpty(text) ? null : text;
    }

    private static string? ReadTrendline(C.PlotArea? plotArea)
    {
        var type = plotArea?.Descendants<C.Trendline>().FirstOrDefault()?.GetFirstChild<C.TrendlineType>()?.Val?.Value;
        return type switch
        {
            { } t when t == C.TrendlineValues.Exponential => "exponential",
            { } t when t == C.TrendlineValues.MovingAverage => "movingAverage",
            { } t when t == C.TrendlineValues.Linear => "linear",
            _ => null,
        };
    }

    private static string? ReadErrorBars(C.PlotArea? plotArea)
    {
        var type = plotArea?.Descendants<C.ErrorBars>().FirstOrDefault()?.GetFirstChild<C.ErrorBarValueType>()?.Val?.Value;
        return type switch
        {
            { } t when t == C.ErrorValues.StandardDeviation => "stdDev",
            { } t when t == C.ErrorValues.Percentage => "percent",
            { } t when t == C.ErrorValues.StandardError => "stdErr",
            _ => null,
        };
    }

    private static object? ReadGridlines(C.PlotArea? plotArea)
    {
        var valueAxis = plotArea?.GetFirstChild<C.ValueAxis>();
        if (valueAxis is null)
        {
            return null;
        }

        var major = valueAxis.GetFirstChild<C.MajorGridlines>() is not null;
        var minor = valueAxis.GetFirstChild<C.MinorGridlines>() is not null;
        return major || minor ? new { Major = major, Minor = minor } : null;
    }

    private static IReadOnlyList<string>? ReadSecondaryAxis(C.PlotArea? plotArea)
    {
        if (plotArea is null || !plotArea.Elements<C.ValueAxis>().Any(a => a.AxisId?.Val?.Value == SecondaryValueAxisId))
        {
            return null;
        }

        var names = new List<string>();
        foreach (var group in ChartGroups(plotArea))
        {
            var axisIds = group.Elements<C.AxisId>().Select(a => a.Val?.Value).ToList();
            if (axisIds.Contains(SecondaryValueAxisId))
            {
                foreach (var ser in group.ChildElements.Where(e => e.LocalName == "ser").Cast<OpenXmlCompositeElement>())
                {
                    names.Add(SeriesName(ser));
                }
            }
        }

        return names.Count == 0 ? null : names;
    }

    // ----- render hints --------------------------------------------------------

    /// <summary>The legend position ("none"/"right"/...) and whether data labels are on, for the SVG mini-chart.</summary>
    public static (string Legend, bool DataLabels) RenderHints(C.ChartSpace chartSpace)
    {
        var chart = chartSpace.GetFirstChild<C.Chart>();
        var plotArea = chart?.GetFirstChild<C.PlotArea>();
        return (ReadLegend(chart) ?? "none", ReadDataLabels(plotArea) is not null);
    }

    // ----- shared helpers ------------------------------------------------------

    /// <summary>Every plot-area chart group (barChart/lineChart/…), in document order.</summary>
    private static IEnumerable<OpenXmlCompositeElement> ChartGroups(C.PlotArea plotArea) =>
        plotArea.ChildElements
            .Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal))
            .Cast<OpenXmlCompositeElement>();

    /// <summary>Every series across every chart group.</summary>
    private static IEnumerable<OpenXmlCompositeElement> AllSeries(C.PlotArea plotArea) =>
        ChartGroups(plotArea).SelectMany(g => g.ChildElements.Where(e => e.LocalName == "ser").Cast<OpenXmlCompositeElement>());

    private static string SeriesName(OpenXmlCompositeElement ser) =>
        ser.ChildElements.FirstOrDefault(e => e.LocalName == "tx")?
            .Descendants<C.NumericValue>().FirstOrDefault()?.Text ?? string.Empty;

    /// <summary>Inserts a c:dLbls into a chart group at a schema-valid spot (right after the last c:ser).</summary>
    private static void InsertGroupChild(OpenXmlCompositeElement group, OpenXmlElement child)
    {
        var lastSeries = group.ChildElements.LastOrDefault(e => e.LocalName == "ser");
        if (lastSeries is not null)
        {
            group.InsertAfter(child, lastSeries);
            return;
        }

        // No series (defensive): place before the first c:axId.
        var firstAxisId = group.ChildElements.FirstOrDefault(e => e.LocalName == "axId");
        if (firstAxisId is not null)
        {
            group.InsertBefore(child, firstAxisId);
        }
        else
        {
            group.Append(child);
        }
    }

    private static string Canonical(string raw, IReadOnlyList<string> allowed, string what)
    {
        foreach (var option in allowed)
        {
            if (string.Equals(option, raw, StringComparison.OrdinalIgnoreCase))
            {
                return option;
            }
        }

        throw Unsupported(Units.Inv($"{what} '{raw}'"), string.Join(", ", allowed));
    }

    private static AiofficeException Unsupported(string what, string supported) => new(
        ErrorCodes.UnsupportedFeature,
        $"Unsupported {what}.",
        $"Supported: {supported}.");

    private static bool AsBool(string key, JsonNode node)
    {
        if (node is JsonValue value)
        {
            if (value.TryGetValue<bool>(out var flag))
            {
                return flag;
            }

            if (value.TryGetValue<string>(out var text) && bool.TryParse(text, out var parsed))
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Property '{key}' is not a boolean: {node.ToJsonString()}",
            "Use true or false.");
    }
}
