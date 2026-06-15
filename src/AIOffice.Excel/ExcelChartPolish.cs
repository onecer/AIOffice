using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// The v1.3 chart-polish layer for xlsx charts (additive, shared with the create
/// path). It parses the polish props — <c>dataLabels</c>, <c>legend</c>,
/// <c>axisTitles</c>, <c>trendline</c>, <c>errorBars</c>, <c>gridlines</c>,
/// <c>secondaryAxis</c> — and applies them to a chart's OpenXml DOM. The same
/// <see cref="Apply"/> serves both <c>add type:chart</c> (the chart built fresh)
/// and <c>set /Sheet/chart[i]</c> (an existing chart reopened raw in a post-save
/// pass), so the two routes can never drift.
///
/// chartml element map: <c>dataLabels</c> → <c>c:dLbls</c> (per-series, schema
/// slot after <c>order</c>/<c>tx</c>, before <c>cat</c>); <c>legend</c> →
/// <c>c:legend</c>; <c>axisTitles</c> → <c>c:title</c> on the category / value
/// axis; <c>trendline</c> → <c>c:trendline</c> per series; <c>errorBars</c> →
/// <c>c:errBars</c> per series; <c>gridlines</c> → <c>c:majorGridlines</c> /
/// <c>c:minorGridlines</c> on the value axis; <c>secondaryAxis</c> moves the
/// named series onto a second value axis (a new <c>c:valAx</c> plus a hidden
/// secondary <c>c:catAx</c>). Every result stays OpenXmlValidator-clean.
/// </summary>
internal static class ExcelChartPolish
{
    /// <summary>The polish prop keys, accepted at chart create and on a chart set.</summary>
    public static readonly IReadOnlyList<string> Props =
        ["dataLabels", "legend", "axisTitles", "trendline", "errorBars", "gridlines", "secondaryAxis"];

    private static readonly IReadOnlyList<string> LegendPositions = ["none", "right", "left", "top", "bottom"];

    private static readonly IReadOnlyList<string> DataLabelShows = ["value", "percent", "category", "seriesName"];

    private static readonly IReadOnlyList<string> DataLabelPositions = ["outEnd", "inEnd", "center", "bestFit"];

    private static readonly IReadOnlyList<string> TrendlineKinds = ["none", "linear", "exponential", "movingAverage"];

    private static readonly IReadOnlyList<string> ErrorBarKinds = ["none", "stdErr", "stdDev", "percent"];

    private const uint SecondaryValueAxisId = 100003u;
    private const uint SecondaryCategoryAxisId = 100004u;

    // The primary axis ids the create path uses (ExcelCharts mirrors these).
    private const uint CategoryAxisId = 100001u;
    private const uint ValueAxisId = 100002u;

    // ----- parse --------------------------------------------------------------

    /// <summary>The parsed, validated polish settings for one chart.</summary>
    internal sealed record PolishSettings(
        DataLabelSetting? DataLabels,
        string? Legend,
        AxisTitleSetting? AxisTitles,
        string? Trendline,
        string? ErrorBars,
        GridlineSetting? Gridlines,
        IReadOnlyList<string>? SecondaryAxisSeries)
    {
        /// <summary>True when no polish prop was supplied (nothing to apply).</summary>
        public bool IsEmpty =>
            DataLabels is null && Legend is null && AxisTitles is null && Trendline is null &&
            ErrorBars is null && Gridlines is null && SecondaryAxisSeries is null;
    }

    internal sealed record DataLabelSetting(bool Show, string? What, string? Position);

    internal sealed record AxisTitleSetting(string? Category, string? Value);

    internal sealed record GridlineSetting(bool? Major, bool? Minor);

    /// <summary>A queued chart-polish edit (set on a chart path) for the post-save raw pass.</summary>
    internal sealed record EditSpec(string SheetName, int ChartIndex, PolishSettings Settings);

    /// <summary>True when at least one polish prop is present in the op props.</summary>
    public static bool HasPolishProps(JsonObject? props) =>
        props is not null && props.Any(p => Props.Contains(p.Key, StringComparer.Ordinal));

    /// <summary>
    /// Parses the polish props out of an op's props (ignoring non-polish keys —
    /// the caller validates those). Throws <c>invalid_args</c> /
    /// <c>unsupported_feature</c> for malformed or unsupported sub-values.
    /// </summary>
    public static PolishSettings Parse(JsonObject props, int opIndex)
    {
        return new PolishSettings(
            ParseDataLabels(props, opIndex),
            ParseLegend(props, opIndex),
            ParseAxisTitles(props, opIndex),
            ParseTrendline(props, opIndex),
            ParseErrorBars(props, opIndex),
            ParseGridlines(props, opIndex),
            ParseSecondaryAxis(props, opIndex));
    }

    private static DataLabelSetting? ParseDataLabels(JsonObject props, int opIndex)
    {
        if (!props.TryGetPropertyValue("dataLabels", out var node) || node is null)
        {
            return null;
        }

        // dataLabels: true → show value at the default position.
        if (node is JsonValue boolValue && boolValue.GetValueKind() is JsonValueKind.True or JsonValueKind.False)
        {
            return boolValue.GetValue<bool>() ? new DataLabelSetting(true, "value", null) : new DataLabelSetting(false, null, null);
        }

        if (node is not JsonObject obj)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'dataLabels' must be true or an object like {{\"show\":\"value\"}}.",
                "Pass dataLabels:true to show values, or {\"show\":\"percent\",\"position\":\"outEnd\"}.");
        }

        var what = StringField(obj, "show", opIndex, "dataLabels") ?? "value";
        if (!DataLabelShows.Contains(what, StringComparer.Ordinal))
        {
            throw Unsupported(opIndex, $"dataLabels.show '{what}'", DataLabelShows);
        }

        var position = StringField(obj, "position", opIndex, "dataLabels");
        if (position is not null && !DataLabelPositions.Contains(position, StringComparer.Ordinal))
        {
            throw Unsupported(opIndex, $"dataLabels.position '{position}'", DataLabelPositions);
        }

        foreach (var (key, _) in obj)
        {
            if (key is not ("show" or "position"))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown dataLabels field '{key}'.",
                    "dataLabels accepts {show, position}.",
                    candidates: ["show", "position"]);
            }
        }

        return new DataLabelSetting(true, what, position);
    }

    private static string? ParseLegend(JsonObject props, int opIndex)
    {
        var legend = StringField(props, "legend", opIndex, "legend");
        if (legend is null)
        {
            return null;
        }

        if (!LegendPositions.Contains(legend, StringComparer.Ordinal))
        {
            throw Unsupported(opIndex, $"legend '{legend}'", LegendPositions);
        }

        return legend;
    }

    private static AxisTitleSetting? ParseAxisTitles(JsonObject props, int opIndex)
    {
        if (!props.TryGetPropertyValue("axisTitles", out var node) || node is null)
        {
            return null;
        }

        if (node is not JsonObject obj)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'axisTitles' must be an object like {{\"category\":\"X\",\"value\":\"Y\"}}.",
                "Pass axisTitles:{category:\"Month\", value:\"Sales\"}; either key is optional.");
        }

        foreach (var (key, _) in obj)
        {
            if (key is not ("category" or "value"))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown axisTitles field '{key}'.",
                    "axisTitles accepts {category, value}.",
                    candidates: ["category", "value"]);
            }
        }

        var category = StringField(obj, "category", opIndex, "axisTitles");
        var value = StringField(obj, "value", opIndex, "axisTitles");
        if (category is null && value is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: axisTitles has neither a category nor a value title.",
                "Pass at least one, e.g. {\"value\":\"Sales\"}.");
        }

        return new AxisTitleSetting(category, value);
    }

    private static string? ParseTrendline(JsonObject props, int opIndex)
    {
        var trendline = StringField(props, "trendline", opIndex, "trendline");
        if (trendline is null)
        {
            return null;
        }

        if (!TrendlineKinds.Contains(trendline, StringComparer.Ordinal))
        {
            throw Unsupported(opIndex, $"trendline '{trendline}'", TrendlineKinds);
        }

        return trendline;
    }

    private static string? ParseErrorBars(JsonObject props, int opIndex)
    {
        var errorBars = StringField(props, "errorBars", opIndex, "errorBars");
        if (errorBars is null)
        {
            return null;
        }

        if (!ErrorBarKinds.Contains(errorBars, StringComparer.Ordinal))
        {
            throw Unsupported(opIndex, $"errorBars '{errorBars}'", ErrorBarKinds);
        }

        return errorBars;
    }

    private static GridlineSetting? ParseGridlines(JsonObject props, int opIndex)
    {
        if (!props.TryGetPropertyValue("gridlines", out var node) || node is null)
        {
            return null;
        }

        if (node is not JsonObject obj)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'gridlines' must be an object like {{\"major\":true,\"minor\":false}}.",
                "Pass gridlines:{major:true} to show major value-axis gridlines.");
        }

        foreach (var (key, _) in obj)
        {
            if (key is not ("major" or "minor"))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown gridlines field '{key}'.",
                    "gridlines accepts {major, minor}.",
                    candidates: ["major", "minor"]);
            }
        }

        var major = BoolField(obj, "major", opIndex);
        var minor = BoolField(obj, "minor", opIndex);
        if (major is null && minor is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: gridlines sets neither major nor minor.",
                "Pass at least one, e.g. {\"major\":true}.");
        }

        return new GridlineSetting(major, minor);
    }

    private static IReadOnlyList<string>? ParseSecondaryAxis(JsonObject props, int opIndex)
    {
        if (!props.TryGetPropertyValue("secondaryAxis", out var node) || node is null)
        {
            return null;
        }

        if (node is not JsonArray array)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'secondaryAxis' must be an array of series names like [\"Cost\"].",
                "Name the series to move onto the secondary value axis, e.g. secondaryAxis:[\"Growth %\"].");
        }

        var names = new List<string>(array.Count);
        foreach (var entry in array)
        {
            if (entry is not JsonValue value || value.GetValueKind() != JsonValueKind.String ||
                value.GetValue<string>() is not { Length: > 0 } name)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: secondaryAxis entries must be non-empty series names.",
                    "Pass e.g. secondaryAxis:[\"Cost\"].");
            }

            names.Add(name);
        }

        if (names.Count == 0)
        {
            // An empty list means "no series on the secondary axis": treat as a clear.
            return [];
        }

        return names;
    }

    // ----- apply (shared by create + edit) -----------------------------------

    /// <summary>
    /// Applies the polish settings to a chart's <c>c:chart</c> DOM in place. Used
    /// both when building a chart fresh (create) and when re-polishing an existing
    /// one (set). Idempotent per facet: each facet is rebuilt from scratch so a
    /// later set overrides an earlier one rather than stacking.
    /// </summary>
    public static void Apply(C.Chart chart, PolishSettings settings)
    {
        if (settings.IsEmpty)
        {
            return;
        }

        var plotArea = chart.GetFirstChild<C.PlotArea>();
        if (plotArea is null)
        {
            return;
        }

        if (settings.DataLabels is { } dataLabels)
        {
            ApplyDataLabels(plotArea, dataLabels);
        }

        if (settings.Trendline is { } trendline)
        {
            ApplyTrendline(plotArea, trendline);
        }

        if (settings.ErrorBars is { } errorBars)
        {
            ApplyErrorBars(plotArea, errorBars);
        }

        if (settings.SecondaryAxisSeries is { } secondary)
        {
            ApplySecondaryAxis(plotArea, secondary);
        }

        if (settings.AxisTitles is { } axisTitles)
        {
            ApplyAxisTitles(plotArea, axisTitles);
        }

        if (settings.Gridlines is { } gridlines)
        {
            ApplyGridlines(plotArea, gridlines);
        }

        if (settings.Legend is { } legend)
        {
            ApplyLegend(chart, legend);
        }
    }

    private static IEnumerable<OpenXmlCompositeElement> SeriesOf(C.PlotArea plotArea) =>
        ChartGroups(plotArea)
            .SelectMany(g => g.ChildElements.Where(e => e.LocalName == "ser"))
            .Cast<OpenXmlCompositeElement>();

    private static IEnumerable<OpenXmlCompositeElement> ChartGroups(C.PlotArea plotArea) =>
        plotArea.ChildElements
            .Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal))
            .Cast<OpenXmlCompositeElement>();

    private static void ApplyDataLabels(C.PlotArea plotArea, DataLabelSetting setting)
    {
        foreach (var series in SeriesOf(plotArea))
        {
            series.RemoveAllChildren<C.DataLabels>();
            if (!setting.Show)
            {
                continue;
            }

            var dLbls = BuildDataLabels(setting);
            InsertSeriesChild(series, dLbls);
        }
    }

    private static C.DataLabels BuildDataLabels(DataLabelSetting setting)
    {
        var dLbls = new C.DataLabels();
        if (setting.Position is { } position)
        {
            dLbls.Append(new C.DataLabelPosition { Val = DataLabelPositionValue(position) });
        }

        // The five show-flags are required children in schema order; flip the one
        // the agent asked for.
        dLbls.Append(new C.ShowLegendKey { Val = false });
        dLbls.Append(new C.ShowValue { Val = setting.What == "value" });
        dLbls.Append(new C.ShowCategoryName { Val = setting.What == "category" });
        dLbls.Append(new C.ShowSeriesName { Val = setting.What == "seriesName" });
        dLbls.Append(new C.ShowPercent { Val = setting.What == "percent" });
        dLbls.Append(new C.ShowBubbleSize { Val = false });
        return dLbls;
    }

    private static void ApplyTrendline(C.PlotArea plotArea, string trendline)
    {
        foreach (var series in SeriesOf(plotArea))
        {
            series.RemoveAllChildren<C.Trendline>();
            if (trendline == "none")
            {
                continue;
            }

            var element = new C.Trendline(new C.TrendlineType { Val = TrendlineValue(trendline) });
            if (trendline == "movingAverage")
            {
                element.Append(new C.Period { Val = 2 });
            }

            InsertSeriesChild(series, element);
        }
    }

    private static void ApplyErrorBars(C.PlotArea plotArea, string errorBars)
    {
        foreach (var series in SeriesOf(plotArea))
        {
            series.RemoveAllChildren<C.ErrorBars>();
            if (errorBars == "none")
            {
                continue;
            }

            InsertSeriesChild(series, BuildErrorBars(errorBars));
        }
    }

    private static C.ErrorBars BuildErrorBars(string kind)
    {
        var bars = new C.ErrorBars(
            new C.ErrorDirection { Val = C.ErrorBarDirectionValues.Y },
            new C.ErrorBarType { Val = C.ErrorBarValues.Both },
            new C.ErrorBarValueType { Val = ErrorValue(kind) },
            new C.NoEndCap { Val = false });

        // A percentage error bar carries its percent in c:val; the stat-derived
        // ones (stdErr/stdDev) compute their own magnitude and take none.
        if (kind == "percent")
        {
            bars.Append(new C.ErrorBarValue { Val = 5d });
        }

        return bars;
    }

    /// <summary>
    /// Inserts a per-series polish child (<c>c:dLbls</c> / <c>c:trendline</c> /
    /// <c>c:errBars</c>) in its schema slot: after the last of
    /// idx/order/tx/spPr/invertIfNegative/pictureOptions/dPt and before
    /// cat/xVal/yVal/val/bubbleSize/smooth. We anchor on the data refs that every
    /// series carries (cat/val for category kinds, xVal for scatter/bubble).
    /// </summary>
    private static void InsertSeriesChild(OpenXmlCompositeElement series, OpenXmlElement child)
    {
        var anchor = series.ChildElements.FirstOrDefault(e =>
            e.LocalName is "cat" or "xVal" or "yVal" or "val" or "bubbleSize" or "smooth");
        if (anchor is null)
        {
            series.Append(child);
        }
        else
        {
            series.InsertBefore(child, anchor);
        }
    }

    private static void ApplyAxisTitles(C.PlotArea plotArea, AxisTitleSetting setting)
    {
        if (setting.Category is { } category)
        {
            // Scatter/bubble have no category axis, so the "category" title lands
            // on the bottom value axis (the X axis there); otherwise the catAx.
            OpenXmlCompositeElement? axis = PrimaryCategoryAxis(plotArea);
            axis ??= BottomValueAxis(plotArea);
            if (axis is not null)
            {
                SetAxisTitle(axis, category);
            }
        }

        if (setting.Value is { } value)
        {
            var axis = PrimaryValueAxis(plotArea);
            if (axis is not null)
            {
                SetAxisTitle(axis, value);
            }
        }
    }

    private static void ApplyGridlines(C.PlotArea plotArea, GridlineSetting setting)
    {
        var axis = PrimaryValueAxis(plotArea);
        if (axis is null)
        {
            return;
        }

        if (setting.Major is { } major)
        {
            axis.RemoveAllChildren<C.MajorGridlines>();
            if (major)
            {
                InsertAxisChild(axis, new C.MajorGridlines(), beforeTitle: true);
            }
        }

        if (setting.Minor is { } minor)
        {
            axis.RemoveAllChildren<C.MinorGridlines>();
            if (minor)
            {
                InsertAxisChild(axis, new C.MinorGridlines(), beforeTitle: true);
            }
        }
    }

    private static void ApplyLegend(C.Chart chart, string legend)
    {
        chart.RemoveAllChildren<C.Legend>();
        if (legend == "none")
        {
            return;
        }

        var element = new C.Legend(
            new C.LegendPosition { Val = LegendPositionValue(legend) },
            new C.Overlay { Val = false });

        // c:legend sits after plotArea and before plotVisibleOnly in c:chart.
        var plotVisible = chart.GetFirstChild<C.PlotVisibleOnly>();
        if (plotVisible is not null)
        {
            chart.InsertBefore(element, plotVisible);
        }
        else
        {
            var plotArea = chart.GetFirstChild<C.PlotArea>();
            chart.InsertAfter(element, (OpenXmlElement?)plotArea ?? chart.LastChild!);
        }
    }

    // ----- secondary axis -----------------------------------------------------

    /// <summary>
    /// Moves the named series onto a secondary value axis. The series are pulled
    /// out of their current group into a fresh group of the same kind that points
    /// at a new value axis (and a hidden secondary category axis); the existing
    /// primary axes stay put. Re-running with a new list rebuilds from a single
    /// primary group first, so the operation is idempotent.
    /// </summary>
    private static void ApplySecondaryAxis(C.PlotArea plotArea, IReadOnlyList<string> seriesNames)
    {
        // Collapse any prior secondary split back into one primary group so a
        // re-set starts from a clean slate.
        CollapseSecondaryAxis(plotArea);
        if (seriesNames.Count == 0)
        {
            return;
        }

        var groups = ChartGroups(plotArea).ToList();
        var primaryGroup = groups.FirstOrDefault();
        if (primaryGroup is null)
        {
            return;
        }

        var wanted = new HashSet<string>(seriesNames, StringComparer.OrdinalIgnoreCase);
        var moved = new List<OpenXmlElement>();
        foreach (var series in primaryGroup.ChildElements.Where(e => e.LocalName == "ser").ToList())
        {
            if (wanted.Contains(SeriesName(series)))
            {
                series.Remove();
                moved.Add(series);
            }
        }

        if (moved.Count == 0)
        {
            // No series matched: nothing to split. The caller validates names, so
            // this is only reachable for a chart with renamed series.
            return;
        }

        // A new group of the same kind on the secondary axes, holding the moved
        // series. Axis-id elements go last in a chart group.
        var secondaryGroup = CloneGroupShell(primaryGroup);
        foreach (var series in moved)
        {
            secondaryGroup.Append(series);
        }

        secondaryGroup.Append(new C.AxisId { Val = SecondaryCategoryAxisId });
        secondaryGroup.Append(new C.AxisId { Val = SecondaryValueAxisId });

        // Insert the secondary group right after the last existing chart group.
        var lastGroup = ChartGroups(plotArea).Last();
        plotArea.InsertAfter(secondaryGroup, lastGroup);

        AppendSecondaryAxes(plotArea);
    }

    /// <summary>A bare group of the same kind as <paramref name="group"/> (grouping flags copied, series dropped).</summary>
    private static OpenXmlCompositeElement CloneGroupShell(OpenXmlCompositeElement group)
    {
        var shell = (OpenXmlCompositeElement)group.CloneNode(deep: false);
        // Copy the leading config children (barDir, grouping, varyColors, …) up to
        // the first series; drop everything from the first series onward.
        foreach (var child in group.ChildElements)
        {
            if (child.LocalName is "ser" or "axId" or "marker" or "gapWidth" or "overlap")
            {
                continue;
            }

            shell.Append(child.CloneNode(deep: true));
        }

        return shell;
    }

    /// <summary>Appends the secondary value axis (right) and a hidden secondary category axis.</summary>
    private static void AppendSecondaryAxes(C.PlotArea plotArea)
    {
        if (plotArea.Descendants<C.ValueAxis>()
            .Any(a => a.GetFirstChild<C.AxisId>()?.Val?.Value == SecondaryValueAxisId))
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

    /// <summary>
    /// Folds any secondary-axis group back into the primary group and removes the
    /// secondary axes — the inverse of <see cref="ApplySecondaryAxis"/>, so a
    /// re-set is a clean rebuild.
    /// </summary>
    private static void CollapseSecondaryAxis(C.PlotArea plotArea)
    {
        var secondaryGroups = ChartGroups(plotArea)
            .Where(g => g.ChildElements.OfType<C.AxisId>().Any(a => a.Val?.Value == SecondaryValueAxisId))
            .ToList();
        if (secondaryGroups.Count == 0)
        {
            return;
        }

        var primaryGroup = ChartGroups(plotArea)
            .FirstOrDefault(g => g.ChildElements.OfType<C.AxisId>().Any(a => a.Val?.Value == ValueAxisId));
        foreach (var group in secondaryGroups)
        {
            var movedBack = group.ChildElements.Where(e => e.LocalName == "ser").ToList();
            if (primaryGroup is not null)
            {
                var anchor = primaryGroup.ChildElements.FirstOrDefault(e => e.LocalName == "axId");
                foreach (var series in movedBack)
                {
                    series.Remove();
                    if (anchor is not null)
                    {
                        primaryGroup.InsertBefore(series, anchor);
                    }
                    else
                    {
                        primaryGroup.Append(series);
                    }
                }
            }

            group.Remove();
        }

        foreach (var axis in plotArea.Descendants<C.ValueAxis>().ToList())
        {
            if (axis.GetFirstChild<C.AxisId>()?.Val?.Value == SecondaryValueAxisId)
            {
                axis.Remove();
            }
        }

        foreach (var axis in plotArea.Descendants<C.CategoryAxis>().ToList())
        {
            if (axis.GetFirstChild<C.AxisId>()?.Val?.Value == SecondaryCategoryAxisId)
            {
                axis.Remove();
            }
        }
    }

    // ----- axis helpers -------------------------------------------------------

    private static C.CategoryAxis? PrimaryCategoryAxis(C.PlotArea plotArea) =>
        plotArea.Elements<C.CategoryAxis>()
            .FirstOrDefault(a => a.GetFirstChild<C.AxisId>()?.Val?.Value != SecondaryCategoryAxisId);

    private static C.ValueAxis? PrimaryValueAxis(C.PlotArea plotArea) =>
        plotArea.Elements<C.ValueAxis>()
            .FirstOrDefault(a => a.GetFirstChild<C.AxisId>()?.Val?.Value != SecondaryValueAxisId);

    private static C.ValueAxis? BottomValueAxis(C.PlotArea plotArea) =>
        plotArea.Elements<C.ValueAxis>()
            .FirstOrDefault(a => a.GetFirstChild<C.AxisPosition>()?.Val?.Value == C.AxisPositionValues.Bottom);

    private static void SetAxisTitle(OpenXmlCompositeElement axis, string text)
    {
        axis.RemoveAllChildren<C.Title>();
        // The title is empty-cleared by an empty string (drop it).
        if (text.Length == 0)
        {
            return;
        }

        var title = new C.Title(
            new C.ChartText(new C.RichText(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.Text(text))))),
            new C.Overlay { Val = false });
        InsertAxisChild(axis, title, beforeTitle: false);
    }

    /// <summary>
    /// Inserts a child into an axis in its schema slot. Axis child order is
    /// axId, scaling, delete, axisPosition, majorGridlines, minorGridlines, title,
    /// numFmt, …, crossingAxis (crossAx), crosses. Gridlines go before the title;
    /// the title goes after gridlines and before crossAx.
    /// </summary>
    private static void InsertAxisChild(OpenXmlCompositeElement axis, OpenXmlElement child, bool beforeTitle)
    {
        if (child is C.MajorGridlines)
        {
            // majorGridlines goes right after axisPosition.
            var axisPosition = axis.GetFirstChild<C.AxisPosition>();
            if (axisPosition is not null)
            {
                axis.InsertAfter(child, axisPosition);
                return;
            }
        }

        if (child is C.MinorGridlines)
        {
            // minorGridlines goes after majorGridlines (or axisPosition).
            var after = (OpenXmlElement?)axis.GetFirstChild<C.MajorGridlines>() ?? axis.GetFirstChild<C.AxisPosition>();
            if (after is not null)
            {
                axis.InsertAfter(child, after);
                return;
            }
        }

        // Title: after gridlines, before crossingAxis (crossAx).
        var crossingAxis = axis.GetFirstChild<C.CrossingAxis>();
        if (crossingAxis is not null)
        {
            axis.InsertBefore(child, crossingAxis);
            return;
        }

        axis.Append(child);
    }

    // ----- post-save edit pass ------------------------------------------------

    /// <summary>
    /// Applies queued chart-polish edits to the file ClosedXML just saved. Each
    /// edit reopens the chart part addressed by (sheet, 1-based index), mutates
    /// its DOM, and saves. ClosedXML preserves chart parts byte-identical, so the
    /// edits survive subsequent saves.
    /// </summary>
    public static void ApplyEdits(string file, IReadOnlyList<EditSpec> edits)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        foreach (var edit in edits)
        {
            var chartPart = ChartPartFor(document, edit.SheetName, edit.ChartIndex);
            if (chartPart?.ChartSpace?.GetFirstChild<C.Chart>() is not { } chart)
            {
                throw new AiofficeException(
                    ErrorCodes.InternalError,
                    $"chart[{edit.ChartIndex}] on '{edit.SheetName}' disappeared between validation and the polish write pass.",
                    "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
            }

            Apply(chart, edit.Settings);
            chartPart.ChartSpace.Save();
        }
    }

    /// <summary>The chart part at a sheet's 1-based drawing index (null when absent).</summary>
    public static ChartPart? ChartPartFor(SpreadsheetDocument document, string sheetName, int index)
    {
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart?.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id?.Value is not { } relationshipId ||
            workbookPart!.GetPartById(relationshipId) is not WorksheetPart { DrawingsPart: { } drawings })
        {
            return null;
        }

        var ordinal = 0;
        foreach (var child in drawings.WorksheetDrawing?.ChildElements ?? [])
        {
            if (child is not (Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor))
            {
                continue;
            }

            if (((OpenXmlCompositeElement)child).Descendants<C.ChartReference>().FirstOrDefault()?.Id?.Value
                is not { } chartRelId)
            {
                continue;
            }

            var pair = drawings.Parts.FirstOrDefault(p => p.RelationshipId == chartRelId);
            if (pair.OpenXmlPart is ChartPart chartPart)
            {
                ordinal++;
                if (ordinal == index)
                {
                    return chartPart;
                }
            }
        }

        return null;
    }

    // ----- describe (get) -----------------------------------------------------

    /// <summary>
    /// Reads a chart's current polish settings back for <c>get</c>. Returns null
    /// for facets that are not set so the get payload only carries what is there.
    /// </summary>
    public static object DescribePolish(C.Chart chart)
    {
        var plotArea = chart.GetFirstChild<C.PlotArea>();
        var firstSeries = plotArea is null ? null : SeriesOf(plotArea).FirstOrDefault();

        return new
        {
            dataLabels = DescribeDataLabels(firstSeries),
            legend = DescribeLegend(chart),
            axisTitles = DescribeAxisTitles(plotArea),
            trendline = DescribeTrendline(firstSeries),
            errorBars = DescribeErrorBars(firstSeries),
            gridlines = DescribeGridlines(plotArea),
            secondaryAxis = DescribeSecondaryAxis(plotArea),
        };
    }

    private static object? DescribeDataLabels(OpenXmlCompositeElement? series)
    {
        if (series?.GetFirstChild<C.DataLabels>() is not { } dLbls)
        {
            return null;
        }

        string? what = null;
        if (dLbls.GetFirstChild<C.ShowValue>()?.Val?.Value == true)
        {
            what = "value";
        }
        else if (dLbls.GetFirstChild<C.ShowPercent>()?.Val?.Value == true)
        {
            what = "percent";
        }
        else if (dLbls.GetFirstChild<C.ShowCategoryName>()?.Val?.Value == true)
        {
            what = "category";
        }
        else if (dLbls.GetFirstChild<C.ShowSeriesName>()?.Val?.Value == true)
        {
            what = "seriesName";
        }

        var position = dLbls.GetFirstChild<C.DataLabelPosition>()?.Val?.Value is { } pos
            ? DataLabelPositionName(pos)
            : null;
        return new { show = what, position };
    }

    private static string? DescribeLegend(C.Chart chart) =>
        chart.GetFirstChild<C.Legend>()?.GetFirstChild<C.LegendPosition>()?.Val?.Value is { } pos
            ? LegendPositionName(pos)
            : "none";

    private static object? DescribeAxisTitles(C.PlotArea? plotArea)
    {
        if (plotArea is null)
        {
            return null;
        }

        var category = PrimaryCategoryAxis(plotArea) is { } catAx ? AxisTitleText(catAx) : null;
        var value = PrimaryValueAxis(plotArea) is { } valAx ? AxisTitleText(valAx) : null;
        if (category is null && value is null)
        {
            return null;
        }

        return new { category, value };
    }

    private static string? DescribeTrendline(OpenXmlCompositeElement? series) =>
        series?.GetFirstChild<C.Trendline>()?.GetFirstChild<C.TrendlineType>()?.Val?.Value is { } type
            ? TrendlineName(type)
            : null;

    private static string? DescribeErrorBars(OpenXmlCompositeElement? series) =>
        series?.GetFirstChild<C.ErrorBars>()?.GetFirstChild<C.ErrorBarValueType>()?.Val?.Value is { } valueType
            ? ErrorBarName(valueType)
            : null;

    private static object? DescribeGridlines(C.PlotArea? plotArea)
    {
        if (plotArea is null || PrimaryValueAxis(plotArea) is not { } axis)
        {
            return null;
        }

        var major = axis.GetFirstChild<C.MajorGridlines>() is not null;
        var minor = axis.GetFirstChild<C.MinorGridlines>() is not null;
        if (!major && !minor)
        {
            return null;
        }

        return new { major, minor };
    }

    private static IReadOnlyList<string>? DescribeSecondaryAxis(C.PlotArea? plotArea)
    {
        if (plotArea is null)
        {
            return null;
        }

        var secondaryGroup = ChartGroups(plotArea)
            .FirstOrDefault(g => g.ChildElements.OfType<C.AxisId>().Any(a => a.Val?.Value == SecondaryValueAxisId));
        if (secondaryGroup is null)
        {
            return null;
        }

        var names = secondaryGroup.ChildElements
            .Where(e => e.LocalName == "ser")
            .Cast<OpenXmlCompositeElement>()
            .Select(SeriesName)
            .Where(n => n.Length > 0)
            .ToList();
        return names.Count > 0 ? names : null;
    }

    private static string AxisTitleText(OpenXmlCompositeElement axis)
    {
        var title = axis.GetFirstChild<C.Title>();
        return title is null ? string.Empty : string.Concat(title.Descendants<A.Text>().Select(t => t.Text));
    }

    /// <summary>The cached series name from its <c>c:tx</c> (string cache or literal).</summary>
    private static string SeriesName(OpenXmlElement series)
    {
        var tx = series.ChildElements.FirstOrDefault(e => e.LocalName == "tx");
        if (tx is null)
        {
            return string.Empty;
        }

        var cachePoint = tx.Descendants<C.StringPoint>().FirstOrDefault()?.NumericValue?.Text;
        if (cachePoint is not null)
        {
            return cachePoint;
        }

        return tx.Descendants<C.NumericValue>().FirstOrDefault()?.Text ?? string.Empty;
    }

    /// <summary>The cached series names on a chart, for the secondary-axis name check.</summary>
    public static IReadOnlyList<string> SeriesNames(C.Chart chart)
    {
        var plotArea = chart.GetFirstChild<C.PlotArea>();
        return plotArea is null
            ? []
            : SeriesOf(plotArea).Select(SeriesName).Where(n => n.Length > 0).ToList();
    }

    // ----- enum mapping -------------------------------------------------------

    private static C.LegendPositionValues LegendPositionValue(string legend) => legend switch
    {
        "right" => C.LegendPositionValues.Right,
        "left" => C.LegendPositionValues.Left,
        "top" => C.LegendPositionValues.Top,
        "bottom" => C.LegendPositionValues.Bottom,
        _ => C.LegendPositionValues.Right,
    };

    private static string LegendPositionName(C.LegendPositionValues value)
    {
        if (value == C.LegendPositionValues.Right) return "right";
        if (value == C.LegendPositionValues.Left) return "left";
        if (value == C.LegendPositionValues.Top) return "top";
        if (value == C.LegendPositionValues.Bottom) return "bottom";
        if (value == C.LegendPositionValues.TopRight) return "right";
        return "right";
    }

    private static C.DataLabelPositionValues DataLabelPositionValue(string position) => position switch
    {
        "outEnd" => C.DataLabelPositionValues.OutsideEnd,
        "inEnd" => C.DataLabelPositionValues.InsideEnd,
        "center" => C.DataLabelPositionValues.Center,
        "bestFit" => C.DataLabelPositionValues.BestFit,
        _ => C.DataLabelPositionValues.OutsideEnd,
    };

    private static string DataLabelPositionName(C.DataLabelPositionValues value)
    {
        if (value == C.DataLabelPositionValues.OutsideEnd) return "outEnd";
        if (value == C.DataLabelPositionValues.InsideEnd) return "inEnd";
        if (value == C.DataLabelPositionValues.Center) return "center";
        if (value == C.DataLabelPositionValues.BestFit) return "bestFit";
        return "outEnd";
    }

    private static C.TrendlineValues TrendlineValue(string trendline) => trendline switch
    {
        "linear" => C.TrendlineValues.Linear,
        "exponential" => C.TrendlineValues.Exponential,
        "movingAverage" => C.TrendlineValues.MovingAverage,
        _ => C.TrendlineValues.Linear,
    };

    private static string TrendlineName(C.TrendlineValues value)
    {
        if (value == C.TrendlineValues.Linear) return "linear";
        if (value == C.TrendlineValues.Exponential) return "exponential";
        if (value == C.TrendlineValues.MovingAverage) return "movingAverage";
        return "linear";
    }

    private static C.ErrorValues ErrorValue(string kind) => kind switch
    {
        "stdErr" => C.ErrorValues.StandardError,
        "stdDev" => C.ErrorValues.StandardDeviation,
        "percent" => C.ErrorValues.Percentage,
        _ => C.ErrorValues.StandardError,
    };

    private static string ErrorBarName(C.ErrorValues value)
    {
        if (value == C.ErrorValues.StandardError) return "stdErr";
        if (value == C.ErrorValues.StandardDeviation) return "stdDev";
        if (value == C.ErrorValues.Percentage) return "percent";
        return "stdErr";
    }

    // ----- small parse helpers -----------------------------------------------

    private static string? StringField(JsonObject obj, string key, int opIndex, string owner)
    {
        if (!obj.TryGetPropertyValue(key, out var node) || node is null)
        {
            return null;
        }

        if (node is JsonValue value && value.GetValueKind() == JsonValueKind.String)
        {
            return value.GetValue<string>();
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: '{owner}.{key}' must be a string.",
            $"Pass {owner}.{key} as text, e.g. \"value\".");
    }

    private static bool? BoolField(JsonObject obj, string key, int opIndex)
    {
        if (!obj.TryGetPropertyValue(key, out var node) || node is null)
        {
            return null;
        }

        if (node is JsonValue value && value.GetValueKind() is JsonValueKind.True or JsonValueKind.False)
        {
            return value.GetValue<bool>();
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: '{key}' must be true or false.",
            $"Pass e.g. {{\"{key}\":true}}.");
    }

    private static AiofficeException Unsupported(int opIndex, string what, IReadOnlyList<string> supported) =>
        new(
            ErrorCodes.UnsupportedFeature,
            $"ops[{opIndex}]: {what} is not supported.",
            "Supported values: " + string.Join(", ", supported) + ".",
            candidates: supported);
}
