using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>One chart series: display name plus literal-cached values (null = gap).</summary>
internal sealed record PptxChartSeries(string Name, IReadOnlyList<double?> Values);

/// <summary>A slide chart's data, as written into (or read back from) its ChartPart.</summary>
internal sealed record PptxChartData(
    string Kind,
    string? Title,
    IReadOnlyList<string> Categories,
    IReadOnlyList<PptxChartSeries> Series);

/// <summary>
/// Native pptx charts: a p:graphicFrame in the slide's shape tree referencing a
/// ChartPart whose series carry literal data caches (c:strLit / c:numLit) — no
/// embedded workbook. PowerPoint renders the cached data as-is; "Edit Data"
/// prompts to create a workbook, which is why every projection reports
/// <c>dataEditable: false</c> and chart creation attaches an envelope warning.
/// </summary>
internal static class PptxCharts
{
    /// <summary>The chart kinds aioffice can create. Everything else is unsupported_feature.</summary>
    public static readonly IReadOnlyList<string> Kinds = ["bar", "line", "pie"];

    private static readonly IReadOnlyList<string> AddProps =
        ["kind", "categories", "series", "title", "x", "y", "w", "h"];

    private const string ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string MainNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const uint CategoryAxisId = 100001u;
    private const uint ValueAxisId = 100002u;

    /// <summary>The warning attached to every edit envelope that created a chart.</summary>
    public static readonly Warning DataNotEditableWarning = new(
        "chart_data_not_editable",
        "Chart data is cached literally in the chart XML (no embedded workbook). PowerPoint renders it " +
        "as-is, but 'Edit Data' will prompt to create a new workbook; to change values, re-add the chart.");

    // ----- create --------------------------------------------------------------

    /// <summary>Adds a chart graphic frame to the slide and returns its stable shape id.</summary>
    public static uint Add(SlidePart slidePart, JsonObject? props)
    {
        props ??= [];
        var data = ParseAdd(props);
        var tree = PptxDoc.RequireShapeTree(slidePart);
        var id = PptxDoc.NextShapeId(tree);

        var x = Length(props, "x", Units.CmToEmu(2));
        var y = Length(props, "y", Units.CmToEmu(3));
        var w = Length(props, "w", Units.CmToEmu(20));
        var h = Length(props, "h", Units.CmToEmu(12));

        var chartPart = slidePart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = BuildChartSpace(data);

        tree.Append(new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = Units.Inv($"Chart {id}") },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = x, Y = y },
                new A.Extents { Cx = w, Cy = h }),
            new A.Graphic(new A.GraphicData(
                new C.ChartReference { Id = slidePart.GetIdOfPart(chartPart) })
            {
                Uri = ChartNs,
            })));
        return id;
    }

    private static long Length(JsonObject props, string key, long fallback) =>
        props.TryGetPropertyValue(key, out var node) ? Units.ParseLengthEmu(key, node) : fallback;

    /// <summary>Validates an add-chart op's props and extracts the chart data.</summary>
    private static PptxChartData ParseAdd(JsonObject props)
    {
        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown chart prop '{key}'.",
                    "Chart props: kind, categories, series, title, x, y, w, h.",
                    candidates: AddProps);
            }
        }

        var kind = props.TryGetPropertyValue("kind", out var kindNode) && kindNode is not null
            ? J.ScalarText(kindNode).Trim().ToLowerInvariant()
            : throw MissingProp("kind", "\"bar\"");
        if (!Kinds.Contains(kind, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Chart kind '{kind}' is not supported.",
                "Supported chart kinds: bar, line, pie. A line chart is the usual stand-in for scatter/area data.",
                candidates: Kinds);
        }

        var categories = ParseCategories(props);
        var series = ParseSeries(props, categories.Count);
        if (kind == "pie" && series.Count != 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"A pie chart takes exactly one series; props.series has {series.Count}.",
                "Pass a single series, or use kind bar/line for multi-series data.");
        }

        var title = props.TryGetPropertyValue("title", out var titleNode) && titleNode is not null
            ? J.ScalarText(titleNode)
            : null;

        return new PptxChartData(kind, title, categories, series);
    }

    private static List<string> ParseCategories(JsonObject props)
    {
        if (!props.TryGetPropertyValue("categories", out var node) || node is not JsonArray array || array.Count == 0)
        {
            throw MissingProp("categories", "[\"Q1\",\"Q2\",\"Q3\",\"Q4\"]");
        }

        var categories = new List<string>(array.Count);
        foreach (var item in array)
        {
            if (item is null or JsonObject or JsonArray)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "props.categories must hold scalar labels.",
                    "Use strings or numbers, e.g. {\"categories\":[\"Q1\",\"Q2\"]}.");
            }

            categories.Add(J.ScalarText(item));
        }

        return categories;
    }

    private static List<PptxChartSeries> ParseSeries(JsonObject props, int categoryCount)
    {
        if (!props.TryGetPropertyValue("series", out var node) || node is not JsonArray array || array.Count == 0)
        {
            throw MissingProp("series", "[{\"name\":\"Sales\",\"values\":[10,20,30,40]}]");
        }

        var series = new List<PptxChartSeries>(array.Count);
        for (var i = 0; i < array.Count; i++)
        {
            if (array[i] is not JsonObject item)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    Units.Inv($"props.series[{i}] is not an object."),
                    "Each series looks like {\"name\":\"Sales\",\"values\":[10,20,30]}.");
            }

            if (!item.TryGetPropertyValue("values", out var valuesNode) || valuesNode is not JsonArray values)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    Units.Inv($"props.series[{i}] has no values array."),
                    "Each series needs numeric values, e.g. {\"name\":\"Sales\",\"values\":[10,20,30]}.");
            }

            if (values.Count != categoryCount)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    Units.Inv($"props.series[{i}] has {values.Count} value(s) but there are {categoryCount} categories."),
                    "Give every series exactly one value per category (null for a gap).");
            }

            var numbers = new List<double?>(values.Count);
            for (var v = 0; v < values.Count; v++)
            {
                numbers.Add(NumericValue(values[v], i, v));
            }

            var name = item.TryGetPropertyValue("name", out var nameNode) && nameNode is not null
                ? J.ScalarText(nameNode)
                : Units.Inv($"Series {i + 1}");
            series.Add(new PptxChartSeries(name, numbers));
        }

        return series;
    }

    private static double? NumericValue(JsonNode? node, int seriesIndex, int valueIndex)
    {
        if (node is null)
        {
            return null; // an explicit gap
        }

        if (node is JsonValue value)
        {
            if (Units.TryNumber(value, out var number))
            {
                return number;
            }

            if (value.TryGetValue<string>(out var text) &&
                double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            Units.Inv($"props.series[{seriesIndex}].values[{valueIndex}] is not a number: {node.ToJsonString()}"),
            "Series values must be numbers (or null for a gap).");
    }

    private static AiofficeException MissingProp(string key, string example) => new(
        ErrorCodes.InvalidArgs,
        $"add chart needs the '{key}' prop (e.g. {example}).",
        "The full shape is {\"op\":\"add\",\"path\":\"/slide[2]\",\"type\":\"chart\",\"props\":{\"kind\":\"bar\"," +
        "\"categories\":[\"Q1\",\"Q2\"],\"series\":[{\"name\":\"Sales\",\"values\":[10,20]}],\"title\":\"Sales\"}}.");

    // ----- chartml -------------------------------------------------------------

    private static C.ChartSpace BuildChartSpace(PptxChartData data)
    {
        var plotArea = new C.PlotArea(new C.Layout());
        plotArea.Append(BuildChartGroup(data));
        if (data.Kind is "bar" or "line")
        {
            plotArea.Append(new C.CategoryAxis(
                new C.AxisId { Val = CategoryAxisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
                new C.CrossingAxis { Val = ValueAxisId }));
            plotArea.Append(new C.ValueAxis(
                new C.AxisId { Val = ValueAxisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = C.AxisPositionValues.Left },
                new C.CrossingAxis { Val = CategoryAxisId }));
        }

        var chart = new C.Chart();
        if (data.Title is { } titleText)
        {
            chart.Append(new C.Title(
                new C.ChartText(new C.RichText(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(titleText))))),
                new C.Overlay { Val = false }));
        }

        chart.Append(new C.AutoTitleDeleted { Val = data.Title is null });
        chart.Append(plotArea);
        chart.Append(new C.PlotVisibleOnly { Val = true });

        var chartSpace = new C.ChartSpace();
        chartSpace.AddNamespaceDeclaration("c", ChartNs);
        chartSpace.AddNamespaceDeclaration("a", MainNs);
        chartSpace.AddNamespaceDeclaration("r", RelNs);
        chartSpace.Append(chart);
        return chartSpace;
    }

    private static OpenXmlCompositeElement BuildChartGroup(PptxChartData data)
    {
        switch (data.Kind)
        {
            case "bar":
            {
                var group = new C.BarChart(
                    new C.BarDirection { Val = C.BarDirectionValues.Column },
                    new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                    new C.VaryColors { Val = false });
                for (var i = 0; i < data.Series.Count; i++)
                {
                    var series = new C.BarChartSeries();
                    AppendSeriesChildren(series, data, i);
                    group.Append(series);
                }

                group.Append(new C.AxisId { Val = CategoryAxisId });
                group.Append(new C.AxisId { Val = ValueAxisId });
                return group;
            }

            case "line":
            {
                var group = new C.LineChart(
                    new C.Grouping { Val = C.GroupingValues.Standard },
                    new C.VaryColors { Val = false });
                for (var i = 0; i < data.Series.Count; i++)
                {
                    var series = new C.LineChartSeries();
                    AppendSeriesChildren(series, data, i);
                    group.Append(series);
                }

                group.Append(new C.AxisId { Val = CategoryAxisId });
                group.Append(new C.AxisId { Val = ValueAxisId });
                return group;
            }

            default: // "pie" — ParseAdd rejected everything else
            {
                var group = new C.PieChart(new C.VaryColors { Val = true });
                var series = new C.PieChartSeries();
                AppendSeriesChildren(series, data, 0);
                group.Append(series);
                group.Append(new C.FirstSliceAngle { Val = 0 });
                return group;
            }
        }
    }

    /// <summary>idx/order/tx/cat/val with literal caches — the no-workbook series skeleton.</summary>
    private static void AppendSeriesChildren(OpenXmlCompositeElement series, PptxChartData data, int ordinal)
    {
        var spec = data.Series[ordinal];
        series.Append(new C.Index { Val = (uint)ordinal });
        series.Append(new C.Order { Val = (uint)ordinal });
        series.Append(new C.SeriesText(new C.NumericValue(spec.Name)));

        var stringLiteral = new C.StringLiteral(new C.PointCount { Val = (uint)data.Categories.Count });
        for (var i = 0; i < data.Categories.Count; i++)
        {
            stringLiteral.Append(new C.StringPoint
            {
                Index = (uint)i,
                NumericValue = new C.NumericValue(data.Categories[i]),
            });
        }

        series.Append(new C.CategoryAxisData(stringLiteral));

        var numberLiteral = new C.NumberLiteral(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)spec.Values.Count });
        for (var i = 0; i < spec.Values.Count; i++)
        {
            if (spec.Values[i] is { } value)
            {
                numberLiteral.Append(new C.NumericPoint
                {
                    Index = (uint)i,
                    NumericValue = new C.NumericValue(value.ToString(CultureInfo.InvariantCulture)),
                });
            }
        }

        series.Append(new C.Values(numberLiteral));
    }

    // ----- enumerate / resolve -------------------------------------------------

    /// <summary>The chart part a graphic frame references; null when the element hosts no chart.</summary>
    public static ChartPart? ChartPartOf(SlidePart slidePart, OpenXmlCompositeElement element)
    {
        if (element is not P.GraphicFrame frame ||
            frame.Descendants<C.ChartReference>().FirstOrDefault()?.Id?.Value is not { } relId)
        {
            return null;
        }

        return slidePart.TryGetPartById(relId, out var part) && part is ChartPart chartPart ? chartPart : null;
    }

    /// <summary>Chart-hosting frames on a slide in paint order; chart indices are 1-based.</summary>
    public static List<(int Index, ShapeView View, ChartPart Part)> Charts(SlidePart slidePart)
    {
        var result = new List<(int, ShapeView, ChartPart)>();
        foreach (var view in PptxDoc.Shapes(slidePart))
        {
            if (ChartPartOf(slidePart, view.Element) is { } part)
            {
                result.Add((result.Count + 1, view, part));
            }
        }

        return result;
    }

    /// <summary>Resolves /slide[i]/chart[k] or throws invalid_path with candidates.</summary>
    public static (int Index, ShapeView View, ChartPart Part) Resolve(SlidePart slidePart, PptxAddress address)
    {
        var charts = Charts(slidePart);
        var index = address.ChartIndex!.Value;
        if (index >= 1 && index <= charts.Count)
        {
            return charts[index - 1];
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            Units.Inv($"No chart {index} on slide {address.SlideIndex}; it has {charts.Count} chart(s)."),
            charts.Count > 0
                ? "Chart indices are 1-based per slide; run 'aioffice get <file> /slide[i]' to list shapes."
                : "Add one first: {\"op\":\"add\",\"path\":\"" + address.CanonicalSlidePath + "\",\"type\":\"chart\"," +
                  "\"props\":{\"kind\":\"bar\",\"categories\":[\"Q1\"],\"series\":[{\"name\":\"Sales\",\"values\":[10]}]}}.",
            candidates: [.. charts.Take(10).Select(c => Units.Inv($"{address.CanonicalSlidePath}/chart[{c.Index}]"))]);
    }

    /// <summary>The 1-based chart index of a chart-hosting frame on its slide; null for non-charts.</summary>
    public static int? IndexOf(SlidePart slidePart, OpenXmlCompositeElement element) =>
        Charts(slidePart).Where(c => ReferenceEquals(c.View.Element, element))
            .Select(c => (int?)c.Index)
            .FirstOrDefault();

    // ----- read-back -----------------------------------------------------------

    /// <summary>Reads the cached chart data back from a chart part (literal or ref caches alike).</summary>
    public static PptxChartData ReadData(ChartPart part)
    {
        var chartSpace = part.ChartSpace;
        var plotArea = chartSpace?.Descendants<C.PlotArea>().FirstOrDefault();
        var group = plotArea?.ChildElements
            .FirstOrDefault(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));

        var kind = KindName(group?.LocalName);
        var title = TitleText(chartSpace);
        var serElements = group?.ChildElements.Where(e => e.LocalName == "ser").ToList() ?? [];

        var categories = new List<string>();
        if (serElements.Count > 0 &&
            serElements[0].ChildElements.FirstOrDefault(e => e.LocalName == "cat") is { } cat)
        {
            categories = ReadStringPoints(cat);
        }

        var series = new List<PptxChartSeries>(serElements.Count);
        foreach (var ser in serElements)
        {
            var name = ser.ChildElements.FirstOrDefault(e => e.LocalName == "tx")?
                .Descendants<C.NumericValue>().FirstOrDefault()?.Text ?? string.Empty;
            var val = ser.ChildElements.FirstOrDefault(e => e.LocalName == "val");
            series.Add(new PptxChartSeries(name, val is null ? [] : ReadNumericPoints(val)));
        }

        return new PptxChartData(kind, title, categories, series);
    }

    private static string KindName(string? localName)
    {
        if (string.IsNullOrEmpty(localName))
        {
            return "unknown";
        }

        var name = localName.EndsWith("Chart", StringComparison.Ordinal) ? localName[..^5] : localName;
        return name.EndsWith("3D", StringComparison.Ordinal) ? name[..^2] : name;
    }

    private static string? TitleText(C.ChartSpace? chartSpace)
    {
        var title = chartSpace?.Descendants<C.Title>().FirstOrDefault();
        if (title is null)
        {
            return null;
        }

        var text = string.Concat(title.Descendants<A.Text>().Select(t => t.Text));
        return text.Length == 0 ? null : text;
    }

    private static List<string> ReadStringPoints(OpenXmlElement container)
    {
        var points = container.Descendants<C.StringPoint>().ToList();
        var count = (int)(container.Descendants<C.PointCount>().FirstOrDefault()?.Val?.Value ?? (uint)points.Count);
        var result = new List<string>(count);
        for (var i = 0; i < count; i++)
        {
            result.Add(string.Empty);
        }

        foreach (var point in points)
        {
            var index = (int)(point.Index?.Value ?? 0);
            if (index >= 0 && index < count)
            {
                result[index] = point.NumericValue?.Text ?? string.Empty;
            }
        }

        return result;
    }

    private static List<double?> ReadNumericPoints(OpenXmlElement container)
    {
        var points = container.Descendants<C.NumericPoint>().ToList();
        var count = (int)(container.Descendants<C.PointCount>().FirstOrDefault()?.Val?.Value ?? (uint)points.Count);
        var result = new List<double?>(count);
        for (var i = 0; i < count; i++)
        {
            result.Add(null);
        }

        foreach (var point in points)
        {
            var index = (int)(point.Index?.Value ?? 0);
            if (index >= 0 && index < count &&
                double.TryParse(point.NumericValue?.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            {
                result[index] = value;
            }
        }

        return result;
    }

    // ----- projections ---------------------------------------------------------

    /// <summary>The `get` projection for /slide[i]/chart[k].</summary>
    public static object Detail(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var (index, view, part) = Resolve(slidePart, address);
        var data = ReadData(part);
        var geometry = PptxDoc.Geometry(view.Element);
        return new
        {
            Path = Units.Inv($"/slide[{address.SlideIndex}]/chart[{index}]"),
            ShapePath = view.CanonicalPath(address.SlideIndex),
            Slide = address.SlideIndex,
            Id = view.Id,
            Kind = "chart",
            ChartKind = data.Kind,
            Title = data.Title,
            Categories = data.Categories,
            Series = data.Series.Select(s => (object)new { Name = s.Name, Values = s.Values }).ToList(),
            DataEditable = false,
            X = geometry is { } g1 ? Units.EmuToCm(g1.X) : (double?)null,
            Y = geometry is { } g2 ? Units.EmuToCm(g2.Y) : (double?)null,
            W = geometry is { } g3 ? Units.EmuToCm(g3.Cx) : (double?)null,
            H = geometry is { } g4 ? Units.EmuToCm(g4.Cy) : (double?)null,
            ZIndex = view.Ordinal,
        };
    }

    /// <summary>The chart summary embedded in a shape `get` when the shape hosts a chart.</summary>
    public static object? Summary(SlidePart slidePart, OpenXmlCompositeElement element)
    {
        if (ChartPartOf(slidePart, element) is not { } part)
        {
            return null;
        }

        var data = ReadData(part);
        return new
        {
            Kind = data.Kind,
            Title = data.Title,
            Categories = data.Categories,
            SeriesNames = data.Series.Select(s => s.Name).ToList(),
            DataEditable = false,
        };
    }

    // ----- remove --------------------------------------------------------------

    /// <summary>remove /slide[i]/chart[k]: drops the graphic frame and deletes its chart part.</summary>
    public static string Remove(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var (index, view, part) = Resolve(slidePart, address);
        view.Element.Remove();
        slidePart.DeletePart(part);
        return Units.Inv($"/slide[{address.SlideIndex}]/chart[{index}]");
    }

    /// <summary>Deletes the chart part when a shape being removed hosts one (shape-path removes).</summary>
    public static void DeletePartFor(SlidePart slidePart, OpenXmlCompositeElement element)
    {
        if (ChartPartOf(slidePart, element) is { } part)
        {
            slidePart.DeletePart(part);
        }
    }
}
