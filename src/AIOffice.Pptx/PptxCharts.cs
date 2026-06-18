using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// One chart series: display name plus literal-cached values (null = gap). For a
/// bubble chart the parallel <see cref="XValues"/>/<see cref="Sizes"/> carry the
/// x/y/size triple per point (<see cref="Values"/> holds the y values); they are
/// empty for every other chart kind.
/// </summary>
internal sealed record PptxChartSeries(
    string Name,
    IReadOnlyList<double?> Values,
    IReadOnlyList<double?>? XValues = null,
    IReadOnlyList<double?>? Sizes = null);

/// <summary>A slide chart's data, as written into (or read back from) its ChartPart.</summary>
internal sealed record PptxChartData(
    string Kind,
    string? Title,
    IReadOnlyList<string> Categories,
    IReadOnlyList<PptxChartSeries> Series);

/// <summary>
/// Native pptx charts: a p:graphicFrame in the slide's shape tree referencing a
/// ChartPart whose series carry reference caches (c:strRef / c:numRef) backed by
/// a minimal real workbook embedded in the chart part (c:externalData with
/// autoUpdate=false) — so PowerPoint's "Edit Data" opens a live sheet. Charts
/// created by aioffice embed the workbook from the start; foreign cached-only
/// charts report <c>dataEditable: false</c> until retrofitted via
/// <c>{"op":"set","path":"/slide[i]/chart[k]","props":{"embedData":true}}</c>.
/// </summary>
internal static class PptxCharts
{
    /// <summary>The chart kinds aioffice can create. Everything else is unsupported_feature.</summary>
    public static readonly IReadOnlyList<string> Kinds =
    [
        "bar", "line", "pie",
        "doughnut", "radar", "bubble",
        "stackedBar", "percentStackedBar", "stackedArea", "combo",
    ];

    /// <summary>Kinds that draw on a category + value axis pair (so the plot area carries both axes).</summary>
    private static readonly IReadOnlyList<string> CategoryValueAxisKinds =
        ["bar", "line", "stackedBar", "percentStackedBar", "stackedArea", "combo"];

    private static readonly IReadOnlyList<string> AddProps =
        ["kind", "categories", "series", "title", "x", "y", "w", "h", .. PptxChartPolish.PropKeys];

    private const string ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string MainNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const uint CategoryAxisId = 100001u;
    private const uint ValueAxisId = 100002u;

    // ----- create --------------------------------------------------------------

    /// <summary>Adds a chart graphic frame (with an embedded data workbook) and returns its stable shape id.</summary>
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
        chartPart.ChartSpace = BuildChartSpace(data, EmbedWorkbook(chartPart, data));

        // Chart-polish props (dataLabels/legend/axisTitles/trendline/errorBars/
        // gridlines/secondaryAxis) are accepted at create alongside the data.
        var polish = PptxChartPolish.Split(props, out _);
        if (polish.Count > 0)
        {
            PptxChartPolish.Apply(chartPart.ChartSpace, polish);
        }

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
            ? NormalizeKind(J.ScalarText(kindNode).Trim())
            : throw MissingProp("kind", "\"bar\"");
        if (!Kinds.Contains(kind, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Chart kind '{kind}' is not supported.",
                "Supported chart kinds: " + string.Join(", ", Kinds) +
                ". A line chart is the usual stand-in for scatter/area data; bubble needs an x/y/size triple per point.",
                candidates: Kinds);
        }

        var categories = ParseCategories(props);

        // Bubble plots an x/y/size triple per point, so its series carry three
        // value arrays rather than the single value-per-category shape.
        if (kind == "bubble")
        {
            var bubbleSeries = ParseBubbleSeries(props, categories.Count);
            return new PptxChartData(kind, ReadTitle(props), categories, bubbleSeries);
        }

        var series = ParseSeries(props, categories.Count);
        if (kind is "pie" or "doughnut" && series.Count != 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"A {kind} chart takes exactly one series; props.series has {series.Count}.",
                "Pass a single series, or use kind bar/line for multi-series data.");
        }

        if (kind == "combo" && series.Count < 2)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"A combo chart needs at least two series (column(s) plus a line); props.series has {series.Count}.",
                "Give it two or more series — the first becomes columns, the rest a line — or use kind bar/line.");
        }

        return new PptxChartData(kind, ReadTitle(props), categories, series);
    }

    /// <summary>Reads the optional chart title; null when no title prop is present.</summary>
    private static string? ReadTitle(JsonObject props) =>
        props.TryGetPropertyValue("title", out var titleNode) && titleNode is not null
            ? J.ScalarText(titleNode)
            : null;

    /// <summary>
    /// Canonicalizes the kind token: lower-cases the simple kinds but keeps the
    /// camelCase compound kinds (stackedBar, percentStackedBar, stackedArea) as
    /// the contract spells them, accepting any case the caller sends.
    /// </summary>
    private static string NormalizeKind(string raw)
    {
        foreach (var kind in Kinds)
        {
            if (string.Equals(kind, raw, StringComparison.OrdinalIgnoreCase))
            {
                return kind;
            }
        }

        return raw.ToLowerInvariant();
    }

    /// <summary>
    /// Parses a bubble chart's series: each point is an {x, y, size} triple (or
    /// the bare [y] arrays for x/size, matching the categories count). x defaults
    /// to the 1-based point index and size to 1 when omitted, so a caller can pass
    /// the same {name, values} shape as the other kinds and still get a valid chart.
    /// </summary>
    private static List<PptxChartSeries> ParseBubbleSeries(JsonObject props, int categoryCount)
    {
        if (!props.TryGetPropertyValue("series", out var node) || node is not JsonArray array || array.Count == 0)
        {
            throw MissingProp("series", "[{\"name\":\"Deals\",\"values\":[10,20,30],\"x\":[1,2,3],\"size\":[5,8,3]}]");
        }

        var series = new List<PptxChartSeries>(array.Count);
        for (var i = 0; i < array.Count; i++)
        {
            if (array[i] is not JsonObject item)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    Units.Inv($"props.series[{i}] is not an object."),
                    "Each bubble series looks like {\"name\":\"Deals\",\"values\":[10,20],\"x\":[1,2],\"size\":[5,8]}.");
            }

            var yValues = ParseValueArray(item, "values", i, categoryCount, required: true)!;
            var xValues = ParseValueArray(item, "x", i, categoryCount, required: false)
                ?? [.. Enumerable.Range(1, yValues.Count).Select(n => (double?)n)];
            var sizes = ParseValueArray(item, "size", i, categoryCount, required: false)
                ?? [.. Enumerable.Repeat((double?)1, yValues.Count)];

            var name = item.TryGetPropertyValue("name", out var nameNode) && nameNode is not null
                ? J.ScalarText(nameNode)
                : Units.Inv($"Series {i + 1}");
            series.Add(new PptxChartSeries(name, yValues, xValues, sizes));
        }

        return series;
    }

    /// <summary>Parses one numeric array prop of a bubble series (length must match the category count).</summary>
    private static List<double?>? ParseValueArray(JsonObject item, string key, int seriesIndex, int categoryCount, bool required)
    {
        if (!item.TryGetPropertyValue(key, out var node) || node is null)
        {
            if (!required)
            {
                return null;
            }

            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"props.series[{seriesIndex}] has no '{key}' array."),
                "A bubble series needs at least 'values' (the y values); x and size are optional.");
        }

        if (node is not JsonArray values)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"props.series[{seriesIndex}].{key} is not an array."),
                "Pass numbers (or null for a gap), e.g. {\"" + key + "\":[10,20,30]}.");
        }

        if (values.Count != categoryCount)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"props.series[{seriesIndex}].{key} has {values.Count} value(s) but there are {categoryCount} categories."),
                "Every bubble value array (values, x, size) must hold exactly one entry per category.");
        }

        var numbers = new List<double?>(values.Count);
        for (var v = 0; v < values.Count; v++)
        {
            numbers.Add(NumericValue(values[v], seriesIndex, v));
        }

        return numbers;
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

    // ----- embedded workbook ----------------------------------------------------

    /// <summary>Embeds the data workbook in the chart part and returns its relationship id.</summary>
    private static string EmbedWorkbook(ChartPart chartPart, PptxChartData data)
    {
        var embedded = chartPart.AddNewPart<EmbeddedPackagePart>(PptxChartWorkbook.ContentType);
        using (var workbook = new MemoryStream(PptxChartWorkbook.Build(data)))
        {
            embedded.FeedData(workbook);
        }

        return chartPart.GetIdOfPart(embedded);
    }

    /// <summary>True when the chart references an embedded workbook (PowerPoint "Edit Data" works).</summary>
    public static bool DataEditable(ChartPart part) =>
        part.ChartSpace?.Elements<C.ExternalData>().FirstOrDefault()?.Id?.Value is { } relId &&
        part.TryGetPartById(relId, out var referenced) &&
        referenced is EmbeddedPackagePart;

    /// <summary>
    /// set /slide[i]/chart[k] {embedData:true}: retrofits a cached-only chart with
    /// an embedded workbook and rewrites its series to reference caches. Already
    /// embedded charts are rebuilt from their current caches (idempotent).
    /// </summary>
    public static string EmbedData(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var (index, _, part) = Resolve(slidePart, address);
        var chartSpace = part.ChartSpace ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The chart part has no chartSpace XML.",
            "Remove the chart and add it again, or restore a snapshot.");

        var serElements = chartSpace.Descendants<C.PlotArea>().FirstOrDefault()?.ChildElements
            .FirstOrDefault(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal))?.ChildElements
            .Where(e => e.LocalName == "ser")
            .Cast<OpenXmlCompositeElement>()
            .ToList() ?? [];
        if (serElements.Count == 0 || serElements.Any(ser => ser.ChildElements.All(c => c.LocalName != "val")))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "This chart's series carry no category/value caches aioffice can embed.",
                "Only bar/line/pie-style charts (c:cat/c:val series) can be retrofitted; " +
                "remove the chart and re-add it with explicit data instead.");
        }

        var data = ReadData(part);

        // Idempotent rebuild: drop any previous embedded workbook and wiring first.
        foreach (var external in chartSpace.Elements<C.ExternalData>().ToList())
        {
            external.Remove();
        }

        foreach (var embedded in part.Parts.Select(p => p.OpenXmlPart).OfType<EmbeddedPackagePart>().ToList())
        {
            part.DeletePart(embedded);
        }

        chartSpace.Append(new C.ExternalData(new C.AutoUpdate { Val = false }) { Id = EmbedWorkbook(part, data) });

        for (var i = 0; i < serElements.Count; i++)
        {
            var ser = serElements[i];
            ReplaceSerChild(ser, "tx", BuildSeriesText(data.Series[i].Name, i), insertAfter: ["order", "idx"]);
            if (data.Categories.Count > 0)
            {
                ReplaceSerChild(ser, "cat", BuildCategoryData(data.Categories), insertBefore: "val");
            }

            ReplaceSerChild(ser, "val", BuildValues(data.Series[i].Values, i), insertAfter: ["cat", "order", "idx"]);
        }

        return Units.Inv($"/slide[{address.SlideIndex}]/chart[{index}]");
    }

    /// <summary>
    /// set /slide[i]/chart[k] with chart-polish props: applies dataLabels/legend/
    /// axisTitles/trendline/errorBars/gridlines/secondaryAxis to the existing chart
    /// in place. Returns the canonical chart path.
    /// </summary>
    public static string SetPolish(PresentationPart presentation, PptxAddress address, JsonObject polish)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var (index, _, part) = Resolve(slidePart, address);
        var chartSpace = part.ChartSpace ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The chart part has no chartSpace XML.",
            "Remove the chart and add it again, or restore a snapshot.");

        PptxChartPolish.Apply(chartSpace, polish);
        return Units.Inv($"/slide[{address.SlideIndex}]/chart[{index}]");
    }

    /// <summary>The data props this module edits in place (v1.12, additive).</summary>
    public static readonly IReadOnlyList<string> DataProps = ["title", "categories", "series"];

    /// <summary>True when the props object carries any in-place chart-data key (title/categories/series).</summary>
    public static bool HandlesData(JsonObject props) =>
        props.Any(kv => DataProps.Contains(kv.Key, StringComparer.Ordinal));

    /// <summary>
    /// set /slide[i]/chart[k] with chart-data props (v1.12, additive): edits an
    /// existing chart's <c>title</c> (string to set/replace, <c>false</c> to remove),
    /// its <c>categories</c> (replaces every series' c:cat cache and the embedded
    /// sheet's column A) and its <c>series</c> (replaces each c:ser c:tx name and
    /// c:val cache). Both the chart-XML caches and the embedded "Edit Data" workbook
    /// are rewritten so render/get and PowerPoint both reflect the new data.
    ///
    /// Series are matched by index: passing fewer series than exist leaves the
    /// trailing existing series untouched; passing more updates the overlapping
    /// ones and ignores the surplus (adding/removing series groups would restructure
    /// the plot area — remove and re-add the chart for that). Each replacement
    /// series' values must match the (new or existing) category count.
    /// </summary>
    public static string SetData(PresentationPart presentation, PptxAddress address, JsonObject props)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var (index, _, part) = Resolve(slidePart, address);
        var chartSpace = part.ChartSpace ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The chart part has no chartSpace XML.",
            "Remove the chart and add it again, or restore a snapshot.");

        var chart = chartSpace.GetFirstChild<C.Chart>() ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The chart part has no c:chart element.",
            "Remove the chart and add it again, or restore a snapshot.");

        var serElements = chart.Descendants<C.PlotArea>().FirstOrDefault()?.ChildElements
            .Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal))
            .SelectMany(g => g.ChildElements.Where(e => e.LocalName == "ser"))
            .Cast<OpenXmlCompositeElement>()
            .ToList() ?? [];

        var current = ReadData(part);

        // ----- title --------------------------------------------------------
        if (props.TryGetPropertyValue("title", out var titleNode))
        {
            ApplyTitle(chart, titleNode);
        }

        // ----- categories ---------------------------------------------------
        var categories = current.Categories;
        if (props.TryGetPropertyValue("categories", out var catNode))
        {
            categories = ParseCategoriesArray(catNode);
            // Bubble carries c:xVal as the x channel rather than a c:cat string axis.
            var isBubble = serElements.Any(s => s.ChildElements.Any(c => c.LocalName == "yVal"));
            if (isBubble)
            {
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    "A bubble chart has no category axis to relabel.",
                    "Bubble points are x/y/size triples; remove and re-add the chart to change them.");
            }

            // Changing the category count would leave any series whose values are not
            // also re-supplied mismatched against the new axis; require all series.
            if (categories.Count != current.Categories.Count &&
                !props.ContainsKey("series"))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    Units.Inv($"New 'categories' has {categories.Count} label(s) but the chart had {current.Categories.Count}."),
                    "When you change the number of categories, pass 'series' in the same op with one value " +
                    "per new category for every series, so the data stays aligned.");
            }

            foreach (var ser in serElements)
            {
                ReplaceSerChild(ser, "cat", BuildCategoryData(categories), insertBefore: "val");
            }
        }

        // ----- series -------------------------------------------------------
        if (props.TryGetPropertyValue("series", out var seriesNode))
        {
            // Bubble's yVal/xVal/bubbleSize triples have no single-array c:val to
            // swap; restructuring them in place would mis-pair the channels.
            if (serElements.Any(s => s.ChildElements.Any(c => c.LocalName == "yVal")))
            {
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    "A bubble chart's x/y/size series cannot be replaced in place.",
                    "Remove the chart and add it again with the new bubble series.");
            }

            var replacements = ParseSeriesReplacements(seriesNode, categories.Count, serElements.Count);

            // A category-count change must re-supply every series so none is left
            // with a stale-length value cache against the new axis.
            if (categories.Count != current.Categories.Count && replacements.Count < serElements.Count)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    Units.Inv($"Changing the category count needs all {serElements.Count} series re-supplied; got {replacements.Count}."),
                    "Pass one series object per existing series (in order), each with one value per new category.");
            }

            for (var i = 0; i < serElements.Count && i < replacements.Count; i++)
            {
                var ser = serElements[i];
                var (name, values) = replacements[i];
                if (name is not null)
                {
                    ReplaceSerChild(ser, "tx", BuildSeriesText(name, i), insertAfter: ["order", "idx"]);
                }

                ReplaceSerChild(ser, "val", BuildValues(values, i), insertAfter: ["cat", "order", "idx"]);
            }
        }

        // ----- rebuild the embedded "Edit Data" workbook --------------------
        // Re-read so the workbook mirrors exactly what the caches now hold (names,
        // categories and values), keeping PowerPoint's "Edit Data" sheet in sync.
        if (DataEditable(part) &&
            (props.ContainsKey("categories") || props.ContainsKey("series")))
        {
            RebuildEmbeddedWorkbook(part, chartSpace);
        }

        return Units.Inv($"/slide[{address.SlideIndex}]/chart[{index}]");
    }

    /// <summary>
    /// Sets, replaces or removes the chart's c:title. A string sets/replaces the
    /// title text (autoTitleDeleted=false); <c>false</c> removes it
    /// (autoTitleDeleted=true). null/missing is a no-op (handled by the caller).
    /// </summary>
    private static void ApplyTitle(C.Chart chart, JsonNode? titleNode)
    {
        if (titleNode is JsonValue boolValue && boolValue.TryGetValue<bool>(out var flag))
        {
            if (flag)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "Chart title true is not a title.",
                    "Pass a string to set the title, or false to remove it, e.g. {\"title\":\"Revenue\"} or {\"title\":false}.");
            }

            // false: remove the title and mark it auto-deleted.
            foreach (var existing in chart.Elements<C.Title>().ToList())
            {
                existing.Remove();
            }

            SetAutoTitleDeleted(chart, deleted: true);
            return;
        }

        var text = J.ScalarText(titleNode ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "Chart title cannot be null.",
            "Pass a string to set the title, or false to remove it."));

        var title = new C.Title(
            new C.ChartText(new C.RichText(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.Text(text))))),
            new C.Overlay { Val = false });

        if (chart.Elements<C.Title>().FirstOrDefault() is { } old)
        {
            chart.ReplaceChild(title, old);
        }
        else
        {
            chart.InsertAt(title, 0); // c:title is the first child of c:chart
        }

        SetAutoTitleDeleted(chart, deleted: false);
    }

    /// <summary>Sets c:autoTitleDeleted (creating it after the title if absent).</summary>
    private static void SetAutoTitleDeleted(C.Chart chart, bool deleted)
    {
        if (chart.Elements<C.AutoTitleDeleted>().FirstOrDefault() is { } existing)
        {
            existing.Val = deleted;
            return;
        }

        var node = new C.AutoTitleDeleted { Val = deleted };
        if (chart.Elements<C.Title>().FirstOrDefault() is { } title)
        {
            chart.InsertAfter(node, title);
        }
        else
        {
            chart.InsertAt(node, 0);
        }
    }

    /// <summary>Parses a categories array prop (same rules as add) for an in-place edit.</summary>
    private static List<string> ParseCategoriesArray(JsonNode? node)
    {
        if (node is not JsonArray array || array.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "set chart 'categories' must be a non-empty array.",
                "Pass scalar labels, e.g. {\"categories\":[\"Q1\",\"Q2\",\"Q3\",\"Q4\"]}.");
        }

        var categories = new List<string>(array.Count);
        foreach (var item in array)
        {
            if (item is null or JsonObject or JsonArray)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "set chart 'categories' must hold scalar labels.",
                    "Use strings or numbers, e.g. {\"categories\":[\"Q1\",\"Q2\"]}.");
            }

            categories.Add(J.ScalarText(item));
        }

        return categories;
    }

    /// <summary>
    /// Parses the replacement series for an in-place edit: each {name?, values:[…]}.
    /// Values length must match the (new or kept) category count. Surplus series
    /// beyond the existing count are reported so the caller can warn, but parsing
    /// itself does not reject them — SetData simply ignores the overflow.
    /// </summary>
    private static List<(string? Name, IReadOnlyList<double?> Values)> ParseSeriesReplacements(
        JsonNode? node, int categoryCount, int existingCount)
    {
        if (node is not JsonArray array || array.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "set chart 'series' must be a non-empty array.",
                "Each series looks like {\"name\":\"Sales\",\"values\":[10,20,30,40]}.");
        }

        var result = new List<(string?, IReadOnlyList<double?>)>(array.Count);
        for (var i = 0; i < array.Count; i++)
        {
            if (array[i] is not JsonObject item)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    Units.Inv($"set chart series[{i}] is not an object."),
                    "Each series looks like {\"name\":\"Sales\",\"values\":[10,20,30]}.");
            }

            if (!item.TryGetPropertyValue("values", out var valuesNode) || valuesNode is not JsonArray values)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    Units.Inv($"set chart series[{i}] has no values array."),
                    "Each series needs numeric values, e.g. {\"name\":\"Sales\",\"values\":[10,20,30]}.");
            }

            if (values.Count != categoryCount)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    Units.Inv($"set chart series[{i}] has {values.Count} value(s) but there are {categoryCount} categories."),
                    "Give every series exactly one value per category (null for a gap); " +
                    "pass new 'categories' in the same op to change the count.");
            }

            var numbers = new List<double?>(values.Count);
            for (var v = 0; v < values.Count; v++)
            {
                numbers.Add(NumericValue(values[v], i, v));
            }

            var name = item.TryGetPropertyValue("name", out var nameNode) && nameNode is not null
                ? J.ScalarText(nameNode)
                : null;
            result.Add((name, numbers));
        }

        if (result.Count > existingCount)
        {
            // Surplus series cannot be added without restructuring the plot area;
            // SetData updates the overlapping ones and leaves the rest of the chart
            // as-is. This is the documented "more series than exist" behavior.
        }

        return result;
    }

    /// <summary>
    /// Rebuilds the embedded "Edit Data" workbook from the chart's current caches so
    /// PowerPoint's live sheet matches the edited values, then re-points the chart at
    /// the fresh package part. No-op when the chart carries no embedded workbook.
    /// </summary>
    private static void RebuildEmbeddedWorkbook(ChartPart part, C.ChartSpace chartSpace)
    {
        var data = ReadData(part);

        foreach (var embedded in part.Parts.Select(p => p.OpenXmlPart).OfType<EmbeddedPackagePart>().ToList())
        {
            part.DeletePart(embedded);
        }

        var newRelId = EmbedWorkbook(part, data);
        if (chartSpace.Elements<C.ExternalData>().FirstOrDefault() is { } external)
        {
            external.Id = newRelId;
        }
        else
        {
            chartSpace.Append(new C.ExternalData(new C.AutoUpdate { Val = false }) { Id = newRelId });
        }
    }

    /// <summary>Replaces a ser child by local name, inserting at a schema-valid spot when absent.</summary>
    private static void ReplaceSerChild(
        OpenXmlCompositeElement ser,
        string localName,
        OpenXmlElement replacement,
        string[]? insertAfter = null,
        string? insertBefore = null)
    {
        if (ser.ChildElements.FirstOrDefault(c => c.LocalName == localName) is { } existing)
        {
            ser.ReplaceChild(replacement, existing);
            return;
        }

        if (insertBefore is not null &&
            ser.ChildElements.FirstOrDefault(c => c.LocalName == insertBefore) is { } anchorBefore)
        {
            ser.InsertBefore(replacement, anchorBefore);
            return;
        }

        foreach (var name in insertAfter ?? [])
        {
            if (ser.ChildElements.FirstOrDefault(c => c.LocalName == name) is { } anchorAfter)
            {
                ser.InsertAfter(replacement, anchorAfter);
                return;
            }
        }

        ser.Append(replacement);
    }

    // ----- chartml -------------------------------------------------------------

    private static C.ChartSpace BuildChartSpace(PptxChartData data, string externalDataRelId)
    {
        var plotArea = new C.PlotArea(new C.Layout());
        foreach (var group in BuildChartGroups(data))
        {
            plotArea.Append(group);
        }

        if (CategoryValueAxisKinds.Contains(data.Kind, StringComparer.Ordinal))
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
        else if (data.Kind is "radar" or "bubble")
        {
            // Radar plots categories around the rim against a value axis; bubble
            // plots value-against-value (both axes are value axes).
            plotArea.Append(data.Kind == "radar"
                ? new C.CategoryAxis(
                    new C.AxisId { Val = CategoryAxisId },
                    new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                    new C.Delete { Val = false },
                    new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
                    new C.CrossingAxis { Val = ValueAxisId })
                : new C.ValueAxis(
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
        chartSpace.Append(new C.ExternalData(new C.AutoUpdate { Val = false }) { Id = externalDataRelId });
        return chartSpace;
    }

    /// <summary>
    /// The plot-area chart group(s) for the kind. Every kind yields one group
    /// except combo, which yields a column group (the first series) followed by a
    /// line group (the rest) sharing the same axis pair.
    /// </summary>
    private static IEnumerable<OpenXmlCompositeElement> BuildChartGroups(PptxChartData data)
    {
        switch (data.Kind)
        {
            case "bar":
                yield return BuildBarGroup(data, C.BarGroupingValues.Clustered, Enumerable.Range(0, data.Series.Count));
                break;

            case "stackedBar":
                yield return BuildBarGroup(data, C.BarGroupingValues.Stacked, Enumerable.Range(0, data.Series.Count));
                break;

            case "percentStackedBar":
                yield return BuildBarGroup(data, C.BarGroupingValues.PercentStacked, Enumerable.Range(0, data.Series.Count));
                break;

            case "line":
                yield return BuildLineGroup(data, Enumerable.Range(0, data.Series.Count));
                break;

            case "stackedArea":
                yield return BuildAreaGroup(data, C.GroupingValues.Stacked, Enumerable.Range(0, data.Series.Count));
                break;

            case "radar":
                yield return BuildRadarGroup(data);
                break;

            case "bubble":
                yield return BuildBubbleGroup(data);
                break;

            case "doughnut":
                yield return BuildDoughnutGroup(data);
                break;

            case "combo":
                // First series as columns, the remaining series as a line — both
                // groups reference the shared category/value axis pair.
                yield return BuildBarGroup(data, C.BarGroupingValues.Clustered, [0]);
                yield return BuildLineGroup(data, Enumerable.Range(1, data.Series.Count - 1));
                break;

            default: // "pie" — ParseAdd rejected everything ParseAdd does not handle here
            {
                var group = new C.PieChart(new C.VaryColors { Val = true });
                var series = new C.PieChartSeries();
                AppendSeriesChildren(series, data, 0);
                group.Append(series);
                group.Append(new C.FirstSliceAngle { Val = 0 });
                yield return group;
            }

            break;
        }
    }

    private static C.BarChart BuildBarGroup(PptxChartData data, C.BarGroupingValues grouping, IEnumerable<int> ordinals)
    {
        var group = new C.BarChart(
            new C.BarDirection { Val = C.BarDirectionValues.Column },
            new C.BarGrouping { Val = grouping },
            new C.VaryColors { Val = false });
        foreach (var i in ordinals)
        {
            var series = new C.BarChartSeries();
            AppendSeriesChildren(series, data, i);
            group.Append(series);
        }

        // Stacked bars overlap fully (overlap 100); clustered keep the default.
        if (grouping != C.BarGroupingValues.Clustered)
        {
            group.Append(new C.Overlap { Val = 100 });
        }

        group.Append(new C.AxisId { Val = CategoryAxisId });
        group.Append(new C.AxisId { Val = ValueAxisId });
        return group;
    }

    private static C.LineChart BuildLineGroup(PptxChartData data, IEnumerable<int> ordinals)
    {
        var group = new C.LineChart(
            new C.Grouping { Val = C.GroupingValues.Standard },
            new C.VaryColors { Val = false });
        foreach (var i in ordinals)
        {
            var series = new C.LineChartSeries();
            AppendSeriesChildren(series, data, i);
            group.Append(series);
        }

        group.Append(new C.AxisId { Val = CategoryAxisId });
        group.Append(new C.AxisId { Val = ValueAxisId });
        return group;
    }

    private static C.AreaChart BuildAreaGroup(PptxChartData data, C.GroupingValues grouping, IEnumerable<int> ordinals)
    {
        var group = new C.AreaChart(
            new C.Grouping { Val = grouping },
            new C.VaryColors { Val = false });
        foreach (var i in ordinals)
        {
            var series = new C.AreaChartSeries();
            AppendSeriesChildren(series, data, i);
            group.Append(series);
        }

        group.Append(new C.AxisId { Val = CategoryAxisId });
        group.Append(new C.AxisId { Val = ValueAxisId });
        return group;
    }

    private static C.RadarChart BuildRadarGroup(PptxChartData data)
    {
        var group = new C.RadarChart(
            new C.RadarStyle { Val = C.RadarStyleValues.Marker },
            new C.VaryColors { Val = false });
        for (var i = 0; i < data.Series.Count; i++)
        {
            var series = new C.RadarChartSeries();
            AppendSeriesChildren(series, data, i);
            group.Append(series);
        }

        group.Append(new C.AxisId { Val = CategoryAxisId });
        group.Append(new C.AxisId { Val = ValueAxisId });
        return group;
    }

    /// <summary>A doughnut is a pie variant: a single PieChartSeries plus a holeSize.</summary>
    private static C.DoughnutChart BuildDoughnutGroup(PptxChartData data)
    {
        var group = new C.DoughnutChart(new C.VaryColors { Val = true });
        var series = new C.PieChartSeries();
        AppendSeriesChildren(series, data, 0);
        group.Append(series);
        group.Append(new C.FirstSliceAngle { Val = 0 });
        group.Append(new C.HoleSize { Val = 50 });
        return group;
    }

    /// <summary>idx/order/tx/xVal/yVal/bubbleSize per series — value-against-value with a size channel.</summary>
    private static C.BubbleChart BuildBubbleGroup(PptxChartData data)
    {
        var group = new C.BubbleChart(new C.VaryColors { Val = false });
        for (var i = 0; i < data.Series.Count; i++)
        {
            var spec = data.Series[i];
            var series = new C.BubbleChartSeries(
                new C.Index { Val = (uint)i },
                new C.Order { Val = (uint)i },
                BuildSeriesText(spec.Name, i));
            series.Append(new C.XValues(new C.NumberReference(
                new C.Formula(PptxChartWorkbook.CategoriesRange(data.Categories.Count)),
                BuildNumberCache(spec.XValues ?? spec.Values))));
            series.Append(new C.YValues(new C.NumberReference(
                new C.Formula(PptxChartWorkbook.ValuesRange(i, spec.Values.Count)),
                BuildNumberCache(spec.Values))));
            series.Append(new C.BubbleSize(new C.NumberReference(
                new C.Formula(PptxChartWorkbook.ValuesRange(i, spec.Values.Count)),
                BuildNumberCache(spec.Sizes ?? spec.Values))));
            series.Append(new C.Bubble3D { Val = false });
            group.Append(series);
        }

        group.Append(new C.AxisId { Val = CategoryAxisId });
        group.Append(new C.AxisId { Val = ValueAxisId });
        return group;
    }

    /// <summary>idx/order/tx/cat/val with reference caches pointing at the embedded Sheet1.</summary>
    private static void AppendSeriesChildren(OpenXmlCompositeElement series, PptxChartData data, int ordinal)
    {
        var spec = data.Series[ordinal];
        series.Append(new C.Index { Val = (uint)ordinal });
        series.Append(new C.Order { Val = (uint)ordinal });
        series.Append(BuildSeriesText(spec.Name, ordinal));
        series.Append(BuildCategoryData(data.Categories));
        series.Append(BuildValues(spec.Values, ordinal));
    }

    /// <summary>c:tx as a string reference into the embedded sheet's header row.</summary>
    private static C.SeriesText BuildSeriesText(string name, int ordinal) => new(
        new C.StringReference(
            new C.Formula(PptxChartWorkbook.SeriesNameReference(ordinal)),
            new C.StringCache(
                new C.PointCount { Val = 1u },
                new C.StringPoint { Index = 0u, NumericValue = new C.NumericValue(name) })));

    /// <summary>c:cat as a string reference (with cache) into column A of the embedded sheet.</summary>
    private static C.CategoryAxisData BuildCategoryData(IReadOnlyList<string> categories)
    {
        var cache = new C.StringCache(new C.PointCount { Val = (uint)categories.Count });
        for (var i = 0; i < categories.Count; i++)
        {
            cache.Append(new C.StringPoint
            {
                Index = (uint)i,
                NumericValue = new C.NumericValue(categories[i]),
            });
        }

        return new C.CategoryAxisData(new C.StringReference(
            new C.Formula(PptxChartWorkbook.CategoriesRange(categories.Count)),
            cache));
    }

    /// <summary>c:val as a number reference (with cache) into the series' column of the embedded sheet.</summary>
    private static C.Values BuildValues(IReadOnlyList<double?> values, int ordinal) =>
        new(new C.NumberReference(
            new C.Formula(PptxChartWorkbook.ValuesRange(ordinal, values.Count)),
            BuildNumberCache(values)));

    /// <summary>A c:numCache of the values (null entries are dropped, matching the embedded sheet's gaps).</summary>
    private static C.NumberingCache BuildNumberCache(IReadOnlyList<double?> values)
    {
        var cache = new C.NumberingCache(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)values.Count });
        for (var i = 0; i < values.Count; i++)
        {
            if (values[i] is { } value)
            {
                cache.Append(new C.NumericPoint
                {
                    Index = (uint)i,
                    NumericValue = new C.NumericValue(value.ToString(CultureInfo.InvariantCulture)),
                });
            }
        }

        return cache;
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
        var groups = plotArea?.ChildElements
            .Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal))
            .ToList() ?? [];

        var kind = ResolveKind(groups);
        var title = TitleText(chartSpace);

        // Every series across every group (combo has a bar group + a line group).
        var serElements = groups.SelectMany(g => g.ChildElements.Where(e => e.LocalName == "ser")).ToList();

        var categories = new List<string>();
        var firstCat = serElements
            .Select(ser => ser.ChildElements.FirstOrDefault(e => e.LocalName == "cat"))
            .FirstOrDefault(c => c is not null);
        if (firstCat is not null)
        {
            categories = ReadStringPoints(firstCat);
        }

        var series = new List<PptxChartSeries>(serElements.Count);
        foreach (var ser in serElements)
        {
            var name = ser.ChildElements.FirstOrDefault(e => e.LocalName == "tx")?
                .Descendants<C.NumericValue>().FirstOrDefault()?.Text ?? string.Empty;
            // bubble carries yVal; bar/line/area/pie/radar carry val.
            var val = ser.ChildElements.FirstOrDefault(e => e.LocalName is "val" or "yVal");
            series.Add(new PptxChartSeries(name, val is null ? [] : ReadNumericPoints(val)));
        }

        return new PptxChartData(kind, title, categories, series);
    }

    /// <summary>
    /// The aioffice kind token of a chart's plot-area groups: two groups (bar +
    /// line) is combo; a single group maps by its element and grouping (stacked /
    /// percentStacked variants are distinguished by the c:grouping attribute).
    /// </summary>
    private static string ResolveKind(IReadOnlyList<OpenXmlElement> groups)
    {
        if (groups.Count == 0)
        {
            return "unknown";
        }

        if (groups.Count > 1 &&
            groups.Any(g => g.LocalName == "barChart") && groups.Any(g => g.LocalName == "lineChart"))
        {
            return "combo";
        }

        var group = groups[0];
        var grouping = group.ChildElements
            .FirstOrDefault(e => e.LocalName is "grouping" or "barGrouping")?
            .GetAttribute("val", string.Empty).Value;

        return group.LocalName switch
        {
            "barChart" => grouping switch
            {
                "stacked" => "stackedBar",
                "percentStacked" => "percentStackedBar",
                _ => "bar",
            },
            "areaChart" => grouping == "stacked" ? "stackedArea" : "area",
            "lineChart" => "line",
            "pieChart" => "pie",
            "doughnutChart" => "doughnut",
            "radarChart" => "radar",
            "bubbleChart" => "bubble",
            "scatterChart" => "scatter",
            _ => KindName(group.LocalName),
        };
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
            DataEditable = DataEditable(part),
            Polish = part.ChartSpace is { } chartSpace ? PptxChartPolish.ReadSettings(chartSpace) : null,
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
            DataEditable = DataEditable(part),
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
