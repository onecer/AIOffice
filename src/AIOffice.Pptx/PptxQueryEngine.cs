using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>Selector evaluation and node projections (the read side of query/get).</summary>
internal static class PptxQueryEngine
{
    private static readonly IReadOnlyList<string> QueryElements = ["slide", "shape", "master", "layout", "*"];
    private static readonly IReadOnlyList<string> ShapeAttributes = ["id", "name", "fill", "text", "kind", "placeholder"];
    private static readonly IReadOnlyList<string> LayoutAttributes = ["name", "type"];

    public static List<object> Query(PresentationPart presentation, Selector selector, PptxAddress? scope = null)
    {
        if (!QueryElements.Contains(selector.Element, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"pptx documents have no queryable element '{selector.Element}'.",
                "Query slide, shape, master or layout (* matches slides and shapes), e.g. shape:contains('Q3') or layout[type=blank].",
                candidates: QueryElements);
        }

        if (scope is not null)
        {
            return QueryScoped(presentation, selector, scope);
        }

        return selector.Element switch
        {
            "master" => QueryMasters(presentation, selector),
            "layout" => QueryLayouts(presentation, selector),
            _ => QuerySlides(presentation, selector, onlySlideIndex: null),
        };
    }

    private static List<object> QuerySlides(PresentationPart presentation, Selector selector, int? onlySlideIndex)
    {
        var results = new List<object>();
        var slides = PptxDoc.Slides(presentation);
        for (var i = 0; i < slides.Count; i++)
        {
            var slideIndex = i + 1;
            if (onlySlideIndex is { } only && slideIndex != only)
            {
                continue;
            }

            var shapes = PptxDoc.Shapes(slides[i].Part);

            if (selector.Element is "slide" or "*" && MatchesSlide(shapes, selector))
            {
                results.Add(new
                {
                    Path = Units.Inv($"/slide[{slideIndex}]"),
                    Kind = "slide",
                    Slide = slideIndex,
                    Text = Snippet(string.Join('\n', shapes.Select(s => PptxDoc.ShapeText(s.Element)).Where(t => t.Length > 0))),
                });
            }

            if (selector.Element is "shape" or "*")
            {
                foreach (var shape in shapes.Where(s => MatchesShape(s, selector)))
                {
                    results.Add(new
                    {
                        Path = shape.CanonicalPath(slideIndex),
                        OrdinalPath = shape.OrdinalPath(slideIndex),
                        Kind = shape.Kind,
                        Slide = slideIndex,
                        Name = shape.Name,
                        Text = Snippet(PptxDoc.ShapeText(shape.Element)),
                    });
                }
            }
        }

        return results;
    }

    /// <summary>Query restricted to one container: a slide, a master or one of its layouts.</summary>
    private static List<object> QueryScoped(PresentationPart presentation, Selector selector, PptxAddress scope)
    {
        if (scope.HasShape || scope.ParagraphIndex is not null || scope.IsNotes || scope.IsChart ||
            scope.IsAnimation || scope.IsComment)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Query scope must be a container, not '{scope.Raw}'.",
                "Scope to /slide[2], /master[1] or /master[1]/layout[2].");
        }

        if (!scope.IsMaster)
        {
            _ = PptxDoc.ResolveSlide(presentation, scope.SlideIndex, scope.Raw); // throws invalid_path with candidates
            return QuerySlides(presentation, selector, scope.SlideIndex);
        }

        if (selector.Element is not ("shape" or "*"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Only shapes can be queried inside a master/layout scope; got '{selector.Element}'.",
                "Query shape (or *) with --scope, e.g. shape:contains('Title') scoped to /master[1]/layout[2].",
                candidates: ["shape", "*"]);
        }

        var masterPart = PptxDoc.ResolveMaster(presentation, scope.MasterIndex, scope.Raw);
        var tree = scope.LayoutIndex is { } layoutIndex
            ? PptxDoc.RequireShapeTree(PptxDoc.ResolveLayout(masterPart, scope.MasterIndex, layoutIndex, scope.Raw))
            : PptxDoc.RequireShapeTree(masterPart);

        var containerPath = scope.CanonicalContainerPath;
        var results = new List<object>();
        foreach (var shape in PptxDoc.Shapes(tree).Where(s => MatchesShape(s, selector)))
        {
            results.Add(new
            {
                Path = shape.CanonicalPathIn(containerPath),
                OrdinalPath = shape.OrdinalPathIn(containerPath),
                Kind = shape.Kind,
                Master = scope.MasterIndex,
                Layout = scope.LayoutIndex,
                Name = shape.Name,
                Placeholder = PptxDoc.PlaceholderType(shape.Element),
                Text = Snippet(PptxDoc.ShapeText(shape.Element)),
            });
        }

        return results;
    }

    private static List<object> QueryMasters(PresentationPart presentation, Selector selector)
    {
        var results = new List<object>();
        foreach (var (masterIndex, masterPart) in PptxDoc.Masters(presentation))
        {
            var shapes = PptxDoc.Shapes(PptxDoc.RequireShapeTree(masterPart));
            if (!MatchesContainer(shapes, selector, "master", layoutAttributes: null, name: null, type: null))
            {
                continue;
            }

            results.Add(new
            {
                Path = Units.Inv($"/master[{masterIndex}]"),
                Kind = "master",
                Master = masterIndex,
                LayoutCount = PptxDoc.Layouts(masterPart).Count,
                Text = Snippet(string.Join('\n', shapes.Select(s => PptxDoc.ShapeText(s.Element)).Where(t => t.Length > 0))),
            });
        }

        return results;
    }

    private static List<object> QueryLayouts(PresentationPart presentation, Selector selector)
    {
        var results = new List<object>();
        foreach (var (masterIndex, masterPart) in PptxDoc.Masters(presentation))
        {
            foreach (var (layoutIndex, layoutPart) in PptxDoc.Layouts(masterPart))
            {
                var shapes = PptxDoc.Shapes(PptxDoc.RequireShapeTree(layoutPart));
                var name = PptxDoc.LayoutName(layoutPart);
                var type = PptxDoc.LayoutType(layoutPart);
                if (!MatchesContainer(shapes, selector, "layout", LayoutAttributes, name, type))
                {
                    continue;
                }

                results.Add(new
                {
                    Path = Units.Inv($"/master[{masterIndex}]/layout[{layoutIndex}]"),
                    Kind = "layout",
                    Master = masterIndex,
                    Name = name,
                    Type = type,
                    Text = Snippet(string.Join('\n', shapes.Select(s => PptxDoc.ShapeText(s.Element)).Where(t => t.Length > 0))),
                });
            }
        }

        return results;
    }

    /// <summary>Predicate evaluation for master/layout containers: :contains over shape text, plus name/type for layouts.</summary>
    private static bool MatchesContainer(
        List<ShapeView> shapes,
        Selector selector,
        string elementName,
        IReadOnlyList<string>? layoutAttributes,
        string? name,
        string? type)
    {
        foreach (var predicate in selector.Predicates)
        {
            switch (predicate)
            {
                case ContainsPredicate contains:
                    if (!shapes.Any(s => PptxDoc.ShapeText(s.Element).Contains(contains.Text, StringComparison.OrdinalIgnoreCase)))
                    {
                        return false;
                    }

                    break;

                case AttributePredicate attribute when layoutAttributes is not null && attribute.Attribute == "name":
                    if (!CompareText(name ?? string.Empty, attribute))
                    {
                        return false;
                    }

                    break;

                case AttributePredicate attribute when layoutAttributes is not null && attribute.Attribute == "type":
                    if (!CompareText(type ?? string.Empty, attribute))
                    {
                        return false;
                    }

                    break;

                case AttributePredicate attribute:
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"'{elementName}' has no queryable attribute '{attribute.Attribute}'.",
                        layoutAttributes is null
                            ? "Masters support only :contains('text'); name/type belong to layout."
                            : "Layouts support name, type and :contains('text').",
                        candidates: layoutAttributes ?? [":contains('text')"]);
            }
        }

        return true;
    }

    private static bool MatchesSlide(List<ShapeView> shapes, Selector selector)
    {
        foreach (var predicate in selector.Predicates)
        {
            switch (predicate)
            {
                case ContainsPredicate contains:
                    if (!shapes.Any(s => PptxDoc.ShapeText(s.Element).Contains(contains.Text, StringComparison.OrdinalIgnoreCase)))
                    {
                        return false;
                    }

                    break;

                case AttributePredicate attribute:
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Slides have no queryable attribute '{attribute.Attribute}'.",
                        "Slides support only :contains('text'); attributes like fill/name/id belong to shape.",
                        candidates: [":contains('text')"]);
            }
        }

        return true;
    }

    private static bool MatchesShape(ShapeView shape, Selector selector)
    {
        var text = PptxDoc.ShapeText(shape.Element);
        foreach (var predicate in selector.Predicates)
        {
            switch (predicate)
            {
                case ContainsPredicate contains:
                    if (!text.Contains(contains.Text, StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }

                    break;

                case AttributePredicate attribute:
                    if (!MatchesAttribute(shape, text, attribute))
                    {
                        return false;
                    }

                    break;
            }
        }

        return true;
    }

    private static bool MatchesAttribute(ShapeView shape, string text, AttributePredicate predicate)
    {
        switch (predicate.Attribute)
        {
            case "id":
                return CompareNumeric(shape.Id, predicate);
            case "name":
                return CompareText(shape.Name, predicate);
            case "text":
                return CompareText(text, predicate);
            case "kind":
                return CompareText(shape.Kind, predicate);
            case "placeholder":
                return CompareText(PptxDoc.PlaceholderType(shape.Element) ?? string.Empty, predicate);
            case "fill":
                var fill = PptxDoc.FillHex(shape.Element);
                var wanted = predicate.Value.TrimStart('#').ToUpperInvariant();
                var equal = string.Equals(fill, wanted, StringComparison.Ordinal);
                return predicate.Op switch
                {
                    SelectorOperator.Equals => equal,
                    SelectorOperator.NotEquals => !equal,
                    _ => throw UnsupportedOperator(predicate, "fill supports = and !="),
                };
            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Shapes have no queryable attribute '{predicate.Attribute}'.",
                    "Queryable shape attributes: id, name, fill, text, kind, placeholder.",
                    candidates: ShapeAttributes);
        }
    }

    private static bool CompareNumeric(uint actual, AttributePredicate predicate)
    {
        if (predicate.NumericValue is not { } wanted)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Attribute '{predicate.Attribute}' needs a numeric value; got '{predicate.Value}'.",
                "Compare ids numerically, e.g. shape[id=4].");
        }

        return predicate.Op switch
        {
            SelectorOperator.Equals => actual == wanted,
            SelectorOperator.NotEquals => Math.Abs(actual - wanted) >= double.Epsilon,
            SelectorOperator.GreaterThan => actual > wanted,
            SelectorOperator.GreaterOrEqual => actual >= wanted,
            SelectorOperator.LessThan => actual < wanted,
            SelectorOperator.LessOrEqual => actual <= wanted,
            _ => throw UnsupportedOperator(predicate, "ids support = != > >= < <="),
        };
    }

    private static bool CompareText(string actual, AttributePredicate predicate) => predicate.Op switch
    {
        SelectorOperator.Equals => string.Equals(actual, predicate.Value, StringComparison.Ordinal),
        SelectorOperator.NotEquals => !string.Equals(actual, predicate.Value, StringComparison.Ordinal),
        SelectorOperator.ContainsText => actual.Contains(predicate.Value, StringComparison.OrdinalIgnoreCase),
        _ => throw UnsupportedOperator(predicate, "text attributes support =, != and *="),
    };

    private static AiofficeException UnsupportedOperator(AttributePredicate predicate, string hint) => new(
        ErrorCodes.InvalidArgs,
        $"Operator not supported for attribute '{predicate.Attribute}' in this comparison.",
        hint + ". Run 'aioffice help selectors' for the grammar.");

    /// <summary>The `get` projection for a slide path.</summary>
    public static object SlideDetail(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var shapes = PptxDoc.Shapes(slidePart);
        var notes = PptxNotes.Text(slidePart);
        var transition = PptxTransitions.Read(slidePart);
        var charts = PptxCharts.Charts(slidePart);
        return new
        {
            Path = address.CanonicalSlidePath,
            Index = address.SlideIndex,
            Background = PptxDoc.BackgroundHex(slidePart),
            Transition = transition?.Kind,
            TransitionDuration = transition?.Duration,
            ShapeCount = shapes.Count,
            Shapes = shapes.Select(s => new
            {
                Path = s.CanonicalPath(address.SlideIndex),
                OrdinalPath = s.OrdinalPath(address.SlideIndex),
                Kind = s.Kind,
                Name = s.Name,
                Text = Snippet(PptxDoc.ShapeText(s.Element)),
            }).ToList<object>(),
            Charts = charts.Count == 0
                ? null
                : charts.Select(c => (object)Units.Inv($"{address.CanonicalSlidePath}/chart[{c.Index}]")).ToList(),
            Notes = notes.Length == 0 ? null : Snippet(notes),
        };
    }

    /// <summary>The `get` projection for a shape path: identity, text, geometry in cm, fill, font.</summary>
    public static object ShapeDetail(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var view = PptxDoc.ResolveShape(slidePart, address);

        if (address.ParagraphIndex is { } paragraphIndex)
        {
            var paragraph = PptxDoc.ResolveParagraph(view, address);
            return new
            {
                Path = Units.Inv($"{view.CanonicalPath(address.SlideIndex)}/p[{paragraphIndex}]"),
                Text = PptxDoc.ParagraphText(paragraph),
                Font = FontInfo(paragraph),
            };
        }

        var geometry = PptxDoc.Geometry(view.Element);
        var firstParagraph = (view.Element as P.Shape)?.TextBody?.Elements<A.Paragraph>().FirstOrDefault();
        var chartIndex = PptxCharts.IndexOf(slidePart, view.Element);
        return new
        {
            Path = view.CanonicalPath(address.SlideIndex),
            OrdinalPath = view.OrdinalPath(address.SlideIndex),
            ChartPath = chartIndex is { } ci ? Units.Inv($"{address.CanonicalSlidePath}/chart[{ci}]") : null,
            Slide = address.SlideIndex,
            Id = view.Id,
            Ordinal = view.Ordinal,
            ZIndex = view.Ordinal, // paint order: 1 = painted first (bottom)
            Kind = view.Kind,
            Name = view.Name,
            Placeholder = PptxDoc.PlaceholderType(view.Element),
            Geometry = PptxDoc.GeometryToken(view.Element),
            Flip = PptxDoc.FlipToken(view.Element),
            Text = PptxDoc.ShapeText(view.Element),
            X = geometry is { } g1 ? Units.EmuToCm(g1.X) : (double?)null,
            Y = geometry is { } g2 ? Units.EmuToCm(g2.Y) : (double?)null,
            W = geometry is { } g3 ? Units.EmuToCm(g3.Cx) : (double?)null,
            H = geometry is { } g4 ? Units.EmuToCm(g4.Cy) : (double?)null,
            Fill = PptxDoc.FillHex(view.Element),
            LineColor = PptxDoc.LineHex(view.Element),
            Font = firstParagraph is null ? null : FontInfo(firstParagraph),
            Chart = PptxCharts.Summary(slidePart, view.Element),
        };
    }

    /// <summary>The `get` projection for /master[m]: identity, its layouts (with usage) and its shapes.</summary>
    public static object MasterDetail(PresentationPart presentation, PptxAddress address)
    {
        var masterPart = PptxDoc.ResolveMaster(presentation, address.MasterIndex, address.Raw);
        var containerPath = address.CanonicalContainerPath;
        var shapes = PptxDoc.Shapes(PptxDoc.RequireShapeTree(masterPart));
        var layouts = PptxDoc.Layouts(masterPart);
        return new
        {
            Path = containerPath,
            Index = address.MasterIndex,
            Kind = "master",
            Theme = masterPart.ThemePart?.Theme?.Name?.Value,
            LayoutCount = layouts.Count,
            Layouts = layouts.Select(l => LayoutSummary(presentation, address.MasterIndex, l.Index, l.Part)).ToList<object>(),
            ShapeCount = shapes.Count,
            Shapes = ShapeSummaries(shapes, containerPath),
        };
    }

    /// <summary>The `get` projection for /master[m]/layout[l]: identity, which slides use it, its shapes.</summary>
    public static object LayoutDetail(PresentationPart presentation, PptxAddress address)
    {
        var masterPart = PptxDoc.ResolveMaster(presentation, address.MasterIndex, address.Raw);
        var layoutPart = PptxDoc.ResolveLayout(masterPart, address.MasterIndex, address.LayoutIndex!.Value, address.Raw);
        var containerPath = address.CanonicalContainerPath;
        var shapes = PptxDoc.Shapes(PptxDoc.RequireShapeTree(layoutPart));
        var summary = LayoutSummary(presentation, address.MasterIndex, address.LayoutIndex.Value, layoutPart);
        return new
        {
            Path = containerPath,
            Index = address.LayoutIndex.Value,
            Master = address.MasterIndex,
            Kind = "layout",
            Name = summary.Name,
            Type = summary.Type,
            UsedBySlides = summary.UsedBySlides,
            ShapeCount = shapes.Count,
            Shapes = ShapeSummaries(shapes, containerPath),
        };
    }

    /// <summary>The `get` projection for a shape under a master or layout (read-only in this milestone).</summary>
    public static object MasterShapeDetail(PresentationPart presentation, PptxAddress address)
    {
        var masterPart = PptxDoc.ResolveMaster(presentation, address.MasterIndex, address.Raw);
        var tree = address.LayoutIndex is { } layoutIndex
            ? PptxDoc.RequireShapeTree(PptxDoc.ResolveLayout(masterPart, address.MasterIndex, layoutIndex, address.Raw))
            : PptxDoc.RequireShapeTree(masterPart);

        var containerPath = address.CanonicalContainerPath;
        var label = address.LayoutIndex is { } li
            ? Units.Inv($"on layout {li} of master {address.MasterIndex}")
            : Units.Inv($"on master {address.MasterIndex}");
        var view = PptxDoc.ResolveShape(PptxDoc.Shapes(tree), address, containerPath, label);

        var geometry = PptxDoc.Geometry(view.Element);
        var firstParagraph = (view.Element as P.Shape)?.TextBody?.Elements<A.Paragraph>().FirstOrDefault();
        return new
        {
            Path = view.CanonicalPathIn(containerPath),
            OrdinalPath = view.OrdinalPathIn(containerPath),
            Master = address.MasterIndex,
            Layout = address.LayoutIndex,
            Id = view.Id,
            Ordinal = view.Ordinal,
            Kind = view.Kind,
            Name = view.Name,
            Placeholder = PptxDoc.PlaceholderType(view.Element),
            Text = PptxDoc.ShapeText(view.Element),
            X = geometry is { } g1 ? Units.EmuToCm(g1.X) : (double?)null,
            Y = geometry is { } g2 ? Units.EmuToCm(g2.Y) : (double?)null,
            W = geometry is { } g3 ? Units.EmuToCm(g3.Cx) : (double?)null,
            H = geometry is { } g4 ? Units.EmuToCm(g4.Cy) : (double?)null,
            Fill = PptxDoc.FillHex(view.Element),
            Font = firstParagraph is null ? null : FontInfo(firstParagraph),
        };
    }

    internal sealed record LayoutSummaryView(
        string Path, int Index, string? Name, string? Type, IReadOnlyList<int> UsedBySlides);

    /// <summary>One layout row shared by get /master[m] and read --view structure.</summary>
    internal static LayoutSummaryView LayoutSummary(
        PresentationPart presentation, int masterIndex, int layoutIndex, SlideLayoutPart layoutPart)
    {
        var usedBy = new List<int>();
        var slides = PptxDoc.Slides(presentation);
        for (var i = 0; i < slides.Count; i++)
        {
            if (slides[i].Part.SlideLayoutPart?.Uri == layoutPart.Uri)
            {
                usedBy.Add(i + 1);
            }
        }

        return new LayoutSummaryView(
            Units.Inv($"/master[{masterIndex}]/layout[{layoutIndex}]"),
            layoutIndex,
            PptxDoc.LayoutName(layoutPart),
            PptxDoc.LayoutType(layoutPart),
            usedBy);
    }

    private static List<object> ShapeSummaries(List<ShapeView> shapes, string containerPath) => shapes
        .Select(s => (object)new
        {
            Path = s.CanonicalPathIn(containerPath),
            OrdinalPath = s.OrdinalPathIn(containerPath),
            Kind = s.Kind,
            Name = s.Name,
            Placeholder = PptxDoc.PlaceholderType(s.Element),
            Text = Snippet(PptxDoc.ShapeText(s.Element)),
        })
        .ToList();

    /// <summary>Font properties of a paragraph's first run (the inspectable summary).</summary>
    public static object? FontInfo(A.Paragraph paragraph)
    {
        var runProperties = paragraph.Elements<A.Run>().FirstOrDefault()?.RunProperties;
        var alignment = paragraph.ParagraphProperties?.Alignment;
        string? align = null;
        if (alignment is not null)
        {
            if (alignment.Value == A.TextAlignmentTypeValues.Left)
            {
                align = "left";
            }
            else if (alignment.Value == A.TextAlignmentTypeValues.Center)
            {
                align = "center";
            }
            else if (alignment.Value == A.TextAlignmentTypeValues.Right)
            {
                align = "right";
            }
            else if (alignment.Value == A.TextAlignmentTypeValues.Justified)
            {
                align = "justify";
            }
        }

        var sizePt = runProperties?.FontSize?.Value is { } size
            ? Math.Round(size / 100.0, 2)
            : (double?)null;
        var bold = runProperties?.Bold?.Value;
        var color = runProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant();

        if (sizePt is null && bold is null && color is null && align is null)
        {
            return null;
        }

        return new { SizePt = sizePt, Bold = bold, Color = color, Align = align };
    }

    public static string Snippet(string text)
    {
        var flat = text.Replace('\n', ' ').Trim();
        return flat.Length <= 80 ? flat : flat[..77] + "...";
    }
}
