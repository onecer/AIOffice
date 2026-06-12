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
    private static readonly IReadOnlyList<string> QueryElements = ["slide", "shape", "*"];
    private static readonly IReadOnlyList<string> ShapeAttributes = ["id", "name", "fill", "text", "kind"];

    public static List<object> Query(PresentationPart presentation, Selector selector)
    {
        if (!QueryElements.Contains(selector.Element, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"pptx documents have no queryable element '{selector.Element}'.",
                "Query slide or shape (or * for both), e.g. shape:contains('Q3') or shape[fill=FF0000].",
                candidates: QueryElements);
        }

        var results = new List<object>();
        var slides = PptxDoc.Slides(presentation);
        for (var i = 0; i < slides.Count; i++)
        {
            var slideIndex = i + 1;
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
                    "Queryable shape attributes: id, name, fill, text, kind.",
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
        return new
        {
            Path = address.CanonicalSlidePath,
            Index = address.SlideIndex,
            ShapeCount = shapes.Count,
            Shapes = shapes.Select(s => new
            {
                Path = s.CanonicalPath(address.SlideIndex),
                OrdinalPath = s.OrdinalPath(address.SlideIndex),
                Kind = s.Kind,
                Name = s.Name,
                Text = Snippet(PptxDoc.ShapeText(s.Element)),
            }).ToList<object>(),
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
        return new
        {
            Path = view.CanonicalPath(address.SlideIndex),
            OrdinalPath = view.OrdinalPath(address.SlideIndex),
            Slide = address.SlideIndex,
            Id = view.Id,
            Ordinal = view.Ordinal,
            Kind = view.Kind,
            Name = view.Name,
            Text = PptxDoc.ShapeText(view.Element),
            X = geometry is { } g1 ? Units.EmuToCm(g1.X) : (double?)null,
            Y = geometry is { } g2 ? Units.EmuToCm(g2.Y) : (double?)null,
            W = geometry is { } g3 ? Units.EmuToCm(g3.Cx) : (double?)null,
            H = geometry is { } g4 ? Units.EmuToCm(g4.Cy) : (double?)null,
            Fill = PptxDoc.FillHex(view.Element),
            Font = firstParagraph is null ? null : FontInfo(firstParagraph),
        };
    }

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
