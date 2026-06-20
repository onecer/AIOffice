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
    private static readonly IReadOnlyList<string> QueryElements =
        ["slide", "shape", "table", "tc", "smartart", "master", "layout", "*"];
    private static readonly IReadOnlyList<string> ShapeAttributes = ["id", "name", "fill", "text", "kind", "placeholder", "hyperlink"];
    private static readonly IReadOnlyList<string> LayoutAttributes = ["name", "type"];

    public static List<object> Query(PresentationPart presentation, Selector selector, PptxAddress? scope = null)
    {
        if (!QueryElements.Contains(selector.Element, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"pptx documents have no queryable element '{selector.Element}'.",
                "Query slide, shape, table, tc (table cells), smartart, master or layout " +
                "(* matches slides and shapes), e.g. shape:contains('Q3') or tc:contains('total').",
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

            var slidePart = slides[i].Part;
            var shapes = PptxDoc.Shapes(slidePart);

            if (selector.Element is "slide" or "*" && MatchesSlide(slidePart, shapes, selector))
            {
                results.Add(new
                {
                    Path = Units.Inv($"/slide[{slideIndex}]"),
                    Kind = "slide",
                    Slide = slideIndex,
                    Text = Snippet(SlideAggregateText(slidePart, shapes)),
                });
            }

            if (selector.Element is "shape" or "*")
            {
                string? ResolveLink(ShapeView s) => PptxHyperlinks.Resolve(presentation, slidePart, s.Element);
                foreach (var shape in shapes.Where(s => MatchesShape(s, selector, ResolveLink)))
                {
                    results.Add(new
                    {
                        Path = shape.CanonicalPath(slideIndex),
                        OrdinalPath = shape.OrdinalPath(slideIndex),
                        Kind = shape.Kind,
                        Slide = slideIndex,
                        Name = shape.Name,
                        Text = Snippet(PptxDoc.ShapeText(shape.Element)),
                        Hyperlink = ResolveLink(shape),
                    });
                }
            }

            if (selector.Element is "table" or "tc")
            {
                QueryTables(slidePart, slideIndex, selector, results);
            }

            if (selector.Element is "smartart")
            {
                QuerySmartArt(slidePart, slideIndex, selector, results);
            }
        }

        return results;
    }

    /// <summary>All text a slide carries: shape bodies plus table cells plus SmartArt nodes.</summary>
    private static string SlideAggregateText(SlidePart slidePart, List<ShapeView> shapes)
    {
        var parts = new List<string>();
        foreach (var shape in shapes)
        {
            var text = PptxDoc.ShapeText(shape.Element);
            if (text.Length == 0 && PptxTables.TableOf(shape.Element) is { } table)
            {
                text = PptxTables.TableText(table);
            }

            if (text.Length == 0 && PptxSmartArt.DataPartOf(slidePart, shape.Element) is { } dataPart)
            {
                text = PptxSmartArt.FlatText(dataPart);
            }

            if (text.Length > 0)
            {
                parts.Add(text);
            }
        }

        return string.Join('\n', parts);
    }

    /// <summary>table / tc queries on one slide: :contains over cell text.</summary>
    private static void QueryTables(SlidePart slidePart, int slideIndex, Selector selector, List<object> results)
    {
        foreach (var (index, _, table) in PptxTables.Tables(slidePart))
        {
            var tablePath = Units.Inv($"/slide[{slideIndex}]/table[{index}]");
            if (selector.Element == "table")
            {
                if (MatchesTextOnly(PptxTables.TableText(table), selector, "table"))
                {
                    results.Add(new
                    {
                        Path = tablePath,
                        Kind = "table",
                        Slide = slideIndex,
                        Rows = table.Elements<A.TableRow>().Count(),
                        Cols = table.TableGrid?.Elements<A.GridColumn>().Count() ?? 0,
                        Text = Snippet(PptxTables.TableText(table)),
                    });
                }

                continue;
            }

            var rowIndex = 0;
            foreach (var row in table.Elements<A.TableRow>())
            {
                rowIndex++;
                var cellIndex = 0;
                foreach (var cell in row.Elements<A.TableCell>())
                {
                    cellIndex++;
                    var text = PptxTables.CellText(cell);
                    if (!MatchesTextOnly(text, selector, "tc"))
                    {
                        continue;
                    }

                    results.Add(new
                    {
                        Path = Units.Inv($"{tablePath}/tr[{rowIndex}]/tc[{cellIndex}]"),
                        Kind = "tc",
                        Slide = slideIndex,
                        Row = rowIndex,
                        Col = cellIndex,
                        Text = Snippet(text),
                    });
                }
            }
        }
    }

    /// <summary>smartart queries on one slide: :contains over the diagram's node texts.</summary>
    private static void QuerySmartArt(SlidePart slidePart, int slideIndex, Selector selector, List<object> results)
    {
        foreach (var (index, view, part) in PptxSmartArt.List(slidePart))
        {
            var text = PptxSmartArt.FlatText(part);
            if (!MatchesTextOnly(text, selector, "smartart"))
            {
                continue;
            }

            results.Add(new
            {
                Path = Units.Inv($"/slide[{slideIndex}]/smartart[{index}]"),
                Kind = "smartart",
                Slide = slideIndex,
                Layout = PptxSmartArt.LayoutName(slidePart, view.Element),
                Text = Snippet(text),
            });
        }
    }

    /// <summary>Predicate evaluation for text-only elements (table, tc, smartart): just :contains.</summary>
    private static bool MatchesTextOnly(string text, Selector selector, string elementName)
    {
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
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"'{elementName}' has no queryable attribute '{attribute.Attribute}'.",
                        $"'{elementName}' supports only :contains('text').",
                        candidates: [":contains('text')"]);
            }
        }

        return true;
    }

    /// <summary>Query restricted to one container: a slide, a master or one of its layouts.</summary>
    private static List<object> QueryScoped(PresentationPart presentation, Selector selector, PptxAddress scope)
    {
        if (scope.HasShape || scope.ParagraphIndex is not null || scope.IsNotes || scope.IsChart ||
            scope.IsTable || scope.IsSmartArt || scope.IsAnimation || scope.IsComment)
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

    private static bool MatchesSlide(SlidePart slidePart, List<ShapeView> shapes, Selector selector)
    {
        foreach (var predicate in selector.Predicates)
        {
            switch (predicate)
            {
                case ContainsPredicate contains:
                    if (!SlideAggregateText(slidePart, shapes).Contains(contains.Text, StringComparison.OrdinalIgnoreCase))
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

    private static bool MatchesShape(ShapeView shape, Selector selector, Func<ShapeView, string?>? resolveLink = null)
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
                    if (!MatchesAttribute(shape, text, attribute, resolveLink))
                    {
                        return false;
                    }

                    break;
            }
        }

        return true;
    }

    private static bool MatchesAttribute(ShapeView shape, string text, AttributePredicate predicate, Func<ShapeView, string?>? resolveLink)
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
            case "hyperlink":
                return MatchesHyperlink(resolveLink?.Invoke(shape), predicate);
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
                    "Queryable shape attributes: id, name, fill, text, kind, placeholder, hyperlink.",
                    candidates: ShapeAttributes);
        }
    }

    /// <summary>
    /// Matches the shape's resolved hyperlink form. The sentinel value <c>*</c>
    /// means "has any hyperlink" (and <c>!=*</c> means "has none"); otherwise the
    /// canonical form (url / #slide:N / #first…) is compared with =, != or *=.
    /// </summary>
    private static bool MatchesHyperlink(string? link, AttributePredicate predicate)
    {
        if (predicate.Value == "*")
        {
            return predicate.Op switch
            {
                SelectorOperator.Equals => link is not null,
                SelectorOperator.NotEquals => link is null,
                _ => throw UnsupportedOperator(predicate, "hyperlink=* / hyperlink!=* test presence"),
            };
        }

        return CompareText(link ?? string.Empty, predicate);
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
        var tables = PptxTables.Tables(slidePart);
        var smartArts = PptxSmartArt.List(slidePart);
        return new
        {
            Path = address.CanonicalSlidePath,
            Index = address.SlideIndex,
            Background = PptxDoc.BackgroundHex(slidePart),
            BackgroundKind = PptxDoc.BackgroundKind(slidePart),
            Transition = transition?.Kind,
            TransitionDuration = transition?.Duration,
            Footer = PptxFooters.SlideState(slidePart),
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
            Tables = tables.Count == 0
                ? null
                : tables.Select(t => (object)Units.Inv($"{address.CanonicalSlidePath}/table[{t.Index}]")).ToList(),
            SmartArt = smartArts.Count == 0
                ? null
                : smartArts.Select(d => (object)Units.Inv($"{address.CanonicalSlidePath}/smartart[{d.Index}]")).ToList(),
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
        var tableIndex = PptxTables.IndexOf(slidePart, view.Element);
        var smartArtIndex = PptxSmartArt.IndexOf(slidePart, view.Element);
        return new
        {
            Path = view.CanonicalPath(address.SlideIndex),
            OrdinalPath = view.OrdinalPath(address.SlideIndex),
            ChartPath = chartIndex is { } ci ? Units.Inv($"{address.CanonicalSlidePath}/chart[{ci}]") : null,
            TablePath = tableIndex is { } ti ? Units.Inv($"{address.CanonicalSlidePath}/table[{ti}]") : null,
            SmartArtPath = smartArtIndex is { } si ? Units.Inv($"{address.CanonicalSlidePath}/smartart[{si}]") : null,
            Slide = address.SlideIndex,
            Id = view.Id,
            Ordinal = view.Ordinal,
            ZIndex = view.Ordinal, // paint order: 1 = painted first (bottom)
            Kind = view.Kind,
            Name = view.Name,
            Placeholder = PptxDoc.PlaceholderType(view.Element),
            AltText = PptxDoc.AltText(view.Element),
            AltTitle = PptxDoc.AltTitle(view.Element),
            Geometry = PptxDoc.GeometryToken(view.Element),
            Flip = PptxDoc.FlipToken(view.Element),
            Text = PptxDoc.ShapeText(view.Element),
            X = geometry is { } g1 ? Units.EmuToCm(g1.X) : (double?)null,
            Y = geometry is { } g2 ? Units.EmuToCm(g2.Y) : (double?)null,
            W = geometry is { } g3 ? Units.EmuToCm(g3.Cx) : (double?)null,
            H = geometry is { } g4 ? Units.EmuToCm(g4.Cy) : (double?)null,
            Fill = PptxDoc.FillHex(view.Element),
            LineColor = PptxDoc.LineHex(view.Element),
            Hyperlink = PptxHyperlinks.Resolve(presentation, slidePart, view.Element),
            Font = firstParagraph is null ? null : FontInfo(firstParagraph),
            Autofit = PptxDoc.Autofit(view.Element),
            Chart = PptxCharts.Summary(slidePart, view.Element),
            Effects = PptxEffects.Read(view.Element),
            Media = view.Element is P.Picture picture && PptxMedia.MediaKindOf(picture) is { } mediaKind
                ? new { Kind = mediaKind, Path = Units.Inv($"{address.CanonicalSlidePath}/media[@id={view.Id}]") }
                : null,
            Connector = view.Element is P.ConnectionShape connector ? ConnectorInfo(connector) : null,
            ActionButton = PptxActionButtons.Summary(presentation, slidePart, view.Element),
        };
    }

    /// <summary>The endpoints a connector references (the from/to shape ids), when it has cxn refs.</summary>
    private static object? ConnectorInfo(P.ConnectionShape connector)
    {
        var (startId, endId) = PptxConnectors.Endpoints(connector);
        if (startId is null && endId is null)
        {
            return null;
        }

        return new { From = startId, To = endId };
    }

    /// <summary>
    /// The `get` projection for a group path: identity, geometry, and either the group's
    /// children (/slide[i]/group[@id=N]) or the addressed child's shape detail
    /// (/slide[i]/group[@id=N]/shape[...]).
    /// </summary>
    public static object GroupDetail(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var tree = PptxDoc.RequireShapeTree(slidePart);
        var group = PptxGroups.ResolveGroup(tree, address);

        if (address.HasShape)
        {
            var child = PptxGroups.ResolveChild(group, address);
            var childGeometry = PptxDoc.Geometry(child.Element);
            var childParagraph = (child.Element as P.Shape)?.TextBody?.Elements<A.Paragraph>().FirstOrDefault();
            return new
            {
                Path = Units.Inv($"{address.CanonicalGroupPath}/shape[@id={child.Id}]"),
                GroupPath = address.CanonicalGroupPath,
                Slide = address.SlideIndex,
                Id = child.Id,
                Ordinal = child.Ordinal,
                Kind = child.Kind,
                Name = child.Name,
                AltText = PptxDoc.AltText(child.Element),
                Geometry = PptxDoc.GeometryToken(child.Element),
                Text = PptxDoc.ShapeText(child.Element),
                X = childGeometry is { } cg1 ? Units.EmuToCm(cg1.X) : (double?)null,
                Y = childGeometry is { } cg2 ? Units.EmuToCm(cg2.Y) : (double?)null,
                W = childGeometry is { } cg3 ? Units.EmuToCm(cg3.Cx) : (double?)null,
                H = childGeometry is { } cg4 ? Units.EmuToCm(cg4.Cy) : (double?)null,
                Fill = PptxDoc.FillHex(child.Element),
                Font = childParagraph is null ? null : FontInfo(childParagraph),
                Autofit = PptxDoc.Autofit(child.Element),
            };
        }

        var geometry = PptxDoc.Geometry(group);
        var children = PptxGroups.Children(group);
        var nv = group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties;
        return new
        {
            Path = address.CanonicalGroupPath,
            Slide = address.SlideIndex,
            Id = address.GroupId,
            Kind = "group",
            Name = nv?.Name?.Value,
            AltText = PptxDoc.AltText(group),
            AltTitle = PptxDoc.AltTitle(group),
            X = geometry is { } g1 ? Units.EmuToCm(g1.X) : (double?)null,
            Y = geometry is { } g2 ? Units.EmuToCm(g2.Y) : (double?)null,
            W = geometry is { } g3 ? Units.EmuToCm(g3.Cx) : (double?)null,
            H = geometry is { } g4 ? Units.EmuToCm(g4.Cy) : (double?)null,
            ChildCount = children.Count,
            Children = children.Select(c => new
            {
                Path = Units.Inv($"{address.CanonicalGroupPath}/shape[@id={c.Id}]"),
                Id = c.Id,
                Kind = c.Kind,
                Name = c.Name,
                Text = Snippet(PptxDoc.ShapeText(c.Element)),
            }).ToList<object>(),
        };
    }

    /// <summary>The `get` projection for /master[m]: identity, its layouts (with usage) and its shapes.</summary>
    public static object MasterDetail(PresentationPart presentation, PptxAddress address)
    {
        var masterPart = PptxDoc.ResolveMaster(presentation, address.MasterIndex, address.Raw);
        var containerPath = address.CanonicalContainerPath;
        var shapes = PptxDoc.Shapes(PptxDoc.RequireShapeTree(masterPart));
        var layouts = PptxDoc.Layouts(masterPart);
        var fontScheme = masterPart.ThemePart?.Theme?.ThemeElements?.FontScheme;
        return new
        {
            Path = containerPath,
            Index = address.MasterIndex,
            Kind = "master",
            Theme = masterPart.ThemePart?.Theme?.Name?.Value,
            MajorFont = fontScheme?.MajorFont?.LatinFont?.Typeface?.Value,
            MinorFont = fontScheme?.MinorFont?.LatinFont?.Typeface?.Value,
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
        var placeholders = shapes
            .Where(s => PptxDoc.PlaceholderType(s.Element) is not null)
            .Select(s =>
            {
                var geometry = PptxDoc.Geometry(s.Element);
                return (object)new
                {
                    Path = s.CanonicalPathIn(containerPath),
                    Type = PptxDoc.PlaceholderType(s.Element),
                    Name = s.Name,
                    X = geometry is { } g1 ? Units.EmuToCm(g1.X) : (double?)null,
                    Y = geometry is { } g2 ? Units.EmuToCm(g2.Y) : (double?)null,
                    W = geometry is { } g3 ? Units.EmuToCm(g3.Cx) : (double?)null,
                    H = geometry is { } g4 ? Units.EmuToCm(g4.Cy) : (double?)null,
                };
            })
            .ToList();
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
            Placeholders = placeholders.Count == 0 ? null : placeholders,
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
