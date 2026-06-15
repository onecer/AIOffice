using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>One zoom navigation object on a slide: its 1-based index, kind and resolved target label.</summary>
internal sealed record ZoomView(int Index, P.GraphicFrame Frame, uint Id, string Kind, string? Target);

/// <summary>
/// Zoom links (v1.4.0, modern PowerPoint navigation): a slide/section/summary zoom is a
/// <c>p:graphicFrame</c> whose <c>a:graphicData</c> carries the 2018 p159 zoom payload
/// (<c>p159:slideZoom</c> / <c>p159:sectionZoom</c> / <c>p159:summaryZoom</c>). A slide zoom
/// jumps to a slide; a section zoom jumps to a section's first slide (and records its 1-based
/// section index); a summary zoom is one frame referencing each section's first slide. The
/// reference is a real <c>r:embed</c> slide relationship on the host slide part, so it opens in
/// PowerPoint 2019+ and stays OpenXmlValidator-clean. The host frame's shape id is the zoom's
/// stable identity; zooms are addressed by 1-based ordinal among the slide's zoom frames
/// (<c>/slide[i]/zoom[k]</c>).
/// </summary>
internal static class PptxZoom
{
    /// <summary>The 2018/8 main namespace PowerPoint uses for the zoom graphic payload.</summary>
    private const string ZoomNs = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";

    /// <summary>The relationships namespace the zoom's r:embed attribute lives in.</summary>
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly XNamespace P159 = ZoomNs;
    private static readonly XNamespace R = RelNs;
    private static readonly XNamespace AOmml = DrawingNs;

    /// <summary>The zoom kinds add zoom understands.</summary>
    public static readonly IReadOnlyList<string> Kinds = ["slide", "section", "summary"];

    private static readonly IReadOnlyList<string> AddProps = ["kind", "target", "x", "y", "w", "h", "name"];

    // ----- add -------------------------------------------------------------------

    /// <summary>
    /// add /slide[i] {type:zoom}: builds a slide/section/summary zoom graphic frame on the slide and
    /// returns its canonical /slide[i]/zoom[k] path. The target slide(s) are referenced by a real slide
    /// relationship on the host slide part, so the zoom opens in PowerPoint 2019+.
    /// </summary>
    public static string Add(PresentationPart presentation, PptxAddress address, JsonObject? props)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown zoom prop '{key}'.",
                    "Zoom props: kind (slide/section/summary), target (\"slide N\" for slide, the section name for " +
                    "section; summary takes none), x, y, w, h, name.",
                    candidates: AddProps);
            }
        }

        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var kind = ParseKind(props);

        var x = Length(props, "x", Units.CmToEmu(2));
        var y = Length(props, "y", Units.CmToEmu(2));
        var w = Length(props, "w", Units.CmToEmu(8));
        var h = Length(props, "h", Units.CmToEmu(5));

        var tree = PptxDoc.RequireShapeTree(slidePart);
        var id = PptxDoc.NextShapeId(tree);

        var payload = kind switch
        {
            "slide" => BuildSlideZoom(presentation, slidePart, props, address),
            "section" => BuildSectionZoom(presentation, slidePart, props, address),
            _ => BuildSummaryZoom(presentation, slidePart, address),
        };

        var name = props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null
            ? J.ScalarText(nameNode)
            : DefaultName(kind, id);

        var frameXml = BuildFrameXml(id, name, x, y, w, h, payload);
        tree.Append(new P.GraphicFrame(frameXml));

        var index = List(slidePart).Count;
        return Units.Inv($"/slide[{address.SlideIndex}]/zoom[{index}]");
    }

    private static string ParseKind(JsonObject props)
    {
        if (!props.TryGetPropertyValue("kind", out var node) || node is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add zoom needs a 'kind'.",
                "Use kind slide (jump to a slide), section (jump to a section), or summary " +
                "(one zoom referencing every section's first slide).",
                candidates: Kinds);
        }

        var token = J.ScalarText(node).Trim().ToLowerInvariant();
        if (!Kinds.Contains(token, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Zoom kind '{token}' is not supported.",
                "Supported zoom kinds: slide, section, summary.",
                candidates: Kinds);
        }

        return token;
    }

    /// <summary>The slide-zoom payload XElement (p159:slideZoom r:embed=relId), wiring a slide relationship.</summary>
    private static XElement BuildSlideZoom(PresentationPart presentation, SlidePart slidePart, JsonObject props, PptxAddress address)
    {
        var targetIndex = ParseSlideTarget(props, presentation, address);
        var targetPart = PptxDoc.Slides(presentation)[targetIndex - 1].Part;
        var relId = slidePart.CreateRelationshipToPart(targetPart);

        return new XElement(
            P159 + "slideZoom",
            new XAttribute(R + "embed", relId),
            new XElement(P159 + "slideZoomObjBg", new XElement(AOmml + "noFill")));
    }

    /// <summary>The section-zoom payload (p159:sectionZoom r:embed=firstSlide sectionIdx=N).</summary>
    private static XElement BuildSectionZoom(PresentationPart presentation, SlidePart slidePart, JsonObject props, PptxAddress address)
    {
        var (section, firstSlideIndex) = ParseSectionTarget(props, presentation);
        var targetPart = PptxDoc.Slides(presentation)[firstSlideIndex - 1].Part;
        var relId = slidePart.CreateRelationshipToPart(targetPart);

        return new XElement(
            P159 + "sectionZoom",
            new XAttribute(R + "embed", relId),
            new XAttribute("sectionIdx", section.Index.ToString(CultureInfo.InvariantCulture)),
            new XElement(P159 + "sectionZoomObjBg", new XElement(AOmml + "noFill")));
    }

    /// <summary>The summary-zoom payload: one p159:section child per deck section, each a first-slide reference.</summary>
    private static XElement BuildSummaryZoom(PresentationPart presentation, SlidePart slidePart, PptxAddress address)
    {
        var sections = PptxSections.List(presentation).Where(s => s.Slides.Count > 0).ToList();
        if (sections.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "A summary zoom needs sections, but the deck has none.",
                "Add sections first ({\"op\":\"add\",\"path\":\"/\",\"type\":\"section\",\"props\":{\"name\":\"Intro\"}}), " +
                "then add a summary zoom; or use kind slide/section.");
        }

        var summary = new XElement(P159 + "summaryZoom");
        foreach (var section in sections)
        {
            var firstSlideIndex = section.Slides.Min();
            var targetPart = PptxDoc.Slides(presentation)[firstSlideIndex - 1].Part;
            var relId = slidePart.CreateRelationshipToPart(targetPart);
            summary.Add(new XElement(
                P159 + "section",
                new XAttribute(R + "embed", relId),
                new XAttribute("sectionIdx", section.Index.ToString(CultureInfo.InvariantCulture)),
                new XElement(P159 + "sectionZoomObjBg", new XElement(AOmml + "noFill"))));
        }

        return summary;
    }

    /// <summary>Parses props.target ("slide N") into the 1-based target slide index, validating it.</summary>
    private static int ParseSlideTarget(JsonObject props, PresentationPart presentation, PptxAddress address)
    {
        if (!props.TryGetPropertyValue("target", out var node) || node is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "A slide zoom needs a 'target' slide.",
                "Pass {\"kind\":\"slide\",\"target\":\"slide 3\"} (or a path like \"/slide[3]\").");
        }

        var raw = J.ScalarText(node).Trim();
        var digits = raw;
        if (raw.StartsWith("slide", StringComparison.OrdinalIgnoreCase))
        {
            digits = raw["slide".Length..].Trim().TrimStart('[').TrimEnd(']').Trim();
        }
        else if (raw.StartsWith("/slide[", StringComparison.OrdinalIgnoreCase))
        {
            digits = raw["/slide[".Length..].TrimEnd(']').Trim();
        }

        if (!int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out var index))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"A slide zoom target must name a slide like \"slide 3\"; got {node.ToJsonString()}.",
                "Pass {\"target\":\"slide 3\"} or {\"target\":\"/slide[3]\"}.");
        }

        var slideCount = PptxDoc.Slides(presentation).Count;
        if (index < 1 || index > slideCount)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                Units.Inv($"Zoom target slide {index} is out of range 1..{slideCount}."),
                "Point the zoom at an existing slide; run 'aioffice read <file> --view outline' to list slides.",
                candidates: [.. Enumerable.Range(1, Math.Min(slideCount, 10)).Select(i => Units.Inv($"slide {i}"))]);
        }

        return index;
    }

    /// <summary>Resolves props.target (a section name) to its section view + first slide index.</summary>
    private static (SectionView Section, int FirstSlideIndex) ParseSectionTarget(JsonObject props, PresentationPart presentation)
    {
        if (!props.TryGetPropertyValue("target", out var node) || node is null || J.ScalarText(node).Trim().Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "A section zoom needs a 'target' section name.",
                "Pass {\"kind\":\"section\",\"target\":\"Appendix\"}; run 'aioffice read <file> --view outline' to list sections.");
        }

        var name = J.ScalarText(node).Trim();
        var sections = PptxSections.List(presentation);
        var match = sections.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.Ordinal));
        if (match is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No section named '{name}' in the deck.",
                sections.Count > 0
                    ? "Use one of the deck's section names; run 'aioffice read <file> --view outline' to list them."
                    : "The deck has no sections; add one first ({\"op\":\"add\",\"path\":\"/\",\"type\":\"section\",\"props\":{\"name\":\"Intro\"}}).",
                candidates: [.. sections.Take(10).Select(s => s.Name)]);
        }

        if (match.Slides.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Section '{name}' has no slides to zoom to.",
                "A section zoom jumps to the section's first slide; assign slides to the section first.");
        }

        return (match, match.Slides.Min());
    }

    /// <summary>The raw p:graphicFrame XML hosting the zoom payload (built as text so the p159 payload rides verbatim).</summary>
    private static string BuildFrameXml(uint id, string name, long x, long y, long w, long h, XElement payload)
    {
        var graphicData = new XElement(
            AOmml + "graphicData",
            new XAttribute("uri", ZoomNs),
            payload);

        var frame = new XElement(
            XName.Get("graphicFrame", "http://schemas.openxmlformats.org/presentationml/2006/main"),
            new XAttribute(XNamespace.Xmlns + "p", "http://schemas.openxmlformats.org/presentationml/2006/main"),
            new XAttribute(XNamespace.Xmlns + "a", DrawingNs),
            new XAttribute(XNamespace.Xmlns + "r", RelNs),
            new XAttribute(XNamespace.Xmlns + "p159", ZoomNs),
            NvGraphicFrameProps(id, name),
            Xfrm(x, y, w, h),
            new XElement(AOmml + "graphic", graphicData));

        return frame.ToString(SaveOptions.DisableFormatting);
    }

    private static XElement NvGraphicFrameProps(uint id, string name)
    {
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        return new XElement(
            p + "nvGraphicFramePr",
            new XElement(p + "cNvPr", new XAttribute("id", id.ToString(CultureInfo.InvariantCulture)), new XAttribute("name", name)),
            new XElement(p + "cNvGraphicFramePr"),
            new XElement(p + "nvPr"));
    }

    private static XElement Xfrm(long x, long y, long w, long h)
    {
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        return new XElement(
            p + "xfrm",
            new XElement(AOmml + "off", new XAttribute("x", x.ToString(CultureInfo.InvariantCulture)), new XAttribute("y", y.ToString(CultureInfo.InvariantCulture))),
            new XElement(AOmml + "ext", new XAttribute("cx", w.ToString(CultureInfo.InvariantCulture)), new XAttribute("cy", h.ToString(CultureInfo.InvariantCulture))));
    }

    private static string DefaultName(string kind, uint id) => kind switch
    {
        "slide" => Units.Inv($"Slide Zoom {id}"),
        "section" => Units.Inv($"Section Zoom {id}"),
        _ => Units.Inv($"Summary Zoom {id}"),
    };

    private static long Length(JsonObject props, string key, long fallback) =>
        props.TryGetPropertyValue(key, out var node) && node is not null ? Units.ParseLengthEmu(key, node) : fallback;

    // ----- enumerate / resolve ---------------------------------------------------

    /// <summary>Every zoom graphic frame on a slide in paint order, 1-based.</summary>
    public static List<ZoomView> List(SlidePart slidePart)
    {
        var result = new List<ZoomView>();
        foreach (var view in PptxDoc.Shapes(slidePart))
        {
            if (view.Element is P.GraphicFrame frame && ZoomPayload(frame) is { } payload)
            {
                var (kind, target) = Describe(slidePart, payload);
                result.Add(new ZoomView(result.Count + 1, frame, view.Id, kind, target));
            }
        }

        return result;
    }

    /// <summary>The zoom view for a slide element when it is a zoom graphic frame, else null (for the renderer).</summary>
    public static ZoomView? ZoomViewOf(SlidePart slidePart, DocumentFormat.OpenXml.OpenXmlElement element) =>
        element is P.GraphicFrame frame && ZoomPayload(frame) is not null
            ? List(slidePart).FirstOrDefault(z => ReferenceEquals(z.Frame, frame))
            : null;

    /// <summary>The p159 zoom payload element a frame hosts (slideZoom/sectionZoom/summaryZoom), or null.</summary>
    private static XElement? ZoomPayload(P.GraphicFrame frame)
    {
        var graphicData = frame.Graphic?.GraphicData;
        if (graphicData?.Uri?.Value != ZoomNs)
        {
            return null;
        }

        var element = XElement.Parse(graphicData.OuterXml);
        return element.Elements().FirstOrDefault(e =>
            e.Name == P159 + "slideZoom" || e.Name == P159 + "sectionZoom" || e.Name == P159 + "summaryZoom");
    }

    /// <summary>The kind + resolved target label for a zoom payload (best-effort; foreign decks stay truthful).</summary>
    private static (string Kind, string? Target) Describe(SlidePart slidePart, XElement payload)
    {
        if (payload.Name == P159 + "slideZoom")
        {
            var relId = payload.Attribute(R + "embed")?.Value;
            return ("slide", SlideLabelFor(slidePart, relId));
        }

        if (payload.Name == P159 + "sectionZoom")
        {
            var sectionIdx = payload.Attribute("sectionIdx")?.Value;
            return ("section", sectionIdx is null ? null : Units.Inv($"section {sectionIdx}"));
        }

        // summaryZoom: report how many sections it references.
        var count = payload.Elements(P159 + "section").Count();
        return ("summary", Units.Inv($"{count} section(s)"));
    }

    /// <summary>The "slide N" label for a zoom's r:embed relationship, resolved through the slide part's rels.</summary>
    private static string? SlideLabelFor(SlidePart slidePart, string? relId)
    {
        if (relId is null)
        {
            return null;
        }

        var presentation = slidePart.OpenXmlPackage.GetPartsOfType<PresentationPart>().FirstOrDefault();
        if (presentation is null || !slidePart.TryGetPartById(relId, out var part) || part is not SlidePart target)
        {
            return null;
        }

        var slides = PptxDoc.Slides(presentation);
        for (var i = 0; i < slides.Count; i++)
        {
            if (ReferenceEquals(slides[i].Part, target))
            {
                return Units.Inv($"slide {i + 1}");
            }
        }

        return null;
    }

    /// <summary>Resolves /slide[i]/zoom[k] or throws invalid_path with candidates.</summary>
    public static ZoomView Resolve(SlidePart slidePart, PptxAddress address)
    {
        var zooms = List(slidePart);
        var index = address.ZoomIndex!.Value;
        if (index >= 1 && index <= zooms.Count)
        {
            return zooms[index - 1];
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            Units.Inv($"No zoom {index} on slide {address.SlideIndex}; it has {zooms.Count} zoom(s)."),
            zooms.Count > 0
                ? "Zoom indices are 1-based per slide; run 'aioffice read <file> --view structure' to list them."
                : "Add one first: {\"op\":\"add\",\"path\":\"" + address.CanonicalSlidePath +
                  "\",\"type\":\"zoom\",\"props\":{\"kind\":\"slide\",\"target\":\"slide 2\"}}.",
            candidates: [.. zooms.Take(10).Select(z => Units.Inv($"{address.CanonicalSlidePath}/zoom[{z.Index}]"))]);
    }

    // ----- get -------------------------------------------------------------------

    /// <summary>The `get` projection for /slide[i]/zoom[k].</summary>
    public static object Detail(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var view = Resolve(slidePart, address);
        return Project(view, address.SlideIndex);
    }

    /// <summary>The shared zoom row (get and read --view structure).</summary>
    public static object Project(ZoomView view, int slideIndex) => new
    {
        Path = Units.Inv($"/slide[{slideIndex}]/zoom[{view.Index}]"),
        Slide = slideIndex,
        Index = view.Index,
        Id = view.Id,
        Kind = view.Kind,
        Target = view.Target,
    };

    // ----- remove ----------------------------------------------------------------

    /// <summary>
    /// remove /slide[i]/zoom[k]: drops the zoom frame, then prunes any slide relationship that only
    /// this zoom referenced (a slide-to-slide rel to a target the slide does not otherwise link). The
    /// target slide part itself is never removed — it is still referenced by the deck's slide list, so
    /// DeletePart only severs the host slide's relationship to it.
    /// </summary>
    public static string Remove(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var view = Resolve(slidePart, address);

        var relIds = ZoomRelationshipIds(view.Frame).ToList();
        view.Frame.Remove();

        // Drop each rel only if no surviving zoom on the slide still references it (so two zooms to
        // the same slide are safe). The target slide survives — it is held by the presentation's
        // slide list, so DeletePart just removes the host slide's now-dangling relationship.
        var stillReferenced = List(slidePart)
            .SelectMany(z => ZoomRelationshipIds(z.Frame))
            .ToHashSet(StringComparer.Ordinal);
        foreach (var relId in relIds)
        {
            if (!stillReferenced.Contains(relId) && slidePart.TryGetPartById(relId, out _))
            {
                slidePart.DeletePart(relId);
            }
        }

        return address.CanonicalZoomPath;
    }

    /// <summary>Every r:embed relationship id the zoom payload references (one for slide/section, many for summary).</summary>
    private static IEnumerable<string> ZoomRelationshipIds(P.GraphicFrame frame)
    {
        if (ZoomPayload(frame) is not { } payload)
        {
            yield break;
        }

        if (payload.Attribute(R + "embed")?.Value is { } direct)
        {
            yield return direct;
        }

        foreach (var section in payload.Elements(P159 + "section"))
        {
            if (section.Attribute(R + "embed")?.Value is { } id)
            {
                yield return id;
            }
        }
    }

    // ----- render ----------------------------------------------------------------

    /// <summary>
    /// Draws a zoom navigation placeholder: a thumbnail rect + a label naming the kind and target,
    /// wrapped (by the caller) in the data-aio-path group. Honest stand-in, not a slide thumbnail.
    /// </summary>
    public static void AppendPlaceholder(StringBuilder svg, ZoomView view, double x, double y, double w, double h)
    {
        svg.Append(Units.Inv($"    <rect class=\"aio-zoom\" x=\"{x:0.#}\" y=\"{y:0.#}\" width=\"{w:0.#}\" height=\"{h:0.#}\" "));
        svg.Append("fill=\"#f1f5f9\" stroke=\"#475569\" stroke-dasharray=\"5 3\"/>\n");

        // A small inner "thumbnail" rect reads as a zoomed slide preview.
        var pad = Math.Min(w, h) * 0.18;
        var thumbW = Math.Max(w - (pad * 2), 4);
        var thumbH = Math.Max(h - (pad * 2) - 12, 4);
        svg.Append(Units.Inv($"    <rect class=\"aio-zoom-thumb\" x=\"{x + pad:0.#}\" y=\"{y + pad:0.#}\" width=\"{thumbW:0.#}\" height=\"{thumbH:0.#}\" "));
        svg.Append("fill=\"#ffffff\" stroke=\"#94a3b8\"/>\n");

        var label = view.Target is null ? Units.Inv($"[{view.Kind} zoom]") : Units.Inv($"[{view.Kind} zoom] {view.Target}");
        svg.Append(Units.Inv($"    <text x=\"{x + (w / 2):0.#}\" y=\"{y + h - 6:0.#}\" font-size=\"11\" "));
        svg.Append(Units.Inv($"text-anchor=\"middle\" fill=\"#475569\">{PptxRenderer.Escape(label)}</text>\n"));
    }
}
