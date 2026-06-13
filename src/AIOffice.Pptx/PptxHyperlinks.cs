using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// M8 shape hyperlinks / actions. A shape's click action lives on its
/// <c>p:cNvPr</c> as an <c>a:hlinkClick</c>:
/// <list type="bullet">
/// <item>external url — <c>{"hyperlink":"https://…"}</c> → a relationship to the
///   url + <c>a:hlinkClick r:id="…"</c>;</item>
/// <item>slide jump — <c>{"hyperlink":"#slide:4"}</c> → a relationship to the
///   target slide + <c>a:hlinkClick action="ppaction://hlinksldjump"</c>;</item>
/// <item>show action — <c>{"hyperlink":"#first|#last|#next|#prev|#end"}</c> →
///   <c>a:hlinkClick action="ppaction://hlinkshowjump?jump=…"</c> (no
///   relationship).</item>
/// </list>
/// <c>{"hyperlink":""}</c> clears it. <see cref="Resolve"/> reports the canonical
/// form (url / #slide:N / #first…) for get and query.
/// </summary>
internal static class PptxHyperlinks
{
    /// <summary>The ppaction verbs aioffice exposes, keyed by their canonical shorthand.</summary>
    private static readonly IReadOnlyDictionary<string, string> ShowActions =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["#first"] = "firstslide",
            ["#last"] = "lastslide",
            ["#next"] = "nextslide",
            ["#prev"] = "previousslide",
            ["#end"] = "endshow",
        };

    private const string SlideJumpAction = "ppaction://hlinksldjump";
    private const string ShowJumpActionPrefix = "ppaction://hlinkshowjump?jump=";

    /// <summary>True when a set op's props target the shape hyperlink (so the editor routes here).</summary>
    public static bool Handles(JsonObject props) =>
        props.ContainsKey("hyperlink") || props.ContainsKey("linkText");

    /// <summary>
    /// Applies a hyperlink/linkText set to the shape. <c>hyperlink</c> sets the
    /// shape-level click action; <c>linkText</c> wraps the shape's runs in a
    /// run-level link to the same target (optional sugar). Returns nothing; the
    /// caller already knows the canonical shape path.
    /// </summary>
    public static void Apply(PresentationPart presentation, SlidePart slidePart, ShapeView view, JsonObject props)
    {
        var nonVisual = PptxDoc.NonVisualProps(view.Element)
            ?? throw Corrupt("the shape has no p:cNvPr to carry a hyperlink");

        // Clear any existing click action + its dangling relationship first, so a
        // re-set never leaves an orphan relationship behind.
        var existingRelId = nonVisual.GetFirstChild<A.HyperlinkOnClick>()?.Id?.Value;
        foreach (var existing in nonVisual.Elements<A.HyperlinkOnClick>().ToList())
        {
            existing.Remove();
        }

        if (existingRelId is { Length: > 0 })
        {
            DeleteSlideRelationshipIfUnused(slidePart, existingRelId);
        }

        if (props.TryGetPropertyValue("hyperlink", out var hyperlinkNode))
        {
            var raw = hyperlinkNode is null ? string.Empty : J.ScalarText(hyperlinkNode).Trim();
            if (raw.Length > 0)
            {
                var hlink = BuildHyperlink(presentation, slidePart, raw);
                // a:hlinkClick precedes a:hlinkHover/a:extLst in the cNvPr schema:
                // insert after the last preceding child, or first when there is none.
                if (LastBeforeHlink(nonVisual) is { } anchor)
                {
                    nonVisual.InsertAfter(hlink, anchor);
                }
                else
                {
                    nonVisual.InsertAt(hlink, 0);
                }
            }
        }

        // Optional run-level mirror: wrap each run's text in the same target.
        if (props.TryGetPropertyValue("linkText", out var linkTextNode) && view.Element is P.Shape shape)
        {
            ApplyRunHyperlink(presentation, slidePart, shape, linkTextNode is null ? string.Empty : J.ScalarText(linkTextNode).Trim());
        }
    }

    /// <summary>Builds the a:hlinkClick for a raw target (url / #slide:N / #show), creating any relationship.</summary>
    private static A.HyperlinkOnClick BuildHyperlink(PresentationPart presentation, SlidePart slidePart, string raw)
    {
        // Show actions: first/last/next/prev/end (no relationship needed).
        if (ShowActions.TryGetValue(raw, out var jump))
        {
            return new A.HyperlinkOnClick { Id = string.Empty, Action = ShowJumpActionPrefix + jump };
        }

        // Slide jump: #slide:N references the Nth slide by a relationship. The
        // single-arg overload infers the standard slide relationship type from the
        // target part and returns the generated, schema-valid relationship id.
        if (raw.StartsWith("#slide:", StringComparison.Ordinal))
        {
            var targetIndex = ParseSlideTarget(raw, presentation);
            var slides = PptxDoc.Slides(presentation);
            var targetPart = slides[targetIndex - 1].Part;
            var relId = slidePart.CreateRelationshipToPart(targetPart);
            return new A.HyperlinkOnClick { Id = relId, Action = SlideJumpAction };
        }

        // A bare '#...' that is not a known action is a mistake worth flagging.
        if (raw.StartsWith('#'))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown hyperlink action '{raw}'.",
                "Use #slide:N to jump to a slide, or #first/#last/#next/#prev/#end for a show action; " +
                "an external link is an absolute http(s)/mailto url.",
                candidates: [.. ShowActions.Keys, "#slide:N"]);
        }

        // External url.
        if (!Uri.TryCreate(raw, UriKind.Absolute, out var uri) ||
            uri.Scheme is not ("http" or "https" or "mailto"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{raw}' is not a supported hyperlink target.",
                "External links use absolute http(s) or mailto urls; for in-deck navigation use " +
                "#slide:N or #first/#last/#next/#prev/#end.");
        }

        var externalRelId = slidePart.AddHyperlinkRelationship(uri, isExternal: true).Id;
        return new A.HyperlinkOnClick { Id = externalRelId };
    }

    /// <summary>Parses #slide:N and validates it against the deck's slide count.</summary>
    private static int ParseSlideTarget(string raw, PresentationPart presentation)
    {
        var rest = raw["#slide:".Length..];
        if (!int.TryParse(rest, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var index) || index < 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{raw}' is not a valid slide-jump target.",
                "Use #slide:N with a 1-based slide number, e.g. #slide:4.");
        }

        var count = PptxDoc.Slides(presentation).Count;
        if (index > count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"Slide-jump target {index} is out of range; the deck has {count} slide(s)."),
                Units.Inv($"Use #slide:N with N in 1..{count}."));
        }

        return index;
    }

    /// <summary>Wraps every run of a shape's text body in an a:hlinkClick (or clears them when target is empty).</summary>
    private static void ApplyRunHyperlink(PresentationPart presentation, SlidePart slidePart, P.Shape shape, string raw)
    {
        foreach (var run in (shape.TextBody?.Descendants<A.Run>() ?? []).ToList())
        {
            var rPr = run.RunProperties ??= new A.RunProperties();
            foreach (var existing in rPr.Elements<A.HyperlinkOnClick>().ToList())
            {
                if (existing.Id?.Value is { Length: > 0 } relId)
                {
                    DeleteSlideRelationshipIfUnused(slidePart, relId);
                }

                existing.Remove();
            }

            if (raw.Length > 0)
            {
                // a:hlinkClick is the first child of a:rPr in the schema order.
                rPr.InsertAt(BuildHyperlink(presentation, slidePart, raw), 0);
            }
        }
    }

    /// <summary>The resolved, canonical hyperlink form of a shape (url / #slide:N / #first…), or null.</summary>
    public static string? Resolve(PresentationPart presentation, SlidePart slidePart, OpenXmlCompositeElement element)
    {
        var hlink = PptxDoc.NonVisualProps(element)?.GetFirstChild<A.HyperlinkOnClick>();
        return hlink is null ? null : ResolveHlink(presentation, slidePart, hlink);
    }

    /// <summary>Resolves one a:hlinkClick to its canonical aioffice form.</summary>
    public static string? ResolveHlink(PresentationPart presentation, SlidePart slidePart, A.HyperlinkOnClick hlink)
    {
        var action = hlink.Action?.Value;

        // Show action: ppaction://hlinkshowjump?jump=firstslide → #first.
        if (action is { } a && a.StartsWith(ShowJumpActionPrefix, StringComparison.Ordinal))
        {
            var jump = a[ShowJumpActionPrefix.Length..];
            foreach (var (shorthand, verb) in ShowActions)
            {
                if (string.Equals(verb, jump, StringComparison.Ordinal))
                {
                    return shorthand;
                }
            }

            return a; // an unknown show verb passes through raw
        }

        // Slide jump: resolve the relationship target (an internal part
        // relationship) back to a 1-based slide index.
        if (action is SlideJumpAction && hlink.Id?.Value is { Length: > 0 } jumpRelId)
        {
            if (slidePart.Parts.FirstOrDefault(p => p.RelationshipId == jumpRelId).OpenXmlPart is SlidePart targetPart)
            {
                var slides = PptxDoc.Slides(presentation);
                for (var i = 0; i < slides.Count; i++)
                {
                    if (slides[i].Part.Uri == targetPart.Uri)
                    {
                        return Units.Inv($"#slide:{i + 1}");
                    }
                }
            }

            return "#slide:?";
        }

        // External url.
        if (hlink.Id?.Value is { Length: > 0 } relId)
        {
            var rel = slidePart.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
            return rel?.Uri.OriginalString;
        }

        return null;
    }

    /// <summary>
    /// Drops a shape hyperlink's relationship when no other hlink on the slide
    /// still uses it. External-url links are reference relationships; slide-jump
    /// links are internal part relationships — the relationship-id deletion
    /// overload covers both, and a part relationship leaves the target slide
    /// itself untouched (it is only un-referenced, never deleted).
    /// </summary>
    private static void DeleteSlideRelationshipIfUnused(SlidePart slidePart, string relId)
    {
        var stillUsed = slidePart.Slide?.Descendants<A.HyperlinkOnClick>()
            .Any(h => h.Id?.Value == relId) == true;
        if (stillUsed)
        {
            return;
        }

        // External-url links are reference relationships; clear them with
        // DeleteReferenceRelationship. Slide-jump links are part relationships;
        // DeletePart drops only this relationship edge — the target slide stays
        // (it is still referenced by the presentation's slide list).
        if (slidePart.HyperlinkRelationships.Any(r => r.Id == relId))
        {
            slidePart.DeleteReferenceRelationship(relId);
        }
        else if (slidePart.Parts.Any(p => p.RelationshipId == relId))
        {
            slidePart.DeletePart(relId);
        }
    }

    /// <summary>
    /// The element a:hlinkClick must be inserted after inside a:cNvPr (it follows
    /// a:extLst's siblings but the schema puts hlinkClick before hlinkHover and
    /// extLst). Inserting after the last non-hlink/extLst child keeps it valid.
    /// </summary>
    private static OpenXmlElement? LastBeforeHlink(P.NonVisualDrawingProperties nonVisual) =>
        nonVisual.Elements()
            .LastOrDefault(e => e is not (A.HyperlinkOnClick or A.HyperlinkOnHover or A.NonVisualDrawingPropertiesExtensionList));

    private static AiofficeException Corrupt(string detail) => new(
        ErrorCodes.FormatCorrupt,
        $"The presentation is malformed: {detail}.",
        "Re-export the file from PowerPoint/Keynote, or restore a snapshot.");
}
