using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Action buttons (v1.5.0): navigation shapes that build on M8 shape hyperlinks.
/// An action button is a <c>p:sp</c> with a preset action-button geometry
/// (<c>actionButtonBackPrevious</c>, <c>actionButtonHome</c>, …) plus the matching
/// click action on its <c>p:cNvPr</c> (<c>a:hlinkClick</c> with a
/// <c>ppaction://</c> verb):
/// <list type="bullet">
/// <item><c>first/last/next/prev/home/end</c> — a show-jump action (no
///   relationship);</item>
/// <item><c>slide</c> with <c>target:"slide N"</c> — a slide-jump action (a part
///   relationship to that slide);</item>
/// <item><c>url</c> with <c>target:"https://…"</c> — an external-link action.</item>
/// </list>
/// <c>add /slide[i] {type:actionButton}</c> creates one; <c>get</c> on its shape
/// path reports the resolved action; <c>remove</c> deletes it. The SVG render draws
/// the button glyph.
/// </summary>
internal static class PptxActionButtons
{
    private static readonly IReadOnlyList<string> AddPropKeys =
        ["action", "target", "x", "y", "w", "h", "label", "fill", "name"];

    /// <summary>The supported actions, in the order the suggestion lists them.</summary>
    private static readonly IReadOnlyList<string> Actions =
        ["first", "last", "next", "prev", "home", "end", "slide", "url"];

    /// <summary>The preset action-button geometry per action (slide/url reuse the blank/info button face).</summary>
    private static readonly IReadOnlyDictionary<string, A.ShapeTypeValues> Geometries =
        new Dictionary<string, A.ShapeTypeValues>(StringComparer.Ordinal)
        {
            ["first"] = A.ShapeTypeValues.ActionButtonBeginning,
            ["last"] = A.ShapeTypeValues.ActionButtonEnd,
            ["next"] = A.ShapeTypeValues.ActionButtonForwardNext,
            ["prev"] = A.ShapeTypeValues.ActionButtonBackPrevious,
            ["home"] = A.ShapeTypeValues.ActionButtonHome,
            ["end"] = A.ShapeTypeValues.ActionButtonReturn,
            ["slide"] = A.ShapeTypeValues.ActionButtonBlank,
            ["url"] = A.ShapeTypeValues.ActionButtonInformation,
        };

    /// <summary>The render glyph per action-button OOXML preset token.</summary>
    private static readonly IReadOnlyDictionary<string, string> GlyphByToken =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["actionButtonBeginning"] = "|◄",
            ["actionButtonEnd"] = "►|",
            ["actionButtonForwardNext"] = "►",
            ["actionButtonBackPrevious"] = "◄",
            ["actionButtonHome"] = "⌂",
            ["actionButtonReturn"] = "↩",
            ["actionButtonBlank"] = "→",
            ["actionButtonInformation"] = "i",
            ["actionButtonHelp"] = "?",
            ["actionButtonMovie"] = "▶",
            ["actionButtonSound"] = "♪",
            ["actionButtonDocument"] = "▤",
        };

    /// <summary>The show-jump verb each non-target action maps to (ppaction://hlinkshowjump?jump=...).</summary>
    private static readonly IReadOnlyDictionary<string, string> ShowJumps =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["first"] = "firstslide",
            ["last"] = "lastslide",
            ["next"] = "nextslide",
            ["prev"] = "previousslide",
            ["home"] = "firstslide", // "home" navigates to the first slide
            ["end"] = "lastslide",   // "return"/end button: jump to the last slide
        };

    private const string ShowJumpActionPrefix = "ppaction://hlinkshowjump?jump=";
    private const string SlideJumpAction = "ppaction://hlinksldjump";
    private const string FileAction = "ppaction://hlinkfile";

    /// <summary>Default action-button size when x/y/w/h are omitted (2cm square near the bottom-left).</summary>
    private const long DefaultSizeEmu = 720_000; // 2cm

    // ---- add ----------------------------------------------------------------

    /// <summary>
    /// add /slide[i] {type:actionButton}: appends an action-button shape with the
    /// preset geometry + navigation action and returns its canonical shape path.
    /// </summary>
    public static string Add(PresentationPart presentation, SlidePart slidePart, int slideIndex, JsonObject? props)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            if (!AddPropKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown actionButton prop '{key}'.",
                    "actionButton props: action (required), target (for slide/url), x, y, w, h, label, fill, name.",
                    candidates: AddPropKeys);
            }
        }

        if (!props.TryGetPropertyValue("action", out var actionNode) || actionNode is null || J.ScalarText(actionNode).Trim().Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add actionButton requires props.action.",
                "Use action first/last/next/prev/home/end (show navigation), slide (with target \"slide N\") " +
                "or url (with target \"https://…\"), e.g. " +
                "{\"op\":\"add\",\"path\":\"/slide[1]\",\"type\":\"actionButton\",\"props\":{\"action\":\"next\"}}.");
        }

        var action = J.ScalarText(actionNode).Trim().ToLowerInvariant();
        if (!Geometries.TryGetValue(action, out var geometry))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Action-button action '{action}' is not supported.",
                "Supported actions: first, last, next, prev, home, end, slide (target \"slide N\"), url (target \"https://…\").",
                candidates: Actions);
        }

        var tree = PptxDoc.RequireShapeTree(slidePart);
        var id = PptxDoc.NextShapeId(tree);

        var x = props.TryGetPropertyValue("x", out var xNode) ? Units.ParseLengthEmu("x", xNode) : DefaultSizeEmu;
        var y = props.TryGetPropertyValue("y", out var yNode)
            ? Units.ParseLengthEmu("y", yNode)
            : (presentation.Presentation?.SlideSize?.Cy?.Value ?? PptxFactory.SlideHeightEmu) - DefaultSizeEmu - 360_000L;
        var w = props.TryGetPropertyValue("w", out var wNode) ? Units.ParseLengthEmu("w", wNode) : DefaultSizeEmu;
        var h = props.TryGetPropertyValue("h", out var hNode) ? Units.ParseLengthEmu("h", hNode) : DefaultSizeEmu;

        var name = props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null && J.ScalarText(nameNode).Trim().Length > 0
            ? J.ScalarText(nameNode).Trim()
            : Units.Inv($"Action Button: {action} {id}");

        var hlink = BuildAction(presentation, slidePart, action, props);

        var cNvPr = new P.NonVisualDrawingProperties { Id = id, Name = name };
        cNvPr.Append(hlink);

        var shapeProperties = new P.ShapeProperties(
            new A.Transform2D(new A.Offset { X = x, Y = y }, new A.Extents { Cx = w, Cy = h }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = geometry });

        if (props.TryGetPropertyValue("fill", out var fillNode))
        {
            shapeProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = Units.ParseColorHex("fill", fillNode) }));
        }

        var body = new P.TextBody(new A.BodyProperties { Anchor = A.TextAnchoringTypeValues.Center }, new A.ListStyle());
        var label = props.TryGetPropertyValue("label", out var labelNode) && labelNode is not null
            ? J.ScalarText(labelNode)
            : string.Empty;
        if (label.Length > 0)
        {
            var paragraph = new A.Paragraph(new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Center });
            paragraph.Append(new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text(label)));
            body.Append(paragraph);
        }
        else
        {
            body.Append(new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" }));
        }

        tree.Append(new P.Shape(
            new P.NonVisualShapeProperties(cNvPr, new P.NonVisualShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
            shapeProperties,
            body));

        return Units.Inv($"/slide[{slideIndex}]/shape[@id={id}]");
    }

    /// <summary>Builds the a:hlinkClick action for an action button (show-jump / slide-jump / external url).</summary>
    private static A.HyperlinkOnClick BuildAction(PresentationPart presentation, SlidePart slidePart, string action, JsonObject props)
    {
        if (ShowJumps.TryGetValue(action, out var jump))
        {
            RejectTarget(props, action);
            return new A.HyperlinkOnClick { Id = string.Empty, Action = ShowJumpActionPrefix + jump };
        }

        var target = props.TryGetPropertyValue("target", out var targetNode) && targetNode is not null
            ? J.ScalarText(targetNode).Trim()
            : string.Empty;

        if (action == "slide")
        {
            var slideTarget = ParseSlideTarget(target, presentation);
            var slides = PptxDoc.Slides(presentation);
            var relId = slidePart.CreateRelationshipToPart(slides[slideTarget - 1].Part);
            return new A.HyperlinkOnClick { Id = relId, Action = SlideJumpAction };
        }

        // action == "url"
        if (!Uri.TryCreate(target, UriKind.Absolute, out var uri) || uri.Scheme is not ("http" or "https" or "mailto"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"action 'url' needs an absolute http(s)/mailto target; got '{(target.Length == 0 ? "(none)" : target)}'.",
                "Pass target as an absolute url, e.g. {\"action\":\"url\",\"target\":\"https://example.com\"}.");
        }

        var urlRelId = slidePart.AddHyperlinkRelationship(uri, isExternal: true).Id;
        return new A.HyperlinkOnClick { Id = urlRelId, Action = FileAction };
    }

    private static void RejectTarget(JsonObject props, string action)
    {
        if (props.ContainsKey("target"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"action '{action}' takes no target (it is a fixed show-navigation jump).",
                "Drop props.target; target only applies to action slide (\"slide N\") or url (\"https://…\").");
        }
    }

    /// <summary>Parses "slide N" (or a bare 1-based number) into a validated 1-based slide index.</summary>
    private static int ParseSlideTarget(string target, PresentationPart presentation)
    {
        var text = target.Trim();
        if (text.StartsWith("slide", StringComparison.OrdinalIgnoreCase))
        {
            text = text["slide".Length..].Trim();
        }

        if (!int.TryParse(text, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var index) || index < 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"action 'slide' needs target \"slide N\" (a 1-based slide number); got '{(target.Length == 0 ? "(none)" : target)}'.",
                "Pass target like {\"action\":\"slide\",\"target\":\"slide 3\"}.");
        }

        var count = PptxDoc.Slides(presentation).Count;
        if (index > count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"Action-button slide target {index} is out of range; the deck has {count} slide(s)."),
                Units.Inv($"Use a target slide in 1..{count}."));
        }

        return index;
    }

    // ---- get ----------------------------------------------------------------

    /// <summary>
    /// The action-button summary for a shape, when it carries an action-button preset
    /// geometry: the action token + resolved target. Null when the shape is not one.
    /// </summary>
    public static object? Summary(PresentationPart presentation, SlidePart slidePart, OpenXmlCompositeElement element)
    {
        if (element is not P.Shape shape)
        {
            return null;
        }

        var preset = shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset;
        if (preset is null || !IsActionButtonGeometry(preset))
        {
            return null;
        }

        var hlink = PptxDoc.NonVisualProps(element)?.GetFirstChild<A.HyperlinkOnClick>();
        var (action, target) = ResolveAction(presentation, slidePart, hlink);
        return new
        {
            Geometry = preset.InnerText,
            Action = action,
            Target = target,
        };
    }

    /// <summary>True when a preset geometry is one of the action-button faces.</summary>
    public static bool IsActionButtonGeometry(EnumValue<A.ShapeTypeValues>? preset) =>
        preset?.InnerText is { } token && token.StartsWith("actionButton", StringComparison.Ordinal);

    /// <summary>The render glyph for an action-button preset geometry (a small navigation symbol), or null.</summary>
    public static string? GlyphFor(EnumValue<A.ShapeTypeValues>? preset)
    {
        if (preset?.InnerText is not { } token)
        {
            return null;
        }

        return GlyphByToken.TryGetValue(token, out var glyph)
            ? glyph
            : "•"; // a foreign action-button face (e.g. one PowerPoint wrote) still gets a glyph
    }

    /// <summary>Resolves an action button's a:hlinkClick back to its (action, target) form.</summary>
    private static (string Action, string? Target) ResolveAction(PresentationPart presentation, SlidePart slidePart, A.HyperlinkOnClick? hlink)
    {
        var verb = hlink?.Action?.Value;

        if (verb is { } v && v.StartsWith(ShowJumpActionPrefix, StringComparison.Ordinal))
        {
            var jump = v[ShowJumpActionPrefix.Length..];
            var action = jump switch
            {
                "firstslide" => "first",
                "lastslide" => "last",
                "nextslide" => "next",
                "previousslide" => "prev",
                _ => jump,
            };
            return (action, null);
        }

        if (verb is SlideJumpAction && hlink?.Id?.Value is { Length: > 0 } slideRelId)
        {
            if (slidePart.Parts.FirstOrDefault(p => p.RelationshipId == slideRelId).OpenXmlPart is SlidePart targetPart)
            {
                var slides = PptxDoc.Slides(presentation);
                for (var i = 0; i < slides.Count; i++)
                {
                    if (slides[i].Part.Uri == targetPart.Uri)
                    {
                        return ("slide", Units.Inv($"slide {i + 1}"));
                    }
                }
            }

            return ("slide", "slide ?");
        }

        if (hlink?.Id?.Value is { Length: > 0 } relId)
        {
            var rel = slidePart.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
            return ("url", rel?.Uri.OriginalString);
        }

        return ("none", null);
    }
}
