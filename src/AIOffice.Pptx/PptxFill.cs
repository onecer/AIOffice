using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// 1.8 visual fills: the <c>gradient</c> and <c>image</c> props that sit beside
/// the plain <c>fill</c>/<c>background</c> color. A gradient builds an
/// <c>a:gradFill</c> (linear <c>a:lin@ang</c> or radial <c>a:path@path=circle</c>)
/// from a stops list; an image builds an <c>a:blipFill</c> whose <c>a:blip@embed</c>
/// points at a sandbox-resolved picture part on the shape's host part. Both
/// replace any prior solid/gradient/blip fill on the target so the result is one
/// unambiguous fill, mirroring <see cref="PptxEditor"/>'s SetFill behaviour.
/// </summary>
internal static class PptxFill
{
    /// <summary>The 1.8 fill-style props (beyond the legacy solid <c>fill</c>/<c>background</c>).</summary>
    public static readonly IReadOnlyList<string> StyleKeys = ["gradient", "image"];

    /// <summary>True when these props carry a gradient or image fill (so the host threads a part/workspace in).</summary>
    public static bool Handles(JsonObject props) =>
        props.ContainsKey("gradient") || props.ContainsKey("image");

    /// <summary>The fill children a fresh fill replaces (so a target carries exactly one fill).</summary>
    public static bool IsFillElement(OpenXmlElement element) =>
        element is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill;

    // ---------------------------------------------------------------- gradient

    /// <summary>
    /// Builds an <c>a:gradFill</c> from <c>{type:"linear"|"radial", angle?, stops:[{color, at}]}</c>.
    /// Linear writes an <c>a:lin@ang</c> (degrees → 1/60000); radial writes
    /// <c>a:path@path="circle"</c>. Stops are <c>a:gs@pos</c> (0–100 → 0–100000) with an RGB hex color.
    /// </summary>
    public static A.GradientFill BuildGradientFill(JsonNode? node)
    {
        if (node is not JsonObject spec)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "The 'gradient' prop must be an object.",
                "Example: {\"gradient\":{\"type\":\"linear\",\"angle\":90,\"stops\":[" +
                "{\"color\":\"0EA5E9\",\"at\":0},{\"color\":\"6366F1\",\"at\":100}]}}.");
        }

        foreach (var (key, _) in spec)
        {
            if (key is not ("type" or "angle" or "stops"))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown gradient prop '{key}'.",
                    "A gradient takes type (\"linear\"|\"radial\"), angle (degrees, linear only) and stops " +
                    "([{color, at}]).",
                    candidates: ["type", "angle", "stops"]);
            }
        }

        var type = spec.TryGetPropertyValue("type", out var typeNode) && typeNode is not null
            ? J.ScalarText(typeNode).Trim().ToLowerInvariant()
            : "linear";
        if (type is not ("linear" or "radial"))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Gradient type '{type}' is not supported.",
                "Use \"linear\" (default) or \"radial\".",
                candidates: ["linear", "radial"]);
        }

        var stops = BuildStops(spec);

        // The direction child follows the stop list (a:lin / a:path).
        OpenXmlElement direction = type == "radial"
            ? new A.PathGradientFill(new A.FillToRectangle { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 })
            {
                Path = A.PathShadeValues.Circle,
            }
            : new A.LinearGradientFill { Angle = LinearAngle(spec), Scaled = true };

        return new A.GradientFill(stops, direction);
    }

    /// <summary>The stop list: at least two <c>a:gs</c> stops, each a position (0–100) and an RGB color.</summary>
    private static A.GradientStopList BuildStops(JsonObject spec)
    {
        if (!spec.TryGetPropertyValue("stops", out var stopsNode) || stopsNode is not JsonArray stopsArray || stopsArray.Count < 2)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "A gradient needs a 'stops' array of at least two {color, at} entries.",
                "Example: \"stops\":[{\"color\":\"0EA5E9\",\"at\":0},{\"color\":\"6366F1\",\"at\":100}].");
        }

        var list = new A.GradientStopList();
        foreach (var entry in stopsArray)
        {
            if (entry is not JsonObject stop)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "Each gradient stop must be a {color, at} object.",
                    "Example: {\"color\":\"0EA5E9\",\"at\":0}.");
            }

            foreach (var (key, _) in stop)
            {
                if (key is not ("color" or "at"))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Unknown gradient stop prop '{key}'.",
                        "A stop takes color (RRGGBB hex) and at (0–100, the position percent).",
                        candidates: ["color", "at"]);
                }
            }

            if (!stop.TryGetPropertyValue("color", out var colorNode) || colorNode is null)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "A gradient stop needs a 'color'.",
                    "Example: {\"color\":\"0EA5E9\",\"at\":0}.");
            }

            var hex = Units.ParseColorHex("color", colorNode);
            var position = StopPosition(stop);
            list.Append(new A.GradientStop(new A.RgbColorModelHex { Val = hex }) { Position = position });
        }

        return list;
    }

    /// <summary>One stop's position: a 0–100 percent mapped to OOXML's 0–100000 thousandths.</summary>
    private static int StopPosition(JsonObject stop)
    {
        if (!stop.TryGetPropertyValue("at", out var atNode) || atNode is not JsonValue value || !Units.TryNumber(value, out var percent))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "A gradient stop needs an 'at' position (0–100).",
                "Example: {\"color\":\"0EA5E9\",\"at\":0} and {\"color\":\"6366F1\",\"at\":100}.");
        }

        if (percent is < 0 or > 100)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"A gradient stop 'at' must be between 0 and 100; got {percent.ToString(CultureInfo.InvariantCulture)}."),
                "Positions are percentages along the gradient (0 = start, 100 = end).");
        }

        return (int)Math.Round(percent * 1000);
    }

    /// <summary>The linear gradient angle in 1/60000-degree units (default 90° = top-to-bottom).</summary>
    private static int LinearAngle(JsonObject spec)
    {
        var degrees = 90.0;
        if (spec.TryGetPropertyValue("angle", out var angleNode) && angleNode is JsonValue value)
        {
            if (!Units.TryNumber(value, out degrees))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"The gradient 'angle' is not a number: {angleNode.ToJsonString()}.",
                    "Pass degrees, e.g. 0 (left-to-right), 90 (top-to-bottom), 45 (diagonal).");
            }
        }

        // Normalize into [0,360) then scale to OOXML's 60000ths of a degree.
        var normalized = ((degrees % 360) + 360) % 360;
        return (int)Math.Round(normalized * 60000);
    }

    // ------------------------------------------------------------------- image

    /// <summary>
    /// Builds an <c>a:blipFill</c> whose <c>a:blip@embed</c> points at a sandbox-resolved
    /// picture embedded on <paramref name="host"/>. Accepts <c>{image:"banner.jpg"}</c> or
    /// <c>{image:{src, mode:"stretch"|"tile", tint?}}</c>; the image fills the shape area.
    /// </summary>
    public static A.BlipFill BuildImageFill(JsonNode? node, OpenXmlPartContainer host, Workspace workspace)
    {
        string srcText;
        var mode = "stretch";
        string? tint = null;

        switch (node)
        {
            case JsonValue scalar when scalar.TryGetValue<string>(out var raw):
                srcText = raw;
                break;
            case JsonObject spec:
                foreach (var (key, _) in spec)
                {
                    if (key is not ("src" or "mode" or "tint"))
                    {
                        throw new AiofficeException(
                            ErrorCodes.InvalidArgs,
                            $"Unknown image-fill prop '{key}'.",
                            "An image fill takes src (required), mode (\"stretch\"|\"tile\") and tint (RRGGBB hex).",
                            candidates: ["src", "mode", "tint"]);
                    }
                }

                if (!spec.TryGetPropertyValue("src", out var srcNode) || srcNode is null || J.ScalarText(srcNode).Trim().Length == 0)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        "An image fill needs a 'src'.",
                        "Example: {\"image\":{\"src\":\"banner.jpg\",\"mode\":\"tile\"}}.");
                }

                srcText = J.ScalarText(srcNode);
                if (spec.TryGetPropertyValue("mode", out var modeNode) && modeNode is not null)
                {
                    mode = J.ScalarText(modeNode).Trim().ToLowerInvariant();
                    if (mode is not ("stretch" or "tile"))
                    {
                        throw new AiofficeException(
                            ErrorCodes.UnsupportedFeature,
                            $"Image-fill mode '{mode}' is not supported.",
                            "Use \"stretch\" (default, fits the shape) or \"tile\" (repeats).",
                            candidates: ["stretch", "tile"]);
                    }
                }

                if (spec.TryGetPropertyValue("tint", out var tintNode) && tintNode is not null)
                {
                    tint = Units.ParseColorHex("tint", tintNode);
                }

                break;
            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "The 'image' prop must be a path string or a {src, mode, tint} object.",
                    "Example: {\"image\":\"banner.jpg\"} or {\"image\":{\"src\":\"banner.jpg\",\"mode\":\"tile\"}}.");
        }

        // Sandbox first: a path outside the workspace is sandbox_denied, never read.
        var resolved = workspace.Resolve(srcText.Trim(), mustExist: true);
        var relId = EmbedImage(host, resolved);

        var blip = new A.Blip { Embed = relId };
        if (tint is { } hex)
        {
            // A duotone colour overlay over the picture (a:blip/a:duotone): the
            // shape's area shows the image washed toward the tint colour.
            blip.Append(new A.Duotone(
                new A.RgbColorModelHex(new A.LuminanceModulation { Val = 50000 }, new A.LuminanceOffset { Val = 50000 }) { Val = "FFFFFF" },
                new A.RgbColorModelHex { Val = hex }));
        }

        OpenXmlElement fillMode = mode == "tile"
            ? new A.Tile
            {
                HorizontalRatio = 100000,
                VerticalRatio = 100000,
                Flip = A.TileFlipValues.None,
                Alignment = A.RectangleAlignmentValues.TopLeft,
            }
            : new A.Stretch(new A.FillRectangle());

        return new A.BlipFill(blip, fillMode);
    }

    /// <summary>
    /// Embeds an already-sandbox-resolved PNG/JPEG on the shape's host part (a slide,
    /// layout or master part), reusing the picture-add embed path, and returns its rel id.
    /// </summary>
    private static string EmbedImage(OpenXmlPartContainer host, string resolvedSrc)
    {
        // AddImagePart is a typed extension (ISupportedRelationship<ImagePart>), so
        // the concrete part type matters; the three pptx hosts each support it.
        var bytes = File.ReadAllBytes(resolvedSrc);
        var partType = PptxImages.SniffPartType(bytes, resolvedSrc);
        ImagePart imagePart = host switch
        {
            SlidePart slide => slide.AddImagePart(partType),
            SlideLayoutPart layout => layout.AddImagePart(partType),
            SlideMasterPart master => master.AddImagePart(partType),
            _ => throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "An image fill cannot embed a picture on this part.",
                "Apply image fills on a slide, master or layout shape/background."),
        };

        using (var stream = new MemoryStream(bytes, writable: false))
        {
            imagePart.FeedData(stream);
        }

        return host.GetIdOfPart(imagePart);
    }

    // ------------------------------------------------------------------ shared

    /// <summary>
    /// Replaces any prior fill on a shape's <c>spPr</c> with <paramref name="fill"/>
    /// (a gradFill or blipFill), inserting it after the geometry so the part stays
    /// schema-ordered. Mirrors <see cref="PptxEditor"/>'s solid SetFill.
    /// </summary>
    public static void ApplyToShape(P.Shape shape, OpenXmlElement fill)
    {
        var properties = shape.ShapeProperties;
        if (properties is null)
        {
            properties = new P.ShapeProperties();
            shape.InsertAfter(properties, shape.NonVisualShapeProperties);
        }

        foreach (var existing in properties.ChildElements.Where(IsFillElement).ToList())
        {
            existing.Remove();
        }

        OpenXmlElement? anchor = (OpenXmlElement?)properties.GetFirstChild<A.PresetGeometry>()
            ?? (OpenXmlElement?)properties.GetFirstChild<A.CustomGeometry>()
            ?? properties.Transform2D;
        if (anchor is not null)
        {
            properties.InsertAfter(fill, anchor);
        }
        else
        {
            properties.InsertAt(fill, 0);
        }
    }

    /// <summary>
    /// Writes a proper <c>p:bg/p:bgPr</c> carrying <paramref name="fill"/> on any
    /// <c>p:cSld</c> (slide, master or layout), replacing any previous background.
    /// </summary>
    public static void ApplyToBackground(P.CommonSlideData slideData, OpenXmlElement fill)
    {
        slideData.Background = new P.Background(
            new P.BackgroundProperties(fill, new A.EffectList()));
    }
}
