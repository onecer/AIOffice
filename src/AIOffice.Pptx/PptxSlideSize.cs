using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// M6 custom slide size — set / {slideSize:"16:9|4:3|16:10|A4|letter"} or an
/// explicit {width,height} (cm/in/emu). Updates p:sldSz only; existing shapes
/// keep their EMU coordinates (nothing is rescaled or repositioned). get / reports
/// the named preset (when it matches one) plus the dimensions in cm.
/// </summary>
internal static class PptxSlideSize
{
    /// <summary>Named presets (width, height) in EMU, plus the sldSz Type token PowerPoint expects.</summary>
    private static readonly IReadOnlyDictionary<string, (long Width, long Height, P.SlideSizeValues? Type)> Presets =
        new Dictionary<string, (long, long, P.SlideSizeValues?)>(StringComparer.OrdinalIgnoreCase)
        {
            ["16:9"] = (12_192_000, 6_858_000, P.SlideSizeValues.Custom),
            ["4:3"] = (9_144_000, 6_858_000, P.SlideSizeValues.Screen4x3),
            ["16:10"] = (10_972_800, 6_858_000, P.SlideSizeValues.Screen16x10),
            ["a4"] = (10_692_000, 7_560_000, P.SlideSizeValues.A4),
            ["a3"] = (15_120_000, 10_692_000, P.SlideSizeValues.A3),
            ["letter"] = (9_144_000, 6_858_000, P.SlideSizeValues.Letter),
            ["ledger"] = (12_192_000, 9_144_000, P.SlideSizeValues.Ledger),
        };

    /// <summary>The named presets, lower-cased, for error candidates.</summary>
    public static IReadOnlyList<string> PresetNames => [.. Presets.Keys];

    /// <summary>set / {slideSize | width+height}: rewrites p:sldSz. Returns the canonical "/" path.</summary>
    public static string Set(PresentationPart presentation, JsonObject props)
    {
        var presentationXml = presentation.Presentation ?? throw Corrupt("the presentation part has no presentation XML");
        var size = presentationXml.SlideSize ??= new P.SlideSize
        {
            Cx = PptxFactory.SlideWidthEmu,
            Cy = PptxFactory.SlideHeightEmu,
        };

        var hasSlideSize = props.ContainsKey("slideSize");
        var hasWidth = props.ContainsKey("width");
        var hasHeight = props.ContainsKey("height");

        foreach (var (key, _) in props)
        {
            if (key is not ("slideSize" or "width" or "height"))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Prop '{key}' does not apply to the presentation root.",
                    "Root props: slideSize (a named preset) or width + height (explicit dimensions); " +
                    "to add a section use {\"op\":\"add\",\"path\":\"/\",\"type\":\"section\"}.",
                    candidates: ["slideSize", "width", "height"]);
            }
        }

        if (hasSlideSize)
        {
            if (hasWidth || hasHeight)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "slideSize and width/height cannot be combined.",
                    "Pass a named preset (e.g. {\"slideSize\":\"16:9\"}) or explicit dimensions " +
                    "({\"width\":\"33.87cm\",\"height\":\"19.05cm\"}), not both.");
            }

            ApplyPreset(size, J.ScalarText(props["slideSize"]!));
            return "/";
        }

        if (!hasWidth || !hasHeight)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "Explicit slide dimensions need both width and height.",
                "Pass {\"width\":\"33.87cm\",\"height\":\"19.05cm\"}, or use a named preset {\"slideSize\":\"16:9\"}.");
        }

        var width = Units.ParseLengthEmu("width", props["width"]);
        var height = Units.ParseLengthEmu("height", props["height"]);
        RequireInRange(width, "width");
        RequireInRange(height, "height");
        size.Cx = (int)width;
        size.Cy = (int)height;
        size.Type = P.SlideSizeValues.Custom;
        return "/";
    }

    private static void ApplyPreset(P.SlideSize size, string token)
    {
        if (!Presets.TryGetValue(token.Trim(), out var preset))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown slide size preset '{token}'.",
                "Use 16:9, 4:3, 16:10, A4, A3, letter or ledger; or pass explicit width and height.",
                candidates: PresetNames);
        }

        size.Cx = (int)preset.Width;
        size.Cy = (int)preset.Height;
        if (preset.Type is { } type)
        {
            size.Type = type;
        }
    }

    /// <summary>The `get` projection for "/": the slide size, dimensions and slide count.</summary>
    public static object Detail(PresentationPart presentation)
    {
        var size = presentation.Presentation?.SlideSize;
        var cx = size?.Cx?.Value ?? PptxFactory.SlideWidthEmu;
        var cy = size?.Cy?.Value ?? PptxFactory.SlideHeightEmu;
        return new
        {
            Path = "/",
            SlideSize = MatchPreset(cx, cy),
            WidthCm = Units.EmuToCm(cx),
            HeightCm = Units.EmuToCm(cy),
            SlideCount = PptxDoc.Slides(presentation).Count,
            SectionCount = PptxSections.List(presentation).Count,
        };
    }

    /// <summary>The named preset matching these dimensions, when one does (else null for a custom size).</summary>
    public static string? MatchPreset(long cx, long cy)
    {
        foreach (var (name, preset) in Presets)
        {
            if (preset.Width == cx && preset.Height == cy)
            {
                return name;
            }
        }

        return null;
    }

    private static void RequireInRange(long emu, string key)
    {
        // PowerPoint's sldSz bounds: 914400 (1in) .. 51206400 (56in).
        const long min = 914_400;
        const long max = 51_206_400;
        if (emu < min || emu > max)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"{key} {Units.EmuToCm(emu)}cm is out of range."),
                "Slide dimensions must be between 2.54cm (1in) and 142.24cm (56in).");
        }
    }

    private static AiofficeException Corrupt(string detail) => new(
        ErrorCodes.FormatCorrupt,
        $"The presentation is malformed: {detail}.",
        "Re-export the file from PowerPoint/Keynote, or restore a snapshot.");
}
