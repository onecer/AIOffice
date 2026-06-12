using System.Globalization;
using System.Text.RegularExpressions;
using AIOffice.Core;

namespace AIOffice.Pptx;

/// <summary>
/// A parsed pptx address. Accepted forms (1-based indices):
/// <code>
/// /slide[2]
/// /slide[2]/shape[3]          (ordinal)
/// /slide[2]/shape[@id=7]      (stable id — the canonical form aioffice emits)
/// /slide[2]/shape[3]/p[1]
/// </code>
/// The id form survives shape insertions/removals; query/get always return it.
/// </summary>
internal sealed partial record PptxAddress
{
    public const string GrammarHint =
        "pptx paths look like /slide[2], /slide[2]/shape[3], /slide[2]/shape[@id=7] or " +
        "/slide[2]/shape[3]/p[1]; indices are 1-based, @id is the stable shape id from query.";

    [GeneratedRegex(@"^slide\[([0-9]+)\]$")]
    private static partial Regex SlideSegment();

    [GeneratedRegex(@"^shape\[([0-9]+)\]$")]
    private static partial Regex ShapeOrdinalSegment();

    [GeneratedRegex(@"^shape\[@id=([0-9]+)\]$")]
    private static partial Regex ShapeIdSegment();

    [GeneratedRegex(@"^p\[([0-9]+)\]$")]
    private static partial Regex ParagraphSegment();

    [GeneratedRegex(@"^run\[([0-9]+)\]$")]
    private static partial Regex RunSegment();

    private static readonly string[] ReservedRoots = ["master", "layout", "notes", "notesmaster", "handout", "handoutmaster"];

    public required string Raw { get; init; }

    public required int SlideIndex { get; init; }

    public int? ShapeOrdinal { get; init; }

    public uint? ShapeId { get; init; }

    public int? ParagraphIndex { get; init; }

    public int? RunIndex { get; init; }

    public bool HasShape => ShapeOrdinal.HasValue || ShapeId.HasValue;

    /// <summary>Parses an address or throws a typed <c>invalid_path</c>/<c>unsupported_feature</c>.</summary>
    public static PptxAddress Parse(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw) || raw[0] != '/' || raw.Length < 2 || raw[^1] == '/')
        {
            throw Invalid(raw, "Paths start with '/' and must not end with one.");
        }

        var segments = raw[1..].Split('/');
        var rootName = segments[0].Split('[')[0].ToLowerInvariant();
        if (ReservedRoots.Contains(rootName, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Master/layout/notes addressing is reserved for a later milestone: {raw}",
                "Address slide content directly instead, e.g. /slide[2]/shape[3].");
        }

        var slideMatch = SlideSegment().Match(segments[0]);
        if (!slideMatch.Success)
        {
            throw Invalid(raw, $"The first segment must be slide[i]; got '{segments[0]}'.");
        }

        var address = new PptxAddress
        {
            Raw = raw,
            SlideIndex = ParseIndex(slideMatch.Groups[1].Value, raw),
        };

        if (segments.Length == 1)
        {
            return address;
        }

        var ordinal = ShapeOrdinalSegment().Match(segments[1]);
        var byId = ShapeIdSegment().Match(segments[1]);
        if (ordinal.Success)
        {
            address = address with { ShapeOrdinal = ParseIndex(ordinal.Groups[1].Value, raw) };
        }
        else if (byId.Success)
        {
            address = address with { ShapeId = uint.Parse(byId.Groups[1].Value, CultureInfo.InvariantCulture) };
        }
        else
        {
            throw Invalid(raw, $"The second segment must be shape[j] or shape[@id=N]; got '{segments[1]}'.");
        }

        if (segments.Length == 2)
        {
            return address;
        }

        var paragraph = ParagraphSegment().Match(segments[2]);
        if (!paragraph.Success)
        {
            throw Invalid(raw, $"The third segment must be p[k]; got '{segments[2]}'.");
        }

        address = address with { ParagraphIndex = ParseIndex(paragraph.Groups[1].Value, raw) };

        if (segments.Length == 3)
        {
            return address;
        }

        var run = RunSegment().Match(segments[3]);
        if (!run.Success || segments.Length > 4)
        {
            throw Invalid(raw, "Nothing can follow p[k] except run[m].");
        }

        return address with { RunIndex = ParseIndex(run.Groups[1].Value, raw) };
    }

    public string CanonicalSlidePath => Units.Inv($"/slide[{SlideIndex}]");

    private static int ParseIndex(string digits, string raw)
    {
        var value = int.Parse(digits, CultureInfo.InvariantCulture);
        if (value < 1)
        {
            throw Invalid(raw, "Indices are 1-based.");
        }

        return value;
    }

    private static AiofficeException Invalid(string raw, string detail) =>
        new(ErrorCodes.InvalidPath, $"Not a valid pptx path: '{raw}'. {detail}", GrammarHint);
}
