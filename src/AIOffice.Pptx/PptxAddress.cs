using System.Globalization;
using System.Text.RegularExpressions;
using AIOffice.Core;

namespace AIOffice.Pptx;

/// <summary>What the first path segment addresses.</summary>
internal enum PptxRootKind
{
    Slide,
    Master,
}

/// <summary>
/// A parsed pptx address. Accepted forms (1-based indices):
/// <code>
/// /slide[2]
/// /slide[2]/shape[3]          (ordinal)
/// /slide[2]/shape[@id=7]      (stable id — the canonical form aioffice emits)
/// /slide[2]/shape[3]/p[1]
/// /master[1]                  (read-only in this milestone)
/// /master[1]/layout[2]
/// /master[1]/layout[2]/shape[1]
/// </code>
/// The id form survives shape insertions/removals; query/get always return it.
/// </summary>
internal sealed partial record PptxAddress
{
    public const string GrammarHint =
        "pptx paths look like /slide[2], /slide[2]/shape[3], /slide[2]/shape[@id=7], " +
        "/slide[2]/shape[3]/p[1], /master[1], /master[1]/layout[2] or /master[1]/shape[1]; " +
        "indices are 1-based, @id is the stable shape id from query.";

    [GeneratedRegex(@"^slide\[([0-9]+)\]$")]
    private static partial Regex SlideSegment();

    [GeneratedRegex(@"^master\[([0-9]+)\]$")]
    private static partial Regex MasterSegment();

    [GeneratedRegex(@"^layout\[([0-9]+)\]$")]
    private static partial Regex LayoutSegment();

    [GeneratedRegex(@"^shape\[([0-9]+)\]$")]
    private static partial Regex ShapeOrdinalSegment();

    [GeneratedRegex(@"^shape\[@id=([0-9]+)\]$")]
    private static partial Regex ShapeIdSegment();

    [GeneratedRegex(@"^p\[([0-9]+)\]$")]
    private static partial Regex ParagraphSegment();

    [GeneratedRegex(@"^run\[([0-9]+)\]$")]
    private static partial Regex RunSegment();

    private static readonly string[] ReservedRoots = ["notes", "notesmaster", "handout", "handoutmaster"];

    public required string Raw { get; init; }

    public PptxRootKind Root { get; init; } = PptxRootKind.Slide;

    /// <summary>1-based slide index; 0 when <see cref="Root"/> is Master.</summary>
    public int SlideIndex { get; init; }

    /// <summary>1-based master index; 0 when <see cref="Root"/> is Slide.</summary>
    public int MasterIndex { get; init; }

    /// <summary>1-based layout index beneath the master; null for the master itself.</summary>
    public int? LayoutIndex { get; init; }

    public int? ShapeOrdinal { get; init; }

    public uint? ShapeId { get; init; }

    public int? ParagraphIndex { get; init; }

    public int? RunIndex { get; init; }

    public bool HasShape => ShapeOrdinal.HasValue || ShapeId.HasValue;

    public bool IsMaster => Root == PptxRootKind.Master;

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
                $"Notes/handout addressing is reserved for a later milestone: {raw}",
                "Address slide content directly instead, e.g. /slide[2]/shape[3].");
        }

        var masterMatch = MasterSegment().Match(segments[0]);
        if (masterMatch.Success)
        {
            return ParseMasterTail(raw, segments, ParseIndex(masterMatch.Groups[1].Value, raw));
        }

        var slideMatch = SlideSegment().Match(segments[0]);
        if (!slideMatch.Success)
        {
            throw Invalid(raw, $"The first segment must be slide[i] or master[i]; got '{segments[0]}'.");
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

        var shaped = WithShapeSegment(address, segments[1], raw,
            $"The second segment must be shape[j] or shape[@id=N]; got '{segments[1]}'.");

        if (segments.Length == 2)
        {
            return shaped;
        }

        var paragraph = ParagraphSegment().Match(segments[2]);
        if (!paragraph.Success)
        {
            throw Invalid(raw, $"The third segment must be p[k]; got '{segments[2]}'.");
        }

        shaped = shaped with { ParagraphIndex = ParseIndex(paragraph.Groups[1].Value, raw) };

        if (segments.Length == 3)
        {
            return shaped;
        }

        var run = RunSegment().Match(segments[3]);
        if (!run.Success || segments.Length > 4)
        {
            throw Invalid(raw, "Nothing can follow p[k] except run[m].");
        }

        return shaped with { RunIndex = ParseIndex(run.Groups[1].Value, raw) };
    }

    private static PptxAddress ParseMasterTail(string raw, string[] segments, int masterIndex)
    {
        var address = new PptxAddress
        {
            Raw = raw,
            Root = PptxRootKind.Master,
            MasterIndex = masterIndex,
        };

        var next = 1;
        if (segments.Length > 1 && LayoutSegment().Match(segments[1]) is { Success: true } layoutMatch)
        {
            address = address with { LayoutIndex = ParseIndex(layoutMatch.Groups[1].Value, raw) };
            next = 2;
        }

        if (segments.Length == next)
        {
            return address;
        }

        var expectation = next == 1
            ? $"After master[i] comes layout[j] or shape[k]; got '{segments[next]}'."
            : $"After layout[j] comes shape[k] or shape[@id=N]; got '{segments[next]}'.";
        address = WithShapeSegment(address, segments[next], raw, expectation);

        if (segments.Length > next + 1)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Paragraph/run addressing under masters and layouts is not supported yet (planned M2): {raw}",
                "Get the shape itself (e.g. /master[1]/shape[1]) — its full text is included — or address slide content via /slide[i]/shape[j]/p[k].");
        }

        return address;
    }

    private static PptxAddress WithShapeSegment(PptxAddress address, string segment, string raw, string expectation)
    {
        var ordinal = ShapeOrdinalSegment().Match(segment);
        if (ordinal.Success)
        {
            return address with { ShapeOrdinal = ParseIndex(ordinal.Groups[1].Value, raw) };
        }

        var byId = ShapeIdSegment().Match(segment);
        if (byId.Success)
        {
            return address with { ShapeId = uint.Parse(byId.Groups[1].Value, CultureInfo.InvariantCulture) };
        }

        throw Invalid(raw, expectation);
    }

    public string CanonicalSlidePath => Units.Inv($"/slide[{SlideIndex}]");

    /// <summary>The container the address points into: /slide[i], /master[m] or /master[m]/layout[l].</summary>
    public string CanonicalContainerPath => Root == PptxRootKind.Master
        ? LayoutIndex is { } layout
            ? Units.Inv($"/master[{MasterIndex}]/layout[{layout}]")
            : Units.Inv($"/master[{MasterIndex}]")
        : CanonicalSlidePath;

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
