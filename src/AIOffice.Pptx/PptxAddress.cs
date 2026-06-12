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
/// /slide[2]/notes             (the slide's speaker notes)
/// /slide[2]/shape[3]          (ordinal)
/// /slide[2]/shape[@id=7]      (stable id — the canonical form aioffice emits)
/// /slide[2]/shape[3]/p[1]
/// /slide[2]/chart[1]          (1-based chart index on the slide)
/// /slide[2]/table[1]          (1-based table index on the slide)
/// /slide[2]/table[1]/tr[2]    (1-based row in the table)
/// /slide[2]/table[1]/tr[2]/tc[3] (1-based cell in the row)
/// /slide[2]/smartart[1]       (1-based SmartArt index on the slide; read-only)
/// /slide[2]/animation[1]      (1-based animation index on the slide)
/// /slide[2]/comment[@id=3]    (stable comment id on the slide)
/// /master[1]                  (read-only in this milestone)
/// /master[1]/layout[2]
/// /master[1]/layout[2]/shape[1]
/// </code>
/// The id form survives shape insertions/removals; query/get always return it.
/// </summary>
internal sealed partial record PptxAddress
{
    public const string GrammarHint =
        "pptx paths look like /slide[2], /slide[2]/notes, /slide[2]/shape[3], /slide[2]/shape[@id=7], " +
        "/slide[2]/shape[3]/p[1], /slide[2]/chart[1], /slide[2]/table[1], /slide[2]/table[1]/tr[2]/tc[3], " +
        "/slide[2]/smartart[1], /slide[2]/animation[1], /slide[2]/comment[@id=3], " +
        "/master[1], /master[1]/layout[2] or /master[1]/shape[1]; " +
        "indices are 1-based, @id is the stable id from query/get.";

    [GeneratedRegex(@"^slide\[([0-9]+)\]$")]
    private static partial Regex SlideSegment();

    [GeneratedRegex(@"^master\[([0-9]+)\]$")]
    private static partial Regex MasterSegment();

    [GeneratedRegex(@"^layout\[([0-9]+)\]$")]
    private static partial Regex LayoutSegment();

    [GeneratedRegex(@"^shape\[([0-9]+)\]$")]
    private static partial Regex ShapeOrdinalSegment();

    [GeneratedRegex(@"^chart\[([0-9]+)\]$")]
    private static partial Regex ChartSegment();

    [GeneratedRegex(@"^table\[([0-9]+)\]$")]
    private static partial Regex TableSegment();

    [GeneratedRegex(@"^tr\[([0-9]+)\]$")]
    private static partial Regex TableRowSegment();

    [GeneratedRegex(@"^tc\[([0-9]+)\]$")]
    private static partial Regex TableCellSegment();

    [GeneratedRegex(@"^smartart\[([0-9]+)\]$")]
    private static partial Regex SmartArtSegment();

    [GeneratedRegex(@"^animation\[([0-9]+)\]$")]
    private static partial Regex AnimationSegment();

    [GeneratedRegex(@"^comment\[@id=([0-9]+)\]$")]
    private static partial Regex CommentSegment();

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

    /// <summary>1-based chart index on the slide (/slide[i]/chart[k]); null otherwise.</summary>
    public int? ChartIndex { get; init; }

    /// <summary>1-based table index on the slide (/slide[i]/table[k]); null otherwise.</summary>
    public int? TableIndex { get; init; }

    /// <summary>1-based row index inside the table (/slide[i]/table[k]/tr[r]); null otherwise.</summary>
    public int? TableRowIndex { get; init; }

    /// <summary>1-based cell index inside the row (/slide[i]/table[k]/tr[r]/tc[c]); null otherwise.</summary>
    public int? TableCellIndex { get; init; }

    /// <summary>1-based SmartArt index on the slide (/slide[i]/smartart[k]); null otherwise.</summary>
    public int? SmartArtIndex { get; init; }

    /// <summary>1-based animation index on the slide (/slide[i]/animation[k]); null otherwise.</summary>
    public int? AnimationIndex { get; init; }

    /// <summary>Stable comment id on the slide (/slide[i]/comment[@id=N]); null otherwise.</summary>
    public uint? CommentId { get; init; }

    public int? ParagraphIndex { get; init; }

    public int? RunIndex { get; init; }

    /// <summary>True when the path addresses a slide's speaker notes (/slide[i]/notes).</summary>
    public bool IsNotes { get; init; }

    public bool HasShape => ShapeOrdinal.HasValue || ShapeId.HasValue;

    /// <summary>True when the path addresses a chart by index (/slide[i]/chart[k]).</summary>
    public bool IsChart => ChartIndex.HasValue;

    /// <summary>True when the path addresses a table (or a row/cell inside one).</summary>
    public bool IsTable => TableIndex.HasValue;

    /// <summary>True when the path addresses a SmartArt diagram by index (/slide[i]/smartart[k]).</summary>
    public bool IsSmartArt => SmartArtIndex.HasValue;

    /// <summary>True when the path addresses an animation by index (/slide[i]/animation[k]).</summary>
    public bool IsAnimation => AnimationIndex.HasValue;

    /// <summary>True when the path addresses a comment by id (/slide[i]/comment[@id=N]).</summary>
    public bool IsComment => CommentId.HasValue;

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
                "Speaker notes live under their slide — use /slide[2]/notes; address slide content via /slide[2]/shape[3].");
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

        if (string.Equals(segments[1], "notes", StringComparison.OrdinalIgnoreCase))
        {
            if (segments.Length > 2)
            {
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"Paragraph/run addressing inside notes is not supported yet: {raw}",
                    "Target the whole notes body: get/set /slide[i]/notes, or use op 'add' to append one paragraph.");
            }

            return address with { IsNotes = true };
        }

        if (ChartSegment().Match(segments[1]) is { Success: true } chartMatch)
        {
            if (segments.Length > 2)
            {
                throw Invalid(raw, "Nothing can follow chart[k].");
            }

            return address with { ChartIndex = ParseIndex(chartMatch.Groups[1].Value, raw) };
        }

        if (TableSegment().Match(segments[1]) is { Success: true } tableMatch)
        {
            return ParseTableTail(raw, segments, address with
            {
                TableIndex = ParseIndex(tableMatch.Groups[1].Value, raw),
            });
        }

        if (SmartArtSegment().Match(segments[1]) is { Success: true } smartArtMatch)
        {
            if (segments.Length > 2)
            {
                throw Invalid(raw, "Nothing can follow smartart[k].");
            }

            return address with { SmartArtIndex = ParseIndex(smartArtMatch.Groups[1].Value, raw) };
        }

        if (AnimationSegment().Match(segments[1]) is { Success: true } animationMatch)
        {
            if (segments.Length > 2)
            {
                throw Invalid(raw, "Nothing can follow animation[k].");
            }

            return address with { AnimationIndex = ParseIndex(animationMatch.Groups[1].Value, raw) };
        }

        if (CommentSegment().Match(segments[1]) is { Success: true } commentMatch)
        {
            if (segments.Length > 2)
            {
                throw Invalid(raw, "Nothing can follow comment[@id=N].");
            }

            return address with { CommentId = uint.Parse(commentMatch.Groups[1].Value, CultureInfo.InvariantCulture) };
        }

        var shaped = WithShapeSegment(address, segments[1], raw,
            $"The second segment must be notes, chart[k], table[k], smartart[k], animation[k], comment[@id=N], " +
            $"shape[j] or shape[@id=N]; got '{segments[1]}'.");

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

    /// <summary>Parses the optional /tr[r]/tc[c] tail after table[k].</summary>
    private static PptxAddress ParseTableTail(string raw, string[] segments, PptxAddress address)
    {
        if (segments.Length == 2)
        {
            return address;
        }

        var rowMatch = TableRowSegment().Match(segments[2]);
        if (!rowMatch.Success)
        {
            throw Invalid(raw, $"After table[k] comes tr[r]; got '{segments[2]}'.");
        }

        address = address with { TableRowIndex = ParseIndex(rowMatch.Groups[1].Value, raw) };
        if (segments.Length == 3)
        {
            return address;
        }

        var cellMatch = TableCellSegment().Match(segments[3]);
        if (!cellMatch.Success || segments.Length > 4)
        {
            throw Invalid(raw, "Nothing can follow tr[r] except tc[c].");
        }

        return address with { TableCellIndex = ParseIndex(cellMatch.Groups[1].Value, raw) };
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

    /// <summary>The canonical speaker-notes path of the addressed slide.</summary>
    public string CanonicalNotesPath => Units.Inv($"/slide[{SlideIndex}]/notes");

    /// <summary>The canonical chart-index path (/slide[i]/chart[k]) of the addressed chart.</summary>
    public string CanonicalChartPath => Units.Inv($"/slide[{SlideIndex}]/chart[{ChartIndex}]");

    /// <summary>The canonical table path (/slide[i]/table[k]) of the addressed table.</summary>
    public string CanonicalTablePath => Units.Inv($"/slide[{SlideIndex}]/table[{TableIndex}]");

    /// <summary>The canonical SmartArt path (/slide[i]/smartart[k]) of the addressed diagram.</summary>
    public string CanonicalSmartArtPath => Units.Inv($"/slide[{SlideIndex}]/smartart[{SmartArtIndex}]");

    /// <summary>The canonical animation-index path (/slide[i]/animation[k]) of the addressed animation.</summary>
    public string CanonicalAnimationPath => Units.Inv($"/slide[{SlideIndex}]/animation[{AnimationIndex}]");

    /// <summary>The canonical comment path (/slide[i]/comment[@id=N]) of the addressed comment.</summary>
    public string CanonicalCommentPath => Units.Inv($"/slide[{SlideIndex}]/comment[@id={CommentId}]");

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
