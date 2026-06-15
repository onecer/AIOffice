using System.Globalization;
using System.Text.RegularExpressions;
using AIOffice.Core;

namespace AIOffice.Pptx;

/// <summary>What the first path segment addresses.</summary>
internal enum PptxRootKind
{
    Slide,
    Master,

    /// <summary>The presentation itself ("/"): slide size and section management.</summary>
    Presentation,

    /// <summary>A named slide section ("/section[i]").</summary>
    Section,

    /// <summary>The package document properties ("/properties"): core + custom metadata.</summary>
    Properties,

    /// <summary>The deck's embedded-font list ("/fonts"): embed/list/remove fonts.</summary>
    Fonts,
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
/// /slide[2]/group[@id=7]      (a shape group, by its p:grpSp shape id)
/// /slide[2]/group[@id=7]/shape[@id=9] (a shape inside the group)
/// /slide[2]/animation[1]      (1-based animation index on the slide)
/// /slide[2]/comment[@id=3]    (stable comment id on the slide)
/// /slide[2]/embed[1]          (1-based embedded-object index on the slide)
/// /slide[2]/embed[@id=7]      (stable embed id — the host graphicFrame's shape id)
/// /slide[2]/shape[@id=7]/omath[1] (1-based equation in the shape's text body)
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
        "/slide[2]/smartart[1], /slide[2]/group[@id=7], /slide[2]/group[@id=7]/shape[@id=9], " +
        "/slide[2]/animation[1], /slide[2]/zoom[1], /slide[2]/comment[@id=3], " +
        "/slide[2]/embed[1], /slide[2]/embed[@id=7], /slide[2]/media[1], /slide[2]/media[@id=7], " +
        "/slide[2]/model3d[1], /slide[2]/model3d[@id=7], " +
        "/slide[2]/shape[@id=7]/omath[1], " +
        "/master[1], /master[1]/layout[2], /master[1]/layout[@name=My Layout], /master[1]/shape[1], " +
        "/section[1] (a slide section), /properties (document core + custom metadata), " +
        "/fonts (embedded fonts) or /fonts/font[@name=MyFont] " +
        "or / (the presentation: slide size and sections); " +
        "indices are 1-based, @id is the stable id from query/get.";

    [GeneratedRegex(@"^slide\[([0-9]+)\]$")]
    private static partial Regex SlideSegment();

    [GeneratedRegex(@"^master\[([0-9]+)\]$")]
    private static partial Regex MasterSegment();

    [GeneratedRegex(@"^section\[([0-9]+)\]$")]
    private static partial Regex SectionSegment();

    [GeneratedRegex(@"^layout\[([0-9]+)\]$")]
    private static partial Regex LayoutSegment();

    [GeneratedRegex(@"^layout\[@name=(.+)\]$")]
    private static partial Regex LayoutNameSegment();

    [GeneratedRegex(@"^font\[@name=(.+)\]$")]
    private static partial Regex FontNameSegment();

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

    [GeneratedRegex(@"^group\[@id=([0-9]+)\]$")]
    private static partial Regex GroupSegment();

    [GeneratedRegex(@"^animation\[([0-9]+)\]$")]
    private static partial Regex AnimationSegment();

    [GeneratedRegex(@"^zoom\[([0-9]+)\]$")]
    private static partial Regex ZoomSegment();

    [GeneratedRegex(@"^comment\[@id=([0-9]+)\]$")]
    private static partial Regex CommentSegment();

    [GeneratedRegex(@"^embed\[([0-9]+)\]$")]
    private static partial Regex EmbedOrdinalSegment();

    [GeneratedRegex(@"^embed\[@id=([0-9]+)\]$")]
    private static partial Regex EmbedIdSegment();

    [GeneratedRegex(@"^media\[([0-9]+)\]$")]
    private static partial Regex MediaOrdinalSegment();

    [GeneratedRegex(@"^media\[@id=([0-9]+)\]$")]
    private static partial Regex MediaIdSegment();

    [GeneratedRegex(@"^model3d\[([0-9]+)\]$")]
    private static partial Regex Model3DOrdinalSegment();

    [GeneratedRegex(@"^model3d\[@id=([0-9]+)\]$")]
    private static partial Regex Model3DIdSegment();

    [GeneratedRegex(@"^shape\[@id=([0-9]+)\]$")]
    private static partial Regex ShapeIdSegment();

    [GeneratedRegex(@"^p\[([0-9]+)\]$")]
    private static partial Regex ParagraphSegment();

    [GeneratedRegex(@"^run\[([0-9]+)\]$")]
    private static partial Regex RunSegment();

    [GeneratedRegex(@"^omath\[([0-9]+)\]$")]
    private static partial Regex OMathSegment();

    private static readonly string[] ReservedRoots = ["notes", "notesmaster", "handout", "handoutmaster"];

    public required string Raw { get; init; }

    public PptxRootKind Root { get; init; } = PptxRootKind.Slide;

    /// <summary>1-based slide index; 0 when <see cref="Root"/> is Master.</summary>
    public int SlideIndex { get; init; }

    /// <summary>1-based master index; 0 when <see cref="Root"/> is Slide.</summary>
    public int MasterIndex { get; init; }

    /// <summary>1-based layout index beneath the master; null for the master itself.</summary>
    public int? LayoutIndex { get; init; }

    /// <summary>Layout display name ("/master[m]/layout[@name=...]"); null when the layout is addressed by index.</summary>
    public string? LayoutName { get; init; }

    /// <summary>Embedded-font display name ("/fonts/font[@name=...]"); null for the whole /fonts list.</summary>
    public string? FontName { get; init; }

    /// <summary>1-based slide-section index ("/section[i]"); 0 when <see cref="Root"/> is not Section.</summary>
    public int SectionIndex { get; init; }

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

    /// <summary>Stable group id on the slide (/slide[i]/group[@id=N], the p:grpSp's shape id); null otherwise.</summary>
    public uint? GroupId { get; init; }

    /// <summary>1-based animation index on the slide (/slide[i]/animation[k]); null otherwise.</summary>
    public int? AnimationIndex { get; init; }

    /// <summary>1-based zoom index on the slide (/slide[i]/zoom[k], a zoom navigation object); null otherwise.</summary>
    public int? ZoomIndex { get; init; }

    /// <summary>Stable comment id on the slide (/slide[i]/comment[@id=N]); null otherwise.</summary>
    public uint? CommentId { get; init; }

    /// <summary>1-based embed index on the slide (/slide[i]/embed[k]); null otherwise.</summary>
    public int? EmbedOrdinal { get; init; }

    /// <summary>Stable embed id on the slide (/slide[i]/embed[@id=N], the host graphicFrame's shape id); null otherwise.</summary>
    public uint? EmbedId { get; init; }

    /// <summary>1-based media index on the slide (/slide[i]/media[k]); null otherwise.</summary>
    public int? MediaOrdinal { get; init; }

    /// <summary>Stable media id on the slide (/slide[i]/media[@id=N], the host picture's shape id); null otherwise.</summary>
    public uint? MediaId { get; init; }

    /// <summary>1-based 3D-model index on the slide (/slide[i]/model3d[k]); null otherwise.</summary>
    public int? Model3DOrdinal { get; init; }

    /// <summary>Stable 3D-model id on the slide (/slide[i]/model3d[@id=N], the host picture's shape id); null otherwise.</summary>
    public uint? Model3DId { get; init; }

    public int? ParagraphIndex { get; init; }

    public int? RunIndex { get; init; }

    /// <summary>1-based equation index inside the shape's text body (/slide[i]/shape[@id=N]/omath[k]); null otherwise.</summary>
    public int? OMathIndex { get; init; }

    /// <summary>True when the path addresses a slide's speaker notes (/slide[i]/notes).</summary>
    public bool IsNotes { get; init; }

    public bool HasShape => ShapeOrdinal.HasValue || ShapeId.HasValue;

    /// <summary>True when the path addresses a chart by index (/slide[i]/chart[k]).</summary>
    public bool IsChart => ChartIndex.HasValue;

    /// <summary>True when the path addresses a table (or a row/cell inside one).</summary>
    public bool IsTable => TableIndex.HasValue;

    /// <summary>True when the path addresses a SmartArt diagram by index (/slide[i]/smartart[k]).</summary>
    public bool IsSmartArt => SmartArtIndex.HasValue;

    /// <summary>
    /// True when the path addresses a shape group (/slide[i]/group[@id=N]) or a shape
    /// inside one (/slide[i]/group[@id=N]/shape[...]).
    /// </summary>
    public bool IsGroup => GroupId.HasValue;

    /// <summary>True when the path addresses an animation by index (/slide[i]/animation[k]).</summary>
    public bool IsAnimation => AnimationIndex.HasValue;

    /// <summary>True when the path addresses a zoom navigation object by index (/slide[i]/zoom[k]).</summary>
    public bool IsZoom => ZoomIndex.HasValue;

    /// <summary>True when the path addresses a comment by id (/slide[i]/comment[@id=N]).</summary>
    public bool IsComment => CommentId.HasValue;

    /// <summary>True when the path addresses an embedded object (/slide[i]/embed[k] or /slide[i]/embed[@id=N]).</summary>
    public bool IsEmbed => EmbedOrdinal.HasValue || EmbedId.HasValue;

    /// <summary>True when the path addresses a media object (/slide[i]/media[k] or /slide[i]/media[@id=N]).</summary>
    public bool IsMedia => MediaOrdinal.HasValue || MediaId.HasValue;

    /// <summary>True when the path addresses a 3D model (/slide[i]/model3d[k] or /slide[i]/model3d[@id=N]).</summary>
    public bool IsModel3D => Model3DOrdinal.HasValue || Model3DId.HasValue;

    /// <summary>True when the path addresses an equation inside a shape (/slide[i]/shape[@id=N]/omath[k]).</summary>
    public bool IsOMath => OMathIndex.HasValue;

    public bool IsMaster => Root == PptxRootKind.Master;

    /// <summary>True when the path is the presentation root ("/").</summary>
    public bool IsPresentation => Root == PptxRootKind.Presentation;

    /// <summary>True when the path addresses a slide section ("/section[i]").</summary>
    public bool IsSection => Root == PptxRootKind.Section;

    /// <summary>True when the path is the document-properties root ("/properties").</summary>
    public bool IsProperties => Root == PptxRootKind.Properties;

    /// <summary>True when the path is the embedded-font root ("/fonts" or "/fonts/font[@name=...]").</summary>
    public bool IsFonts => Root == PptxRootKind.Fonts;

    /// <summary>Parses an address or throws a typed <c>invalid_path</c>/<c>unsupported_feature</c>.</summary>
    public static PptxAddress Parse(string raw)
    {
        // "/" is the presentation root (slide size and sections); every other path is at least "/x".
        if (raw == "/")
        {
            return new PptxAddress { Raw = raw, Root = PptxRootKind.Presentation };
        }

        // "/properties" is the package metadata root (core + custom document properties).
        if (string.Equals(raw, "/properties", StringComparison.OrdinalIgnoreCase))
        {
            return new PptxAddress { Raw = raw, Root = PptxRootKind.Properties };
        }

        // "/fonts" is the embedded-font list root; "/fonts/font[@name=...]" addresses one font.
        if (string.Equals(raw, "/fonts", StringComparison.OrdinalIgnoreCase))
        {
            return new PptxAddress { Raw = raw, Root = PptxRootKind.Fonts };
        }

        if (raw.StartsWith("/fonts/", StringComparison.OrdinalIgnoreCase))
        {
            var fontSeg = raw["/fonts/".Length..];
            if (FontNameSegment().Match(fontSeg) is { Success: true } fontMatch)
            {
                return new PptxAddress
                {
                    Raw = raw,
                    Root = PptxRootKind.Fonts,
                    FontName = fontMatch.Groups[1].Value,
                };
            }

            throw Invalid(raw, "After /fonts comes font[@name=...]; got '" + fontSeg + "'.");
        }

        if (string.IsNullOrWhiteSpace(raw) || raw[0] != '/' || raw.Length < 2 || raw[^1] == '/')
        {
            throw Invalid(raw, "Paths start with '/' and must not end with one ('/' alone is the presentation root).");
        }

        var segments = raw[1..].Split('/');
        if (SectionSegment().Match(segments[0]) is { Success: true } sectionMatch)
        {
            if (segments.Length > 1)
            {
                throw Invalid(raw, "Nothing can follow section[i]; assign slides by range when adding/moving sections.");
            }

            return new PptxAddress
            {
                Raw = raw,
                Root = PptxRootKind.Section,
                SectionIndex = ParseIndex(sectionMatch.Groups[1].Value, raw),
            };
        }

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

        if (GroupSegment().Match(segments[1]) is { Success: true } groupMatch)
        {
            var grouped = address with { GroupId = uint.Parse(groupMatch.Groups[1].Value, CultureInfo.InvariantCulture) };
            if (segments.Length == 2)
            {
                return grouped;
            }

            // A child shape inside the group: /slide[i]/group[@id=N]/shape[k] or shape[@id=M].
            grouped = WithShapeSegment(grouped, segments[2], raw,
                $"After group[@id=N] comes shape[k] or shape[@id=M]; got '{segments[2]}'.");
            if (segments.Length > 3)
            {
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"Paragraph/run addressing inside a group's child shape is not supported: {raw}",
                    "Set the child shape's text with a 'text' prop on the shape path " +
                    "(e.g. /slide[1]/group[@id=5]/shape[@id=7]), or ungroup first and address /slide[i]/shape[j]/p[k].");
            }

            return grouped;
        }

        if (AnimationSegment().Match(segments[1]) is { Success: true } animationMatch)
        {
            if (segments.Length > 2)
            {
                throw Invalid(raw, "Nothing can follow animation[k].");
            }

            return address with { AnimationIndex = ParseIndex(animationMatch.Groups[1].Value, raw) };
        }

        if (ZoomSegment().Match(segments[1]) is { Success: true } zoomMatch)
        {
            if (segments.Length > 2)
            {
                throw Invalid(raw, "Nothing can follow zoom[k].");
            }

            return address with { ZoomIndex = ParseIndex(zoomMatch.Groups[1].Value, raw) };
        }

        if (CommentSegment().Match(segments[1]) is { Success: true } commentMatch)
        {
            if (segments.Length > 2)
            {
                throw Invalid(raw, "Nothing can follow comment[@id=N].");
            }

            return address with { CommentId = uint.Parse(commentMatch.Groups[1].Value, CultureInfo.InvariantCulture) };
        }

        if (EmbedOrdinalSegment().Match(segments[1]) is { Success: true } embedOrdinalMatch)
        {
            if (segments.Length > 2)
            {
                throw Invalid(raw, "Nothing can follow embed[k].");
            }

            return address with { EmbedOrdinal = ParseIndex(embedOrdinalMatch.Groups[1].Value, raw) };
        }

        if (EmbedIdSegment().Match(segments[1]) is { Success: true } embedIdMatch)
        {
            if (segments.Length > 2)
            {
                throw Invalid(raw, "Nothing can follow embed[@id=N].");
            }

            return address with { EmbedId = uint.Parse(embedIdMatch.Groups[1].Value, CultureInfo.InvariantCulture) };
        }

        if (MediaOrdinalSegment().Match(segments[1]) is { Success: true } mediaOrdinalMatch)
        {
            if (segments.Length > 2)
            {
                throw Invalid(raw, "Nothing can follow media[k].");
            }

            return address with { MediaOrdinal = ParseIndex(mediaOrdinalMatch.Groups[1].Value, raw) };
        }

        if (MediaIdSegment().Match(segments[1]) is { Success: true } mediaIdMatch)
        {
            if (segments.Length > 2)
            {
                throw Invalid(raw, "Nothing can follow media[@id=N].");
            }

            return address with { MediaId = uint.Parse(mediaIdMatch.Groups[1].Value, CultureInfo.InvariantCulture) };
        }

        if (Model3DOrdinalSegment().Match(segments[1]) is { Success: true } model3DOrdinalMatch)
        {
            if (segments.Length > 2)
            {
                throw Invalid(raw, "Nothing can follow model3d[k].");
            }

            return address with { Model3DOrdinal = ParseIndex(model3DOrdinalMatch.Groups[1].Value, raw) };
        }

        if (Model3DIdSegment().Match(segments[1]) is { Success: true } model3DIdMatch)
        {
            if (segments.Length > 2)
            {
                throw Invalid(raw, "Nothing can follow model3d[@id=N].");
            }

            return address with { Model3DId = uint.Parse(model3DIdMatch.Groups[1].Value, CultureInfo.InvariantCulture) };
        }

        var shaped = WithShapeSegment(address, segments[1], raw,
            $"The second segment must be notes, chart[k], table[k], smartart[k], animation[k], zoom[k], comment[@id=N], " +
            $"embed[k], embed[@id=N], media[k], media[@id=N], model3d[k], model3d[@id=N], group[@id=N], shape[j] or shape[@id=N]; got '{segments[1]}'.");

        if (segments.Length == 2)
        {
            return shaped;
        }

        // An equation inside a shape's text body: /slide[i]/shape[@id=N]/omath[k] (1-based).
        if (OMathSegment().Match(segments[2]) is { Success: true } omathMatch)
        {
            if (segments.Length > 3)
            {
                throw Invalid(raw, "Nothing can follow omath[k].");
            }

            return shaped with { OMathIndex = ParseIndex(omathMatch.Groups[1].Value, raw) };
        }

        var paragraph = ParagraphSegment().Match(segments[2]);
        if (!paragraph.Success)
        {
            throw Invalid(raw, $"The third segment must be p[k] or omath[k]; got '{segments[2]}'.");
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
        else if (segments.Length > 1 && LayoutNameSegment().Match(segments[1]) is { Success: true } layoutNameMatch)
        {
            address = address with { LayoutName = layoutNameMatch.Groups[1].Value };
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
                $"Paragraph/run addressing under masters and layouts is not supported: {raw}",
                "Set the whole shape's text with a 'text' prop on the shape path (e.g. /master[1]/shape[1]), " +
                "or address slide content via /slide[i]/shape[j]/p[k].");
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

    /// <summary>The canonical group path (/slide[i]/group[@id=N]) of the addressed group, by its p:grpSp shape id.</summary>
    public string CanonicalGroupPath => Units.Inv($"/slide[{SlideIndex}]/group[@id={GroupId}]");

    /// <summary>The canonical animation-index path (/slide[i]/animation[k]) of the addressed animation.</summary>
    public string CanonicalAnimationPath => Units.Inv($"/slide[{SlideIndex}]/animation[{AnimationIndex}]");

    /// <summary>The canonical zoom-index path (/slide[i]/zoom[k]) of the addressed zoom object.</summary>
    public string CanonicalZoomPath => Units.Inv($"/slide[{SlideIndex}]/zoom[{ZoomIndex}]");

    /// <summary>The canonical comment path (/slide[i]/comment[@id=N]) of the addressed comment.</summary>
    public string CanonicalCommentPath => Units.Inv($"/slide[{SlideIndex}]/comment[@id={CommentId}]");

    /// <summary>The canonical embed path (/slide[i]/embed[@id=N]) of the addressed embed, by its host graphicFrame id.</summary>
    public string CanonicalEmbedPath(uint id) => Units.Inv($"/slide[{SlideIndex}]/embed[@id={id}]");

    /// <summary>The canonical media path (/slide[i]/media[@id=N]) of the addressed media, by its host picture id.</summary>
    public string CanonicalMediaPath(uint id) => Units.Inv($"/slide[{SlideIndex}]/media[@id={id}]");

    /// <summary>The canonical 3D-model path (/slide[i]/model3d[@id=N]) of the addressed model, by its host picture id.</summary>
    public string CanonicalModel3DPath(uint id) => Units.Inv($"/slide[{SlideIndex}]/model3d[@id={id}]");

    /// <summary>The canonical equation path (/slide[i]/shape[@id=N]/omath[k]) for the host shape id and 1-based index.</summary>
    public string CanonicalOMathPath(uint shapeId, int index) =>
        Units.Inv($"/slide[{SlideIndex}]/shape[@id={shapeId}]/omath[{index}]");

    /// <summary>The canonical slide-section path (/section[i]) of the addressed section.</summary>
    public string CanonicalSectionPath => Units.Inv($"/section[{SectionIndex}]");

    /// <summary>The canonical master path (/master[m]) of the addressed master.</summary>
    public string CanonicalMasterPath => Units.Inv($"/master[{MasterIndex}]");

    /// <summary>The canonical layout path (/master[m]/layout[l]) of the addressed layout.</summary>
    public string CanonicalLayoutPath => Units.Inv($"/master[{MasterIndex}]/layout[{LayoutIndex}]");

    /// <summary>The canonical embedded-font path (/fonts/font[@name=...]) for a font name.</summary>
    public static string CanonicalFontPath(string name) => Units.Inv($"/fonts/font[@name={name}]");

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
