using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace AIOffice.Core;

/// <summary>What a single path segment addresses.</summary>
public enum PathSegmentKind
{
    /// <summary>A named element, optionally 1-based indexed: <c>body</c>, <c>p[3]</c>, <c>slide[2]</c>, <c>row[3]</c>.</summary>
    Element,

    /// <summary>A quoted name (sheet names with spaces/specials): <c>'Q3 Data'</c>.</summary>
    Name,

    /// <summary>A single spreadsheet cell: <c>A1</c>.</summary>
    Cell,

    /// <summary>A rectangular cell range: <c>A1:C10</c>.</summary>
    Range,

    /// <summary>
    /// A contiguous run of indexed/letter-indexed elements of the same name
    /// (xlsx outline groups): <c>row[2]:row[6]</c>, <c>col[B]:col[E]</c>. The
    /// handler owns the span semantics; the grammar only validates the shape.
    /// </summary>
    ElementSpan,
}

/// <summary>A spreadsheet cell reference (column letters + 1-based row).</summary>
public readonly record struct CellRef(string Column, int Row)
{
    public override string ToString() => Column + Row.ToString(CultureInfo.InvariantCulture);

    /// <summary>Column letters converted to a 1-based column number (A=1, Z=26, AA=27).</summary>
    public int ColumnNumber
    {
        get
        {
            var n = 0;
            foreach (var c in Column)
            {
                n = (n * 26) + (c - 'A' + 1);
            }

            return n;
        }
    }
}

/// <summary>
/// One segment of a document path. <see cref="Name"/> holds the element or
/// quoted name; cells/ranges expose <see cref="Start"/> and <see cref="End"/>.
/// </summary>
public sealed record PathSegment
{
    public required PathSegmentKind Kind { get; init; }

    /// <summary>Element name (<c>p</c>, <c>table</c>, <c>slide</c>…) or unquoted sheet name. Null for cells/ranges.</summary>
    public string? Name { get; init; }

    /// <summary>1-based index for indexed elements; null when the element is unindexed.</summary>
    public int? Index { get; init; }

    /// <summary>Uppercase column letters for letter-indexed elements (xlsx <c>col[C]</c>); null otherwise.</summary>
    public string? Letter { get; init; }

    /// <summary>Stable-id selector value for <c>element[@id=…]</c> / <c>element[@name=…]</c> segments (e.g. shape[@id=7], revision[@id=3], bookmark[@name=Results]); null otherwise.</summary>
    public string? Id { get; init; }

    /// <summary>Named-variant selector for the fixed variant vocabulary (docx M5 <c>header[firstPage]</c>, <c>footer[even]</c>); null otherwise.</summary>
    public string? Variant { get; init; }

    /// <summary>The attribute the id selector matched on: "id" (default), "name" (e.g. bookmark[@name=Results]) or "tag" (e.g. sdt[@tag=clientName]); null for non-id segments.</summary>
    public string? IdAttribute { get; init; }

    /// <summary>Cell (Cell kind) or top-left of a range.</summary>
    public CellRef? Start { get; init; }

    /// <summary>Bottom-right of a range; null unless Kind is Range.</summary>
    public CellRef? End { get; init; }

    /// <summary>
    /// For an <see cref="PathSegmentKind.ElementSpan"/>, the raw start/end tokens
    /// as written (<c>"2"</c>/<c>"6"</c> or <c>"B"</c>/<c>"E"</c>); null otherwise.
    /// The handler interprets them against the document.
    /// </summary>
    public string? SpanFrom { get; init; }

    public string? SpanTo { get; init; }

    public string ToCanonicalString() => Kind switch
    {
        PathSegmentKind.Element => Id is { } id
            ? string.Create(CultureInfo.InvariantCulture, $"{Name}[@{IdAttribute ?? "id"}={id}]")
            : Variant is { } variant
                ? string.Create(CultureInfo.InvariantCulture, $"{Name}[{variant}]")
                : Letter is { } letter
                    ? string.Create(CultureInfo.InvariantCulture, $"{Name}[{letter}]")
                    : Index is { } i
                        ? string.Create(CultureInfo.InvariantCulture, $"{Name}[{i}]")
                        : Name!,
        PathSegmentKind.Name => Quote(Name!),
        PathSegmentKind.Cell => Start!.Value.ToString(),
        PathSegmentKind.Range => $"{Start!.Value}:{End!.Value}",
        PathSegmentKind.ElementSpan => $"{Name}[{SpanFrom}]:{Name}[{SpanTo}]",
        _ => throw new InvalidOperationException($"Unknown segment kind: {Kind}"),
    };

    internal static string Quote(string name) =>
        DocPath.BareNameRegex.IsMatch(name) && !DocPath.LooksLikeCellOrRange(name)
            ? name
            : "'" + name.Replace("'", "''", StringComparison.Ordinal) + "'";
}

/// <summary>
/// A parsed document path. Grammar (1-based indices throughout):
/// <code>
/// docx:  /body/p[3]   /body/table[1]/tr[2]/tc[1]   /body/p[3]/run[2]   /header[1]/p[1]
/// xlsx:  /Sheet1/A1   /Sheet1/A1:C10   /Sheet1/row[3]   /Sheet1/col[C]   /'Q3 Data'/B2
/// pptx:  /slide[2]    /slide[2]/shape[3]   /slide[2]/shape[3]/p[1]
/// </code>
/// </summary>
public sealed partial record DocPath
{
    [GeneratedRegex(@"^[A-Za-z_][A-Za-z0-9_.-]*$")]
    private static partial Regex BareName();

    [GeneratedRegex(@"^([A-Za-z_][A-Za-z0-9_.-]*)\[([0-9]+)\]$")]
    private static partial Regex IndexedElement();

    [GeneratedRegex(@"^([A-Za-z_][A-Za-z0-9_.-]*)\[([A-Z]{1,3})\]$")]
    private static partial Regex LetterElement();

    [GeneratedRegex(@"^([A-Za-z_][A-Za-z0-9_.-]*)\[@(id|name|tag)=([A-Za-z0-9_.-]+)\]$")]
    private static partial Regex IdElement();

    // Named variants are a closed vocabulary (docx M5 header/footer types), so
    // arbitrary words in brackets (p[abc]) stay rejected as malformed indices.
    [GeneratedRegex(@"^([A-Za-z_][A-Za-z0-9_.-]*)\[(default|firstPage|even)\]$")]
    private static partial Regex NamedVariantElement();

    // A contiguous element span: same element name on both ends, numeric bounds
    // (row[2]:row[6]) or letter bounds (col[B]:col[E]). The handler resolves it.
    [GeneratedRegex(@"^([A-Za-z_][A-Za-z0-9_.-]*)\[([0-9]+|[A-Z]{1,3})\]:\1\[([0-9]+|[A-Z]{1,3})\]$")]
    private static partial Regex ElementSpanPattern();

    [GeneratedRegex(@"^([A-Z]{1,3})([0-9]{1,7})$")]
    private static partial Regex CellPattern();

    [GeneratedRegex(@"^([A-Z]{1,3})([0-9]{1,7}):([A-Z]{1,3})([0-9]{1,7})$")]
    private static partial Regex RangePattern();

    internal static Regex BareNameRegex => BareName();

    internal static bool LooksLikeCellOrRange(string text) =>
        CellPattern().IsMatch(text) || RangePattern().IsMatch(text);

    private const string GrammarHint =
        "Paths look like /body/p[3], /Sheet1/A1:C10, /'Q3 Data'/B2 or /slide[2]/shape[3]; " +
        "indices are 1-based. Run 'aioffice help addressing' for the full grammar.";

    public required IReadOnlyList<PathSegment> Segments { get; init; }

    /// <summary>
    /// True for the document root <c>"/"</c> (zero segments): pptx slide size and
    /// slide-section management, and document-level <c>get</c>. Handlers that have
    /// no document-level surface reject it with <c>unsupported_feature</c>.
    /// </summary>
    public bool IsRoot => Segments.Count == 0;

    public string ToCanonicalString() =>
        Segments.Count == 0 ? "/" : "/" + string.Join('/', Segments.Select(s => s.ToCanonicalString()));

    public override string ToString() => ToCanonicalString();

    /// <summary>Parses a path or throws <c>invalid_path</c> with a grammar hint.</summary>
    public static DocPath Parse(string text)
    {
        if (!TryParse(text, out var path, out var error))
        {
            throw new AiofficeException(ErrorCodes.InvalidPath, error!, GrammarHint);
        }

        return path!;
    }

    public static bool TryParse(string? text, out DocPath? path, out string? error)
    {
        path = null;
        error = null;

        if (string.IsNullOrWhiteSpace(text))
        {
            error = "Path is empty.";
            return false;
        }

        if (text[0] != '/')
        {
            error = $"Path must start with '/': {text}";
            return false;
        }

        // "/" alone is the document root: pptx slide size + sections and
        // document-level get. It carries zero segments.
        if (text.Length == 1)
        {
            path = new DocPath { Segments = [] };
            return true;
        }

        var rawSegments = SplitSegments(text, ref error);
        if (rawSegments is null)
        {
            return false;
        }

        if (rawSegments.Count == 0)
        {
            error = "Path has no segments.";
            return false;
        }

        var segments = new List<PathSegment>(rawSegments.Count);
        foreach (var (raw, wasQuoted) in rawSegments)
        {
            var segment = ParseSegment(raw, wasQuoted, ref error);
            if (segment is null)
            {
                return false;
            }

            segments.Add(segment);
        }

        path = new DocPath { Segments = segments };
        return true;
    }

    private static PathSegment? ParseSegment(string raw, bool wasQuoted, ref string? error)
    {
        if (raw.Length == 0)
        {
            error = "Path contains an empty segment (double '/'?).";
            return null;
        }

        if (wasQuoted)
        {
            return new PathSegment { Kind = PathSegmentKind.Name, Name = raw };
        }

        // Element spans (xlsx outline groups) carry both bounds in one segment:
        // row[2]:row[6], col[B]:col[E]. The bounds must be the same kind (both
        // numeric or both letters); the handler interprets them.
        var span = ElementSpanPattern().Match(raw);
        if (span.Success)
        {
            var from = span.Groups[2].Value;
            var to = span.Groups[3].Value;
            var fromNumeric = char.IsDigit(from[0]);
            if (fromNumeric != char.IsDigit(to[0]))
            {
                error = $"Span bounds must both be numeric or both be column letters: {raw}";
                return null;
            }

            return new PathSegment
            {
                Kind = PathSegmentKind.ElementSpan,
                Name = span.Groups[1].Value,
                SpanFrom = from,
                SpanTo = to,
            };
        }

        var range = RangePattern().Match(raw);
        if (range.Success)
        {
            var start = new CellRef(range.Groups[1].Value, int.Parse(range.Groups[2].Value, CultureInfo.InvariantCulture));
            var end = new CellRef(range.Groups[3].Value, int.Parse(range.Groups[4].Value, CultureInfo.InvariantCulture));
            if (start.Row < 1 || end.Row < 1)
            {
                error = $"Range rows are 1-based: {raw}";
                return null;
            }

            if (start.ColumnNumber > end.ColumnNumber || start.Row > end.Row)
            {
                error = $"Range start must not be past its end: {raw}";
                return null;
            }

            return new PathSegment { Kind = PathSegmentKind.Range, Start = start, End = end };
        }

        var cell = CellPattern().Match(raw);
        if (cell.Success)
        {
            var cellRef = new CellRef(cell.Groups[1].Value, int.Parse(cell.Groups[2].Value, CultureInfo.InvariantCulture));
            if (cellRef.Row < 1)
            {
                error = $"Cell rows are 1-based: {raw}";
                return null;
            }

            return new PathSegment { Kind = PathSegmentKind.Cell, Start = cellRef };
        }

        var byId = IdElement().Match(raw);
        if (byId.Success)
        {
            return new PathSegment
            {
                Kind = PathSegmentKind.Element,
                Name = byId.Groups[1].Value,
                IdAttribute = byId.Groups[2].Value,
                Id = byId.Groups[3].Value,
            };
        }

        var variant = NamedVariantElement().Match(raw);
        if (variant.Success)
        {
            return new PathSegment
            {
                Kind = PathSegmentKind.Element,
                Name = variant.Groups[1].Value,
                Variant = variant.Groups[2].Value,
            };
        }

        var indexed = IndexedElement().Match(raw);
        if (indexed.Success)
        {
            var index = int.Parse(indexed.Groups[2].Value, CultureInfo.InvariantCulture);
            if (index < 1)
            {
                error = $"Indices are 1-based; got [{index}] in: {raw}";
                return null;
            }

            return new PathSegment { Kind = PathSegmentKind.Element, Name = indexed.Groups[1].Value, Index = index };
        }

        // Letter-indexed elements address spreadsheet columns: col[C], col[AB].
        var lettered = LetterElement().Match(raw);
        if (lettered.Success)
        {
            return new PathSegment
            {
                Kind = PathSegmentKind.Element,
                Name = lettered.Groups[1].Value,
                Letter = lettered.Groups[2].Value,
            };
        }

        if (BareName().IsMatch(raw))
        {
            return new PathSegment { Kind = PathSegmentKind.Element, Name = raw };
        }

        error = $"Unrecognized path segment: '{raw}'";
        return null;
    }

    /// <summary>Splits on '/' while honoring single-quoted names ('' escapes a quote).</summary>
    private static List<(string Text, bool WasQuoted)>? SplitSegments(string text, ref string? error)
    {
        var segments = new List<(string, bool)>();
        var i = 1; // skip leading '/'
        while (i < text.Length)
        {
            if (text[i] == '\'')
            {
                var sb = new StringBuilder();
                i++; // consume opening quote
                var closed = false;
                while (i < text.Length)
                {
                    if (text[i] == '\'')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '\'')
                        {
                            sb.Append('\'');
                            i += 2;
                            continue;
                        }

                        closed = true;
                        i++;
                        break;
                    }

                    sb.Append(text[i]);
                    i++;
                }

                if (!closed)
                {
                    error = $"Unterminated quoted name in path: {text}";
                    return null;
                }

                if (i < text.Length && text[i] != '/')
                {
                    error = $"Quoted name must be a whole segment: {text}";
                    return null;
                }

                if (sb.Length == 0)
                {
                    error = "Quoted name is empty.";
                    return null;
                }

                segments.Add((sb.ToString(), true));
                i++; // consume '/' (or move past end)
            }
            else
            {
                var slash = text.IndexOf('/', i);
                var end = slash < 0 ? text.Length : slash;
                segments.Add((text[i..end], false));
                i = end + 1;
            }
        }

        if (text[^1] == '/')
        {
            error = $"Path must not end with '/': {text}";
            return null;
        }

        return segments;
    }
}
