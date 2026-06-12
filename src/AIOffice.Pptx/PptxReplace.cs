using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>The per-op result of a replace: match count plus up to 20 canonical locations.</summary>
internal sealed record PptxReplaceResult(string Target, int Replacements, IReadOnlyList<string> Locations);

/// <summary>
/// The shared M4 find/replace contract for pptx:
/// <c>{"op":"replace","path":"/slide[2]" | "/slide[2]/shape[@id=N]" | "/slide[2]/notes",
/// "props":{"find","replace","regex","matchCase","wholeWord","includeNotes"}}</c>.
/// Matching runs over the concatenated paragraph text so finds split across runs
/// are found; rewritten spans keep the formatting of the first affected run.
/// Deck-wide scopes are expanded to per-slide ops by the command layer.
/// </summary>
internal static class PptxReplace
{
    private static readonly IReadOnlyList<string> PropKeys =
        ["find", "replace", "regex", "matchCase", "wholeWord", "includeNotes"];

    private const int MaxLocations = 20;

    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);

    /// <summary>Applies one replace op and returns its target, count and locations.</summary>
    public static PptxReplaceResult Apply(PresentationPart presentation, EditOp op)
    {
        var address = PptxAddress.Parse(op.Path);
        if (address.IsChart || address.IsTable || address.IsSmartArt || address.IsAnimation || address.IsComment ||
            address.ParagraphIndex is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"replace scopes to a slide, a shape or the notes, not '{op.Path}'.",
                "Use /slide[2] (all its text), /slide[2]/shape[@id=N] (one shape) or /slide[2]/notes; " +
                "set table cell text via {\"op\":\"set\",...,\"props\":{\"text\":\"...\"}}.");
        }

        var spec = ParseSpec(op, allowIncludeNotes: !address.HasShape && !address.IsNotes);
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, op.Path);

        var replacements = 0;
        var locations = new List<string>();

        if (address.IsNotes)
        {
            ReplaceInNotes(slidePart, address, spec, ref replacements, locations);
            return new PptxReplaceResult(address.CanonicalNotesPath, replacements, locations);
        }

        if (address.HasShape)
        {
            var view = PptxDoc.ResolveShape(slidePart, address);
            if (view.Element is not P.Shape shape || shape.TextBody is null)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"'{view.Kind}' shapes have no replaceable text.",
                    "Scope replace to a text shape, or to the whole slide: {\"op\":\"replace\",\"path\":\"" +
                    address.CanonicalSlidePath + "\",\"props\":{...}}.");
            }

            ReplaceInShape(shape, view.CanonicalPath(address.SlideIndex), spec, ref replacements, locations);
            return new PptxReplaceResult(view.CanonicalPath(address.SlideIndex), replacements, locations);
        }

        foreach (var view in PptxDoc.Shapes(slidePart))
        {
            if (view.Element is P.Shape shape && shape.TextBody is not null)
            {
                ReplaceInShape(shape, view.CanonicalPath(address.SlideIndex), spec, ref replacements, locations);
            }
        }

        if (spec.IncludeNotes)
        {
            ReplaceInNotes(slidePart, address, spec, ref replacements, locations);
        }

        return new PptxReplaceResult(address.CanonicalSlidePath, replacements, locations);
    }

    // ----- props ---------------------------------------------------------------------

    private sealed record ReplaceSpec(Regex Pattern, string Replacement, bool IsRegex, bool IncludeNotes);

    private static ReplaceSpec ParseSpec(EditOp op, bool allowIncludeNotes)
    {
        var props = op.Props ?? throw MissingFind(op);
        foreach (var (key, _) in props)
        {
            if (!PropKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown replace prop '{key}'.",
                    "Replace props: find, replace, regex, matchCase, wholeWord, includeNotes.",
                    candidates: PropKeys);
            }

            if (!allowIncludeNotes && key == "includeNotes")
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Prop 'includeNotes' does not apply to scope '{op.Path}'.",
                    "includeNotes belongs to slide scopes (/slide[2]); shape and notes scopes are already exact.");
            }
        }

        if (!props.TryGetPropertyValue("find", out var findNode) || findNode is null ||
            J.ScalarText(findNode).Length == 0)
        {
            throw MissingFind(op);
        }

        var find = J.ScalarText(findNode);
        var replacement = props.TryGetPropertyValue("replace", out var replaceNode) && replaceNode is not null
            ? J.ScalarText(replaceNode)
            : string.Empty;
        var isRegex = Flag(props, "regex");
        var matchCase = Flag(props, "matchCase");
        var wholeWord = Flag(props, "wholeWord");
        var includeNotes = !props.TryGetPropertyValue("includeNotes", out _) || Flag(props, "includeNotes");

        var pattern = isRegex ? find : Regex.Escape(find);
        if (wholeWord)
        {
            pattern = Units.Inv($@"\b(?:{pattern})\b");
        }

        var options = RegexOptions.CultureInvariant | (matchCase ? RegexOptions.None : RegexOptions.IgnoreCase);
        Regex regex;
        try
        {
            regex = new Regex(pattern, options, RegexTimeout);
        }
        catch (ArgumentException ex)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.find is not a valid regular expression: {ex.Message}",
                "Fix the pattern, or drop \"regex\":true to search for the literal text.",
                innerException: ex);
        }

        return new ReplaceSpec(regex, replacement, isRegex, includeNotes);
    }

    private static AiofficeException MissingFind(EditOp op) => new(
        ErrorCodes.InvalidArgs,
        "replace requires a non-empty props.find.",
        "The full shape is {\"op\":\"replace\",\"path\":\"" + op.Path + "\",\"props\":{\"find\":\"Q3\"," +
        "\"replace\":\"Q4\",\"regex\":false,\"matchCase\":false,\"wholeWord\":false}}.");

    private static bool Flag(JsonObject props, string key)
    {
        if (!props.TryGetPropertyValue(key, out var node) || node is null)
        {
            return false;
        }

        if (node is JsonValue value)
        {
            if (value.TryGetValue<bool>(out var flag))
            {
                return flag;
            }

            if (value.TryGetValue<string>(out var text) && bool.TryParse(text, out var parsed))
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Replace prop '{key}' is not a boolean: {node.ToJsonString()}",
            "Use true or false.");
    }

    // ----- rewrite -------------------------------------------------------------------

    private static void ReplaceInNotes(
        SlidePart slidePart, PptxAddress address, ReplaceSpec spec, ref int replacements, List<string> locations)
    {
        if (PptxNotes.Body(slidePart)?.TextBody is not { } body)
        {
            return;
        }

        var count = 0;
        foreach (var paragraph in body.Elements<A.Paragraph>().ToList())
        {
            count += ReplaceInParagraph(paragraph, spec);
        }

        if (count > 0)
        {
            replacements += count;
            AddLocation(locations, address.CanonicalNotesPath);
        }
    }

    private static void ReplaceInShape(
        P.Shape shape, string canonicalPath, ReplaceSpec spec, ref int replacements, List<string> locations)
    {
        var paragraphs = shape.TextBody!.Elements<A.Paragraph>().ToList();
        for (var i = 0; i < paragraphs.Count; i++)
        {
            var count = ReplaceInParagraph(paragraphs[i], spec);
            if (count > 0)
            {
                replacements += count;
                AddLocation(locations, Units.Inv($"{canonicalPath}/p[{i + 1}]"));
            }
        }
    }

    private static void AddLocation(List<string> locations, string path)
    {
        if (locations.Count < MaxLocations)
        {
            locations.Add(path);
        }
    }

    /// <summary>
    /// Replaces within one paragraph. Runs separated by breaks/fields form
    /// independent segments; within a segment the match runs over the
    /// concatenated text, so finds split across runs are rewritten correctly.
    /// </summary>
    private static int ReplaceInParagraph(A.Paragraph paragraph, ReplaceSpec spec)
    {
        var total = 0;
        var segment = new List<A.Run>();
        foreach (var child in paragraph.ChildElements.ToList())
        {
            if (child is A.Run run && run.Text is not null)
            {
                segment.Add(run);
                continue;
            }

            total += ReplaceSegment(paragraph, segment, spec);
            segment.Clear();
        }

        total += ReplaceSegment(paragraph, segment, spec);
        return total;
    }

    private static int ReplaceSegment(A.Paragraph paragraph, List<A.Run> runs, ReplaceSpec spec)
    {
        if (runs.Count == 0)
        {
            return 0;
        }

        var text = string.Concat(runs.Select(r => r.Text!.Text));
        if (text.Length == 0)
        {
            return 0;
        }

        MatchCollection matches;
        try
        {
            matches = spec.Pattern.Matches(text);
            if (matches.Count == 0)
            {
                return 0;
            }
        }
        catch (RegexMatchTimeoutException ex)
        {
            throw RegexTimedOut(ex);
        }

        // Run start offsets in the concatenated segment text.
        var starts = new int[runs.Count];
        for (var i = 1; i < runs.Count; i++)
        {
            starts[i] = starts[i - 1] + runs[i - 1].Text!.Text.Length;
        }

        var pieces = new List<(string Text, A.RunProperties? Props)>();
        var pos = 0;
        try
        {
            foreach (Match match in matches)
            {
                AppendSpan(pieces, runs, starts, text, pos, match.Index);
                var replaced = spec.IsRegex ? match.Result(spec.Replacement) : spec.Replacement;
                pieces.Add((replaced, RunAt(runs, starts, match.Index).RunProperties));
                pos = match.Index + match.Length;
            }
        }
        catch (RegexMatchTimeoutException ex)
        {
            throw RegexTimedOut(ex); // lazy MatchCollection: the timeout can also fire during enumeration
        }

        AppendSpan(pieces, runs, starts, text, pos, text.Length);

        var anchor = runs[0];
        foreach (var (pieceText, props) in pieces)
        {
            if (pieceText.Length == 0)
            {
                continue;
            }

            var run = new A.Run();
            if (props is not null)
            {
                run.Append((A.RunProperties)props.CloneNode(true));
            }

            run.Append(new A.Text(pieceText));
            paragraph.InsertBefore(run, anchor);
        }

        foreach (var run in runs)
        {
            run.Remove();
        }

        return matches.Count;
    }

    private static AiofficeException RegexTimedOut(RegexMatchTimeoutException ex) => new(
        ErrorCodes.InvalidArgs,
        "The regex timed out after 2 seconds (catastrophic backtracking?).",
        "Simplify the pattern (avoid nested quantifiers like (a+)+), or search for literal text without \"regex\":true.",
        innerException: ex);

    /// <summary>Emits the [from, to) span as one piece per overlapped run, keeping each run's formatting.</summary>
    private static void AppendSpan(
        List<(string Text, A.RunProperties? Props)> pieces,
        List<A.Run> runs,
        int[] starts,
        string text,
        int from,
        int to)
    {
        for (var i = 0; i < runs.Count && from < to; i++)
        {
            var runStart = starts[i];
            var runEnd = runStart + runs[i].Text!.Text.Length;
            var sliceFrom = Math.Max(from, runStart);
            var sliceTo = Math.Min(to, runEnd);
            if (sliceTo > sliceFrom)
            {
                pieces.Add((text[sliceFrom..sliceTo], runs[i].RunProperties));
            }
        }
    }

    /// <summary>The run containing the given text offset (clamped to the last run at the end).</summary>
    private static A.Run RunAt(List<A.Run> runs, int[] starts, int offset)
    {
        for (var i = runs.Count - 1; i >= 0; i--)
        {
            if (offset >= starts[i])
            {
                return runs[i];
            }
        }

        return runs[0];
    }
}
