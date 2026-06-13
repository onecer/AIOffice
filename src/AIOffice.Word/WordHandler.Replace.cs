using System.Text;
using System.Text.RegularExpressions;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// One inline content piece of a paragraph: a w:t (length = its text) or a
/// non-text inline element (break, tab, field char, footnote reference, …)
/// that occupies one object-replacement character in the match string so
/// literal finds can never silently span it.
/// </summary>
internal sealed record InlinePiece(OpenXmlElement Element, Run Run, int Start, int Length);

public sealed partial class WordHandler
{
    /// <summary>Placeholder for non-text inline content in the concatenated match text.</summary>
    private const char InlineObjectChar = '￼';

    private static readonly TimeSpan RegexBudget = TimeSpan.FromSeconds(2);

    /// <summary>
    /// The shared find/replace contract:
    /// <c>{"op":"replace","path":scope,"props":{"find","replace","regex","matchCase","wholeWord"}}</c>.
    /// The scope is any container path (/body, /body/p[3], /header[1], /footer[1],
    /// tables and cells included). Matching runs against the concatenated text of
    /// each paragraph, so finds that Word users see as one string match even when
    /// the XML splits them across runs; rewritten text keeps the formatting of the
    /// first affected run. With ctx track=true every replacement becomes a
    /// w:del + w:ins revision pair. Zero matches is ok:true plus a find_no_match
    /// meta warning.
    /// </summary>
    private static object ApplyReplace(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        var author = session.ResolveAuthor(props);

        var find = props["find"] is { } findNode ? NodeToString(findNode) : null;
        if (string.IsNullOrEmpty(find))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "replace needs props.find (the text or pattern to search for).",
                "Example: {\"op\":\"replace\",\"path\":\"/body\",\"props\":{\"find\":\"2025\",\"replace\":\"2026\"}}.");
        }

        var replace = props["replace"] is { } replaceNode ? NodeToString(replaceNode) : string.Empty;
        var isRegex = PropFlag(props, "regex");
        var matchCase = PropFlag(props, "matchCase");
        var wholeWord = PropFlag(props, "wholeWord");

        var scope = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (scope.Element is Run or Hyperlink)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"replace scopes are containers (/body, /header[n], /footer[n], a paragraph, table or cell), not '{scope.Type}'.",
                $"Target the paragraph instead: {scope.CanonicalPath[..scope.CanonicalPath.LastIndexOf('/')]}.");
        }

        if (session.Track)
        {
            RequireBodyScope(scope.CanonicalPath, "replace");
        }

        var regex = BuildFindRegex(find, isRegex, matchCase, wholeWord);

        // Canonical paths for location reporting, captured before any mutation
        // (replace never adds or removes paragraphs, so they stay valid).
        var paths = new Dictionary<OpenXmlElement, string>(ReferenceEqualityComparer.Instance);
        foreach (var node in WordAddress.EnumerateAll(doc))
        {
            paths[node.Element] = node.CanonicalPath;
        }

        List<Paragraph> paragraphs = scope.Element is Paragraph scopeParagraph
            ? [scopeParagraph]
            : [.. scope.Element.Descendants<Paragraph>()];

        var replacements = 0;
        var locations = new List<string>();
        var date = DateTime.UtcNow;
        var nextId = session.Track ? NextRevisionId(doc) : 0;

        foreach (var paragraph in paragraphs)
        {
            ComputeInlinePieces(paragraph, out var text);
            if (text.Length == 0)
            {
                continue;
            }

            List<Match> matches;
            try
            {
                // Zero-length matches (e.g. regex "x*") are insertion points, not text; skip them.
                matches = regex.Matches(text).Where(m => m.Length > 0).ToList();
            }
            catch (RegexMatchTimeoutException ex)
            {
                throw RegexTimeout(find, ex);
            }

            if (matches.Count == 0)
            {
                continue;
            }

            var paragraphPath = paths.TryGetValue(paragraph, out var known) ? known : scope.CanonicalPath;
            foreach (var _ in matches)
            {
                if (locations.Count < 20)
                {
                    locations.Add(paragraphPath);
                }
            }

            replacements += matches.Count;

            // Last match first, so earlier match offsets stay valid in the
            // (unchanged) prefix of the concatenated text.
            for (var i = matches.Count - 1; i >= 0; i--)
            {
                var match = matches[i];
                var replacement = isRegex ? match.Result(replace) : replace;
                ReplaceRange(paragraph, match.Index, match.Index + match.Length, replacement, session.Track, author, date, ref nextId);
            }
        }

        if (replacements == 0)
        {
            session.Warnings.Add(new Warning(
                ErrorCodes.FindNoMatch,
                $"replace: no match for '{find}' in {scope.CanonicalPath}; 0 replacements made."));
        }

        return new
        {
            op = "replace",
            path = scope.CanonicalPath,
            replacements,
            locations,
            tracked = session.Track ? true : (bool?)null,
        };
    }

    private static bool PropFlag(System.Text.Json.Nodes.JsonObject props, string name) =>
        props[name] is { } node && WordFormatting.ParseBool(name, NodeToString(node));

    private static Regex BuildFindRegex(string find, bool isRegex, bool matchCase, bool wholeWord)
    {
        var pattern = isRegex ? find : Regex.Escape(find);
        if (wholeWord)
        {
            // Non-capturing wrap keeps user group numbering intact in regex mode.
            pattern = @"(?<!\w)(?:" + pattern + @")(?!\w)";
        }

        var options = RegexOptions.CultureInvariant;
        if (!matchCase)
        {
            options |= RegexOptions.IgnoreCase;
        }

        try
        {
            return new Regex(pattern, options, RegexBudget);
        }
        catch (ArgumentException ex)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.find is not a valid .NET regex: {ex.Message}",
                "Fix the pattern, or drop regex:true to search for the literal text.",
                innerException: ex);
        }
    }

    private static AiofficeException RegexTimeout(string find, Exception inner) => new(
        ErrorCodes.InvalidArgs,
        $"regex_timeout: pattern '{find}' exceeded the 2s match budget.",
        "Simplify the regex (catastrophic backtracking is the usual cause), anchor it, or replace literal text instead.",
        innerException: inner);

    // ----------------------------------------------------------- match model

    /// <summary>
    /// The paragraph's inline pieces in document order plus the concatenated
    /// match text. Pending deletions (w:del) are not current content and are
    /// excluded; pending insertions (w:ins) and hyperlink runs are included.
    /// </summary>
    private static List<InlinePiece> ComputeInlinePieces(Paragraph paragraph, out string text)
    {
        var pieces = new List<InlinePiece>();
        var sb = new StringBuilder();

        foreach (var run in paragraph.Descendants<Run>())
        {
            if (IsInsideDeleted(run, paragraph))
            {
                continue;
            }

            foreach (var child in run.ChildElements)
            {
                switch (child)
                {
                    case Text t:
                        pieces.Add(new InlinePiece(t, run, sb.Length, t.Text.Length));
                        sb.Append(t.Text);
                        break;

                    case Break or TabChar or NoBreakHyphen or SymbolChar or PositionalTab
                        or FootnoteReference or EndnoteReference or FootnoteReferenceMark or EndnoteReferenceMark
                        or CommentReference or Drawing or FieldChar or FieldCode
                        or DocumentFormat.OpenXml.Wordprocessing.EmbeddedObject or Picture:
                        pieces.Add(new InlinePiece(child, run, sb.Length, 1));
                        sb.Append(InlineObjectChar);
                        break;

                    default: // rPr, lastRenderedPageBreak, proofing marks, … — zero width
                        break;
                }
            }
        }

        text = sb.ToString();
        return pieces;
    }

    private static bool IsInsideDeleted(Run run, Paragraph paragraph)
    {
        for (var node = run.Parent; node is not null && !ReferenceEquals(node, paragraph); node = node.Parent)
        {
            if (node is DeletedRun)
            {
                return true;
            }
        }

        return false;
    }

    // -------------------------------------------------------------- rewrite

    /// <summary>
    /// Rewrites the [start, end) range of the paragraph's concatenated text:
    /// run boundaries are split so the range covers whole runs, then the runs
    /// are replaced by one run with the new text (formatting cloned from the
    /// first affected run) — or wrapped in w:del with a w:ins when tracking.
    /// </summary>
    private static void ReplaceRange(
        Paragraph paragraph, int start, int end, string replacement, bool track, string author, DateTime date, ref int nextId)
    {
        EnsureRunBoundary(paragraph, end);
        EnsureRunBoundary(paragraph, start);

        var pieces = ComputeInlinePieces(paragraph, out _);
        var affected = new List<Run>();
        foreach (var piece in pieces)
        {
            if (piece.Length > 0 && piece.Start >= start && piece.Start + piece.Length <= end &&
                (affected.Count == 0 || !ReferenceEquals(affected[^1], piece.Run)))
            {
                affected.Add(piece.Run);
            }
        }

        if (affected.Count == 0)
        {
            return; // defensive: boundary splitting guarantees coverage
        }

        if (track)
        {
            TrackedReplaceMatchedRuns(affected, replacement, author, date, ref nextId);
        }
        else
        {
            ReplaceMatchedRuns(affected, replacement);
        }
    }

    /// <summary>Untracked rewrite: one replacement run (first affected run's formatting), matched runs removed.</summary>
    private static void ReplaceMatchedRuns(List<Run> runs, string replacement)
    {
        if (replacement.Length > 0)
        {
            var replacementRun = new Run();
            if (runs[0].RunProperties is { } rPr)
            {
                replacementRun.RunProperties = (RunProperties)rPr.CloneNode(true);
            }

            replacementRun.AppendChild(NewText(replacement));
            runs[0].InsertBeforeSelf(replacementRun);
        }

        foreach (var run in runs)
        {
            run.Remove();
        }
    }

    /// <summary>
    /// Tracked rewrite: contiguous sibling groups of matched runs each go into
    /// one w:del (w:t becomes w:delText); the replacement lands in a w:ins right
    /// after the first w:del — the delete-old + insert-new revision pair.
    /// </summary>
    private static void TrackedReplaceMatchedRuns(
        List<Run> runs, string replacement, string author, DateTime date, ref int nextId)
    {
        var keepFormatting = runs[0].RunProperties?.CloneNode(true) as RunProperties;

        DeletedRun? firstDel = null;
        DeletedRun? current = null;
        foreach (var run in runs)
        {
            if (current is null ||
                !ReferenceEquals(current.Parent, run.Parent) ||
                !ReferenceEquals(current.NextSibling(), run))
            {
                current = NewTrackChange(new DeletedRun(), author, date, nextId++);
                run.InsertBeforeSelf(current);
                firstDel ??= current;
            }

            run.Remove();
            ConvertTextToDeleted(run);
            current.AppendChild(run);
        }

        if (replacement.Length > 0 && firstDel is not null)
        {
            var insertedRun = new Run();
            if (keepFormatting is not null)
            {
                insertedRun.RunProperties = keepFormatting;
            }

            insertedRun.AppendChild(NewText(replacement));
            var ins = NewTrackChange(new InsertedRun(), author, date, nextId++);
            ins.AppendChild(insertedRun);
            firstDel.InsertAfterSelf(ins);
        }
    }

    // ----------------------------------------------------------- run splits

    /// <summary>
    /// Guarantees a run boundary exists at the given offset of the paragraph's
    /// concatenated text, splitting a w:t (and its run) or a run between two
    /// inline children when the offset falls inside one.
    /// </summary>
    private static void EnsureRunBoundary(Paragraph paragraph, int offset)
    {
        var pieces = ComputeInlinePieces(paragraph, out var text);
        if (offset <= 0 || offset >= text.Length)
        {
            return;
        }

        var inside = pieces.FirstOrDefault(p => p.Start < offset && offset < p.Start + p.Length);
        if (inside is not null)
        {
            SplitRunAtText(inside.Run, (Text)inside.Element, offset - inside.Start);
            return;
        }

        var before = pieces.LastOrDefault(p => p.Length > 0 && p.Start + p.Length == offset);
        var after = pieces.FirstOrDefault(p => p.Length > 0 && p.Start == offset);
        if (before is not null && after is not null && ReferenceEquals(before.Run, after.Run))
        {
            SplitRunBetweenChildren(before.Run, after.Element);
        }
    }

    /// <summary>Splits a run inside a w:t: the suffix text and later children move into a cloned-format sibling run.</summary>
    private static void SplitRunAtText(Run run, Text text, int localOffset)
    {
        var suffix = text.Text[localOffset..];
        SetTextValue(text, text.Text[..localOffset]);

        var newRun = new Run();
        if (run.RunProperties is { } rPr)
        {
            newRun.RunProperties = (RunProperties)rPr.CloneNode(true);
        }

        var moving = run.ChildElements.SkipWhile(c => !ReferenceEquals(c, text)).Skip(1).ToList();
        newRun.AppendChild(NewText(suffix));
        foreach (var child in moving)
        {
            child.Remove();
            newRun.AppendChild(child);
        }

        run.InsertAfterSelf(newRun);
    }

    /// <summary>Splits a run between two children: splitChild and everything after it move into a cloned-format sibling run.</summary>
    private static void SplitRunBetweenChildren(Run run, OpenXmlElement splitChild)
    {
        var newRun = new Run();
        if (run.RunProperties is { } rPr)
        {
            newRun.RunProperties = (RunProperties)rPr.CloneNode(true);
        }

        var moving = run.ChildElements.SkipWhile(c => !ReferenceEquals(c, splitChild)).ToList();
        foreach (var child in moving)
        {
            child.Remove();
            newRun.AppendChild(child);
        }

        run.InsertAfterSelf(newRun);
    }

    private static void SetTextValue(Text text, string value)
    {
        text.Text = value;
        text.Space = value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]))
            ? SpaceProcessingModeValues.Preserve
            : null;
    }
}
