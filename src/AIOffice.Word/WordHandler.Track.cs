using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>One tracked revision: a w:ins / w:del wrapper, a paragraph-mark
/// ins/del, or a (read-only in M2) formatting change.</summary>
internal sealed record RevisionRecord(
    OpenXmlElement Element,
    int Id,
    string Kind, // insert | delete | format
    string? Author,
    string? Date,
    string Text,
    string At,
    bool ParagraphMark);

public sealed partial class WordHandler
{
    // ------------------------------------------------------------ enumeration

    /// <summary>
    /// Every revision in the document in document order (body first, then
    /// headers/footers): w:ins / w:del run wrappers, paragraph-mark ins/del,
    /// and formatting changes (w:pPrChange / w:rPrChange) as kind "format".
    /// </summary>
    private static List<RevisionRecord> EnumerateRevisions(WordprocessingDocument doc)
    {
        // Canonical paths for the nearest addressable ancestor ("at").
        var paths = new Dictionary<OpenXmlElement, string>(ReferenceEqualityComparer.Instance);
        foreach (var node in WordAddress.EnumerateAll(doc))
        {
            paths[node.Element] = node.CanonicalPath;
        }

        var roots = new List<(OpenXmlElement Root, string Path)>();
        if (doc.MainDocumentPart?.Document?.Body is { } body)
        {
            roots.Add((body, "/body"));
        }

        foreach (var hf in WordAddress.HeaderFooterRoots(doc))
        {
            roots.Add((hf.Element, hf.CanonicalPath));
        }

        var revisions = new List<RevisionRecord>();
        foreach (var (root, rootPath) in roots)
        {
            foreach (var element in root.Descendants())
            {
                var record = element switch
                {
                    InsertedRun ins => Describe(ins, "insert", ins.InnerText, paragraphMark: false),
                    DeletedRun del => Describe(del, "delete", del.InnerText, paragraphMark: false),
                    Inserted insMark when insMark.Parent is ParagraphMarkRunProperties =>
                        Describe(insMark, "insert", string.Empty, paragraphMark: true),
                    Deleted delMark when delMark.Parent is ParagraphMarkRunProperties =>
                        Describe(delMark, "delete", string.Empty, paragraphMark: true),
                    ParagraphPropertiesChange pPrChange => Describe(pPrChange, "format", string.Empty, paragraphMark: false),
                    RunPropertiesChange rPrChange => Describe(rPrChange, "format", string.Empty, paragraphMark: false),
                    _ => null,
                };

                if (record is not null)
                {
                    revisions.Add(record);
                }
            }

            RevisionRecord? Describe(OpenXmlElement element, string kind, string text, bool paragraphMark)
            {
                var (id, author, date) = TrackChangeAttributes(element);
                return new RevisionRecord(element, id, kind, author, date, text, NearestPath(element), paragraphMark);
            }

            string NearestPath(OpenXmlElement element)
            {
                for (var node = element.Parent; node is not null; node = node.Parent)
                {
                    if (paths.TryGetValue(node, out var path))
                    {
                        return path;
                    }
                }

                return rootPath;
            }
        }

        return revisions;
    }

    /// <summary>w:id / w:author / w:date off any track-change element.</summary>
    private static (int Id, string? Author, string? Date) TrackChangeAttributes(OpenXmlElement element)
    {
        string? id = null, author = null, date = null;
        foreach (var attribute in element.GetAttributes())
        {
            switch (attribute.LocalName)
            {
                case "id":
                    id = attribute.Value;
                    break;
                case "author":
                    author = attribute.Value;
                    break;
                case "date":
                    date = attribute.Value;
                    break;
                default:
                    break;
            }
        }

        return (int.TryParse(id, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0, author, date);
    }

    /// <summary>The next free w:id across all parts (sequential ids per batch).</summary>
    private static int NextRevisionId(WordprocessingDocument doc) =>
        EnumerateRevisions(doc).Select(r => r.Id).DefaultIfEmpty(0).Max() + 1;

    private static string RevisionPath(int id) =>
        string.Create(CultureInfo.InvariantCulture, $"/revision[@id={id}]");

    /// <summary>The wire shape of one revision (read --view revisions / get /revision[...]).</summary>
    private static object RevisionShape(RevisionRecord r) => new
    {
        path = RevisionPath(r.Id),
        id = r.Id,
        kind = r.Kind,
        author = r.Author,
        date = r.Date,
        text = r.Text,
        at = r.At,
        mark = r.ParagraphMark ? "paragraph" : null,
    };

    private static object RevisionsView(WordprocessingDocument doc)
    {
        var revisions = EnumerateRevisions(doc).Select(RevisionShape).ToList();
        return new { view = "revisions", count = revisions.Count, revisions };
    }

    /// <summary>Resolves /revision[@id=N] or positional /revision[i] or throws invalid_path with candidates.</summary>
    private static RevisionRecord ResolveRevision(WordprocessingDocument doc, DocPath path)
    {
        var revisions = EnumerateRevisions(doc);
        var segment = path.Segments[0];

        if (path.Segments.Count > 1)
        {
            throw RevisionNotFound($"'{path.ToCanonicalString()}' has segments after /revision[…].", revisions);
        }

        if (segment.Id is { } idText)
        {
            if (!int.TryParse(idText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id))
            {
                throw RevisionNotFound($"Revision ids are numeric; got '@id={idText}'.", revisions);
            }

            return revisions.FirstOrDefault(r => r.Id == id)
                ?? throw RevisionNotFound($"No revision has id {id}.", revisions);
        }

        var index = segment.Index ?? 1;
        if (revisions.Count == 0)
        {
            throw RevisionNotFound("This document has no tracked revisions.", revisions);
        }

        if (index > revisions.Count)
        {
            throw RevisionNotFound($"/revision[{index}] does not exist; there are {revisions.Count} revision(s).", revisions);
        }

        return revisions[index - 1];
    }

    private static AiofficeException RevisionNotFound(string message, List<RevisionRecord> revisions) => new(
        ErrorCodes.InvalidPath,
        message,
        "Run 'aioffice read <file> --view revisions' to list revisions with their ids.",
        candidates: [.. revisions.Take(5).Select(r => RevisionPath(r.Id))]);

    // -------------------------------------------------------- tracked writes

    /// <summary>Tracked ops apply to body content only.</summary>
    private static void RequireBodyScope(string canonicalPath, string opName)
    {
        if (!canonicalPath.StartsWith("/body", StringComparison.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Tracked {opName} is supported on body content only; got {canonicalPath}.",
                "Edit headers/footers without track:true, or target a /body path.");
        }
    }

    /// <summary>
    /// Tracked set: only text changes can be tracked in M2 (formatting needs
    /// w:rPrChange, planned M3). Old runs are wrapped in w:del with w:delText,
    /// the new text lands in a w:ins — the delete-old + insert-new pattern.
    /// </summary>
    private static object ApplyTrackedSet(WordprocessingDocument doc, ResolvedNode node, JsonObject props, EditSession session)
    {
        RequireBodyScope(node.CanonicalPath, "set");

        var effective = props.DeepClone().AsObject();
        var author = session.ResolveAuthor(effective);

        if (effective.Select(kv => kv.Key).Any(k => k != "text"))
        {
            var others = string.Join(", ", effective.Select(kv => kv.Key).Where(k => k != "text"));
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Producing tracked formatting changes (w:rPrChange) is not supported; got: {others}.",
                "Track text changes only ({\"props\":{\"text\":\"…\"}}), or apply formatting without track:true. " +
                "Existing formatting revisions can be resolved with accept/reject.");
        }

        if (effective["text"] is not { } textNode)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "Tracked set needs props.text.",
                "Example: {\"op\":\"set\",\"path\":\"/body/p[1]\",\"props\":{\"text\":\"New wording\"}} with track:true.");
        }

        var newText = NodeToString(textNode);
        var date = DateTime.UtcNow;
        var nextId = NextRevisionId(doc);

        switch (node.Element)
        {
            case Paragraph paragraph:
                TrackedReplaceRuns(paragraph, [.. paragraph.ChildElements.OfType<Run>()], newText, author, date, ref nextId);
                break;

            case Run run:
                TrackedReplaceRuns((Paragraph)run.Parent!, [run], newText, author, date, ref nextId);
                break;

            case TableCell:
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    "Tracked set on a table cell is not supported; track the paragraph inside it.",
                    $"Target {node.CanonicalPath}/p[1] instead.");

            default:
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"Tracked set is not supported on '{node.Type}'.",
                    "Track text on p or run elements, or run the op without track:true.");
        }

        return new { op = "set", path = node.CanonicalPath, type = node.Type, tracked = true, author };
    }

    /// <summary>Wraps the given old runs in one w:del and inserts one w:ins with the new text.</summary>
    private static void TrackedReplaceRuns(
        Paragraph paragraph, IReadOnlyList<Run> oldRuns, string newText, string author, DateTime date, ref int nextId)
    {
        var keepFormatting = oldRuns.FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
        OpenXmlElement? anchor = oldRuns.FirstOrDefault();

        if (oldRuns.Count > 0)
        {
            var del = NewTrackChange(new DeletedRun(), author, date, nextId++);
            oldRuns[0].InsertBeforeSelf(del);
            anchor = del;
            foreach (var run in oldRuns)
            {
                run.Remove();
                ConvertTextToDeleted(run);
                del.AppendChild(run);
            }
        }

        var insertedRun = new Run();
        if (keepFormatting is not null)
        {
            insertedRun.RunProperties = keepFormatting;
        }

        insertedRun.AppendChild(NewText(newText));
        var ins = NewTrackChange(new InsertedRun(), author, date, nextId++);
        ins.AppendChild(insertedRun);

        if (anchor is not null)
        {
            anchor.InsertAfterSelf(ins);
        }
        else
        {
            paragraph.AppendChild(ins);
        }
    }

    /// <summary>
    /// Tracked remove: runs go into w:del (w:t becomes w:delText) and the
    /// paragraph mark is flagged deleted, so accept removes the paragraph and
    /// reject restores it.
    /// </summary>
    private static object ApplyTrackedRemove(WordprocessingDocument doc, ResolvedNode node, EditOp op, EditSession session)
    {
        RequireBodyScope(node.CanonicalPath, "remove");

        var props = op.Props?.DeepClone().AsObject();
        var author = session.ResolveAuthor(props);
        var date = DateTime.UtcNow;
        var nextId = NextRevisionId(doc);

        switch (node.Element)
        {
            case Paragraph paragraph:
            {
                var runs = paragraph.ChildElements.OfType<Run>().ToList();
                if (runs.Count > 0)
                {
                    var del = NewTrackChange(new DeletedRun(), author, date, nextId++);
                    runs[0].InsertBeforeSelf(del);
                    foreach (var run in runs)
                    {
                        run.Remove();
                        ConvertTextToDeleted(run);
                        del.AppendChild(run);
                    }
                }

                var markRPr = EnsureParagraphMarkRunProperties(paragraph);
                if (!markRPr.Elements<Deleted>().Any())
                {
                    markRPr.PrependChild(NewTrackChange(new Deleted(), author, date, nextId++));
                }

                break;
            }

            case Run run:
            {
                var del = NewTrackChange(new DeletedRun(), author, date, nextId++);
                run.InsertBeforeSelf(del);
                run.Remove();
                ConvertTextToDeleted(run);
                del.AppendChild(run);
                break;
            }

            default:
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"Tracked remove is not supported on '{node.Type}' (tracking covers paragraph content).",
                    "Remove it without track:true, or tracked-remove the paragraphs inside it one by one.");
        }

        return new { op = "remove", path = node.CanonicalPath, type = node.Type, tracked = true, author };
    }

    /// <summary>Wraps a fresh paragraph's runs in one w:ins and flags its paragraph mark inserted.</summary>
    private static void MarkParagraphInserted(WordprocessingDocument doc, Paragraph paragraph, string author)
    {
        var date = DateTime.UtcNow;
        var nextId = NextRevisionId(doc);

        var runs = paragraph.ChildElements.OfType<Run>().ToList();
        if (runs.Count > 0)
        {
            var ins = NewTrackChange(new InsertedRun(), author, date, nextId++);
            runs[0].InsertBeforeSelf(ins);
            foreach (var run in runs)
            {
                run.Remove();
                ins.AppendChild(run);
            }
        }

        var markRPr = EnsureParagraphMarkRunProperties(paragraph);
        markRPr.PrependChild(NewTrackChange(new Inserted(), author, date, nextId));
    }

    private static T NewTrackChange<T>(T element, string author, DateTime date, int id)
        where T : OpenXmlElement
    {
        element.SetAttribute(new OpenXmlAttribute(
            "w", "id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            id.ToString(CultureInfo.InvariantCulture)));
        element.SetAttribute(new OpenXmlAttribute(
            "w", "author", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", author));
        element.SetAttribute(new OpenXmlAttribute(
            "w", "date", "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            date.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture)));
        return element;
    }

    private static ParagraphMarkRunProperties EnsureParagraphMarkRunProperties(Paragraph paragraph)
    {
        var pPr = paragraph.ParagraphProperties ??= new ParagraphProperties();
        return pPr.ParagraphMarkRunProperties ??= new ParagraphMarkRunProperties();
    }

    /// <summary>w:t inside w:del must be w:delText (Word refuses the file otherwise).</summary>
    private static void ConvertTextToDeleted(Run run)
    {
        foreach (var text in run.Elements<Text>().ToList())
        {
            var deleted = new DeletedText(text.Text) { Space = text.Space };
            run.ReplaceChild(deleted, text);
        }
    }

    private static void ConvertDeletedToText(Run run)
    {
        foreach (var deleted in run.Elements<DeletedText>().ToList())
        {
            var text = new Text(deleted.Text) { Space = deleted.Space };
            run.ReplaceChild(text, deleted);
        }
    }

    // --------------------------------------------------------- accept/reject

    /// <summary>
    /// <c>{"op":"accept"|"reject","path":P}</c> where P is /revision[@id=N],
    /// /revision[i], or a scope path (/body, /body/p[3]) meaning every
    /// revision inside it — insert/delete wrappers, paragraph marks and
    /// formatting changes (w:rPrChange / w:pPrChange) alike.
    /// </summary>
    private static object ApplyAcceptOrReject(WordprocessingDocument doc, EditOp op)
    {
        var accept = op.Op == "accept";
        var path = DocPath.Parse(op.Path);

        List<RevisionRecord> targets;
        if (path.Segments[0].Name == "revision")
        {
            targets = [ResolveRevision(doc, path)];
        }
        else
        {
            var scope = WordAddress.Resolve(doc, path);
            targets = [.. EnumerateRevisions(doc).Where(r => IsSelfOrDescendantOf(r.Element, scope.Element))];
        }

        // Reverse document order: content revisions resolve before the paragraph
        // mark of the same paragraph, so mark handling sees the settled content.
        for (var i = targets.Count - 1; i >= 0; i--)
        {
            ResolveOne(targets[i], accept);
        }

        return new
        {
            op = accept ? "accept" : "reject",
            path = op.Path,
            applied = targets.Count,
        };
    }

    private static bool IsSelfOrDescendantOf(OpenXmlElement element, OpenXmlElement scope)
    {
        for (OpenXmlElement? node = element; node is not null; node = node.Parent)
        {
            if (ReferenceEquals(node, scope))
            {
                return true;
            }
        }

        return false;
    }

    private static void ResolveOne(RevisionRecord revision, bool accept)
    {
        if (revision.ParagraphMark)
        {
            ResolveParagraphMark(revision, accept);
            return;
        }

        if (revision.Kind == "format")
        {
            ResolveFormatChange(revision.Element, accept);
            return;
        }

        switch (revision.Element, accept)
        {
            case (InsertedRun ins, true): // keep text, drop markup
                UnwrapChildren(ins);
                break;

            case (InsertedRun ins, false): // never happened
                ins.Remove();
                break;

            case (DeletedRun del, true): // the delete becomes real
                del.Remove();
                break;

            case (DeletedRun del, false): // restore the original text
                foreach (var run in del.ChildElements.OfType<Run>())
                {
                    ConvertDeletedToText(run);
                }

                UnwrapChildren(del);
                break;

            default:
                break;
        }
    }

    /// <summary>
    /// Formatting revisions: accept keeps the current (new) formatting and
    /// drops the marker; reject restores the previous properties the marker
    /// preserved (w:rPrChange/w:pPrChange children are the OLD props).
    /// </summary>
    private static void ResolveFormatChange(OpenXmlElement marker, bool accept)
    {
        if (accept)
        {
            var host = marker.Parent;
            marker.Remove();
            TidyFormatHost(host);
            return;
        }

        switch (marker)
        {
            case RunPropertiesChange rPrChange when rPrChange.Parent is RunProperties rPr && rPr.Parent is Run run:
            {
                var previous = rPrChange.PreviousRunProperties;
                if (previous is null || !previous.HasChildren)
                {
                    run.RunProperties = null;
                }
                else
                {
                    var restored = new RunProperties();
                    foreach (var child in previous.ChildElements)
                    {
                        restored.AppendChild(child.CloneNode(true));
                    }

                    run.RunProperties = restored;
                }

                break;
            }

            case RunPropertiesChange rPrChange when rPrChange.Parent is ParagraphMarkRunProperties markRPr:
            {
                var previous = rPrChange.PreviousRunProperties;
                foreach (var child in markRPr.ChildElements.Where(c => c is not (Inserted or Deleted)).ToList())
                {
                    child.Remove();
                }

                foreach (var child in previous?.ChildElements ?? [])
                {
                    markRPr.AppendChild(child.CloneNode(true));
                }

                TidyFormatHost(markRPr);
                break;
            }

            case ParagraphPropertiesChange pPrChange when pPrChange.Parent is ParagraphProperties pPr:
            {
                var previous = pPrChange.ParagraphPropertiesExtended;

                // Restore the old pPr base; the paragraph-mark rPr and any sectPr are not part of the change.
                foreach (var child in pPr.ChildElements
                    .Where(c => c is not (ParagraphMarkRunProperties or SectionProperties or ParagraphPropertiesChange))
                    .ToList())
                {
                    child.Remove();
                }

                // Base props precede the paragraph-mark rPr / sectPr in CT_PPr; restore them
                // in their original order ahead of whatever survived the sweep.
                var anchor = pPr.ChildElements.FirstOrDefault(c =>
                    c is ParagraphMarkRunProperties or SectionProperties or ParagraphPropertiesChange);
                foreach (var child in previous?.ChildElements ?? [])
                {
                    if (anchor is null)
                    {
                        pPr.AppendChild(child.CloneNode(true));
                    }
                    else
                    {
                        pPr.InsertBefore(child.CloneNode(true), anchor);
                    }
                }

                pPrChange.Remove();
                TidyFormatHost(pPr);
                break;
            }

            default:
                marker.Remove();
                break;
        }
    }

    /// <summary>Empty rPr/pPr wrappers left after a resolved format change are noise; drop them.</summary>
    private static void TidyFormatHost(OpenXmlElement? host)
    {
        if (host is RunProperties or ParagraphMarkRunProperties && !host.HasChildren)
        {
            var pPr = host.Parent as ParagraphProperties;
            host.Remove();
            TidyFormatHost(pPr);
        }
        else if (host is ParagraphProperties pPrHost && !pPrHost.HasChildren)
        {
            pPrHost.Remove();
        }
    }

    /// <summary>
    /// Paragraph-mark revisions: accepting an inserted mark (or rejecting a
    /// deleted one) just drops the marker; rejecting an inserted mark (or
    /// accepting a deleted one) removes the paragraph once no untracked or
    /// pending content remains — the end state Word users expect for whole
    /// paragraphs added/removed under track.
    /// </summary>
    private static void ResolveParagraphMark(RevisionRecord revision, bool accept)
    {
        var marker = revision.Element;
        var paragraph = marker.Ancestors<Paragraph>().FirstOrDefault();
        marker.Remove();
        if (paragraph is null)
        {
            return;
        }

        TidyEmptyParagraphProperties(paragraph);

        var removeParagraph = (revision.Kind, accept) switch
        {
            ("insert", false) => true, // the inserted paragraph never happened
            ("delete", true) => true, // the deletion becomes real
            _ => false,
        };

        if (removeParagraph && !HasRemainingContent(paragraph))
        {
            paragraph.Remove();
        }
    }

    private static bool HasRemainingContent(Paragraph paragraph) =>
        paragraph.ChildElements.Any(c => c is Run or InsertedRun or DeletedRun or Hyperlink)
        || paragraph.InnerText.Length > 0;

    private static void TidyEmptyParagraphProperties(Paragraph paragraph)
    {
        var pPr = paragraph.ParagraphProperties;
        if (pPr?.ParagraphMarkRunProperties is { } rPr && !rPr.HasChildren)
        {
            rPr.Remove();
        }

        if (pPr is not null && !pPr.HasChildren)
        {
            pPr.Remove();
        }
    }

    private static void UnwrapChildren(OpenXmlElement wrapper)
    {
        foreach (var child in wrapper.ChildElements.ToList())
        {
            child.Remove();
            wrapper.InsertBeforeSelf(child);
        }

        wrapper.Remove();
    }
}
