using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// Batch-level options threaded through every op: the sandbox (file-valued
/// props MUST resolve through it), the tracked-changes flag and the batch
/// author (op props.author &gt; batch author &gt; "AIOffice").
/// </summary>
internal sealed record EditSession(Workspace Workspace, bool Track, string? Author)
{
    public const string DefaultAuthor = "AIOffice";

    /// <summary>Non-fatal warnings raised by ops (find_no_match, toc_pages_unknown, …); surfaced on envelope meta.</summary>
    public List<Warning> Warnings { get; } = [];

    /// <summary>Resolves the revision/comment author and consumes props.author so generic prop application never sees it.</summary>
    public string ResolveAuthor(JsonObject? props)
    {
        if (props is not null && props.TryGetPropertyValue("author", out var node))
        {
            props.Remove("author");
            var author = WordHandler.NodeToString(node);
            if (author.Length > 0)
            {
                return author;
            }
        }

        return Author is { Length: > 0 } ? Author : DefaultAuthor;
    }
}

public sealed partial class WordHandler
{
    public Envelope Edit(CommandContext ctx, IReadOnlyList<EditOp> ops)
    {
        var file = RequireFile(ctx, mustExist: true);
        var dryRun = BoolArg(ctx.Args, "dryRun") || BoolArg(ctx.Args, "dry-run");
        var expectRev = StringArg(ctx.Args, "expectRev") ?? StringArg(ctx.Args, "expect-rev");
        var session = new EditSession(ctx.Workspace, BoolArg(ctx.Args, "track"), StringArg(ctx.Args, "author"));

        var originalBytes = File.ReadAllBytes(file);
        var currentRev = Rev.OfBytes(originalBytes);
        if (expectRev is not null && !expectRev.Equals(currentRev, StringComparison.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.StaleAddress,
                $"The file changed since you last read it: rev is {currentRev}, you expected {expectRev}.",
                "Re-run 'aioffice read' or 'aioffice query' to get fresh paths and the current rev, then retry.");
        }

        // When the batch sets document properties, upgrade any legacy .psmdcp
        // core-properties part to the standard docProps/core.xml first, so the write
        // lands in the conventional part instead of stranding values in the old one.
        // Scoped to property writes so unrelated edits never rewrite the core part.
        var workingBytes = ops.Any(o => o.Op == "set" && IsPropertiesPath(o.Path))
            ? MigrateLegacyCorePropertiesBytes(originalBytes, file)
            : originalBytes;

        // Atomic: every op is applied to an in-memory copy; the file is written only when all succeed.
        var ms = new MemoryStream();
        ms.Write(workingBytes);
        ms.Position = 0;

        var summaries = new List<object>(ops.Count);
        using (var doc = OpenPackage(ms, file, editable: true))
        {
            for (var i = 0; i < ops.Count; i++)
            {
                try
                {
                    summaries.Add(ApplyOp(doc, file, ops[i], session));
                }
                catch (AiofficeException ex)
                {
                    throw new AiofficeException(
                        ex.Code,
                        $"ops[{i}] ({ops[i].Op} {ops[i].Path}): {ex.Message}",
                        ex.Suggestion,
                        ex.Candidates,
                        ex);
                }
            }
        }

        var newBytes = ms.ToArray();

        if (dryRun)
        {
            return Envelope.Ok(
                new { applied = summaries.Count, dryRun = true, ops = summaries },
                MetaFor(file, currentRev, session.Warnings));
        }

        var snapshot = _snapshots.Save(file); // pre-image, so the edit is undoable
        File.WriteAllBytes(file, newBytes);

        return Envelope.Ok(
            new { applied = summaries.Count, snapshot = snapshot.Number, ops = summaries },
            MetaFor(file, Rev.OfBytes(newBytes), session.Warnings));
    }

    // ------------------------------------------------------------------- ops

    private static object ApplyOp(WordprocessingDocument doc, string file, EditOp op, EditSession session)
    {
        var parsedPath = DocPath.Parse(op.Path);
        if (parsedPath.IsRoot)
        {
            // The document root carries document-level protection since 1.13: set /
            // writes w:documentProtection / w:writeProtection into settings.xml,
            // mirroring how xlsx 'set /' carries workbook protection/calc props.
            // Any other root-targeted op (or a propless set /) keeps the old
            // helpful refusal that points at the real edit targets.
            if (op.Op == "set" && op.Props is { Count: > 0 })
            {
                return ApplySetDocumentProtection(doc, op);
            }

            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "docx has no document-root edit target ('/').",
                "Set document protection with {op:set, path:/, props:{protection:{edit:\"readOnly\"}}}; "
                + "edit page setup on /section[1], body content under /body, or use 'replace' with path '/' for document-wide find/replace.");
        }

        var rootName = parsedPath.Segments[0].Name;
        return op.Op switch
        {
            // extract is a producing-but-not-mutating op: it writes an embed's
            // payload to a sandbox-resolved dest and leaves the document unchanged.
            "extract" => ApplyExtractEmbed(doc, op, session),
            "accept" or "reject" => ApplyAcceptOrReject(doc, op),
            "replace" => ApplyReplace(doc, op, session),
            "set" when rootName == "style" => ApplySetStyle(doc, op),
            "set" when rootName == "section" => ApplySetSection(doc, op),
            "set" when rootName == "properties" => ApplySetProperties(doc, op),
            "set" when rootName == "sdt" => ApplySetContentControl(doc, op),
            "set" when rootName == "theme" => ApplySetTheme(doc, op),
            "set" when rootName == "formField" => ApplySetFormField(doc, op),
            // /body/shape[i] and /body/textBox[i] are two-segment body paths.
            "set" when IsBodyShapePath(parsedPath, "textBox") => ApplySetBodyShape(doc, op, isTextBox: true),
            "set" when IsBodyShapePath(parsedPath, "shape") => ApplySetBodyShape(doc, op, isTextBox: false),
            "set" => ApplySet(doc, op, session),
            // Part-backed adds cannot resolve their anchor through WordAddress
            // (the target part/list, not body content, receives the element).
            "add" when op.Type is "header" or "footer" => ApplyAddHeaderFooter(doc, file, op),
            "add" when op.Type == "style" => ApplyAddStyle(doc, op),
            "add" when op.Type == "comment" => ApplyAddComment(doc, op, session),
            "add" when op.Type == "reply" => ApplyAddCommentReply(doc, op, session),
            "add" when op.Type == "image" => ApplyAddImage(doc, op, session),
            "add" when op.Type == "embed" => ApplyAddEmbed(doc, op, session),
            "add" when op.Type == "link" && session.Track => throw TrackedStructureUnsupported("link"),
            "add" when op.Type == "link" => ApplyAddLink(doc, op, session),
            "add" when op.Type == "field" && session.Track => throw TrackedStructureUnsupported("field"),
            "add" when op.Type == "field" => ApplyAddField(doc, file, op),
            "add" when op.Type == "bookmark" => ApplyAddBookmark(doc, op),
            "add" when op.Type == "footnote" && session.Track => throw TrackedStructureUnsupported("footnote"),
            "add" when op.Type == "footnote" => ApplyAddFootnote(doc, op),
            // Endnotes are real since M4 (the M3 unsupported_feature refusal is gone).
            "add" when op.Type == "endnote" && session.Track => throw TrackedStructureUnsupported("endnote"),
            "add" when op.Type == "endnote" => ApplyAddEndnote(doc, op),
            "add" when op.Type == "toc" && session.Track => throw TrackedStructureUnsupported("toc"),
            "add" when op.Type == "toc" => ApplyAddToc(doc, file, op, session),
            "add" when op.Type == "tableOfFigures" && session.Track => throw TrackedStructureUnsupported("tableOfFigures"),
            "add" when op.Type == "tableOfFigures" => ApplyAddTableOfFigures(doc, op, session),
            "add" when op.Type == "indexEntry" && session.Track => throw TrackedStructureUnsupported("indexEntry"),
            "add" when op.Type == "indexEntry" => ApplyAddIndexEntry(doc, op),
            "add" when op.Type == "index" && session.Track => throw TrackedStructureUnsupported("index"),
            "add" when op.Type == "index" => ApplyAddIndex(doc, op, session),
            "add" when op.Type == "mergeField" && session.Track => throw TrackedStructureUnsupported("mergeField"),
            "add" when op.Type == "mergeField" => ApplyAddMergeField(doc, op),
            "add" when op.Type == "ifField" && session.Track => throw TrackedStructureUnsupported("ifField"),
            "add" when op.Type == "ifField" => ApplyAddIfField(doc, op),
            "add" when op.Type == "watermark" && session.Track => throw TrackedStructureUnsupported("watermark"),
            "add" when op.Type == "watermark" => ApplyAddWatermark(doc, file, op, session),
            "add" when op.Type == "sectionBreak" && session.Track => throw TrackedStructureUnsupported("sectionBreak"),
            "add" when op.Type == "sectionBreak" => ApplyAddSectionBreak(doc, op),
            "add" when op.Type == "equation" && session.Track => throw TrackedStructureUnsupported("equation"),
            "add" when op.Type == "equation" => ApplyAddEquation(doc, op, session),
            "add" when op.Type == "columnBreak" => ApplyAddColumnBreak(doc, op, session),
            "add" when op.Type == "caption" && session.Track => throw TrackedStructureUnsupported("caption"),
            "add" when op.Type == "caption" => ApplyAddCaption(doc, op, session),
            "add" when op.Type == "crossRef" && session.Track => throw TrackedStructureUnsupported("crossRef"),
            "add" when op.Type == "crossRef" => ApplyAddCrossRef(doc, op, session),
            "add" when op.Type == "contentControl" && session.Track => throw TrackedStructureUnsupported("contentControl"),
            "add" when op.Type == "contentControl" => ApplyAddContentControl(doc, op),
            "add" when op.Type == "shape" && session.Track => throw TrackedStructureUnsupported("shape"),
            "add" when op.Type == "shape" => ApplyAddBodyShape(doc, op, session, isTextBox: false),
            "add" when op.Type == "textBox" && session.Track => throw TrackedStructureUnsupported("textBox"),
            "add" when op.Type == "textBox" => ApplyAddBodyShape(doc, op, session, isTextBox: true),
            "add" when op.Type == "formField" && session.Track => throw TrackedStructureUnsupported("formField"),
            "add" when op.Type == "formField" => ApplyAddFormField(doc, op),
            "add" when op.Type == "source" => ApplyAddSource(doc, op),
            "add" when op.Type == "citation" && session.Track => throw TrackedStructureUnsupported("citation"),
            "add" when op.Type == "citation" => ApplyAddCitation(doc, op),
            "add" when op.Type == "bibliography" && session.Track => throw TrackedStructureUnsupported("bibliography"),
            "add" when op.Type == "bibliography" => ApplyAddBibliography(doc, file, op, session),
            "add" when op.Type == "buildingBlock" && session.Track => throw TrackedStructureUnsupported("buildingBlock"),
            "add" when op.Type == "buildingBlock" => ApplyAddBuildingBlock(doc, op),
            "add" when op.Type == "buildingBlockRef" && session.Track => throw TrackedStructureUnsupported("buildingBlockRef"),
            "add" when op.Type == "buildingBlockRef" => ApplyAddBuildingBlockRef(doc, op),
            "add" => ApplyAdd(doc, op, session),
            "remove" when rootName == "source" => ApplyRemoveSource(doc, op),
            "remove" when rootName == "buildingBlock" => ApplyRemoveBuildingBlock(doc, op),
            "remove" when rootName == "bibliography" => ApplyRemoveBibliography(doc, op),
            "remove" when rootName == "embed" => ApplyRemoveEmbed(doc, op),
            "remove" when rootName == "style" => ApplyRemoveStyle(doc, op),
            "remove" when rootName == "comment" => ApplyRemoveComment(doc, op),
            "remove" when rootName == "bookmark" => ApplyRemoveBookmark(doc, op),
            "remove" when rootName == "footnote" => ApplyRemoveFootnote(doc, op),
            "remove" when rootName == "endnote" => ApplyRemoveEndnote(doc, op),
            "remove" when rootName == "toc" => ApplyRemoveToc(doc, op),
            "remove" when rootName == "tableOfFigures" => ApplyRemoveTableOfFigures(doc, op),
            "remove" when rootName == "index" => ApplyRemoveIndex(doc, op),
            "remove" when rootName == "watermark" => ApplyRemoveWatermark(doc, op),
            "remove" when rootName == "section" => ApplyRemoveSection(doc, op),
            "remove" when rootName == "sdt" => ApplyRemoveContentControl(doc, op),
            "remove" when rootName == "formField" => ApplyRemoveFormField(doc, op),
            "remove" when IsBodyShapePath(parsedPath, "textBox") => ApplyRemoveBodyShape(doc, op, isTextBox: true),
            "remove" when IsBodyShapePath(parsedPath, "shape") => ApplyRemoveBodyShape(doc, op, isTextBox: false),
            "remove" when DocPath.Parse(op.Path).Segments[^1].Name == "omath" => ApplyRemoveEquation(doc, op),
            "remove" when rootName == "revision" => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "Revisions are not removed; they are accepted or rejected.",
                "Use {\"op\":\"accept\",\"path\":\"" + op.Path + "\"} to apply it or {\"op\":\"reject\",...} to undo it."),
            "remove" => ApplyRemove(doc, op, session),
            _ => ApplyMove(doc, op),
        };
    }

    private static object ApplySet(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var node = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        var props = RequireProps(op);

        if (session.Track)
        {
            return ApplyTrackedSet(doc, node, props, session);
        }

        // Tables and cells take structured props (arrays, merge counts), so they
        // bypass the stringly-typed paragraph/run prop loop.
        if (node.Element is Table table)
        {
            return ApplySetTable(table, node, props);
        }

        if (node.Element is TableCell cell)
        {
            return ApplySetCell(cell, node, props, session);
        }

        // Image-carrying paragraphs/runs take an alt/descr property: it sets the
        // wp:docPr description (the accessibility alt text the auditor checks).
        if (node.Element.Descendants<Drawing>().Any() &&
            (props.ContainsKey("alt") || props.ContainsKey("descr")))
        {
            return ApplySetImageAlt(node, props);
        }

        // Text effects (shadow/glow/reflection/outline) are structured props that
        // emit w14 run effects. Peel them out of props first so the stringly-typed
        // loop below never sees them, but apply them AFTER the loop runs — a text
        // prop in the same op rebuilds the runs, and effects must land on the final
        // run text. They can be the whole op, or ride alongside plain formatting.
        var effectProps = ExtractTextEffectProps(props);

        // (1.7) Drop-cap props restructure the paragraph (the first letter moves to
        // a framed paragraph). Peel them out so the formatting loop never sees them;
        // apply after the loop so a same-op text change has already rebuilt the runs.
        var dropCapProps = ExtractDropCapProps(props);
        if (dropCapProps is not null && node.Element is not Paragraph)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"dropCap applies to a paragraph, not '{node.Type}'.",
                "Target a paragraph path, e.g. {\"op\":\"set\",\"path\":\"/body/p[1]\",\"props\":{\"dropCap\":\"drop\"}}.");
        }

        // (1.8) Structured paragraph-visual props (shading fill, border box) take an
        // object value, so peel them out of the stringly-typed loop; the scalar
        // companions (spacingBefore/After, indentLeft/Right) flow through as usual.
        var paragraphVisualProps = ExtractParagraphVisualProps(props);
        if (paragraphVisualProps is not null && node.Element is not Paragraph)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"shading/border apply to a paragraph, not '{node.Type}'.",
                "Target a paragraph path, e.g. {\"op\":\"set\",\"path\":\"/body/p[1]\",\"props\":{\"shading\":\"FEF9C3\"}}. " +
                "For a table cell use the cell's own shading prop.");
        }

        foreach (var (name, value) in OrderedProps(props))
        {
            switch (node.Element)
            {
                case Paragraph p:
                    if (name == "style")
                    {
                        EnsureStyleDefined(doc, value);
                    }

                    WordFormatting.SetParagraphProp(p, name, value);
                    break;

                case Run r:
                    WordFormatting.SetRunProp(r, name, value);
                    break;

                default:
                    throw new AiofficeException(
                        ErrorCodes.UnsupportedFeature,
                        $"set is not supported on '{node.Type}' (property '{name}').",
                        node.Type switch
                        {
                            "header" or "footer" => "Address a paragraph inside it instead: " + node.CanonicalPath + "/p[1].",
                            "tr" => "Set table-level props on the table (headerRow styles row 1), or cell props on a tc.",
                            _ => "Set properties on p or run elements; address one with query first.",
                        });
            }
        }

        var effects = effectProps is null ? [] : ApplyTextEffects(doc, node, effectProps);

        var visuals = paragraphVisualProps is not null && node.Element is Paragraph visualParagraph
            ? ApplyParagraphVisuals(visualParagraph, paragraphVisualProps)
            : [];

        if (dropCapProps is not null && node.Element is Paragraph dropCapParagraph)
        {
            var dropCap = ApplyDropCap(dropCapParagraph, dropCapProps);
            return new { op = "set", path = node.CanonicalPath, type = node.Type, dropCap };
        }

        if (visuals.Count > 0)
        {
            return new { op = "set", path = node.CanonicalPath, type = node.Type, visuals };
        }

        return effects.Count > 0
            ? (object)new { op = "set", path = node.CanonicalPath, type = node.Type, effects }
            : new { op = "set", path = node.CanonicalPath, type = node.Type };
    }

    private static object ApplyAdd(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        var type = op.Type ?? "p";
        var position = op.Position;
        if (position is not (null or "before" or "after" or "inside"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add position '{position}' is not valid.",
                "Use position before, after or inside (inside only for /body, header/footer, table or tc targets).",
                candidates: ["before", "after", "inside"]);
        }

        // Tracked insertion is paragraph-scoped; author resolution must run
        // before BuildParagraph so props.author never reaches generic prop handling.
        var props = op.Props?.DeepClone().AsObject();
        var author = session.ResolveAuthor(props);
        var listRequest = type == "p" && props is not null ? PopListProps(props) : null;

        OpenXmlElement created = type switch
        {
            "p" => BuildParagraph(doc, props),
            "tr" when session.Track => throw TrackedStructureUnsupported("tr"),
            "table" when session.Track => throw TrackedStructureUnsupported("table"),
            "tr" => BuildRow(props),
            "table" => WordFactory.Table(
                rows: PropInt(props, "rows") ?? 2,
                columns: PropInt(props, "columns") ?? PropInt(props, "cols") ?? 2),
            _ => throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"add --type {type} is not supported for docx.",
                "Add p (props.style=Heading1 for headings, props.list=bullet|number for lists), tr (props.cells=[…]), " +
                "table (props.rows/columns), image (props.src), link (props.url), bookmark (props.name), " +
                "footnote/endnote (props.text), comment/reply, style, toc (props.levels), watermark (props.text), " +
                "sectionBreak (props.kind), field (props.kind=pageNumber|numPages|date|docTitle), " +
                "equation (props.latex, props.display), columnBreak, " +
                "caption (props.label=Figure|Table|Equation, props.text), crossRef (props.to, props.show), " +
                "tableOfFigures (props.label — lists captions of one label), " +
                "indexEntry (props.text — marks an XE field), index (props.columns — builds the index), " +
                "mergeField (props.name — a MERGEFIELD the template verb fills by name), " +
                "ifField (props.field/operator/value/trueText/falseText — an «IF» field resolved during merge), " +
                "embed (props.src — embeds any file as an OLE package object), " +
                "source (props.tag/kind/title — adds a bibliography source at /sources), " +
                "citation (props.source — cites a source by tag), bibliography (props.style), " +
                "shape (props.shape=rect|roundRect|ellipse|line|arrow — a floating body drawing), " +
                "textBox (props.text — a floating text box), " +
                "formField (props.kind=text|checkbox|dropdown, props.name — a legacy form field), " +
                "or header/footer targeting /header[1]|/header[firstPage]|/header[even]. " +
                "For runs, set text on the paragraph instead.",
                candidates: ["p", "tr", "table", "image", "link", "bookmark", "footnote", "endnote", "comment", "reply",
                    "style", "header", "footer", "toc", "watermark", "sectionBreak", "field", "equation", "columnBreak",
                    "caption", "crossRef", "tableOfFigures", "indexEntry", "index", "mergeField", "ifField",
                    "embed", "source", "citation", "bibliography", "shape", "textBox", "formField"]),
        };

        // Default placement: containers receive children, blocks get siblings after them.
        var isContainerAnchor = anchor.Type is "body" or "tc" or "header" or "footer"
            || (anchor.Type == "table" && created is TableRow);
        var pos = position ?? (isContainerAnchor ? "inside" : "after");

        var valid = (created, anchor.Type, pos) switch
        {
            (Paragraph or Table, "body" or "tc" or "header" or "footer", "inside") => true,
            (Paragraph or Table, "p" or "table", "before" or "after") => true,
            (TableRow, "table", "inside") => true,
            (TableRow, "tr", "before" or "after") => true,
            _ => false,
        };

        if (!valid)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Cannot add a '{type}' {pos} {anchor.CanonicalPath} ({anchor.Type}).",
                "Add p/table inside /body, /header[n], /footer[n] or a tc, or before/after an existing p/table; " +
                "add tr inside a table, or before/after an existing tr.");
        }

        if (session.Track && created is Paragraph trackedParagraph)
        {
            RequireBodyScope(anchor.CanonicalPath, "add");
            MarkParagraphInserted(doc, trackedParagraph, author);
        }

        var canonical = Insert(doc, anchor, created, pos);

        // List numbering needs the final position: neighbors decide whether the
        // new item continues their sequence or starts a fresh one.
        if (listRequest is not null && created is Paragraph listParagraph)
        {
            ApplyListNumbering(doc, listParagraph, listRequest);
        }

        return new { op = "add", type, anchor = anchor.CanonicalPath, path = canonical };
    }

    private static AiofficeException TrackedStructureUnsupported(string type) => new(
        ErrorCodes.UnsupportedFeature,
        $"Tracked '{type}' additions are not supported (tracking covers paragraph text content).",
        "Run the op without track:true, or add tracked paragraphs instead.");

    /// <summary>Inserts at the validated position; appending to body keeps sectPr last.</summary>
    private static string Insert(WordprocessingDocument doc, ResolvedNode anchor, OpenXmlElement created, string pos)
    {
        switch (pos)
        {
            case "before":
                anchor.Element.InsertBeforeSelf(created);
                break;

            case "after":
                anchor.Element.InsertAfterSelf(created);
                break;

            default: // "inside"
                if (anchor.Element is Body b && b.Elements<SectionProperties>().FirstOrDefault() is { } sectPr)
                {
                    sectPr.InsertBeforeSelf(created);
                }
                else
                {
                    anchor.Element.AppendChild(created);
                }

                break;
        }

        // Report where the new node landed as a canonical path (body, header or footer scope).
        var match = WordAddress.EnumerateAll(doc).FirstOrDefault(n => ReferenceEquals(n.Element, created));
        return match?.CanonicalPath ?? anchor.CanonicalPath;
    }

    private static object ApplyRemove(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var node = WordAddress.Resolve(doc, DocPath.Parse(op.Path));

        if (session.Track)
        {
            return ApplyTrackedRemove(doc, node, op, session);
        }

        if (node.Element is Hyperlink hyperlink)
        {
            RemoveLink(doc, hyperlink);
            return new { op = "remove", path = node.CanonicalPath, type = "link" };
        }

        if (node.Element is not (Paragraph or Table or TableRow))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"remove is not supported on '{node.Type}'.",
                node.Type is "header" or "footer"
                    ? $"Removing a whole {node.Type} is not supported yet. Blank it instead: set {node.CanonicalPath}/p[1] text to \"\"."
                    : "Remove p, table, tr or link elements. To clear a run or cell, set its text to \"\" instead.");
        }

        if (node.Element is Paragraph && node.Element.Parent is TableCell cell &&
            cell.Elements<Paragraph>().Count() == 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"{node.CanonicalPath} is the only paragraph in its table cell; a cell must keep one.",
                "Set its text to \"\" instead: {\"op\":\"set\",\"path\":\"" + node.CanonicalPath + "\",\"props\":{\"text\":\"\"}}.");
        }

        // A w:hdr/w:ftr must keep at least one block-level element to stay schema-valid.
        if (node.Element is Paragraph or Table && node.Element.Parent is Header or Footer &&
            node.Element.Parent!.ChildElements.Count(c => c is Paragraph or Table) == 1)
        {
            var container = node.Element.Parent is Header ? "header" : "footer";
            var containerPath = node.CanonicalPath[..node.CanonicalPath.LastIndexOf('/')];
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"{node.CanonicalPath} is the last block in its {container}; a {container} must keep at least one block.",
                node.Element is Paragraph
                    ? "Set its text to \"\" instead: {\"op\":\"set\",\"path\":\"" + node.CanonicalPath + "\",\"props\":{\"text\":\"\"}}."
                    : "Add a paragraph to the " + container + " first ({\"op\":\"add\",\"path\":\"" + containerPath + "\",\"type\":\"p\"}), then remove the table.");
        }

        node.Element.Remove();
        return new { op = "remove", path = node.CanonicalPath, type = node.Type };
    }

    private static object ApplyMove(WordprocessingDocument doc, EditOp op)
    {
        var source = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (source.Element is not (Paragraph or Table or TableRow))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"move is not supported on '{source.Type}'.",
                "Move p, table or tr elements.");
        }

        var (before, targetPath) = ParseMovePosition(op.Position);
        var target = WordAddress.Resolve(doc, DocPath.Parse(targetPath));
        if (ReferenceEquals(source.Element, target.Element))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "move source and target are the same element.",
                "Pick a different target path in position.");
        }

        source.Element.Remove();
        if (before)
        {
            target.Element.InsertBeforeSelf(source.Element);
        }
        else
        {
            target.Element.InsertAfterSelf(source.Element);
        }

        return new { op = "move", path = source.CanonicalPath, target = target.CanonicalPath, placement = before ? "before" : "after" };
    }

    private static (bool Before, string TargetPath) ParseMovePosition(string? position)
    {
        var parts = position?.Split([' ', ':'], 2, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts is [var direction, var target] && direction is "before" or "after" && target.StartsWith('/'))
        {
            return (direction == "before", target);
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"move needs position '<before|after> <path>', got '{position}'.",
            "Example: {\"op\":\"move\",\"path\":\"/body/p[3]\",\"position\":\"before /body/p[1]\"}.");
    }

    // --------------------------------------------------------------- helpers

    private static Paragraph BuildParagraph(WordprocessingDocument doc, JsonObject? props)
    {
        var paragraph = new Paragraph();
        var working = props ?? [];
        var effectProps = ExtractTextEffectProps(working);
        foreach (var (name, value) in OrderedProps(working))
        {
            if (name == "style")
            {
                EnsureStyleDefined(doc, value);
            }

            WordFormatting.SetParagraphProp(paragraph, name, value);
        }

        if (effectProps is not null)
        {
            ApplyTextEffects(doc, new ResolvedNode(paragraph, "/body/p", "p"), effectProps);
        }

        return paragraph;
    }

    private static TableRow BuildRow(JsonObject? props)
    {
        var cells = props?["cells"] as JsonArray;
        if (cells is null || cells.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type tr needs props.cells with one text per cell.",
                "Example: {\"op\":\"add\",\"path\":\"/body/table[1]\",\"type\":\"tr\",\"props\":{\"cells\":[\"a\",\"b\"]}}.");
        }

        return WordFactory.Row([.. cells.Select(c => NodeToString(c))]);
    }

    private static JsonObject RequireProps(EditOp op) =>
        op.Props is { Count: > 0 }
            ? op.Props
            : throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "set needs at least one property in props.",
                "Example: {\"op\":\"set\",\"path\":\"/body/p[1]\",\"props\":{\"text\":\"Hello\",\"bold\":true}}.");

    /// <summary>Props in application order: text first (creates runs), then style, then formatting.</summary>
    private static IEnumerable<(string Name, string Value)> OrderedProps(JsonObject props) =>
        props
            .OrderBy(kv => kv.Key switch { "text" => 0, "style" => 1, _ => 2 })
            .Select(kv => (kv.Key, NodeToString(kv.Value)));

    private static int? PropInt(JsonObject? props, string name) =>
        props?[name] is { } node && int.TryParse(NodeToString(node), out var n) ? n : null;

    internal static string NodeToString(JsonNode? node)
    {
        if (node is null)
        {
            return string.Empty;
        }

        if (node is JsonValue v)
        {
            if (v.TryGetValue<string>(out var s))
            {
                return s;
            }

            if (v.TryGetValue<bool>(out var b))
            {
                return b ? "true" : "false";
            }

            if (v.TryGetValue<double>(out var d))
            {
                return d.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        return node.ToJsonString();
    }

    /// <summary>Adds a built-in Heading1..6 definition when an edit references one that is missing.</summary>
    private static void EnsureStyleDefined(WordprocessingDocument doc, string styleId)
    {
        if (HeadingLevel(styleId) is not { } level || level > 6 || doc.MainDocumentPart is not { } main)
        {
            return; // custom styles pass through untouched; Word renders them unstyled if undefined
        }

        var stylesPart = main.StyleDefinitionsPart;
        if (stylesPart is null)
        {
            WordFactory.AddDefaultStylesPart(main);
            stylesPart = main.StyleDefinitionsPart!;
        }

        var styles = stylesPart.Styles ??= new Styles();
        var canonicalId = "Heading" + level;
        if (!styles.Elements<Style>().Any(s => s.StyleId?.Value == canonicalId))
        {
            var halfPoints = level switch { 1 => "32", 2 => "28", 3 => "26", _ => "24" };
            styles.AppendChild(WordFactory.HeadingStyle(level, halfPoints));
        }
    }
}
