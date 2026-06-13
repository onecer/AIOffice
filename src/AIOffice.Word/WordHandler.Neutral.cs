using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Dw = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// The M9 conversion surface for docx: project the document to the
/// format-neutral model (<see cref="ExportNeutral"/>) and absorb a neutral model
/// into a freshly created docx (<see cref="ImportNeutral"/>). The neutral model
/// is the lingua franca <c>convert</c> uses to move content between formats; it
/// carries headings, paragraphs, formatted runs, lists, tables and images, and
/// is deliberately lossy about everything format-specific. Anything the model
/// cannot carry (or that a freshly created docx cannot embed) is named honestly
/// in <see cref="ImportResult.Dropped"/> / on the export warnings so the
/// <c>convert</c> verb can surface it as a <c>convert_lossy</c> note.
/// </summary>
public sealed partial class WordHandler : INeutralConvertible
{
    // ------------------------------------------------------------------ export

    /// <summary>
    /// Walks the document body into a <see cref="NeutralDoc"/>: heading-styled
    /// paragraphs become Heading blocks (Level from the style), list paragraphs
    /// become ListItem blocks (Level from the numbering level, Ordered from the
    /// numbering kind), plain paragraphs become Paragraph blocks carrying
    /// formatted <see cref="NeutralRun"/>s (bold/italic/underline/color and
    /// hyperlink href preserved), tables become Table blocks (cell-text grid,
    /// HeaderRow when the first row is a header) and inline images become Image
    /// blocks. The image bytes live inside the package, not on disk, so an Image
    /// block carries a stable <c>embedded:image{n}</c> token in <c>Source</c>
    /// plus its alt text; the integrator's media-extraction step (not this
    /// method) turns that into a real path. The title comes from the core Title
    /// property, falling back to the first Heading1.
    /// </summary>
    public NeutralDoc ExportNeutral(CommandContext ctx)
    {
        var file = RequireFile(ctx, mustExist: true);
        var (doc, ms, _) = OpenCopy(file, editable: false);
        using (doc)
        using (ms)
        {
            var body = GetBody(doc, file);
            var blocks = new List<NeutralBlock>();
            var imageCounter = 0;

            foreach (var child in body.ChildElements)
            {
                switch (child)
                {
                    case Paragraph paragraph:
                        AppendNeutralParagraph(doc, paragraph, blocks, ref imageCounter);
                        break;

                    case Table table:
                        blocks.Add(NeutralTable(doc, table));
                        break;

                    default:
                        break; // sectPr and other body-level metadata carry no neutral block
                }
            }

            return new NeutralDoc(ExportTitle(doc, blocks), blocks);
        }
    }

    /// <summary>Title from the standardized core property (docProps/core.xml), else the first Heading1 block.</summary>
    private static string? ExportTitle(WordprocessingDocument doc, IReadOnlyList<NeutralBlock> blocks)
    {
        if (ReadCoreTitle(doc) is { Length: > 0 } title)
        {
            return title;
        }

        var firstH1 = blocks.FirstOrDefault(b => b.Kind == NeutralBlockKind.Heading && b.Level == 1);
        var text = firstH1?.Runs is { } runs ? string.Concat(runs.Select(r => r.Text)) : null;
        return text is { Length: > 0 } ? text : null;
    }

    /// <summary>
    /// One body paragraph → 0..1 neutral blocks. An image-only paragraph yields
    /// an Image block per embedded drawing; otherwise the paragraph maps to a
    /// Heading / ListItem / Paragraph block by its style and numbering.
    /// </summary>
    private static void AppendNeutralParagraph(
        WordprocessingDocument doc, Paragraph paragraph, List<NeutralBlock> blocks, ref int imageCounter)
    {
        // Inline images: emit an Image block for each drawing the paragraph holds.
        var drawings = paragraph.Descendants<Drawing>().ToList();
        var runs = ReadNeutralRuns(doc, paragraph);

        if (drawings.Count > 0 && runs.Count == 0)
        {
            foreach (var drawing in drawings)
            {
                imageCounter++;
                var alt = drawing.Descendants<Dw.DocProperties>().FirstOrDefault()?.Description?.Value;
                blocks.Add(new NeutralBlock(
                    NeutralBlockKind.Image,
                    Source: $"embedded:image{imageCounter}",
                    Alt: alt is { Length: > 0 } ? alt : null));
            }

            return;
        }

        var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (HeadingLevel(style) is { } level)
        {
            blocks.Add(new NeutralBlock(NeutralBlockKind.Heading, Level: Math.Clamp(level, 1, 6), Runs: runs));
            return;
        }

        if (ListInfoOf(doc, paragraph) is { } info)
        {
            blocks.Add(new NeutralBlock(
                NeutralBlockKind.ListItem,
                Level: Math.Clamp(info.Level, 0, 8),
                Ordered: info.Kind == "number",
                Runs: runs));
            return;
        }

        // An empty / whitespace-only paragraph with no runs and no drawings that is
        // neither a heading nor a list item is a layout artifact — e.g. the blank
        // trailing paragraph 'create' leaves behind. Skipping it keeps junk empty
        // blocks (and the empty slides/blocks they spawn downstream) out of the
        // neutral model. Paragraphs that DO carry visible text still round-trip.
        if (drawings.Count == 0 && !runs.Any(r => r.Text.Any(c => !char.IsWhiteSpace(c))))
        {
            return;
        }

        blocks.Add(new NeutralBlock(NeutralBlockKind.Paragraph, Runs: runs));
    }

    /// <summary>
    /// The inline content of a paragraph as neutral runs: plain runs and the
    /// runs inside a hyperlink (each carrying the link href). Bold/italic/
    /// underline/color come off the run properties; empty-text runs are dropped.
    /// </summary>
    private static IReadOnlyList<NeutralRun> ReadNeutralRuns(WordprocessingDocument doc, Paragraph paragraph)
    {
        var result = new List<NeutralRun>();
        foreach (var child in paragraph.ChildElements)
        {
            switch (child)
            {
                case Run run:
                    AppendNeutralRun(result, run, href: null);
                    break;

                case Hyperlink hyperlink:
                {
                    var href = ResolveLinkUrl(doc, hyperlink)
                        ?? (hyperlink.Anchor?.Value is { Length: > 0 } anchor ? "#" + anchor : null);
                    foreach (var run in hyperlink.ChildElements.OfType<Run>())
                    {
                        AppendNeutralRun(result, run, href);
                    }

                    break;
                }

                case SimpleField field: // cached field result is the honest text shape
                    foreach (var run in field.ChildElements.OfType<Run>())
                    {
                        AppendNeutralRun(result, run, href: null);
                    }

                    break;

                case InsertedRun ins: // tracked insertions export at their accepted state
                    foreach (var run in ins.ChildElements.OfType<Run>())
                    {
                        AppendNeutralRun(result, run, href: null);
                    }

                    break;

                default: // DeletedRun, bookmarks, math, comment markers: no neutral run
                    break;
            }
        }

        return result;
    }

    /// <summary>Appends one run's text as a neutral run, unless it is image-only / empty.</summary>
    private static void AppendNeutralRun(List<NeutralRun> result, Run run, string? href)
    {
        var text = RunText(run);
        if (text.Length == 0)
        {
            return; // drawing-only or marker-only runs carry no inline text
        }

        var rPr = run.RunProperties;
        result.Add(new NeutralRun(
            text,
            Bold: WordFormatting.IsOn(rPr?.Bold) == true,
            Italic: WordFormatting.IsOn(rPr?.Italic) == true,
            Underline: WordFormatting.IsUnderlined(rPr) == true,
            Color: NormalizeColor(rPr?.Color?.Val?.Value),
            Href: href));
    }

    /// <summary>The visible text of a run: w:t plus tabs and breaks, no drawings.</summary>
    private static string RunText(Run run)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var piece in run.ChildElements)
        {
            switch (piece)
            {
                case Text text:
                    sb.Append(text.Text);
                    break;

                case TabChar:
                    sb.Append('\t');
                    break;

                case Break:
                    sb.Append('\n');
                    break;

                default:
                    break;
            }
        }

        return sb.ToString();
    }

    /// <summary>Drops the "auto" sentinel; keeps a real RRGGBB hex color.</summary>
    private static string? NormalizeColor(string? color) =>
        color is { Length: > 0 } && !color.Equals("auto", StringComparison.OrdinalIgnoreCase) ? color : null;

    /// <summary>A table → a Table block: a rectangular cell-text grid, HeaderRow when row 0 is a header.</summary>
    private static NeutralBlock NeutralTable(WordprocessingDocument doc, Table table)
    {
        var tableRows = table.ChildElements.OfType<TableRow>().ToList();
        var columns = Math.Max(1, GridColumnCount(table));
        var rows = new List<IReadOnlyList<string>>();

        foreach (var tableRow in tableRows)
        {
            var cells = new List<string>();
            foreach (var cell in tableRow.ChildElements.OfType<TableCell>())
            {
                var text = VerticalMergeState(cell) == "continue"
                    ? string.Empty // a swallowed vertical-merge slot reads as empty
                    : CellText(doc, cell);
                var span = CellSpan(cell);
                cells.Add(text);
                for (var s = 1; s < span; s++)
                {
                    cells.Add(string.Empty); // the neutral grid has no colspan; pad to grid width
                }
            }

            while (cells.Count < columns)
            {
                cells.Add(string.Empty);
            }

            if (cells.Count > columns)
            {
                cells.RemoveRange(columns, cells.Count - columns);
            }

            rows.Add(cells);
        }

        var headerRow = tableRows.Count > 0 && IsHeaderRow(tableRows[0]);
        return new NeutralBlock(NeutralBlockKind.Table, Rows: rows, HeaderRow: headerRow);
    }

    /// <summary>The text of one cell: its paragraphs joined by newlines (plain text, no markup).</summary>
    private static string CellText(WordprocessingDocument doc, TableCell cell)
    {
        var paragraphs = cell.ChildElements.OfType<Paragraph>()
            .Select(p => string.Concat(ReadNeutralRuns(doc, p).Select(r => r.Text)))
            .Where(t => t.Length > 0);
        return string.Join('\n', paragraphs);
    }

    /// <summary>A header row carries w:tblHeader, or every cell is bold (the convention the importer writes).</summary>
    private static bool IsHeaderRow(TableRow row)
    {
        if (row.TableRowProperties?.Elements<TableHeader>().Any(h => h.Val?.Value != OnOffOnlyValues.Off) == true)
        {
            return true;
        }

        var cells = row.ChildElements.OfType<TableCell>().ToList();
        if (cells.Count == 0)
        {
            return false;
        }

        return cells.All(c =>
        {
            var runs = c.Descendants<Run>().Where(r => r.InnerText.Length > 0).ToList();
            return runs.Count > 0 && runs.All(r => WordFormatting.IsOn(r.RunProperties?.Bold) == true);
        });
    }

    // ------------------------------------------------------------------ import

    /// <summary>
    /// Writes a <see cref="NeutralDoc"/> into THIS (freshly created, empty) docx:
    /// Heading blocks become Heading{level} paragraphs, Paragraph blocks become
    /// runs (bold/italic/underline/color, hyperlinks for Href), ListItem blocks
    /// ride the M3 numbering machinery (bullet/numbered, indent from Level),
    /// Table blocks become real Word tables (bold + repeated header row) and
    /// Image blocks embed when <c>Source</c> resolves to a readable image,
    /// otherwise land in <see cref="ImportResult.Dropped"/>. The neutral title
    /// becomes the core Title property. The doc is opened editable in place and
    /// saved on dispose.
    /// </summary>
    public ImportResult ImportNeutral(CommandContext ctx, NeutralDoc doc)
    {
        var file = RequireFile(ctx, mustExist: true);
        var dropped = new List<string>();
        var written = 0;

        // Atomic: build the whole package in an in-memory copy, save it (on the
        // wDoc using-scope close), then read the bytes back from the still-open ms.
        var (wDoc, ms, _) = OpenCopy(file, editable: true);
        using (ms)
        {
            using (wDoc)
            {
                var body = GetBody(wDoc, file);

                // A freshly created docx carries one empty placeholder paragraph; clear
                // the body so the imported content is the whole document.
                foreach (var stale in body.ChildElements.OfType<Paragraph>().ToList())
                {
                    if (stale.InnerText.Length == 0 && !stale.Descendants<Drawing>().Any())
                    {
                        stale.Remove();
                    }
                }

                // Per-kind numbering instances so each ordered/bullet run numbers cleanly.
                var listNumIds = new Dictionary<(bool Ordered, int Level), int>();
                var imageOrdinal = 0;

                foreach (var block in doc.Blocks)
                {
                    switch (block.Kind)
                    {
                        case NeutralBlockKind.Heading:
                            body.AppendChild(ImportHeading(wDoc, block));
                            written++;
                            break;

                        case NeutralBlockKind.Paragraph:
                            body.AppendChild(ImportParagraph(wDoc, block, list: null));
                            written++;
                            break;

                        case NeutralBlockKind.ListItem:
                            body.AppendChild(ImportListItem(wDoc, block, listNumIds));
                            written++;
                            break;

                        case NeutralBlockKind.Table:
                            body.AppendChild(ImportTable(wDoc, block));
                            written++;
                            break;

                        case NeutralBlockKind.Image:
                            imageOrdinal++;
                            if (TryImportImage(wDoc, ctx.Workspace, body, block, imageOrdinal, out var note))
                            {
                                written++;
                            }
                            else
                            {
                                dropped.Add(note);
                            }

                            break;

                        default:
                            dropped.Add($"Unknown neutral block kind '{block.Kind}' was skipped.");
                            break;
                    }
                }

                // An all-empty import still leaves a valid one-paragraph docx.
                if (!body.ChildElements.Any(c => c is Paragraph or Table))
                {
                    body.AppendChild(new Paragraph());
                }

                if (doc.Title is { Length: > 0 } title)
                {
                    WriteCoreField(wDoc, CoreField.Title, title);
                }
            } // wDoc disposed here → the package is flushed into the still-open ms

            // Write the rebuilt package back, snapshotting the on-disk pre-image first.
            var bytes = ms.ToArray();
            _snapshots.Save(file); // pre-image, so the import is undoable
            File.WriteAllBytes(file, bytes);
        }

        return new ImportResult(written, dropped);
    }

    private static Paragraph ImportHeading(WordprocessingDocument doc, NeutralBlock block)
    {
        var level = Math.Clamp(block.Level == 0 ? 1 : block.Level, 1, 6);
        EnsureStyleDefined(doc, "Heading" + level);
        var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading" + level }));
        AppendNeutralRuns(doc, paragraph, block.Runs);
        return paragraph;
    }

    private static Paragraph ImportParagraph(
        WordprocessingDocument doc, NeutralBlock block, (int NumId, int Level)? list)
    {
        var paragraph = new Paragraph();
        AppendNeutralRuns(doc, paragraph, block.Runs);
        if (list is { } marker)
        {
            var pPr = paragraph.ParagraphProperties ??= new ParagraphProperties();
            pPr.NumberingProperties = new NumberingProperties(
                new NumberingLevelReference { Val = marker.Level },
                new NumberingId { Val = marker.NumId });
        }

        return paragraph;
    }

    private static Paragraph ImportListItem(
        WordprocessingDocument doc, NeutralBlock block, Dictionary<(bool, int), int> listNumIds)
    {
        var level = Math.Clamp(block.Level, 0, MaxListLevel);
        var key = (block.Ordered, level);
        if (!listNumIds.TryGetValue(key, out var numId))
        {
            numId = NewNumberingInstance(doc, block.Ordered ? "number" : "bullet");
            listNumIds[key] = numId;
        }

        return ImportParagraph(doc, block, (numId, level));
    }

    private static Table ImportTable(WordprocessingDocument doc, NeutralBlock block)
    {
        var rows = block.Rows ?? [];
        var columns = Math.Max(1, rows.Select(r => r.Count).DefaultIfEmpty(0).Max());

        var table = new Table();
        table.AppendChild(new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })));

        var grid = new TableGrid();
        for (var c = 0; c < columns; c++)
        {
            grid.AppendChild(new GridColumn());
        }

        table.AppendChild(grid);

        for (var r = 0; r < rows.Count; r++)
        {
            var isHeader = block.HeaderRow && r == 0;
            var row = new TableRow();
            if (isHeader)
            {
                // Mark the row to repeat as a header on every page (w:tblHeader).
                row.AppendChild(new TableRowProperties(new TableHeader()));
            }

            for (var c = 0; c < columns; c++)
            {
                var text = c < rows[r].Count ? rows[r][c] : string.Empty;
                var paragraph = new Paragraph();
                if (text.Length > 0)
                {
                    var run = new Run(NewText(text));
                    if (isHeader)
                    {
                        run.RunProperties = new RunProperties(new Bold());
                    }

                    paragraph.AppendChild(run);
                }

                row.AppendChild(new TableCell(paragraph));
            }

            table.AppendChild(row);
        }

        if (rows.Count == 0)
        {
            table.AppendChild(new TableRow(new TableCell(new Paragraph())));
        }

        return table;
    }

    /// <summary>
    /// Embeds the Image block when its <c>Source</c> resolves through the sandbox
    /// to a readable PNG/JPEG; otherwise reports a Dropped note (an
    /// <c>embedded:</c> export token, a missing file, a sandbox escape or an
    /// unsupported format never fails the import).
    /// </summary>
    private static bool TryImportImage(
        WordprocessingDocument doc, Workspace workspace, Body body, NeutralBlock block, int ordinal, out string note)
    {
        var src = block.Source;
        if (string.IsNullOrWhiteSpace(src) || src.StartsWith("embedded:", StringComparison.Ordinal))
        {
            note = $"Image {ordinal}" +
                (block.Alt is { Length: > 0 } alt ? $" (\"{alt}\")" : string.Empty) +
                " had no readable source path and was dropped; its bytes live inside the source package, not on disk.";
            return false;
        }

        try
        {
            var resolved = workspace.Resolve(src, mustExist: true);
            var bytes = File.ReadAllBytes(resolved);
            var (format, pixelWidth, pixelHeight) = SniffImage(bytes, src);

            var main = doc.MainDocumentPart!;
            var imagePart = main.AddImagePart(format == "png" ? ImagePartType.Png : ImagePartType.Jpeg);
            using (var stream = new MemoryStream(bytes, writable: false))
            {
                imagePart.FeedData(stream);
            }

            var cx = (long)Math.Round(Math.Max(1, pixelWidth) * EmuPerPixel);
            var cy = (long)Math.Round(Math.Max(1, pixelHeight) * EmuPerPixel);
            var name = block.Alt is { Length: > 0 } ? block.Alt : Path.GetFileName(src);
            body.AppendChild(new Paragraph(new Run(BuildInlineDrawing(
                main.GetIdOfPart(imagePart),
                NextDrawingId(doc),
                name,
                cx,
                cy,
                block.Alt is { Length: > 0 } ? block.Alt : null))));

            note = string.Empty;
            return true;
        }
        catch (AiofficeException ex)
        {
            note = $"Image '{Snippet(src)}' was dropped: {ex.Message}";
            return false;
        }
    }

    /// <summary>Writes a block's neutral runs into a paragraph, grouping consecutive same-href runs into one hyperlink.</summary>
    private static void AppendNeutralRuns(WordprocessingDocument doc, Paragraph paragraph, IReadOnlyList<NeutralRun>? runs)
    {
        if (runs is null || runs.Count == 0)
        {
            return;
        }

        var i = 0;
        while (i < runs.Count)
        {
            var href = runs[i].Href;
            if (href is { Length: > 0 })
            {
                // Coalesce the run(s) sharing this href into one hyperlink element.
                var group = new List<NeutralRun>();
                while (i < runs.Count && runs[i].Href == href)
                {
                    group.Add(runs[i]);
                    i++;
                }

                AppendNeutralHyperlink(doc, paragraph, group, href);
            }
            else
            {
                paragraph.AppendChild(BuildNeutralRun(runs[i]));
                i++;
            }
        }
    }

    private static void AppendNeutralHyperlink(
        WordprocessingDocument doc, Paragraph paragraph, IReadOnlyList<NeutralRun> runs, string href)
    {
        // Internal anchors ("#bookmark") become an anchor link; absolute http(s)/mailto
        // urls become an external relationship; anything else degrades to plain runs.
        if (href.StartsWith('#'))
        {
            EnsureHyperlinkStyle(doc);
            var anchorLink = new Hyperlink { Anchor = href[1..] };
            foreach (var run in runs)
            {
                anchorLink.AppendChild(BuildNeutralRun(run, hyperlinkStyle: true));
            }

            paragraph.AppendChild(anchorLink);
            return;
        }

        if (Uri.TryCreate(href, UriKind.Absolute, out var uri) &&
            AllowedLinkSchemes.Contains(uri.Scheme, StringComparer.OrdinalIgnoreCase))
        {
            EnsureHyperlinkStyle(doc);
            var hyperlink = new Hyperlink
            {
                Id = doc.MainDocumentPart!.AddHyperlinkRelationship(uri, isExternal: true).Id,
            };
            foreach (var run in runs)
            {
                hyperlink.AppendChild(BuildNeutralRun(run, hyperlinkStyle: true));
            }

            paragraph.AppendChild(hyperlink);
            return;
        }

        // Non-absolute, non-anchor href: keep the text, drop the (unrepresentable) link.
        foreach (var run in runs)
        {
            paragraph.AppendChild(BuildNeutralRun(run));
        }
    }

    /// <summary>One neutral run → a w:r with the matching run properties.</summary>
    private static Run BuildNeutralRun(NeutralRun neutral, bool hyperlinkStyle = false)
    {
        var run = new Run();
        RunProperties? rPr = null;

        if (hyperlinkStyle)
        {
            rPr = new RunProperties { RunStyle = new RunStyle { Val = "Hyperlink" } };
        }

        if (neutral.Bold)
        {
            (rPr ??= new RunProperties()).Bold = new Bold();
        }

        if (neutral.Italic)
        {
            (rPr ??= new RunProperties()).Italic = new Italic();
        }

        if (neutral.Underline)
        {
            (rPr ??= new RunProperties()).Underline = new Underline { Val = UnderlineValues.Single };
        }

        if (NormalizeColor(neutral.Color) is { } color)
        {
            (rPr ??= new RunProperties()).Color = new Color { Val = SanitizeHex(color) };
        }

        if (rPr is not null)
        {
            run.RunProperties = rPr;
        }

        AppendRunText(run, neutral.Text);
        return run;
    }

    /// <summary>Splits a run's text on newlines into w:t + w:br so embedded breaks survive.</summary>
    private static void AppendRunText(Run run, string text)
    {
        var lines = text.Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            if (i > 0)
            {
                run.AppendChild(new Break());
            }

            run.AppendChild(NewText(lines[i]));
        }
    }

    /// <summary>Keeps a 6-hex-digit color, normalizing case; falls back to "auto" for anything else.</summary>
    private static string SanitizeHex(string color)
    {
        var trimmed = color.TrimStart('#');
        return trimmed.Length == 6 && trimmed.All(Uri.IsHexDigit)
            ? trimmed.ToUpperInvariant()
            : "auto";
    }
}
