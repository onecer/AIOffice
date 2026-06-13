using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// The M9 conversion surface for pptx: project a deck down to the format-neutral
/// model (<see cref="ExportNeutral"/>) and absorb a neutral model into a freshly
/// created deck as slides (<see cref="ImportNeutral"/>).
///
/// Mapping (export): each slide contributes a Heading (its title, level 1)
/// followed by its body — body/text-box paragraphs become Paragraph or ListItem
/// blocks (a bullet's indent level rides on <see cref="NeutralBlock.Level"/>),
/// tables become Table blocks, pictures become Image blocks (Alt from alt-text),
/// and speaker notes become a "Notes: "-prefixed Paragraph. Animations,
/// transitions, charts and SmartArt cannot cross to the neutral model and are
/// named in the export's <see cref="NeutralDoc"/> only through the import side's
/// Dropped notes — on export they would be silently lost, so the round-trip law
/// is "content survives; effects do not".
///
/// Mapping (import): every top-level Heading (level &lt;= 2) opens a NEW slide
/// titled with its text; the Paragraph/ListItem blocks that follow become that
/// slide's body bullets (Level → bullet indent); a Table block lands as a native
/// pptx table; an Image lands as a picture; deeper headings become bold body
/// lines. Anything the neutral model could not carry from the source is honestly
/// reported in <see cref="ImportResult.Dropped"/>.
/// </summary>
public sealed partial class PptxHandler : INeutralConvertible
{
    /// <summary>Body bullets layout box (below the title band AddTitleShape uses), in EMU.</summary>
    private const long BodyOffsetXEmu = 831_850L;

    private const long BodyOffsetYEmu = 1_825_625L;

    private const long BodyWidthEmu = 10_515_600L;

    private const long BodyHeightEmu = 4_351_338L;

    /// <summary>The deepest bullet indent the schema allows (a:pPr/@lvl is 0..8).</summary>
    private const int MaxBulletLevel = 8;

    public NeutralDoc ExportNeutral(CommandContext ctx) => ExportNeutral(ctx, out _);

    /// <summary>
    /// Projects the deck to the neutral model and, via <paramref name="dropped"/>,
    /// names the export-side losses the neutral model cannot carry (animations,
    /// transitions, charts, SmartArt). The command layer folds these into the
    /// same <c>convert_lossy</c> report as the import-side <see cref="ImportResult.Dropped"/>.
    /// </summary>
    public NeutralDoc ExportNeutral(CommandContext ctx, out IReadOnlyList<string> dropped)
    {
        var file = ctx.File ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "Conversion requires a target file.",
            "Pass the .pptx path to read from.");

        using var stream = PptxDoc.LoadStream(file);
        using var doc = PptxDoc.Open(stream, editable: false, file);
        var presentation = PptxDoc.RequirePresentationPart(doc, file);

        var blocks = new List<NeutralBlock>();
        var lost = new List<string>();
        var slides = PptxDoc.Slides(presentation);

        string? firstTitle = null;
        foreach (var (_, slidePart) in slides)
        {
            ExportSlide(slidePart, blocks, lost, ref firstTitle);
        }

        dropped = lost.Distinct(StringComparer.Ordinal).ToList();

        // The deck title is the standard core property (docProps/core.xml) when
        // set, else the first slide's title. PptxCoreProps migrates a legacy-only
        // file on read so older decks still surface their title here.
        var title = NullIfBlank(PptxCoreProps.Read(doc)?.Title)
            ?? NullIfBlank(doc.PackageProperties.Title)
            ?? firstTitle;
        return new NeutralDoc(title, blocks);
    }

    private static void ExportSlide(SlidePart slidePart, List<NeutralBlock> blocks, List<string> dropped, ref string? firstTitle)
    {
        var shapes = PptxDoc.Shapes(slidePart);

        // The title placeholder opens the slide as a level-1 heading.
        var titleShape = shapes.FirstOrDefault(s => PptxDoc.PlaceholderType(s.Element) is "title" or "ctrTitle");
        var titleText = titleShape is not null ? PptxDoc.ShapeText(titleShape.Element).Trim() : string.Empty;
        if (titleText.Length > 0)
        {
            firstTitle ??= titleText;
            blocks.Add(new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun(titleText)]));
        }

        // Body content in document order: text shapes → paragraphs/bullets,
        // tables → Table blocks, pictures → Image blocks. Charts and SmartArt
        // frames have no neutral home — their content is dropped (named below).
        foreach (var shape in shapes)
        {
            if (ReferenceEquals(shape, titleShape))
            {
                continue;
            }

            switch (shape.Element)
            {
                case P.Shape textShape when PptxSmartArt.DataPartOf(slidePart, textShape) is null:
                    ExportTextShape(textShape, blocks);
                    break;

                case P.Picture picture:
                    blocks.Add(new NeutralBlock(
                        NeutralBlockKind.Image,
                        Source: NullIfBlank(picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value),
                        Alt: PptxDoc.AltText(picture)));
                    break;

                case P.GraphicFrame frame when PptxTables.TableOf(frame) is { } table:
                    ExportTable(table, blocks);
                    break;

                default:
                    break; // charts, SmartArt frames, connectors, groups: no neutral block
            }
        }

        // Effects and rich graphics the neutral model cannot carry: name them
        // honestly so the round-trip reports the loss instead of hiding it.
        if (PptxCharts.Charts(slidePart).Count > 0)
        {
            dropped.Add("charts (transferred as data is not supported; the chart is not converted)");
        }

        if (PptxSmartArt.List(slidePart).Count > 0)
        {
            dropped.Add("SmartArt diagrams (no neutral equivalent; the diagram is not converted)");
        }

        if (PptxAnimations.List(slidePart).Count > 0)
        {
            dropped.Add("slide animations (the neutral model carries content, not motion)");
        }

        if (PptxTransitions.Read(slidePart) is { Kind: not "none" })
        {
            dropped.Add("slide transitions (the neutral model carries content, not motion)");
        }

        if (PptxEmbeds.SlideEmbedCount(slidePart) > 0)
        {
            dropped.Add("embedded objects (OLE/package attachments have no neutral equivalent; the embed is not converted)");
        }

        // Speaker notes ride along as a tagged paragraph so they survive the trip.
        var notes = PptxNotes.Text(slidePart).Trim();
        if (notes.Length > 0)
        {
            blocks.Add(new NeutralBlock(
                NeutralBlockKind.Paragraph,
                Runs: [new NeutralRun("Notes: " + notes)]));
        }
    }

    /// <summary>A non-title text shape: each paragraph becomes a Paragraph or, when it is a body bullet, a ListItem.</summary>
    private static void ExportTextShape(P.Shape shape, List<NeutralBlock> blocks)
    {
        var body = shape.TextBody;
        if (body is null)
        {
            return;
        }

        // A body/subtitle placeholder treats its paragraphs as bullets; a plain
        // text box treats them as paragraphs (unless a paragraph carries an
        // explicit indent level, which always means a bullet).
        var isBulletShape = PptxDoc.PlaceholderType(shape) is "body" or "subTitle";

        foreach (var paragraph in body.Elements<A.Paragraph>())
        {
            var runs = ExportRuns(paragraph);
            if (runs.Count == 0)
            {
                continue; // skip empty paragraphs — they carry no transferable content
            }

            var lvl = paragraph.ParagraphProperties?.Level?.Value;
            if (isBulletShape || lvl is not null)
            {
                blocks.Add(new NeutralBlock(
                    NeutralBlockKind.ListItem,
                    Level: Math.Clamp(lvl ?? 0, 0, MaxBulletLevel),
                    Ordered: false,
                    Runs: runs));
            }
            else
            {
                blocks.Add(new NeutralBlock(NeutralBlockKind.Paragraph, Runs: runs));
            }
        }
    }

    /// <summary>The inline runs of a paragraph, carrying the formatting the neutral model preserves.</summary>
    private static List<NeutralRun> ExportRuns(A.Paragraph paragraph)
    {
        var runs = new List<NeutralRun>();
        foreach (var run in paragraph.Elements<A.Run>())
        {
            var text = run.Text?.Text ?? string.Empty;
            if (text.Length == 0)
            {
                continue;
            }

            var rPr = run.RunProperties;
            runs.Add(new NeutralRun(
                text,
                Bold: rPr?.Bold?.Value == true,
                Italic: rPr?.Italic?.Value == true,
                Underline: rPr?.Underline?.Value is { } u && u != A.TextUnderlineValues.None,
                Color: rPr?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant()));
        }

        return runs;
    }

    private static void ExportTable(A.Table table, List<NeutralBlock> blocks)
    {
        var rows = new List<IReadOnlyList<string>>();
        foreach (var row in table.Elements<A.TableRow>())
        {
            var cells = new List<string>();
            foreach (var cell in row.Elements<A.TableCell>())
            {
                cells.Add(PptxTables.CellText(cell));
            }

            if (cells.Count > 0)
            {
                rows.Add(cells);
            }
        }

        if (rows.Count == 0)
        {
            return;
        }

        var headerRow = table.TableProperties?.FirstRow?.Value == true;
        blocks.Add(new NeutralBlock(NeutralBlockKind.Table, Rows: rows, HeaderRow: headerRow));
    }

    // ----------------------------------------------------------------- import

    public ImportResult ImportNeutral(CommandContext ctx, NeutralDoc doc)
    {
        var file = ctx.File ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "Conversion requires a target file.",
            "Pass the .pptx path to write into.");

        var dropped = new List<string>();
        var written = 0;

        using var stream = PptxDoc.LoadStream(file);
        using (var doc2 = PptxDoc.Open(stream, editable: true, file))
        {
            var presentation = PptxDoc.RequirePresentationPart(doc2, file);
            if (NullIfBlank(doc.Title) is { } title)
            {
                // Write the converted deck title to the standard docProps/core.xml
                // part so it is visible to Office and round-trips through
                // read --view properties (not the non-standard PackageProperties).
                // Seed from any existing core props so a title set on an existing
                // deck does not wipe its other metadata.
                var core = PptxCoreProps.Read(doc2) ?? new PptxCoreProps.CoreModel();
                core.Title = title;
                PptxCoreProps.Write(doc2, core);
            }

            // A freshly created deck already holds one blank slide; the first
            // heading reuses it, every later heading appends a new one.
            var slides = PptxDoc.Slides(presentation);
            var blankReused = false;
            SlidePart? current = slides.Count > 0 ? slides[0].Part : null;
            var bodyShape = (P.Shape?)null; // the current slide's body placeholder, lazily created

            foreach (var block in doc.Blocks)
            {
                switch (block.Kind)
                {
                    case NeutralBlockKind.Heading when block.Level <= 2:
                    {
                        var text = RunsText(block.Runs);
                        if (!blankReused && current is not null)
                        {
                            // Reuse the deck's seed slide for the first heading.
                            blankReused = true;
                        }
                        else
                        {
                            current = AppendSlide(presentation);
                        }

                        bodyShape = null;
                        if (text.Length > 0)
                        {
                            PptxEditor.AddTitleShape(current!, text);
                        }

                        written++;
                        break;
                    }

                    case NeutralBlockKind.Heading:
                        // A deeper heading (level >= 3) has no slide of its own:
                        // it becomes a bold body line on the current slide.
                        current = EnsureSlide(presentation, current, ref blankReused);
                        AppendBodyParagraph(ref bodyShape, current, RunsText(block.Runs), level: 0, bold: true);
                        written++;
                        break;

                    case NeutralBlockKind.Paragraph:
                        current = EnsureSlide(presentation, current, ref blankReused);
                        AppendBodyParagraph(ref bodyShape, current, RunsText(block.Runs), level: 0, bold: false);
                        written++;
                        break;

                    case NeutralBlockKind.ListItem:
                        current = EnsureSlide(presentation, current, ref blankReused);
                        AppendBodyParagraph(
                            ref bodyShape,
                            current,
                            RunsText(block.Runs),
                            level: Math.Clamp(block.Level, 0, MaxBulletLevel),
                            bold: false);
                        if (block.Ordered)
                        {
                            dropped.Add("ordered list numbering on a bullet item (rendered as a plain bullet)");
                        }

                        written++;
                        break;

                    case NeutralBlockKind.Table:
                        current = EnsureSlide(presentation, current, ref blankReused);
                        ImportTable(current, block);
                        bodyShape = null; // a table closes the running body shape
                        written++;
                        break;

                    case NeutralBlockKind.Image:
                        current = EnsureSlide(presentation, current, ref blankReused);
                        if (ImportImage(current, block, ctx.Workspace, dropped))
                        {
                            written++;
                        }

                        bodyShape = null;
                        break;

                    default:
                        break;
                }
            }
        }

        File.WriteAllBytes(file, stream.ToArray());

        // De-duplicate the per-occurrence notes so the lossy report stays terse.
        var distinctDropped = dropped.Distinct(StringComparer.Ordinal).ToList();
        return new ImportResult(written, distinctDropped);
    }

    /// <summary>
    /// Resolves the slide a body block lands on: the current slide when one is
    /// open, else the deck's seed slide (consumed on first use), else a fresh
    /// slide. Either way the seed slide counts as used, so a later first heading
    /// starts its own slide instead of overwriting this content.
    /// </summary>
    private static SlidePart EnsureSlide(PresentationPart presentation, SlidePart? current, ref bool blankReused)
    {
        blankReused = true;
        if (current is not null)
        {
            return current;
        }

        var slides = PptxDoc.Slides(presentation);
        return slides.Count > 0 ? slides[^1].Part : AppendSlide(presentation);
    }

    /// <summary>Appends a fresh blank slide bound to the master's first layout, returning its part.</summary>
    private static SlidePart AppendSlide(PresentationPart presentation)
    {
        var slideIdList = presentation.Presentation?.SlideIdList
            ?? throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                "The presentation has no slide id list (p:sldIdLst).",
                "Re-create the deck with 'aioffice create', then convert into it.");

        var masters = PptxDoc.Masters(presentation);
        if (masters.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                "The presentation has no slide master.",
                "Re-create the deck with 'aioffice create', then convert into it.");
        }

        var layouts = PptxDoc.Layouts(masters[0].Part);
        if (layouts.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                "The slide master has no layout.",
                "Re-create the deck with 'aioffice create', then convert into it.");
        }

        var slidePart = presentation.AddNewPart<SlidePart>();
        slidePart.Slide = PptxFactory.BuildBlankSlide();
        slidePart.AddPart(layouts[0].Part);

        slideIdList.Append(new P.SlideId
        {
            Id = PptxDoc.NextSlideId(slideIdList),
            RelationshipId = presentation.GetIdOfPart(slidePart),
        });

        return slidePart;
    }

    /// <summary>
    /// Appends one body paragraph to the slide's running body placeholder,
    /// creating the placeholder on first use. The indent <paramref name="level"/>
    /// rides on a:pPr/@lvl so bullets nest the way export reads them back.
    /// </summary>
    private static void AppendBodyParagraph(ref P.Shape? bodyShape, SlidePart slidePart, string text, int level, bool bold)
    {
        bodyShape ??= EnsureBodyShape(slidePart);
        var paragraph = new A.Paragraph();
        if (level > 0)
        {
            paragraph.Append(new A.ParagraphProperties { Level = level });
        }

        var runProperties = new A.RunProperties { Language = "en-US" };
        if (bold)
        {
            runProperties.Bold = true;
        }

        paragraph.Append(new A.Run(runProperties, new A.Text(text)));
        bodyShape.TextBody!.Append(paragraph);
    }

    /// <summary>The slide's body placeholder, created (empty) on first use under the title band.</summary>
    private static P.Shape EnsureBodyShape(SlidePart slidePart)
    {
        var tree = PptxDoc.RequireShapeTree(slidePart);
        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = PptxDoc.NextShapeId(tree), Name = "Content Placeholder" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(
                    new P.PlaceholderShape { Type = P.PlaceholderValues.Body, Index = (UInt32Value)1U })),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = BodyOffsetXEmu, Y = BodyOffsetYEmu },
                    new A.Extents { Cx = BodyWidthEmu, Cy = BodyHeightEmu }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new P.TextBody(new A.BodyProperties(), new A.ListStyle()));
        tree.Append(shape);
        return shape;
    }

    private static void ImportTable(SlidePart slidePart, NeutralBlock block)
    {
        var rows = block.Rows ?? [];
        var rowCount = Math.Max(1, rows.Count);
        var colCount = Math.Max(1, rows.Select(r => r.Count).DefaultIfEmpty(1).Max());

        var props = new System.Text.Json.Nodes.JsonObject
        {
            ["rows"] = rowCount,
            ["cols"] = colCount,
            ["headerRow"] = block.HeaderRow,
        };
        PptxTables.Add(slidePart, props);

        // Fill the freshly added table's cells with the block's text.
        var tables = PptxTables.Tables(slidePart);
        var table = tables[^1].Table;
        var tableRows = table.Elements<A.TableRow>().ToList();
        for (var r = 0; r < tableRows.Count && r < rows.Count; r++)
        {
            var cells = tableRows[r].Elements<A.TableCell>().ToList();
            var sourceRow = rows[r];
            for (var c = 0; c < cells.Count && c < sourceRow.Count; c++)
            {
                SetCellText(cells[c], sourceRow[c]);
            }
        }
    }

    /// <summary>Replaces a cell's text, keeping the look's first run formatting as the prototype.</summary>
    private static void SetCellText(A.TableCell cell, string text)
    {
        var textBody = cell.TextBody ??= new A.TextBody(new A.BodyProperties(), new A.ListStyle());
        var prototype = textBody.Descendants<A.RunProperties>().FirstOrDefault();

        foreach (var paragraph in textBody.Elements<A.Paragraph>().ToList())
        {
            paragraph.Remove();
        }

        var run = new A.Run();
        run.Append(prototype is not null
            ? (A.RunProperties)prototype.CloneNode(true)
            : new A.RunProperties { Language = "en-US" });
        run.Append(new A.Text(text));
        textBody.Append(new A.Paragraph(run));
    }

    /// <summary>
    /// Embeds the image when its source resolves to a real PNG/JPEG in the
    /// workspace; otherwise records an honest Dropped note and skips it (a
    /// missing image is never a failed conversion).
    /// </summary>
    private static bool ImportImage(SlidePart slidePart, NeutralBlock block, Workspace workspace, List<string> dropped)
    {
        if (NullIfBlank(block.Source) is not { } source)
        {
            dropped.Add("an image with no source path");
            return false;
        }

        var props = new System.Text.Json.Nodes.JsonObject { ["src"] = source };
        if (NullIfBlank(block.Alt) is { } alt)
        {
            props["name"] = alt;
        }

        try
        {
            var id = PptxImages.AddImage(slidePart, props, workspace);
            if (NullIfBlank(block.Alt) is { } altText)
            {
                // Carry the alt text onto the picture's accessibility description.
                var picture = PptxDoc.Shapes(slidePart).FirstOrDefault(s => s.Id == id)?.Element;
                if (PptxDoc.NonVisualProps((OpenXmlCompositeElement)picture!) is { } nv)
                {
                    nv.Description = altText;
                }
            }

            return true;
        }
        catch (AiofficeException ex)
        {
            dropped.Add($"image '{source}' ({ex.Message})");
            return false;
        }
    }

    private static string RunsText(IReadOnlyList<NeutralRun>? runs) =>
        runs is null ? string.Empty : string.Concat(runs.Select(r => r.Text));

    private static string? NullIfBlank(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;
}
