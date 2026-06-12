using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig;
using Markdig.Extensions.EmphasisExtras;
using Md = Markdig.Syntax;
using MdInline = Markdig.Syntax.Inlines;
using MdTable = Markdig.Extensions.Tables;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    /// <summary>Inline formatting carried down the Markdig inline tree.</summary>
    private readonly record struct MdFormat(bool Bold, bool Italic, bool Strike, bool Code, bool Link);

    /// <summary>Everything the import walker threads along.</summary>
    private sealed record MdImportState(
        WordprocessingDocument Doc,
        Workspace Workspace,
        string? SourceDir,
        List<Warning> Warnings);

    /// <summary>
    /// The markdown bridge, import side: <c>aioffice create report.docx --from notes.md</c>.
    /// Parses the source through Markdig (GFM pipe tables, strikethrough, YAML
    /// front matter ignored) and emits a real docx: heading styles, formatted
    /// runs, code blocks, nested lists on the M3 numbering machinery, real
    /// hyperlinks/images/tables, blockquotes and horizontal rules. Raw HTML and
    /// missing images degrade to warnings, never failures.
    /// </summary>
    public Envelope CreateFrom(CommandContext ctx, string sourcePath)
    {
        var file = RequireFile(ctx);
        if (File.Exists(file))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"File already exists: {file}",
                "Pick a new file name, or edit the existing document with 'aioffice edit'.");
        }

        if (string.IsNullOrWhiteSpace(sourcePath))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "create --from needs a source file path.",
                "Example: aioffice create report.docx --from notes.md.");
        }

        // SECURITY: the only road to the source bytes is through the sandbox.
        var resolvedSource = ctx.Workspace.Resolve(sourcePath, mustExist: true);
        if (!File.Exists(resolvedSource))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"Source file not found: {sourcePath}",
                "Check the path; it is resolved relative to the workspace root.");
        }

        var extension = Path.GetExtension(resolvedSource);
        if (!extension.Equals(".md", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".markdown", StringComparison.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"docx import reads markdown sources only; got '{extension}'.",
                "Pass a .md file, or create an empty document and add content with 'aioffice edit'.",
                candidates: [".md", ".markdown"]);
        }

        FileSizeGuard.Ensure(resolvedSource);
        var markdown = File.ReadAllText(resolvedSource);
        var warnings = new List<Warning>();

        byte[] bytes;
        using (var ms = new MemoryStream())
        {
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body());
                WordFactory.AddDefaultStylesPart(main);

                var body = main.Document.Body!;
                var state = new MdImportState(doc, ctx.Workspace, Path.GetDirectoryName(resolvedSource), warnings);
                ImportMarkdown(body, markdown, state);

                if (!body.ChildElements.Any(c => c is Paragraph or Table))
                {
                    body.AppendChild(new Paragraph()); // an empty source still yields a valid one-paragraph docx
                }
            }

            bytes = ms.ToArray();
        }

        var dir = Path.GetDirectoryName(file);
        if (dir is { Length: > 0 })
        {
            Directory.CreateDirectory(dir);
        }

        File.WriteAllBytes(file, bytes);

        return Envelope.Ok(
            new { created = file, kind = "docx", source = sourcePath, format = "markdown" },
            MetaFor(file, Rev.OfBytes(bytes), warnings));
    }

    // ----------------------------------------------------------------- blocks

    private static void ImportMarkdown(Body body, string markdown, MdImportState state)
    {
        var pipeline = new MarkdownPipelineBuilder()
            .UsePipeTables()
            .UseEmphasisExtras(EmphasisExtraOptions.Strikethrough)
            .UseYamlFrontMatter() // front matter parses as a block we then ignore
            .Build();

        foreach (var block in Markdown.Parse(markdown, pipeline))
        {
            AppendMdBlock(body, block, state, quoteDepth: 0, list: null);
        }
    }

    private static void AppendMdBlock(
        OpenXmlElement container,
        Md.Block block,
        MdImportState state,
        int quoteDepth,
        (int NumId, int Level)? list)
    {
        switch (block)
        {
            case Md.HeadingBlock heading:
            {
                var level = Math.Clamp(heading.Level, 1, 6);
                EnsureStyleDefined(state.Doc, "Heading" + level);
                var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading" + level }));
                AppendMdInlines(paragraph, heading.Inline, state, default);
                ApplyQuoteFormatting(paragraph, quoteDepth);
                container.AppendChild(paragraph);
                break;
            }

            case Md.ParagraphBlock paragraphBlock:
            {
                var paragraph = new Paragraph();
                AppendMdInlines(paragraph, paragraphBlock.Inline, state, default);
                ApplyQuoteFormatting(paragraph, quoteDepth);
                if (list is { } marker)
                {
                    var pPr = paragraph.ParagraphProperties ??= new ParagraphProperties();
                    pPr.NumberingProperties = new NumberingProperties(
                        new NumberingLevelReference { Val = marker.Level },
                        new NumberingId { Val = marker.NumId });
                }

                container.AppendChild(paragraph);
                break;
            }

            // Front matter derives from CodeBlock in Markdig, so it must match first.
            case Markdig.Extensions.Yaml.YamlFrontMatterBlock:
                break; // metadata, not content

            case Md.CodeBlock code: // fenced and indented both derive from CodeBlock
                container.AppendChild(BuildCodeParagraph(state.Doc, code, quoteDepth));
                break;

            case Md.ListBlock listBlock:
                AppendMdList(container, listBlock, state, quoteDepth, level: list?.Level + 1 ?? 0, numIds: null);
                break;

            case Md.QuoteBlock quote:
                foreach (var child in quote)
                {
                    AppendMdBlock(container, child, state, quoteDepth + 1, list: null);
                }

                break;

            case MdTable.Table table:
                container.AppendChild(BuildMdTable(table, state));
                break;

            case Md.ThematicBreakBlock:
                container.AppendChild(HorizontalRuleParagraph());
                break;

            case Md.HtmlBlock html:
                state.Warnings.Add(new Warning(
                    "md_html_skipped",
                    $"Raw HTML block at line {html.Line + 1} was skipped; markdown import emits OOXML only."));
                break;

            case Md.LinkReferenceDefinitionGroup:
            case Md.LinkReferenceDefinition:
                break; // metadata, not content

            default:
                state.Warnings.Add(new Warning(
                    "md_block_skipped",
                    $"Unsupported markdown construct '{block.GetType().Name}' at line {block.Line + 1} was skipped."));
                break;
        }
    }

    /// <summary>
    /// Lists ride the M3 numbering machinery: one numbering instance per kind
    /// per top-level list (so a fresh list restarts at 1, and deeper levels
    /// reset whenever a shallower item appears — Word's own multi-level model).
    /// </summary>
    private static void AppendMdList(
        OpenXmlElement container,
        Md.ListBlock listBlock,
        MdImportState state,
        int quoteDepth,
        int level,
        Dictionary<string, int>? numIds)
    {
        numIds ??= [];
        var kind = listBlock.IsOrdered ? "number" : "bullet";
        if (!numIds.TryGetValue(kind, out var numId))
        {
            numId = NewNumberingInstance(state.Doc, kind);
            numIds[kind] = numId;
        }

        var effectiveLevel = Math.Min(level, MaxListLevel);
        foreach (var item in listBlock.OfType<Md.ListItemBlock>())
        {
            var firstParagraphDone = false;
            foreach (var child in item)
            {
                switch (child)
                {
                    case Md.ParagraphBlock paragraph when !firstParagraphDone:
                        firstParagraphDone = true;
                        AppendMdBlock(container, paragraph, state, quoteDepth, (numId, effectiveLevel));
                        break;

                    case Md.ListBlock nested:
                        AppendMdList(container, nested, state, quoteDepth, effectiveLevel + 1, numIds);
                        break;

                    default: // loose continuation paragraphs, code blocks, quotes, …
                        AppendMdBlock(container, child, state, quoteDepth, list: null);
                        break;
                }
            }
        }
    }

    /// <summary>
    /// A fenced/indented code block: one "Code"-styled paragraph (monospace,
    /// shaded), newlines preserved as w:br so the block round-trips losslessly.
    /// </summary>
    private static Paragraph BuildCodeParagraph(WordprocessingDocument doc, Md.CodeBlock code, int quoteDepth)
    {
        EnsureCodeStyle(doc);
        var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Code" }));

        var lines = code.Lines;
        for (var i = 0; i < lines.Count; i++)
        {
            if (i > 0)
            {
                paragraph.AppendChild(new Run(new Break()));
            }

            paragraph.AppendChild(new Run(NewText(lines.Lines[i].Slice.ToString())));
        }

        ApplyQuoteFormatting(paragraph, quoteDepth);
        return paragraph;
    }

    /// <summary>The monospace + shaded paragraph style code blocks import as (and export detects).</summary>
    private static void EnsureCodeStyle(WordprocessingDocument doc)
    {
        var styles = EnsureStylesRoot(doc);
        if (FindStyle(styles, "Code") is not null)
        {
            return;
        }

        styles.AppendChild(new Style(
            new StyleName { Val = "Code" },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "F2F2F2" }),
            new StyleRunProperties(
                new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" },
                new FontSize { Val = "20" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = "Code",
        });
    }

    /// <summary>Markdown blockquotes: left border + indent, one step per nesting level.</summary>
    private static void ApplyQuoteFormatting(Paragraph paragraph, int quoteDepth)
    {
        if (quoteDepth <= 0)
        {
            return;
        }

        var pPr = paragraph.ParagraphProperties ??= new ParagraphProperties();
        pPr.ParagraphBorders ??= new ParagraphBorders();
        pPr.ParagraphBorders.LeftBorder = new LeftBorder
        {
            Val = BorderValues.Single,
            Size = 12,
            Space = 4,
            Color = "CCCCCC",
        };
        pPr.Indentation = new Indentation
        {
            Left = (quoteDepth * 360).ToString(System.Globalization.CultureInfo.InvariantCulture),
        };
    }

    /// <summary>A horizontal rule, the way Word draws one: an empty paragraph with a bottom border.</summary>
    private static Paragraph HorizontalRuleParagraph() => new(new ParagraphProperties(
        new ParagraphBorders(new BottomBorder
        {
            Val = BorderValues.Single,
            Size = 6,
            Space = 1,
            Color = "auto",
        })));

    /// <summary>GFM pipe tables become real Word tables: grid, borders, bold header row, column alignments.</summary>
    private static Table BuildMdTable(MdTable.Table mdTable, MdImportState state)
    {
        var rows = mdTable.OfType<MdTable.TableRow>().ToList();
        var columns = Math.Max(
            mdTable.ColumnDefinitions?.Count ?? 0,
            rows.Select(r => r.Count).DefaultIfEmpty(0).Max());

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

        foreach (var mdRow in rows)
        {
            var row = new TableRow();
            if (mdRow.IsHeader)
            {
                row.AppendChild(new TableRowProperties(new TableHeader()));
            }

            for (var c = 0; c < columns; c++)
            {
                var cell = new TableCell();
                if (c < mdRow.Count && mdRow[c] is MdTable.TableCell mdCell)
                {
                    foreach (var block in mdCell)
                    {
                        AppendMdBlock(cell, block, state, quoteDepth: 0, list: null);
                    }
                }

                if (!cell.ChildElements.OfType<Paragraph>().Any())
                {
                    cell.AppendChild(new Paragraph()); // a tc must hold at least one block
                }

                var alignment = c < (mdTable.ColumnDefinitions?.Count ?? 0)
                    ? mdTable.ColumnDefinitions![c].Alignment
                    : null;
                foreach (var paragraph in cell.ChildElements.OfType<Paragraph>())
                {
                    if (alignment is MdTable.TableColumnAlign.Center)
                    {
                        WordFormatting.SetParagraphProp(paragraph, "alignment", "center");
                    }
                    else if (alignment is MdTable.TableColumnAlign.Right)
                    {
                        WordFormatting.SetParagraphProp(paragraph, "alignment", "right");
                    }
                    else if (alignment is MdTable.TableColumnAlign.Left)
                    {
                        WordFormatting.SetParagraphProp(paragraph, "alignment", "left");
                    }

                    if (mdRow.IsHeader)
                    {
                        if (!paragraph.ChildElements.OfType<Run>().Any())
                        {
                            paragraph.AppendChild(new Run(NewText(string.Empty)));
                        }

                        foreach (var run in paragraph.ChildElements.OfType<Run>())
                        {
                            WordFormatting.SetRunProp(run, "bold", "true");
                        }
                    }
                }

                row.AppendChild(cell);
            }

            table.AppendChild(row);
        }

        return table;
    }

    // ---------------------------------------------------------------- inlines

    private static void AppendMdInlines(
        OpenXmlElement target,
        MdInline.ContainerInline? container,
        MdImportState state,
        MdFormat format)
    {
        for (var inline = container?.FirstChild; inline is not null; inline = inline.NextSibling)
        {
            switch (inline)
            {
                case MdInline.LiteralInline literal:
                    target.AppendChild(MdRun(literal.Content.ToString(), format));
                    break;

                case MdInline.CodeInline code:
                    target.AppendChild(MdRun(code.Content, format with { Code = true }));
                    break;

                case MdInline.EmphasisInline emphasis:
                {
                    var inner = emphasis.DelimiterChar switch
                    {
                        '~' when emphasis.DelimiterCount >= 2 => format with { Strike = true },
                        '*' or '_' when emphasis.DelimiterCount >= 2 => format with { Bold = true },
                        '*' or '_' => format with { Italic = true },
                        _ => format,
                    };
                    AppendMdInlines(target, emphasis, state, inner);
                    break;
                }

                case MdInline.LinkInline { IsImage: true } image:
                    AppendMdImage(target, image, state);
                    break;

                case MdInline.LinkInline link:
                    AppendMdLink(target, link, state, format);
                    break;

                case MdInline.AutolinkInline autolink:
                    AppendMdAutolink(target, autolink, state, format);
                    break;

                case MdInline.LineBreakInline lineBreak:
                    if (lineBreak.IsHard)
                    {
                        target.AppendChild(new Run(new Break()));
                    }
                    else
                    {
                        target.AppendChild(MdRun(" ", format)); // soft wrap is just whitespace
                    }

                    break;

                case MdInline.HtmlInline html:
                    state.Warnings.Add(new Warning(
                        "md_html_skipped",
                        $"Inline HTML '{Snippet(html.Tag ?? string.Empty)}' was skipped; markdown import emits OOXML only."));
                    break;

                case MdInline.HtmlEntityInline entity:
                    target.AppendChild(MdRun(entity.Transcoded.ToString(), format));
                    break;

                case MdInline.ContainerInline nested: // unknown containers: keep their text
                    AppendMdInlines(target, nested, state, format);
                    break;

                default:
                    break;
            }
        }
    }

    private static Run MdRun(string text, MdFormat format)
    {
        var run = new Run();
        if (format != default)
        {
            var rPr = new RunProperties();
            if (format.Link)
            {
                rPr.RunStyle = new RunStyle { Val = "Hyperlink" };
            }

            if (format.Code)
            {
                rPr.RunFonts = new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" };
                rPr.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "EDEDED" };
            }

            if (format.Bold)
            {
                rPr.Bold = new Bold();
            }

            if (format.Italic)
            {
                rPr.Italic = new Italic();
            }

            if (format.Strike)
            {
                rPr.Strike = new Strike();
            }

            run.RunProperties = rPr;
        }

        run.AppendChild(NewText(text));
        return run;
    }

    private static void AppendMdLink(OpenXmlElement target, MdInline.LinkInline link, MdImportState state, MdFormat format)
    {
        if (!TryBuildHyperlink(state, link.Url, out var hyperlink))
        {
            state.Warnings.Add(new Warning(
                "md_link_skipped",
                $"Link target '{Snippet(link.Url ?? string.Empty)}' is not an absolute http(s)/mailto url; the text was kept, the link dropped."));
            AppendMdInlines(target, link, state, format);
            return;
        }

        AppendMdInlines(hyperlink, link, state, format with { Link = true });
        if (!hyperlink.ChildElements.OfType<Run>().Any())
        {
            hyperlink.AppendChild(MdRun(link.Url ?? string.Empty, format with { Link = true }));
        }

        target.AppendChild(hyperlink);
    }

    private static void AppendMdAutolink(OpenXmlElement target, MdInline.AutolinkInline autolink, MdImportState state, MdFormat format)
    {
        var url = autolink.IsEmail ? "mailto:" + autolink.Url : autolink.Url;
        if (!TryBuildHyperlink(state, url, out var hyperlink))
        {
            target.AppendChild(MdRun(autolink.Url, format));
            return;
        }

        hyperlink.AppendChild(MdRun(autolink.Url, format with { Link = true }));
        target.AppendChild(hyperlink);
    }

    /// <summary>Reuses the M2 hyperlink machinery: relationship + Hyperlink style.</summary>
    private static bool TryBuildHyperlink(MdImportState state, string? url, out Hyperlink hyperlink)
    {
        hyperlink = new Hyperlink();
        if (url is null ||
            !Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
            !AllowedLinkSchemes.Contains(uri.Scheme, StringComparer.OrdinalIgnoreCase))
        {
            return false;
        }

        EnsureHyperlinkStyle(state.Doc);
        hyperlink.Id = state.Doc.MainDocumentPart!.AddHyperlinkRelationship(uri, isExternal: true).Id;
        return true;
    }

    /// <summary>
    /// <c>![alt](path)</c>: the path resolves through the sandbox (relative to
    /// the markdown file first, then the workspace root) and embeds via the M2
    /// image machinery at natural size. Anything that fails — missing file,
    /// sandbox escape, unsupported format — is a per-image
    /// <c>md_image_skipped</c> warning, not a failed import.
    /// </summary>
    private static void AppendMdImage(OpenXmlElement target, MdInline.LinkInline image, MdImportState state)
    {
        var src = image.Url;
        var alt = (image.FirstChild as MdInline.LiteralInline)?.Content.ToString() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(src))
        {
            state.Warnings.Add(new Warning("md_image_skipped", "An image had no path and was skipped."));
            return;
        }

        try
        {
            var resolved = ResolveMdImagePath(state, src);
            var bytes = File.ReadAllBytes(resolved);
            var (imageFormat, pixelWidth, pixelHeight) = SniffImage(bytes, src);

            var main = state.Doc.MainDocumentPart!;
            var imagePart = main.AddImagePart(imageFormat == "png" ? ImagePartType.Png : ImagePartType.Jpeg);
            using (var stream = new MemoryStream(bytes, writable: false))
            {
                imagePart.FeedData(stream);
            }

            var cx = (long)Math.Round(Math.Max(1, pixelWidth) * EmuPerPixel);
            var cy = (long)Math.Round(Math.Max(1, pixelHeight) * EmuPerPixel);
            var name = alt is { Length: > 0 } ? alt : Path.GetFileName(src);
            target.AppendChild(new Run(BuildInlineDrawing(
                main.GetIdOfPart(imagePart),
                NextDrawingId(state.Doc),
                name,
                cx,
                cy)));
        }
        catch (AiofficeException ex)
        {
            state.Warnings.Add(new Warning(
                "md_image_skipped",
                $"Image '{Snippet(src)}' was skipped: {ex.Message}"));
        }
    }

    /// <summary>Sandbox-resolved image path, preferring md-file-relative over workspace-relative.</summary>
    private static string ResolveMdImagePath(MdImportState state, string src)
    {
        if (!Path.IsPathRooted(src) && state.SourceDir is { Length: > 0 })
        {
            try
            {
                return state.Workspace.Resolve(Path.Combine(state.SourceDir, src), mustExist: true);
            }
            catch (AiofficeException)
            {
                // fall through to workspace-root resolution
            }
        }

        return state.Workspace.Resolve(src, mustExist: true);
    }
}
