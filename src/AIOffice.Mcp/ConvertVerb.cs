using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;

namespace AIOffice.Mcp;

/// <summary>
/// The M9 <c>convert</c> verb (and <c>office_convert</c> tool): the one
/// entrypoint that moves a document between formats. Routing is by the
/// (source extension, destination extension) pair:
/// <list type="bullet">
/// <item>Same-family text bridges reuse the existing code paths — <c>docx↔md</c>
/// (the markdown bridge) and <c>xlsx↔csv</c> (the csv bridge).</item>
/// <item>Cross-format pairs go through the format-neutral model
/// (<see cref="INeutralConvertible"/>): the source handler projects to a
/// <see cref="NeutralDoc"/>, the destination handler absorbs it into a freshly
/// created file. Markdown participates via <see cref="NeutralMarkdown"/>, so any
/// office format can convert to/from <c>.md</c>.</item>
/// <item><c>any → pdf|png|svg|html|text</c> routes to the render layer (open the
/// source, render to the destination file).</item>
/// </list>
/// Conversion is content-transfer and inherently lossy across formats: every
/// drop (export-side and import-side) is aggregated into a single
/// <c>convert_lossy</c> meta warning so the result never hides what did not
/// survive. Both paths are sandbox-resolved; the destination is created fresh
/// (overwriting an existing destination warns <c>convert_overwrite</c>).
/// </summary>
public static class ConvertVerb
{
    /// <summary>The office formats convert can read from and write to.</summary>
    private static readonly string[] OfficeExtensions = [".docx", ".xlsx", ".pptx"];

    /// <summary>The render targets convert can produce (source opened, rendered to dest).</summary>
    private static readonly string[] RenderExtensions = [".pdf", ".png", ".svg", ".html", ".txt"];

    /// <summary>The text-bridge format the neutral model serializes to/from.</summary>
    private static readonly string[] MarkdownExtensions = [".md", ".markdown"];

    private static readonly string[] CsvExtensions = [".csv", ".tsv"];

    /// <summary>All extensions convert understands as a destination, for the unsupported-target hint.</summary>
    private static readonly string[] AllTargets =
        [.. OfficeExtensions, .. MarkdownExtensions, .. CsvExtensions, .. RenderExtensions];

    /// <summary>The content-source extensions convert can read (office + the text bridges).</summary>
    public static IReadOnlyList<string> SupportedSources =>
        [.. OfficeExtensions, .. MarkdownExtensions, .. CsvExtensions];

    /// <summary>The content-destination extensions convert can write a new file in (office + text bridges).</summary>
    public static IReadOnlyList<string> SupportedContentTargets =>
        [.. OfficeExtensions, .. MarkdownExtensions, .. CsvExtensions];

    /// <summary>The render-destination extensions convert routes to the render layer.</summary>
    public static IReadOnlyList<string> SupportedRenderTargets => RenderExtensions;

    /// <summary>
    /// Runs a conversion from <paramref name="srcArg"/> to <paramref name="destArg"/>
    /// (both user-supplied, sandbox-resolved here). <paramref name="byKind"/> maps a
    /// <see cref="DocumentKind"/> to its handler (for the neutral bridge);
    /// <paramref name="registry"/> resolves a path's handler for the text bridges.
    /// </summary>
    public static Envelope Run(
        Workspace workspace,
        HandlerRegistry registry,
        IReadOnlyDictionary<DocumentKind, IFormatHandler> byKind,
        string srcArg,
        string destArg)
    {
        if (string.IsNullOrWhiteSpace(srcArg) || string.IsNullOrWhiteSpace(destArg))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "convert needs a source and a destination: convert <src> <dest>.",
                "Example: aioffice convert report.docx deck.pptx — or report.docx report.md.");
        }

        var srcExt = Path.GetExtension(srcArg).ToLowerInvariant();
        var destExt = Path.GetExtension(destArg).ToLowerInvariant();

        if (srcExt.Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Cannot infer the source format: '{srcArg}' has no extension.",
                "Name the source with its real extension, e.g. report.docx.");
        }

        if (destExt.Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Cannot infer the destination format: '{destArg}' has no extension.",
                $"Name the destination with a target extension: {string.Join(", ", AllTargets)}.");
        }

        // Same extension is an edit, not a convert.
        if (srcExt == destExt || (MarkdownExtensions.Contains(srcExt) && MarkdownExtensions.Contains(destExt)))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "source and destination are already the same format; use edit.",
                "convert moves content BETWEEN formats. To change one file, use 'aioffice edit'.");
        }

        // The source must exist and is read-only; resolve it inside the sandbox.
        var src = workspace.Resolve(srcArg, mustExist: true);

        // Render targets: open the source, render it to the destination file.
        if (RenderExtensions.Contains(destExt))
        {
            return Render(workspace, registry, src, srcArg, destArg, destExt);
        }

        if (!IsConvertibleSource(srcExt))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"convert cannot read '{srcExt}' as a content source.",
                $"Convert FROM one of: {string.Join(", ", OfficeExtensions)}, {string.Join(", ", MarkdownExtensions)}, {string.Join(", ", CsvExtensions)}. " +
                $"To produce {srcExt}, convert INTO it from an office document.",
                candidates: [.. OfficeExtensions, .. MarkdownExtensions, .. CsvExtensions]);
        }

        if (!IsConvertibleDest(destExt))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"convert does not know how to write '{destExt}'.",
                $"Supported destinations: office {string.Join("/", OfficeExtensions)}, " +
                $"text {string.Join("/", MarkdownExtensions)}/{string.Join("/", CsvExtensions)}, " +
                $"render {string.Join("/", RenderExtensions)}.",
                candidates: AllTargets);
        }

        // The destination is created fresh; warn (don't refuse) if it already existed.
        var dest = workspace.Resolve(destArg);
        var overwrote = File.Exists(dest);
        if (Path.GetDirectoryName(dest) is { Length: > 0 } parent)
        {
            Directory.CreateDirectory(parent);
        }

        // csv is xlsx-only family; md bridges every office format via the neutral model.
        if (CsvExtensions.Contains(srcExt) || CsvExtensions.Contains(destExt))
        {
            return CsvBridge(workspace, registry, byKind, src, srcExt, srcArg, dest, destExt, destArg, overwrote);
        }

        // docx↔md reuses the existing markdown bridge verbatim; every other md pair
        // (pptx↔md, xlsx↔md) and every office↔office pair goes through the neutral model.
        var fromKind = OfficeKind(srcExt);
        var toKind = OfficeKind(destExt);
        var fromMd = MarkdownExtensions.Contains(srcExt);
        var toMd = MarkdownExtensions.Contains(destExt);

        if ((fromKind == DocumentKind.Docx && toMd) || (fromMd && toKind == DocumentKind.Docx))
        {
            return DocxMarkdownBridge(workspace, registry, src, srcExt, srcArg, dest, destExt, destArg, overwrote, fromMd);
        }

        return NeutralBridge(workspace, byKind, src, srcExt, srcArg, dest, destExt, destArg, overwrote);
    }

    // ─────────────────────────────────────────────────────── render route

    private static Envelope Render(
        Workspace workspace, HandlerRegistry registry, string src, string srcArg, string destArg, string destExt)
    {
        if (OfficeKind(Path.GetExtension(src).ToLowerInvariant()) is null)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Rendering to '{destExt}' needs an office source (.docx/.xlsx/.pptx), not '{Path.GetExtension(src)}'.",
                "Convert the source to an office format first, then render it.",
                candidates: OfficeExtensions);
        }

        var handler = registry.Resolve(src);
        var dest = workspace.Resolve(destArg);
        var overwrote = File.Exists(dest);
        if (Path.GetDirectoryName(dest) is { Length: > 0 } parent)
        {
            Directory.CreateDirectory(parent);
        }

        var to = destExt switch
        {
            ".pdf" => "pdf",
            ".png" => "png",
            ".svg" => "svg",
            ".html" => "html",
            _ => "text",
        };

        var ctx = new CommandContext
        {
            Workspace = workspace,
            File = src,
            Args = new JsonObject { ["to"] = to, ["output"] = dest },
        };

        var rendered = to switch
        {
            "png" => AIOffice.Render.PngRenderVerb.Execute(handler, ctx),
            "pdf" => AIOffice.Render.PdfRenderVerb.Execute(handler, ctx),
            _ => handler.Render(ctx),
        };

        if (!rendered.IsOk)
        {
            return rendered;
        }

        // text/svg/html handlers inline their content; persist it to the destination
        // file so 'convert' always yields a file on disk.
        if (to is "text" or "svg" or "html" && DataNode(rendered)?["content"]?.GetValue<string>() is { } content)
        {
            File.WriteAllText(dest, content);
        }

        var warnings = new List<Warning>();
        if (overwrote)
        {
            warnings.Add(OverwriteWarning(destArg));
        }

        return Envelope.Ok(
            new
            {
                from = Path.GetExtension(src).TrimStart('.').ToLowerInvariant(),
                to = destExt.TrimStart('.'),
                blocksWritten = 0,
                dropped = Array.Empty<string>(),
                written = dest,
            },
            new Meta { Warnings = warnings.Count > 0 ? warnings : null });
    }

    // ─────────────────────────────────────────────────────── docx ↔ md bridge

    private static Envelope DocxMarkdownBridge(
        Workspace workspace, HandlerRegistry registry,
        string src, string srcExt, string srcArg, string dest, string destExt, string destArg, bool overwrote, bool fromMd)
    {
        var dropped = new List<string>();

        if (fromMd)
        {
            // md → docx: the markdown bridge builds the docx from scratch.
            var handler = registry.Resolve(dest);
            if (File.Exists(dest))
            {
                File.Delete(dest); // CreateFrom refuses an existing target; convert may overwrite
            }

            var ctx = new CommandContext { Workspace = workspace, File = dest, Args = new JsonObject() };
            var env = Bridge.CreateFrom(handler, ctx, srcArg);
            if (!env.IsOk)
            {
                return env;
            }

            return Result(srcExt, destExt, blocksWritten: 0, dropped, dest, overwrote);
        }

        // docx → md: the markdown read view is the exported text.
        var docHandler = registry.Resolve(src);
        var readCtx = new CommandContext
        {
            Workspace = workspace,
            File = src,
            Args = new JsonObject { ["view"] = "markdown" },
        };
        var read = docHandler.Read(readCtx);
        if (!read.IsOk)
        {
            return read;
        }

        var markdown = DataNode(read)?["markdown"]?.GetValue<string>() ?? string.Empty;

        // Carry the document Title into the markdown as a leading "# Title" so the
        // converted file is self-describing and the title survives a later md → docx
        // round trip (the docx importer reads the first H1 as the title). The
        // general-purpose 'read --view markdown' view stays title-free; only the
        // convert bridge prepends it. Skip when the body already opens with that
        // exact heading, so a title that doubles as the first H1 is not duplicated.
        if (DocTitle(docHandler, workspace, src) is { Length: > 0 } title &&
            !MarkdownOpensWithTitle(markdown, title))
        {
            markdown = "# " + title + "\n\n" + markdown.TrimStart('\n');
        }

        File.WriteAllText(dest, markdown);
        return Result(srcExt, destExt, blocksWritten: 0, dropped, dest, overwrote);
    }

    /// <summary>The document's core Title (via <c>read --view properties</c>), or null when unset.</summary>
    private static string? DocTitle(IFormatHandler handler, Workspace workspace, string src)
    {
        var ctx = new CommandContext
        {
            Workspace = workspace,
            File = src,
            Args = new JsonObject { ["view"] = "properties" },
        };

        var read = handler.Read(ctx);
        if (!read.IsOk)
        {
            return null;
        }

        var title = DataNode(read)?["properties"]?["core"]?["title"]?.GetValue<string>();
        return string.IsNullOrWhiteSpace(title) ? null : title;
    }

    /// <summary>True when the markdown's first non-blank line is exactly "# {title}".</summary>
    private static bool MarkdownOpensWithTitle(string markdown, string title)
    {
        foreach (var line in markdown.Split('\n'))
        {
            var trimmed = line.Trim();
            if (trimmed.Length == 0)
            {
                continue;
            }

            // Skip a leading omission comment (read --view markdown emits one for
            // headers/footers/watermarks) so the H1 check sees the first real line.
            if (trimmed.StartsWith("<!--", StringComparison.Ordinal))
            {
                continue;
            }

            return string.Equals(trimmed, "# " + title, StringComparison.Ordinal);
        }

        return false;
    }

    // ─────────────────────────────────────────────────────── csv bridge

    private static Envelope CsvBridge(
        Workspace workspace, HandlerRegistry registry, IReadOnlyDictionary<DocumentKind, IFormatHandler> byKind,
        string src, string srcExt, string srcArg, string dest, string destExt, string destArg, bool overwrote)
    {
        // csv → xlsx is the only direct csv import; csv → other office formats go
        // csv → xlsx (neutral) is overkill, so route csv only to/from xlsx directly
        // and let the neutral bridge cover xlsx → (docx/pptx/md), etc.
        var fromCsv = CsvExtensions.Contains(srcExt);
        var toCsv = CsvExtensions.Contains(destExt);

        if (fromCsv && OfficeKind(destExt) == DocumentKind.Xlsx)
        {
            var handler = registry.Resolve(dest);
            if (File.Exists(dest))
            {
                File.Delete(dest);
            }

            var ctx = new CommandContext { Workspace = workspace, File = dest, Args = new JsonObject() };
            var env = Bridge.CreateFrom(handler, ctx, srcArg);
            return env.IsOk ? Result(srcExt, destExt, 0, [], dest, overwrote) : env;
        }

        if (toCsv && OfficeKind(srcExt) == DocumentKind.Xlsx)
        {
            var handler = registry.Resolve(src);
            var readCtx = new CommandContext
            {
                Workspace = workspace,
                File = src,
                Args = new JsonObject { ["view"] = "csv" },
            };
            var read = handler.Read(readCtx);
            if (!read.IsOk)
            {
                return read;
            }

            var content = DataNode(read)?["content"]?.GetValue<string>() ?? string.Empty;
            File.WriteAllText(dest, content);
            var dropped = new List<string>();
            if (DataNode(read)?["truncated"]?.GetValue<bool>() == true)
            {
                dropped.Add("the workbook exceeded one csv window; only the first sheet's used range was exported");
            }

            return Result(srcExt, destExt, 0, dropped, dest, overwrote);
        }

        // csv ↔ a non-xlsx office format: there is no single-file csv home for it.
        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"convert moves csv only to/from xlsx, not '{(fromCsv ? destExt : srcExt)}'.",
            "Convert via a workbook: csv → xlsx → docx/pptx, or docx/pptx → xlsx → csv.",
            candidates: [".xlsx"]);
    }

    // ─────────────────────────────────────────────────────── neutral bridge

    private static Envelope NeutralBridge(
        Workspace workspace, IReadOnlyDictionary<DocumentKind, IFormatHandler> byKind,
        string src, string srcExt, string srcArg, string dest, string destExt, string destArg, bool overwrote)
    {
        var dropped = new List<string>();

        // 1. EXPORT the source to the neutral model (md source = parse text).
        NeutralDoc neutral;
        if (MarkdownExtensions.Contains(srcExt))
        {
            FileSizeGuard.Ensure(src);
            neutral = NeutralMarkdown.Parse(File.ReadAllText(src));
        }
        else
        {
            var fromKind = OfficeKind(srcExt)!.Value;
            var fromHandler = RequireConvertible(byKind, fromKind, "read from");
            var exportCtx = new CommandContext { Workspace = workspace, File = src, Args = new JsonObject() };

            // The pptx exporter has an overload that names its export-side losses
            // (animations/charts/SmartArt) — capture them when present.
            if (fromHandler is AIOffice.Pptx.PptxHandler pptx)
            {
                neutral = pptx.ExportNeutral(exportCtx, out var pptxDropped);
                dropped.AddRange(pptxDropped);
            }
            else
            {
                neutral = ((INeutralConvertible)fromHandler).ExportNeutral(exportCtx);
            }
        }

        var blocksWritten = neutral.Blocks.Count;

        // 2. IMPORT the neutral model into a freshly created destination file.
        if (MarkdownExtensions.Contains(destExt))
        {
            File.WriteAllText(dest, NeutralMarkdown.Write(neutral));
            // Markdown carries headings/lists/tables/links/images but not colors
            // or exact styling; name that once if the source had any.
            if (HasColorOrUnderline(neutral))
            {
                dropped.Add("run colors and underline are dropped: markdown carries bold/italic/links only");
            }
        }
        else
        {
            var toKind = OfficeKind(destExt)!.Value;
            var toHandler = RequireConvertible(byKind, toKind, "write to");
            var convertible = (INeutralConvertible)toHandler;

            // Create the destination fresh, then import into it.
            if (File.Exists(dest))
            {
                File.Delete(dest);
            }

            var createCtx = new CommandContext { Workspace = workspace, File = dest, Args = new JsonObject() };
            var create = toHandler.Create(createCtx);
            if (!create.IsOk)
            {
                return create;
            }

            var importCtx = new CommandContext { Workspace = workspace, File = dest, Args = new JsonObject() };
            var result = convertible.ImportNeutral(importCtx, neutral);
            blocksWritten = result.BlocksWritten;
            dropped.AddRange(result.Dropped);
        }

        return Result(srcExt, destExt, blocksWritten, dropped, dest, overwrote);
    }

    // ─────────────────────────────────────────────────────── helpers

    private static Envelope Result(
        string srcExt, string destExt, int blocksWritten, List<string> dropped, string dest, bool overwrote)
    {
        var distinct = dropped.Distinct(StringComparer.Ordinal).ToList();
        var warnings = new List<Warning>();
        if (distinct.Count > 0)
        {
            warnings.Add(new Warning(
                "convert_lossy",
                "Some content did not survive the conversion: " + string.Join("; ", distinct) + "."));
        }

        if (overwrote)
        {
            warnings.Add(OverwriteWarning(dest));
        }

        return Envelope.Ok(
            new
            {
                from = srcExt.TrimStart('.'),
                to = destExt.TrimStart('.'),
                blocksWritten,
                dropped = distinct,
                written = dest,
            },
            new Meta { Warnings = warnings.Count > 0 ? warnings : null });
    }

    private static Warning OverwriteWarning(string dest) =>
        new("convert_overwrite", $"The destination '{dest}' already existed and was overwritten.");

    private static IFormatHandler RequireConvertible(
        IReadOnlyDictionary<DocumentKind, IFormatHandler> byKind, DocumentKind kind, string direction)
    {
        if (byKind.TryGetValue(kind, out var handler) && handler is INeutralConvertible)
        {
            return handler;
        }

        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"The {kind.ToString().ToLowerInvariant()} handler cannot {direction} the neutral model in this build.",
            "Run 'aioffice doctor' for handler status; the docx/xlsx/pptx handlers all convert.");
    }

    private static bool HasColorOrUnderline(NeutralDoc doc) =>
        doc.Blocks.Any(b => b.Runs is { } runs && runs.Any(r => r.Color is not null || r.Underline));

    private static bool IsConvertibleSource(string ext) =>
        OfficeExtensions.Contains(ext) || MarkdownExtensions.Contains(ext) || CsvExtensions.Contains(ext);

    private static bool IsConvertibleDest(string ext) =>
        OfficeExtensions.Contains(ext) || MarkdownExtensions.Contains(ext) || CsvExtensions.Contains(ext);

    private static DocumentKind? OfficeKind(string ext) => ext switch
    {
        ".docx" => DocumentKind.Docx,
        ".xlsx" => DocumentKind.Xlsx,
        ".pptx" => DocumentKind.Pptx,
        _ => null,
    };

    private static JsonObject? DataNode(Envelope envelope) =>
        JsonSerializer.SerializeToNode(envelope.Data, JsonDefaults.Options) as JsonObject;
}
