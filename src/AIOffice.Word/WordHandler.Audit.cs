using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Dw = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// The docx accessibility + quality auditor (M7). Each check fires on a bad
/// document and stays silent on a clean one; the safe subset is autofixable.
/// Findings carry a stable id (<c>code#path</c>) so <c>--fix</c> can target one.
/// </summary>
public sealed partial class WordHandler : IAuditor
{
    public AuditResult Audit(CommandContext ctx, AuditOptions opts)
    {
        var file = RequireFile(ctx, mustExist: true);
        var (doc, ms, _) = OpenCopy(file, editable: false);
        using (doc)
        using (ms)
        {
            var findings = CollectFindings(doc).ToList();
            return FilterAndSummarize(findings, opts);
        }
    }

    public int Fix(CommandContext ctx, IReadOnlyList<string> findingIds)
    {
        var file = RequireFile(ctx, mustExist: true);

        // Atomic, like Edit: apply to an in-memory copy and write back only on success.
        var originalBytes = File.ReadAllBytes(file);

        // a11y_no_doc_title writes the core Title; migrate any legacy .psmdcp part to
        // the standard docProps/core.xml first so the fix lands in the conventional part.
        var workingBytes = MigrateLegacyCorePropertiesBytes(originalBytes, file);

        var ms = new MemoryStream();
        ms.Write(workingBytes);
        ms.Position = 0;

        var fixedCount = 0;
        using (var doc = OpenPackage(ms, file, editable: true))
        {
            // Re-collect against the live doc so ids line up with the current state.
            var byId = CollectFindings(doc).ToDictionary(f => f.Id, StringComparer.Ordinal);

            // Empty id list = fix every autofixable finding (the default --fix).
            var targets = findingIds.Count == 0
                ? byId.Values.Where(f => f.Autofixable).Select(f => f.Id).ToList()
                : findingIds;

            foreach (var id in targets)
            {
                if (byId.TryGetValue(id, out var finding) && finding.Autofixable && ApplyAutofix(doc, file, finding))
                {
                    fixedCount++;
                }
            }
        }

        if (fixedCount > 0)
        {
            _snapshots.Save(file); // pre-image, so the fix is undoable
            File.WriteAllBytes(file, ms.ToArray());
        }

        return fixedCount;
    }

    /// <summary>Applies severity/category filtering and tallies the summary.</summary>
    private static AuditResult FilterAndSummarize(List<AuditFinding> findings, AuditOptions opts)
    {
        var minRank = AuditOptions.SeverityRank(opts.MinSeverity);
        var filtered = findings
            .Where(f => opts.Category is "all" || f.Category == opts.Category)
            .Where(f => AuditOptions.SeverityRank(f.Severity) >= minRank)
            .ToList();

        return new AuditResult
        {
            Findings = filtered,
            Summary = new AuditSummary(
                Errors: filtered.Count(f => f.Severity == "error"),
                Warnings: filtered.Count(f => f.Severity == "warning"),
                Infos: filtered.Count(f => f.Severity == "info")),
        };
    }

    // -------------------------------------------------------------- collect

    private IEnumerable<AuditFinding> CollectFindings(WordprocessingDocument doc)
    {
        foreach (var f in AuditNoDocTitle(doc))
        {
            yield return f;
        }

        foreach (var f in AuditImages(doc))
        {
            yield return f;
        }

        foreach (var f in AuditHeadings(doc))
        {
            yield return f;
        }

        foreach (var f in AuditTables(doc))
        {
            yield return f;
        }

        foreach (var f in AuditContrast(doc))
        {
            yield return f;
        }

        foreach (var f in AuditLinks(doc))
        {
            yield return f;
        }

        foreach (var f in AuditOrphanBookmarks(doc))
        {
            yield return f;
        }
    }

    // ------------------------------------------------------- a11y_no_doc_title

    private static IEnumerable<AuditFinding> AuditNoDocTitle(WordprocessingDocument doc)
    {
        if (string.IsNullOrWhiteSpace(ReadCoreTitle(doc)))
        {
            yield return Finding(
                "a11y_no_doc_title", "warning", "accessibility", path: "/properties",
                "The document has no title (core property Title is empty).",
                "Set a title: {\"op\":\"set\",\"path\":\"/properties\",\"props\":{\"title\":\"…\"}}. " +
                "--fix uses the first Heading1 or the file name.",
                autofixable: true);
        }
    }

    // ------------------------------------------------------- a11y_no_alt_text

    /// <summary>Inline/floating images whose wp:docPr has no description (alt text).</summary>
    private static IEnumerable<AuditFinding> AuditImages(WordprocessingDocument doc)
    {
        foreach (var (element, path) in DrawingCarriers(doc))
        {
            var docPr = element.Descendants<Dw.DocProperties>().FirstOrDefault();
            if (docPr is null)
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(docPr.Description?.Value))
            {
                yield return Finding(
                    "a11y_no_alt_text", "error", "accessibility", path,
                    "An image has no alternative text (descr).",
                    "Add a description: {\"op\":\"set\",\"path\":\"" + path + "\",\"props\":{\"alt\":\"…\"}} — " +
                    "or --fix to insert a placeholder.",
                    autofixable: true);
            }
        }
    }

    /// <summary>
    /// One carrier per drawing, addressed at its most specific node (the run when
    /// the drawing sits in one, else the paragraph). Deduped by docPr so a single
    /// image is never reported twice (once for its run and once for its paragraph).
    /// </summary>
    private static IEnumerable<(OpenXmlElement Element, string Path)> DrawingCarriers(WordprocessingDocument doc)
    {
        var seen = new HashSet<Dw.DocProperties>();
        foreach (var node in WordAddress.EnumerateAll(doc))
        {
            if (node.Type is not ("p" or "run"))
            {
                continue;
            }

            var docPr = node.Element.Descendants<Dw.DocProperties>().FirstOrDefault();
            if (docPr is null || !seen.Add(docPr))
            {
                continue;
            }

            yield return (node.Element, node.CanonicalPath);
        }
    }

    // ------------------------------------------------ heading_skip / empty

    private static IEnumerable<AuditFinding> AuditHeadings(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart?.Document?.Body is not { } body)
        {
            yield break;
        }

        var previousLevel = 0;
        foreach (var node in WordAddress.EnumerateBody(body).Where(n => n.Type == "p"))
        {
            var paragraph = (Paragraph)node.Element;
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (HeadingLevel(styleId) is not { } level)
            {
                continue;
            }

            // quality_empty_heading: a heading paragraph with no visible text.
            if (string.IsNullOrWhiteSpace(paragraph.InnerText))
            {
                yield return Finding(
                    "quality_empty_heading", "warning", "quality", node.CanonicalPath,
                    $"Heading {level} at {node.CanonicalPath} has no text.",
                    "Give the heading text, or remove the empty heading paragraph.",
                    autofixable: false);
            }

            // a11y_heading_skip: jumping more than one level deeper (H1 -> H3).
            if (previousLevel > 0 && level > previousLevel + 1)
            {
                yield return Finding(
                    "a11y_heading_skip", "warning", "accessibility", node.CanonicalPath,
                    $"Heading level jumps from H{previousLevel} to H{level}, skipping H{previousLevel + 1}.",
                    $"Add an intermediate Heading{previousLevel + 1}, or demote this heading so levels increase by one.",
                    autofixable: false);
            }

            previousLevel = level;
        }
    }

    // ------------------------------------------------ a11y_no_table_header

    private static IEnumerable<AuditFinding> AuditTables(WordprocessingDocument doc)
    {
        foreach (var node in WordAddress.EnumerateAll(doc).Where(n => n.Type == "table"))
        {
            var table = (Table)node.Element;
            var firstRow = table.Elements<TableRow>().FirstOrDefault();
            if (firstRow is null)
            {
                continue; // an empty table is a different problem
            }

            var hasHeader = firstRow.TableRowProperties?.GetFirstChild<TableHeader>() is not null
                || firstRow.Descendants<TableHeader>().Any();
            if (!hasHeader)
            {
                yield return Finding(
                    "a11y_no_table_header", "error", "accessibility", node.CanonicalPath,
                    $"Table {node.CanonicalPath} has no header row marked to repeat.",
                    "Mark the first row a header: {\"op\":\"set\",\"path\":\"" + node.CanonicalPath +
                    "\",\"props\":{\"headerRow\":true}} — or --fix.",
                    autofixable: true);
            }
        }
    }

    // ------------------------------------------------------ a11y_low_contrast

    /// <summary>
    /// Run colour vs its background (cell/paragraph/table shading, default white)
    /// with a WCAG relative-luminance ratio below 4.5:1.
    /// </summary>
    private static IEnumerable<AuditFinding> AuditContrast(WordprocessingDocument doc)
    {
        foreach (var node in WordAddress.EnumerateAll(doc).Where(n => n.Type == "run"))
        {
            var run = (Run)node.Element;
            var fg = run.RunProperties?.Color?.Val?.Value;
            if (fg is null || string.IsNullOrWhiteSpace(run.InnerText))
            {
                continue;
            }

            if (!TryParseHexColor(fg, out var foreground))
            {
                continue; // "auto" or a theme colour — not a concrete hex we can score
            }

            var background = BackgroundColorOf(run);
            var ratio = ContrastRatio(foreground, background);
            if (ratio < 4.5)
            {
                yield return Finding(
                    "a11y_low_contrast", "warning", "accessibility", node.CanonicalPath,
                    $"Text colour #{fg.ToUpperInvariant()} on #{ColorHex(background)} has a contrast ratio of " +
                    $"{ratio.ToString("0.0", CultureInfo.InvariantCulture)}:1, below the 4.5:1 minimum.",
                    "Use a darker text colour or a lighter background so the ratio reaches 4.5:1 (WCAG AA).",
                    autofixable: false);
            }
        }
    }

    /// <summary>The effective background of a run: nearest shading (run/cell/paragraph), default white.</summary>
    private static (int R, int G, int B) BackgroundColorOf(Run run)
    {
        var runShade = run.RunProperties?.Shading?.Fill?.Value;
        if (runShade is not null && TryParseHexColor(runShade, out var rc))
        {
            return rc;
        }

        var cell = run.Ancestors<TableCell>().FirstOrDefault();
        var cellShade = cell?.TableCellProperties?.Shading?.Fill?.Value;
        if (cellShade is not null && TryParseHexColor(cellShade, out var cc))
        {
            return cc;
        }

        var paragraph = run.Ancestors<Paragraph>().FirstOrDefault();
        var paraShade = paragraph?.ParagraphProperties?.GetFirstChild<Shading>()?.Fill?.Value;
        if (paraShade is not null && TryParseHexColor(paraShade, out var pc))
        {
            return pc;
        }

        return (255, 255, 255); // default page background
    }

    // ----------------------------------------------------- quality_broken_link

    /// <summary>Internal hyperlinks (w:hyperlink/@w:anchor) pointing at a missing bookmark.</summary>
    private static IEnumerable<AuditFinding> AuditLinks(WordprocessingDocument doc)
    {
        var bookmarks = EnumerateBookmarks(doc)
            .Select(b => b.Name?.Value)
            .OfType<string>()
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var node in WordAddress.EnumerateAll(doc).Where(n => n.Type == "link"))
        {
            var link = (Hyperlink)node.Element;
            var anchor = link.Anchor?.Value;
            if (anchor is { Length: > 0 } && !bookmarks.Contains(anchor))
            {
                yield return Finding(
                    "quality_broken_link", "error", "quality", node.CanonicalPath,
                    $"Hyperlink at {node.CanonicalPath} targets bookmark '{anchor}', which does not exist.",
                    $"Add the bookmark, or repoint the link. Existing bookmarks: " +
                    (bookmarks.Count > 0 ? string.Join(", ", bookmarks) : "(none)") + ".",
                    autofixable: false);
            }
        }
    }

    // -------------------------------------------------- quality_orphan_bookmark

    /// <summary>Bookmarks referenced by no internal link (and not the implicit _GoBack).</summary>
    private static IEnumerable<AuditFinding> AuditOrphanBookmarks(WordprocessingDocument doc)
    {
        var referenced = WordAddress.EnumerateAll(doc)
            .Where(n => n.Type == "link")
            .Select(n => ((Hyperlink)n.Element).Anchor?.Value)
            .OfType<string>()
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        // Field-based references (REF/PAGEREF Bookmark) also count as usage.
        foreach (var instr in doc.MainDocumentPart?.Document?.Descendants<FieldCode>() ?? [])
        {
            var tokens = instr.Text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length >= 2 && tokens[0] is "REF" or "PAGEREF")
            {
                referenced.Add(tokens[1]);
            }
        }

        foreach (var instr in doc.MainDocumentPart?.Document?.Descendants<SimpleField>() ?? [])
        {
            var tokens = (instr.Instruction?.Value ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length >= 2 && tokens[0] is "REF" or "PAGEREF")
            {
                referenced.Add(tokens[1]);
            }
        }

        foreach (var start in EnumerateBookmarks(doc))
        {
            var name = start.Name?.Value;
            if (name is null || name.StartsWith('_'))
            {
                continue; // Word's implicit bookmarks (_GoBack, _Toc…) are not orphans
            }

            if (!referenced.Contains(name))
            {
                yield return Finding(
                    "quality_orphan_bookmark", "info", "quality", BookmarkPath(name),
                    $"Bookmark '{name}' is referenced by nothing.",
                    "Remove it if unused ({\"op\":\"remove\",\"path\":\"" + BookmarkPath(name) + "\"}), or --fix.",
                    autofixable: true);
            }
        }
    }

    // -------------------------------------------------------------- autofix

    /// <summary>Applies one safe autofix in place; returns whether it changed anything.</summary>
    private static bool ApplyAutofix(WordprocessingDocument doc, string file, AuditFinding finding)
    {
        switch (finding.Code)
        {
            case "a11y_no_alt_text":
            {
                if (finding.Path is null)
                {
                    return false;
                }

                var node = WordAddress.Resolve(doc, DocPath.Parse(finding.Path));
                var docPr = node.Element.Descendants<Dw.DocProperties>().FirstOrDefault();
                if (docPr is null || !string.IsNullOrWhiteSpace(docPr.Description?.Value))
                {
                    return false;
                }

                docPr.Description = "(describe this image)";
                return true;
            }

            case "a11y_no_table_header":
            {
                if (finding.Path is null)
                {
                    return false;
                }

                var node = WordAddress.Resolve(doc, DocPath.Parse(finding.Path));
                if (node.Element is not Table table)
                {
                    return false;
                }

                ApplyHeaderRow(table, node, on: true);
                return true;
            }

            case "a11y_no_doc_title":
            {
                var title = FirstHeadingText(doc) ?? Path.GetFileNameWithoutExtension(file);
                if (string.IsNullOrWhiteSpace(title))
                {
                    return false;
                }

                WriteCoreField(doc, CoreField.Title, title);
                return true;
            }

            case "quality_orphan_bookmark":
            {
                if (finding.Path is null)
                {
                    return false;
                }

                ApplyRemoveBookmark(doc, new EditOp { Op = "remove", Path = finding.Path });
                return true;
            }

            default:
                return false;
        }
    }

    private static string? FirstHeadingText(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart?.Document?.Body is not { } body)
        {
            return null;
        }

        foreach (var node in WordAddress.EnumerateBody(body).Where(n => n.Type == "p"))
        {
            var paragraph = (Paragraph)node.Element;
            if (HeadingLevel(paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value) == 1 &&
                !string.IsNullOrWhiteSpace(paragraph.InnerText))
            {
                return paragraph.InnerText;
            }
        }

        return null;
    }

    // ----------------------------------------------------- WCAG luminance math

    /// <summary>WCAG 2.x contrast ratio between two colours ((L1+0.05)/(L2+0.05)).</summary>
    internal static double ContrastRatio((int R, int G, int B) a, (int R, int G, int B) b)
    {
        var la = RelativeLuminance(a);
        var lb = RelativeLuminance(b);
        var (lighter, darker) = la >= lb ? (la, lb) : (lb, la);
        return (lighter + 0.05) / (darker + 0.05);
    }

    /// <summary>WCAG relative luminance of an sRGB colour (0..1).</summary>
    internal static double RelativeLuminance((int R, int G, int B) color)
    {
        return (0.2126 * Channel(color.R)) + (0.7152 * Channel(color.G)) + (0.0722 * Channel(color.B));

        static double Channel(int raw)
        {
            var c = raw / 255.0;
            return c <= 0.03928 ? c / 12.92 : Math.Pow((c + 0.055) / 1.055, 2.4);
        }
    }

    internal static bool TryParseHexColor(string? value, out (int R, int G, int B) color)
    {
        color = default;
        if (value is null)
        {
            return false;
        }

        var hex = value.TrimStart('#');
        if (hex.Length != 6 || !hex.All(Uri.IsHexDigit))
        {
            return false; // "auto", named/theme colours, malformed — not scoreable
        }

        color = (
            int.Parse(hex.AsSpan(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
            int.Parse(hex.AsSpan(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
            int.Parse(hex.AsSpan(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
        return true;
    }

    private static string ColorHex((int R, int G, int B) c) =>
        string.Create(CultureInfo.InvariantCulture, $"{c.R:X2}{c.G:X2}{c.B:X2}");

    // ----------------------------------------------------------------- helpers

    private static AuditFinding Finding(
        string code, string severity, string category, string? path,
        string message, string suggestion, bool autofixable) => new()
    {
        Id = path is null ? code : $"{code}#{path}",
        Code = code,
        Severity = severity,
        Category = category,
        Path = path,
        Message = message,
        Suggestion = suggestion,
        Autofixable = autofixable,
    };
}
