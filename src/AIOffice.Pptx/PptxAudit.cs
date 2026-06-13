using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// The pptx accessibility + quality auditor (<c>aioffice audit</c>). Each check
/// fires on a crafted bad deck and stays silent on a clean one; <c>--fix</c>
/// applies only the safe, never-destructive autofixes (placeholder alt text, a
/// present title, and removal of empty placeholders).
/// </summary>
public sealed partial class PptxHandler : IAuditor
{
    // Accessibility thresholds.
    private const double TinyFontWarnPt = 12.0; // < 12pt is hard to read from the back of a room
    private const double TinyFontErrorPt = 8.0;  // < 8pt is effectively unreadable
    private const double MinContrastRatio = 4.5; // WCAG AA for normal text
    private const string PlaceholderAltText = "(describe this image)";
    private const string PlaceholderTitleText = "(slide title)";

    public AuditResult Audit(CommandContext ctx, AuditOptions opts)
    {
        var file = ctx.File ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "audit requires a target file.",
            "Pass the .pptx path as the first argument.");

        using var stream = PptxDoc.LoadStream(file);
        using var doc = PptxDoc.Open(stream, editable: false, file);
        var presentation = PptxDoc.RequirePresentationPart(doc, file);

        var findings = Collect(presentation)
            .Where(f => opts.Category is "all" || string.Equals(f.Category, opts.Category, StringComparison.Ordinal))
            .Where(f => AuditOptions.SeverityRank(f.Severity) >= AuditOptions.SeverityRank(opts.MinSeverity))
            .ToList();

        return new AuditResult
        {
            Findings = findings,
            Summary = new AuditSummary(
                findings.Count(f => f.Severity == "error"),
                findings.Count(f => f.Severity == "warning"),
                findings.Count(f => f.Severity == "info")),
        };
    }

    public int Fix(CommandContext ctx, IReadOnlyList<string> findingIds)
    {
        var file = ctx.File ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "audit --fix requires a target file.",
            "Pass the .pptx path as the first argument.");

        // An empty id list means "fix every autofixable finding".
        var wanted = findingIds is { Count: > 0 }
            ? new HashSet<string>(findingIds, StringComparer.Ordinal)
            : null;

        int fixed_;
        using var stream = PptxDoc.LoadStream(file);
        using (var doc = PptxDoc.Open(stream, editable: true, file))
        {
            var presentation = PptxDoc.RequirePresentationPart(doc, file);

            // Re-derive the findings from the live tree so ids line up with Audit().
            var autofixable = Collect(presentation)
                .Where(f => f.Autofixable && (wanted is null || wanted.Contains(f.Id)))
                .ToList();

            fixed_ = ApplyFixes(presentation, autofixable);
        }

        if (fixed_ > 0)
        {
            _snapshots?.Save(file);
            File.WriteAllBytes(file, stream.ToArray());
        }

        return fixed_;
    }

    // ---- collection ---------------------------------------------------------

    /// <summary>Every finding, unfiltered and in deck order (slide 1 → N, doc order within a slide).</summary>
    private static List<AuditFinding> Collect(PresentationPart presentation)
    {
        var findings = new List<AuditFinding>();
        var bounds = SlideBoundsEmu(presentation);
        var slides = PptxDoc.Slides(presentation);
        for (var i = 0; i < slides.Count; i++)
        {
            var slideIndex = i + 1;
            var slidePart = slides[i].Part;
            AuditSlide(findings, slidePart, slideIndex, bounds);
        }

        return findings;
    }

    private static void AuditSlide(List<AuditFinding> findings, SlidePart slidePart, int slideIndex, (long Width, long Height) bounds)
    {
        var slidePath = Units.Inv($"/slide[{slideIndex}]");
        var shapes = PptxDoc.Shapes(slidePart);
        var slideBgHex = PptxDoc.BackgroundHex(slidePart);
        var (slideW, slideH) = bounds;

        CheckSlideTitle(findings, shapes, slidePath, slideIndex);
        CheckDuplicateIds(findings, shapes, slidePath);
        CheckReadingOrder(findings, shapes, slidePath);

        foreach (var shape in shapes)
        {
            var shapePath = shape.CanonicalPath(slideIndex);
            CheckAltText(findings, shape, shapePath);
            CheckEmptyPlaceholder(findings, shape, shapePath);
            CheckOffCanvas(findings, shape, shapePath, slideW, slideH);
            CheckFontSizes(findings, shape, shapePath);
            CheckContrast(findings, shape, shapePath, slideBgHex);
        }
    }

    // ---- accessibility checks ----------------------------------------------

    /// <summary>a11y_no_alt_text: pictures, charts, tables and groups need a description.</summary>
    private static void CheckAltText(List<AuditFinding> findings, ShapeView shape, string shapePath)
    {
        if (!NeedsAltText(shape))
        {
            return;
        }

        if (PptxDoc.AltText(shape.Element) is not null)
        {
            return;
        }

        findings.Add(new AuditFinding
        {
            Id = Id("a11y_no_alt_text", shapePath),
            Severity = "warning",
            Category = "accessibility",
            Code = "a11y_no_alt_text",
            Path = shapePath,
            Message = $"The {KindLabel(shape)} '{NameOrKind(shape)}' has no alt text; screen readers cannot describe it.",
            Suggestion = $"Set alt text: {{\"op\":\"set\",\"path\":\"{shapePath}\",\"props\":{{\"altText\":\"…\"}}}} — " +
                "or run 'aioffice audit <file> --fix' to insert a placeholder.",
            Autofixable = true,
        });
    }

    /// <summary>Visual, non-text shapes (pictures, charts, tables, groups) carry meaning that needs alt text.</summary>
    private static bool NeedsAltText(ShapeView shape) => shape.Element switch
    {
        P.Picture => true,
        P.GraphicFrame => true, // charts and tables
        P.GroupShape => true,
        _ => false,
    };

    /// <summary>a11y_no_slide_title: every slide needs a title placeholder with text for navigation.</summary>
    private static void CheckSlideTitle(List<AuditFinding> findings, List<ShapeView> shapes, string slidePath, int slideIndex)
    {
        var titleShape = shapes.FirstOrDefault(IsTitlePlaceholder);
        var hasTitleText = titleShape is not null && PptxDoc.ShapeText(titleShape.Element).Trim().Length > 0;
        if (hasTitleText)
        {
            return;
        }

        findings.Add(new AuditFinding
        {
            Id = Id("a11y_no_slide_title", slidePath),
            Severity = "warning",
            Category = "accessibility",
            Code = "a11y_no_slide_title",
            Path = slidePath,
            Message = titleShape is null
                ? $"Slide {slideIndex} has no title placeholder; screen-reader navigation relies on slide titles."
                : $"Slide {slideIndex} has an empty title placeholder.",
            Suggestion = $"Give it a title: {{\"op\":\"set\",\"path\":\"{slidePath}\",\"props\":{{\"title\":\"…\"}}}} — " +
                "or run 'aioffice audit <file> --fix' to insert a placeholder title.",
            Autofixable = true,
        });
    }

    /// <summary>
    /// a11y_reading_order (info): the narration order is document order, but the
    /// visual order is top-to-bottom then left-to-right. When they disagree the
    /// deck reads out of sequence for screen-reader users.
    /// </summary>
    private static void CheckReadingOrder(List<AuditFinding> findings, List<ShapeView> shapes, string slidePath)
    {
        var positioned = shapes
            .Select(s => (Shape: s, Geo: PptxDoc.Geometry(s.Element)))
            .Where(t => t.Geo is not null)
            .Select(t => (t.Shape, Geo: t.Geo!.Value))
            .ToList();
        if (positioned.Count < 2)
        {
            return;
        }

        // Visual order: rows top-to-bottom (with a tolerance), then left-to-right.
        const long rowToleranceEmu = 457_200; // ~0.5 inch: shapes within this band are "the same row"
        var visual = positioned
            .OrderBy(t => t.Geo.Y / Math.Max(rowToleranceEmu, 1))
            .ThenBy(t => t.Geo.X)
            .Select(t => t.Shape.Id)
            .ToList();
        var document = positioned.Select(t => t.Shape.Id).ToList();

        if (visual.SequenceEqual(document))
        {
            return;
        }

        findings.Add(new AuditFinding
        {
            Id = Id("a11y_reading_order", slidePath),
            Severity = "info",
            Category = "accessibility",
            Code = "a11y_reading_order",
            Path = slidePath,
            Message = "Shapes are narrated in document order but laid out in a different visual order; " +
                "the slide may read out of sequence.",
            Suggestion = "Reorder shapes to match the visual flow with " +
                "{\"op\":\"move\",\"path\":\"<shape>\",\"position\":\"readingOrder N\"} " +
                "(1 = narrated first); 'aioffice read <file> --view structure' shows the current order.",
            Autofixable = false,
        });
    }

    /// <summary>a11y_tiny_font: text runs below the legibility thresholds (< 12pt warn, < 8pt error).</summary>
    private static void CheckFontSizes(List<AuditFinding> findings, ShapeView shape, string shapePath)
    {
        if (shape.Element is not P.Shape { TextBody: { } body })
        {
            return;
        }

        double? smallestPt = null;
        foreach (var run in body.Descendants<A.Run>())
        {
            if (run.Text?.Text is not { Length: > 0 })
            {
                continue;
            }

            if (run.RunProperties?.FontSize?.Value is { } sizeHundredths)
            {
                var pt = sizeHundredths / 100.0;
                smallestPt = smallestPt is { } s ? Math.Min(s, pt) : pt;
            }
        }

        if (smallestPt is not { } smallest || smallest >= TinyFontWarnPt)
        {
            return;
        }

        var severity = smallest < TinyFontErrorPt ? "error" : "warning";
        findings.Add(new AuditFinding
        {
            Id = Id("a11y_tiny_font", shapePath),
            Severity = severity,
            Category = "accessibility",
            Code = "a11y_tiny_font",
            Path = shapePath,
            Message = Units.Inv($"Text is {smallest:0.#}pt — below the {TinyFontWarnPt:0}pt readability threshold."),
            Suggestion = $"Increase the size: {{\"op\":\"set\",\"path\":\"{shapePath}\",\"props\":{{\"fontSize\":18}}}}; " +
                "aim for 18pt+ for body text and 24pt+ for titles.",
            Autofixable = false,
        });
    }

    /// <summary>
    /// a11y_low_contrast: the explicit run color vs the shape's own fill (or the
    /// slide background) falls below WCAG AA 4.5:1. Skipped when either color is
    /// theme-driven (we only judge explicit RGB).
    /// </summary>
    private static void CheckContrast(List<AuditFinding> findings, ShapeView shape, string shapePath, string? slideBgHex)
    {
        if (shape.Element is not P.Shape { TextBody: { } body })
        {
            return;
        }

        // Background behind the text: the shape's own solid fill, else the slide background.
        var bgHex = PptxDoc.FillHex(shape.Element) ?? slideBgHex;
        if (bgHex is null || !TryParseHex(bgHex, out var bg))
        {
            return; // a theme/inherited background can't be judged offline
        }

        foreach (var run in body.Descendants<A.Run>())
        {
            if (run.Text?.Text is not { Length: > 0 })
            {
                continue;
            }

            var colorHex = run.RunProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value;
            if (colorHex is null || !TryParseHex(colorHex, out var fg))
            {
                continue;
            }

            var ratio = ContrastRatio(fg, bg);
            if (ratio >= MinContrastRatio)
            {
                continue;
            }

            findings.Add(new AuditFinding
            {
                Id = Id("a11y_low_contrast", shapePath),
                Severity = "warning",
                Category = "accessibility",
                Code = "a11y_low_contrast",
                Path = shapePath,
                Message = Units.Inv($"Text #{colorHex.ToUpperInvariant()} on #{bgHex.ToUpperInvariant()} has a {ratio:0.0}:1 contrast ratio — below WCAG AA 4.5:1."),
                Suggestion = $"Darken or lighten the text/fill so the ratio reaches 4.5:1, e.g. " +
                    $"{{\"op\":\"set\",\"path\":\"{shapePath}\",\"props\":{{\"color\":\"000000\"}}}}.",
                Autofixable = false,
            });
            return; // one finding per shape is enough to flag it
        }
    }

    // ---- quality checks -----------------------------------------------------

    /// <summary>quality_off_canvas: a shape whose box lies entirely outside the slide bounds is invisible.</summary>
    private static void CheckOffCanvas(List<AuditFinding> findings, ShapeView shape, string shapePath, long slideW, long slideH)
    {
        if (PptxDoc.Geometry(shape.Element) is not { } g)
        {
            return;
        }

        var fullyOff = g.X + g.Cx <= 0 || g.Y + g.Cy <= 0 || g.X >= slideW || g.Y >= slideH;
        if (!fullyOff)
        {
            return;
        }

        findings.Add(new AuditFinding
        {
            Id = Id("quality_off_canvas", shapePath),
            Severity = "warning",
            Category = "quality",
            Code = "quality_off_canvas",
            Path = shapePath,
            Message = $"The {KindLabel(shape)} '{NameOrKind(shape)}' is positioned entirely off the slide and will not be visible.",
            Suggestion = $"Move it onto the canvas, e.g. {{\"op\":\"set\",\"path\":\"{shapePath}\",\"props\":{{\"x\":2,\"y\":2}}}}, " +
                "or remove it if it is leftover.",
            Autofixable = false,
        });
    }

    /// <summary>quality_empty_placeholder: an empty (non-title) placeholder adds clutter with no content.</summary>
    private static void CheckEmptyPlaceholder(List<AuditFinding> findings, ShapeView shape, string shapePath)
    {
        var placeholder = PptxDoc.PlaceholderType(shape.Element);
        if (placeholder is null)
        {
            return;
        }

        // The empty-title case is owned by a11y_no_slide_title; here we flag other empties.
        if (IsTitlePlaceholder(shape))
        {
            return;
        }

        if (PptxDoc.ShapeText(shape.Element).Trim().Length > 0)
        {
            return;
        }

        findings.Add(new AuditFinding
        {
            Id = Id("quality_empty_placeholder", shapePath),
            Severity = "warning",
            Category = "quality",
            Code = "quality_empty_placeholder",
            Path = shapePath,
            Message = $"The '{placeholder}' placeholder is empty; an unused placeholder clutters the slide.",
            Suggestion = $"Fill it with {{\"op\":\"set\",\"path\":\"{shapePath}\",\"props\":{{\"text\":\"…\"}}}}, " +
                "or run 'aioffice audit <file> --fix' to remove it.",
            Autofixable = true,
        });
    }

    /// <summary>quality_duplicate_id: two shapes on one slide sharing a cNvPr id breaks @id addressing.</summary>
    private static void CheckDuplicateIds(List<AuditFinding> findings, List<ShapeView> shapes, string slidePath)
    {
        foreach (var group in shapes.GroupBy(s => s.Id).Where(g => g.Count() > 1))
        {
            findings.Add(new AuditFinding
            {
                Id = Id("quality_duplicate_id", Units.Inv($"{slidePath}/shape[@id={group.Key}]")),
                Severity = "error",
                Category = "quality",
                Code = "quality_duplicate_id",
                Path = slidePath,
                Message = Units.Inv($"{group.Count()} shapes on this slide share id {group.Key}; @id addressing is ambiguous."),
                Suggestion = "Give each shape a unique id in PowerPoint (cut and re-paste the duplicate), " +
                    "then re-run the audit.",
                Autofixable = false,
            });
        }
    }

    // ---- fixes --------------------------------------------------------------

    /// <summary>Applies the safe autofixes for the supplied findings; returns how many landed.</summary>
    private static int ApplyFixes(PresentationPart presentation, List<AuditFinding> findings)
    {
        var slides = PptxDoc.Slides(presentation);
        var fixed_ = 0;

        foreach (var finding in findings)
        {
            if (finding.Path is null)
            {
                continue;
            }

            var address = PptxAddress.Parse(finding.Path);
            if (address.SlideIndex < 1 || address.SlideIndex > slides.Count)
            {
                continue;
            }

            var slidePart = slides[address.SlideIndex - 1].Part;
            switch (finding.Code)
            {
                case "a11y_no_alt_text":
                    if (FixAltText(slidePart, address))
                    {
                        fixed_++;
                    }

                    break;

                case "a11y_no_slide_title":
                    if (FixSlideTitle(slidePart))
                    {
                        fixed_++;
                    }

                    break;

                case "quality_empty_placeholder":
                    if (FixEmptyPlaceholder(slidePart, address))
                    {
                        fixed_++;
                    }

                    break;
            }
        }

        return fixed_;
    }

    /// <summary>Sets a placeholder description on the shape, when it still lacks one.</summary>
    private static bool FixAltText(SlidePart slidePart, PptxAddress address)
    {
        var view = PptxDoc.ResolveShape(slidePart, address);
        if (PptxDoc.AltText(view.Element) is not null)
        {
            return false;
        }

        var nonVisual = PptxDoc.NonVisualProps(view.Element);
        if (nonVisual is null)
        {
            return false;
        }

        nonVisual.Description = PlaceholderAltText;
        return true;
    }

    /// <summary>
    /// Ensures the slide has a title with text: fills an empty title placeholder,
    /// or adds a title shape carrying the placeholder text.
    /// </summary>
    private static bool FixSlideTitle(SlidePart slidePart)
    {
        var existing = PptxDoc.Shapes(slidePart).FirstOrDefault(IsTitlePlaceholder);
        if (existing is not null)
        {
            if (existing.Element is P.Shape shape && PptxDoc.ShapeText(shape).Trim().Length == 0)
            {
                PptxEditor.ReplaceText(shape, PlaceholderTitleText);
                return true;
            }

            return false;
        }

        PptxEditor.AddTitleShape(slidePart, PlaceholderTitleText);
        return true;
    }

    /// <summary>Removes the empty placeholder shape (safe: it carries no content).</summary>
    private static bool FixEmptyPlaceholder(SlidePart slidePart, PptxAddress address)
    {
        var view = PptxDoc.ResolveShape(slidePart, address);
        if (PptxDoc.PlaceholderType(view.Element) is null ||
            IsTitlePlaceholder(view) ||
            PptxDoc.ShapeText(view.Element).Trim().Length > 0)
        {
            return false;
        }

        view.Element.Remove();
        return true;
    }

    // ---- helpers ------------------------------------------------------------

    private static bool IsTitlePlaceholder(ShapeView shape) =>
        PptxDoc.PlaceholderType(shape.Element) is "title" or "ctrTitle";

    private static (long Width, long Height) SlideBoundsEmu(PresentationPart presentation)
    {
        var size = presentation.Presentation?.SlideSize;
        return (size?.Cx?.Value ?? PptxFactory.SlideWidthEmu, size?.Cy?.Value ?? PptxFactory.SlideHeightEmu);
    }

    private static string KindLabel(ShapeView shape) => shape.Element switch
    {
        P.Picture => "picture",
        P.GraphicFrame => "chart/table",
        P.GroupShape => "group",
        _ => shape.Kind,
    };

    private static string NameOrKind(ShapeView shape) =>
        shape.Name.Length > 0 ? shape.Name : KindLabel(shape);

    private static string Id(string code, string path) => Units.Inv($"{code}#{path}");

    // ---- color math (WCAG relative luminance) -------------------------------

    private static bool TryParseHex(string hex, out (double R, double G, double B) rgb)
    {
        rgb = default;
        var text = hex.Trim().TrimStart('#');
        if (text.Length != 6 || !text.All(Uri.IsHexDigit))
        {
            return false;
        }

        var r = int.Parse(text[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var g = int.Parse(text[2..4], NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        var b = int.Parse(text[4..6], NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        rgb = (r / 255.0, g / 255.0, b / 255.0);
        return true;
    }

    private static double ContrastRatio((double R, double G, double B) a, (double R, double G, double B) b)
    {
        var la = RelativeLuminance(a);
        var lb = RelativeLuminance(b);
        var (light, dark) = la >= lb ? (la, lb) : (lb, la);
        return (light + 0.05) / (dark + 0.05);
    }

    private static double RelativeLuminance((double R, double G, double B) c) =>
        (0.2126 * LinearizeChannel(c.R)) + (0.7152 * LinearizeChannel(c.G)) + (0.0722 * LinearizeChannel(c.B));

    private static double LinearizeChannel(double channel) =>
        channel <= 0.03928 ? channel / 12.92 : Math.Pow((channel + 0.055) / 1.055, 2.4);
}
