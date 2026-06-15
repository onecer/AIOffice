using System.Diagnostics;
using System.Globalization;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// The .pptx format handler. Reads open the package from an in-memory copy and
/// never write; edits are atomic (all ops apply to the in-memory package, the
/// file is only rewritten when every op succeeded) and snapshot the pre-image
/// when a <see cref="SnapshotStore"/> is provided.
/// </summary>
public sealed partial class PptxHandler : IFormatHandler, IEmbedHost
{
    private static readonly IReadOnlyList<string> Views = ["outline", "text", "stats", "structure", "comments", "properties", "embeds"];
    private static readonly IReadOnlyList<string> RenderTargets = ["svg", "html", "text"];

    [GeneratedRegex(@"^([0-9]+)(?:\.\.([0-9]+))?$")]
    private static partial Regex RangePattern();

    private readonly SnapshotStore? _snapshots;

    /// <param name="snapshots">When provided, every successful edit/template-in-place snapshots the pre-image.</param>
    public PptxHandler(SnapshotStore? snapshots = null) => _snapshots = snapshots;

    public DocumentKind Kind => DocumentKind.Pptx;

    public Envelope Create(CommandContext ctx) => Execute(ctx, file =>
    {
        if (File.Exists(file))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"File already exists: {Path.GetFileName(file)}",
                "Pick a new path, or edit the existing file with 'aioffice edit'.");
        }

        if (Path.GetDirectoryName(file) is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
        }

        PptxFactory.CreateMinimal(file, J.Str(ctx.Args, "title"));
        return new { File = file, Kind = "pptx", Slides = 1 };
    });

    public Envelope Read(CommandContext ctx) => Execute(ctx, file =>
    {
        var view = J.Str(ctx.Args, "view") ?? "outline";
        if (!Views.Contains(view, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown view '{view}' for pptx.",
                "Use --view outline, text, stats, structure, comments, properties or embeds.",
                candidates: Views);
        }

        using var stream = PptxDoc.LoadStream(file);
        using var doc = PptxDoc.Open(stream, editable: false, file);

        // Document properties are a package-level node, not slide content.
        if (view == "properties")
        {
            return PptxProperties.Shape(doc);
        }

        var presentation = PptxDoc.RequirePresentationPart(doc, file);
        var slides = SlidesInRange(presentation, J.Str(ctx.Args, "range"));

        return view switch
        {
            "outline" => BuildOutline(presentation, slides),
            "text" => BuildTextView(slides),
            "stats" => BuildStats(slides),
            "comments" => PptxComments.CommentsView(presentation, slides.Select(s => (s.Index, s.Part))),
            "embeds" => BuildEmbedsView(presentation),
            _ => BuildStructure(presentation, slides),
        };
    });

    public Envelope Get(CommandContext ctx) => Execute(ctx, file =>
    {
        var address = PptxAddress.Parse(RequireArg(ctx, "path"));
        if (address.RunIndex is not null)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "Run-level get is not supported yet.",
                "Get the paragraph instead: /slide[i]/shape[j]/p[k].");
        }

        using var stream = PptxDoc.LoadStream(file);
        using var doc = PptxDoc.Open(stream, editable: false, file);

        // Document properties are a package-level node, before the presentation part.
        if (address.IsProperties)
        {
            return PptxProperties.Shape(doc);
        }

        var presentation = PptxDoc.RequirePresentationPart(doc, file);

        if (address.IsPresentation)
        {
            return PptxSlideSize.Detail(presentation);
        }

        if (address.IsSection)
        {
            return PptxSections.Detail(presentation, address);
        }

        if (address.IsNotes)
        {
            return PptxNotes.NotesDetail(presentation, address);
        }

        if (address.IsChart)
        {
            return PptxCharts.Detail(presentation, address);
        }

        if (address.IsTable)
        {
            return PptxTables.Detail(presentation, address);
        }

        if (address.IsSmartArt)
        {
            return PptxSmartArt.Detail(presentation, address);
        }

        if (address.IsGroup)
        {
            return PptxQueryEngine.GroupDetail(presentation, address);
        }

        if (address.IsAnimation)
        {
            return PptxAnimations.Detail(presentation, address);
        }

        if (address.IsComment)
        {
            return PptxComments.Detail(presentation, address);
        }

        if (address.IsEmbed)
        {
            return PptxEmbeds.Detail(presentation, address);
        }

        if (address.IsMedia)
        {
            return PptxMedia.Detail(presentation, address);
        }

        if (address.IsOMath)
        {
            return PptxEquations.Detail(presentation, address);
        }

        if (address.IsMaster)
        {
            return address.HasShape
                ? PptxQueryEngine.MasterShapeDetail(presentation, address)
                : address.LayoutIndex is null
                    ? PptxQueryEngine.MasterDetail(presentation, address)
                    : PptxQueryEngine.LayoutDetail(presentation, address);
        }

        return address.HasShape
            ? PptxQueryEngine.ShapeDetail(presentation, address)
            : PptxQueryEngine.SlideDetail(presentation, address);
    });

    public Envelope Query(CommandContext ctx) => Execute(ctx, file =>
    {
        var selector = Selector.Parse(RequireArg(ctx, "selector"));
        var scope = J.Str(ctx.Args, "scope") is { } scopeRaw ? PptxAddress.Parse(scopeRaw) : null;
        using var stream = PptxDoc.LoadStream(file);
        using var doc = PptxDoc.Open(stream, editable: false, file);
        var presentation = PptxDoc.RequirePresentationPart(doc, file);
        var matches = PptxQueryEngine.Query(presentation, selector, scope);
        return new
        {
            Selector = selector.ToCanonicalString(),
            Scope = scope?.CanonicalContainerPath,
            Count = matches.Count,
            Matches = matches,
        };
    });

    public Envelope Edit(CommandContext ctx, IReadOnlyList<EditOp> ops) => Execute(ctx, (file, warnings) =>
    {
        if (ops.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "The edit batch is empty.",
                "Pass at least one operation in --ops.");
        }

        CheckExpectedRev(ctx, file);

        var results = new List<object>();
        int slideCount;
        var mutated = false;
        using var stream = PptxDoc.LoadStream(file);
        using (var doc = PptxDoc.Open(stream, editable: true, file))
        {
            var presentation = PptxDoc.RequirePresentationPart(doc, file);
            foreach (var op in ops)
            {
                var outcome = PptxEditor.Apply(doc, presentation, op, ctx.Workspace);
                mutated |= outcome.Mutated;
                if (outcome.Warnings is { Count: > 0 } opWarnings)
                {
                    warnings.AddRange(opWarnings);
                }

                if (outcome.Replacements is { } replacements)
                {
                    results.Add(new
                    {
                        Op = op.Op,
                        Path = op.Path,
                        Target = outcome.Target,
                        Replacements = replacements,
                        Locations = outcome.Locations,
                    });
                    if (replacements == 0)
                    {
                        warnings.Add(new Warning(
                            "find_no_match",
                            Units.Inv($"replace on '{op.Path}' matched nothing; the deck is unchanged for this op.")));
                    }
                }
                else
                {
                    results.Add(new { Op = op.Op, Path = op.Path, Target = outcome.Target });
                }
            }

            slideCount = PptxDoc.Slides(presentation).Count;
        }

        // Producing-only batches (extract) read the deck but must not rewrite it —
        // re-serializing could perturb embedded sub-packages and the rev. Only
        // snapshot + write when at least one op actually mutated the document.
        int? snapshotNumber = null;
        if (mutated)
        {
            snapshotNumber = _snapshots?.Save(file)?.Number;
            File.WriteAllBytes(file, stream.ToArray());
        }

        return new { Applied = ops.Count, Snapshot = snapshotNumber, Results = results, Slides = slideCount };
    });

    public Envelope Render(CommandContext ctx) => Execute(ctx, file =>
    {
        var to = J.Str(ctx.Args, "to") ?? "svg";
        if (to == "png")
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "PNG rasterization is not available yet (planned for M1).",
                "Render --to svg (or html) and rasterize externally, e.g. 'rsvg-convert slide.svg -o slide.png'.");
        }

        if (!RenderTargets.Contains(to, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown render target '{to}' for pptx.",
                "Use --to svg, html or text.",
                candidates: RenderTargets);
        }

        using var stream = PptxDoc.LoadStream(file);
        using var doc = PptxDoc.Open(stream, editable: false, file);
        var presentation = PptxDoc.RequirePresentationPart(doc, file);
        var slides = SlidesInScope(presentation, J.Str(ctx.Args, "scope"));

        if (to == "text")
        {
            var text = BuildPlainText(slides);
            return WithOptionalOutput(ctx, new { To = to, Text = text }, text, to, slides.Count);
        }

        var rendered = slides
            .Select(s => (Path: Units.Inv($"/slide[{s.Index}]"), Svg: PptxRenderer.RenderSlideSvg(presentation, s.Part, s.Index)))
            .ToList();

        if (to == "html")
        {
            var html = PptxRenderer.WrapHtml(rendered, PptxRenderer.SlideSizePx(presentation).WidthPx);
            return WithOptionalOutput(ctx, new { To = to, SlideCount = rendered.Count, Html = html }, html, to, rendered.Count);
        }

        var payload = new
        {
            To = to,
            SlideCount = rendered.Count,
            Slides = rendered.Select(r => new { Path = r.Path, Svg = r.Svg }).ToList<object>(),
        };
        return WithOptionalOutput(ctx, payload, rendered.Count == 1 ? rendered[0].Svg : null, to, rendered.Count);
    });

    public Envelope Validate(CommandContext ctx) => Execute(ctx, file =>
    {
        using var stream = PptxDoc.LoadStream(file);
        using var doc = PptxDoc.Open(stream, editable: false, file);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2021);
        var issues = validator.Validate(doc).Select(error => new
        {
            Id = error.Id,
            Description = error.Description,
            Part = error.Part?.Uri.ToString(),
            XPath = error.Path?.XPath,
            Severity = error.ErrorType.ToString(),
        }).ToList<object>();

        return new { Valid = issues.Count == 0, Count = issues.Count, Issues = issues };
    });

    public Envelope Template(CommandContext ctx) => Execute(ctx, file =>
    {
        if (ctx.Args["data"] is not JsonObject dataNode || dataNode.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "template requires --data with a non-empty JSON object.",
                "Pass merge values like --data '{\"title\":\"Q3 Review\",\"owner\":\"Dana\"}'.");
        }

        var data = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (key, value) in dataNode)
        {
            if (value is null or JsonObject or JsonArray)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"--data value for '{key}' must be a scalar.",
                    "Use strings/numbers/booleans; nested objects are not merged in this milestone.");
            }

            data[key] = J.ScalarText(value);
        }

        var output = J.Str(ctx.Args, "output") is { } outputArg
            ? ctx.Workspace.Resolve(outputArg)
            : file;

        int replacements;
        using var stream = PptxDoc.LoadStream(file);
        using (var doc = PptxDoc.Open(stream, editable: true, file))
        {
            var presentation = PptxDoc.RequirePresentationPart(doc, file);
            replacements = PptxTemplater.Apply(presentation, data);
        }

        if (string.Equals(output, file, StringComparison.Ordinal))
        {
            _snapshots?.Save(file);
        }
        else if (Path.GetDirectoryName(output) is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllBytes(output, stream.ToArray());
        return new { Replacements = replacements, Output = output };
    });

    // ---- IEmbedHost (M10) ---------------------------------------------------

    /// <summary>Lists every embedded object in the deck, in canonical-path order.</summary>
    public IReadOnlyList<EmbeddedObject> ListEmbeds(CommandContext ctx)
    {
        var file = ctx.File ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "Listing embeds requires a target file.",
            "Pass the .pptx path to read from.");

        using var stream = PptxDoc.LoadStream(file);
        using var doc = PptxDoc.Open(stream, editable: false, file);
        var presentation = PptxDoc.RequirePresentationPart(doc, file);
        return PptxEmbeds.List(presentation);
    }

    /// <summary>Writes the addressed embed's payload to <paramref name="destPath"/> (already sandbox-resolved); the deck is not modified.</summary>
    public void ExtractEmbed(CommandContext ctx, string embedPath, string destPath)
    {
        var file = ctx.File ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "Extracting an embed requires a target file.",
            "Pass the .pptx path to read from.");

        var address = PptxAddress.Parse(embedPath);
        if (!address.IsEmbed)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{embedPath}' is not an embed path.",
                "Use /slide[i]/embed[k] or /slide[i]/embed[@id=N]; run 'aioffice read <file> --view embeds' to list them.");
        }

        using var stream = PptxDoc.LoadStream(file);
        using var doc = PptxDoc.Open(stream, editable: false, file);
        var presentation = PptxDoc.RequirePresentationPart(doc, file);
        PptxEmbeds.Extract(presentation, address, destPath);
    }

    // ---- shared plumbing ----------------------------------------------------

    private static Envelope Execute(CommandContext ctx, Func<string, object?> action) =>
        Execute(ctx, (file, _) => action(file));

    private static Envelope Execute(CommandContext ctx, Func<string, List<Warning>, object?> action)
    {
        var stopwatch = Stopwatch.StartNew();
        string? file = ctx.File;
        var warnings = new List<Warning>();
        try
        {
            if (string.IsNullOrWhiteSpace(file))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "This command requires a target file.",
                    "Pass the .pptx path as the first argument.");
            }

            var data = action(file, warnings);
            return Envelope.Ok(data, BuildMeta(file, stopwatch, warnings));
        }
        catch (Exception exception)
        {
            return Envelope.FromException(exception, BuildMeta(file, stopwatch, warnings: null));
        }
    }

    private static Meta BuildMeta(string? file, Stopwatch stopwatch, List<Warning>? warnings = null) => new()
    {
        File = file,
        Rev = file is not null && File.Exists(file) ? Rev.OfFile(file) : null,
        ElapsedMs = stopwatch.ElapsedMilliseconds,
        Warnings = warnings is { Count: > 0 } ? warnings : null,
    };

    private static string RequireArg(CommandContext ctx, string key)
    {
        return J.Str(ctx.Args, key) ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Missing required argument '{key}'.",
            key == "selector"
                ? "Pass a selector, e.g. shape:contains('Q3') — run 'aioffice help selectors'."
                : "Pass a path, e.g. /slide[1]/shape[2] — run 'aioffice help addressing'.");
    }

    private static void CheckExpectedRev(CommandContext ctx, string file)
    {
        var expected = J.Str(ctx.Args, "expectRev") ?? J.Str(ctx.Args, "expect-rev");
        if (expected is null)
        {
            return;
        }

        var current = Rev.OfFile(file);
        if (!string.Equals(expected, current, StringComparison.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.StaleAddress,
                $"The file changed since it was read: expected rev {expected}, but it is now {current}.",
                "Re-run 'aioffice read' or 'aioffice query' to refresh paths, then retry the edit with the new rev.");
        }
    }

    private sealed record SlideRef(int Index, SlidePart Part);

    private static List<SlideRef> SlidesInRange(PresentationPart presentation, string? range)
    {
        var all = PptxDoc.Slides(presentation)
            .Select((slide, i) => new SlideRef(i + 1, slide.Part))
            .ToList();
        if (range is null)
        {
            return all;
        }

        var match = RangePattern().Match(range.Trim());
        if (!match.Success)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a valid slide range: '{range}'.",
                "Use --range a..b (1-based, inclusive) or a single slide number like --range 2.");
        }

        var from = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
        var to = match.Groups[2].Success ? int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) : from;
        if (from < 1 || to < from)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Slide range bounds must be 1-based and ordered: '{range}'.",
                "Use --range a..b with a <= b, e.g. --range 2..5.");
        }

        return [.. all.Where(s => s.Index >= from && s.Index <= to)];
    }

    private static List<SlideRef> SlidesInScope(PresentationPart presentation, string? scope)
    {
        if (scope is null)
        {
            return SlidesInRange(presentation, null);
        }

        var address = PptxAddress.Parse(scope);
        if (address.IsMaster)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Rendering a master/layout directly is not supported: '{scope}'.",
                "Render a slide that uses the layout instead, e.g. --scope /slide[2].");
        }

        if (address.HasShape || address.IsNotes || address.IsChart || address.IsTable || address.IsSmartArt ||
            address.IsAnimation || address.IsComment)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Render scope must be a slide, not '{scope}'.",
                "Use --scope /slide[2]; shape/chart/table/notes-level rendering is not supported yet.");
        }

        var part = PptxDoc.ResolveSlide(presentation, address.SlideIndex, scope);
        return [new SlideRef(address.SlideIndex, part)];
    }

    private static object BuildOutline(PresentationPart presentation, List<SlideRef> slides)
    {
        var sections = PptxSections.List(presentation);
        return new
        {
            View = "outline",
            Slides = slides.Select(s =>
            {
                var shapes = PptxDoc.Shapes(s.Part);
                var notes = PptxNotes.Text(s.Part);
                var transition = PptxTransitions.Read(s.Part);
                return new
                {
                    Path = Units.Inv($"/slide[{s.Index}]"),
                    Index = s.Index,
                    Section = SectionNameOf(sections, s.Index),
                    Transition = transition?.Kind,
                    TransitionDuration = transition?.Duration,
                    ShapeCount = shapes.Count,
                    Shapes = shapes.Select(shape => new
                    {
                        Path = shape.CanonicalPath(s.Index),
                        Name = shape.Name,
                        Kind = shape.Kind,
                        Text = PptxQueryEngine.Snippet(PptxDoc.ShapeText(shape.Element)),
                    }).ToList<object>(),
                    Notes = notes.Length == 0 ? null : PptxQueryEngine.Snippet(notes),
                };
            }).ToList<object>(),

            // The same slides grouped under their sections (when the deck has any).
            Sections = sections.Count == 0
                ? null
                : sections.Select(section => (object)new
                {
                    Path = Units.Inv($"/section[{section.Index}]"),
                    Index = section.Index,
                    section.Name,
                    Slides = section.Slides
                        .Where(n => slides.Any(s => s.Index == n))
                        .Select(n => Units.Inv($"/slide[{n}]"))
                        .ToList(),
                }).ToList(),
        };
    }

    /// <summary>The name of the section a slide belongs to, or null when it is unsectioned.</summary>
    private static string? SectionNameOf(List<SectionView> sections, int slideIndex) =>
        sections.FirstOrDefault(s => s.Slides.Contains(slideIndex))?.Name;

    private static object BuildTextView(List<SlideRef> slides) => new
    {
        View = "text",
        Text = BuildPlainText(slides, includeNotes: true),
    };

    /// <summary>Slide text blocks; --view text includes a [notes] section per slide, render --to text stays slide-only.</summary>
    private static string BuildPlainText(List<SlideRef> slides, bool includeNotes = false)
    {
        var perSlide = slides.Select(s =>
        {
            var body = string.Join(
                '\n',
                PptxDoc.Shapes(s.Part)
                    .Select(shape => ShapeOrSmartArtText(s.Part, shape))
                    .Where(text => text.Length > 0));
            if (!includeNotes)
            {
                return body;
            }

            var notes = PptxNotes.Text(s.Part);
            if (notes.Length == 0)
            {
                return body;
            }

            var section = "[notes]\n" + notes;
            return body.Length == 0 ? section : body + "\n" + section;
        });
        return string.Join("\n\n", perSlide);
    }

    /// <summary>Shape text, with SmartArt frames contributing their node texts (indented per level).</summary>
    private static string ShapeOrSmartArtText(SlidePart slidePart, ShapeView shape)
    {
        if (PptxSmartArt.DataPartOf(slidePart, shape.Element) is { } dataPart)
        {
            return string.Join('\n', PptxSmartArt.IndentedLines(dataPart));
        }

        return PptxDoc.ShapeText(shape.Element);
    }

    private static object BuildStats(List<SlideRef> slides)
    {
        var shapeCount = 0;
        var characters = 0;
        var words = 0;
        foreach (var slide in slides)
        {
            foreach (var shape in PptxDoc.Shapes(slide.Part))
            {
                shapeCount++;
                var text = PptxDoc.ShapeText(shape.Element);
                characters += text.Length;
                words += text.Split([' ', '\n', '\t'], StringSplitOptions.RemoveEmptyEntries).Length;
            }
        }

        return new { View = "stats", Slides = slides.Count, Shapes = shapeCount, Words = words, Characters = characters };
    }

    /// <summary>The embedded-objects view: every embed's metadata (path, name, mediaType, size), never bytes.</summary>
    private static object BuildEmbedsView(PresentationPart presentation)
    {
        var embeds = PptxEmbeds.List(presentation);
        return new
        {
            View = "embeds",
            Count = embeds.Count,
            Embeds = embeds.Select(e => new
            {
                Path = e.Path,
                Name = e.Name,
                MediaType = e.MediaType,
                Size = e.Size,
                Container = e.Container,
            }).ToList<object>(),
        };
    }

    private static object BuildStructure(PresentationPart presentation, List<SlideRef> slides)
    {
        var size = presentation.Presentation?.SlideSize;
        var cx = size?.Cx?.Value ?? PptxFactory.SlideWidthEmu;
        var cy = size?.Cy?.Value ?? PptxFactory.SlideHeightEmu;
        var sections = PptxSections.List(presentation);
        return new
        {
            View = "structure",
            SlideSize = PptxSlideSize.MatchPreset(cx, cy),
            SlideWidthCm = Units.EmuToCm(cx),
            SlideHeightCm = Units.EmuToCm(cy),
            Sections = sections.Count == 0
                ? null
                : sections.Select(s => (object)new
                {
                    Path = Units.Inv($"/section[{s.Index}]"),
                    Index = s.Index,
                    s.Name,
                    Slides = s.Slides.Select(n => Units.Inv($"/slide[{n}]")).ToList(),
                }).ToList(),
            Masters = PptxDoc.Masters(presentation).Select(m =>
            {
                var masterPath = Units.Inv($"/master[{m.Index}]");
                var layouts = PptxDoc.Layouts(m.Part);
                return new
                {
                    Path = masterPath,
                    Index = m.Index,
                    Theme = m.Part.ThemePart?.Theme?.Name?.Value,
                    ShapeCount = PptxDoc.Shapes(PptxDoc.RequireShapeTree(m.Part)).Count,
                    LayoutCount = layouts.Count,
                    Layouts = layouts
                        .Select(l => (object)PptxQueryEngine.LayoutSummary(presentation, m.Index, l.Index, l.Part))
                        .ToList(),
                };
            }).ToList<object>(),
            Slides = slides.Select(s =>
            {
                var animations = PptxAnimations.List(s.Part);
                var smartArts = PptxSmartArt.List(s.Part);
                var embeds = PptxEmbeds.SlideEmbeds(s.Part, s.Index);
                var media = PptxMedia.SlideMedia(s.Part, s.Index);
                return new
                {
                    Path = Units.Inv($"/slide[{s.Index}]"),
                    Index = s.Index,
                    Layout = PptxDoc.LayoutPathOf(presentation, s.Part),
                    Shapes = PptxDoc.Shapes(s.Part).Select(shape =>
                    {
                        var geometry = PptxDoc.Geometry(shape.Element);
                        var paragraphs = (shape.Element as P.Shape)?.TextBody?.Elements<A.Paragraph>().Count() ?? 0;
                        return new
                        {
                            Path = shape.CanonicalPath(s.Index),
                            OrdinalPath = shape.OrdinalPath(s.Index),
                            Id = shape.Id,
                            Kind = shape.Kind,
                            Name = shape.Name,
                            X = geometry is { } g1 ? Units.EmuToCm(g1.X) : (double?)null,
                            Y = geometry is { } g2 ? Units.EmuToCm(g2.Y) : (double?)null,
                            W = geometry is { } g3 ? Units.EmuToCm(g3.Cx) : (double?)null,
                            H = geometry is { } g4 ? Units.EmuToCm(g4.Cy) : (double?)null,
                            Fill = PptxDoc.FillHex(shape.Element),
                            Paragraphs = paragraphs,
                        };
                    }).ToList<object>(),
                    Animations = animations.Count == 0
                        ? null
                        : animations.Select(a => PptxAnimations.Project(a, s.Index, s.Part)).ToList(),
                    SmartArt = smartArts.Count == 0
                        ? null
                        : smartArts.Select(d => PptxSmartArt.StructureRow(s.Part, s.Index, d.Index, d.View, d.Part)).ToList(),
                    Embeds = embeds.Count == 0
                        ? null
                        : embeds.Select(e => (object)new
                        {
                            Path = e.Path,
                            Name = e.Name,
                            MediaType = e.MediaType,
                            Size = e.Size,
                        }).ToList(),
                    Media = media.Count == 0 ? null : media,
                };
            }).ToList<object>(),
        };
    }

    /// <summary>Writes the rendered payload to --output when requested, returning the enriched data.</summary>
    private static object WithOptionalOutput(CommandContext ctx, object data, string? writable, string to, int slideCount)
    {
        var outputArg = J.Str(ctx.Args, "output");
        if (outputArg is null)
        {
            return data;
        }

        if (writable is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"Cannot write {slideCount} svg slides to a single output file."),
                "Narrow to one slide with --scope /slide[2], or use --to html for a single multi-slide file.");
        }

        var output = ctx.Workspace.Resolve(outputArg);
        if (Path.GetDirectoryName(output) is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(output, writable);
        return new { To = to, SlideCount = slideCount, Output = output };
    }
}
