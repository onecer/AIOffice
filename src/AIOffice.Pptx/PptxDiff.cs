using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// The pptx semantic differ (<c>aioffice diff</c>). It opens the baseline (OLD)
/// and current (NEW) decks, matches slides by their stable p:sldId (falling back
/// to a title+text content hash so a re-saved deck still lines up), classifies
/// each slide as added/removed/moved/modified, and within matched slides
/// compares shapes by their cNvPr id (text, position, fill, font, hyperlink).
/// Presentation-level changes — slide size, the first master's background and
/// the section list — surface as <c>modified</c> entries on their canonical
/// paths. Every change is emitted through <see cref="DiffResult.FromChanges"/>
/// so output is sorted deterministically and identical on every platform.
/// </summary>
public sealed partial class PptxHandler : IDiffer
{
    public DiffResult Diff(CommandContext ctx, string baselineFile)
    {
        var current = ctx.File ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "diff requires a target file.",
            "Pass the current .pptx as the first argument and the baseline as the second.");

        // The baseline must be the same format; a foreign extension is invalid_args
        // naming the mismatch (the differ compares like with like).
        var baselineExt = Path.GetExtension(baselineFile);
        if (!string.Equals(baselineExt, ".pptx", StringComparison.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"diff compares two .pptx decks, but the baseline is '{baselineExt}'.",
                "Pass a .pptx baseline (the OLD deck); diffing across formats is not supported.");
        }

        using var currentStream = PptxDoc.LoadStream(current);
        using var currentDoc = PptxDoc.Open(currentStream, editable: false, current);
        using var baselineStream = PptxDoc.LoadStream(baselineFile);
        using var baselineDoc = PptxDoc.Open(baselineStream, editable: false, baselineFile);

        var newPres = PptxDoc.RequirePresentationPart(currentDoc, current);
        var oldPres = PptxDoc.RequirePresentationPart(baselineDoc, baselineFile);

        var changes = new List<DiffChange>();
        DiffSlideSize(changes, oldPres, newPres);
        DiffMasterBackground(changes, oldPres, newPres);
        DiffSections(changes, oldPres, newPres);
        DiffSlides(changes, oldPres, newPres);

        return DiffResult.FromChanges(changes);
    }

    // ---- presentation-level diffs ------------------------------------------

    /// <summary>A changed p:sldSz surfaces as a modified entry on the presentation root.</summary>
    private static void DiffSlideSize(List<DiffChange> changes, PresentationPart oldPres, PresentationPart newPres)
    {
        var (oldCx, oldCy) = SlideSizeEmu(oldPres);
        var (newCx, newCy) = SlideSizeEmu(newPres);
        if (oldCx == newCx && oldCy == newCy)
        {
            return;
        }

        changes.Add(new DiffChange
        {
            Kind = "modified",
            Path = "/",
            Before = SlideSizeLabel(oldCx, oldCy),
            After = SlideSizeLabel(newCx, newCy),
            Detail = "slide size",
        });
    }

    /// <summary>A changed first-master solid background surfaces as a modified entry on /master[1].</summary>
    private static void DiffMasterBackground(List<DiffChange> changes, PresentationPart oldPres, PresentationPart newPres)
    {
        var oldMasters = PptxDoc.Masters(oldPres);
        var newMasters = PptxDoc.Masters(newPres);
        if (oldMasters.Count == 0 || newMasters.Count == 0)
        {
            return;
        }

        var oldBg = MasterBackgroundHex(oldMasters[0].Part);
        var newBg = MasterBackgroundHex(newMasters[0].Part);
        if (string.Equals(oldBg, newBg, StringComparison.Ordinal))
        {
            return;
        }

        changes.Add(new DiffChange
        {
            Kind = "modified",
            Path = "/master[1]",
            Before = oldBg ?? "(theme)",
            After = newBg ?? "(theme)",
            Detail = "master background",
        });
    }

    /// <summary>
    /// Section changes by 1-based position: a renamed section is modified, a new
    /// section is added, a dropped one is removed. The slide membership of a
    /// section follows the slide diff, so only the section identity is compared.
    /// </summary>
    private static void DiffSections(List<DiffChange> changes, PresentationPart oldPres, PresentationPart newPres)
    {
        var oldSections = PptxSections.List(oldPres);
        var newSections = PptxSections.List(newPres);
        var max = Math.Max(oldSections.Count, newSections.Count);

        for (var i = 0; i < max; i++)
        {
            var path = Units.Inv($"/section[{i + 1}]");
            var oldName = i < oldSections.Count ? oldSections[i].Name : null;
            var newName = i < newSections.Count ? newSections[i].Name : null;

            if (oldName is null && newName is not null)
            {
                changes.Add(new DiffChange { Kind = "added", Path = path, After = newName, Detail = "section" });
            }
            else if (oldName is not null && newName is null)
            {
                changes.Add(new DiffChange { Kind = "removed", Path = path, Before = oldName, Detail = "section" });
            }
            else if (oldName is not null && newName is not null && !string.Equals(oldName, newName, StringComparison.Ordinal))
            {
                changes.Add(new DiffChange
                {
                    Kind = "modified",
                    Path = path,
                    Before = oldName,
                    After = newName,
                    Detail = "section name",
                });
            }
        }
    }

    // ---- slide diff ---------------------------------------------------------

    private sealed record SlideEntry(int Index, SlidePart Part, uint SldId, string Hash);

    /// <summary>
    /// Matches slides between the two decks (sldId first, then content hash),
    /// classifying each as moved/modified (matched) or added/removed (unmatched),
    /// and recurses into matched slides to diff their shapes.
    /// </summary>
    private static void DiffSlides(List<DiffChange> changes, PresentationPart oldPres, PresentationPart newPres)
    {
        var oldSlides = SlideEntries(oldPres);
        var newSlides = SlideEntries(newPres);

        // newIndex(1-based) -> matched old entry. We match in two passes so a
        // stable sldId always wins over a coincidental content-hash collision.
        var matchOldToNew = new Dictionary<int, SlideEntry>(); // oldIndex -> new entry
        var matchNewToOld = new Dictionary<int, SlideEntry>(); // newIndex -> old entry

        // Pass 1: identical p:sldId values (the common "edited copy" case).
        var oldById = oldSlides
            .GroupBy(s => s.SldId)
            .Where(g => g.Count() == 1)
            .ToDictionary(g => g.Key, g => g.Single());
        var newById = newSlides
            .GroupBy(s => s.SldId)
            .Where(g => g.Count() == 1)
            .ToDictionary(g => g.Key, g => g.Single());
        foreach (var newSlide in newSlides)
        {
            if (oldById.TryGetValue(newSlide.SldId, out var oldSlide) && newById[newSlide.SldId] == newSlide)
            {
                matchNewToOld[newSlide.Index] = oldSlide;
                matchOldToNew[oldSlide.Index] = newSlide;
            }
        }

        // Pass 2: identical content hash among the still-unmatched slides (so a
        // re-saved deck whose sldIds were renumbered still lines up). Greedy, in
        // order, one-to-one.
        var unmatchedOld = oldSlides.Where(s => !matchOldToNew.ContainsKey(s.Index)).ToList();
        foreach (var newSlide in newSlides.Where(s => !matchNewToOld.ContainsKey(s.Index)))
        {
            var oldSlide = unmatchedOld.FirstOrDefault(o => string.Equals(o.Hash, newSlide.Hash, StringComparison.Ordinal));
            if (oldSlide is not null)
            {
                matchNewToOld[newSlide.Index] = oldSlide;
                matchOldToNew[oldSlide.Index] = newSlide;
                unmatchedOld.Remove(oldSlide);
            }
        }

        // Removed: an old slide with no match. Path is the baseline path.
        foreach (var oldSlide in oldSlides.Where(s => !matchOldToNew.ContainsKey(s.Index)))
        {
            changes.Add(new DiffChange
            {
                Kind = "removed",
                Path = Units.Inv($"/slide[{oldSlide.Index}]"),
                Before = SlideLabel(oldSlide),
                Detail = "slide removed",
            });
        }

        // Added: a new slide with no match. Path is the current path.
        foreach (var newSlide in newSlides.Where(s => !matchNewToOld.ContainsKey(s.Index)))
        {
            changes.Add(new DiffChange
            {
                Kind = "added",
                Path = Units.Inv($"/slide[{newSlide.Index}]"),
                After = SlideLabel(newSlide),
                Detail = "slide added",
            });
        }

        // Matched: a slide is "moved" when its relative order changed. We compare
        // the matched old indices in new order against their sorted order: a slide
        // whose old index is out of ascending order relative to its neighbours
        // moved. We use the longest increasing subsequence so the fewest slides
        // are flagged as moved (the rest are the stable anchor).
        var matchedNew = newSlides.Where(s => matchNewToOld.ContainsKey(s.Index)).ToList();
        var oldOrder = matchedNew.Select(s => matchNewToOld[s.Index].Index).ToList();
        var stable = LongestIncreasingSubsequenceMembers(oldOrder);

        for (var i = 0; i < matchedNew.Count; i++)
        {
            var newSlide = matchedNew[i];
            var oldSlide = matchNewToOld[newSlide.Index];
            if (!stable.Contains(i))
            {
                changes.Add(new DiffChange
                {
                    Kind = "moved",
                    Path = Units.Inv($"/slide[{newSlide.Index}]"),
                    Detail = Units.Inv($"slide reordered {oldSlide.Index} -> {newSlide.Index}"),
                });
            }

            DiffSlideShapes(changes, oldPres, newPres, oldSlide, newSlide);
        }
    }

    /// <summary>
    /// Compares the shapes of two matched slides by their cNvPr id: a shape on
    /// the new slide with no old counterpart is added, an old one with no new
    /// counterpart is removed, and a matched pair is compared for text, position,
    /// fill, font and hyperlink changes (one modified entry per facet).
    /// </summary>
    private static void DiffSlideShapes(
        List<DiffChange> changes, PresentationPart oldPres, PresentationPart newPres, SlideEntry oldSlide, SlideEntry newSlide)
    {
        var newIndex = newSlide.Index;
        var oldShapes = PptxDoc.Shapes(oldSlide.Part).GroupBy(s => s.Id).Where(g => g.Count() == 1).ToDictionary(g => g.Key, g => g.Single());
        var newShapes = PptxDoc.Shapes(newSlide.Part);
        var matchedOldIds = new HashSet<uint>();

        foreach (var newShape in newShapes)
        {
            var path = newShape.CanonicalPath(newIndex);
            if (!oldShapes.TryGetValue(newShape.Id, out var oldShape))
            {
                changes.Add(new DiffChange
                {
                    Kind = "added",
                    Path = path,
                    After = ShapeLabel(newShape),
                    Detail = "shape added",
                });
                continue;
            }

            matchedOldIds.Add(newShape.Id);
            DiffShapePair(changes, oldPres, oldSlide.Part, newPres, newSlide.Part, oldShape, newShape, path);
        }

        // Removed shapes keep their baseline path (they no longer exist in new).
        foreach (var (id, oldShape) in oldShapes.Where(kv => !matchedOldIds.Contains(kv.Key)))
        {
            changes.Add(new DiffChange
            {
                Kind = "removed",
                Path = Units.Inv($"/slide[{newIndex}]/shape[@id={id}]"),
                Before = ShapeLabel(oldShape),
                Detail = "shape removed",
            });
        }
    }

    /// <summary>Compares one matched shape pair, emitting a modified entry per changed facet.</summary>
    private static void DiffShapePair(
        List<DiffChange> changes,
        PresentationPart oldPres,
        SlidePart oldSlidePart,
        PresentationPart newPres,
        SlidePart newSlidePart,
        ShapeView oldShape,
        ShapeView newShape,
        string path)
    {
        // Text.
        var oldText = PptxDoc.ShapeText(oldShape.Element);
        var newText = PptxDoc.ShapeText(newShape.Element);
        if (!string.Equals(oldText, newText, StringComparison.Ordinal))
        {
            changes.Add(new DiffChange
            {
                Kind = "modified",
                Path = path,
                Before = PptxQueryEngine.Snippet(oldText),
                After = PptxQueryEngine.Snippet(newText),
                Detail = "text",
            });
        }

        // Position / size (cm-rounded so sub-pixel jitter is not a change).
        var oldGeo = PptxDoc.Geometry(oldShape.Element);
        var newGeo = PptxDoc.Geometry(newShape.Element);
        if (!GeometryEqual(oldGeo, newGeo))
        {
            changes.Add(new DiffChange
            {
                Kind = "modified",
                Path = path,
                Before = GeometryLabel(oldGeo),
                After = GeometryLabel(newGeo),
                Detail = "moved on slide",
            });
        }

        // Fill.
        var oldFill = PptxDoc.FillHex(oldShape.Element);
        var newFill = PptxDoc.FillHex(newShape.Element);
        if (!string.Equals(oldFill, newFill, StringComparison.Ordinal))
        {
            changes.Add(new DiffChange
            {
                Kind = "modified",
                Path = path,
                Before = oldFill ?? "(none)",
                After = newFill ?? "(none)",
                Detail = "fill",
            });
        }

        // Font (size/bold/color of the first run, the shape's dominant style).
        var oldFont = FontSignature(oldShape.Element);
        var newFont = FontSignature(newShape.Element);
        if (!string.Equals(oldFont, newFont, StringComparison.Ordinal))
        {
            changes.Add(new DiffChange
            {
                Kind = "modified",
                Path = path,
                Before = oldFont,
                After = newFont,
                Detail = "font",
            });
        }

        // Hyperlink (the M8 interactive-deck facet).
        var oldLink = PptxHyperlinks.Resolve(oldPres, oldSlidePart, oldShape.Element);
        var newLink = PptxHyperlinks.Resolve(newPres, newSlidePart, newShape.Element);
        if (!string.Equals(oldLink, newLink, StringComparison.Ordinal))
        {
            changes.Add(new DiffChange
            {
                Kind = "modified",
                Path = path,
                Before = oldLink ?? "(none)",
                After = newLink ?? "(none)",
                Detail = "hyperlink",
            });
        }
    }

    // ---- fingerprints & labels ---------------------------------------------

    private static List<SlideEntry> SlideEntries(PresentationPart presentation)
    {
        var entries = new List<SlideEntry>();
        var index = 0;
        foreach (var (slideId, part) in PptxDoc.Slides(presentation))
        {
            index++;
            entries.Add(new SlideEntry(index, part, slideId.Id?.Value ?? 0, SlideHash(part)));
        }

        return entries;
    }

    /// <summary>
    /// A stable content hash of a slide: every shape's kind and visible text, in
    /// document order. It deliberately excludes shape ids (which a re-save can
    /// renumber) so two slides with the same visible content hash-match even when
    /// their sldId differs — the fallback after the sldId pass.
    /// </summary>
    private static string SlideHash(SlidePart slidePart)
    {
        var builder = new StringBuilder();
        foreach (var shape in PptxDoc.Shapes(slidePart))
        {
            builder.Append(shape.Kind).Append('\x1f');
            builder.Append(PptxDoc.ShapeText(shape.Element)).Append('\x1e');
        }

        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(builder.ToString()));
        return Convert.ToHexString(bytes);
    }

    /// <summary>The slide's title text (its title placeholder), else its first shape text, for labels.</summary>
    private static string SlideLabel(SlideEntry entry)
    {
        var shapes = PptxDoc.Shapes(entry.Part);
        var title = shapes.FirstOrDefault(s => PptxDoc.PlaceholderType(s.Element) is "title" or "ctrTitle");
        var text = title is not null
            ? PptxDoc.ShapeText(title.Element)
            : shapes.Select(s => PptxDoc.ShapeText(s.Element)).FirstOrDefault(t => t.Length > 0) ?? string.Empty;
        var snippet = PptxQueryEngine.Snippet(text);
        return snippet.Length == 0 ? Units.Inv($"slide {entry.Index}") : snippet;
    }

    private static string ShapeLabel(ShapeView shape)
    {
        var text = PptxQueryEngine.Snippet(PptxDoc.ShapeText(shape.Element));
        if (text.Length > 0)
        {
            return text;
        }

        return shape.Name.Length > 0 ? shape.Name : shape.Kind;
    }

    /// <summary>A compact size/bold/color signature of a shape's first text run.</summary>
    private static string FontSignature(OpenXmlCompositeElement element)
    {
        var run = (element as P.Shape)?.TextBody?.Descendants<A.Run>()
            .FirstOrDefault(r => r.Text?.Text is { Length: > 0 });
        var rPr = run?.RunProperties;
        var size = rPr?.FontSize?.Value is { } s ? (s / 100.0).ToString("0.#", CultureInfo.InvariantCulture) : "-";
        var bold = rPr?.Bold?.Value == true ? "b" : "-";
        var color = rPr?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant() ?? "-";
        return Units.Inv($"{size}/{bold}/{color}");
    }

    private static bool GeometryEqual(GeometryEmu? a, GeometryEmu? b)
    {
        if (a is null && b is null)
        {
            return true;
        }

        if (a is not { } ga || b is not { } gb)
        {
            return false;
        }

        // Compare at cm precision so an export's sub-EMU rounding is not a change.
        return Units.EmuToCm(ga.X) == Units.EmuToCm(gb.X)
            && Units.EmuToCm(ga.Y) == Units.EmuToCm(gb.Y)
            && Units.EmuToCm(ga.Cx) == Units.EmuToCm(gb.Cx)
            && Units.EmuToCm(ga.Cy) == Units.EmuToCm(gb.Cy);
    }

    private static string GeometryLabel(GeometryEmu? geo)
    {
        if (geo is not { } g)
        {
            return "(no position)";
        }

        return Units.Inv($"{Units.EmuToCm(g.X)},{Units.EmuToCm(g.Y)} {Units.EmuToCm(g.Cx)}x{Units.EmuToCm(g.Cy)}cm");
    }

    private static (long Cx, long Cy) SlideSizeEmu(PresentationPart presentation)
    {
        var size = presentation.Presentation?.SlideSize;
        return (size?.Cx?.Value ?? PptxFactory.SlideWidthEmu, size?.Cy?.Value ?? PptxFactory.SlideHeightEmu);
    }

    private static string SlideSizeLabel(long cx, long cy)
    {
        var preset = PptxSlideSize.MatchPreset(cx, cy);
        return preset ?? Units.Inv($"{Units.EmuToCm(cx)}x{Units.EmuToCm(cy)}cm");
    }

    /// <summary>The first master's solid RRGGBB background, when set with an explicit RGB color.</summary>
    private static string? MasterBackgroundHex(SlideMasterPart masterPart) =>
        masterPart.SlideMaster?.CommonSlideData?.Background?
            .BackgroundProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant();

    /// <summary>
    /// The indices (positions in the supplied sequence) that form a longest
    /// strictly-increasing subsequence; these are the slides that did NOT move, so
    /// everything else is flagged as moved. Deterministic for any input.
    /// </summary>
    private static HashSet<int> LongestIncreasingSubsequenceMembers(IReadOnlyList<int> sequence)
    {
        var n = sequence.Count;
        if (n == 0)
        {
            return [];
        }

        // tails[k] = index (into sequence) of the smallest tail of an increasing
        // subsequence of length k+1; prev[i] links the chain for reconstruction.
        var tailsIndex = new List<int>();
        var prev = new int[n];
        for (var i = 0; i < n; i++)
        {
            prev[i] = -1;
            var lo = 0;
            var hi = tailsIndex.Count;
            while (lo < hi)
            {
                var mid = (lo + hi) / 2;
                if (sequence[tailsIndex[mid]] < sequence[i])
                {
                    lo = mid + 1;
                }
                else
                {
                    hi = mid;
                }
            }

            if (lo > 0)
            {
                prev[i] = tailsIndex[lo - 1];
            }

            if (lo == tailsIndex.Count)
            {
                tailsIndex.Add(i);
            }
            else
            {
                tailsIndex[lo] = i;
            }
        }

        var members = new HashSet<int>();
        for (var k = tailsIndex[^1]; k >= 0; k = prev[k])
        {
            members.Add(k);
        }

        return members;
    }
}
