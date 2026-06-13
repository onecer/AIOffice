using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace AIOffice.Pptx;

/// <summary>One slide section: its 1-based index, name, GUID id and the 1-based slide indices it owns.</summary>
internal sealed record SectionView(int Index, string Name, string Id, IReadOnlyList<int> Slides);

/// <summary>
/// M6 slide sections — the standard PowerPoint sections stored as a
/// p14:sectionLst in presentation.xml's extLst (the
/// {521415D9-36F7-43E2-AB2F-B90AF26B5E84} extension). Sections group slides for
/// the outline view; removing one keeps its slides (they just go unsectioned).
/// Each p14:section references slides by their p:sldId/@id, so the section list
/// is resilient to slide reordering.
/// </summary>
internal static class PptxSections
{
    /// <summary>The OOXML extension URI PowerPoint uses for the section list.</summary>
    private const string SectionExtensionUri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";

    // ---- read --------------------------------------------------------------

    /// <summary>All sections in document order, 1-based, with the slide indices each owns.</summary>
    public static List<SectionView> List(PresentationPart presentation)
    {
        var views = new List<SectionView>();
        var list = FindSectionList(presentation);
        if (list is null)
        {
            return views;
        }

        var slideIdToIndex = SlideIdToIndex(presentation);
        var index = 0;
        foreach (var section in list.Elements<P14.Section>())
        {
            index++;
            var slides = new List<int>();
            foreach (var entry in section.SectionSlideIdList?.Elements<P14.SectionSlideIdListEntry>() ?? [])
            {
                if (entry.Id?.Value is { } id && slideIdToIndex.TryGetValue(id, out var slideIndex))
                {
                    slides.Add(slideIndex);
                }
            }

            slides.Sort();
            views.Add(new SectionView(index, section.Name?.Value ?? string.Empty, section.Id?.Value ?? string.Empty, slides));
        }

        return views;
    }

    /// <summary>Resolves /section[i] or throws invalid_path with candidates.</summary>
    public static SectionView Resolve(PresentationPart presentation, PptxAddress address)
    {
        var sections = List(presentation);
        var index = address.SectionIndex;
        if (index >= 1 && index <= sections.Count)
        {
            return sections[index - 1];
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            Units.Inv($"No section {index}; the deck has {sections.Count} section(s)."),
            sections.Count > 0
                ? "Section indices are 1-based; run 'aioffice read <file> --view outline' to list them."
                : "Add one first: {\"op\":\"add\",\"path\":\"/\",\"type\":\"section\",\"props\":{\"name\":\"Intro\"}}.",
            candidates: [.. sections.Take(10).Select(s => Units.Inv($"/section[{s.Index}]"))]);
    }

    /// <summary>The `get` projection for /section[i].</summary>
    public static object Detail(PresentationPart presentation, PptxAddress address)
    {
        var view = Resolve(presentation, address);
        return new
        {
            Path = Units.Inv($"/section[{view.Index}]"),
            Index = view.Index,
            view.Name,
            SlideCount = view.Slides.Count,
            Slides = view.Slides.Select(s => Units.Inv($"/slide[{s}]")).ToList(),
        };
    }

    // ---- add ---------------------------------------------------------------

    /// <summary>
    /// add /{type:section}: appends a named section. props.afterSlide (0-based; 0 = "before slide 1")
    /// claims every still-unsectioned slide from that point up to the next section's first slide. With
    /// no afterSlide the section claims all slides not yet owned by an earlier section.
    /// </summary>
    public static string Add(PresentationPart presentation, PptxAddress address, JsonObject? props)
    {
        if (!address.IsPresentation)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add section targets the presentation root, not '{address.Raw}'.",
                "Use {\"op\":\"add\",\"path\":\"/\",\"type\":\"section\",\"props\":{\"name\":\"Intro\",\"afterSlide\":0}}.");
        }

        props ??= [];
        string? name = null;
        int? afterSlide = null;
        foreach (var (key, value) in props)
        {
            switch (key)
            {
                case "name":
                    name = value is null ? string.Empty : J.ScalarText(value);
                    break;
                case "afterSlide":
                    afterSlide = ParseAfterSlide(value);
                    break;
                default:
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Unknown add-section prop '{key}'.",
                        "add section takes name (required) and afterSlide (0-based slide index to start after).",
                        candidates: ["name", "afterSlide"]);
            }
        }

        if (string.IsNullOrWhiteSpace(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add section needs a non-empty 'name'.",
                "Pass {\"props\":{\"name\":\"Intro\"}}.");
        }

        var slideIds = SlideIdsInOrder(presentation);
        var slideCount = slideIds.Count;
        var start = afterSlide ?? 0; // claim from slide (start+1) onward
        if (start < 0 || start > slideCount)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"afterSlide {start} is out of range; the deck has {slideCount} slide(s)."),
                "afterSlide is 0-based: 0 starts the section before slide 1, N starts it after slide N.");
        }

        var (_, sectionList) = EnsureSectionList(presentation);

        // The new section spans [start+1 .. nextStart], where nextStart is the first slide of the next
        // existing section that begins after this one. Slides in that span are taken from whatever section
        // currently owns them (a split), so adjacent sections never overlap.
        var existing = List(presentation);
        var nextStart1 = existing
            .Select(s => s.Slides.Count > 0 ? s.Slides.Min() : int.MaxValue)
            .Where(first => first > start) // existing sections whose first slide is after our start (1-based > 0-based start)
            .DefaultIfEmpty(slideCount + 1)
            .Min();

        var claimed = new List<uint>();
        for (var slideIndex1 = start + 1; slideIndex1 < nextStart1 && slideIndex1 <= slideCount; slideIndex1++)
        {
            claimed.Add(slideIds[slideIndex1 - 1]);
        }

        // Take the claimed slides away from any section that currently owns them.
        if (claimed.Count > 0)
        {
            var claimedSet = claimed.ToHashSet();
            foreach (var owner in sectionList.Elements<P14.Section>())
            {
                foreach (var entry in owner.SectionSlideIdList?.Elements<P14.SectionSlideIdListEntry>().ToList() ?? [])
                {
                    if (entry.Id?.Value is { } id && claimedSet.Contains(id))
                    {
                        entry.Remove();
                    }
                }
            }
        }

        var section = new P14.Section { Name = name!, Id = NewSectionId() };
        if (claimed.Count > 0)
        {
            var ids = new P14.SectionSlideIdList();
            foreach (var id in claimed)
            {
                ids.Append(new P14.SectionSlideIdListEntry { Id = id });
            }

            section.SectionSlideIdList = ids;
        }

        // Insert the section in slide order so the outline reads top-to-bottom.
        InsertInSlideOrder(presentation, sectionList, section, start);

        var newIndex = List(presentation).FindIndex(s => string.Equals(s.Id, section.Id?.Value, StringComparison.Ordinal)) + 1;
        return Units.Inv($"/section[{newIndex}]");
    }

    // ---- set (rename) ------------------------------------------------------

    /// <summary>set /section[i] {name}: renames a section.</summary>
    public static string Set(PresentationPart presentation, PptxAddress address, JsonObject props)
    {
        var view = Resolve(presentation, address);
        var list = FindSectionList(presentation)!;
        var section = list.Elements<P14.Section>().ElementAt(view.Index - 1);

        foreach (var (key, value) in props)
        {
            if (!string.Equals(key, "name", StringComparison.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Prop '{key}' does not apply to a section.",
                    "A section's only editable prop is name; reassign slides by removing and re-adding the section.",
                    candidates: ["name"]);
            }

            var name = value is null ? string.Empty : J.ScalarText(value);
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "A section name cannot be empty.",
                    "Pass a non-empty name, e.g. {\"props\":{\"name\":\"Appendix\"}}.");
            }

            section.Name = name;
        }

        return address.CanonicalSectionPath;
    }

    // ---- remove ------------------------------------------------------------

    /// <summary>remove /section[i]: drops the section (its slides survive, just unsectioned). Drops an empty list.</summary>
    public static string Remove(PresentationPart presentation, PptxAddress address)
    {
        var view = Resolve(presentation, address);
        var list = FindSectionList(presentation)!;
        list.Elements<P14.Section>().ElementAt(view.Index - 1).Remove();

        if (!list.Elements<P14.Section>().Any())
        {
            // An empty p14:sectionLst (and its now-empty ext) is pointless; prune it.
            var ext = list.Parent as P.PresentationExtension;
            list.Remove();
            ext?.Remove();
            if (presentation.Presentation?.PresentationExtensionList is { } extList && !extList.HasChildren)
            {
                extList.Remove();
            }
        }

        return address.CanonicalSectionPath;
    }

    // ---- plumbing ----------------------------------------------------------

    private static P14.SectionList? FindSectionList(PresentationPart presentation) =>
        presentation.Presentation?.PresentationExtensionList?
            .Elements<P.PresentationExtension>()
            .Select(e => e.GetFirstChild<P14.SectionList>())
            .FirstOrDefault(s => s is not null);

    private static (P.PresentationExtensionList ExtList, P14.SectionList SectionList) EnsureSectionList(
        PresentationPart presentation)
    {
        var presentationXml = presentation.Presentation ?? throw Corrupt("the presentation part has no presentation XML");
        var existing = FindSectionList(presentation);
        if (existing is not null)
        {
            return ((P.PresentationExtensionList)existing.Parent!.Parent!, existing);
        }

        var extList = presentationXml.PresentationExtensionList ??= new P.PresentationExtensionList();
        var sectionList = new P14.SectionList();
        extList.Append(new P.PresentationExtension(sectionList) { Uri = SectionExtensionUri });
        return (extList, sectionList);
    }

    /// <summary>The p:sldId/@id values in show order.</summary>
    private static List<uint> SlideIdsInOrder(PresentationPart presentation)
    {
        var ids = new List<uint>();
        foreach (var slideId in presentation.Presentation?.SlideIdList?.Elements<P.SlideId>() ?? [])
        {
            if (slideId.Id?.Value is { } id)
            {
                ids.Add(id);
            }
        }

        return ids;
    }

    private static Dictionary<uint, int> SlideIdToIndex(PresentationPart presentation)
    {
        var map = new Dictionary<uint, int>();
        var index = 0;
        foreach (var slideId in presentation.Presentation?.SlideIdList?.Elements<P.SlideId>() ?? [])
        {
            index++;
            if (slideId.Id?.Value is { } id)
            {
                map[id] = index;
            }
        }

        return map;
    }

    /// <summary>Inserts a section so existing sections stay ordered by their first slide.</summary>
    private static void InsertInSlideOrder(
        PresentationPart presentation, P14.SectionList list, P14.Section section, int startSlide0)
    {
        var slideIdToIndex = SlideIdToIndex(presentation);
        P14.Section? insertBefore = null;
        foreach (var existing in list.Elements<P14.Section>())
        {
            var firstSlide = existing.SectionSlideIdList?.Elements<P14.SectionSlideIdListEntry>()
                .Select(e => e.Id?.Value is { } id && slideIdToIndex.TryGetValue(id, out var idx) ? idx : int.MaxValue)
                .DefaultIfEmpty(int.MaxValue)
                .Min() ?? int.MaxValue;

            // startSlide0 is 0-based "after slide N"; the new section's first slide is N+1.
            if (firstSlide > startSlide0)
            {
                insertBefore = existing;
                break;
            }
        }

        if (insertBefore is null)
        {
            list.Append(section);
        }
        else
        {
            list.InsertBefore(section, insertBefore);
        }
    }

    private static int ParseAfterSlide(JsonNode? node)
    {
        if (node is JsonValue value)
        {
            if (Units.TryNumber(value, out var number) && number == Math.Floor(number) && number >= 0)
            {
                return (int)number;
            }

            if (value.TryGetValue<string>(out var raw) &&
                int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) && parsed >= 0)
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"afterSlide is not a valid 0-based slide index: {node?.ToJsonString() ?? "null"}",
            "Use 0 (start the section before slide 1) or a positive integer N (start it after slide N).");
    }

    private static string NewSectionId() => "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";

    private static AiofficeException Corrupt(string detail) => new(
        ErrorCodes.FormatCorrupt,
        $"The presentation is malformed: {detail}.",
        "Re-export the file from PowerPoint/Keynote, or restore a snapshot.");
}
