using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// v1.13 slide footer / slide-number / date placeholders. Three OOXML
/// placeholder shapes share the slide footer strip:
/// <list type="bullet">
/// <item><c>ph type="ftr"</c> — a plain text run (the footer caption).</item>
/// <item><c>ph type="sldNum"</c> — a <c>slidenum</c> <c>a:fld</c> field (auto page number).</item>
/// <item><c>ph type="dt"</c> — a <c>datetime</c> <c>a:fld</c> field (auto date, or a fixed string).</item>
/// </list>
/// Each shape's visibility is mirrored in the host's <c>p:hf</c> (header/footer)
/// element. The slide-level form (<c>set /slide[i] {...}</c>) edits one slide;
/// the deck-wide form (<c>set / {...}</c> or <c>set /master[1] {...}</c>) stamps
/// the master, its layouts and every slide in one op so an agent can number a
/// whole deck at once. Title slides hide the footer by convention unless
/// <c>skipTitle:false</c> is passed. Fields carry a cached <c>a:t</c> so they
/// render (SVG/LibreOffice) and open live in PowerPoint.
/// </summary>
internal static class PptxFooters
{
    /// <summary>The footer/number/date props this module owns (slide- and deck-level).</summary>
    public static readonly IReadOnlyList<string> PropKeys = ["footer", "slideNumber", "date"];

    /// <summary>The deck-wide form adds skipTitle to the three placeholder props.</summary>
    private static readonly IReadOnlyList<string> DeckPropKeys = ["footer", "slideNumber", "date", "skipTitle"];

    // Standard PowerPoint 16:9 footer-strip geometry (EMU). The three shapes
    // share the bottom band: date on the left, footer centered, number on the
    // right. Coordinates scale-independently (existing decks keep their own).
    private static readonly (long X, long Y, long Cx, long Cy) DateBox = (838_200, 6_356_350, 2_743_200, 365_125);
    private static readonly (long X, long Y, long Cx, long Cy) FooterBox = (4_038_600, 6_356_350, 4_114_800, 365_125);
    private static readonly (long X, long Y, long Cx, long Cy) NumberBox = (8_685_213, 6_356_350, 2_667_000, 365_125);

    /// <summary>True when these props belong to this module (so callers can split them out).</summary>
    public static bool Handles(JsonObject props) =>
        PropKeys.Any(k => props.ContainsKey(k));

    // ---- slide-level -------------------------------------------------------

    /// <summary>
    /// set /slide[i] {footer, slideNumber, date}: add/update the slide's footer,
    /// slide-number and date placeholder shapes. (A p:sld has no p:hf — the schema
    /// keeps header/footer visibility on the master/layout; on a single slide the
    /// presence of the placeholder shape *is* its visibility.) Strings set the
    /// footer caption / a fixed date; <c>false</c> removes the shape; <c>true</c>
    /// adds the auto field. Only the keys present are touched.
    /// </summary>
    public static void ApplyToSlide(SlidePart slidePart, JsonObject props)
    {
        var slide = slidePart.Slide ?? throw Corrupt("the slide has no p:sld");
        var tree = (slide.CommonSlideData ??= new P.CommonSlideData(PptxFactory.EmptyShapeTree())).ShapeTree
            ??= PptxFactory.EmptyShapeTree();

        if (props.TryGetPropertyValue("footer", out var footerNode))
        {
            ApplyFooter(tree, footerNode);
        }

        if (props.TryGetPropertyValue("slideNumber", out var numberNode))
        {
            ApplyNumber(tree, numberNode);
        }

        if (props.TryGetPropertyValue("date", out var dateNode))
        {
            ApplyDate(tree, dateNode);
        }
    }

    // ---- deck-wide ---------------------------------------------------------

    /// <summary>
    /// set / {footer, slideNumber, date, skipTitle?} or set /master[m] {...}:
    /// stamps the footer/number/date onto the addressed master, all its layouts
    /// and every slide, so the whole deck is numbered in one op. Title-layout
    /// slides hide the footer by convention unless <c>skipTitle:false</c>.
    /// Returns the canonical path ("/" or "/master[m]").
    /// </summary>
    public static string ApplyDeckWide(PresentationPart presentation, PptxAddress address, JsonObject props, out JsonObject rest)
    {
        var ours = Split(props, out rest);

        foreach (var (key, _) in ours)
        {
            if (!DeckPropKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Prop '{key}' does not apply to the deck footer.",
                    "Deck footer props: footer (text/false), slideNumber (bool), date (true/false/\"fixed text\"), skipTitle (bool).",
                    candidates: DeckPropKeys);
            }
        }

        var skipTitle = true;
        if (ours.TryGetPropertyValue("skipTitle", out var skipNode))
        {
            skipTitle = AsBool("skipTitle", skipNode);
        }

        // The master(s) carry the placeholders the slides inherit; when a deck-wide
        // op is addressed at "/", every master is stamped, at /master[m] just that one.
        var masters = address.IsMaster
            ? [(address.MasterIndex, PptxDoc.ResolveMaster(presentation, address.MasterIndex, address.Raw))]
            : PptxDoc.Masters(presentation);

        foreach (var (_, masterPart) in masters)
        {
            var master = masterPart.SlideMaster ?? throw Corrupt("the master has no p:sldMaster");
            StampMasterHost(
                (master.CommonSlideData ??= new P.CommonSlideData(PptxFactory.EmptyShapeTree())).ShapeTree
                    ??= PptxFactory.EmptyShapeTree(),
                master.HeaderFooter ??= new P.HeaderFooter(),
                ours,
                hideFooter: false);
            PruneEmptyHeaderFooter(master);

            foreach (var (_, layoutPart) in PptxDoc.Layouts(masterPart))
            {
                var layout = layoutPart.SlideLayout ?? throw Corrupt("the layout has no p:sldLayout");
                var isTitle = IsTitleLayout(layoutPart);
                StampMasterHost(
                    (layout.CommonSlideData ??= new P.CommonSlideData(PptxFactory.EmptyShapeTree())).ShapeTree
                        ??= PptxFactory.EmptyShapeTree(),
                    layout.HeaderFooter ??= new P.HeaderFooter(),
                    ours,
                    hideFooter: skipTitle && isTitle);
                PruneEmptyHeaderFooter(layout);
            }
        }

        // Every slide gets the concrete shapes so the placeholders actually render
        // (PowerPoint inherits from the layout, but our SVG renderer reads the slide).
        var slidesToStamp = address.IsMaster
            ? PptxDoc.Slides(presentation).Where(s => UsesMaster(presentation, s.Part, address.MasterIndex))
            : PptxDoc.Slides(presentation);

        foreach (var (_, slidePart) in slidesToStamp)
        {
            var slide = slidePart.Slide ?? throw Corrupt("the slide has no p:sld");
            var hideFooter = skipTitle && IsTitleSlide(slidePart);
            StampSlide(
                (slide.CommonSlideData ??= new P.CommonSlideData(PptxFactory.EmptyShapeTree())).ShapeTree
                    ??= PptxFactory.EmptyShapeTree(),
                ours,
                hideFooter);
        }

        return address.IsMaster ? address.CanonicalMasterPath : "/";
    }

    /// <summary>
    /// Stamps one master/layout host with the deck-wide footer props, mirroring each
    /// shape's visibility into the host's p:hf (only masters/layouts carry p:hf).
    /// </summary>
    private static void StampMasterHost(P.ShapeTree tree, P.HeaderFooter hf, JsonObject ours, bool hideFooter)
    {
        if (ours.TryGetPropertyValue("footer", out var footerNode))
        {
            hf.Footer = ApplyFooter(tree, hideFooter ? (JsonNode?)false : footerNode);
        }

        if (ours.TryGetPropertyValue("slideNumber", out var numberNode))
        {
            hf.SlideNumber = ApplyNumber(tree, numberNode);
        }

        if (ours.TryGetPropertyValue("date", out var dateNode))
        {
            hf.DateTime = ApplyDate(tree, dateNode);
        }
    }

    /// <summary>Stamps one slide's shape tree with the deck-wide footer props (a slide has no p:hf).</summary>
    private static void StampSlide(P.ShapeTree tree, JsonObject ours, bool hideFooter)
    {
        if (ours.TryGetPropertyValue("footer", out var footerNode))
        {
            ApplyFooter(tree, hideFooter ? (JsonNode?)false : footerNode);
        }

        if (ours.TryGetPropertyValue("slideNumber", out var numberNode))
        {
            ApplyNumber(tree, numberNode);
        }

        if (ours.TryGetPropertyValue("date", out var dateNode))
        {
            ApplyDate(tree, dateNode);
        }
    }

    /// <summary>Peels footer/slideNumber/date/skipTitle out of a props object (the original is not mutated).</summary>
    public static JsonObject Split(JsonObject props, out JsonObject rest)
    {
        var taken = new JsonObject();
        rest = new JsonObject();
        foreach (var (key, value) in props)
        {
            if (DeckPropKeys.Contains(key, StringComparer.Ordinal))
            {
                taken[key] = value?.DeepClone();
            }
            else
            {
                rest[key] = value?.DeepClone();
            }
        }

        return taken;
    }

    // ---- the three placeholders -------------------------------------------

    /// <summary>Add/update or remove the footer placeholder shape; returns whether it is now visible.</summary>
    private static bool ApplyFooter(P.ShapeTree tree, JsonNode? value)
    {
        if (IsFalse(value))
        {
            RemovePlaceholder(tree, P.PlaceholderValues.Footer);
            return false;
        }

        var text = J.ScalarText(value ?? string.Empty);
        var shape = EnsurePlaceholder(tree, P.PlaceholderValues.Footer, "Footer Placeholder", FooterBox);
        SetParagraph(shape, new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text(text)));
        return true;
    }

    /// <summary>Add or remove the slide-number field placeholder; returns whether it is now visible.</summary>
    private static bool ApplyNumber(P.ShapeTree tree, JsonNode? value)
    {
        if (!AsBool("slideNumber", value))
        {
            RemovePlaceholder(tree, P.PlaceholderValues.SlideNumber);
            return false;
        }

        var shape = EnsurePlaceholder(tree, P.PlaceholderValues.SlideNumber, "Slide Number Placeholder", NumberBox);
        SetParagraph(shape, NumberField());
        return true;
    }

    /// <summary>Add/update or remove the date placeholder (auto field or fixed caption); returns visibility.</summary>
    private static bool ApplyDate(P.ShapeTree tree, JsonNode? value)
    {
        if (IsFalse(value))
        {
            RemovePlaceholder(tree, P.PlaceholderValues.DateAndTime);
            return false;
        }

        var shape = EnsurePlaceholder(tree, P.PlaceholderValues.DateAndTime, "Date Placeholder", DateBox);

        // A bare true → the live auto-updating date field; a string → a fixed
        // caption (a plain run, not a field) the way PowerPoint stores "fixed" dates.
        if (value is JsonValue v && v.TryGetValue<string>(out var fixedText) && fixedText.Length > 0)
        {
            SetParagraph(shape, new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text(fixedText)));
        }
        else
        {
            SetParagraph(shape, DateField());
        }

        return true;
    }

    /// <summary>A slidenum field carrying a cached "1" so it renders before PowerPoint recomputes it.</summary>
    private static A.Field NumberField() => new(
        new A.RunProperties { Language = "en-US" },
        new A.Text("1"))
    {
        Id = Guid.NewGuid().ToString("B").ToUpperInvariant(),
        Type = "slidenum",
    };

    /// <summary>A datetime1 field carrying today's cached date so it renders before PowerPoint recomputes it.</summary>
    private static A.Field DateField() => new(
        new A.RunProperties { Language = "en-US" },
        new A.Text(DateTime.Now.ToString("M/d/yyyy", CultureInfo.InvariantCulture)))
    {
        Id = Guid.NewGuid().ToString("B").ToUpperInvariant(),
        Type = "datetime1",
    };

    // ---- shape helpers -----------------------------------------------------

    private static P.Shape EnsurePlaceholder(P.ShapeTree tree, P.PlaceholderValues type, string name, (long X, long Y, long Cx, long Cy) box)
    {
        var shape = FindPlaceholder(tree, type);
        if (shape is not null)
        {
            return shape;
        }

        shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = NextShapeId(tree), Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = type })),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = box.X, Y = box.Y },
                    new A.Extents { Cx = box.Cx, Cy = box.Cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph()));
        tree.Append(shape);
        return shape;
    }

    /// <summary>Replaces the placeholder's single paragraph with one carrying the given run/field.</summary>
    private static void SetParagraph(P.Shape shape, OpenXmlElement content)
    {
        var body = shape.TextBody ??= new P.TextBody(new A.BodyProperties(), new A.ListStyle());
        foreach (var paragraph in body.Elements<A.Paragraph>().ToList())
        {
            paragraph.Remove();
        }

        body.Append(new A.Paragraph(content));
    }

    private static void RemovePlaceholder(P.ShapeTree tree, P.PlaceholderValues type)
    {
        FindPlaceholder(tree, type)?.Remove();
    }

    private static P.Shape? FindPlaceholder(P.ShapeTree tree, P.PlaceholderValues type) =>
        tree.Elements<P.Shape>().FirstOrDefault(s =>
            s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == type);

    private static uint NextShapeId(P.ShapeTree tree)
    {
        uint max = 1;
        foreach (var props in tree.Descendants<P.NonVisualDrawingProperties>())
        {
            if (props.Id?.Value is { } id && id > max)
            {
                max = id;
            }
        }

        return max + 1;
    }

    /// <summary>Drops a header/footer element whose every flag is off (an all-false p:hf is noise).</summary>
    private static void PruneEmptyHeaderFooter(P.SlideMaster master)
    {
        if (master.HeaderFooter is { } hf && IsEmpty(hf))
        {
            hf.Remove();
        }
    }

    private static void PruneEmptyHeaderFooter(P.SlideLayout layout)
    {
        if (layout.HeaderFooter is { } hf && IsEmpty(hf))
        {
            hf.Remove();
        }
    }

    private static bool IsEmpty(P.HeaderFooter hf) =>
        hf.Footer?.Value is not true && hf.SlideNumber?.Value is not true && hf.DateTime?.Value is not true
        && hf.Header?.Value is not true;

    // ---- get / report ------------------------------------------------------

    /// <summary>
    /// The footer/slideNumber/date state of one slide, for the get projection. A
    /// slide carries no p:hf, so visibility is read off the presence of each
    /// placeholder shape on the slide itself.
    /// </summary>
    public static object SlideState(SlidePart slidePart)
    {
        var tree = slidePart.Slide?.CommonSlideData?.ShapeTree;
        return new
        {
            Footer = tree is null ? null : PlaceholderText(tree, P.PlaceholderValues.Footer),
            SlideNumber = tree is not null && FindPlaceholder(tree, P.PlaceholderValues.SlideNumber) is not null,
            Date = tree is not null && FindPlaceholder(tree, P.PlaceholderValues.DateAndTime) is not null,
            DateText = tree is null ? null : FixedDateText(tree),
        };
    }

    /// <summary>The deck-wide footer/slideNumber/date state read off the first master's p:hf, for get /.</summary>
    public static object DeckState(PresentationPart presentation)
    {
        var master = PptxDoc.Masters(presentation).Select(m => m.Part.SlideMaster).FirstOrDefault();
        var tree = master?.CommonSlideData?.ShapeTree;
        var hf = master?.HeaderFooter;
        return new
        {
            Footer = tree is null ? null : PlaceholderText(tree, P.PlaceholderValues.Footer),
            SlideNumber = hf?.SlideNumber?.Value ?? false,
            Date = hf?.DateTime?.Value ?? false,
            DateText = tree is null ? null : FixedDateText(tree),
        };
    }

    private static string? PlaceholderText(P.ShapeTree tree, P.PlaceholderValues type)
    {
        var shape = FindPlaceholder(tree, type);
        if (shape?.TextBody is not { } body)
        {
            return null;
        }

        var text = string.Join("\n", body.Elements<A.Paragraph>().Select(PptxDoc.ParagraphText));
        return text.Length == 0 ? null : text;
    }

    /// <summary>The fixed date caption (a plain run), or null when the date is an auto field / absent.</summary>
    private static string? FixedDateText(P.ShapeTree tree)
    {
        var shape = FindPlaceholder(tree, P.PlaceholderValues.DateAndTime);
        if (shape?.TextBody is not { } body)
        {
            return null;
        }

        // A fixed date is a plain a:r run; an auto date is an a:fld field. Only the
        // former is a user caption worth reporting.
        var hasField = body.Descendants<A.Field>().Any();
        if (hasField)
        {
            return null;
        }

        var text = string.Join("\n", body.Elements<A.Paragraph>().Select(PptxDoc.ParagraphText));
        return text.Length == 0 ? null : text;
    }

    // ---- title detection ---------------------------------------------------

    private static bool IsTitleSlide(SlidePart slidePart) =>
        slidePart.SlideLayoutPart is { } layout && IsTitleLayout(layout);

    private static bool IsTitleLayout(SlideLayoutPart layoutPart)
    {
        var type = layoutPart.SlideLayout?.Type?.Value;
        return type == P.SlideLayoutValues.Title || type == P.SlideLayoutValues.TitleOnly;
    }

    private static bool UsesMaster(PresentationPart presentation, SlidePart slidePart, int masterIndex)
    {
        var layout = slidePart.SlideLayoutPart;
        if (layout is null)
        {
            return false;
        }

        var masters = PptxDoc.Masters(presentation);
        var target = masters.FirstOrDefault(m => m.Index == masterIndex).Part;
        return target is not null && PptxDoc.Layouts(target).Any(l => l.Part.Uri == layout.Uri);
    }

    // ---- value parsing -----------------------------------------------------

    private static bool IsFalse(JsonNode? node) =>
        node is JsonValue v && v.TryGetValue<bool>(out var flag) && !flag;

    private static bool AsBool(string key, JsonNode? node)
    {
        if (node is JsonValue value)
        {
            if (value.TryGetValue<bool>(out var flag))
            {
                return flag;
            }

            if (value.TryGetValue<string>(out var text) && bool.TryParse(text, out var parsed))
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"{key} must be true or false; got {node?.ToJsonString() ?? "null"}.",
            "Pass a boolean, e.g. {\"slideNumber\":true}.");
    }

    private static AiofficeException Corrupt(string detail) => new(
        ErrorCodes.FormatCorrupt,
        $"The presentation is malformed: {detail}.",
        "Re-export the file from PowerPoint/Keynote, or restore a snapshot.");
}
