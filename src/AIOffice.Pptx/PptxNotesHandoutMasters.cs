using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// v1.7.0 notes-master and handout-master editing (M6 added slide-master editing;
/// this extends the same idea to the two remaining master parts):
/// <list type="bullet">
/// <item>set /notesMaster {background, bodyFont} — the shared look of every slide's notes page.</item>
/// <item>set /handoutMaster {background, headerFooter:{header,footer,date,pageNumber}, slidesPerPage} —
/// the print-handout layout.</item>
/// </list>
/// Both parts are created (with a minimal valid body) on first edit and reported
/// by get and by <c>read --view structure</c>. Every edit keeps the package
/// validator-clean.
/// </summary>
internal static class PptxNotesHandoutMasters
{
    private static readonly IReadOnlyList<string> NotesMasterProps = ["background", "bodyFont"];

    private static readonly IReadOnlyList<string> HandoutMasterProps = ["background", "headerFooter", "slidesPerPage"];

    private static readonly IReadOnlyList<int> SlidesPerPageChoices = [1, 2, 3, 4, 6, 9];

    private static readonly IReadOnlyList<string> HeaderFooterKeys = ["header", "footer", "date", "pageNumber"];

    // ---- set ---------------------------------------------------------------

    /// <summary>Routes a set op on /notesMaster or /handoutMaster to the right editor.</summary>
    public static string Set(PresentationPart presentation, PptxAddress address, JsonObject props) =>
        address.IsNotesMaster
            ? SetNotesMaster(presentation, props)
            : SetHandoutMaster(presentation, props);

    private static string SetNotesMaster(PresentationPart presentation, JsonObject props)
    {
        var masterPart = EnsureNotesMaster(presentation);
        var master = masterPart.NotesMaster ?? throw Corrupt("the notes master has no p:notesMaster");

        foreach (var (key, value) in props)
        {
            switch (key)
            {
                case "background":
                    PptxEditor.SetBackground(
                        master.CommonSlideData ??= new P.CommonSlideData(PptxFactory.EmptyShapeTree()),
                        value);
                    break;
                case "bodyFont":
                    SetThemeBodyFont(masterPart, "notes master", value);
                    break;
                default:
                    throw UnknownProp(key, "notes master", NotesMasterProps);
            }
        }

        return "/notesMaster";
    }

    private static string SetHandoutMaster(PresentationPart presentation, JsonObject props)
    {
        var masterPart = EnsureHandoutMaster(presentation);
        var master = masterPart.HandoutMaster ?? throw Corrupt("the handout master has no p:handoutMaster");

        foreach (var (key, value) in props)
        {
            switch (key)
            {
                case "background":
                    PptxEditor.SetBackground(
                        master.CommonSlideData ??= new P.CommonSlideData(PptxFactory.EmptyShapeTree()),
                        value);
                    break;
                case "headerFooter":
                    SetHeaderFooter(master, value);
                    break;
                case "slidesPerPage":
                    SetSlidesPerPage(presentation, value);
                    break;
                default:
                    throw UnknownProp(key, "handout master", HandoutMasterProps);
            }
        }

        return "/handoutMaster";
    }

    /// <summary>
    /// Sets the header/footer text (in the handout master's header/footer placeholder
    /// shapes) plus the date/pageNumber visibility flags (in p:hf). Strings create or
    /// rewrite the matching placeholder; booleans toggle the hf visibility.
    /// </summary>
    private static void SetHeaderFooter(P.HandoutMaster master, JsonNode? value)
    {
        if (value is not JsonObject hf)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "headerFooter must be an object.",
                "Pass {\"header\":\"...\",\"footer\":\"...\",\"date\":true,\"pageNumber\":true} — " +
                "header/footer are text strings; date/pageNumber are booleans.");
        }

        foreach (var (key, _) in hf)
        {
            if (!HeaderFooterKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown headerFooter key '{key}'.",
                    "headerFooter keys: header, footer (text), date, pageNumber (booleans).",
                    candidates: HeaderFooterKeys);
            }
        }

        var tree = (master.CommonSlideData ??= new P.CommonSlideData(PptxFactory.EmptyShapeTree())).ShapeTree
            ??= PptxFactory.EmptyShapeTree();
        var headerFooter = master.HeaderFooter ??= new P.HeaderFooter();

        if (hf.TryGetPropertyValue("header", out var headerNode) && headerNode is not null)
        {
            SetPlaceholderText(tree, P.PlaceholderValues.Header, "Header Placeholder", J.ScalarText(headerNode));
            headerFooter.Header = true;
        }

        if (hf.TryGetPropertyValue("footer", out var footerNode) && footerNode is not null)
        {
            SetPlaceholderText(tree, P.PlaceholderValues.Footer, "Footer Placeholder", J.ScalarText(footerNode));
            headerFooter.Footer = true;
        }

        if (hf.TryGetPropertyValue("date", out var dateNode))
        {
            headerFooter.DateTime = AsBool("date", dateNode);
        }

        if (hf.TryGetPropertyValue("pageNumber", out var pageNode))
        {
            headerFooter.SlideNumber = AsBool("pageNumber", pageNode);
        }
    }

    /// <summary>
    /// Records the handout slides-per-page in presProps.xml (p:prnPr/@prnWhat), the
    /// part PowerPoint reads for its print-handout layout. 1/2/3/4/6/9 map to the
    /// six handout PrintOutputValues.
    /// </summary>
    private static void SetSlidesPerPage(PresentationPart presentation, JsonNode? value)
    {
        if (!TryInteger(value, out var perPage) || !SlidesPerPageChoices.Contains(perPage))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"slidesPerPage must be one of 1, 2, 3, 4, 6, 9; got {value?.ToJsonString() ?? "null"}.",
                "PowerPoint lays out handouts at 1, 2, 3, 4, 6 or 9 slides per page.",
                candidates: [.. SlidesPerPageChoices.Select(n => n.ToString(CultureInfo.InvariantCulture))]);
        }

        var propsPart = presentation.PresentationPropertiesPart ?? presentation.AddNewPart<PresentationPropertiesPart>();
        var props = propsPart.PresentationProperties ??= new P.PresentationProperties();
        var printing = props.PrintingProperties ??= new P.PrintingProperties();
        printing.PrintWhat = perPage switch
        {
            1 => P.PrintOutputValues.Handouts1,
            2 => P.PrintOutputValues.Handouts2,
            3 => P.PrintOutputValues.Handouts3,
            4 => P.PrintOutputValues.Handouts4,
            6 => P.PrintOutputValues.Handouts6,
            _ => P.PrintOutputValues.Handouts9,
        };
    }

    /// <summary>Sets the deck-wide minor (body) font typeface on the notes master's theme part.</summary>
    private static void SetThemeBodyFont(NotesMasterPart masterPart, string label, JsonNode? value)
    {
        var typeface = J.ScalarText(value ?? string.Empty).Trim();
        if (typeface.Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "bodyFont must be a non-empty font name.",
                "Pass a typeface, e.g. {\"bodyFont\":\"Calibri\"}.");
        }

        var minorFont = masterPart.ThemePart?.Theme?.ThemeElements?.FontScheme?.MinorFont;
        if (minorFont is null)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"The {label} has no theme font scheme to retypeface.",
                "Re-export the deck from PowerPoint to add a theme, or set fonts on individual shapes instead.");
        }

        (minorFont.LatinFont ??= new A.LatinFont()).Typeface = typeface;
    }

    /// <summary>Sets a header/footer placeholder shape's text, creating the placeholder when absent.</summary>
    private static void SetPlaceholderText(P.ShapeTree tree, P.PlaceholderValues type, string name, string text)
    {
        var shape = FindPlaceholder(tree, type);
        if (shape is null)
        {
            shape = new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = NextShapeId(tree), Name = name },
                    new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = type })),
                new P.ShapeProperties(),
                new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph()));
            tree.Append(shape);
        }

        var body = shape.TextBody ??= new P.TextBody(new A.BodyProperties(), new A.ListStyle());
        foreach (var paragraph in body.Elements<A.Paragraph>().ToList())
        {
            paragraph.Remove();
        }

        body.Append(new A.Paragraph(new A.Run(new A.Text(text))));
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

    // ---- get / report ------------------------------------------------------

    /// <summary>The get projection for /notesMaster (state report; creates nothing).</summary>
    public static object NotesMasterDetail(PresentationPart presentation)
    {
        var masterPart = presentation.NotesMasterPart;
        var master = masterPart?.NotesMaster;
        return new
        {
            Path = "/notesMaster",
            Kind = "notesMaster",
            Present = masterPart is not null,
            Background = master?.CommonSlideData?.Background is { } bg
                ? bg.BackgroundProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant()
                : null,
            BodyFont = masterPart?.ThemePart?.Theme?.ThemeElements?.FontScheme?.MinorFont?.LatinFont?.Typeface?.Value,
            ShapeCount = master?.CommonSlideData?.ShapeTree is { } tree ? PptxDoc.Shapes(tree).Count : 0,
        };
    }

    /// <summary>The get projection for /handoutMaster (state report; creates nothing).</summary>
    public static object HandoutMasterDetail(PresentationPart presentation)
    {
        var masterPart = presentation.HandoutMasterPart;
        var master = masterPart?.HandoutMaster;
        var hf = master?.HeaderFooter;
        var tree = master?.CommonSlideData?.ShapeTree;
        return new
        {
            Path = "/handoutMaster",
            Kind = "handoutMaster",
            Present = masterPart is not null,
            Background = master?.CommonSlideData?.Background is { } bg
                ? bg.BackgroundProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant()
                : null,
            HeaderFooter = master is null ? null : new
            {
                Header = tree is null ? null : PlaceholderText(tree, P.PlaceholderValues.Header),
                Footer = tree is null ? null : PlaceholderText(tree, P.PlaceholderValues.Footer),
                Date = hf?.DateTime?.Value,
                PageNumber = hf?.SlideNumber?.Value,
            },
            SlidesPerPage = ReadSlidesPerPage(presentation),
        };
    }

    private static string? PlaceholderText(P.ShapeTree tree, P.PlaceholderValues type)
    {
        var shape = FindPlaceholder(tree, type);
        if (shape?.TextBody is not { } body)
        {
            return null;
        }

        var text = string.Join("\n", body.Elements<A.Paragraph>().Select(p =>
            string.Concat(p.Descendants<A.Text>().Select(t => t.Text))));
        return text.Length == 0 ? null : text;
    }

    private static int? ReadSlidesPerPage(PresentationPart presentation)
    {
        var printWhat = presentation.PresentationPropertiesPart?.PresentationProperties?.PrintingProperties?.PrintWhat?.Value;
        if (printWhat is null)
        {
            return null;
        }

        if (printWhat == P.PrintOutputValues.Handouts1)
        {
            return 1;
        }

        if (printWhat == P.PrintOutputValues.Handouts2)
        {
            return 2;
        }

        if (printWhat == P.PrintOutputValues.Handouts3)
        {
            return 3;
        }

        if (printWhat == P.PrintOutputValues.Handouts4)
        {
            return 4;
        }

        if (printWhat == P.PrintOutputValues.Handouts6)
        {
            return 6;
        }

        return printWhat == P.PrintOutputValues.Handouts9 ? 9 : (int?)null;
    }

    // ---- part creation -----------------------------------------------------

    /// <summary>The deck's notes master, created with a minimal valid body and registered on first use.</summary>
    private static NotesMasterPart EnsureNotesMaster(PresentationPart presentation)
    {
        var masterPart = presentation.NotesMasterPart;
        if (masterPart is null)
        {
            masterPart = presentation.AddNewPart<NotesMasterPart>();
            masterPart.NotesMaster = new P.NotesMaster(
                new P.CommonSlideData(PptxFactory.EmptyShapeTree()),
                PptxFactory.BuildColorMap(),
                new P.NotesStyle());
            masterPart.AddNewPart<ThemePart>().Theme = PptxFactory.BuildTheme();
        }

        var root = RequireRoot(presentation);
        root.NotesMasterIdList ??= new P.NotesMasterIdList();
        if (!root.NotesMasterIdList.Elements<P.NotesMasterId>().Any())
        {
            root.NotesMasterIdList.Append(new P.NotesMasterId { Id = presentation.GetIdOfPart(masterPart) });
        }

        return masterPart;
    }

    /// <summary>The deck's handout master, created with a minimal valid body and registered on first use.</summary>
    private static HandoutMasterPart EnsureHandoutMaster(PresentationPart presentation)
    {
        var masterPart = presentation.HandoutMasterPart;
        if (masterPart is null)
        {
            masterPart = presentation.AddNewPart<HandoutMasterPart>();
            masterPart.HandoutMaster = new P.HandoutMaster(
                new P.CommonSlideData(PptxFactory.EmptyShapeTree()),
                PptxFactory.BuildColorMap());
            masterPart.AddNewPart<ThemePart>().Theme = PptxFactory.BuildTheme();
        }

        var root = RequireRoot(presentation);
        root.HandoutMasterIdList ??= new P.HandoutMasterIdList();
        if (!root.HandoutMasterIdList.Elements<P.HandoutMasterId>().Any())
        {
            root.HandoutMasterIdList.Append(new P.HandoutMasterId { Id = presentation.GetIdOfPart(masterPart) });
        }

        return masterPart;
    }

    // ---- helpers -----------------------------------------------------------

    private static P.Presentation RequireRoot(PresentationPart presentation) =>
        presentation.Presentation ?? throw Corrupt("p:presentation is missing");

    private static bool TryInteger(JsonNode? node, out int number)
    {
        number = 0;
        if (node is not JsonValue value)
        {
            return false;
        }

        if (value.TryGetValue<int>(out number))
        {
            return true;
        }

        if (value.TryGetValue<double>(out var d) && d == Math.Floor(d))
        {
            number = (int)d;
            return true;
        }

        return value.TryGetValue<string>(out var text) &&
            int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out number);
    }

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
            $"headerFooter.{key} must be true or false; got {node?.ToJsonString() ?? "null"}.",
            "Pass a boolean, e.g. {\"date\":true,\"pageNumber\":false}.");
    }

    private static AiofficeException UnknownProp(string key, string what, IReadOnlyList<string> allowed) => new(
        ErrorCodes.InvalidArgs,
        $"Prop '{key}' does not apply to the {what}.",
        what == "notes master"
            ? "Notes-master props: background (a solid color), bodyFont (the notes body typeface)."
            : "Handout-master props: background, headerFooter ({header,footer,date,pageNumber}), slidesPerPage (1/2/3/4/6/9).",
        candidates: allowed);

    private static AiofficeException Corrupt(string detail) => new(
        ErrorCodes.FormatCorrupt,
        $"The presentation is malformed: {detail}.",
        "Re-export the file from PowerPoint/Keynote, or restore a snapshot.");
}
