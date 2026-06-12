using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Builds the minimal valid .pptx package: one slide master, one blank layout,
/// one theme and one blank slide (16:9). Every part is required by the OOXML
/// schema for a deck that opens in PowerPoint/Keynote; nothing here is optional.
/// </summary>
internal static class PptxFactory
{
    public const int SlideWidthEmu = 12_192_000; // 16:9
    public const int SlideHeightEmu = 6_858_000;

    public static void CreateMinimal(string path, string? title)
    {
        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();
        presentationPart.Presentation = new P.Presentation(
            new P.SlideMasterIdList(new P.SlideMasterId { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" }),
            new P.SlideIdList(new P.SlideId { Id = (UInt32Value)256U, RelationshipId = "rId2" }),
            new P.SlideSize { Cx = SlideWidthEmu, Cy = SlideHeightEmu },
            new P.NotesSize { Cx = 6_858_000L, Cy = 9_144_000L },
            new P.DefaultTextStyle());

        var slidePart = presentationPart.AddNewPart<SlidePart>("rId2");
        slidePart.Slide = BuildBlankSlide();

        var layoutPart = slidePart.AddNewPart<SlideLayoutPart>("rId1");
        layoutPart.SlideLayout = BuildBlankLayout();

        var masterPart = layoutPart.AddNewPart<SlideMasterPart>("rId1");
        masterPart.SlideMaster = BuildMaster();

        var themePart = masterPart.AddNewPart<ThemePart>("rId2");
        themePart.Theme = BuildTheme();

        masterPart.AddPart(layoutPart, "rId1");
        presentationPart.AddPart(masterPart, "rId1");

        if (!string.IsNullOrWhiteSpace(title))
        {
            PptxEditor.AddTitleShape(slidePart, title);
        }
    }

    internal static P.Slide BuildBlankSlide() => new(
        new P.CommonSlideData(EmptyShapeTree()),
        new P.ColorMapOverride(new A.MasterColorMapping()));

    internal static P.ShapeTree EmptyShapeTree() => new(
        new P.NonVisualGroupShapeProperties(
            new P.NonVisualDrawingProperties { Id = (UInt32Value)1U, Name = string.Empty },
            new P.NonVisualGroupShapeDrawingProperties(),
            new P.ApplicationNonVisualDrawingProperties()),
        new P.GroupShapeProperties(new A.TransformGroup()));

    private static P.SlideLayout BuildBlankLayout() => new(
        new P.CommonSlideData(EmptyShapeTree()) { Name = "Blank" },
        new P.ColorMapOverride(new A.MasterColorMapping()))
    {
        Type = P.SlideLayoutValues.Blank,
    };

    private static P.SlideMaster BuildMaster() => new(
        new P.CommonSlideData(EmptyShapeTree()),
        new P.ColorMap
        {
            Background1 = A.ColorSchemeIndexValues.Light1,
            Text1 = A.ColorSchemeIndexValues.Dark1,
            Background2 = A.ColorSchemeIndexValues.Light2,
            Text2 = A.ColorSchemeIndexValues.Dark2,
            Accent1 = A.ColorSchemeIndexValues.Accent1,
            Accent2 = A.ColorSchemeIndexValues.Accent2,
            Accent3 = A.ColorSchemeIndexValues.Accent3,
            Accent4 = A.ColorSchemeIndexValues.Accent4,
            Accent5 = A.ColorSchemeIndexValues.Accent5,
            Accent6 = A.ColorSchemeIndexValues.Accent6,
            Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
            FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
        },
        new P.SlideLayoutIdList(new P.SlideLayoutId { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }));

    private static A.Theme BuildTheme()
    {
        static A.SolidFill PlaceholderFill() => new(new A.SchemeColor { Val = A.SchemeColorValues.PhColor });

        static A.Outline Line() => new(PlaceholderFill()) { Width = 9525 };

        static A.EffectStyle Effect() => new(new A.EffectList());

        var colorScheme = new A.ColorScheme(
            new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
            new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
            new A.Dark2Color(new A.RgbColorModelHex { Val = "44546A" }),
            new A.Light2Color(new A.RgbColorModelHex { Val = "E7E6E6" }),
            new A.Accent1Color(new A.RgbColorModelHex { Val = "4472C4" }),
            new A.Accent2Color(new A.RgbColorModelHex { Val = "ED7D31" }),
            new A.Accent3Color(new A.RgbColorModelHex { Val = "A5A5A5" }),
            new A.Accent4Color(new A.RgbColorModelHex { Val = "FFC000" }),
            new A.Accent5Color(new A.RgbColorModelHex { Val = "5B9BD5" }),
            new A.Accent6Color(new A.RgbColorModelHex { Val = "70AD47" }),
            new A.Hyperlink(new A.RgbColorModelHex { Val = "0563C1" }),
            new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "954F72" }))
        {
            Name = "AIOffice",
        };

        var fontScheme = new A.FontScheme(
            new A.MajorFont(
                new A.LatinFont { Typeface = "Calibri Light" },
                new A.EastAsianFont { Typeface = string.Empty },
                new A.ComplexScriptFont { Typeface = string.Empty }),
            new A.MinorFont(
                new A.LatinFont { Typeface = "Calibri" },
                new A.EastAsianFont { Typeface = string.Empty },
                new A.ComplexScriptFont { Typeface = string.Empty }))
        {
            Name = "AIOffice",
        };

        var formatScheme = new A.FormatScheme(
            new A.FillStyleList(PlaceholderFill(), PlaceholderFill(), PlaceholderFill()),
            new A.LineStyleList(Line(), Line(), Line()),
            new A.EffectStyleList(Effect(), Effect(), Effect()),
            new A.BackgroundFillStyleList(PlaceholderFill(), PlaceholderFill(), PlaceholderFill()))
        {
            Name = "AIOffice",
        };

        return new A.Theme(new A.ThemeElements(colorScheme, fontScheme, formatScheme))
        {
            Name = "AIOffice",
        };
    }
}
