using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class ThemeTests : WordTestBase
{
    private JsonNode Theme(string file) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/theme" })));

    [Fact]
    public void Set_theme_creates_part_on_demand_and_edits_accent_colors()
    {
        var file = CreateDoc(title: "Theme");

        // A freshly created docx has no theme part; set /theme must create one.
        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.Null(doc.MainDocumentPart!.ThemePart);
        }

        Edit(file, """[{"op":"set","path":"/theme","props":{"accent1":"38BDF8","accent2":"0EA5E9","dk1":"111827"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var scheme = doc.MainDocumentPart!.ThemePart!.Theme!.ThemeElements!.ColorScheme!;
            Assert.Equal("38BDF8", scheme.Accent1Color!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal("0EA5E9", scheme.Accent2Color!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            // dk1 default is a sysColor; setting it replaces with an explicit RGB.
            Assert.Equal("111827", scheme.Dark1Color!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
        }

        var theme = Theme(file);
        Assert.Equal("38BDF8", theme["properties"]!["accent1"]!.GetValue<string>());
        Assert.Equal("0EA5E9", theme["properties"]!["accent2"]!.GetValue<string>());
        Assert.Equal("111827", theme["properties"]!["dk1"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Set_theme_edits_major_and_minor_fonts()
    {
        var file = CreateDoc(title: "ThemeFont");
        Edit(file, """[{"op":"set","path":"/theme","props":{"majorFont":"Georgia","minorFont":"Verdana"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var fonts = doc.MainDocumentPart!.ThemePart!.Theme!.ThemeElements!.FontScheme!;
            Assert.Equal("Georgia", fonts.MajorFont!.LatinFont!.Typeface!.Value);
            Assert.Equal("Verdana", fonts.MinorFont!.LatinFont!.Typeface!.Value);
        }

        var theme = Theme(file);
        Assert.Equal("Georgia", theme["properties"]!["majorFont"]!.GetValue<string>());
        Assert.Equal("Verdana", theme["properties"]!["minorFont"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void A_run_referencing_accent1_resolves_to_the_edited_theme_value()
    {
        var file = CreateDoc(title: "ThemeRef");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"Accented"}}]""");

        // Wire a run that references the theme's accent1 slot by name (themeColor),
        // the way a theme-aware style/run does. The rendered color follows the theme.
        using (var doc = WordprocessingDocument.Open(file, isEditable: true))
        {
            var run = doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().Last();
            (run.RunProperties ??= new RunProperties()).Color = new Color
            {
                Val = "4472C4",
                ThemeColor = ThemeColorValues.Accent1,
            };
        }

        Edit(file, """[{"op":"set","path":"/theme","props":{"accent1":"38BDF8"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var color = doc.MainDocumentPart!.Document!.Body!.Descendants<Color>()
                .Single(c => c.ThemeColor is not null);
            Assert.Equal(ThemeColorValues.Accent1, color.ThemeColor!.Value);
            var accent1 = doc.MainDocumentPart!.ThemePart!.Theme!.ThemeElements!.ColorScheme!
                .Accent1Color!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value;
            // The run still references accent1 by name; that name now resolves to the new hex.
            Assert.Equal("38BDF8", accent1);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_theme_without_a_part_reports_nulls()
    {
        var file = CreateDoc(title: "NoTheme");
        var theme = Theme(file);
        Assert.Equal("theme", theme["type"]!.GetValue<string>());
        Assert.Null(theme["properties"]!["accent1"]);
    }

    [Fact]
    public void Unknown_theme_slot_is_unsupported_feature()
    {
        var file = CreateDoc(title: "BadSlot");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/theme","props":{"accent7":"123456"}}]"""));
        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("accent1", ex.Candidates!);
    }

    [Fact]
    public void Structure_view_notes_theme_colors_when_present()
    {
        var file = CreateDoc(title: "StructTheme");
        Edit(file, """[{"op":"set","path":"/theme","props":{"accent1":"38BDF8"}}]""");

        var structure = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));
        Assert.NotNull(structure["theme"]);
        Assert.Equal("38BDF8", structure["theme"]!["accent1"]!.GetValue<string>());
    }
}
