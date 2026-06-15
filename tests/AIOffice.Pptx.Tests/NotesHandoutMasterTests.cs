using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// v1.7.0 notes-master and handout-master editing (extends M6 slide-master editing
/// to the two remaining master parts): set /notesMaster {background, bodyFont} and
/// set /handoutMaster {background, headerFooter, slidesPerPage}. Each part is created
/// on first edit, reopens validator-clean, and is reported by get and read --view
/// structure.
/// </summary>
public sealed class NotesHandoutMasterTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private string Create()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Cover"))));
        return _ws.PathOf("deck.pptx");
    }

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private JsonObject Get(string path) => TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));

    // ---- notes master --------------------------------------------------------

    [Fact]
    public void SetNotesMasterBackground_CreatesPartWithBgFill_ReopenVerified()
    {
        Create();
        var data = Edit(TestEnv.Op("set", "/notesMaster", props: TestEnv.Props(("background", "0F172A"))));
        Assert.Equal("/notesMaster", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var master = doc.PresentationPart!.NotesMasterPart!.NotesMaster!;
            Assert.Equal(
                "0F172A",
                master.CommonSlideData!.Background!.BackgroundProperties!
                    .GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);

            // The notes master is registered in p:notesMasterIdLst.
            Assert.Single(doc.PresentationPart.Presentation!.NotesMasterIdList!.Elements<P.NotesMasterId>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetNotesMasterBodyFont_WritesThemeMinorFont_ReopenVerified()
    {
        Create();
        Edit(TestEnv.Op("set", "/notesMaster", props: TestEnv.Props(("bodyFont", "Georgia"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var theme = doc.PresentationPart!.NotesMasterPart!.ThemePart!.Theme!;
            Assert.Equal("Georgia", theme.ThemeElements!.FontScheme!.MinorFont!.LatinFont!.Typeface!.Value);
        }

        var detail = Get("/notesMaster");
        Assert.True(detail["present"]!.GetValue<bool>());
        Assert.Equal("Georgia", detail["bodyFont"]!.GetValue<string>());

        // bodyFont alone leaves the background unset (null is omitted from the wire form).
        Assert.False(detail.ContainsKey("background"));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- handout master ------------------------------------------------------

    [Fact]
    public void SetHandoutMasterHeaderFooter_WritesPlaceholdersAndHfFlags_ReopenVerified()
    {
        Create();
        var data = Edit(TestEnv.Op("set", "/handoutMaster", props: TestEnv.Props(
            ("headerFooter", new JsonObject
            {
                ["header"] = "Acme Q3 Review",
                ["footer"] = "Confidential",
                ["date"] = true,
                ["pageNumber"] = true,
            }))));
        Assert.Equal("/handoutMaster", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var master = doc.PresentationPart!.HandoutMasterPart!.HandoutMaster!;
            var hf = master.HeaderFooter!;
            Assert.True(hf.DateTime!.Value);
            Assert.True(hf.SlideNumber!.Value);

            var tree = master.CommonSlideData!.ShapeTree!;
            var header = tree.Elements<P.Shape>().Single(s =>
                s.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.PlaceholderShape!.Type!.Value
                    == P.PlaceholderValues.Header);
            Assert.Equal("Acme Q3 Review", header.TextBody!.InnerText);

            Assert.Single(doc.PresentationPart.Presentation!.HandoutMasterIdList!.Elements<P.HandoutMasterId>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetHandoutMasterSlidesPerPage_WritesPresPropsPrintWhat_ReopenVerified()
    {
        Create();
        Edit(TestEnv.Op("set", "/handoutMaster", props: TestEnv.Props(("slidesPerPage", 6))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var printWhat = doc.PresentationPart!.PresentationPropertiesPart!
                .PresentationProperties!.PrintingProperties!.PrintWhat!.Value;
            Assert.Equal(P.PrintOutputValues.Handouts6, printWhat);
        }

        var detail = Get("/handoutMaster");
        Assert.Equal(6, detail["slidesPerPage"]!.GetValue<int>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Get_HandoutMaster_ReportsHeaderFooterAndSlidesPerPage()
    {
        Create();
        Edit(
            TestEnv.Op("set", "/handoutMaster", props: TestEnv.Props(
                ("background", "FFFFFF"),
                ("headerFooter", new JsonObject { ["header"] = "Hdr", ["pageNumber"] = false }),
                ("slidesPerPage", 4))));

        var detail = Get("/handoutMaster");
        Assert.True(detail["present"]!.GetValue<bool>());
        Assert.Equal("FFFFFF", detail["background"]!.GetValue<string>());
        Assert.Equal("Hdr", detail["headerFooter"]!["header"]!.GetValue<string>());
        Assert.False(detail["headerFooter"]!["pageNumber"]!.GetValue<bool>());
        Assert.Equal(4, detail["slidesPerPage"]!.GetValue<int>());
    }

    [Fact]
    public void Get_BeforeAnyEdit_ReportsMastersAbsent()
    {
        Create();
        var notes = Get("/notesMaster");
        var handout = Get("/handoutMaster");
        Assert.False(notes["present"]!.GetValue<bool>());
        Assert.False(handout["present"]!.GetValue<bool>());
    }

    // ---- structure view ------------------------------------------------------

    [Fact]
    public void Structure_ListsNotesAndHandoutMasters()
    {
        Create();
        Edit(
            TestEnv.Op("set", "/notesMaster", props: TestEnv.Props(("background", "112233"))),
            TestEnv.Op("set", "/handoutMaster", props: TestEnv.Props(("slidesPerPage", 9))));

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        Assert.True(data["notesMaster"]!["present"]!.GetValue<bool>());
        Assert.Equal("112233", data["notesMaster"]!["background"]!.GetValue<string>());
        Assert.True(data["handoutMaster"]!["present"]!.GetValue<bool>());
        Assert.Equal(9, data["handoutMaster"]!["slidesPerPage"]!.GetValue<int>());
    }

    // ---- validation / guards -------------------------------------------------

    [Fact]
    public void SetNotesMaster_UnknownProp_Fails()
    {
        Create();
        var env = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/notesMaster", props: TestEnv.Props(("slidesPerPage", 4)))]);
        TestEnv.AssertFail(env, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void SetHandoutMaster_BadSlidesPerPage_Fails()
    {
        Create();
        var env = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/handoutMaster", props: TestEnv.Props(("slidesPerPage", 5)))]);
        TestEnv.AssertFail(env, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void AddOnNotesMaster_IsRejected()
    {
        Create();
        var env = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/notesMaster", type: "shape")]);
        TestEnv.AssertFail(env, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void RemoveHandoutMaster_IsRejected()
    {
        Create();
        var env = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("remove", "/handoutMaster")]);
        TestEnv.AssertFail(env, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void NotesMasterSubSegment_IsInvalidPath()
    {
        Create();
        var env = _handler.Get(_ws.Ctx("deck.pptx", ("path", "/notesMaster/shape[1]")));
        TestEnv.AssertFail(env, ErrorCodes.InvalidPath);
    }

    [Fact]
    public void EditingTwice_IsIdempotentOnTheSamePart()
    {
        Create();
        Edit(TestEnv.Op("set", "/notesMaster", props: TestEnv.Props(("background", "112233"))));
        Edit(TestEnv.Op("set", "/notesMaster", props: TestEnv.Props(("bodyFont", "Verdana"))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        // A single notes master part is reused, not duplicated.
        Assert.Single(doc.PresentationPart!.Parts, p => p.OpenXmlPart is NotesMasterPart);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
