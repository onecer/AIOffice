using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Dw = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class BodyShapeTests : WordTestBase
{
    private JsonNode Get(string file, string path) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    [Fact]
    public void Add_rect_shape_reopens_with_wsp_and_geometry()
    {
        var file = CreateDoc(title: "Shapes");
        Edit(file, """[{"op":"add","path":"/body","type":"shape","props":{"shape":"rect","x":"2cm","y":"3cm","w":"6cm","h":"4cm","fill":"38BDF8","text":"Hi box"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var wsp = Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<Wps.WordprocessingShape>());
            var preset = wsp.Descendants<A.PresetGeometry>().Single().Preset!.Value;
            Assert.Equal(A.ShapeTypeValues.Rectangle, preset);
            Assert.Equal("Hi box", wsp.InnerText);
            var anchor = Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<Dw.Anchor>());
            Assert.Equal(6 * 360000L, anchor.Extent!.Cx!.Value);
            Assert.Equal(4 * 360000L, anchor.Extent!.Cy!.Value);
        }

        // A shape stays a shape even when it carries text (addressing is fixed at
        // creation, not by content) — so it lives at /body/shape[1].
        var got = Get(file, "/body/shape[1]");
        Assert.Equal("shape", got["type"]!.GetValue<string>());
        Assert.Equal("Hi box", got["properties"]!["text"]!.GetValue<string>());
        Assert.Equal(2.0, got["properties"]!["xCm"]!.GetValue<double>());
        Assert.Equal(6.0, got["properties"]!["wCm"]!.GetValue<double>());
        Assert.Equal("38BDF8", got["properties"]!["fill"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Add_plain_shape_without_text_is_addressed_as_shape()
    {
        var file = CreateDoc(title: "PlainShape");
        Edit(file, """[{"op":"add","path":"/body","type":"shape","props":{"shape":"ellipse","x":"1cm","y":"1cm","w":"5cm","h":"5cm","fill":"FF0000"}}]""");

        var got = Get(file, "/body/shape[1]");
        Assert.Equal("shape", got["type"]!.GetValue<string>());
        Assert.Equal("ellipse", got["properties"]!["shape"]!.GetValue<string>());
        Assert.Null(got["properties"]!["text"]);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Add_line_with_arrow_carries_tail_end()
    {
        var file = CreateDoc(title: "Arrow");
        Edit(file, """[{"op":"add","path":"/body","type":"shape","props":{"shape":"arrow","x":"1cm","y":"1cm","w":"8cm","h":"0cm","line":"333333"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var wsp = doc.MainDocumentPart!.Document!.Body!.Descendants<Wps.WordprocessingShape>().Single();
            Assert.Equal(A.ShapeTypeValues.Line, wsp.Descendants<A.PresetGeometry>().Single().Preset!.Value);
            Assert.Equal(A.LineEndValues.Arrow, wsp.Descendants<A.TailEnd>().Single().Type!.Value);
        }

        var got = Get(file, "/body/shape[1]");
        Assert.Equal("333333", got["properties"]!["line"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Add_text_box_reopens_with_text_and_square_wrap()
    {
        var file = CreateDoc(title: "TextBox");
        Edit(file, """[{"op":"add","path":"/body","type":"textBox","props":{"x":"2cm","y":"2cm","w":"7cm","h":"3cm","text":"Sidebar note"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var anchor = doc.MainDocumentPart!.Document!.Body!.Descendants<Dw.Anchor>().Single();
            Assert.NotNull(anchor.GetFirstChild<Dw.WrapSquare>());
            Assert.Equal("Sidebar note", anchor.Descendants<Wps.WordprocessingShape>().Single().InnerText);
        }

        var got = Get(file, "/body/textBox[1]");
        Assert.Equal("Sidebar note", got["properties"]!["text"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Set_updates_geometry_fill_and_text()
    {
        var file = CreateDoc(title: "SetShape");
        Edit(file, """[{"op":"add","path":"/body","type":"textBox","props":{"x":"1cm","y":"1cm","w":"4cm","h":"2cm","text":"Before"}}]""");

        Edit(file, """[{"op":"set","path":"/body/textBox[1]","props":{"w":"9cm","fill":"00FF00","text":"After"}}]""");

        var got = Get(file, "/body/textBox[1]");
        Assert.Equal(9.0, got["properties"]!["wCm"]!.GetValue<double>());
        Assert.Equal("00FF00", got["properties"]!["fill"]!.GetValue<string>());
        Assert.Equal("After", got["properties"]!["text"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Setting_text_on_a_shape_keeps_it_addressed_as_a_shape()
    {
        var file = CreateDoc(title: "StableAddr");
        Edit(file, """[{"op":"add","path":"/body","type":"shape","props":{"shape":"rect","x":"1cm","y":"1cm","w":"4cm","h":"2cm"}}]""");

        Edit(file, """[{"op":"set","path":"/body/shape[1]","props":{"text":"Now labelled"}}]""");

        // It now has text, but it is still a shape (not promoted to a text box).
        var got = Get(file, "/body/shape[1]");
        Assert.Equal("Now labelled", got["properties"]!["text"]!.GetValue<string>());
        var ex = Assert.Throws<AiofficeException>(() => Get(file, "/body/textBox[1]"));
        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Remove_shape_drops_its_carrier_paragraph()
    {
        var file = CreateDoc(title: "RemoveShape");
        Edit(file, """[{"op":"add","path":"/body","type":"shape","props":{"shape":"rect","x":"1cm","y":"1cm","w":"4cm","h":"2cm"}}]""");

        Edit(file, """[{"op":"remove","path":"/body/shape[1]"}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<Wps.WordprocessingShape>());
        }

        var ex = Assert.Throws<AiofficeException>(() => Get(file, "/body/shape[1]"));
        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Unsupported_shape_kind_is_unsupported_feature_with_candidates()
    {
        var file = CreateDoc(title: "BadShape");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"shape","props":{"shape":"star","x":"1cm","y":"1cm","w":"4cm","h":"4cm"}}]"""));
        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("rect", ex.Candidates!);
    }

    [Fact]
    public void Missing_text_box_path_is_invalid_path()
    {
        var file = CreateDoc(title: "Empty");
        var ex = Assert.Throws<AiofficeException>(() => Get(file, "/body/textBox[1]"));
        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
    }

    [Fact]
    public void Structure_view_lists_shapes_and_text_boxes()
    {
        var file = CreateDoc(title: "StructShapes");
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"shape","props":{"shape":"rect","x":"1cm","y":"1cm","w":"4cm","h":"2cm"}},
              {"op":"add","path":"/body","type":"textBox","props":{"x":"1cm","y":"5cm","w":"4cm","h":"2cm","text":"Box"}}
            ]
            """);

        var structure = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));
        Assert.Single(structure["shapes"]!.AsArray());
        Assert.Single(structure["textBoxes"]!.AsArray());
        Assert.Equal("/body/shape[1]", structure["shapes"]![0]!["path"]!.GetValue<string>());
        Assert.Equal("/body/textBox[1]", structure["textBoxes"]![0]!["path"]!.GetValue<string>());
    }
}
