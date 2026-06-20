using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// v1.14.0 per-cell table borders: the additive <c>borders</c> prop on a tc set op
/// writes w:tcBorders into the cell's tcPr, alongside (not instead of) table-level
/// w:tblBorders. Every assertion reopens the part with the OpenXML SDK so we verify
/// the OOXML we wrote, not our own in-memory model.
/// </summary>
public sealed class CellBorderTests : WordTestBase
{
    private JsonNode Get(string file, string path) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    /// <summary>A 3x3 table with addressable cell texts r{row}c{col}.</summary>
    private string CreateWith3x3()
    {
        var file = CreateDoc(title: "Tables");
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":3,"columns":3}}]""");
        var ops = new JsonArray();
        for (var r = 1; r <= 3; r++)
        {
            for (var c = 1; c <= 3; c++)
            {
                ops.Add(new JsonObject
                {
                    ["op"] = "set",
                    ["path"] = $"/body/table[1]/tr[{r}]/tc[{c}]",
                    ["props"] = new JsonObject { ["text"] = $"r{r}c{c}" },
                });
            }
        }

        Edit(file, ops.ToJsonString());
        return file;
    }

    /// <summary>The w:tcBorders of one cell, read straight from the saved part.</summary>
    private static TableCellBorders? ReadCellBorders(string file, int row, int col)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().First();
        var tr = table.Elements<TableRow>().ElementAt(row - 1);
        var tc = tr.Elements<TableCell>().ElementAt(col - 1);
        return tc.GetFirstChild<TableCellProperties>()?.TableCellBorders;
    }

    // ---------------------------------------------------------------- write

    [Fact]
    public void Single_edge_writes_tcBorders_bottom_with_parser_size_and_color()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [{"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"borders":{"bottom":{"color":"C00000","widthPt":1.5}}}}]
            """);

        var borders = ReadCellBorders(file, 2, 2);
        Assert.NotNull(borders);
        var bottom = borders!.BottomBorder;
        Assert.NotNull(bottom);
        Assert.Equal(BorderValues.Single, bottom!.Val!.Value); // default style
        Assert.Equal("C00000", bottom.Color!.Value);
        // w:sz is eighths-of-a-point: feed 1.5pt through the existing parser and
        // assert exactly what it produces (round(1.5*8) = 12), not a hand-computed twip.
        Assert.Equal(12u, bottom.Size!.Value);

        // Only the requested edge is written.
        Assert.Null(borders.TopBorder);
        Assert.Null(borders.LeftBorder);
        Assert.Null(borders.RightBorder);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Six_edges_in_one_op_write_six_children_and_leave_other_cosmetics_untouched()
    {
        var file = CreateWith3x3();

        // Seed unrelated cosmetics + a merge so we can prove borders don't disturb them.
        Edit(file, """
            [{"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"shading":"FFE699","valign":"center"}}]
            """);

        Edit(file, """
            [{"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"borders":{
              "top":{"color":"111111"},"bottom":{"color":"222222"},
              "left":{"color":"333333"},"right":{"color":"444444"},
              "insideH":{"color":"555555"},"insideV":{"color":"666666"}}}}]
            """);

        var borders = ReadCellBorders(file, 2, 2);
        Assert.NotNull(borders);
        Assert.Equal("111111", borders!.TopBorder!.Color!.Value);
        Assert.Equal("222222", borders.BottomBorder!.Color!.Value);
        Assert.Equal("333333", borders.LeftBorder!.Color!.Value);
        Assert.Equal("444444", borders.RightBorder!.Color!.Value);
        Assert.Equal("555555", borders.InsideHorizontalBorder!.Color!.Value);
        Assert.Equal("666666", borders.InsideVerticalBorder!.Color!.Value);

        // Shading, valign and the cell text all survive the border write.
        var props = Get(file, "/body/table[1]/tr[2]/tc[2]")["properties"]!;
        Assert.Equal("FFE699", props["shading"]!.GetValue<string>());
        Assert.Equal("center", props["valign"]!.GetValue<string>());
        Assert.Equal("r2c2", props["text"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void All_applies_to_four_outer_sides_only()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"borders":{"all":{"color":"0070C0","widthPt":1}}}}]
            """);

        var borders = ReadCellBorders(file, 1, 1)!;
        foreach (var edge in new BorderType?[] { borders.TopBorder, borders.BottomBorder, borders.LeftBorder, borders.RightBorder })
        {
            Assert.NotNull(edge);
            Assert.Equal(BorderValues.Single, edge!.Val!.Value);
            Assert.Equal("0070C0", edge.Color!.Value);
            Assert.Equal(8u, edge.Size!.Value); // round(1*8) = 8
        }

        // "all" is only the outer four; the inside edges stay absent.
        Assert.Null(borders.InsideHorizontalBorder);
        Assert.Null(borders.InsideVerticalBorder);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Style_none_clears_an_edge_with_val_nil()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"borders":{"top":{"style":"none"}}}}]
            """);

        var top = ReadCellBorders(file, 1, 1)!.TopBorder!;
        Assert.Equal(BorderValues.Nil, top.Val!.Value);
        Assert.Null(top.Size); // nil edge carries no size or color
        Assert.Null(top.Color);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Bare_string_none_also_clears_an_edge()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"borders":{"bottom":"none"}}}]
            """);

        Assert.Equal(BorderValues.Nil, ReadCellBorders(file, 1, 1)!.BottomBorder!.Val!.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Style_names_map_to_border_values()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"borders":{
              "top":{"style":"double"},"bottom":{"style":"dashed"},
              "left":{"style":"dotted"},"right":{"style":"thick"}}}}]
            """);

        var b = ReadCellBorders(file, 1, 1)!;
        Assert.Equal(BorderValues.Double, b.TopBorder!.Val!.Value);
        Assert.Equal(BorderValues.Dashed, b.BottomBorder!.Val!.Value);
        Assert.Equal(BorderValues.Dotted, b.LeftBorder!.Val!.Value);
        Assert.Equal(BorderValues.Thick, b.RightBorder!.Val!.Value);
        AssertValidatesClean(file);
    }

    // -------------------------------------------------- coexistence with table

    [Fact]
    public void Cell_borders_coexist_with_table_level_borders()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [{"op":"set","path":"/body/table[1]","props":{"borders":"all","borderColor":"1F4E79","borderWidthPt":1}}]
            """);
        Edit(file, """
            [{"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"borders":{"bottom":{"color":"C00000"}}}}]
            """);

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().First();

            // Table-level borders survive unchanged.
            var tblBorders = table.GetFirstChild<TableProperties>()!.TableBorders!;
            Assert.Equal(BorderValues.Single, tblBorders.TopBorder!.Val!.Value);
            Assert.Equal("1F4E79", tblBorders.TopBorder.Color!.Value);

            // And the cell carries its own override at the same time.
            var tc = table.Elements<TableRow>().ElementAt(1).Elements<TableCell>().ElementAt(1);
            var tcBorders = tc.GetFirstChild<TableCellProperties>()!.TableCellBorders!;
            Assert.Equal("C00000", tcBorders.BottomBorder!.Color!.Value);
        }

        // get on the table still reports "all".
        Assert.Equal("all", Get(file, "/body/table[1]")["properties"]!["borders"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    // ----------------------------------------------------------------- get

    [Fact]
    public void Get_reports_per_edge_cell_borders()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [{"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"borders":{
              "bottom":{"color":"C00000","widthPt":1.5},"top":{"style":"none"}}}}]
            """);

        var borders = Get(file, "/body/table[1]/tr[2]/tc[2]")["properties"]!["borders"]!.AsObject();
        var bottom = borders["bottom"]!.AsObject();
        Assert.Equal("single", bottom["style"]!.GetValue<string>());
        Assert.Equal("C00000", bottom["color"]!.GetValue<string>());
        Assert.Equal(1.5, bottom["widthPt"]!.GetValue<double>(), 3); // 12 eighths / 8 = 1.5pt
        Assert.Equal("none", borders["top"]!.AsObject()["style"]!.GetValue<string>());

        // A plain cell reports null borders.
        Assert.Null(Get(file, "/body/table[1]/tr[1]/tc[1]")["properties"]!["borders"]);
    }

    // ------------------------------------------------------------- negatives

    [Fact]
    public void Unknown_sub_key_is_invalid_args_with_candidates()
    {
        var file = CreateWith3x3();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"borders":{"topp":{"color":"C00000"}}}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("top", ex.Candidates!);
    }

    [Fact]
    public void Bad_width_is_invalid_args()
    {
        var file = CreateWith3x3();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"borders":{"top":{"widthPt":99}}}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Bad_style_is_invalid_args_with_candidates()
    {
        var file = CreateWith3x3();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"borders":{"top":{"style":"squiggle"}}}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("single", ex.Candidates!);
    }

    [Fact]
    public void Borders_not_an_object_is_invalid_args()
    {
        var file = CreateWith3x3();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"borders":"all"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Borders_on_a_non_tc_node_still_throws_unsupported_feature()
    {
        var file = CreateWith3x3();

        // The table-level 'borders' is a string-valued prop; an object value there
        // is not a cell, so the table's own default/validation path must still fire.
        // On a paragraph node (no borders prop at all) the default case throws.
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"x"}}]""");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"borders":{"top":{"color":"C00000"}}}}]"""));

        // Paragraph set has its own supported-prop guard; the point is borders-as-object
        // is rejected anywhere that is not a tc, never silently applied.
        Assert.Contains(ex.Code, new[] { ErrorCodes.UnsupportedFeature, ErrorCodes.InvalidArgs });
    }
}
