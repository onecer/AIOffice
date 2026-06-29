using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// v1.20.0 per-cell table text flow: the additive <c>textDirection</c> prop on a tc set
/// op writes w:tcPr/w:textDirection (@w:val), bringing docx to parity with the pptx
/// table-cell textDirection. Every assertion reopens the part with the OpenXML SDK so we
/// verify the OOXML we wrote, not our in-memory model.
/// </summary>
public sealed class CellTextDirectionTests : WordTestBase
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

    /// <summary>
    /// The w:textDirection @w:val of one cell, read straight from the saved part as the
    /// serialized OOXML token (e.g. "tbRl"), not the SDK enum's ToString.
    /// </summary>
    private static string? ReadTextDirectionVal(string file, int row, int col)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().First();
        var tr = table.Elements<TableRow>().ElementAt(row - 1);
        var tc = tr.Elements<TableCell>().ElementAt(col - 1);
        var textDirection = tc.GetFirstChild<TableCellProperties>()?.TextDirection;
        return textDirection?.Val is { } val ? val.InnerText : null;
    }

    // ---------------------------------------------------------------- write

    [Fact]
    public void Tbrl_writes_text_direction_val_and_get_reports_it_on_reopen()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [{"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"textDirection":"tbRl"}}]
            """);

        // Raw XML carries w:textDirection @w:val=tbRl.
        Assert.Equal("tbRl", ReadTextDirectionVal(file, 2, 2));

        // get reports the same token (this reopens the part, so it persists on reopen).
        Assert.Equal("tbRl", Get(file, "/body/table[1]/tr[2]/tc[2]")["properties"]!["textDirection"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Lrtb_and_btlr_each_write_their_val_and_persist()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [
              {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"textDirection":"lrTb"}},
              {"op":"set","path":"/body/table[1]/tr[3]/tc[3]","props":{"textDirection":"btLr"}}
            ]
            """);

        Assert.Equal("lrTb", ReadTextDirectionVal(file, 1, 1));
        Assert.Equal("btLr", ReadTextDirectionVal(file, 3, 3));

        Assert.Equal("lrTb", Get(file, "/body/table[1]/tr[1]/tc[1]")["properties"]!["textDirection"]!.GetValue<string>());
        Assert.Equal("btLr", Get(file, "/body/table[1]/tr[3]/tc[3]")["properties"]!["textDirection"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    // ----------------------------------------------------------- round-trip

    [Fact]
    public void Hand_coded_direction_reads_back_then_overwrites_leaving_siblings_untouched()
    {
        var file = CreateWith3x3();

        // Hand-code a w:textDirection(btLr) straight onto the part, bypassing the handler,
        // to prove get reads OOXML we did not write ourselves.
        using (var doc = WordprocessingDocument.Open(file, isEditable: true))
        {
            var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().First();
            var tc = table.Elements<TableRow>().ElementAt(0).Elements<TableCell>().ElementAt(0);
            var tcPr = tc.GetFirstChild<TableCellProperties>() ?? tc.PrependChild(new TableCellProperties());
            tcPr.TextDirection = new TextDirection { Val = TextDirectionValues.BottomToTopLeftToRight };
            doc.MainDocumentPart.Document.Save();
        }

        Assert.Equal("btLr", Get(file, "/body/table[1]/tr[1]/tc[1]")["properties"]!["textDirection"]!.GetValue<string>());

        // Set a new value; it overwrites the hand-coded one and persists on reopen.
        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"textDirection":"lrTb"}}]""");
        Assert.Equal("lrTb", ReadTextDirectionVal(file, 1, 1));
        Assert.Equal("lrTb", Get(file, "/body/table[1]/tr[1]/tc[1]")["properties"]!["textDirection"]!.GetValue<string>());

        // A sibling cell stays free of any text direction.
        Assert.Null(ReadTextDirectionVal(file, 1, 2));
        Assert.Null(Get(file, "/body/table[1]/tr[1]/tc[2]")["properties"]!["textDirection"]);
        AssertValidatesClean(file);
    }

    // --------------------------------------------------------- byte-stable

    [Fact]
    public void Cell_without_text_direction_reports_null_and_legacy_doc_is_byte_identical()
    {
        var file = CreateWith3x3();

        // The textDirection key is present (like shading/valign) but null on a plain cell.
        var props = Get(file, "/body/table[1]/tr[1]/tc[1]")["properties"]!.AsObject();
        Assert.True(props.ContainsKey("textDirection"));
        Assert.Null(props["textDirection"]);

        // A read-only get does not rewrite the package: the bytes are unchanged, so a
        // legacy doc's cells stay byte-identical when no textDirection is involved.
        var before = File.ReadAllBytes(file);
        _ = Get(file, "/body/table[1]/tr[1]/tc[1]");
        Assert.Equal(before, File.ReadAllBytes(file));
    }

    // --------------------------------------------------------------- negative

    [Fact]
    public void Invalid_token_is_invalid_args_with_candidates()
    {
        var file = CreateWith3x3();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"textDirection":"invalid"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("lrTb", ex.Candidates!);
        Assert.Contains("tbRl", ex.Candidates!);
        Assert.Contains("btLr", ex.Candidates!);
    }

    // ------------------------------------------------------------ coexistence

    [Fact]
    public void Text_direction_coexists_with_valign_shading_and_borders()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [{"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{
              "valign":"center","shading":"FFE699","textDirection":"tbRl",
              "borders":{"bottom":{"color":"C00000","widthPt":1.5}}}}]
            """);

        // All four cosmetics read back together.
        var props = Get(file, "/body/table[1]/tr[2]/tc[2]")["properties"]!;
        Assert.Equal("center", props["valign"]!.GetValue<string>());
        Assert.Equal("FFE699", props["shading"]!.GetValue<string>());
        Assert.Equal("tbRl", props["textDirection"]!.GetValue<string>());
        Assert.Equal("C00000", props["borders"]!.AsObject()["bottom"]!.AsObject()["color"]!.GetValue<string>());

        // And the raw element is present at its schema-ordered position.
        Assert.Equal("tbRl", ReadTextDirectionVal(file, 2, 2));

        // OpenXmlValidator passes: the tcPr child order (shd, tcBorders, textDirection,
        // vAlign) is correct.
        AssertValidatesClean(file);
    }
}
