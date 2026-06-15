using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>v1.5.0 Word table formulas: =SUM(ABOVE)/=AVERAGE(LEFT)/cell-ref forms compute cached values.</summary>
public sealed class TableFormulaTests : WordTestBase
{
    private JsonNode Get(string file, string path) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    private static List<string> WarningCodes(Envelope envelope) =>
        [.. (envelope.Meta?.Warnings ?? []).Select(w => w.Code)];

    /// <summary>A single-column table: 10, 20, 30 then a blank fourth row to carry a formula.</summary>
    private string CreateColumnTable(params string[] values)
    {
        var file = CreateDoc(title: "Formulas");
        var rows = values.Length + 1;
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":ROWS,"columns":1}}]"""
            .Replace("ROWS", rows.ToString(System.Globalization.CultureInfo.InvariantCulture), System.StringComparison.Ordinal));
        var ops = new JsonArray();
        for (var r = 0; r < values.Length; r++)
        {
            ops.Add(new JsonObject
            {
                ["op"] = "set",
                ["path"] = $"/body/table[1]/tr[{r + 1}]/tc[1]",
                ["props"] = new JsonObject { ["text"] = values[r] },
            });
        }

        Edit(file, ops.ToJsonString());
        return file;
    }

    // ------------------------------------------------------------- SUM(ABOVE)

    [Fact]
    public void Sum_above_computes_cached_value_and_writes_a_field()
    {
        var file = CreateColumnTable("10", "20", "30");
        const int lastRow = 4;

        var envelope = Edit(file, """[{"op":"set","path":"/body/table[1]/tr[4]/tc[1]","props":{"formula":"=SUM(ABOVE)"}}]""");
        Assert.True(envelope.IsOk, envelope.ToJson());

        // The op result echoes the formula and cached value.
        var op = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("=SUM(ABOVE)", op["formula"]!.GetValue<string>());
        Assert.Equal("60", op["cached"]!.GetValue<string>());

        // The cell now holds a w:fldSimple whose instruction is the = formula.
        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!
                .Descendants<SimpleField>().Single();
            Assert.Contains("=SUM(ABOVE)", field.Instruction!.Value!, StringComparison.Ordinal);
            Assert.Equal("60", field.InnerText);
        }

        // get reports the formula + cached value, reopen-verifiable.
        var props = Get(file, $"/body/table[1]/tr[{lastRow}]/tc[1]")["properties"]!;
        Assert.Equal("=SUM(ABOVE)", props["formula"]!.GetValue<string>());
        Assert.Equal("60", props["cached"]!.GetValue<string>());

        AssertValidatesClean(file);
    }

    [Fact]
    public void Sum_above_stops_at_non_numeric_header_text()
    {
        // A "Total" header above numeric cells: SUM still adds only the numbers.
        var file = CreateColumnTable("Amount", "5", "7", "8");

        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[5]/tc[1]","props":{"formula":"=SUM(ABOVE)"}}]""");

        Assert.Equal("20", Get(file, "/body/table[1]/tr[5]/tc[1]")["properties"]!["cached"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    // ----------------------------------------------------------- AVERAGE(LEFT)

    [Fact]
    public void Average_left_computes_over_the_row()
    {
        var file = CreateDoc(title: "Avg");
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":1,"columns":4}}]""");
        Edit(file, """
        [{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"10"}},
         {"op":"set","path":"/body/table[1]/tr[1]/tc[2]","props":{"text":"20"}},
         {"op":"set","path":"/body/table[1]/tr[1]/tc[3]","props":{"text":"30"}}]
        """);

        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[4]","props":{"formula":"=AVERAGE(LEFT)"}}]""");

        Assert.Equal("20", Get(file, "/body/table[1]/tr[1]/tc[4]")["properties"]!["cached"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Product_above_multiplies()
    {
        var file = CreateColumnTable("2", "3", "4");
        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[4]/tc[1]","props":{"formula":"=PRODUCT(ABOVE)"}}]""");
        Assert.Equal("24", Get(file, "/body/table[1]/tr[4]/tc[1]")["properties"]!["cached"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    // ---------------------------------------------------------------- cell-ref

    [Fact]
    public void Cell_ref_arithmetic_computes_using_a1_addressing()
    {
        // A 2x2 grid: A1=5 B1=6 / A2=7 B2=8; a formula cell computes =A1*B2.
        var file = CreateDoc(title: "Refs");
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":3,"columns":2}}]""");
        Edit(file, """
        [{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"5"}},
         {"op":"set","path":"/body/table[1]/tr[1]/tc[2]","props":{"text":"6"}},
         {"op":"set","path":"/body/table[1]/tr[2]/tc[1]","props":{"text":"7"}},
         {"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"text":"8"}}]
        """);

        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[3]/tc[1]","props":{"formula":"=A1*B2"}}]""");
        Assert.Equal("40", Get(file, "/body/table[1]/tr[3]/tc[1]")["properties"]!["cached"]!.GetValue<string>());

        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[3]/tc[2]","props":{"formula":"=(A1+A2)/2"}}]""");
        Assert.Equal("6", Get(file, "/body/table[1]/tr[3]/tc[2]")["properties"]!["cached"]!.GetValue<string>());

        AssertValidatesClean(file);
    }

    // ---------------------------------------------------------- number format

    [Fact]
    public void Number_format_preset_shapes_the_cached_text()
    {
        var file = CreateColumnTable("100", "250.5");
        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[3]/tc[1]","props":{"formula":"=SUM(ABOVE)","numberFormat":"currency"}}]""");

        var props = Get(file, "/body/table[1]/tr[3]/tc[1]")["properties"]!;
        Assert.Equal("$350.50", props["cached"]!.GetValue<string>());

        // The picture is stored on the field instruction for Word to reuse.
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
        Assert.Contains("\\#", field.Instruction!.Value!, StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Number_format_without_formula_is_rejected()
    {
        var file = CreateColumnTable("1", "2");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"numberFormat":"integer"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    // ------------------------------------------------------- text view + note

    [Fact]
    public void Read_text_view_shows_the_computed_value()
    {
        var file = CreateColumnTable("10", "20", "30");
        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[4]/tc[1]","props":{"formula":"=SUM(ABOVE)"}}]""");

        var texts = BodyTexts(file);
        Assert.Contains(texts, t => t.Contains("60", StringComparison.Ordinal));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Static_inputs_compute_without_a_cached_warning()
    {
        var file = CreateColumnTable("10", "20", "30");
        var envelope = Edit(file, """[{"op":"set","path":"/body/table[1]/tr[4]/tc[1]","props":{"formula":"=SUM(ABOVE)"}}]""");
        Assert.DoesNotContain("table_formula_cached", WarningCodes(envelope));
    }

    [Fact]
    public void Formula_over_a_field_input_emits_table_formula_cached_note()
    {
        // Put a PAGE field in a cell, then sum a column that includes it: the
        // field's number is a cache, so the formula warns it may change on F9.
        var file = CreateColumnTable("10", "20");
        // Replace the second data cell's content with a numeric field-bearing cell
        // by chaining: a formula cell ABOVE another formula triggers the field input.
        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[2]/tc[1]","props":{"formula":"=SUM(ABOVE)"}}]""");

        var envelope = Edit(file, """[{"op":"set","path":"/body/table[1]/tr[3]/tc[1]","props":{"formula":"=SUM(ABOVE)"}}]""");
        Assert.Contains("table_formula_cached", WarningCodes(envelope));
    }

    // ----------------------------------------------------------------- guards

    [Fact]
    public void Unknown_function_is_unsupported_feature()
    {
        var file = CreateColumnTable("1", "2");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]/tr[3]/tc[1]","props":{"formula":"=MEDIAN(ABOVE)"}}]"""));
        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }

    [Fact]
    public void Formula_without_equals_is_invalid_args()
    {
        var file = CreateColumnTable("1", "2");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]/tr[3]/tc[1]","props":{"formula":"SUM(ABOVE)"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Reopen_verify_formula_survives_round_trip()
    {
        var file = CreateColumnTable("3", "4", "5");
        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[4]/tc[1]","props":{"formula":"=SUM(ABOVE)"}}]""");

        // Re-read after a second unrelated edit: the formula + cache persist.
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"after"}}]""");
        var props = Get(file, "/body/table[1]/tr[4]/tc[1]")["properties"]!;
        Assert.Equal("=SUM(ABOVE)", props["formula"]!.GetValue<string>());
        Assert.Equal("12", props["cached"]!.GetValue<string>());
        AssertValidatesClean(file);
    }
}
