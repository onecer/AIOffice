using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The v1.2 number-format preset library (additive). A named preset resolves to
/// its OOXML format code; any other string stays a literal custom format (the
/// v1.0 behavior, never broken). Raw assertions reopen the file with the OpenXml
/// SDK to read the actual stored format code.
/// </summary>
public sealed class NumberFormatPresetTests : ExcelTestBase
{
    /// <summary>The stored numberFormat code on a cell, read raw from the styles part.</summary>
    private static string? RawNumberFormat(string file, string sheetName, string address)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        var cell = worksheetPart.Worksheet!.Descendants<Cell>()
            .FirstOrDefault(c => c.CellReference?.Value == address);
        if (cell?.StyleIndex?.Value is not { } styleIndex)
        {
            return null;
        }

        var stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet!;
        var cellFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)styleIndex);
        if (cellFormat.NumberFormatId?.Value is not { } numberFormatId)
        {
            return null;
        }

        var custom = stylesheet.NumberingFormats?.Elements<NumberingFormat>()
            .FirstOrDefault(n => n.NumberFormatId?.Value == numberFormatId);
        return custom?.FormatCode?.Value;
    }

    [Theory]
    [InlineData("currency-usd", "\"$\"#,##0.00")]
    [InlineData("currency-eur", "\"€\"#,##0.00")]
    [InlineData("percent", "0%")]
    [InlineData("percent2", "0.00%")]
    [InlineData("scientific", "0.00E+00")]
    [InlineData("thousands", "#,##0")]
    [InlineData("date-iso", "yyyy-mm-dd")]
    [InlineData("datetime-iso", "yyyy-mm-dd hh:mm:ss")]
    [InlineData("duration", "[h]:mm:ss")]
    [InlineData("text", "@")]
    public void Named_preset_resolves_to_its_ooxml_code(string preset, string expectedCode)
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/B2", ("value", 1234.5), ("numberFormat", preset)));
        Assert.True(envelope.IsOk, envelope.ToJson());

        Assert.Equal(expectedCode, RawNumberFormat(file, "Sheet1", "B2"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Accounting_preset_resolves_to_the_accounting_code()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/B2", ("value", -42), ("numberFormat", "accounting-usd")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var code = RawNumberFormat(file, "Sheet1", "B2");
        Assert.NotNull(code);
        Assert.Contains("$", code);
        Assert.Contains("(", code); // negatives parenthesized — the accounting hallmark
        AssertValidatorClean(file);
    }

    // ----- (1.8) additive presets --------------------------------------------

    [Theory]
    [InlineData("usd-millions", "\"$\"#,##0.0,,\"M\"")]
    [InlineData("usd-thousands", "\"$\"#,##0,\"K\"")]
    [InlineData("millions", "#,##0.0,,\"M\"")]
    [InlineData("thousands-k", "#,##0,\"K\"")]
    [InlineData("percent-0", "0%")]
    [InlineData("percent-1", "0.0%")]
    [InlineData("percent-2", "0.00%")]
    public void V18_preset_resolves_to_its_ooxml_code(string preset, string expectedCode)
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/B2", ("value", 1234567), ("numberFormat", preset)));
        Assert.True(envelope.IsOk, envelope.ToJson());

        Assert.Equal(expectedCode, RawNumberFormat(file, "Sheet1", "B2"));
        AssertValidatorClean(file);
    }

    [Theory]
    [InlineData("usd-millions", 1234567, "$1.2M")]
    [InlineData("usd-thousands", 1234567, "$1,235K")]
    [InlineData("millions", 1234567, "1.2M")]
    [InlineData("thousands-k", 1234567, "1,235K")]
    [InlineData("percent-0", 0.1234, "12%")]
    [InlineData("percent-1", 0.1234, "12.3%")]
    [InlineData("percent-2", 0.1234, "12.34%")]
    public void V18_preset_renders_its_display_value(string preset, double value, string expectedText)
    {
        var file = CreateWorkbook();

        EditOps(file, SetOp("/Sheet1/B2", ("value", value), ("numberFormat", preset)));
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))));

        Assert.Equal(expectedText, data["text"]!.GetValue<string>());
    }

    [Theory]
    [InlineData("accounting-eur", "€")]
    [InlineData("accounting-gbp", "£")]
    public void V18_accounting_preset_mirrors_usd_with_its_symbol(string preset, string symbol)
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/B2", ("value", -1234.5), ("numberFormat", preset)));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var code = RawNumberFormat(file, "Sheet1", "B2");
        Assert.NotNull(code);
        Assert.Contains(symbol, code);
        Assert.Contains("(", code); // parenthesized negatives — the accounting hallmark

        // The negative value renders parenthesized with the right symbol.
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))));
        var text = data["text"]!.GetValue<string>();
        Assert.Contains(symbol, text);
        Assert.Contains("(", text);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Literal_custom_format_string_passes_through_unchanged()
    {
        var file = CreateWorkbook();

        // A real OOXML code is never a preset name; it must be stored verbatim.
        var envelope = EditOps(file, SetOp("/Sheet1/B2", ("value", 0.25), ("numberFormat", "0.000%")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        Assert.Equal("0.000%", RawNumberFormat(file, "Sheet1", "B2"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Custom_dollar_format_is_not_mistaken_for_a_preset()
    {
        var file = CreateWorkbook();

        // The pre-v1.2 example code stays literal (not swapped for currency-usd).
        var envelope = EditOps(file, SetOp("/Sheet1/B2", ("value", 100), ("numberFormat", "$#,##0.00")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        Assert.Equal("$#,##0.00", RawNumberFormat(file, "Sheet1", "B2"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Preset_displays_a_typed_value_correctly_via_get_text()
    {
        var file = CreateWorkbook();

        EditOps(file, SetOp("/Sheet1/B2", ("value", 0.5), ("numberFormat", "percent")));
        var envelope = Handler.Get(Ctx(file, ("path", "/Sheet1/B2")));
        var data = OkData(envelope);

        // The cell carries the resolved code and formats as a percentage.
        Assert.Equal("0%", data["numberFormat"]!.GetValue<string>());
        Assert.Equal("50%", data["text"]!.GetValue<string>());
    }

    [Fact]
    public void Preset_works_inside_a_named_cell_style()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(
            file,
            new EditOp
            {
                Op = "add",
                Type = "cellStyle",
                Path = "/styles",
                Props = new JsonObject { ["name"] = "Money", ["numberFormat"] = "currency-usd" },
            },
            SetOp("/Sheet1/C3", ("value", 9.99), ("cellStyle", "Money")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        Assert.Equal("\"$\"#,##0.00", RawNumberFormat(file, "Sheet1", "C3"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Help_text_lists_every_preset_with_its_code()
    {
        var text = ExcelNumberFormats.HelpText();
        foreach (var (name, code) in ExcelNumberFormats.Presets)
        {
            Assert.Contains("`" + name + "`", text);
            Assert.Contains("`" + code + "`", text);
        }
    }

    [Fact]
    public void Nearest_preset_suggests_for_a_typo_but_not_for_a_real_code()
    {
        Assert.Equal("currency-usd", ExcelNumberFormats.NearestPreset("curency-usd"));
        Assert.Equal("percent", ExcelNumberFormats.NearestPreset("percnt"));

        // Real OOXML codes must not be flagged as misspelled presets.
        Assert.Null(ExcelNumberFormats.NearestPreset("#,##0.00"));
        Assert.Null(ExcelNumberFormats.NearestPreset("yyyy-mm-dd hh:mm"));
        Assert.Null(ExcelNumberFormats.NearestPreset("@"));
    }
}
