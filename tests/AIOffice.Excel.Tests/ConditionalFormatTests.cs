using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The M2 conditional-formatting slice: cellIs | colorScale | dataBar |
/// containsText. Raw assertions reopen the file with the OpenXml SDK; every
/// mutating test must end validator-clean — including data bars, whose
/// ClosedXML GUID defects the post-save fix-up corrects.
/// </summary>
public sealed class ConditionalFormatTests : ExcelTestBase
{
    /// <summary>Numbers 1..30 in A1:C10 plus status text in D1:D3.</summary>
    private string CreateDataWorkbook() => CreateDataWorkbookNamed("book.xlsx");

    /// <summary>As <see cref="CreateDataWorkbook"/> but with a caller-chosen file name.</summary>
    private string CreateDataWorkbookNamed(string name)
    {
        var file = CreateWorkbook(name);
        var grid = new JsonArray();
        for (var r = 0; r < 10; r++)
        {
            grid.Add(new JsonArray((r * 3) + 1, (r * 3) + 2, (r * 3) + 3));
        }

        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1:C10", ("values", grid)),
            SetOp("/Sheet1/D1:D3", ("values", new JsonArray(
                new JsonArray("paid"),
                new JsonArray("overdue now"),
                new JsonArray("pending")))));
        Assert.True(envelope.IsOk, envelope.ToJson());
        return file;
    }

    /// <summary>All base cfRule elements of the first sheet, in document order.</summary>
    private static List<S.ConditionalFormattingRule> RawRules(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().First();
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return [.. worksheetPart.Worksheet!.Descendants<S.ConditionalFormattingRule>()];
    }

    [Fact]
    public void CellIs_greaterThan_writes_rule_and_dxf_fill()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:C10", "conditionalFormat",
            ("kind", "cellIs"), ("operator", ">"), ("value", 20),
            ("fill", "FFC7CE"), ("color", "9C0006"), ("bold", true)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal(
            "/Sheet1/conditionalFormat[1]",
            Json(envelope)["data"]!["ops"]![0]!["path"]!.GetValue<string>());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var rule = Assert.Single(RawRules(document));
        Assert.Equal(S.ConditionalFormatValues.CellIs, rule.Type!.Value);
        Assert.Equal(S.ConditionalFormattingOperatorValues.GreaterThan, rule.Operator!.Value);
        Assert.Equal("20", rule.Elements<S.Formula>().Single().Text);
        Assert.Equal("A1:C10", ((S.ConditionalFormatting)rule.Parent!).SequenceOfReferences!.InnerText);

        // The dxf the rule points at must carry the requested style.
        var dxf = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!
            .DifferentialFormats!.Elements<S.DifferentialFormat>().ElementAt((int)rule.FormatId!.Value);
        Assert.Contains("FFC7CE", dxf.OuterXml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("9C0006", dxf.OuterXml, StringComparison.OrdinalIgnoreCase);
        Assert.NotNull(dxf.Font?.Bold);
    }

    [Fact]
    public void CellIs_between_writes_two_formulas()
    {
        var file = CreateDataWorkbook();

        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "cellIs"), ("operator", "between"), ("value", 5), ("value2", 15),
            ("fill", "C6EFCE"))).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var rule = Assert.Single(RawRules(document));
        Assert.Equal(S.ConditionalFormattingOperatorValues.Between, rule.Operator!.Value);
        Assert.Equal(["5", "15"], rule.Elements<S.Formula>().Select(f => f.Text));
    }

    [Fact]
    public void Between_without_value2_fails_actionably()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "cellIs"), ("operator", "between"), ("value", 5), ("fill", "C6EFCE")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("value2", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unknown_operator_is_invalid_args_with_candidates()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "cellIs"), ("operator", "~"), ("value", 5), ("fill", "C6EFCE")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Equal([">", "<", ">=", "<=", "==", "!=", "between"], envelope.Error.Candidates!);
    }

    [Fact]
    public void ColorScale_three_colors_writes_min_percentile_max()
    {
        var file = CreateDataWorkbook();

        Assert.True(EditOps(file, AddOp("/Sheet1/B1:B10", "conditionalFormat",
            ("kind", "colorScale"),
            ("minColor", "F8696B"), ("midColor", "FFEB84"), ("maxColor", "63BE7B"))).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var rule = Assert.Single(RawRules(document));
        Assert.Equal(S.ConditionalFormatValues.ColorScale, rule.Type!.Value);
        var scale = rule.GetFirstChild<S.ColorScale>()!;
        Assert.Equal(
            ["min", "percentile", "max"],
            scale.Elements<S.ConditionalFormatValueObject>().Select(o => o.Type!.InnerText));
        Assert.Equal(
            ["FFF8696B", "FFFFEB84", "FF63BE7B"],
            scale.Elements<S.Color>().Select(c => c.Rgb!.Value));
    }

    [Fact]
    public void ColorScale_two_colors_omits_the_midpoint()
    {
        var file = CreateDataWorkbook();

        Assert.True(EditOps(file, AddOp("/Sheet1/C1:C10", "conditionalFormat",
            ("kind", "colorScale"), ("minColor", "FFFFFF"), ("maxColor", "63BE7B"))).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var scale = Assert.Single(RawRules(document)).GetFirstChild<S.ColorScale>()!;
        Assert.Equal(2, scale.Elements<S.ConditionalFormatValueObject>().Count());
        Assert.Equal(2, scale.Elements<S.Color>().Count());
    }

    [Fact]
    public void ColorScale_two_colors_get_omits_the_midpoint()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/C1:C10", "conditionalFormat",
            ("kind", "colorScale"), ("minColor", "FFFFFF"), ("maxColor", "63BE7B"))).IsOk);

        var scale = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("colorScale", scale["cfKind"]!.GetValue<string>());
        // A 2-color scale has no midpoint at all — no midColor, midType or midValue.
        Assert.Null(scale["midColor"]);
        Assert.Null(scale["midType"]);
        Assert.Null(scale["midValue"]);
    }

    [Fact]
    public void ColorScale_custom_midpoint_roundtrips_through_cfvo_and_get()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C10", "conditionalFormat",
            ("kind", "colorScale"),
            ("minColor", "F8696B"), ("midColor", "FFEB84"), ("maxColor", "63BE7B"),
            ("midType", "percent"), ("midValue", 25))).IsOk);

        AssertValidatorClean(file);
        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var scale = Assert.Single(RawRules(document)).GetFirstChild<S.ColorScale>()!;
            var cfvos = scale.Elements<S.ConditionalFormatValueObject>().ToList();
            Assert.Equal(["min", "percent", "max"], cfvos.Select(o => o.Type!.InnerText));
            Assert.Equal("25", cfvos[1].Val!.Value);
        }

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("FFEB84", data["midColor"]!.GetValue<string>());
        Assert.Equal("percent", data["midType"]!.GetValue<string>());
        Assert.Equal(25d, data["midValue"]!.GetValue<double>());
    }

    [Theory]
    [InlineData("num", 7, "num", "7")]
    [InlineData("percent", 40, "percent", "40")]
    [InlineData("percentile", 60, "percentile", "60")]
    public void ColorScale_each_midType_writes_its_cfvo(
        string midType, double midValue, string expectedCfvoType, string expectedVal)
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C10", "conditionalFormat",
            ("kind", "colorScale"),
            ("minColor", "F8696B"), ("midColor", "FFEB84"), ("maxColor", "63BE7B"),
            ("midType", midType), ("midValue", midValue))).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var scale = Assert.Single(RawRules(document)).GetFirstChild<S.ColorScale>()!;
        var mid = scale.Elements<S.ConditionalFormatValueObject>().ElementAt(1);
        Assert.Equal(expectedCfvoType, mid.Type!.InnerText);
        Assert.Equal(expectedVal, mid.Val!.Value);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal(midType, data["midType"]!.GetValue<string>());
        Assert.Equal(midValue, data["midValue"]!.GetValue<double>());
    }

    [Fact]
    public void ColorScale_default_midpoint_stays_percentile_50_and_get_omits_midType()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C10", "conditionalFormat",
            ("kind", "colorScale"),
            ("minColor", "F8696B"), ("midColor", "FFEB84"), ("maxColor", "63BE7B"))).IsOk);

        AssertValidatorClean(file);
        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var scale = Assert.Single(RawRules(document)).GetFirstChild<S.ColorScale>()!;
            var mid = scale.Elements<S.ConditionalFormatValueObject>().ElementAt(1);
            Assert.Equal("percentile", mid.Type!.InnerText);
            Assert.Equal("50", mid.Val!.Value);
        }

        // Legacy default: midColor surfaces, but midType/midValue are omitted.
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("FFEB84", data["midColor"]!.GetValue<string>());
        Assert.Null(data["midType"]);
        Assert.Null(data["midValue"]);
    }

    [Fact]
    public void ColorScale_midType_without_midColor_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1/A1:C10", "conditionalFormat",
            ("kind", "colorScale"), ("minColor", "F8696B"), ("maxColor", "63BE7B"),
            ("midType", "percent"), ("midValue", 25)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("midColor", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ColorScale_midValue_without_midType_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1/A1:C10", "conditionalFormat",
            ("kind", "colorScale"),
            ("minColor", "F8696B"), ("midColor", "FFEB84"), ("maxColor", "63BE7B"),
            ("midValue", 25)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        // Names the missing prop, not an empty-string "unknown midType ''".
        Assert.Contains("midType", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("percent", envelope.Error.Candidates ?? [], StringComparer.Ordinal);
    }

    [Theory]
    [InlineData("percent", -1)]
    [InlineData("percent", 101)]
    [InlineData("percentile", 150)]
    public void ColorScale_percent_midValue_out_of_range_is_invalid_args(string midType, double midValue)
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1/A1:C10", "conditionalFormat",
            ("kind", "colorScale"),
            ("minColor", "F8696B"), ("midColor", "FFEB84"), ("maxColor", "63BE7B"),
            ("midType", midType), ("midValue", midValue)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("between 0 and 100", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ColorScale_num_midValue_outside_data_range_is_accepted()
    {
        var file = CreateDataWorkbook();
        // A 'num' midpoint outside the data range is legal in Excel — accept it.
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C10", "conditionalFormat",
            ("kind", "colorScale"),
            ("minColor", "F8696B"), ("midColor", "FFEB84"), ("maxColor", "63BE7B"),
            ("midType", "num"), ("midValue", 9999))).IsOk);

        AssertValidatorClean(file);
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("num", data["midType"]!.GetValue<string>());
        Assert.Equal(9999d, data["midValue"]!.GetValue<double>());
    }

    [Fact]
    public void ColorScale_unknown_midType_is_invalid_args_with_candidates()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1/A1:C10", "conditionalFormat",
            ("kind", "colorScale"),
            ("minColor", "F8696B"), ("midColor", "FFEB84"), ("maxColor", "63BE7B"),
            ("midType", "quartile"), ("midValue", 25)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Equal(["num", "percent", "percentile"], envelope.Error.Candidates!);
    }

    [Fact]
    public void DataBar_is_validator_clean_after_the_guid_fixup()
    {
        var file = CreateDataWorkbook();

        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"))).IsOk);

        // The whole point: ClosedXML's lowercase pairing GUID fails the schema
        // pattern; the post-save pass must leave the file validator-clean.
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var rule = Assert.Single(RawRules(document));
        Assert.Equal(S.ConditionalFormatValues.DataBar, rule.Type!.Value);
        Assert.Equal("FF638EC6", rule.GetFirstChild<S.DataBar>()!.GetFirstChild<S.Color>()!.Rgb!.Value);

        // Base id and x14 twin id must still pair exactly (both uppercased).
        var baseId = rule.Descendants<X14.Id>().Single().Text;
        Assert.Equal(baseId, baseId.ToUpperInvariant());
        var worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(
            document.WorkbookPart!.Workbook!.Descendants<S.Sheet>().First().Id!.Value!);
        var twin = worksheetPart.Worksheet!.Descendants<X14.ConditionalFormattingRule>().Single();
        Assert.Equal(baseId, twin.Id!.Value);
    }

    [Fact]
    public void ContainsText_writes_text_rule_with_search_formula()
    {
        var file = CreateDataWorkbook();

        Assert.True(EditOps(file, AddOp("/Sheet1/D1:D3", "conditionalFormat",
            ("kind", "containsText"), ("text", "overdue"), ("fill", "FFC7CE"))).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var rule = Assert.Single(RawRules(document));
        Assert.Equal(S.ConditionalFormatValues.ContainsText, rule.Type!.Value);
        Assert.Equal("overdue", rule.Text?.Value);
        Assert.Contains("SEARCH(\"overdue\"", rule.Elements<S.Formula>().Single().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void Unsupported_kind_names_the_supported_ones()
    {
        var file = CreateDataWorkbook();

        // iconSet joined the supported set in v1.1; formula/topBottom/
        // aboveBelowAverage joined in v1.3. A genuinely-unknown kind still names
        // every supported kind (the now-expanded set) as candidates.
        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "timePeriod")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Equal(
            ["cellIs", "colorScale", "dataBar", "containsText", "iconSet", "formula", "topBottom", "aboveBelowAverage"],
            envelope.Error.Candidates!);
        Assert.Contains("cellIs", envelope.Error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Rule_without_style_props_is_refused()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "cellIs"), ("operator", ">"), ("value", 5)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("fill", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Get_returns_normalized_rule_essentials()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1/A1:C10", "conditionalFormat",
                ("kind", "cellIs"), ("operator", ">="), ("value", 20), ("fill", "FFC7CE"), ("bold", true)),
            AddOp("/Sheet1/B1:B10", "conditionalFormat",
                ("kind", "colorScale"),
                ("minColor", "F8696B"), ("midColor", "FFEB84"), ("maxColor", "63BE7B"))).IsOk);

        var cellIs = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("conditionalFormat", cellIs["kind"]!.GetValue<string>());
        Assert.Equal("cellIs", cellIs["cfKind"]!.GetValue<string>());
        Assert.Equal(">=", cellIs["operator"]!.GetValue<string>());
        Assert.Equal("20", cellIs["value"]!.GetValue<string>());
        Assert.Equal("FFC7CE", cellIs["fill"]!.GetValue<string>());
        Assert.True(cellIs["bold"]!.GetValue<bool>());
        Assert.Equal(["A1:C10"], cellIs["ranges"]!.AsArray().Select(n => n!.GetValue<string>()));

        var scale = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[2]"))));
        Assert.Equal("colorScale", scale["cfKind"]!.GetValue<string>());
        Assert.Equal("F8696B", scale["minColor"]!.GetValue<string>());
        Assert.Equal("FFEB84", scale["midColor"]!.GetValue<string>());
        Assert.Equal("63BE7B", scale["maxColor"]!.GetValue<string>());
    }

    [Fact]
    public void Read_structure_lists_rules_with_ranges_and_kinds()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1/A1:A10", "conditionalFormat",
                ("kind", "dataBar"), ("color", "638EC6")),
            AddOp("/Sheet1/D1:D3", "conditionalFormat",
                ("kind", "containsText"), ("text", "overdue"), ("fill", "FFC7CE"))).IsOk);

        var data = OkData(Handler.Read(Ctx(file, ("view", "structure"))));

        var formats = data["sheets"]![0]!["conditionalFormats"]!.AsArray();
        Assert.Equal(2, formats.Count);
        Assert.Equal("/Sheet1/conditionalFormat[1]", formats[0]!["path"]!.GetValue<string>());
        Assert.Equal("dataBar", formats[0]!["cfKind"]!.GetValue<string>());
        Assert.Equal(["A1:A10"], formats[0]!["ranges"]!.AsArray().Select(n => n!.GetValue<string>()));
        Assert.Equal("containsText", formats[1]!["cfKind"]!.GetValue<string>());
        Assert.Equal("overdue", formats[1]!["text"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_shifts_indices_and_keeps_the_file_clean()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1/A1:A10", "conditionalFormat",
                ("kind", "cellIs"), ("operator", ">"), ("value", 5), ("fill", "FFC7CE")),
            AddOp("/Sheet1/B1:B10", "conditionalFormat",
                ("kind", "colorScale"), ("minColor", "FFFFFF"), ("maxColor", "63BE7B"))).IsOk);

        var remove = EditOps(file, RemoveOp("/Sheet1/conditionalFormat[1]"));
        Assert.True(remove.IsOk, remove.ToJson());
        AssertValidatorClean(file);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("colorScale", data["cfKind"]!.GetValue<string>()); // shifted into [1]

        var missing = Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[2]")));
        Assert.False(missing.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, missing.Error!.Code);
        Assert.Contains("/Sheet1/conditionalFormat[1]", missing.Error.Candidates!);
    }

    [Fact]
    public void Removing_a_dataBar_leaves_no_orphan_x14_twin()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"))).IsOk);

        Assert.True(EditOps(file, RemoveOp("/Sheet1/conditionalFormat[1]")).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var worksheetPart = (WorksheetPart)document.WorkbookPart!.GetPartById(
            document.WorkbookPart!.Workbook!.Descendants<S.Sheet>().First().Id!.Value!);
        Assert.Empty(worksheetPart.Worksheet!.Descendants<S.ConditionalFormattingRule>());
        Assert.Empty(worksheetPart.Worksheet!.Descendants<X14.ConditionalFormattingRule>());
        Assert.Empty(worksheetPart.Worksheet!.Descendants<X14.ConditionalFormattings>());
    }

    [Fact]
    public void Remove_missing_rule_is_invalid_path_with_candidates()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"))).IsOk);

        var envelope = EditOps(file, RemoveOp("/Sheet1/conditionalFormat[5]"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/Sheet1/conditionalFormat[1]", envelope.Error.Candidates!);
    }

    [Fact]
    public void Single_cell_path_works_as_a_one_cell_range()
    {
        var file = CreateDataWorkbook();

        Assert.True(EditOps(file, AddOp("/Sheet1/A1", "conditionalFormat",
            ("kind", "cellIs"), ("operator", "=="), ("value", 1), ("fill", "C6EFCE"))).IsOk);

        AssertValidatorClean(file);
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal(["A1:A1"], data["ranges"]!.AsArray().Select(n => n!.GetValue<string>()));
    }

    [Fact]
    public void Bad_color_fails_actionably()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "reddish")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("hex", envelope.Error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Set_on_a_conditionalFormat_is_a_typed_unsupported_feature()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"))).IsOk);

        var envelope = EditOps(file, SetOp("/Sheet1/conditionalFormat[1]", ("fill", "FFC7CE")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Contains("remove", envelope.Error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DataBar_survives_a_later_unrelated_edit_still_clean()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/F1", ("value", "touched"))).IsOk);

        AssertValidatorClean(file);
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("dataBar", data["cfKind"]!.GetValue<string>());
        Assert.Equal("638EC6", data["color"]!.GetValue<string>());
    }

    // ----- data-bar thresholds + showValue (v1.15) --------------------------------

    /// <summary>The base dataBar element of the first cfRule on the first sheet.</summary>
    private static S.DataBar FirstDataBar(SpreadsheetDocument document) =>
        RawRules(document).Select(r => r.GetFirstChild<S.DataBar>()).First(b => b is not null)!;

    /// <summary>The single x14 dataBar twin on the first sheet (null when absent).</summary>
    private static X14.DataBar? FirstX14DataBar(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().First();
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return worksheetPart.Worksheet!.Descendants<X14.ConditionalFormattingRule>()
            .Select(r => r.GetFirstChild<X14.DataBar>())
            .FirstOrDefault(b => b is not null);
    }

    [Fact]
    public void DataBar_no_threshold_props_keeps_the_v114_auto_cfvo_byte_stable()
    {
        // Two files, identical except the second names the byte-stable default
        // path explicitly is impossible (there is no v1.14 binary here), so prove
        // the SHAPE instead: auto min/max cfvo, showValue="1", and no raw rewrite.
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "FF0000"))).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var bar = FirstDataBar(document);
        var cfvos = bar.Elements<S.ConditionalFormatValueObject>().ToList();
        Assert.Equal(["min", "max"], cfvos.Select(c => c.Type!.InnerText));
        Assert.Equal(["0", "0"], cfvos.Select(c => c.Val!.Value));
        // The default leaves ClosedXML's showValue="1" untouched (never "0").
        Assert.Equal("1", bar.ShowValue!.Value ? "1" : "0");
        Assert.Equal("FFFF0000", bar.GetFirstChild<S.Color>()!.Rgb!.Value);

        // get reports the bar as an auto bar with the value shown.
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("auto", data["minType"]!.GetValue<string>());
        Assert.Equal("auto", data["maxType"]!.GetValue<string>());
        Assert.True(data["showValue"]!.GetValue<bool>());
    }

    [Fact]
    public void DataBar_fixed_min_percentile_max_writes_both_cfvo_and_roundtrips()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"),
            ("minType", "fixed"), ("minValue", 0),
            ("maxType", "percentile"), ("maxValue", 90),
            ("showValue", false))).IsOk);

        AssertValidatorClean(file);
        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var bar = FirstDataBar(document);
            var cfvos = bar.Elements<S.ConditionalFormatValueObject>().ToList();
            Assert.Equal(["num", "percentile"], cfvos.Select(c => c.Type!.InnerText));
            Assert.Equal(["0", "90"], cfvos.Select(c => c.Val!.Value));
            Assert.False(bar.ShowValue!.Value);

            // The x14 twin (authoritative for Excel) carries the same endpoints.
            var x14 = FirstX14DataBar(document)!;
            var x14cfvos = x14.Elements<X14.ConditionalFormattingValueObject>().ToList();
            Assert.Equal(["num", "percentile"], x14cfvos.Select(c => c.Type!.InnerText));
            Assert.Equal(["0", "90"], x14cfvos.Select(c => c.GetFirstChild<Xm.Formula>()!.Text));
            Assert.False(x14.ShowValue!.Value);
        }

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("fixed", data["minType"]!.GetValue<string>());
        Assert.Equal("0", data["minValue"]!.GetValue<string>());
        Assert.Equal("percentile", data["maxType"]!.GetValue<string>());
        Assert.Equal("90", data["maxValue"]!.GetValue<string>());
        Assert.False(data["showValue"]!.GetValue<bool>());
    }

    [Fact]
    public void DataBar_showValue_false_writes_zero_omitted_keeps_default()
    {
        var off = CreateDataWorkbookNamed("off.xlsx");
        Assert.True(EditOps(off, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"), ("showValue", false))).IsOk);
        AssertValidatorClean(off);
        using (var document = SpreadsheetDocument.Open(off, isEditable: false))
        {
            Assert.False(FirstDataBar(document).ShowValue!.Value);
            // showValue-only is still a pure auto bar (min/max cfvo untouched).
            Assert.Equal(["min", "max"],
                FirstDataBar(document).Elements<S.ConditionalFormatValueObject>().Select(c => c.Type!.InnerText));
        }

        var on = CreateDataWorkbookNamed("on.xlsx");
        Assert.True(EditOps(on, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"))).IsOk);
        using (var document = SpreadsheetDocument.Open(on, isEditable: false))
        {
            // Omitted -> ClosedXML's showValue="1", never "0".
            Assert.True(FirstDataBar(document).ShowValue!.Value);
        }
    }

    [Fact]
    public void DataBar_formula_min_writes_formula_cfvo()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"),
            ("minType", "formula"), ("minValue", "$A$1"))).IsOk);

        AssertValidatorClean(file);
        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var min = FirstDataBar(document).Elements<S.ConditionalFormatValueObject>().First();
            Assert.Equal("formula", min.Type!.InnerText);
            Assert.Equal("$A$1", min.Val!.Value);

            var x14min = FirstX14DataBar(document)!.Elements<X14.ConditionalFormattingValueObject>().First();
            Assert.Equal("formula", x14min.Type!.InnerText);
            Assert.Equal("$A$1", x14min.GetFirstChild<Xm.Formula>()!.Text);
        }

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("formula", data["minType"]!.GetValue<string>());
        Assert.Equal("$A$1", data["minValue"]!.GetValue<string>());
    }

    [Fact]
    public void DataBar_leading_equals_formula_is_normalized()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"),
            ("maxType", "formula"), ("maxValue", "=$B$1"))).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var max = FirstDataBar(document).Elements<S.ConditionalFormatValueObject>().Last();
        Assert.Equal("formula", max.Type!.InnerText);
        Assert.Equal("$B$1", max.Val!.Value); // leading '=' stripped, like the formula kind
    }

    [Theory]
    [InlineData("minType")]
    [InlineData("maxType")]
    public void DataBar_invalid_bound_type_is_invalid_args_with_candidates(string typeKey)
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"), (typeKey, "invalid")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Equal(["auto", "fixed", "percent", "percentile", "formula"], envelope.Error.Candidates!);
    }

    [Fact]
    public void DataBar_formula_without_value_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"), ("minType", "formula")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("formula", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DataBar_non_numeric_fixed_value_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"), ("minType", "fixed"), ("minValue", "abc")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("number", envelope.Error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("percent", 150)]
    [InlineData("percentile", -5)]
    public void DataBar_percent_out_of_range_is_invalid_args(string type, int value)
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"), ("minType", type), ("minValue", value)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("between 0 and 100", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DataBar_value_without_type_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"), ("minValue", 5)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("minType", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DataBar_thresholds_survive_save_reopen_and_a_later_edit()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "dataBar"), ("color", "638EC6"),
            ("minType", "fixed"), ("minValue", 10),
            ("maxType", "percent"), ("maxValue", 80))).IsOk);

        // A later unrelated edit reopens with ClosedXML, strips/re-saves the model
        // and re-runs every post-save pass (including the data-bar fix-up that
        // drops orphaned x14 dataBars). The authored thresholds must survive.
        Assert.True(EditOps(file, SetOp("/Sheet1/F1", ("value", "touched"))).IsOk);

        AssertValidatorClean(file);
        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var cfvos = FirstDataBar(document).Elements<S.ConditionalFormatValueObject>().ToList();
            Assert.Equal(["num", "percent"], cfvos.Select(c => c.Type!.InnerText));
            Assert.Equal(["10", "80"], cfvos.Select(c => c.Val!.Value));

            // The x14 twin still pairs and still carries the endpoints.
            var x14 = FirstX14DataBar(document);
            Assert.NotNull(x14);
            Assert.Equal(["num", "percent"],
                x14!.Elements<X14.ConditionalFormattingValueObject>().Select(c => c.Type!.InnerText));
            Assert.Equal(["10", "80"],
                x14.Elements<X14.ConditionalFormattingValueObject>().Select(c => c.GetFirstChild<Xm.Formula>()!.Text));
        }

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("fixed", data["minType"]!.GetValue<string>());
        Assert.Equal("10", data["minValue"]!.GetValue<string>());
        Assert.Equal("percent", data["maxType"]!.GetValue<string>());
        Assert.Equal("80", data["maxValue"]!.GetValue<string>());
    }
}
