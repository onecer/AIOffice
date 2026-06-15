using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

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
    private string CreateDataWorkbook()
    {
        var file = CreateWorkbook();
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

        // iconSet joined the supported set in v1.1; an unknown kind still names
        // every supported kind (including iconSet) as candidates.
        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "topBottom")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Equal(["cellIs", "colorScale", "dataBar", "containsText", "iconSet"], envelope.Error.Candidates!);
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
}
