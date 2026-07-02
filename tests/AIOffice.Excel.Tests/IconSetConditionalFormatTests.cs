using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The v1.1 iconSet conditional-format kind (additive to cellIs | colorScale |
/// dataBar | containsText): each cell gets one of 3/4/5 icons by where its value
/// falls against evenly-spaced percent thresholds. ClosedXML authors iconSet
/// rules natively as base cfRules (no x14 twin), so they need no GUID fix-up.
/// Raw assertions reopen the file with the OpenXml SDK; every mutating test ends
/// OpenXmlValidator-clean.
/// </summary>
public sealed class IconSetConditionalFormatTests : ExcelTestBase
{
    /// <summary>Numbers 1..10 down column A.</summary>
    private string CreateDataWorkbook()
    {
        var file = CreateWorkbook();
        var grid = new JsonArray();
        for (var r = 1; r <= 10; r++)
        {
            grid.Add(new JsonArray(r * 10));
        }

        Assert.True(EditOps(file, SetOp("/Sheet1/A1:A10", ("values", grid))).IsOk);
        return file;
    }

    private static S.ConditionalFormattingRule SingleRawRule(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().First();
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return Assert.Single(worksheetPart.Worksheet!.Descendants<S.ConditionalFormattingRule>());
    }

    [Theory]
    [InlineData("3TrafficLights1", 3)]
    [InlineData("3Arrows", 3)]
    [InlineData("3Flags", 3)]
    [InlineData("4Rating", 4)]
    [InlineData("4RedToBlack", 4)]
    [InlineData("5Quarters", 5)]
    [InlineData("5Rating", 5)]
    public void IconSet_writes_rule_with_set_and_even_percent_thresholds(string set, int iconCount)
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", set)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal(
            "/Sheet1/conditionalFormat[1]",
            Json(envelope)["data"]!["ops"]![0]!["path"]!.GetValue<string>());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var rule = SingleRawRule(document);
        Assert.Equal(S.ConditionalFormatValues.IconSet, rule.Type!.Value);

        var iconSet = rule.GetFirstChild<S.IconSet>()!;
        Assert.Equal(set, iconSet.IconSetValue!.InnerText);

        // One cfvo per icon; thresholds evenly spaced and starting at 0.
        var thresholds = iconSet.Elements<S.ConditionalFormatValueObject>()
            .Select(o => o.Val!.Value).ToList();
        Assert.Equal(iconCount, thresholds.Count);
        Assert.Equal("0", thresholds[0]);
        var expected = iconCount switch
        {
            3 => new[] { "0", "33", "67" },
            4 => new[] { "0", "25", "50", "75" },
            _ => new[] { "0", "20", "40", "60", "80" },
        };
        Assert.Equal(expected, thresholds);
        Assert.All(
            iconSet.Elements<S.ConditionalFormatValueObject>(),
            o => Assert.Equal(S.ConditionalFormatValueObjectValues.Percent, o.Type!.Value));
    }

    [Fact]
    public void IconSet_reverse_and_showValue_false_round_trip()
    {
        var file = CreateDataWorkbook();

        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "3Symbols"), ("reverse", true), ("showValue", false))).IsOk);

        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var iconSet = SingleRawRule(document).GetFirstChild<S.IconSet>()!;
        Assert.True(iconSet.Reverse!.Value);
        Assert.False(iconSet.ShowValue!.Value);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("iconSet", data["cfKind"]!.GetValue<string>());
        Assert.Equal("3Symbols", data["set"]!.GetValue<string>());
        Assert.True(data["reverse"]!.GetValue<bool>());
        Assert.False(data["showValue"]!.GetValue<bool>());
    }

    [Fact]
    public void IconSet_default_showValue_is_true()
    {
        var file = CreateDataWorkbook();

        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "4Arrows"))).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var iconSet = SingleRawRule(document).GetFirstChild<S.IconSet>()!;

        // showValue defaults to true (the cell keeps its number visible).
        Assert.True(iconSet.ShowValue is null || iconSet.ShowValue.Value);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.True(data["showValue"]!.GetValue<bool>());
        Assert.Null(data["reverse"]); // only emitted when reversed
    }

    [Fact]
    public void IconSet_lists_with_other_kinds_in_structure()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1/A1:A10", "conditionalFormat", ("kind", "iconSet"), ("set", "3TrafficLights1")),
            AddOp("/Sheet1/A1:A10", "conditionalFormat",
                ("kind", "cellIs"), ("operator", ">"), ("value", 50), ("fill", "FFC7CE"))).IsOk);

        var formats = OkData(Handler.Read(Ctx(file, ("view", "structure"))))["sheets"]![0]!["conditionalFormats"]!.AsArray();
        Assert.Equal(2, formats.Count);
        Assert.Equal("iconSet", formats[0]!["cfKind"]!.GetValue<string>());
        Assert.Equal("3TrafficLights1", formats[0]!["set"]!.GetValue<string>());
        Assert.Equal("cellIs", formats[1]!["cfKind"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_iconSet_keeps_the_file_clean()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "5Rating"))).IsOk);

        Assert.True(EditOps(file, RemoveOp("/Sheet1/conditionalFormat[1]")).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().First();
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        Assert.Empty(worksheetPart.Worksheet!.Descendants<S.ConditionalFormattingRule>());
    }

    [Fact]
    public void IconSet_survives_a_later_unrelated_edit_still_clean()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "3Arrows"))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "touched"))).IsOk);

        AssertValidatorClean(file);
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("iconSet", data["cfKind"]!.GetValue<string>());
        Assert.Equal("3Arrows", data["set"]!.GetValue<string>());
    }

    [Fact]
    public void Missing_set_is_invalid_args_with_candidates()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat", ("kind", "iconSet")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("set", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("3TrafficLights1", envelope.Error.Candidates!);
    }

    [Fact]
    public void Unsupported_set_name_is_unsupported_feature_listing_supported_sets()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "6Diamonds")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Contains("3TrafficLights1", envelope.Error.Candidates!);
        Assert.Contains("5Quarters", envelope.Error.Candidates!);
        Assert.Contains("3TrafficLights1", envelope.Error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void IconSet_is_now_a_supported_conditionalFormat_kind()
    {
        // The pre-1.1 suite asserted iconSet was unsupported; confirm it is now
        // accepted while a genuinely-unknown kind still names the expanded set
        // (which v1.3 grew with formula/topBottom/aboveBelowAverage and v1.21
        // with duplicateValues/uniqueValues).
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat", ("kind", "timePeriod")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Equal(
            ["cellIs", "colorScale", "dataBar", "containsText", "iconSet", "formula", "topBottom",
                "aboveBelowAverage", "duplicateValues", "uniqueValues"],
            envelope.Error.Candidates!);
        Assert.Contains("iconSet", envelope.Error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Unknown_iconSet_prop_is_rejected()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "3Arrows"), ("color", "FF0000")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("color", envelope.Error.Message, StringComparison.Ordinal);
    }

    // ----- v1.20 caller-controlled per-icon thresholds --------------------------

    /// <summary>One {type, value} threshold entry for the thresholds array.</summary>
    private static JsonObject Th(string type, JsonNode value) =>
        new() { ["type"] = type, ["value"] = value };

    private static JsonArray Thresholds(params JsonObject[] entries)
    {
        var array = new JsonArray();
        foreach (var entry in entries)
        {
            array.Add(entry);
        }

        return array;
    }

    /// <summary>The base iconSet cfvo (in document order) of the sole raw rule.</summary>
    private static List<S.ConditionalFormatValueObject> IconSetCfvos(SpreadsheetDocument document) =>
        SingleRawRule(document).GetFirstChild<S.IconSet>()!
            .Elements<S.ConditionalFormatValueObject>().ToList();

    [Fact]
    public void IconSet_custom_thresholds_write_percent_cfvo_in_document_order()
    {
        var file = CreateDataWorkbook();

        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "3Arrows"),
            ("thresholds", Thresholds(Th("percent", 10), Th("percent", 50), Th("percent", 90))))).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var cfvos = IconSetCfvos(document);
        Assert.Equal(3, cfvos.Count);
        Assert.All(cfvos, c => Assert.Equal(S.ConditionalFormatValueObjectValues.Percent, c.Type!.Value));
        Assert.Equal(["10", "50", "90"], cfvos.Select(c => c.Val!.Value));
    }

    [Fact]
    public void IconSet_mixed_threshold_types_map_to_cfvo_types()
    {
        var file = CreateDataWorkbook();

        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "3Symbols"),
            ("thresholds", Thresholds(
                Th("num", 0), Th("percentile", 50), Th("formula", "=$A$1"))))).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var cfvos = IconSetCfvos(document);
        Assert.Equal(
            [
                S.ConditionalFormatValueObjectValues.Number,
                S.ConditionalFormatValueObjectValues.Percentile,
                S.ConditionalFormatValueObjectValues.Formula,
            ],
            cfvos.Select(c => c.Type!.Value));
        // The formula cfvo stores @val WITHOUT the leading '='.
        Assert.Equal(["0", "50", "$A$1"], cfvos.Select(c => c.Val!.Value));
    }

    [Fact]
    public void IconSet_threshold_count_must_match_icon_count()
    {
        var file = CreateDataWorkbook();

        // A 4-icon set with only 3 thresholds.
        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "4Arrows"),
            ("thresholds", Thresholds(Th("percent", 0), Th("percent", 33), Th("percent", 67)))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("4", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void IconSet_unknown_threshold_type_is_invalid_args_with_candidates()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "3Arrows"),
            ("thresholds", Thresholds(Th("percent", 0), Th("bogus", 50), Th("percent", 90)))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("percent", envelope.Error.Candidates!);
        Assert.Contains("formula", envelope.Error.Candidates!);
    }

    [Theory]
    [InlineData("percent", 150)]
    [InlineData("percentile", -5)]
    public void IconSet_percent_threshold_out_of_range_is_invalid_args(string type, int value)
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "3Arrows"),
            ("thresholds", Thresholds(Th("percent", 0), Th(type, value), Th("percent", 90)))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("between 0 and 100", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void IconSet_threshold_missing_value_is_invalid_args()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "3Arrows"),
            ("thresholds", Thresholds(
                Th("percent", 0), new JsonObject { ["type"] = "num" }, Th("percent", 90)))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("value", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void IconSet_formula_threshold_needs_a_string_value()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "3Arrows"),
            ("thresholds", Thresholds(Th("percent", 0), Th("percent", 50), Th("formula", 90)))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("formula", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void IconSet_default_rule_emits_no_thresholds_key()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "3Arrows"))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("iconSet", data["cfKind"]!.GetValue<string>());
        Assert.Null(data["thresholds"]); // a legacy even-split iconSet projects no thresholds
    }

    [Fact]
    public void IconSet_custom_thresholds_survive_save_reopen_and_a_later_edit()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat",
            ("kind", "iconSet"), ("set", "3Arrows"),
            ("thresholds", Thresholds(
                Th("num", 0), Th("percent", 50), Th("formula", "=$A$1"))))).IsOk);

        // A later unrelated edit reopens with ClosedXML, strips/re-saves the model
        // and re-runs every post-save pass. The authored thresholds must survive
        // (ClosedXML's FixUpAfterSave must not strip/reorder the cfvo).
        Assert.True(EditOps(file, SetOp("/Sheet1/F1", ("value", "touched"))).IsOk);

        AssertValidatorClean(file);
        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var cfvos = IconSetCfvos(document);
            Assert.Equal(3, cfvos.Count); // no stray/duplicated cfvo
            Assert.Equal(
                [
                    S.ConditionalFormatValueObjectValues.Number,
                    S.ConditionalFormatValueObjectValues.Percent,
                    S.ConditionalFormatValueObjectValues.Formula,
                ],
                cfvos.Select(c => c.Type!.Value));
            Assert.Equal(["0", "50", "$A$1"], cfvos.Select(c => c.Val!.Value));
        }

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        var thresholds = data["thresholds"]!.AsArray();
        Assert.Equal(3, thresholds.Count);
        Assert.Equal("num", thresholds[0]!["type"]!.GetValue<string>());
        Assert.Equal("0", thresholds[0]!["value"]!.GetValue<string>());
        Assert.Equal("percent", thresholds[1]!["type"]!.GetValue<string>());
        Assert.Equal("50", thresholds[1]!["value"]!.GetValue<string>());
        Assert.Equal("formula", thresholds[2]!["type"]!.GetValue<string>());
        Assert.Equal("$A$1", thresholds[2]!["value"]!.GetValue<string>());
    }
}
