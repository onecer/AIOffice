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
        // (which v1.3 grew with formula/topBottom/aboveBelowAverage).
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A10", "conditionalFormat", ("kind", "timePeriod")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Equal(
            ["cellIs", "colorScale", "dataBar", "containsText", "iconSet", "formula", "topBottom", "aboveBelowAverage"],
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
}
