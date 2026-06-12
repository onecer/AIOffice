using System.IO.Compression;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// M5 sparklines. Oracles: the raw x14 extLst XML (sparklineGroups with xm:f /
/// xm:sqref pairs — read straight from the zip part, not through ClosedXML)
/// and the schema validator, which must stay at zero errors (the post-save
/// pass strips ClosedXML's stray xr2:uid attribute).
/// </summary>
public sealed class SparklineTests : ExcelTestBase
{
    private static EditOp SparkOp(string path, params (string Key, JsonNode? Value)[] props) =>
        AddOp(path, "sparkline", props);

    private string SeededWorkbook()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A2:D3", ("values", new JsonArray(
            new JsonArray(1, 3, 2, 5),
            new JsonArray(-1, 2, -3, 4))))).IsOk);
        return file;
    }

    private static string RawSheetXml(string file)
    {
        using var zip = ZipFile.OpenRead(file);
        using var stream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    [Fact]
    public void Add_line_sparkline_writes_the_x14_ext_list()
    {
        var file = SeededWorkbook();
        var envelope = EditOps(file, SparkOp(
            "/Sheet1/E2",
            ("dataRange", "A2:D2"), ("kind", "line"), ("color", "FF0000"), ("markers", true)));
        Assert.True(envelope.IsOk, envelope.ToJson());
        var detail = Json(envelope)["data"]!["ops"]![0]!;
        Assert.Equal("/Sheet1/sparkline[1]", detail["path"]!.GetValue<string>());
        Assert.Equal("E2", detail["cell"]!.GetValue<string>());

        var xml = RawSheetXml(file);
        Assert.Contains("extLst", xml, StringComparison.Ordinal); // groups live in the worksheet extLst
        Assert.Contains("sparklineGroups", xml, StringComparison.Ordinal);
        Assert.Contains("type=\"line\"", xml, StringComparison.Ordinal);
        Assert.Contains("markers=\"1\"", xml, StringComparison.Ordinal);
        Assert.Contains("colorSeries rgb=\"FFFF0000\"", xml, StringComparison.Ordinal);
        Assert.Contains("f>Sheet1!A2:D2</", xml, StringComparison.Ordinal); // xm:f
        Assert.Contains("sqref>E2</", xml, StringComparison.Ordinal); // xm:sqref
        Assert.DoesNotContain("xr2:uid", xml, StringComparison.Ordinal); // the fix-up stripped it

        AssertValidatorClean(file);
    }

    [Fact]
    public void Column_and_winLoss_kinds_map_to_their_ooxml_types()
    {
        var file = SeededWorkbook();
        Assert.True(EditOps(
            file,
            SparkOp("/Sheet1/E2", ("dataRange", "A2:D2"), ("kind", "column")),
            SparkOp("/Sheet1/E3", ("dataRange", "A3:D3"), ("kind", "winLoss"))).IsOk);

        var xml = RawSheetXml(file);
        Assert.Contains("type=\"column\"", xml, StringComparison.Ordinal);
        Assert.Contains("type=\"stacked\"", xml, StringComparison.Ordinal); // winLoss stores as stacked

        var got = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/sparkline[2]"))));
        Assert.Equal("winLoss", got["sparklineKind"]!.GetValue<string>());
        Assert.Equal("E3", got["cell"]!.GetValue<string>());
        Assert.Equal("A3:D3", got["dataRange"]!.GetValue<string>());

        AssertValidatorClean(file);
    }

    [Fact]
    public void Get_describes_a_sparkline_and_structure_lists_groups()
    {
        var file = SeededWorkbook();
        Assert.True(EditOps(file, SparkOp(
            "/Sheet1/F2", ("dataRange", "A2:D2"), ("color", "00B050"), ("markers", true))).IsOk);

        var got = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/sparkline[1]"))));
        Assert.Equal("sparkline", got["kind"]!.GetValue<string>());
        Assert.Equal("line", got["sparklineKind"]!.GetValue<string>());
        Assert.Equal("00B050", got["color"]!.GetValue<string>());
        Assert.True(got["markers"]!.GetValue<bool>());

        var structure = OkData(Handler.Read(Ctx(file, ("view", "structure"))));
        var groups = structure["sheets"]![0]!["sparklineGroups"]!.AsArray();
        var group = Assert.Single(groups)!;
        Assert.Equal("line", group["kind"]!.GetValue<string>());
        var member = Assert.Single(group["sparklines"]!.AsArray())!;
        Assert.Equal("/Sheet1/sparkline[1]", member["path"]!.GetValue<string>());
        Assert.Equal("F2", member["cell"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_deletes_the_sparkline_and_its_empty_group()
    {
        var file = SeededWorkbook();
        Assert.True(EditOps(file, SparkOp("/Sheet1/E2", ("dataRange", "A2:D2"))).IsOk);

        var removed = EditOps(file, RemoveOp("/Sheet1/sparkline[1]"));
        Assert.True(removed.IsOk, removed.ToJson());
        Assert.Equal("sparkline", Json(removed)["data"]!["ops"]![0]!["removed"]!.GetValue<string>());

        Assert.DoesNotContain("sparklineGroups", RawSheetXml(file), StringComparison.Ordinal);
        using (var workbook = new XLWorkbook(file))
        {
            Assert.Empty(workbook.Worksheet("Sheet1").SparklineGroups);
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Sparklines_survive_later_unrelated_edits()
    {
        var file = SeededWorkbook();
        Assert.True(EditOps(file, SparkOp("/Sheet1/E2", ("dataRange", "A2:D2"), ("color", "FF0000"))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/A9", ("value", "later"))).IsOk);

        var xml = RawSheetXml(file);
        Assert.Contains("<x14:sparklineGroups", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("xr2:uid", xml, StringComparison.Ordinal); // re-stripped on every save
        AssertValidatorClean(file);
    }

    [Fact]
    public void Bad_ops_fail_with_typed_errors()
    {
        var file = SeededWorkbook();

        var noRange = EditOps(file, SparkOp("/Sheet1/E2", ("kind", "line")));
        Assert.False(noRange.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, noRange.Error!.Code);
        Assert.Contains("dataRange", noRange.Error.Message, StringComparison.Ordinal);

        var badKind = EditOps(file, SparkOp("/Sheet1/E2", ("dataRange", "A2:D2"), ("kind", "pie")));
        Assert.False(badKind.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, badKind.Error!.Code);
        Assert.Contains("winLoss", badKind.Error.Candidates!);

        var markersOnColumn = EditOps(file, SparkOp(
            "/Sheet1/E2", ("dataRange", "A2:D2"), ("kind", "column"), ("markers", true)));
        Assert.False(markersOnColumn.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, markersOnColumn.Error!.Code);

        var badColor = EditOps(file, SparkOp("/Sheet1/E2", ("dataRange", "A2:D2"), ("color", "reddish")));
        Assert.False(badColor.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, badColor.Error!.Code);

        var onSheet = EditOps(file, SparkOp("/Sheet1", ("dataRange", "A2:D2")));
        Assert.False(onSheet.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, onSheet.Error!.Code);

        Assert.True(EditOps(file, SparkOp("/Sheet1/E2", ("dataRange", "A2:D2"))).IsOk);
        var duplicate = EditOps(file, SparkOp("/Sheet1/E2", ("dataRange", "A3:D3")));
        Assert.False(duplicate.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, duplicate.Error!.Code);
        Assert.Contains("already has a sparkline", duplicate.Error.Message, StringComparison.Ordinal);

        var missing = Handler.Get(Ctx(file, ("path", "/Sheet1/sparkline[9]")));
        Assert.False(missing.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, missing.Error!.Code);
        Assert.Contains("/Sheet1/sparkline[1]", missing.Error.Candidates!);

        var setOnSparkline = EditOps(file, SetOp("/Sheet1/sparkline[1]", ("value", 1)));
        Assert.False(setOnSparkline.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, setOnSparkline.Error!.Code);
    }
}
