using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The M2 pivot-table slice (ClosedXML-native). Raw assertions reopen the file
/// with the OpenXml SDK and pin the pivotTableDefinition / pivotCacheDefinition
/// essentials, so ClosedXML cannot grade its own homework.
/// </summary>
public sealed class PivotTableTests : ExcelTestBase
{
    /// <summary>Region/Product/Quarter/Channel/Sales data in A1:E7.</summary>
    private string CreateDataWorkbook()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1:E7", ("values", new JsonArray(
                new JsonArray("Region", "Product", "Quarter", "Channel", "Sales"),
                new JsonArray("East", "Apple", "Q1", "Web", 100),
                new JsonArray("East", "Pear", "Q2", "Store", 150),
                new JsonArray("West", "Apple", "Q1", "Web", 200),
                new JsonArray("West", "Pear", "Q2", "Store", 250),
                new JsonArray("East", "Apple", "Q2", "Web", 120),
                new JsonArray("West", "Pear", "Q1", "Store", 80)))));
        Assert.True(envelope.IsOk, envelope.ToJson());
        return file;
    }

    private static EditOp PivotOp(
        string name = "SalesPivot",
        string sourceRange = "A1:E7",
        string targetSheet = "Pivot",
        params (string Key, JsonNode? Value)[] extra)
    {
        var props = new List<(string, JsonNode?)>
        {
            ("name", name),
            ("sourceRange", sourceRange),
            ("targetSheet", targetSheet),
        };
        props.AddRange(extra.Select(e => ((string)e.Key, e.Value)));
        return AddOp("/Sheet1", "pivot", [.. props]);
    }

    /// <summary>The raw pivotTableDefinition of the first pivot on the named sheet.</summary>
    private static (S.PivotTableDefinition Definition, S.PivotCacheDefinition Cache) RawPivot(
        SpreadsheetDocument document, string sheetName)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        var pivotPart = worksheetPart.PivotTableParts.First();
        return (pivotPart.PivotTableDefinition!, pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!);
    }

    [Fact]
    public void Add_pivot_writes_raw_definition_essentials_and_is_validator_clean()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, PivotOp(extra:
        [
            ("rows", new JsonArray("Region", "Product")),
            ("columns", new JsonArray("Quarter")),
            ("values", new JsonArray(new JsonObject { ["field"] = "Sales", ["agg"] = "sum" })),
            ("filters", new JsonArray("Channel")),
        ]));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var detail = Json(envelope)["data"]!["ops"]![0]!;
        Assert.Equal("/Pivot/pivot[@name=SalesPivot]", detail["path"]!.GetValue<string>());
        Assert.Equal("Sheet1", detail["sourceSheet"]!.GetValue<string>());
        Assert.Equal("A1:E7", detail["sourceRange"]!.GetValue<string>());
        Assert.Equal(["Region", "Product"], detail["rows"]!.AsArray().Select(n => n!.GetValue<string>()));

        AssertValidatorClean(file);

        // Raw oracle: the pivotTableDefinition must carry the configured axes,
        // the data field with the right agg, and a cache pointing at the source.
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var (definition, cache) = RawPivot(document, "Pivot");
        Assert.Equal("SalesPivot", definition.Name?.Value);

        var fieldNames = definition.PivotFields!.Elements<S.PivotField>().ToList();
        Assert.Equal(5, fieldNames.Count); // every cache field is materialized

        // Field indices are cache-ordered: Region=0 Product=1 Quarter=2 Channel=3 Sales=4.
        Assert.Equal(
            [0, 1],
            definition.RowFields!.Elements<S.Field>().Select(f => f.Index!.Value));
        Assert.Equal(
            [2],
            definition.ColumnFields!.Elements<S.Field>().Select(f => f.Index!.Value));
        Assert.Equal(
            [3],
            definition.PageFields!.Elements<S.PageField>().Select(f => (int)f.Field!.Value));

        var dataField = Assert.Single(definition.DataFields!.Elements<S.DataField>());
        Assert.Equal(4u, dataField.Field!.Value);
        Assert.Equal("Sum of Sales", dataField.Name?.Value);
        Assert.Null(dataField.Subtotal); // sum is the schema default

        var source = cache.CacheSource!.WorksheetSource!;
        Assert.Equal("Sheet1", source.Sheet?.Value);
        Assert.Equal("A1:E7", source.Reference?.Value);
        Assert.True(cache.RefreshOnLoad?.Value); // Excel recomputes the pivot on open
        Assert.True(cache.SaveData?.Value); // cache records saved with the file
        Assert.True(
            cache.CacheFields!.Elements<S.CacheField>().First().SharedItems!.ChildElements.Count > 0,
            "cache field items were not populated"); // refresh actually ran
    }

    [Theory]
    [InlineData("average", "average")]
    [InlineData("count", "count")]
    [InlineData("min", "min")]
    [InlineData("max", "max")]
    public void Non_sum_aggs_write_the_matching_subtotal(string agg, string rawSubtotal)
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, PivotOp(extra:
        [
            ("rows", new JsonArray("Region")),
            ("values", new JsonArray(new JsonObject { ["field"] = "Sales", ["agg"] = agg })),
        ])).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var (definition, _) = RawPivot(document, "Pivot");
        var dataField = Assert.Single(definition.DataFields!.Elements<S.DataField>());
        Assert.Equal(rawSubtotal, dataField.Subtotal!.InnerText);

        // And the wire agg round-trips through get.
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Pivot/pivot[1]"))));
        Assert.Equal(agg, data["values"]![0]!["agg"]!.GetValue<string>());
    }

    [Fact]
    public void Unknown_field_is_invalid_args_with_the_actual_headers_as_candidates()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, PivotOp(extra: [("rows", new JsonArray("Regin"))]));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("Regin", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Equal("Region", envelope.Error.Candidates![0]); // nearest match first
        Assert.Equal(
            ["Channel", "Product", "Quarter", "Region", "Sales"],
            envelope.Error.Candidates!.Order(StringComparer.Ordinal)); // all real headers offered
    }

    [Fact]
    public void Unknown_agg_is_invalid_args_with_candidates()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, PivotOp(extra:
        [
            ("values", new JsonArray(new JsonObject { ["field"] = "Sales", ["agg"] = "median" })),
        ]));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Equal(["sum", "count", "average", "min", "max"], envelope.Error.Candidates!);
    }

    [Fact]
    public void Field_on_two_axes_is_refused()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, PivotOp(extra:
        [
            ("rows", new JsonArray("Region")),
            ("columns", new JsonArray("Region")),
        ]));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("Region", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Missing_sourceRange_fails_actionably()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "pivot",
            ("targetSheet", "Pivot"), ("rows", new JsonArray("Region"))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("sourceRange", envelope.Error.Message, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(envelope.Error.Suggestion));
    }

    [Fact]
    public void Target_sheet_is_created_when_absent()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, PivotOp(targetSheet: "Summary", extra:
        [
            ("rows", new JsonArray("Region")),
            ("values", new JsonArray("Sales")), // bare string form: agg defaults to sum
        ])).IsOk);

        var outline = OkData(Handler.Read(Ctx(file, ("view", "outline"))));
        Assert.Contains(
            "Summary",
            outline["sheets"]!.AsArray().Select(s => s!["name"]!.GetValue<string>()));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Get_by_index_and_by_name_round_trips_the_configuration()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, PivotOp(extra:
        [
            ("rows", new JsonArray("Region", "Product")),
            ("columns", new JsonArray("Quarter")),
            ("values", new JsonArray(new JsonObject { ["field"] = "Sales", ["agg"] = "average" })),
            ("filters", new JsonArray("Channel")),
        ])).IsOk);

        foreach (var path in new[] { "/Pivot/pivot[1]", "/Pivot/pivot[@name=SalesPivot]", "/Pivot/pivot[@name='SalesPivot']" })
        {
            var data = OkData(Handler.Get(Ctx(file, ("path", path))));
            Assert.Equal("/Pivot/pivot[@name=SalesPivot]", data["path"]!.GetValue<string>());
            Assert.Equal("pivot", data["kind"]!.GetValue<string>());
            Assert.Equal("SalesPivot", data["name"]!.GetValue<string>());
            Assert.Equal("Sheet1", data["sourceSheet"]!.GetValue<string>());
            Assert.Equal("A1:E7", data["sourceRange"]!.GetValue<string>());
            Assert.Equal(["Region", "Product"], data["rows"]!.AsArray().Select(n => n!.GetValue<string>()));
            Assert.Equal(["Quarter"], data["columns"]!.AsArray().Select(n => n!.GetValue<string>()));
            Assert.Equal(["Channel"], data["filters"]!.AsArray().Select(n => n!.GetValue<string>()));
            Assert.Equal("Sales", data["values"]![0]!["field"]!.GetValue<string>());
            Assert.Equal("average", data["values"]![0]!["agg"]!.GetValue<string>());
            Assert.NotNull(data["location"]);
        }
    }

    [Fact]
    public void Get_missing_pivot_is_invalid_path_with_stable_name_candidates()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, PivotOp(extra: [("rows", new JsonArray("Region"))])).IsOk);

        var byIndex = Handler.Get(Ctx(file, ("path", "/Pivot/pivot[2]")));
        Assert.False(byIndex.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, byIndex.Error!.Code);
        Assert.Contains("/Pivot/pivot[@name=SalesPivot]", byIndex.Error.Candidates!);

        var byName = Handler.Get(Ctx(file, ("path", "/Pivot/pivot[@name=SalesPivto]")));
        Assert.False(byName.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, byName.Error!.Code);
        Assert.Contains("/Pivot/pivot[@name=SalesPivot]", byName.Error.Candidates!);
    }

    [Fact]
    public void Read_structure_lists_pivots()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, PivotOp(extra:
        [
            ("rows", new JsonArray("Region")),
            ("values", new JsonArray(new JsonObject { ["field"] = "Sales" })),
        ])).IsOk);

        var data = OkData(Handler.Read(Ctx(file, ("view", "structure"))));

        var sheets = data["sheets"]!.AsArray();
        var pivotSheet = sheets.Single(s => s!["name"]!.GetValue<string>() == "Pivot")!;
        var pivot = Assert.Single(pivotSheet["pivots"]!.AsArray())!;
        Assert.Equal("/Pivot/pivot[@name=SalesPivot]", pivot["path"]!.GetValue<string>());
        Assert.Equal("A1:E7", pivot["sourceRange"]!.GetValue<string>());
        Assert.Equal("sum", pivot["values"]![0]!["agg"]!.GetValue<string>());

        var sourceSheet = sheets.Single(s => s!["name"]!.GetValue<string>() == "Sheet1")!;
        Assert.Empty(sourceSheet["pivots"]!.AsArray());
    }

    [Fact]
    public void Remove_by_name_and_by_index_deletes_the_pivot()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(
            file,
            PivotOp(name: "P1", extra: [("rows", new JsonArray("Region"))]),
            PivotOp(name: "P2", extra: [("rows", new JsonArray("Product"))])).IsOk);

        var removeByName = EditOps(file, RemoveOp("/Pivot/pivot[@name=P1]"));
        Assert.True(removeByName.IsOk, removeByName.ToJson());
        AssertValidatorClean(file);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Pivot/pivot[1]"))));
        Assert.Equal("P2", data["name"]!.GetValue<string>()); // P2 shifted into pivot[1]

        Assert.True(EditOps(file, RemoveOp("/Pivot/pivot[1]")).IsOk);
        AssertValidatorClean(file);
        var structure = OkData(Handler.Read(Ctx(file, ("view", "structure"))));
        var pivotSheet = structure["sheets"]!.AsArray().Single(s => s!["name"]!.GetValue<string>() == "Pivot")!;
        Assert.Empty(pivotSheet["pivots"]!.AsArray());
    }

    [Fact]
    public void Remove_missing_pivot_is_invalid_path_with_candidates()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, PivotOp(extra: [("rows", new JsonArray("Region"))])).IsOk);

        var envelope = EditOps(file, RemoveOp("/Pivot/pivot[3]"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/Pivot/pivot[@name=SalesPivot]", envelope.Error.Candidates!);
    }

    [Fact]
    public void Duplicate_pivot_name_on_the_target_sheet_is_refused()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, PivotOp(extra: [("rows", new JsonArray("Region"))])).IsOk);

        var envelope = EditOps(file, PivotOp(extra: [("rows", new JsonArray("Product"))]));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("SalesPivot", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Pivot_with_no_fields_is_refused()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, PivotOp());

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("rows", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Set_on_a_pivot_is_a_typed_unsupported_feature()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, PivotOp(extra: [("rows", new JsonArray("Region"))])).IsOk);

        var envelope = EditOps(file, SetOp("/Pivot/pivot[1]", ("value", 1)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Contains("add", envelope.Error.Suggestion, StringComparison.Ordinal); // workaround named
    }

    [Fact]
    public void Pivot_survives_a_later_unrelated_edit()
    {
        // Pins the cooperation contract with ClosedXML's save pipeline: the
        // pivot parts it rewrites on every save must stay valid and complete.
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, PivotOp(extra:
        [
            ("rows", new JsonArray("Region")),
            ("values", new JsonArray(new JsonObject { ["field"] = "Sales", ["agg"] = "sum" })),
        ])).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/G1", ("value", "touched"))).IsOk);

        AssertValidatorClean(file);
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Pivot/pivot[@name=SalesPivot]"))));
        Assert.Equal(["Region"], data["rows"]!.AsArray().Select(n => n!.GetValue<string>()));
        Assert.Equal("A1:E7", data["sourceRange"]!.GetValue<string>());
    }

    [Fact]
    public void Quoted_pivot_name_with_spaces_round_trips()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, PivotOp(name: "Q3 Sales", extra: [("rows", new JsonArray("Region"))])).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Pivot/pivot[@name='Q3 Sales']"))));

        Assert.Equal("/Pivot/pivot[@name='Q3 Sales']", data["path"]!.GetValue<string>());
        Assert.Equal("Q3 Sales", data["name"]!.GetValue<string>());
    }

    [Fact]
    public void Failing_batch_with_pivot_op_leaves_no_trace()
    {
        var file = CreateDataWorkbook();
        var bytesBefore = File.ReadAllBytes(file);

        var envelope = EditOps(
            file,
            PivotOp(extra: [("rows", new JsonArray("Region"))]),
            SetOp("/Nope/A1", ("value", 1))); // unknown sheet aborts the batch

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Equal(bytesBefore, File.ReadAllBytes(file)); // atomic: nothing written
    }
}
