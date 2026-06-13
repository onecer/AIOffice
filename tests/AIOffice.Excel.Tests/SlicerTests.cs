using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace AIOffice.Excel.Tests;

/// <summary>
/// M8 Excel slicers (dashboard power feature). Slicers are authored entirely on
/// raw OpenXml (ClosedXML cannot create them), so every test reopens the file
/// with the OpenXml SDK and pins the slicerCache + slicer parts, then asserts the
/// OpenXmlValidator (Office2019) reports zero errors — ClosedXML never grades its
/// own homework.
/// </summary>
public sealed class SlicerTests : ExcelTestBase
{
    /// <summary>Seeds a Sales table over Region/Sales/Cost in A1:C4.</summary>
    private string SeedSalesTable(string name = "book.xlsx")
    {
        var file = CreateWorkbook(name);
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:C4", ("values", new JsonArray(
            new JsonArray("Region", "Sales", "Cost"),
            new JsonArray("North", 100, 60),
            new JsonArray("South", 200, 90),
            new JsonArray("East", 150, 70))))).IsOk);
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C4", "table", ("name", "Sales"))).IsOk);
        return file;
    }

    /// <summary>Seeds Region/Product/Sales data in A1:C7 plus a pivot on a Pivot sheet.</summary>
    private string SeedPivot(string name = "book.xlsx")
    {
        var file = CreateWorkbook(name);
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:C7", ("values", new JsonArray(
            new JsonArray("Region", "Product", "Sales"),
            new JsonArray("East", "Apple", 100),
            new JsonArray("West", "Apple", 200),
            new JsonArray("East", "Pear", 150),
            new JsonArray("West", "Pear", 250),
            new JsonArray("East", "Apple", 120),
            new JsonArray("West", "Pear", 80))))).IsOk);
        Assert.True(EditOps(file, AddOp("/Sheet1", "pivot",
            ("name", "SalesPivot"),
            ("sourceRange", "A1:C7"),
            ("targetSheet", "Pivot"),
            ("rows", new JsonArray("Region")),
            ("values", new JsonArray(new JsonObject { ["field"] = "Sales", ["agg"] = "sum" })))).IsOk);
        return file;
    }

    // ----- table slicer -------------------------------------------------------

    [Fact]
    public void Add_table_slicer_writes_raw_parts_and_is_validator_clean()
    {
        var file = SeedSalesTable();

        var envelope = EditOps(file, AddOp("/Sheet1/table[@name=Sales]", "slicer", ("column", "Region")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var op = OkData(envelope)["ops"]![0]!;
        Assert.Equal("slicer", op["type"]!.GetValue<string>());
        Assert.Equal("table", op["source"]!.GetValue<string>());
        Assert.Equal("Region", op["column"]!.GetValue<string>());
        Assert.Equal("/Sheet1/slicer[1]", op["path"]!.GetValue<string>());

        // The raw parts really exist in the package.
        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var workbookPart = document.WorkbookPart!;
            var cachePart = Assert.Single(workbookPart.SlicerCacheParts);
            Assert.Equal("Slicer_Region", cachePart.SlicerCacheDefinition!.Name!.Value);
            Assert.Equal("Region", cachePart.SlicerCacheDefinition!.SourceName!.Value);

            // A table slicer's connection lives in an x15 tableSlicerCache ext.
            var tableCache = cachePart.SlicerCacheDefinition!
                .Descendants<DocumentFormat.OpenXml.Office2013.Excel.TableSlicerCache>().Single();
            Assert.NotNull(tableCache.TableId);

            var worksheetPart = WorksheetPartByName(document, "Sheet1");
            var slicersPart = Assert.Single(worksheetPart.SlicersParts);
            var slicer = slicersPart.Slicers!.Elements<X14.Slicer>().Single();
            Assert.Equal("Slicer_Region", slicer.Name!.Value);
            Assert.Equal("Region", slicer.Caption!.Value);

            // The workbook extLst references the cache via an x15 slicerCaches ext.
            Assert.Single(workbookPart.Workbook!
                .Descendants<DocumentFormat.OpenXml.Office2013.Excel.SlicerCaches>());
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Add_table_slicer_with_caption_and_anchor()
    {
        var file = SeedSalesTable();

        var envelope = EditOps(file, AddOp("/Sheet1/table[@name=Sales]", "slicer",
            ("column", "Region"), ("caption", "Pick a region"), ("x", "F2")));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("Pick a region", OkData(envelope)["ops"]![0]!["caption"]!.GetValue<string>());

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var worksheetPart = WorksheetPartByName(document, "Sheet1");
            var slicer = worksheetPart.SlicersParts.Single().Slicers!.Elements<X14.Slicer>().Single();
            Assert.Equal("Pick a region", slicer.Caption!.Value);
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Slicer_on_a_missing_table_column_lists_candidates()
    {
        var file = SeedSalesTable();

        var envelope = EditOps(file, AddOp("/Sheet1/table[@name=Sales]", "slicer", ("column", "Nope")));
        Assert.False(envelope.IsOk);
        var error = Json(envelope)["error"]!;
        Assert.Equal("invalid_args", error["code"]!.GetValue<string>());
        var candidates = error["candidates"]!.AsArray().Select(c => c!.GetValue<string>()).ToList();
        Assert.Contains("Region", candidates);
    }

    [Fact]
    public void Slicer_on_a_bare_range_is_unsupported_feature()
    {
        var file = SeedSalesTable();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:C4", "slicer", ("column", "Region")));
        Assert.False(envelope.IsOk);
        Assert.Equal("unsupported_feature", Json(envelope)["error"]!["code"]!.GetValue<string>());
    }

    // ----- pivot slicer -------------------------------------------------------

    [Fact]
    public void Add_pivot_slicer_writes_raw_parts_and_is_validator_clean()
    {
        var file = SeedPivot();

        var envelope = EditOps(file, AddOp("/Pivot/pivot[@name=SalesPivot]", "slicer", ("field", "Region")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var op = OkData(envelope)["ops"]![0]!;
        Assert.Equal("pivot", op["source"]!.GetValue<string>());
        Assert.Equal("Region", op["column"]!.GetValue<string>());

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var workbookPart = document.WorkbookPart!;
            var cachePart = Assert.Single(workbookPart.SlicerCacheParts);
            // A pivot slicer connects via slicerCachePivotTables + a tabular cache.
            Assert.Single(cachePart.SlicerCacheDefinition!.Descendants<X14.SlicerCachePivotTable>());
            var tabular = cachePart.SlicerCacheDefinition!.Descendants<X14.TabularSlicerCache>().Single();
            // One item per distinct Region value (East, West).
            Assert.True(tabular.Elements<X14.TabularSlicerCacheItems>().Single().Count!.Value >= 1);

            // The workbook extLst references the cache via an x14 slicerCaches ext.
            Assert.Single(workbookPart.Workbook!
                .Descendants<X14.SlicerCaches>());
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Slicer_on_a_missing_pivot_field_lists_candidates()
    {
        var file = SeedPivot();

        var envelope = EditOps(file, AddOp("/Pivot/pivot[@name=SalesPivot]", "slicer", ("field", "Nope")));
        Assert.False(envelope.IsOk);
        var error = Json(envelope)["error"]!;
        Assert.Equal("invalid_args", error["code"]!.GetValue<string>());
        var candidates = error["candidates"]!.AsArray().Select(c => c!.GetValue<string>()).ToList();
        Assert.Contains("Region", candidates);
    }

    // ----- get / structure / remove -------------------------------------------

    [Fact]
    public void Get_slicer_by_index_and_by_name()
    {
        var file = SeedSalesTable();
        Assert.True(EditOps(file, AddOp("/Sheet1/table[@name=Sales]", "slicer", ("column", "Region"))).IsOk);

        var byIndex = Handler.Get(Ctx(file, ("path", "/Sheet1/slicer[1]")));
        var data = OkData(byIndex);
        Assert.Equal("slicer", data["kind"]!.GetValue<string>());
        Assert.Equal("Slicer_Region", data["name"]!.GetValue<string>());
        Assert.Equal("table", data["source"]!.GetValue<string>());
        Assert.Equal("Region", data["column"]!.GetValue<string>());

        var byName = Handler.Get(Ctx(file, ("path", "/Sheet1/slicer[@name=Slicer_Region]")));
        Assert.Equal("Region", OkData(byName)["column"]!.GetValue<string>());
    }

    [Fact]
    public void Structure_view_lists_slicers()
    {
        var file = SeedSalesTable();
        Assert.True(EditOps(file, AddOp("/Sheet1/table[@name=Sales]", "slicer", ("column", "Region"))).IsOk);

        var read = Handler.Read(Ctx(file, ("view", "structure")));
        var sheet = OkData(read)["sheets"]!.AsArray()[0]!;
        var slicers = sheet["slicers"]!.AsArray();
        Assert.Single(slicers);
        Assert.Equal("/Sheet1/slicer[1]", slicers[0]!["path"]!.GetValue<string>());
        Assert.Equal("Region", slicers[0]!["column"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_slicer_drops_parts_and_stays_valid()
    {
        var file = SeedSalesTable();
        Assert.True(EditOps(file, AddOp("/Sheet1/table[@name=Sales]", "slicer", ("column", "Region"))).IsOk);

        var envelope = EditOps(file, RemoveOp("/Sheet1/slicer[1]"));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("slicer", OkData(envelope)["ops"]![0]!["removed"]!.GetValue<string>());

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            Assert.Empty(document.WorkbookPart!.SlicerCacheParts);
            var worksheetPart = WorksheetPartByName(document, "Sheet1");
            Assert.True(worksheetPart.SlicersParts.All(p => !(p.Slicers?.Elements<X14.Slicer>().Any() ?? false)));
            // The table itself survives the slicer removal.
            Assert.Single(worksheetPart.TableDefinitionParts);
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Remove_missing_slicer_lists_candidates()
    {
        var file = SeedSalesTable();

        var envelope = EditOps(file, RemoveOp("/Sheet1/slicer[3]"));
        Assert.False(envelope.IsOk);
        Assert.Equal("invalid_path", Json(envelope)["error"]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void Two_slicers_on_one_sheet_get_distinct_names_and_indices()
    {
        var file = SeedSalesTable();

        Assert.True(EditOps(file,
            AddOp("/Sheet1/table[@name=Sales]", "slicer", ("column", "Region")),
            AddOp("/Sheet1/table[@name=Sales]", "slicer", ("column", "Sales"))).IsOk);

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var names = document.WorkbookPart!.SlicerCacheParts
                .Select(p => p.SlicerCacheDefinition!.Name!.Value)
                .OrderBy(n => n, StringComparer.Ordinal)
                .ToList();
            Assert.Equal(2, names.Count);
            Assert.Equal(names.Count, names.Distinct(StringComparer.OrdinalIgnoreCase).Count());
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Slicer_survives_a_later_unrelated_edit()
    {
        var file = SeedSalesTable();
        Assert.True(EditOps(file, AddOp("/Sheet1/table[@name=Sales]", "slicer", ("column", "Region"))).IsOk);

        // A later edit goes through ClosedXML's save; the slicer parts must
        // survive byte-clean (the same guarantee charts rely on).
        Assert.True(EditOps(file, SetOp("/Sheet1/A10", ("value", "footer"))).IsOk);

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            Assert.Single(document.WorkbookPart!.SlicerCacheParts);
            var worksheetPart = WorksheetPartByName(document, "Sheet1");
            Assert.Single(worksheetPart.SlicersParts.Single().Slicers!.Elements<X14.Slicer>());
        }

        AssertValidatorClean(file);
    }

    private static WorksheetPart WorksheetPartByName(SpreadsheetDocument document, string sheetName)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
    }
}
