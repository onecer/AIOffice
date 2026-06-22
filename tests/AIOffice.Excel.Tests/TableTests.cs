using System.IO;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// M6 Excel tables (ListObjects): create with a built-in style + header/totals/
/// banding flags, totals-row functions, structured-reference formulas that
/// evaluate, get by name, remove-keeps-data, and structure listing — all
/// validator-clean.
/// </summary>
public sealed class TableTests : ExcelTestBase
{
    /// <summary>Seeds a 3-column Sales table's data (headers + two data rows).</summary>
    private string SeedSalesData(string name = "book.xlsx")
    {
        var file = CreateWorkbook(name);
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:C3", ("values", new JsonArray(
            new JsonArray("Region", "Sales", "Cost"),
            new JsonArray("North", 100, 60),
            new JsonArray("South", 200, 90))))).IsOk);
        return file;
    }

    [Fact]
    public void Create_table_with_style_and_banding_reports_columns()
    {
        var file = SeedSalesData();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:C3", "table",
            ("name", "Sales"), ("style", "medium2"), ("bandedRows", true)));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var data = OkData(envelope)["ops"]![0]!;
        Assert.Equal("table", data["type"]!.GetValue<string>());
        Assert.Equal("Sales", data["name"]!.GetValue<string>());
        Assert.Equal("/Sheet1/table[@name=Sales]", data["path"]!.GetValue<string>());
        Assert.Equal("TableStyleMedium2", data["style"]!.GetValue<string>());
        Assert.Equal(["Region", "Sales", "Cost"], data["columns"]!.AsArray().Select(c => c!.GetValue<string>()));
        Assert.True(data["headerRow"]!.GetValue<bool>());

        // The ListObject really exists in the package.
        using var workbook = new XLWorkbook(file);
        var table = workbook.Worksheet("Sheet1").Tables.Single();
        Assert.Equal("Sales", table.Name);
        Assert.Equal(XLTableTheme.TableStyleMedium2.Name, table.Theme.Name);
        Assert.True(table.ShowRowStripes);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Style_none_makes_a_plain_table()
    {
        var file = SeedSalesData();

        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C3", "table",
            ("name", "Sales"), ("style", "none"))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/table[@name=Sales]"))));
        Assert.Equal("none", data["style"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Unknown_style_is_rejected_with_accepted_forms()
    {
        var file = SeedSalesData();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:C3", "table",
            ("name", "Sales"), ("style", "fancy7")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("fancy7", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("medium", envelope.Error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Totals_row_function_sets_a_subtotal_that_evaluates()
    {
        var file = SeedSalesData();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:C3", "table",
            ("name", "Sales"),
            ("totalsRow", true),
            ("totals", new JsonObject { ["Sales"] = "sum", ["Cost"] = "average" })));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var data = OkData(envelope)["ops"]![0]!;
        Assert.True(data["totalsRow"]!.GetValue<bool>());
        Assert.Equal("sum", data["totals"]!["Sales"]!.GetValue<string>());

        // The totals row landed below the data (row 4) and computes.
        using var workbook = new XLWorkbook(file);
        var table = workbook.Worksheet("Sheet1").Tables.Single();
        Assert.True(table.ShowTotalsRow);
        Assert.Equal(XLTotalsRowFunction.Sum, table.Field("Sales").TotalsRowFunction);
        Assert.Equal(300.0, table.Field("Sales").TotalsCell.Value.GetNumber());
        Assert.Equal(75.0, table.Field("Cost").TotalsCell.Value.GetNumber());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Totals_on_unknown_column_lists_real_columns()
    {
        var file = SeedSalesData();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:C3", "table",
            ("name", "Sales"),
            ("totals", new JsonObject { ["Profit"] = "sum" })));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("Profit", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("Region", envelope.Error.Candidates ?? [], StringComparer.Ordinal);
    }

    [Fact]
    public void Unknown_totals_function_is_rejected()
    {
        var file = SeedSalesData();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:C3", "table",
            ("name", "Sales"),
            ("totals", new JsonObject { ["Sales"] = "median" })));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("median", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Structured_reference_formula_evaluates()
    {
        var file = SeedSalesData();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C3", "table", ("name", "Sales"))).IsOk);

        // =SUM(Sales[Sales]) into a free cell evaluates against the table column.
        Assert.True(EditOps(file, SetOp("/Sheet1/E1", ("value", "=SUM(Sales[Sales])"))).IsOk);

        var raw = RawCell(file, "Sheet1", "E1");
        Assert.Equal("SUM(Sales[Sales])", raw.Formula);
        Assert.Equal("300", raw.CachedValue);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Get_table_describes_range_style_and_columns()
    {
        var file = SeedSalesData();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C3", "table",
            ("name", "Sales"), ("style", "medium9"))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/table[@name=Sales]"))));
        Assert.Equal("table", data["kind"]!.GetValue<string>());
        Assert.Equal("Sales", data["name"]!.GetValue<string>());
        Assert.Equal("A1:C3", data["range"]!.GetValue<string>());
        Assert.Equal("TableStyleMedium9", data["style"]!.GetValue<string>());
        Assert.Equal(3, data["columns"]!.AsArray().Count);
        Assert.Equal("Region", data["columns"]![0]!["name"]!.GetValue<string>());
    }

    [Fact]
    public void Get_missing_table_lists_candidates()
    {
        var file = SeedSalesData();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C3", "table", ("name", "Sales"))).IsOk);

        var envelope = Handler.Get(Ctx(file, ("path", "/Sheet1/table[@name=Ghost]")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/Sheet1/table[@name=Sales]", envelope.Error.Candidates ?? [], StringComparer.Ordinal);
    }

    [Fact]
    public void Remove_table_drops_the_listobject_but_keeps_the_data()
    {
        var file = SeedSalesData();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C3", "table", ("name", "Sales"))).IsOk);

        var envelope = EditOps(file, RemoveOp("/Sheet1/table[@name=Sales]"));
        Assert.True(envelope.IsOk, envelope.ToJson());
        var data = OkData(envelope)["ops"]![0]!;
        Assert.Equal("table", data["removed"]!.GetValue<string>());
        Assert.Equal("data kept", data["note"]!.GetValue<string>());

        using var workbook = new XLWorkbook(file);
        var sheet = workbook.Worksheet("Sheet1");
        Assert.Empty(sheet.Tables); // the table is gone
        Assert.Equal("Region", sheet.Cell("A1").GetString()); // …but the data stays
        Assert.Equal(200.0, sheet.Cell("B3").Value.GetNumber());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Set_on_a_table_is_rejected_with_a_remove_readd_hint()
    {
        var file = SeedSalesData();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C3", "table", ("name", "Sales"))).IsOk);

        var envelope = EditOps(file, SetOp("/Sheet1/table[@name=Sales]", ("name", "Renamed")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Contains("table", envelope.Error.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ----- v1.17: edit a table's totals-row function/label after creation -----

    /// <summary>Seeds a Sales table (no totals row) ready for a set-totals op.</summary>
    private string SeedSalesTable(string name = "book.xlsx")
    {
        var file = SeedSalesData(name);
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C3", "table", ("name", "Sales"))).IsOk);
        return file;
    }

    [Fact]
    public void Set_totals_function_and_label_after_creation_round_trips()
    {
        var file = SeedSalesTable();

        // A function on one column and a custom label on another, in one op. Both
        // round-trip through reopen; setting the totals turns the totals row on.
        var envelope = EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject
            {
                ["Sales"] = new JsonObject { ["function"] = "sum" },
                ["Region"] = new JsonObject { ["label"] = "Total Qty" },
            })));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var data = OkData(envelope)["ops"]![0]!;
        Assert.Equal("table", data["type"]!.GetValue<string>());
        Assert.True(data["totalsRow"]!.GetValue<bool>());
        Assert.Equal("sum", data["totals"]!["Sales"]!["function"]!.GetValue<string>());
        Assert.Equal("Total Qty", data["totals"]!["Region"]!["label"]!.GetValue<string>());

        using var workbook = new XLWorkbook(file);
        var table = workbook.Worksheet("Sheet1").Tables.Single();
        Assert.True(table.ShowTotalsRow);
        Assert.Equal(XLTotalsRowFunction.Sum, table.Field("Sales").TotalsRowFunction);
        Assert.Equal("Total Qty", table.Field("Region").TotalsRowLabel);
        AssertValidatorClean(file);
    }

    /// <summary>
    /// A custom label and a built-in function are mutually exclusive in a totals
    /// cell (Excel's own model). When one op sets both on the SAME column the
    /// label wins (the OOXML reader gives a label precedence), and the response
    /// reflects only the label that actually persisted.
    /// </summary>
    [Fact]
    public void Set_totals_both_on_one_column_lets_the_label_win()
    {
        var file = SeedSalesTable();

        var envelope = EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Sales"] = new JsonObject { ["function"] = "sum", ["label"] = "Total Qty" } })));
        Assert.True(envelope.IsOk, envelope.ToJson());
        var sales = OkData(envelope)["ops"]![0]!["totals"]!["Sales"]!;
        Assert.Equal("Total Qty", sales["label"]!.GetValue<string>());
        Assert.Null(sales["function"]);

        using var workbook = new XLWorkbook(file);
        var field = workbook.Worksheet("Sheet1").Tables.Single().Field("Sales");
        Assert.Equal(XLTotalsRowFunction.None, field.TotalsRowFunction);
        Assert.Equal("Total Qty", field.TotalsRowLabel);
        AssertValidatorClean(file);
    }

    /// <summary>
    /// Setting a function on a column that previously held a custom label replaces
    /// the label (a cell is one or the other). The function persists clean.
    /// </summary>
    [Fact]
    public void Set_totals_function_replaces_an_existing_label()
    {
        var file = SeedSalesTable();
        Assert.True(EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Sales"] = new JsonObject { ["label"] = "Old Label" } }))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Sales"] = new JsonObject { ["function"] = "average" } }))).IsOk);

        using var workbook = new XLWorkbook(file);
        var field = workbook.Worksheet("Sheet1").Tables.Single().Field("Sales");
        Assert.Equal(XLTotalsRowFunction.Average, field.TotalsRowFunction);
        Assert.True(string.IsNullOrEmpty(field.TotalsRowLabel));
        AssertValidatorClean(file);
    }

    /// <summary>
    /// Setting a custom label on a column that previously held a function replaces
    /// the function. The label persists clean.
    /// </summary>
    [Fact]
    public void Set_totals_label_replaces_an_existing_function()
    {
        var file = SeedSalesTable();
        Assert.True(EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Sales"] = new JsonObject { ["function"] = "sum" } }))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Sales"] = new JsonObject { ["label"] = "Grand" } }))).IsOk);

        using var workbook = new XLWorkbook(file);
        var field = workbook.Worksheet("Sheet1").Tables.Single().Field("Sales");
        Assert.Equal(XLTotalsRowFunction.None, field.TotalsRowFunction);
        Assert.Equal("Grand", field.TotalsRowLabel);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Set_totals_function_none_clears_the_function()
    {
        var file = SeedSalesTable();
        Assert.True(EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Sales"] = new JsonObject { ["function"] = "sum" } }))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Sales"] = new JsonObject { ["function"] = "none" } }))).IsOk);

        using var workbook = new XLWorkbook(file);
        Assert.Equal(XLTotalsRowFunction.None,
            workbook.Worksheet("Sheet1").Tables.Single().Field("Sales").TotalsRowFunction);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Set_totals_empty_label_clears_the_label()
    {
        var file = SeedSalesTable();
        Assert.True(EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Region"] = new JsonObject { ["label"] = "Totals" } }))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Region"] = new JsonObject { ["label"] = "" } }))).IsOk);

        using var workbook = new XLWorkbook(file);
        Assert.True(string.IsNullOrEmpty(
            workbook.Worksheet("Sheet1").Tables.Single().Field("Region").TotalsRowLabel));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Set_totals_multiple_columns_in_one_op()
    {
        var file = SeedSalesTable();

        Assert.True(EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject
            {
                ["Region"] = new JsonObject { ["label"] = "Total" },
                ["Sales"] = new JsonObject { ["function"] = "sum" },
                ["Cost"] = new JsonObject { ["function"] = "average" },
            }))).IsOk);

        using var workbook = new XLWorkbook(file);
        var table = workbook.Worksheet("Sheet1").Tables.Single();
        Assert.True(table.ShowTotalsRow);
        Assert.Equal("Total", table.Field("Region").TotalsRowLabel);
        Assert.Equal(XLTotalsRowFunction.Sum, table.Field("Sales").TotalsRowFunction);
        Assert.Equal(XLTotalsRowFunction.Average, table.Field("Cost").TotalsRowFunction);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Get_round_trips_the_totals_function_and_label()
    {
        var file = SeedSalesTable();
        Assert.True(EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject
            {
                ["Sales"] = new JsonObject { ["function"] = "sum" },
                ["Region"] = new JsonObject { ["label"] = "Grand Total" },
            }))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/table[@name=Sales]"))));
        var columns = data["columns"]!.AsArray();
        var salesColumn = columns.Single(c => c!["name"]!.GetValue<string>() == "Sales")!;
        var regionColumn = columns.Single(c => c!["name"]!.GetValue<string>() == "Region")!;
        Assert.Equal("sum", salesColumn["totalsFunction"]!.GetValue<string>());
        Assert.Equal("Grand Total", regionColumn["totalsLabel"]!.GetValue<string>());
    }

    [Fact]
    public void Set_totals_unknown_column_lists_real_columns()
    {
        var file = SeedSalesTable();

        var envelope = EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Profit"] = new JsonObject { ["function"] = "sum" } })));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("Profit", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("Region", envelope.Error.Candidates ?? [], StringComparer.Ordinal);
    }

    [Fact]
    public void Set_totals_unknown_function_lists_the_function_candidates()
    {
        var file = SeedSalesTable();

        var envelope = EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Sales"] = new JsonObject { ["function"] = "median" } })));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("median", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("sum", envelope.Error.Candidates ?? [], StringComparer.Ordinal);
    }

    [Fact]
    public void Set_totals_non_string_function_is_invalid_args()
    {
        var file = SeedSalesTable();

        var envelope = EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Sales"] = new JsonObject { ["function"] = 5 } })));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("sum", envelope.Error.Candidates ?? [], StringComparer.Ordinal);
    }

    [Fact]
    public void Set_totals_non_string_label_is_invalid_args()
    {
        var file = SeedSalesTable();

        var envelope = EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Region"] = new JsonObject { ["label"] = 42 } })));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    [Fact]
    public void Set_totals_empty_column_object_is_invalid_args()
    {
        var file = SeedSalesTable();

        // {} carries neither function nor label — reject rather than silently flip the totals row on.
        var envelope = EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Sales"] = new JsonObject() })));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    /// <summary>
    /// PIN: a table set carrying ANY non-'totals' prop still hits the blanket
    /// guard with the identical UnsupportedFeature error — the relaxation never
    /// over-reaches. Locks against accidental widening of the totals carve-out.
    /// </summary>
    [Fact]
    public void Table_SetWithoutTotals_StillUnsupported()
    {
        var file = SeedSalesTable();

        // bold alone on a table path: still unsupported.
        var bold = EditOps(file, SetOp("/Sheet1/table[@name=Sales]", ("bold", true)));
        Assert.False(bold.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, bold.Error!.Code);
        Assert.Contains("table", bold.Error.Message, StringComparison.OrdinalIgnoreCase);

        // totals + a second prop is NOT a totals-only set → still unsupported.
        var mixed = EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject { ["Sales"] = new JsonObject { ["function"] = "sum" } }),
            ("bold", true)));
        Assert.False(mixed.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, mixed.Error!.Code);
        Assert.Equal(bold.Error.Message, mixed.Error.Message);
    }

    /// <summary>
    /// BYTE-STABLE: a table with existing totals, re-saved by an edit that does
    /// NOT carry a totals set op, leaves the table-definition part byte-identical
    /// (the guard relaxation fires ONLY on a totals-only set).
    /// </summary>
    [Fact]
    public void Existing_totals_survive_a_non_totals_resave_byte_for_byte()
    {
        var file = SeedSalesTable();
        Assert.True(EditOps(file, SetOp("/Sheet1/table[@name=Sales]",
            ("totals", new JsonObject
            {
                ["Sales"] = new JsonObject { ["function"] = "sum" },
                ["Region"] = new JsonObject { ["label"] = "T" },
            }))).IsOk);

        var before = TableDefinitionBytes(file);

        // A benign cell set re-saves the workbook without touching the table path.
        Assert.True(EditOps(file, SetOp("/Sheet1/E1", ("value", 1))).IsOk);

        Assert.Equal(before, TableDefinitionBytes(file));
    }

    /// <summary>Reads the raw bytes of the single table-definition part.</summary>
    private static byte[] TableDefinitionBytes(string file)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var part = document.WorkbookPart!.WorksheetParts
            .SelectMany(w => w.TableDefinitionParts)
            .Single();
        using var stream = part.GetStream();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    [Fact]
    public void Header_row_can_be_turned_off()
    {
        var file = SeedSalesData();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:C3", "table",
            ("name", "Sales"), ("headerRow", false)));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.False(OkData(envelope)["ops"]![0]!["headerRow"]!.GetValue<bool>());

        using var workbook = new XLWorkbook(file);
        Assert.False(workbook.Worksheet("Sheet1").Tables.Single().ShowHeaderRow);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Structure_view_lists_tables_with_style_and_totals()
    {
        var file = SeedSalesData();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C3", "table",
            ("name", "Sales"), ("style", "medium2"), ("totalsRow", true))).IsOk);

        var data = OkData(Handler.Read(Ctx(file, ("view", "structure"))));
        var table = data["sheets"]![0]!["tables"]![0]!;
        Assert.Equal("Sales", table["name"]!.GetValue<string>());
        Assert.Equal("/Sheet1/table[@name=Sales]", table["path"]!.GetValue<string>());
        Assert.Equal("TableStyleMedium2", table["style"]!.GetValue<string>());
        Assert.True(table["totalsRow"]!.GetValue<bool>());
    }

    [Fact]
    public void Quoted_table_name_round_trips_through_the_path()
    {
        var file = SeedSalesData();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C3", "table", ("name", "Q3_Sales"))).IsOk);

        // A name with an underscore is a bare name; confirm get resolves it.
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/table[@name=Q3_Sales]"))));
        Assert.Equal("Q3_Sales", data["name"]!.GetValue<string>());
    }
}
