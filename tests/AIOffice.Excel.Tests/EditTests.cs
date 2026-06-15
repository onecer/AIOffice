using AIOffice.Core;
using ClosedXML.Excel;
using Xunit;

namespace AIOffice.Excel.Tests;

public sealed class EditTests : ExcelTestBase
{
    [Fact]
    public void Set_values_and_sum_saves_cached_value_in_v_element()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 1)),
            SetOp("/Sheet1/A2", ("value", 2)),
            SetOp("/Sheet1/A3", ("value", "=SUM(A1:A2)")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Null(envelope.Meta.Warnings); // every listed function evaluated

        // The flagship check: reopen RAW with the OpenXml SDK; the file itself
        // must carry the computed value so agents see results without Excel.
        var raw = RawCell(file, "Sheet1", "A3");
        Assert.Equal("SUM(A1:A2)", raw.Formula);
        Assert.Equal("3", raw.CachedValue);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Set_auto_types_values_roundtrip()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(
            file,
            SetOp("/Sheet1/B1", ("value", 42.5)),
            SetOp("/Sheet1/B2", ("value", true)),
            SetOp("/Sheet1/B3", ("value", "2024-05-01")),
            SetOp("/Sheet1/B4", ("value", "hello")),
            SetOp("/Sheet1/B5", ("value", "0123"), ("valueType", "text")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var b1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B1"))));
        Assert.Equal("number", b1["type"]!.GetValue<string>());
        Assert.Equal(42.5, b1["value"]!.GetValue<double>());

        var b2 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))));
        Assert.Equal("boolean", b2["type"]!.GetValue<string>());
        Assert.True(b2["value"]!.GetValue<bool>());

        var b3 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B3"))));
        Assert.Equal("dateTime", b3["type"]!.GetValue<string>());
        Assert.Equal("2024-05-01T00:00:00", b3["value"]!.GetValue<string>());

        var b4 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B4"))));
        Assert.Equal("text", b4["type"]!.GetValue<string>());
        Assert.Equal("hello", b4["value"]!.GetValue<string>());

        // valueType:text is the escape hatch that keeps "0123" out of auto-typing.
        var b5 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B5"))));
        Assert.Equal("text", b5["type"]!.GetValue<string>());
        Assert.Equal("0123", b5["value"]!.GetValue<string>());

        AssertValidatorClean(file);
    }

    [Fact]
    public void Unsupported_function_warns_and_saves_formula_without_cached_value()
    {
        var file = CreateWorkbook();

        // A stored LAMBDA passed to MAP cannot be evaluated by the built-in engine
        // and aioffice does not special-case it; it rides the honest
        // formula_not_evaluated path. (XLOOKUP/IFS/SWITCH/LET/MAXIFS/… and the
        // dynamic arrays are now EVALUATED — see ScalarFunctionTests /
        // DynamicArrayTests — so they no longer warn.)
        var envelope = EditOps(file, SetOp("/Sheet1/C1", ("value", "=MAP(A1:A3,LAMBDA(x,x*2))")));

        Assert.True(envelope.IsOk, envelope.ToJson()); // the edit itself succeeds
        var warning = Assert.Single(envelope.Meta.Warnings!);
        Assert.Equal(ErrorCodes.FormulaNotEvaluated, warning.Code);
        Assert.Contains("C1", warning.Message, StringComparison.Ordinal);

        // Honesty on disk: formula text is saved, but the stale #NAME? cached
        // value is stripped so Excel recalculates the cell on open.
        var raw = RawCell(file, "Sheet1", "C1");
        Assert.NotNull(raw.Formula);
        Assert.Contains("LAMBDA", raw.Formula, StringComparison.Ordinal);
        Assert.Null(raw.CachedValue);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Merge_range_persists()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "Title")),
            SetOp("/Sheet1/A1:B2", ("merge", true)));
        Assert.True(envelope.IsOk, envelope.ToJson());

        using (var workbook = new XLWorkbook(file))
        {
            var merged = Assert.Single(workbook.Worksheet("Sheet1").MergedRanges);
            Assert.Equal("A1:B2", merged.RangeAddress.ToString());
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Expect_rev_mismatch_fails_with_stale_address_before_any_write()
    {
        var file = CreateWorkbook();
        var bytesBefore = File.ReadAllBytes(file);

        var envelope = Handler.Edit(
            Ctx(file, ("expectRev", "000000000000")),
            [SetOp("/Sheet1/A1", ("value", 1))]);

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.StaleAddress, envelope.Error!.Code);
        Assert.Equal(bytesBefore, File.ReadAllBytes(file)); // untouched
        Assert.Empty(Snapshots.List(file)); // no pre-image either
    }

    [Fact]
    public void Expect_rev_match_allows_the_edit()
    {
        var file = CreateWorkbook();
        var rev = Rev.OfFile(file);

        var envelope = Handler.Edit(
            Ctx(file, ("expectRev", rev)),
            [SetOp("/Sheet1/A1", ("value", 1))]);

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.NotEqual(rev, envelope.Meta.Rev); // meta reports the new rev
    }

    [Fact]
    public void Successful_edit_snapshots_the_pre_image()
    {
        var file = CreateWorkbook();
        var revBefore = Rev.OfFile(file);

        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("value", 1)));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var snapshot = Assert.Single(Snapshots.List(file));
        Assert.Equal(revBefore, snapshot.Rev);
    }

    [Fact]
    public void Dry_run_reports_ops_without_writing()
    {
        var file = CreateWorkbook();
        var revBefore = Rev.OfFile(file);

        var envelope = Handler.Edit(
            Ctx(file, ("dryRun", true)),
            [SetOp("/Sheet1/A1", ("value", 1))]);

        Assert.True(envelope.IsOk, envelope.ToJson());
        var data = Json(envelope)["data"]!;
        Assert.True(data["dryRun"]!.GetValue<bool>());
        Assert.Equal(1, data["wouldApply"]!.GetValue<int>());
        Assert.Equal(revBefore, Rev.OfFile(file)); // not written
        Assert.Empty(Snapshots.List(file));
    }

    [Fact]
    public void Move_is_a_typed_unsupported_feature()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, new EditOp { Op = "move", Path = "/Sheet1/A1" });

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(envelope.Error.Suggestion)); // workaround named
    }

    [Fact]
    public void Failing_op_aborts_the_whole_batch_atomically()
    {
        var file = CreateWorkbook();
        var bytesBefore = File.ReadAllBytes(file);

        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 1)),
            SetOp("/Nope/B2", ("value", 2))); // unknown sheet

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.NotNull(envelope.Error.Candidates); // invalid_path must carry candidates
        Assert.Contains("/Sheet1", envelope.Error.Candidates);
        Assert.Equal(bytesBefore, File.ReadAllBytes(file)); // op 1 was rolled back with the batch
    }

    [Fact]
    public void Add_sheet_add_row_insert_and_remove_row()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(
            file,
            AddOp("/Data", "sheet"),
            AddOp("/Data", "row", ("values", new System.Text.Json.Nodes.JsonArray(1, "two", true))),
            AddOp("/Data", "row", ("values", new System.Text.Json.Nodes.JsonArray(2, "three", false))));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var row1 = OkData(Handler.Get(Ctx(file, ("path", "/Data/row[1]"))));
        Assert.Equal(1.0, row1["values"]![0]!.GetValue<double>());
        Assert.Equal("two", row1["values"]![1]!.GetValue<string>());

        var removeEnvelope = EditOps(file, RemoveOp("/Data/row[1]"));
        Assert.True(removeEnvelope.IsOk, removeEnvelope.ToJson());

        // Row 2 shifted up into row 1.
        var shifted = OkData(Handler.Get(Ctx(file, ("path", "/Data/row[1]"))));
        Assert.Equal(2.0, shifted["values"]![0]!.GetValue<double>());

        AssertValidatorClean(file);
    }

    [Fact]
    public void Removing_the_last_sheet_is_refused()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, RemoveOp("/Sheet1"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("at least one sheet", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Remove_range_clears_contents_but_keeps_the_cells()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:B2", ("values", new System.Text.Json.Nodes.JsonArray(
                new System.Text.Json.Nodes.JsonArray(1, 2),
                new System.Text.Json.Nodes.JsonArray(3, 4))))).IsOk);

        var envelope = EditOps(file, RemoveOp("/Sheet1/A1:B2"));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var a1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        Assert.Equal("blank", a1["type"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Documented_function_support_baseline_holds()
    {
        // Pins the claim in ExcelHandler's doc comment. If a ClosedXML upgrade
        // changes either list, the docs must be updated with this test.
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 10)),
            SetOp("/Sheet1/A2", ("value", 20)),
            SetOp("/Sheet1/A3", ("value", 30)),
            SetOp("/Sheet1/D1", ("value", "key1")),
            SetOp("/Sheet1/E1", ("value", 100)),
            SetOp("/Sheet1/D2", ("value", "key2")),
            SetOp("/Sheet1/E2", ("value", 200))).IsOk);

        // Documented as evaluated: cached values must land on disk, no warning.
        var supported = EditOps(
            file,
            SetOp("/Sheet1/G1", ("value", "=SUM(A1:A3)")),
            SetOp("/Sheet1/G2", ("value", "=AVERAGE(A1:A3)")),
            SetOp("/Sheet1/G3", ("value", "=IF(A1>5,\"big\",\"small\")")),
            SetOp("/Sheet1/G4", ("value", "=VLOOKUP(\"key2\",D1:E2,2,FALSE)")),
            SetOp("/Sheet1/G5", ("value", "=INDEX(E1:E2,2)")),
            SetOp("/Sheet1/G6", ("value", "=MATCH(\"key2\",D1:D2,0)")),
            SetOp("/Sheet1/G7", ("value", "=TEXT(A1,\"0.00\")")),
            SetOp("/Sheet1/G8", ("value", "=DATE(2024,5,1)")),
            SetOp("/Sheet1/G9", ("value", "=COUNTIF(A1:A3,\">15\")")),
            // v1.5: XLOOKUP (and the other scalar functions) are now EVALUATED.
            SetOp("/Sheet1/G10", ("value", "=XLOOKUP(\"key2\",D1:D2,E1:E2)")));
        Assert.True(supported.IsOk, supported.ToJson());
        Assert.Null(supported.Meta.Warnings);
        Assert.Equal("60", RawCell(file, "Sheet1", "G1").CachedValue);
        Assert.Equal("2", RawCell(file, "Sheet1", "G6").CachedValue);
        Assert.Equal("10.00", RawCell(file, "Sheet1", "G7").CachedValue);
        Assert.Equal("2", RawCell(file, "Sheet1", "G9").CachedValue);
        Assert.Equal("200", RawCell(file, "Sheet1", "G10").CachedValue);

        // Documented as NOT evaluated: a stored LAMBDA passed to MAP cannot be
        // evaluated headlessly; warning raised, formula saved, no stale value.
        var unsupported = EditOps(file, SetOp("/Sheet1/H1", ("value", "=MAP(A1:A3,LAMBDA(x,x*2))")));
        Assert.True(unsupported.IsOk, unsupported.ToJson());
        var warning = Assert.Single(unsupported.Meta.Warnings!);
        Assert.Equal(ErrorCodes.FormulaNotEvaluated, warning.Code);
        Assert.Contains("H1", warning.Message, StringComparison.Ordinal);
        Assert.Null(RawCell(file, "Sheet1", "H1").CachedValue);

        AssertValidatorClean(file);
    }

    [Fact]
    public void Unknown_set_prop_lists_the_supported_ones()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("underline", "single")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("fill", envelope.Error.Candidates!);
    }
}
