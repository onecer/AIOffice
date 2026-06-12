using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// M4 bulk 2D writes: the anchor form (extent inferred), the strict range form
/// (mismatch names both shapes), typing parity with single-cell set, and the
/// SAX streaming path for &gt;50k cells into a bare sheet — including the
/// equality law (streamed and DOM paths produce the same workbook content).
/// </summary>
public sealed class BulkWriteTests : ExcelTestBase
{
    [Fact]
    public void Anchor_form_infers_extent_and_types_like_single_cell_set()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/B2", ("values", new JsonArray(
            new JsonArray(1, "x", true),
            new JsonArray(2.5, "2024-05-01", "=B2+B3")))));

        Assert.True(envelope.IsOk, envelope.ToJson());

        Assert.Equal("number", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))))["type"]!.GetValue<string>());
        Assert.Equal("x", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C2"))))["value"]!.GetValue<string>());
        Assert.True(OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D2"))))["value"]!.GetValue<bool>());
        Assert.Equal("dateTime", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C3"))))["type"]!.GetValue<string>());

        // The formula evaluated against the cells written by the SAME op.
        var raw = RawCell(file, "Sheet1", "D3");
        Assert.Equal("B2+B3", raw.Formula);
        Assert.Equal("3.5", raw.CachedValue);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Anchor_form_allows_ragged_rows()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("values", new JsonArray(
            new JsonArray("h1", "h2", "h3"),
            new JsonArray(1)))));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("h3", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))))["value"]!.GetValue<string>());
        Assert.Equal("blank", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))))["type"]!.GetValue<string>());
    }

    [Fact]
    public void Range_form_dimension_mismatch_states_both_shapes()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/A1:B3", ("values", new JsonArray(
            new JsonArray(1, 2),
            new JsonArray(3, 4)))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("2 row(s) x 2 column(s)", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("3 row(s) x 2 column(s)", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("A1:B3", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("anchor", envelope.Error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Range_form_ragged_row_states_both_shapes()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/A1:B2", ("values", new JsonArray(
            new JsonArray(1, 2),
            new JsonArray(3)))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("row 2 has 1 column(s)", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("2 row(s) x 2 column(s)", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Flat_array_on_a_cell_anchor_is_rejected_with_a_2d_hint()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("values", new JsonArray(1, 2, 3))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("row 1 is not an array", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Values_running_past_the_sheet_edge_are_rejected()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/XFD1", ("values", new JsonArray(
            new JsonArray(1, 2)))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("sheet edge", envelope.Error.Message, StringComparison.Ordinal);
    }

    // ----- the streaming performance path -------------------------------------

    private const int StreamRows = 20_001; // x3 columns = 60,003 cells > the 50k threshold

    private static JsonArray BuildBigPayload()
    {
        var rows = new JsonArray();
        for (var r = 1; r <= StreamRows; r++)
        {
            rows.Add(new JsonArray(r * 1.5, "Item_" + (r % 97), r % 2 == 0));
        }

        return rows;
    }

    private static string SheetXml(string file)
    {
        using var zip = ZipFile.OpenRead(file);
        var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;
        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    [Fact]
    public void Streaming_and_dom_paths_produce_identical_workbook_content()
    {
        // Streamed: one bulk op into a bare sheet (>50k cells) takes the SAX path.
        var streamed = CreateWorkbook("streamed.xlsx");
        var streamedEnvelope = EditOps(streamed, SetOp("/Sheet1/A1", ("values", BuildBigPayload())));
        Assert.True(streamedEnvelope.IsOk, streamedEnvelope.ToJson());

        // DOM: an earlier op makes the sheet non-bare, forcing the ClosedXML path
        // for the very same payload (the bulk write overwrites A1 with its own value).
        var dom = CreateWorkbook("dom.xlsx");
        var domEnvelope = EditOps(
            dom,
            SetOp("/Sheet1/A1", ("value", "placeholder")),
            SetOp("/Sheet1/A1", ("values", BuildBigPayload())));
        Assert.True(domEnvelope.IsOk, domEnvelope.ToJson());

        // Path fingerprints: the SAX writer emits inline strings; ClosedXML
        // uses the shared-string table. This proves which path each file took.
        Assert.Contains("inlineStr", SheetXml(streamed), StringComparison.Ordinal);
        Assert.DoesNotContain("inlineStr", SheetXml(dom), StringComparison.Ordinal);

        // The equality law: identical content, cell by cell, typed.
        using (var streamedWb = new XLWorkbook(streamed))
        using (var domWb = new XLWorkbook(dom))
        {
            var streamedSheet = streamedWb.Worksheet("Sheet1");
            var domSheet = domWb.Worksheet("Sheet1");
            Assert.Equal(
                domSheet.RangeUsed()!.RangeAddress.ToString(),
                streamedSheet.RangeUsed()!.RangeAddress.ToString());

            for (var r = 1; r <= StreamRows; r++)
            {
                for (var c = 1; c <= 3; c++)
                {
                    var expected = domSheet.Cell(r, c).Value;
                    var actual = streamedSheet.Cell(r, c).Value;
                    Assert.True(
                        expected.Equals(actual),
                        $"Cell ({r},{c}) differs: dom={expected} streamed={actual}");
                }
            }
        }

        AssertValidatorClean(streamed);
        AssertValidatorClean(dom);
    }

    [Fact]
    public void Streamed_formulas_get_cached_values_through_the_normal_pipeline()
    {
        var file = CreateWorkbook();
        var rows = new JsonArray();
        for (var r = 1; r <= StreamRows; r++)
        {
            // A handful of formulas ride along with the bulk literals.
            rows.Add(r <= 5
                ? new JsonArray(r, "x", $"=A{r}*2")
                : new JsonArray(r, "x", r * 2));
        }

        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("values", rows)));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Null(envelope.Meta.Warnings);

        var raw = RawCell(file, "Sheet1", "C3");
        Assert.Equal("A3*2", raw.Formula);
        Assert.Equal("6", raw.CachedValue); // evaluated by the post-stream ClosedXML pass
        Assert.Equal(
            8.0,
            OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C4"))))["value"]!.GetValue<double>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Later_op_on_the_same_sheet_flushes_the_queued_write_in_order()
    {
        var file = CreateWorkbook();

        // One batch: bulk write, THEN a single-cell overwrite of A1. Sequential
        // semantics demand the later op wins.
        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", BuildBigPayload())),
            SetOp("/Sheet1/A1", ("value", "header")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal(
            "header",
            OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))))["value"]!.GetValue<string>());
        Assert.Equal(
            3.0, // row 2: 2 * 1.5
            OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A2"))))["value"]!.GetValue<double>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Bulk_write_into_a_brand_new_sheet_in_the_same_batch_streams()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(
            file,
            new EditOp { Op = "add", Path = "/Big", Type = "sheet" },
            SetOp("/Big/A1", ("values", BuildBigPayload())));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal(
            StreamRows * 1.5,
            OkData(Handler.Get(Ctx(file, ("path", $"/Big/A{StreamRows}"))))["value"]!.GetValue<double>());
        Assert.Equal(
            "Item_1",
            OkData(Handler.Get(Ctx(file, ("path", "/Big/B1"))))["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Dates_fall_back_to_the_dom_path_and_stay_correct()
    {
        var file = CreateWorkbook();
        var rows = new JsonArray();
        for (var r = 1; r <= StreamRows; r++)
        {
            rows.Add(new JsonArray(r, "2024-05-01", r % 2 == 0)); // dates are not streamable
        }

        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("values", rows)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.DoesNotContain("inlineStr", SheetXml(file), StringComparison.Ordinal); // DOM path
        var b1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B1"))));
        Assert.Equal("dateTime", b1["type"]!.GetValue<string>());
        Assert.Equal("2024-05-01T00:00:00", b1["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }
}
