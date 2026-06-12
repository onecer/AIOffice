using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// Streaming-vs-DOM equivalence, forced with <c>stream=true</c> on small files
/// so both paths read the same bytes fast. The SAX path must agree with the
/// ClosedXML DOM path on values, types, formulas, shared strings, stats and
/// truncation behavior — and its documented divergences (no number formats, so
/// dates surface as raw serial numbers) are pinned here too.
/// </summary>
public sealed class StreamingReadTests : ExcelTestBase
{
    /// <summary>A ClosedXML-written workbook with every value shape the scan must handle.</summary>
    private string CreateMixedWorkbook()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "hello world")),  // shared string
            SetOp("/Sheet1/B1", ("value", 42)),
            SetOp("/Sheet1/C1", ("value", 3.25)),
            SetOp("/Sheet1/D1", ("value", true)),
            SetOp("/Sheet1/A2", ("value", 1)),
            SetOp("/Sheet1/B2", ("value", 2)),
            SetOp("/Sheet1/C2", ("value", "=SUM(A2:B2)")),
            SetOp("/Sheet1/A3:C4", ("values", new JsonArray(
                new JsonArray("x", 10, 1.5),
                new JsonArray("y", 20, 2.5))))).IsOk);
        return file;
    }

    [Fact]
    public void Streamed_stats_equal_dom_stats()
    {
        var file = CreateMixedWorkbook();

        var dom = OkData(Handler.Read(Ctx(file, ("view", "stats"))));
        var streamed = OkData(Handler.Read(Ctx(file, ("view", "stats"), ("stream", true))));

        Assert.Null(dom["streamed"]);
        Assert.True(streamed["streamed"]!.GetValue<bool>());
        Assert.Equal(dom["totals"]!["cells"]!.GetValue<long>(), streamed["totals"]!["cells"]!.GetValue<long>());
        Assert.Equal(dom["totals"]!["formulas"]!.GetValue<long>(), streamed["totals"]!["formulas"]!.GetValue<long>());
        Assert.Equal(dom["totals"]!["sheets"]!.GetValue<int>(), streamed["totals"]!["sheets"]!.GetValue<int>());
        var domSheet = dom["sheets"]![0]!;
        var streamedSheet = streamed["sheets"]![0]!;
        Assert.Equal(domSheet["name"]!.GetValue<string>(), streamedSheet["name"]!.GetValue<string>());
        Assert.Equal(domSheet["position"]!.GetValue<int>(), streamedSheet["position"]!.GetValue<int>());
        Assert.Equal(domSheet["usedRange"]!.GetValue<string>(), streamedSheet["usedRange"]!.GetValue<string>());
        Assert.Equal(domSheet["cellCount"]!.GetValue<long>(), streamedSheet["cellCount"]!.GetValue<long>());
        Assert.Equal(domSheet["formulaCount"]!.GetValue<long>(), streamedSheet["formulaCount"]!.GetValue<long>());
    }

    [Fact]
    public void Streamed_cell_gets_agree_with_dom_on_every_value_shape()
    {
        var file = CreateMixedWorkbook();
        foreach (var address in new[] { "A1", "B1", "C1", "D1", "C2", "A3", "B4", "C4", "F9" })
        {
            var dom = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/" + address))));
            var streamed = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/" + address), ("stream", true))));

            Assert.True(streamed["streamed"]!.GetValue<bool>());
            Assert.Equal(dom["value"]?.ToJsonString(), streamed["value"]?.ToJsonString());
            Assert.Equal(dom["type"]!.GetValue<string>(), streamed["type"]!.GetValue<string>());
            Assert.Equal(dom["formula"]?.GetValue<string>(), streamed["formula"]?.GetValue<string>());
            Assert.Equal(dom["path"]!.GetValue<string>(), streamed["path"]!.GetValue<string>());
            Assert.Equal(dom["address"]!.GetValue<string>(), streamed["address"]!.GetValue<string>());
        }
    }

    [Fact]
    public void Twenty_random_cells_match_between_streaming_and_dom_on_a_generated_file()
    {
        // The same generator as the large-file fixture, sized for a fast DOM load.
        var file = NewFile("generated.xlsx");
        BigWorkbookGenerator.Generate(file, rows: 500);

        var random = new Random(42);
        for (var i = 0; i < 20; i++)
        {
            var row = random.Next(1, 501);
            var column = "ABCDE"[random.Next(5)];
            var path = $"/Sheet1/{column}{row}";

            var dom = OkData(Handler.Get(Ctx(file, ("path", path))));
            var streamed = OkData(Handler.Get(Ctx(file, ("path", path), ("stream", true))));

            Assert.Equal(dom["value"]!.ToJsonString(), streamed["value"]!.ToJsonString());
            Assert.Equal(dom["type"]!.GetValue<string>(), streamed["type"]!.GetValue<string>());
        }
    }

    [Fact]
    public void Streamed_range_get_matches_dom_values_and_truncation()
    {
        var file = CreateMixedWorkbook();

        var dom = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1:D4"))));
        var streamed = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1:D4"), ("stream", true))));

        Assert.Equal(dom["values"]!.ToJsonString(), streamed["values"]!.ToJsonString());
        Assert.Equal(dom["rows"]!.GetValue<int>(), streamed["rows"]!.GetValue<int>());
        Assert.Equal(dom["columns"]!.GetValue<int>(), streamed["columns"]!.GetValue<int>());
        Assert.Equal(dom["range"]!.GetValue<string>(), streamed["range"]!.GetValue<string>());

        // maxCells truncation must behave identically (row-wise, with warning).
        var domCut = Handler.Get(Ctx(file, ("path", "/Sheet1/A1:D4"), ("maxCells", 8)));
        var streamedCut = Handler.Get(Ctx(file, ("path", "/Sheet1/A1:D4"), ("maxCells", 8), ("stream", true)));
        Assert.Equal(
            Json(domCut)["data"]!["values"]!.ToJsonString(),
            Json(streamedCut)["data"]!["values"]!.ToJsonString());
        Assert.True(Json(streamedCut)["data"]!["truncated"]!.GetValue<bool>());
        Assert.Equal("result_truncated", Json(streamedCut)["meta"]!["warnings"]![0]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void Streamed_text_matches_dom_text_for_a_window()
    {
        var file = CreateMixedWorkbook();

        var dom = OkData(Handler.Read(Ctx(file, ("view", "text"), ("range", "A3:C4"))));
        var streamed = OkData(Handler.Read(Ctx(file, ("view", "text"), ("range", "A3:C4"), ("stream", true))));

        Assert.Equal(dom["content"]!.GetValue<string>(), streamed["content"]!.GetValue<string>());
        Assert.False(streamed["truncated"]!.GetValue<bool>());
    }

    [Fact]
    public void Streamed_text_respects_maxBytes_with_a_truncation_warning()
    {
        var file = CreateMixedWorkbook();

        var envelope = Handler.Read(Ctx(file, ("view", "text"), ("maxBytes", 24), ("stream", true)));

        var root = Json(envelope);
        Assert.True(root["data"]!["truncated"]!.GetValue<bool>());
        Assert.True(root["data"]!["content"]!.GetValue<string>().Length <= 24);
        Assert.Equal("result_truncated", root["meta"]!["warnings"]![0]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void Streamed_date_cells_surface_as_raw_serial_numbers_documented_divergence()
    {
        // No styles part is read on the streaming path, so a date-formatted
        // number cell honestly reports its stored serial value.
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "2024-05-01"))).IsOk);

        var dom = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        var streamed = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"), ("stream", true))));

        Assert.Equal("dateTime", dom["type"]!.GetValue<string>());
        Assert.Equal("number", streamed["type"]!.GetValue<string>());
        Assert.Equal(new DateTime(2024, 5, 1).ToOADate(), streamed["value"]!.GetValue<double>());
    }

    [Fact]
    public void Streamed_get_of_a_blank_cell_reports_blank()
    {
        var file = CreateWorkbook();

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/Z99"), ("stream", true))));

        Assert.Equal("blank", data["type"]!.GetValue<string>());
        Assert.Null(data["value"]);
    }

    [Fact]
    public void Streamed_get_with_a_wrong_sheet_is_invalid_path_with_candidates()
    {
        var file = CreateWorkbook();

        var envelope = Handler.Get(Ctx(file, ("path", "/Shet1/A1"), ("stream", true)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/Sheet1", envelope.Error.Candidates!);
    }

    [Fact]
    public void Streamed_structure_view_falls_back_to_dom_with_a_warning()
    {
        var file = CreateWorkbook();

        var envelope = Handler.Read(Ctx(file, ("view", "structure"), ("stream", true)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var warnings = Json(envelope)["meta"]!["warnings"]!;
        Assert.Equal("stream_fallback", warnings[0]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void Streamed_get_of_a_sheet_target_falls_back_to_dom()
    {
        var file = CreateMixedWorkbook();

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"), ("stream", true))));

        Assert.Equal("sheet", data["kind"]!.GetValue<string>());
        Assert.Null(data["streamed"]); // the DOM shape, not the streamed one
    }
}
