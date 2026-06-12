using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The M3 headline: big workbooks WORK. A generated ~39 MB / 310k-row file
/// (see <see cref="BigWorkbookGenerator"/>) is read through the SAX streaming
/// path — stats, single cells, windows — and cross-checked against both the
/// deterministic generator and one full ClosedXML DOM load of the same file.
/// Every test pins <c>AIOFFICE_MAX_FILE_MB</c> explicitly so the suite is
/// order-independent of the integrator flipping the guard default to opt-in.
/// </summary>
[Collection("FileSizeEnv")]
public sealed class LargeWorkbookStreamingTests : ExcelTestBase, IClassFixture<LargeWorkbookFixture>
{
    private readonly LargeWorkbookFixture _fixture;

    public LargeWorkbookStreamingTests(LargeWorkbookFixture fixture) => _fixture = fixture;

    private static EnvScope GuardOff() => new(FileSizeGuard.EnvVar, "100000");

    [Fact]
    public void Generated_file_is_over_the_streaming_threshold_and_under_the_old_guard_default()
    {
        // Above 20 MB (ExcelStreaming.ThresholdBytes, internal): the streaming
        // path activates without stream=true.
        Assert.True(
            _fixture.SizeBytes > 20L * 1024 * 1024,
            $"fixture is {_fixture.SizeBytes / (1024.0 * 1024.0):F1} MB; the streaming threshold needs > 20 MB");

        // Below 50 MB: green whether or not the integrator flipped the guard default yet.
        Assert.True(
            _fixture.SizeBytes < 49L * 1024 * 1024,
            $"fixture is {_fixture.SizeBytes / (1024.0 * 1024.0):F1} MB; keep it under the old 50 MB default");
    }

    [Fact]
    public void Read_stats_streams_and_reports_correct_rows_and_sheets()
    {
        using var guard = GuardOff();

        // No stream arg: size alone must activate the streaming path.
        var data = OkData(Handler.Read(Ctx(_fixture.File, ("view", "stats"))));

        Assert.True(data["streamed"]!.GetValue<bool>());
        Assert.Equal(1, data["totals"]!["sheets"]!.GetValue<int>());
        var sheet = data["sheets"]![0]!;
        Assert.Equal("Sheet1", sheet["name"]!.GetValue<string>());
        Assert.Equal($"A1:F{LargeWorkbookFixture.Rows}", sheet["usedRange"]!.GetValue<string>());
        Assert.Equal(
            (long)LargeWorkbookFixture.Rows * BigWorkbookGenerator.Columns,
            sheet["cellCount"]!.GetValue<long>());
        Assert.Equal(LargeWorkbookFixture.Rows, sheet["formulaCount"]!.GetValue<long>());
    }

    [Fact]
    public void Read_succeeds_without_any_guard_override()
    {
        // The fixture is ~39 MB — under the old 50 MB default and under any
        // "unlimited by default" successor, so with the env var explicitly
        // CLEARED the read must succeed either way (no file_too_large).
        using var guard = new EnvScope(FileSizeGuard.EnvVar, null);

        var envelope = Handler.Read(Ctx(_fixture.File, ("view", "stats")));

        Assert.True(envelope.IsOk, envelope.ToJson());
    }

    [Fact]
    public void Get_single_cell_returns_the_right_value()
    {
        using var guard = GuardOff();

        var data = OkData(Handler.Get(Ctx(_fixture.File, ("path", "/Sheet1/C12345"))));

        Assert.True(data["streamed"]!.GetValue<bool>());
        Assert.Equal("cell", data["kind"]!.GetValue<string>());
        Assert.Equal("C12345", data["address"]!.GetValue<string>());
        Assert.Equal(BigWorkbookGenerator.ExpectedC(12345), data["value"]!.GetValue<double>());
        Assert.Equal("number", data["type"]!.GetValue<string>());
    }

    [Fact]
    public void Get_formula_cell_reports_formula_and_cached_value()
    {
        using var guard = GuardOff();

        var data = OkData(Handler.Get(Ctx(_fixture.File, ("path", "/Sheet1/E200000"))));

        Assert.Equal("=A200000*2", data["formula"]!.GetValue<string>());
        Assert.Equal(BigWorkbookGenerator.ExpectedCachedE(200000), data["cachedValue"]!.GetValue<double>());
        Assert.Equal(BigWorkbookGenerator.ExpectedCachedE(200000), data["value"]!.GetValue<double>());
    }

    [Fact]
    public void Get_shared_string_cell_resolves_through_the_streamed_sst()
    {
        using var guard = GuardOff();

        var data = OkData(Handler.Get(Ctx(_fixture.File, ("path", "/Sheet1/D309876"))));

        Assert.Equal(BigWorkbookGenerator.ExpectedSharedString(309876), data["value"]!.GetValue<string>());
        Assert.Equal("text", data["type"]!.GetValue<string>());
    }

    [Fact]
    public void Get_window_read_returns_the_right_values()
    {
        using var guard = GuardOff();

        var data = OkData(Handler.Get(Ctx(_fixture.File, ("path", "/Sheet1/A100000:E100009"))));

        Assert.True(data["streamed"]!.GetValue<bool>());
        Assert.Equal(10, data["rows"]!.GetValue<int>());
        Assert.Equal(5, data["columns"]!.GetValue<int>());
        Assert.False(data["truncated"]!.GetValue<bool>());
        for (var i = 0; i < 10; i++)
        {
            var row = 100000 + i;
            var values = data["values"]![i]!;
            Assert.Equal(BigWorkbookGenerator.ExpectedA(row), values[0]!.GetValue<double>());
            Assert.Equal(BigWorkbookGenerator.ExpectedB(row), values[1]!.GetValue<double>());
            Assert.Equal(BigWorkbookGenerator.ExpectedC(row), values[2]!.GetValue<double>());
            Assert.Equal(BigWorkbookGenerator.ExpectedSharedString(row), values[3]!.GetValue<string>());
            Assert.Equal(BigWorkbookGenerator.ExpectedCachedE(row), values[4]!.GetValue<double>());
        }
    }

    [Fact]
    public void Twenty_random_cells_match_the_generator_and_one_dom_window_agrees()
    {
        using var guard = GuardOff();

        // 20 random single-cell streamed gets against the generator's truth.
        var random = new Random(20260612);
        for (var i = 0; i < 20; i++)
        {
            var row = random.Next(1, LargeWorkbookFixture.Rows + 1);
            var column = "ABC"[random.Next(3)];
            var data = OkData(Handler.Get(Ctx(_fixture.File, ("path", $"/Sheet1/{column}{row}"))));
            var expected = column switch
            {
                'A' => BigWorkbookGenerator.ExpectedA(row),
                'B' => BigWorkbookGenerator.ExpectedB(row),
                _ => BigWorkbookGenerator.ExpectedC(row),
            };
            Assert.True(data["streamed"]!.GetValue<bool>());
            Assert.Equal(expected, data["value"]!.GetValue<double>());
        }

        // …and one DOM cross-check: load the same file fully through ClosedXML
        // (the engine the rest of aioffice trusts) and compare a window.
        var streamed = OkData(Handler.Get(Ctx(_fixture.File, ("path", "/Sheet1/A12340:E12345"))));
        using var workbook = new ClosedXML.Excel.XLWorkbook(_fixture.File);
        var sheet = workbook.Worksheet("Sheet1");
        for (var r = 12340; r <= 12345; r++)
        {
            var values = streamed["values"]![r - 12340]!;
            Assert.Equal(sheet.Cell(r, 1).GetDouble(), values[0]!.GetValue<double>());
            Assert.Equal(sheet.Cell(r, 2).GetDouble(), values[1]!.GetValue<double>());
            Assert.Equal(sheet.Cell(r, 3).GetDouble(), values[2]!.GetValue<double>());
            Assert.Equal(sheet.Cell(r, 4).GetString(), values[3]!.GetValue<string>());
            Assert.Equal(sheet.Cell(r, 5).CachedValue.GetNumber(), values[4]!.GetValue<double>());
        }
    }

    [Fact]
    public void Window_text_read_works_and_whole_sheet_text_truncates()
    {
        using var guard = GuardOff();

        var window = OkData(Handler.Read(Ctx(
            _fixture.File, ("view", "text"), ("range", "/Sheet1/A42:C43"))));
        var content = window["content"]!.GetValue<string>();
        Assert.Contains("# Sheet1!A42:C43", content, StringComparison.Ordinal);
        var ci = System.Globalization.CultureInfo.InvariantCulture;
        Assert.Contains(
            $"42,{BigWorkbookGenerator.ExpectedB(42).ToString("R", ci)},{BigWorkbookGenerator.ExpectedC(42).ToString("R", ci)}",
            content,
            StringComparison.Ordinal);
        Assert.False(window["truncated"]!.GetValue<bool>());

        // Whole-sheet text on 310k rows must stop at the byte budget, not scan on.
        var whole = OkData(Handler.Read(Ctx(_fixture.File, ("view", "text"))));
        Assert.True(whole["truncated"]!.GetValue<bool>());
        Assert.True(whole["content"]!.GetValue<string>().Length <= 65536);
    }

}
