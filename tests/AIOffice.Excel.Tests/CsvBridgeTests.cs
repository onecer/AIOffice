using System.Text;
using AIOffice.Core;
using ClosedXML.Excel;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// M5 csv bridge. Import: <see cref="ExcelHandler.CreateFrom"/> parses RFC 4180
/// (quotes, embedded commas/newlines, CRLF, BOM, delimiter sniffing) and types
/// values like bulk writes — with the documented leading-zero exception.
/// Export: <c>read --view csv</c> emits RFC-4180-safe text whose re-import
/// reproduces the same typed sheet; formulas export as cached values.
/// </summary>
public sealed class CsvBridgeTests : ExcelTestBase
{
    private string WriteCsv(string name, string content, bool bom = false)
    {
        var path = Path.Combine(Dir, name);
        File.WriteAllText(path, content, new UTF8Encoding(encoderShouldEmitUTF8Identifier: bom));
        return path;
    }

    private Envelope Import(string csvPath, string xlsxName = "out.xlsx", params (string Key, System.Text.Json.Nodes.JsonNode? Value)[] args) =>
        Handler.CreateFrom(Ctx(NewFile(xlsxName), args), csvPath);

    // ----- import -----------------------------------------------------------

    [Fact]
    public void Import_types_values_like_bulk_writes_and_keeps_leading_zeros()
    {
        var csv = WriteCsv("data.csv", "Name,Qty,Price,Active,When,Code\nant,5,1.5,true,2024-05-01,007\n");
        var envelope = Import(csv);

        var data = OkData(envelope);
        Assert.Equal(2, data["rows"]!.GetValue<int>());
        Assert.Equal(6, data["columns"]!.GetValue<int>());
        Assert.Equal(",", data["delimiter"]!.GetValue<string>());
        Assert.False(data["streamed"]!.GetValue<bool>());
        var file = data["file"]!.GetValue<string>();

        using (var workbook = new XLWorkbook(file))
        {
            var ws = workbook.Worksheet("Sheet1");
            Assert.Equal("Name", ws.Cell("A1").GetText());
            Assert.Equal(XLDataType.Number, ws.Cell("B2").Value.Type);
            Assert.Equal(5.0, ws.Cell("B2").GetDouble());
            Assert.Equal(1.5, ws.Cell("C2").GetDouble());
            Assert.Equal(XLDataType.Boolean, ws.Cell("D2").Value.Type);
            Assert.True(ws.Cell("D2").GetBoolean());
            Assert.Equal(XLDataType.DateTime, ws.Cell("E2").Value.Type);
            Assert.Equal(new DateTime(2024, 5, 1), ws.Cell("E2").GetDateTime());
            Assert.Equal(XLDataType.Text, ws.Cell("F2").Value.Type); // leading zero survives
            Assert.Equal("007", ws.Cell("F2").GetText());
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Import_handles_quoted_commas_escaped_quotes_and_embedded_newlines()
    {
        var csv = WriteCsv(
            "tricky.csv",
            "a,b\r\n\"hello, world\",\"line1\r\nline2\"\r\n\"she said \"\"hi\"\"\",plain\r\n");
        var file = OkData(Import(csv))["file"]!.GetValue<string>();

        using var workbook = new XLWorkbook(file);
        var ws = workbook.Worksheet(1);
        Assert.Equal("hello, world", ws.Cell("A2").GetText());
        Assert.Equal("line1\nline2", ws.Cell("B2").GetText()); // CRLF inside quotes normalized to \n
        Assert.Equal("she said \"hi\"", ws.Cell("A3").GetText());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Import_tolerates_a_utf8_bom()
    {
        var csv = WriteCsv("bom.csv", "x,y\n1,2\n", bom: true);
        var file = OkData(Import(csv))["file"]!.GetValue<string>();

        using var workbook = new XLWorkbook(file);
        Assert.Equal("x", workbook.Worksheet(1).Cell("A1").GetText()); // no BOM glued to the first field
    }

    [Fact]
    public void Import_sniffs_semicolon_and_tab_delimiters()
    {
        var semicolon = WriteCsv("semi.csv", "a;b;c\n1;2;3\n");
        var semiData = OkData(Import(semicolon, "semi.xlsx"));
        Assert.Equal(";", semiData["delimiter"]!.GetValue<string>());
        Assert.Equal(3, semiData["columns"]!.GetValue<int>());

        var tab = WriteCsv("tab.tsv", "a\tb\tc\n1\t2\t3\n");
        var tabData = OkData(Import(tab, "tab.xlsx"));
        Assert.Equal("\t", tabData["delimiter"]!.GetValue<string>());
        Assert.Equal(3, tabData["columns"]!.GetValue<int>());
    }

    [Fact]
    public void Import_honors_a_delimiter_override()
    {
        // Commas inside the values would win the sniff; the override pins ';'.
        var csv = WriteCsv("override.csv", "a,x;b,y\n1,2;3,4\n");
        var data = OkData(Import(csv, "o.xlsx", ("delimiter", ";")));
        Assert.Equal(";", data["delimiter"]!.GetValue<string>());
        Assert.Equal(2, data["columns"]!.GetValue<int>());

        using var workbook = new XLWorkbook(data["file"]!.GetValue<string>());
        Assert.Equal("a,x", workbook.Worksheet(1).Cell("A1").GetText());
    }

    [Fact]
    public void Import_keeps_blank_lines_as_empty_rows_but_drops_the_trailing_newline()
    {
        var csv = WriteCsv("rows.csv", "a\n\nb\n");
        var data = OkData(Import(csv));
        Assert.Equal(3, data["rows"]!.GetValue<int>()); // a, blank, b — no fourth row

        using var workbook = new XLWorkbook(data["file"]!.GetValue<string>());
        var ws = workbook.Worksheet(1);
        Assert.Equal("a", ws.Cell("A1").GetText());
        Assert.True(ws.Cell("A2").Value.IsBlank);
        Assert.Equal("b", ws.Cell("A3").GetText());
    }

    [Fact]
    public void Import_over_50k_cells_streams_and_lands_every_value()
    {
        var sb = new StringBuilder();
        for (var r = 1; r <= 10500; r++)
        {
            sb.Append(r).Append(',').Append(r * 2).Append(',').Append("row").Append(r).Append(',').Append(r % 2 == 0 ? "true" : "false").Append(",x\n");
        }

        var csv = WriteCsv("big.csv", sb.ToString());
        var data = OkData(Import(csv));
        Assert.True(data["streamed"]!.GetValue<bool>());
        Assert.Equal(10500, data["rows"]!.GetValue<int>());
        Assert.Equal(52500, data["cells"]!.GetValue<int>());

        var file = data["file"]!.GetValue<string>();
        using (var workbook = new XLWorkbook(file))
        {
            var ws = workbook.Worksheet("Sheet1");
            Assert.Equal(10500.0, ws.Cell("A10500").GetDouble());
            Assert.Equal(21000.0, ws.Cell("B10500").GetDouble());
            Assert.Equal("row10500", ws.Cell("C10500").GetText());
            Assert.True(ws.Cell("D10500").GetBoolean());
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Import_formula_fields_get_cached_values()
    {
        var csv = WriteCsv("formula.csv", "1\n2\n=SUM(A1:A2)\n");
        var file = OkData(Import(csv))["file"]!.GetValue<string>();

        var (formula, cached, _) = RawCell(file, "Sheet1", "A3");
        Assert.Equal("SUM(A1:A2)", formula);
        Assert.Equal("3", cached);
    }

    [Fact]
    public void Import_rejects_bad_inputs_with_typed_errors()
    {
        var csv = WriteCsv("ok.csv", "a\n");

        // Existing target file.
        var existing = CreateWorkbook("exists.xlsx");
        var clash = Handler.CreateFrom(Ctx(existing), csv);
        Assert.False(clash.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, clash.Error!.Code);

        // Missing source.
        var missing = Import(Path.Combine(Dir, "nope.csv"), "m.xlsx");
        Assert.False(missing.IsOk);
        Assert.Equal(ErrorCodes.FileNotFound, missing.Error!.Code);

        // Source outside the sandbox.
        var outsideDir = Path.Combine(Path.GetTempPath(), "aioffice-outside-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outsideDir);
        try
        {
            var outside = Path.Combine(outsideDir, "x.csv");
            File.WriteAllText(outside, "a\n");
            var denied = Import(outside, "d.xlsx");
            Assert.False(denied.IsOk);
            Assert.Equal(ErrorCodes.SandboxDenied, denied.Error!.Code);
        }
        finally
        {
            Directory.Delete(outsideDir, recursive: true);
        }

        // Non-csv extension names the workaround.
        var json = Path.Combine(Dir, "data.json");
        File.WriteAllText(json, "{}");
        var unsupported = Import(json, "j.xlsx");
        Assert.False(unsupported.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, unsupported.Error!.Code);
        Assert.NotEmpty(unsupported.Error.Suggestion);

        // Bad delimiter override.
        var badDelimiter = Import(csv, "b.xlsx", ("delimiter", "|"));
        Assert.False(badDelimiter.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, badDelimiter.Error!.Code);
        Assert.NotNull(badDelimiter.Error.Candidates);

        // Unbalanced quote.
        var torn = WriteCsv("torn.csv", "a,\"unclosed\n");
        var corrupt = Import(torn, "t.xlsx");
        Assert.False(corrupt.IsOk);
        Assert.Equal(ErrorCodes.FormatCorrupt, corrupt.Error!.Code);
    }

    [Fact]
    public void Import_of_an_empty_csv_creates_an_empty_workbook_with_a_warning()
    {
        var csv = WriteCsv("empty.csv", "");
        var envelope = Import(csv);

        var data = OkData(envelope);
        Assert.Equal(0, data["rows"]!.GetValue<int>());
        var warnings = Json(envelope)["meta"]!["warnings"]!.AsArray();
        Assert.Contains(warnings, w => w!["code"]!.GetValue<string>() == "csv_empty");
    }

    // ----- export -----------------------------------------------------------

    [Fact]
    public void Export_emits_rfc4180_with_cached_formula_values()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:C2", ("values", new System.Text.Json.Nodes.JsonArray(
                new System.Text.Json.Nodes.JsonArray("plain", "with, comma", "say \"hi\""),
                new System.Text.Json.Nodes.JsonArray(1.5, true, "2024-05-01")))),
            SetOp("/Sheet1/A3", ("value", "=SUM(A2:A2)"))).IsOk);

        var data = OkData(Handler.Read(Ctx(file, ("view", "csv"))));
        Assert.Equal("csv", data["view"]!.GetValue<string>());
        Assert.Equal("Sheet1", data["sheet"]!.GetValue<string>());

        var lines = data["content"]!.GetValue<string>().Split("\r\n");
        Assert.Equal("plain,\"with, comma\",\"say \"\"hi\"\"\"", lines[0]);
        Assert.Equal("1.5,true,2024-05-01", lines[1]);
        Assert.Equal("1.5,,", lines[2]); // formula exported as its cached value; blanks stay empty
    }

    [Fact]
    public void Export_round_trips_a_typed_sheet_through_import()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:D2", ("values", new System.Text.Json.Nodes.JsonArray(
                new System.Text.Json.Nodes.JsonArray("Name", "Qty", "Active", "When"),
                new System.Text.Json.Nodes.JsonArray("ant, the first", 5, true, "2024-05-01")))),
            SetOp("/Sheet1/A3", ("value", "007"))).IsOk);

        var csvText = OkData(Handler.Read(Ctx(file, ("view", "csv"))))["content"]!.GetValue<string>();
        var csvPath = Path.Combine(Dir, "roundtrip.csv");
        File.WriteAllText(csvPath, csvText);
        var reimported = OkData(Import(csvPath, "roundtrip.xlsx"))["file"]!.GetValue<string>();

        using var before = new XLWorkbook(file);
        using var after = new XLWorkbook(reimported);
        var b = before.Worksheet(1);
        var a = after.Worksheet(1);
        foreach (var address in new[] { "A1", "B1", "C1", "D1", "A2", "B2", "C2", "D2", "A3" })
        {
            Assert.Equal(b.Cell(address).Value.Type, a.Cell(address).Value.Type);
            Assert.Equal(b.Cell(address).Value, a.Cell(address).Value);
        }
    }

    [Fact]
    public void Export_documented_round_trip_exceptions_hold()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "=1/0")), // error value
            SetOp("/Sheet1/A2", ("value", "123"), ("valueType", "text"))).IsOk);

        var csvText = OkData(Handler.Read(Ctx(file, ("view", "csv"))))["content"]!.GetValue<string>();
        var csvPath = Path.Combine(Dir, "exceptions.csv");
        File.WriteAllText(csvPath, csvText);
        var reimported = OkData(Import(csvPath, "exceptions.xlsx"))["file"]!.GetValue<string>();

        using var after = new XLWorkbook(reimported);
        // Documented: the error's display text comes back as text…
        Assert.Equal("#DIV/0!", after.Worksheet(1).Cell("A1").GetText());
        // …and a forced-text number re-imports as a number (no leading zero to save it).
        Assert.Equal(XLDataType.Number, after.Worksheet(1).Cell("A2").Value.Type);
    }

    [Fact]
    public void Export_narrows_by_sheet_and_range()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            AddOp("/Data", "sheet"),
            SetOp("/Data/A1:B2", ("values", new System.Text.Json.Nodes.JsonArray(
                new System.Text.Json.Nodes.JsonArray(1, 2),
                new System.Text.Json.Nodes.JsonArray(3, 4))))).IsOk);

        var bySheet = OkData(Handler.Read(Ctx(file, ("view", "csv"), ("sheet", "Data"))));
        Assert.Equal("Data", bySheet["sheet"]!.GetValue<string>());
        Assert.Equal("1,2\r\n3,4\r\n", bySheet["content"]!.GetValue<string>());

        var byRange = OkData(Handler.Read(Ctx(file, ("view", "csv"), ("range", "/Data/A2:B2"))));
        Assert.Equal("3,4\r\n", byRange["content"]!.GetValue<string>());

        var narrowed = OkData(Handler.Read(Ctx(file, ("view", "csv"), ("sheet", "Data"), ("range", "B1:B2"))));
        Assert.Equal("2\r\n4\r\n", narrowed["content"]!.GetValue<string>());

        var conflict = Handler.Read(Ctx(file, ("view", "csv"), ("sheet", "Data"), ("range", "/Data/A1:B2")));
        Assert.False(conflict.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, conflict.Error!.Code);

        var unknownSheet = Handler.Read(Ctx(file, ("view", "csv"), ("sheet", "Dta")));
        Assert.False(unknownSheet.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, unknownSheet.Error!.Code);
        Assert.Contains("/Data", unknownSheet.Error.Candidates!);
    }

    [Fact]
    public void Export_truncates_at_max_bytes_with_a_warning()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:A50", ("values", new System.Text.Json.Nodes.JsonArray(
                [.. Enumerable.Range(1, 50).Select(i => (System.Text.Json.Nodes.JsonNode)new System.Text.Json.Nodes.JsonArray("some filler text " + i))])))).IsOk);

        var envelope = Handler.Read(Ctx(file, ("view", "csv"), ("maxBytes", 100)));
        var data = OkData(envelope);
        Assert.True(data["truncated"]!.GetValue<bool>());
        var warnings = Json(envelope)["meta"]!["warnings"]!.AsArray();
        Assert.Contains(warnings, w => w!["code"]!.GetValue<string>() == "result_truncated");
    }

    // ----- export delimiter -------------------------------------------------

    [Fact]
    public void Export_without_delimiter_is_byte_identical_comma_output()
    {
        // Acceptance 1: no delimiter prop -> the shipped comma output, unchanged envelope.
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:C2", ("values", new System.Text.Json.Nodes.JsonArray(
                new System.Text.Json.Nodes.JsonArray("plain", "with, comma", "say \"hi\""),
                new System.Text.Json.Nodes.JsonArray(1.5, true, "2024-05-01"))))).IsOk);

        var data = OkData(Handler.Read(Ctx(file, ("view", "csv"))));
        Assert.Equal("csv", data["view"]!.GetValue<string>());
        Assert.Equal("Sheet1", data["sheet"]!.GetValue<string>());
        Assert.Equal("A1:C2", data["range"]!.GetValue<string>());
        Assert.False(data["truncated"]!.GetValue<bool>());
        Assert.Equal(
            "plain,\"with, comma\",\"say \"\"hi\"\"\"\r\n1.5,true,2024-05-01\r\n",
            data["content"]!.GetValue<string>());
    }

    [Fact]
    public void Export_delimiter_emits_tab_or_semicolon_separated_output()
    {
        // Acceptance 3: tab (via both spellings) and semicolon separators.
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:C1", ("values", new System.Text.Json.Nodes.JsonArray(
                new System.Text.Json.Nodes.JsonArray("a", "b", "c"))))).IsOk);

        var tab = OkData(Handler.Read(Ctx(file, ("view", "csv"), ("delimiter", "tab"))))["content"]!.GetValue<string>();
        Assert.Equal("a\tb\tc\r\n", tab);

        var tabEscape = OkData(Handler.Read(Ctx(file, ("view", "csv"), ("delimiter", "\\t"))))["content"]!.GetValue<string>();
        Assert.Equal("a\tb\tc\r\n", tabEscape);

        var semiWord = OkData(Handler.Read(Ctx(file, ("view", "csv"), ("delimiter", "semicolon"))))["content"]!.GetValue<string>();
        Assert.Equal("a;b;c\r\n", semiWord);

        var semiSymbol = OkData(Handler.Read(Ctx(file, ("view", "csv"), ("delimiter", ";"))))["content"]!.GetValue<string>();
        Assert.Equal("a;b;c\r\n", semiSymbol);
    }

    [Fact]
    public void Export_delimiter_tab_round_trips_through_import()
    {
        // Acceptance 4: export tab-separated, re-import with delimiter='tab', identical grid.
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:D2", ("values", new System.Text.Json.Nodes.JsonArray(
                new System.Text.Json.Nodes.JsonArray("Name", "Qty", "Active", "When"),
                new System.Text.Json.Nodes.JsonArray("ant, the first", 5, true, "2024-05-01")))),
            SetOp("/Sheet1/A3", ("value", "007"))).IsOk);

        var tsvText = OkData(Handler.Read(Ctx(file, ("view", "csv"), ("delimiter", "tab"))))["content"]!.GetValue<string>();
        var tsvPath = Path.Combine(Dir, "roundtrip.tsv");
        File.WriteAllText(tsvPath, tsvText);
        var reimported = OkData(Import(tsvPath, "roundtrip-tsv.xlsx", ("delimiter", "tab")))["file"]!.GetValue<string>();

        using var before = new XLWorkbook(file);
        using var after = new XLWorkbook(reimported);
        var b = before.Worksheet(1);
        var a = after.Worksheet(1);
        foreach (var address in new[] { "A1", "B1", "C1", "D1", "A2", "B2", "C2", "D2", "A3" })
        {
            Assert.Equal(b.Cell(address).Value.Type, a.Cell(address).Value.Type);
            Assert.Equal(b.Cell(address).Value, a.Cell(address).Value);
        }

        // The comma inside "ant, the first" is NOT the tab delimiter, so it lands whole in one cell.
        Assert.Equal("ant, the first", a.Cell("A2").GetText());
    }

    [Fact]
    public void Export_delimiter_quoting_follows_the_active_delimiter()
    {
        // Acceptance 5: with delimiter='tab', a comma is NOT special (unquoted); a tab IS.
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:B1", ("values", new System.Text.Json.Nodes.JsonArray(
                new System.Text.Json.Nodes.JsonArray("has, comma", "has\ttab"))))).IsOk);

        var tab = OkData(Handler.Read(Ctx(file, ("view", "csv"), ("delimiter", "tab"))))["content"]!.GetValue<string>();
        // Field with a comma stays bare; field with the active tab delimiter is quoted.
        Assert.Equal("has, comma\t\"has\ttab\"\r\n", tab);
    }

    [Fact]
    public void Export_rejects_an_unsupported_delimiter_token()
    {
        // Acceptance 6: reuse ParseDelimiterArg's vocabulary — 'pipe'/'|' is invalid_args.
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "x"))).IsOk);

        var pipe = Handler.Read(Ctx(file, ("view", "csv"), ("delimiter", "|")));
        Assert.False(pipe.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, pipe.Error!.Code);
        Assert.NotNull(pipe.Error.Candidates);
        Assert.Equal(ExcelCsv.DelimiterNames, pipe.Error.Candidates);
    }
}
