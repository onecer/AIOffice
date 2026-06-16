using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The v1.8 <c>gradientFill</c> cell/range prop (additive). ClosedXML 0.105 has no
/// gradient-fill model, so the gradient is authored raw on the saved styles part
/// and the target cell rebound. Raw assertions reopen the file with the OpenXml SDK
/// to read the actual <c>&lt;gradientFill&gt;</c>; every mutating test ends
/// OpenXmlValidator-clean.
/// </summary>
public sealed class GradientFillTests : ExcelTestBase
{
    /// <summary>The gradientFill on a cell, read raw from the styles part (or null).</summary>
    private static S.GradientFill? RawGradient(string file, string sheetName, string address)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        var cell = worksheetPart.Worksheet!.Descendants<S.Cell>()
            .FirstOrDefault(c => c.CellReference?.Value == address);
        if (cell?.StyleIndex?.Value is not { } styleIndex)
        {
            return null;
        }

        var cellFormat = stylesheet.CellFormats!.Elements<S.CellFormat>().ElementAt((int)styleIndex);
        if (cellFormat.FillId?.Value is not { } fillId)
        {
            return null;
        }

        return stylesheet.Fills!.Elements<S.Fill>().ElementAt((int)fillId).GetFirstChild<S.GradientFill>();
    }

    private static JsonObject Linear(double angle, params (string Color, double Pos)[] stops) =>
        new()
        {
            ["type"] = "linear",
            ["angle"] = angle,
            ["stops"] = Stops(stops),
        };

    private static JsonArray Stops((string Color, double Pos)[] stops)
    {
        var array = new JsonArray();
        foreach (var (color, pos) in stops)
        {
            array.Add(new JsonObject { ["color"] = color, ["pos"] = pos });
        }

        return array;
    }

    // ----- linear ------------------------------------------------------------

    [Fact]
    public void Linear_gradient_is_authored_and_round_trips_via_get()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/A1",
            ("value", "KPI"),
            ("gradientFill", Linear(90, ("#2E5AAC", 0), ("#E8743B", 1)))));
        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        var gradient = RawGradient(file, "Sheet1", "A1");
        Assert.NotNull(gradient);
        Assert.Equal(S.GradientValues.Linear, gradient!.Type!.Value);
        Assert.Equal(90d, gradient.Degree!.Value);
        var stops = gradient.Elements<S.GradientStop>().ToList();
        Assert.Equal(2, stops.Count);
        Assert.Equal("FF2E5AAC", stops[0].GetFirstChild<S.Color>()!.Rgb!.Value);
        Assert.Equal(0d, stops[0].Position!.Value);
        Assert.Equal("FFE8743B", stops[1].GetFirstChild<S.Color>()!.Rgb!.Value);
        Assert.Equal(1d, stops[1].Position!.Value);

        // get reports the gradient back in the same {type, angle, stops} shape.
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        var reported = data["gradientFill"]!;
        Assert.Equal("linear", reported["type"]!.GetValue<string>());
        Assert.Equal(90d, reported["angle"]!.GetValue<double>());
        Assert.Equal("2E5AAC", reported["stops"]![0]!["color"]!.GetValue<string>());
        Assert.Equal("E8743B", reported["stops"]![1]!["color"]!.GetValue<string>());
    }

    [Fact]
    public void Linear_gradient_defaults_to_zero_degrees_when_no_angle()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("gradientFill", new JsonObject
        {
            ["stops"] = Stops([("2E5AAC", 0), ("FFFFFF", 1)]),
        })));
        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        var gradient = RawGradient(file, "Sheet1", "A1");
        Assert.NotNull(gradient);
        Assert.Equal(S.GradientValues.Linear, gradient!.Type?.Value ?? S.GradientValues.Linear);
        Assert.Equal(0d, gradient.Degree?.Value ?? 0d);
    }

    [Fact]
    public void Gradient_preserves_the_cells_existing_number_format()
    {
        var file = CreateWorkbook();

        // A number format first, then a gradient: the rebind must keep the format.
        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 0.5), ("numberFormat", "percent2")),
            SetOp("/Sheet1/A1", ("gradientFill", Linear(0, ("2E5AAC", 0), ("E8743B", 1)))));
        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        Assert.NotNull(RawGradient(file, "Sheet1", "A1"));

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        Assert.Equal("0.00%", data["numberFormat"]!.GetValue<string>());
        Assert.Equal("50.00%", data["text"]!.GetValue<string>());
    }

    // ----- radial / path -----------------------------------------------------

    [Fact]
    public void Radial_gradient_authors_a_centered_path_fill_over_a_range()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/B2:C3", ("gradientFill", new JsonObject
        {
            ["type"] = "radial",
            ["stops"] = Stops([("FFFFFF", 0), ("2E5AAC", 1)]),
        })));
        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        // Every cell in the empty range got the gradient (materialized + rebound).
        foreach (var address in new[] { "B2", "C2", "B3", "C3" })
        {
            var gradient = RawGradient(file, "Sheet1", address);
            Assert.NotNull(gradient);
            Assert.Equal(S.GradientValues.Path, gradient!.Type!.Value);
            Assert.Null(gradient.Degree); // path gradients carry no degree
        }

        // get reports radial back with no angle.
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))));
        Assert.Equal("radial", data["gradientFill"]!["type"]!.GetValue<string>());
        Assert.Null(data["gradientFill"]!["angle"]);
    }

    // ----- validation --------------------------------------------------------

    [Fact]
    public void Fewer_than_two_stops_is_invalid_args()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("gradientFill", new JsonObject
        {
            ["stops"] = Stops([("2E5AAC", 0)]),
        })));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, Json(envelope)["error"]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void A_pos_outside_zero_to_one_is_invalid_args()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/A1",
            ("gradientFill", Linear(0, ("2E5AAC", 0), ("E8743B", 1.5)))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, Json(envelope)["error"]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void An_angle_on_a_radial_gradient_is_invalid_args()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("gradientFill", new JsonObject
        {
            ["type"] = "radial",
            ["angle"] = 45,
            ["stops"] = Stops([("2E5AAC", 0), ("E8743B", 1)]),
        })));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, Json(envelope)["error"]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void A_non_hex_stop_color_is_invalid_args()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/A1",
            ("gradientFill", Linear(0, ("notacolor", 0), ("E8743B", 1)))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, Json(envelope)["error"]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void A_cell_without_a_gradient_reports_null()
    {
        var file = CreateWorkbook();
        EditOps(file, SetOp("/Sheet1/A1", ("value", "plain")));

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        Assert.Null(data["gradientFill"]);
    }

    [Fact]
    public void Identical_gradients_share_one_fill_entry()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1", ("gradientFill", Linear(90, ("2E5AAC", 0), ("E8743B", 1)))),
            SetOp("/Sheet1/A2", ("gradientFill", Linear(90, ("2E5AAC", 0), ("E8743B", 1)))));
        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var gradientFills = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!
            .Fills!.Elements<S.Fill>().Count(f => f.GetFirstChild<S.GradientFill>() is not null);
        Assert.Equal(1, gradientFills);
    }
}
