using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// Real Office cannot be automated here, so this test regenerates
/// <c>fixtures/manual-check/excel-sample.xlsx</c> for a human to open in Excel
/// (the OpenXml validator is only a proxy). Skips silently when the repo root
/// is not reachable from the test bin directory (e.g. detached CI layouts).
/// </summary>
public sealed class ManualCheckFixtureTests : ExcelTestBase
{
    [Fact]
    public void Regenerates_the_manual_check_fixture()
    {
        var root = FindRepoRoot();
        if (root is null)
        {
            return; // not running inside the repo; nothing to regenerate
        }

        var dir = Path.Combine(root, "fixtures", "manual-check");
        Directory.CreateDirectory(dir);
        var file = Path.Combine(dir, "excel-sample.xlsx");
        if (File.Exists(file))
        {
            File.Delete(file);
        }

        var workspace = new Workspace(dir);
        var ctx = new CommandContext
        {
            Workspace = workspace,
            File = file,
            Args = new JsonObject { ["title"] = "Q3 Data" },
        };
        Assert.True(Handler.Create(ctx).IsOk);

        var edit = Handler.Edit(
            new CommandContext { Workspace = workspace, File = file, Args = [] },
            [
                SetOp("/'Q3 Data'/A1", ("value", "Quarterly Report"), ("bold", true)),
                SetOp("/'Q3 Data'/A1:C1", ("merge", true)),
                SetOp("/'Q3 Data'/A3:C6", ("values", new JsonArray(
                    new JsonArray("Item", "Qty", "Price"),
                    new JsonArray("Widget", 12, 9.99),
                    new JsonArray("Gadget", 3, 24.5),
                    new JsonArray("Sprocket", 40, 1.25)))),
                AddOp("/'Q3 Data'/A3:C6", "table", ("name", "Sales")),
                SetOp("/'Q3 Data'/B8", ("value", "Total")),
                SetOp("/'Q3 Data'/C8", ("value", "=SUM(C4:C6)"), ("numberFormat", "#,##0.00"), ("bold", true)),
                SetOp("/'Q3 Data'/C9", ("value", "2026-06-12")),
                SetOp("/'Q3 Data'/A8", ("fill", "#FFF2CC")),
                AddOp("/'Q3 Data'", "chart",
                    ("kind", "bar"), ("dataRange", "A3:C6"), ("anchor", "E3"), ("title", "Sales by item")),
                AddOp("/'Q3 Data'", "pivot",
                    ("name", "SalesPivot"), ("sourceRange", "A3:C6"), ("targetSheet", "Pivot"),
                    ("rows", new JsonArray("Item")),
                    ("values", new JsonArray(
                        new JsonObject { ["field"] = "Qty", ["agg"] = "sum" },
                        new JsonObject { ["field"] = "Price", ["agg"] = "average" }))),
                AddOp("/'Q3 Data'/B4:B6", "conditionalFormat",
                    ("kind", "dataBar"), ("color", "638EC6")),
                AddOp("/'Q3 Data'/C4:C6", "conditionalFormat",
                    ("kind", "colorScale"), ("minColor", "FFFFFF"), ("maxColor", "63BE7B")),
            ]);
        Assert.True(edit.IsOk, edit.ToJson());

        // M3 surface, on its own sheet so the human can eyeball each feature:
        // scatter chart over numeric X/Y, a defined name used by a live SUM,
        // frozen header row, an AutoFilter and a print area.
        var m3 = Handler.Edit(
            new CommandContext { Workspace = workspace, File = file, Args = [] },
            [
                AddOp("/Metrics", "sheet"),
                SetOp("/Metrics/A1:B5", ("values", new JsonArray(
                    new JsonArray("X", "Y"),
                    new JsonArray(1, 2.5),
                    new JsonArray(2, 4.1),
                    new JsonArray(3, 9.8),
                    new JsonArray(4, 16.2)))),
                AddOp("/Metrics", "chart",
                    ("kind", "scatter"), ("dataRange", "A1:B5"), ("anchor", "D2"), ("title", "Growth (scatter)")),
                AddOp("/Metrics/B2:B5", "name", ("name", "GrowthYs")),
                SetOp("/Metrics/B7", ("value", "=SUM(GrowthYs)"), ("bold", true)),
                SetOp("/Metrics", ("freezeRows", 1)),
                SetOp("/Metrics/A1:B5", ("autoFilter", true)),
                SetOp("/Metrics", ("printArea", "A1:F20")),
            ]);
        Assert.True(m3.IsOk, m3.ToJson());

        // Image: a small PNG written into the sandbox, embedded, then cleaned up.
        var logo = Path.Combine(dir, "manual-check-logo.png");
        File.WriteAllBytes(logo, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAAEElEQVR4nGP4z8AARwzIHABvqgf5gNwAKAAAAABJRU5ErkJggg=="));
        try
        {
            var image = Handler.Edit(
                new CommandContext { Workspace = workspace, File = file, Args = [] },
                [AddOp("/'Q3 Data'", "image",
                    ("src", "manual-check-logo.png"), ("anchor", "E20"), ("widthPx", 80))]);
            Assert.True(image.IsOk, image.ToJson());
        }
        finally
        {
            File.Delete(logo);
        }

        // Proxy oracle until a human opens it in real Excel.
        AssertValidatorClean(file);
        var raw = RawCell(file, "Q3 Data", "C8");
        Assert.NotNull(raw.CachedValue);
        var cached = double.Parse(raw.CachedValue, System.Globalization.CultureInfo.InvariantCulture);
        Assert.Equal(35.74, cached, precision: 10); // 9.99 + 24.5 + 1.25

        // The defined-name SUM must carry a real cached value for Excel to show.
        var nameSum = RawCell(file, "Metrics", "B7");
        Assert.NotNull(nameSum.CachedValue);
        Assert.Equal(
            32.6, // 2.5 + 4.1 + 9.8 + 16.2
            double.Parse(nameSum.CachedValue, System.Globalization.CultureInfo.InvariantCulture),
            precision: 10);
    }

    private static string? FindRepoRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "AIOffice.sln")))
        {
            dir = dir.Parent;
        }

        return dir?.FullName;
    }
}
