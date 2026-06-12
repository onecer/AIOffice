using AIOffice.Core;
using ClosedXML.Excel;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// M5 cell hyperlinks: external URLs and internal <c>#Sheet!A1</c> targets via
/// <c>set</c>, cleared with an empty string, reflected by <c>get</c>, and
/// rendered as <c>&lt;a&gt;</c> in html.
/// </summary>
public sealed class HyperlinkTests : ExcelTestBase
{
    [Fact]
    public void External_link_with_tooltip_round_trips()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(file, SetOp(
            "/Sheet1/A1",
            ("value", "docs"),
            ("hyperlink", "https://example.com/docs"),
            ("hyperlinkTooltip", "open the docs")));
        Assert.True(envelope.IsOk, envelope.ToJson());
        var applied = Json(envelope)["data"]!["ops"]![0]!["applied"]!.AsArray().Select(a => a!.GetValue<string>());
        Assert.Contains("hyperlink", applied);

        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        Assert.Equal("https://example.com/docs", cell["hyperlink"]!.GetValue<string>());
        Assert.Equal("open the docs", cell["hyperlinkTooltip"]!.GetValue<string>());

        using (var workbook = new XLWorkbook(file))
        {
            var a1 = workbook.Worksheet("Sheet1").Cell("A1");
            Assert.True(a1.HasHyperlink);
            Assert.True(a1.GetHyperlink().IsExternal);
            Assert.Equal("https://example.com/docs", a1.GetHyperlink().ExternalAddress.ToString());
            Assert.Equal("open the docs", a1.GetHyperlink().Tooltip);
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Internal_link_targets_another_sheet()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            AddOp("/Detail", "sheet"),
            SetOp("/Sheet1/B1", ("value", "jump"), ("hyperlink", "#Detail!A1"))).IsOk);

        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B1"))));
        Assert.Equal("#Detail!A1", cell["hyperlink"]!.GetValue<string>());

        using (var workbook = new XLWorkbook(file))
        {
            var b1 = workbook.Worksheet("Sheet1").Cell("B1");
            Assert.True(b1.HasHyperlink);
            Assert.False(b1.GetHyperlink().IsExternal);
            Assert.Equal("Detail!A1", b1.GetHyperlink().InternalAddress);
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Empty_string_clears_the_link_and_clearing_again_is_a_noop()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp(
            "/Sheet1/A1", ("value", "keep"), ("hyperlink", "https://example.com"))).IsOk);

        var cleared = EditOps(file, SetOp("/Sheet1/A1", ("hyperlink", "")));
        Assert.True(cleared.IsOk, cleared.ToJson());

        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        Assert.Equal("keep", cell["value"]!.GetValue<string>()); // value untouched
        Assert.Null(cell["hyperlink"]);

        using (var workbook = new XLWorkbook(file))
        {
            Assert.False(workbook.Worksheet("Sheet1").Cell("A1").HasHyperlink);
        }

        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("hyperlink", ""))).IsOk); // idempotent
        AssertValidatorClean(file);
    }

    [Fact]
    public void Tooltip_alone_retargets_an_existing_link()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("hyperlink", "https://example.com"))).IsOk);
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("hyperlinkTooltip", "new tip"))).IsOk);

        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        Assert.Equal("https://example.com/", cell["hyperlink"]!.GetValue<string>());
        Assert.Equal("new tip", cell["hyperlinkTooltip"]!.GetValue<string>());

        var orphan = EditOps(file, SetOp("/Sheet1/Z9", ("hyperlinkTooltip", "no link here")));
        Assert.False(orphan.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, orphan.Error!.Code);
    }

    [Fact]
    public void Bad_links_and_range_targets_are_rejected()
    {
        var file = CreateWorkbook();

        var relative = EditOps(file, SetOp("/Sheet1/A1", ("hyperlink", "not a url")));
        Assert.False(relative.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, relative.Error!.Code);
        Assert.Contains("#Sheet2!A1", relative.Error.Suggestion, StringComparison.Ordinal);

        var bareHash = EditOps(file, SetOp("/Sheet1/A1", ("hyperlink", "#")));
        Assert.False(bareHash.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, bareHash.Error!.Code);

        var range = EditOps(file, SetOp("/Sheet1/A1:B2", ("hyperlink", "https://example.com")));
        Assert.False(range.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, range.Error!.Code);
        Assert.Contains("cell", range.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Render_html_emits_anchors_with_the_aio_path_contract_intact()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            AddOp("/Detail", "sheet"),
            SetOp("/Sheet1/A1", ("value", "external"), ("hyperlink", "https://example.com/x?a=1&b=2")),
            SetOp("/Sheet1/B1", ("value", "internal"), ("hyperlink", "#Detail!A1")),
            SetOp("/Sheet1/C1", ("value", "plain"))).IsOk);

        var html = OkData(Handler.Render(Ctx(file, ("to", "html"))))["content"]!.GetValue<string>();

        Assert.Contains(
            "<td data-aio-path=\"/Sheet1/A1\"><a href=\"https://example.com/x?a=1&amp;b=2\">external</a></td>",
            html,
            StringComparison.Ordinal);
        Assert.Contains(
            "<td data-aio-path=\"/Sheet1/B1\"><a href=\"#Detail!A1\">internal</a></td>",
            html,
            StringComparison.Ordinal);
        Assert.Contains("<td data-aio-path=\"/Sheet1/C1\">plain</td>", html, StringComparison.Ordinal);
    }
}
