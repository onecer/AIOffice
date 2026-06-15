using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>M5 native tables: a:tbl in a graphicFrame, /table[k]/tr[r]/tc[c] addressing.</summary>
public sealed class TableTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    /// <summary>Creates deck.pptx with one table on slide 1 and returns its canonical path.</summary>
    private string CreateWithTable(int rows = 3, int cols = 4, bool headerRow = true, string? style = null)
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        var props = TestEnv.Props(
            ("rows", rows), ("cols", cols), ("headerRow", headerRow),
            ("x", "2cm"), ("y", "5cm"), ("w", "28cm"));
        if (style is not null)
        {
            props["style"] = style;
        }

        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "table", props: props));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private A.Table OpenTable(PresentationDocument doc) =>
        doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!
            .Elements<P.GraphicFrame>().Single().Graphic!.GraphicData!.GetFirstChild<A.Table>()!;

    // ----- create -------------------------------------------------------------------

    [Fact]
    public void AddTable_BuildsGraphicFrame_WithGridAndRows()
    {
        var path = CreateWithTable();
        Assert.Equal("/slide[1]/table[1]", path);

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var frame = doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!
                .Elements<P.GraphicFrame>().Single();
            var table = frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;

            Assert.Equal(4, table.TableGrid!.Elements<A.GridColumn>().Count());
            Assert.Equal(3, table.Elements<A.TableRow>().Count());
            Assert.True(table.TableProperties!.FirstRow!.Value); // headerRow

            // The frame box matches the grid: x/y as given, width = sum of columns.
            Assert.Equal(720_000L, frame.Transform!.Offset!.X!.Value); // 2cm
            Assert.Equal(1_800_000L, frame.Transform!.Offset!.Y!.Value); // 5cm
            var gridWidth = table.TableGrid!.Elements<A.GridColumn>().Sum(c => c.Width!.Value);
            Assert.Equal(gridWidth, frame.Transform!.Extents!.Cx!.Value);
            Assert.Equal(10_080_000L, gridWidth); // 28cm

            // Header cells are painted bold; every cell has four borders + a solid fill.
            var headerCell = table.Elements<A.TableRow>().First().Elements<A.TableCell>().First();
            Assert.True(headerCell.TextBody!.Descendants<A.RunProperties>().First().Bold!.Value);
            Assert.NotNull(headerCell.TableCellProperties!.GetFirstChild<A.LeftBorderLineProperties>());
            Assert.NotNull(headerCell.TableCellProperties!.GetFirstChild<A.BottomBorderLineProperties>());
            Assert.NotNull(headerCell.TableCellProperties!.GetFirstChild<A.SolidFill>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void TableStyles_PaintDirectFills_NoThemeTableStyle()
    {
        CreateWithTable(style: "medium");

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var table = OpenTable(doc);

        // Direct paint, not a theme style reference.
        Assert.Empty(table.TableProperties!.Elements<A.TableStyleId>());

        var headerFill = table.Elements<A.TableRow>().First().Elements<A.TableCell>().First()
            .TableCellProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value;
        Assert.Equal("4472C4", headerFill);

        // The first body row is banded, the second is plain.
        var rows = table.Elements<A.TableRow>().ToList();
        Assert.Equal("D9E2F3", rows[1].Elements<A.TableCell>().First()
            .TableCellProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        Assert.Equal("FFFFFF", rows[2].Elements<A.TableCell>().First()
            .TableCellProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
    }

    [Fact]
    public void AddTable_UnknownStyle_IsTypedUnsupported()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            "add", "/slide[1]", type: "table",
            props: TestEnv.Props(("rows", 2), ("cols", 2), ("style", "neon")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        // The direct-paint looks plus the v1.4.0 built-in presets.
        Assert.Equal(
            ["light", "medium", "dark", "none", "light1", "light2", "medium1", "medium2", "medium3", "dark1", "dark2"],
            error.Candidates!);
    }

    [Fact]
    public void AddTable_WithoutRows_IsInvalidArgs()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", "/slide[1]", type: "table",
                props: TestEnv.Props(("cols", 3)))]),
            ErrorCodes.InvalidArgs);
    }

    // ----- get -------------------------------------------------------------------

    [Fact]
    public void Get_TablePath_ReportsDimsWidthsAndCells()
    {
        CreateWithTable();
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]"))));

        Assert.Equal("/slide[1]/table[1]", detail["path"]!.GetValue<string>());
        Assert.Equal(3, detail["rows"]!.GetValue<int>());
        Assert.Equal(4, detail["cols"]!.GetValue<int>());
        Assert.True(detail["headerRow"]!.GetValue<bool>());
        Assert.Equal(4, detail["columnWidths"]!.AsArray().Count);
        Assert.Equal(7, detail["columnWidths"]![0]!.GetValue<double>()); // 28cm / 4
        Assert.StartsWith("/slide[1]/shape[@id=", detail["shapePath"]!.GetValue<string>(), StringComparison.Ordinal);

        var rows = detail["rowsDetail"]!.AsArray();
        Assert.Equal(3, rows.Count);
        Assert.Equal("/slide[1]/table[1]/tr[2]/tc[3]", rows[1]!["cells"]![2]!["path"]!.GetValue<string>());
    }

    [Fact]
    public void Get_TableOutOfRange_IsInvalidPath()
    {
        CreateWithTable();
        TestEnv.AssertFail(
            _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[2]"))),
            ErrorCodes.InvalidPath);
        TestEnv.AssertFail(
            _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]/tr[9]"))),
            ErrorCodes.InvalidPath);
        TestEnv.AssertFail(
            _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]/tr[1]/tc[9]"))),
            ErrorCodes.InvalidPath);
    }

    [Fact]
    public void Get_ShapePathOfTableFrame_LinksTablePath()
    {
        CreateWithTable();
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/shape[1]"))));
        Assert.Equal("/slide[1]/table[1]", detail["tablePath"]!.GetValue<string>());

        var slide = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]"))));
        Assert.Equal("/slide[1]/table[1]", slide["tables"]![0]!.GetValue<string>());
    }

    // ----- cell sets -------------------------------------------------------------

    [Fact]
    public void SetCell_TextStylingAndFill_ReopenVerifies()
    {
        CreateWithTable();
        Edit(
            TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("text", "Region"))),
            TestEnv.Op("set", "/slide[1]/table[1]/tr[2]/tc[2]", props: TestEnv.Props(
                ("text", "42"), ("bold", true), ("color", "CC0000"), ("align", "right"), ("fill", "FFF7E6"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var table = OpenTable(doc);
            var rows = table.Elements<A.TableRow>().ToList();
            var header = rows[0].Elements<A.TableCell>().First();
            Assert.Equal("Region", header.TextBody!.Descendants<A.Text>().Single().Text);
            Assert.True(header.TextBody!.Descendants<A.RunProperties>().First().Bold!.Value); // header style kept

            var cell = rows[1].Elements<A.TableCell>().ToList()[1];
            Assert.Equal("42", cell.TextBody!.Descendants<A.Text>().Single().Text);
            var runProps = cell.TextBody!.Descendants<A.RunProperties>().First();
            Assert.True(runProps.Bold!.Value);
            Assert.Equal("CC0000", runProps.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Equal(A.TextAlignmentTypeValues.Right,
                cell.TextBody!.Elements<A.Paragraph>().Single().ParagraphProperties!.Alignment!.Value);
            Assert.Equal("FFF7E6", cell.TableCellProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]/tr[2]/tc[2]"))));
        Assert.Equal("42", detail["text"]!.GetValue<string>());
        Assert.Equal("FFF7E6", detail["fill"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetCell_UnknownProp_IsInvalidArgs()
    {
        CreateWithTable();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            "set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("w", "3cm")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("mergeRight", error.Candidates!);
    }

    // ----- merges ----------------------------------------------------------------

    [Fact]
    public void MergeRight_WritesGridSpanAndHMerge()
    {
        CreateWithTable();
        Edit(
            TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("text", "left"))),
            TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[2]", props: TestEnv.Props(("text", "right"))),
            TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("mergeRight", 1))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var cells = OpenTable(doc).Elements<A.TableRow>().First().Elements<A.TableCell>().ToList();
            Assert.Equal(2, cells[0].GridSpan!.Value);
            Assert.Null(cells[0].RowSpan);
            Assert.True(cells[1].HorizontalMerge!.Value);
            Assert.Null(cells[1].VerticalMerge);

            // The covered cell's text moved into the origin.
            Assert.Equal("left\nright", string.Join('\n', cells[0].TextBody!
                .Elements<A.Paragraph>().Select(p => p.InnerText).Where(t => t.Length > 0)));
            Assert.Equal(string.Empty, cells[1].TextBody!.InnerText);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void MergeDown_WritesRowSpanAndVMerge()
    {
        CreateWithTable();
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[2]/tc[1]", props: TestEnv.Props(("mergeDown", 1))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var rows = OpenTable(doc).Elements<A.TableRow>().ToList();
            Assert.Equal(2, rows[1].Elements<A.TableCell>().First().RowSpan!.Value);
            Assert.True(rows[2].Elements<A.TableCell>().First().VerticalMerge!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]/tr[3]/tc[1]"))));
        Assert.True(detail["covered"]!.GetValue<bool>());
        Assert.Equal("/slide[1]/table[1]/tr[2]/tc[1]", detail["mergedInto"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void MergeBlock_WritesTheFullMatrix()
    {
        CreateWithTable();
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[2]/tc[2]", props: TestEnv.Props(
            ("mergeRight", 1), ("mergeDown", 1))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var rows = OpenTable(doc).Elements<A.TableRow>().ToList();
            var origin = rows[1].Elements<A.TableCell>().ToList()[1];
            Assert.Equal(2, origin.GridSpan!.Value);
            Assert.Equal(2, origin.RowSpan!.Value);

            var rightOfOrigin = rows[1].Elements<A.TableCell>().ToList()[2];
            Assert.True(rightOfOrigin.HorizontalMerge!.Value);
            Assert.Equal(2, rightOfOrigin.RowSpan!.Value);

            var belowOrigin = rows[2].Elements<A.TableCell>().ToList()[1];
            Assert.True(belowOrigin.VerticalMerge!.Value);
            Assert.Equal(2, belowOrigin.GridSpan!.Value);

            var diagonal = rows[2].Elements<A.TableCell>().ToList()[2];
            Assert.True(diagonal.HorizontalMerge!.Value);
            Assert.True(diagonal.VerticalMerge!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]"))));
        var origin2 = detail["rowsDetail"]![1]!["cells"]![1]!;
        Assert.Equal(2, origin2["gridSpan"]!.GetValue<int>());
        Assert.Equal(2, origin2["rowSpan"]!.GetValue<int>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Merge_Overlapping_IsInvalidArgs()
    {
        CreateWithTable();
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("mergeRight", 1))));

        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
                "set", "/slide[1]/table[1]/tr[1]/tc[2]", props: TestEnv.Props(("mergeRight", 1)))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Merge_OutOfBounds_IsInvalidArgs()
    {
        CreateWithTable();
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
                "set", "/slide[1]/table[1]/tr[1]/tc[4]", props: TestEnv.Props(("mergeRight", 1)))]),
            ErrorCodes.InvalidArgs);
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
                "set", "/slide[1]/table[1]/tr[3]/tc[1]", props: TestEnv.Props(("mergeDown", 1)))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void SetCoveredCell_IsInvalidArgs()
    {
        CreateWithTable();
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("mergeRight", 1))));

        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
                "set", "/slide[1]/table[1]/tr[1]/tc[2]", props: TestEnv.Props(("text", "hidden")))]),
            ErrorCodes.InvalidArgs);
    }

    // ----- column widths ------------------------------------------------------------

    [Fact]
    public void SetColumnWidths_RewritesGridAndFrameWidth()
    {
        CreateWithTable(cols: 3);
        Edit(TestEnv.Op("set", "/slide[1]/table[1]", props: TestEnv.Props(
            ("columnWidths", new JsonArray("10cm", "6cm", "4cm")))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var frame = doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!
                .Elements<P.GraphicFrame>().Single();
            var widths = OpenTable(doc).TableGrid!.Elements<A.GridColumn>().Select(c => c.Width!.Value).ToList();
            Assert.Equal([3_600_000L, 2_160_000L, 1_440_000L], widths); // 10cm/6cm/4cm
            Assert.Equal(7_200_000L, frame.Transform!.Extents!.Cx!.Value); // 20cm
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetColumnWidths_WrongCount_IsInvalidArgs()
    {
        CreateWithTable(cols: 3);
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("set", "/slide[1]/table[1]", props: TestEnv.Props(
                ("columnWidths", new JsonArray("10cm", "6cm"))))]),
            ErrorCodes.InvalidArgs);
    }

    // ----- rows ----------------------------------------------------------------------

    [Fact]
    public void AddRow_AppendsAndInserts_CloningLookWithoutText()
    {
        CreateWithTable(rows: 2);
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[2]/tc[1]", props: TestEnv.Props(("text", "keep"))));

        var appended = Edit(TestEnv.Op("add", "/slide[1]/table[1]", type: "row"));
        Assert.Equal("/slide[1]/table[1]/tr[3]", appended["results"]![0]!["target"]!.GetValue<string>());

        var inserted = Edit(TestEnv.Op("add", "/slide[1]/table[1]/tr[2]", type: "row"));
        Assert.Equal("/slide[1]/table[1]/tr[2]", inserted["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var rows = OpenTable(doc).Elements<A.TableRow>().ToList();
            Assert.Equal(4, rows.Count);
            Assert.Equal(string.Empty, rows[1].Elements<A.TableCell>().First().TextBody!.InnerText); // the new row
            Assert.Equal("keep", rows[2].Elements<A.TableCell>().First().TextBody!.InnerText); // shifted down
            Assert.Equal(4, rows[1].Elements<A.TableCell>().Count()); // full grid width
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RemoveRow_DropsIt_AndFrameHeightShrinks()
    {
        CreateWithTable(rows: 3);
        long heightBefore;
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            heightBefore = doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!
                .Elements<P.GraphicFrame>().Single().Transform!.Extents!.Cy!.Value;
        }

        var data = Edit(TestEnv.Op("remove", "/slide[1]/table[1]/tr[2]"));
        Assert.Equal("/slide[1]/table[1]/tr[2]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Equal(2, OpenTable(doc).Elements<A.TableRow>().Count());
            var heightAfter = doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!
                .Elements<P.GraphicFrame>().Single().Transform!.Extents!.Cy!.Value;
            Assert.True(heightAfter < heightBefore);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RemoveRow_InsideVerticalMerge_IsInvalidArgs()
    {
        CreateWithTable();
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[2]/tc[1]", props: TestEnv.Props(("mergeDown", 1))));

        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("remove", "/slide[1]/table[1]/tr[3]")]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void RemoveOnlyRow_IsInvalidArgs_RemoveCellToo()
    {
        CreateWithTable(rows: 1);
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("remove", "/slide[1]/table[1]/tr[1]")]),
            ErrorCodes.InvalidArgs);
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("remove", "/slide[1]/table[1]/tr[1]/tc[1]")]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void RemoveTable_DropsTheFrame()
    {
        CreateWithTable();
        Edit(TestEnv.Op("remove", "/slide[1]/table[1]"));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Empty(doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!
                .Elements<P.GraphicFrame>());
        }

        TestEnv.AssertFail(
            _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]"))),
            ErrorCodes.InvalidPath);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ----- render ---------------------------------------------------------------------

    [Fact]
    public void Svg_DrawsTheGridWithPerCellPathsTextAndMerges()
    {
        CreateWithTable(rows: 2, cols: 2);
        Edit(
            TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("text", "Quarter"))),
            TestEnv.Op("set", "/slide[1]/table[1]/tr[2]/tc[1]", props: TestEnv.Props(("mergeRight", 1))));

        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        var svg = data["slides"]![0]!["svg"]!.GetValue<string>();

        // 3 visible cells (2 header + 1 merged) — the covered cell is not drawn.
        Assert.Contains("data-aio-path=\"/slide[1]/table[1]/tr[1]/tc[1]\"", svg, StringComparison.Ordinal);
        Assert.Contains("data-aio-path=\"/slide[1]/table[1]/tr[1]/tc[2]\"", svg, StringComparison.Ordinal);
        Assert.Contains("data-aio-path=\"/slide[1]/table[1]/tr[2]/tc[1]\"", svg, StringComparison.Ordinal);
        Assert.DoesNotContain("data-aio-path=\"/slide[1]/table[1]/tr[2]/tc[2]\"", svg, StringComparison.Ordinal);
        Assert.Equal(3, svg.Split("data-aio-path=\"/slide[1]/table[1]/").Length - 1);
        Assert.Contains(">Quarter</text>", svg, StringComparison.Ordinal);

        // The merged cell spans both columns: its rect is twice the single-cell width.
        // 28cm over 2 cols = 14cm = 5040000 EMU = 529.1px... merged = 1058.3px wide.
        Assert.Contains("width=\"1058.3\"", svg, StringComparison.Ordinal);

        // The frame group still carries the shape-level path around the cells.
        Assert.Contains("<g data-aio-path=\"/slide[1]/shape[@id=", svg, StringComparison.Ordinal);
    }

    // ----- query ---------------------------------------------------------------------

    [Fact]
    public void Query_TcContains_FindsTheCell()
    {
        CreateWithTable();
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[2]/tc[3]", props: TestEnv.Props(("text", "Quarterly total"))));

        var data = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "tc:contains('total')"))));
        Assert.Equal(1, data["count"]!.GetValue<int>());
        var match = data["matches"]![0]!;
        Assert.Equal("/slide[1]/table[1]/tr[2]/tc[3]", match["path"]!.GetValue<string>());
        Assert.Equal("tc", match["kind"]!.GetValue<string>());
        Assert.Equal("Quarterly total", match["text"]!.GetValue<string>());
    }

    [Fact]
    public void Query_TableAndSlideContains_SeeTableText()
    {
        CreateWithTable();
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("text", "Zebra"))));

        var tables = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "table:contains('Zebra')"))));
        Assert.Equal(1, tables["count"]!.GetValue<int>());
        Assert.Equal("/slide[1]/table[1]", tables["matches"]![0]!["path"]!.GetValue<string>());
        Assert.Equal(3, tables["matches"]![0]!["rows"]!.GetValue<int>());

        var slides = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "slide:contains('Zebra')"))));
        Assert.Equal(1, slides["count"]!.GetValue<int>());
    }

    [Fact]
    public void Query_TcAttribute_IsInvalidArgs()
    {
        CreateWithTable();
        TestEnv.AssertFail(
            _handler.Query(_ws.Ctx("deck.pptx", ("selector", "tc[name=foo]"))),
            ErrorCodes.InvalidArgs);
    }

    // ----- guards ----------------------------------------------------------------------

    [Fact]
    public void Replace_OnTablePath_IsInvalidArgs()
    {
        CreateWithTable();
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("replace", "/slide[1]/table[1]",
                props: TestEnv.Props(("find", "x")))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void MoveTable_ZOrderWorks_ButRowsAndCellsDoNot()
    {
        CreateWithTable();
        Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "on top"))));

        var moved = Edit(TestEnv.Op("move", "/slide[1]/table[1]", position: "front"));
        Assert.StartsWith("/slide[1]/shape[@id=", moved["results"]![0]!["target"]!.GetValue<string>(), StringComparison.Ordinal);

        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("move", "/slide[1]/table[1]/tr[1]", position: "front")]),
            ErrorCodes.InvalidArgs);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
