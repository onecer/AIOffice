using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>M5 deep tables: merges, borders, shading, widths, valign, padding.</summary>
public sealed class TableDeepTests : WordTestBase
{
    private JsonNode Get(string file, string path) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    private string RenderHtml(string file) =>
        Data(Handler.Render(Ctx(file, new JsonObject { ["to"] = "html" })))["content"]!.GetValue<string>();

    /// <summary>A 3x3 table with addressable cell texts r{row}c{col}.</summary>
    private string CreateWith3x3()
    {
        var file = CreateDoc(title: "Tables");
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":3,"columns":3}}]""");
        var ops = new JsonArray();
        for (var r = 1; r <= 3; r++)
        {
            for (var c = 1; c <= 3; c++)
            {
                ops.Add(new JsonObject
                {
                    ["op"] = "set",
                    ["path"] = $"/body/table[1]/tr[{r}]/tc[{c}]",
                    ["props"] = new JsonObject { ["text"] = $"r{r}c{c}" },
                });
            }
        }

        Edit(file, ops.ToJsonString());
        return file;
    }

    // ----------------------------------------------------------- mergeRight

    [Fact]
    public void Merge_right_sets_grid_span_absorbs_cells_and_renders_colspan()
    {
        var file = CreateWith3x3();

        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"mergeRight":2}}]""");

        var props = Get(file, "/body/table[1]/tr[1]/tc[1]")["properties"]!;
        Assert.Equal(2, props["colspan"]!.GetValue<int>());
        Assert.Equal(1, props["rowspan"]!.GetValue<int>());
        Assert.Contains("r1c2", props["text"]!.GetValue<string>(), StringComparison.Ordinal); // content absorbed

        // Row 1 now has two tc elements; the table still reports 3 grid columns.
        Assert.Equal(2, Get(file, "/body/table[1]/tr[1]")["properties"]!["cells"]!.AsArray().Count);
        Assert.Equal(3, Get(file, "/body/table[1]")["properties"]!["columns"]!.GetValue<int>());

        Assert.Contains("colspan=\"2\"", RenderHtml(file), StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Merge_right_1_unmerges_and_restores_the_grid()
    {
        var file = CreateWith3x3();
        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"mergeRight":3}}]""");

        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"mergeRight":1}}]""");

        Assert.Equal(1, Get(file, "/body/table[1]/tr[1]/tc[1]")["properties"]!["colspan"]!.GetValue<int>());
        Assert.Equal(3, Get(file, "/body/table[1]/tr[1]")["properties"]!["cells"]!.AsArray().Count);
        Assert.DoesNotContain("colspan", RenderHtml(file), StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Merge_right_past_the_row_end_is_invalid_args()
    {
        var file = CreateWith3x3();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[2]","props":{"mergeRight":3}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    // ------------------------------------------------------------ mergeDown

    [Fact]
    public void Merge_down_builds_a_vmerge_chain_and_renders_rowspan()
    {
        var file = CreateWith3x3();

        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"mergeDown":2}}]""");

        var head = Get(file, "/body/table[1]/tr[1]/tc[1]")["properties"]!;
        Assert.Equal(1, head["colspan"]!.GetValue<int>());
        Assert.Equal(2, head["rowspan"]!.GetValue<int>());
        Assert.Contains("r2c1", head["text"]!.GetValue<string>(), StringComparison.Ordinal); // content moved up

        var continuation = Get(file, "/body/table[1]/tr[2]/tc[1]")["properties"]!;
        Assert.Equal(0, continuation["rowspan"]!.GetValue<int>());

        var html = RenderHtml(file);
        Assert.Contains("rowspan=\"2\"", html, StringComparison.Ordinal);
        // The continuation slot does not render its own td.
        Assert.DoesNotContain("data-aio-path=\"/body/table[1]/tr[2]/tc[1]\"", html, StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Merge_down_1_dissolves_the_chain()
    {
        var file = CreateWith3x3();
        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"mergeDown":3}}]""");

        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"mergeDown":1}}]""");

        Assert.Equal(1, Get(file, "/body/table[1]/tr[1]/tc[1]")["properties"]!["rowspan"]!.GetValue<int>());
        Assert.Equal(1, Get(file, "/body/table[1]/tr[2]/tc[1]")["properties"]!["rowspan"]!.GetValue<int>());
        Assert.DoesNotContain("rowspan", RenderHtml(file), StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Merge_both_directions_renders_a_2x2_block()
    {
        var file = CreateWith3x3();

        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"mergeRight":2,"mergeDown":2}}]""");

        var props = Get(file, "/body/table[1]/tr[1]/tc[1]")["properties"]!;
        Assert.Equal(2, props["colspan"]!.GetValue<int>());
        Assert.Equal(2, props["rowspan"]!.GetValue<int>());

        var html = RenderHtml(file);
        Assert.Contains("colspan=\"2\" rowspan=\"2\"", html, StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Merge_down_on_a_continuation_slot_is_invalid_args()
    {
        var file = CreateWith3x3();
        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"mergeDown":2}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]/tr[2]/tc[1]","props":{"mergeDown":2}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("restart", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Merge_down_past_the_last_row_is_invalid_args()
    {
        var file = CreateWith3x3();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]/tr[2]/tc[1]","props":{"mergeDown":3}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    // -------------------------------------------------- borders and shading

    [Fact]
    public void Borders_outer_with_color_and_width_reopen_verify()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [{"op":"set","path":"/body/table[1]","props":{"borders":"outer","borderColor":"1F4E79","borderWidthPt":1.5}}]
            """);

        var props = Get(file, "/body/table[1]")["properties"]!;
        Assert.Equal("outer", props["borders"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Borders_none_and_all_round_trip_through_get()
    {
        var file = CreateWith3x3();

        Edit(file, """[{"op":"set","path":"/body/table[1]","props":{"borders":"none"}}]""");
        Assert.Equal("none", Get(file, "/body/table[1]")["properties"]!["borders"]!.GetValue<string>());

        Edit(file, """[{"op":"set","path":"/body/table[1]","props":{"borders":"all"}}]""");
        Assert.Equal("all", Get(file, "/body/table[1]")["properties"]!["borders"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Table_and_cell_shading_reopen_verify()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [
              {"op":"set","path":"/body/table[1]","props":{"shading":"F2F2F2"}},
              {"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"shading":"FFE699"}}
            ]
            """);

        Assert.Equal("F2F2F2", Get(file, "/body/table[1]")["properties"]!["shading"]!.GetValue<string>());
        Assert.Equal("FFE699", Get(file, "/body/table[1]/tr[2]/tc[2]")["properties"]!["shading"]!.GetValue<string>());

        // "none" removes it again.
        Edit(file, """[{"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"shading":"none"}}]""");
        Assert.Null(Get(file, "/body/table[1]/tr[2]/tc[2]")["properties"]!["shading"]);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Header_row_bolds_and_repeats_row_one()
    {
        var file = CreateWith3x3();

        Edit(file, """[{"op":"set","path":"/body/table[1]","props":{"headerRow":true}}]""");

        Assert.True(Get(file, "/body/table[1]")["properties"]!["headerRow"]!.GetValue<bool>());
        Assert.True(Get(file, "/body/table[1]/tr[1]/tc[1]/p[1]")["properties"]!["bold"]!.GetValue<bool>());

        Edit(file, """[{"op":"set","path":"/body/table[1]","props":{"headerRow":false}}]""");
        Assert.False(Get(file, "/body/table[1]")["properties"]!["headerRow"]!.GetValue<bool>());
        AssertValidatesClean(file);
    }

    // ------------------------------------------------------ widths and more

    [Fact]
    public void Column_widths_apply_to_grid_and_cells()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [{"op":"set","path":"/body/table[1]","props":{"columnWidths":["3cm","auto","2.5cm"]}}]
            """);

        var widths = Get(file, "/body/table[1]")["properties"]!["columnWidthsCm"]!.AsArray();
        Assert.Equal(3.0, widths[0]!.GetValue<double>(), 2);
        Assert.Equal("auto", widths[1]!.GetValue<string>());
        Assert.Equal(2.5, widths[2]!.GetValue<double>(), 2);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Column_width_count_mismatch_is_invalid_args()
    {
        var file = CreateWith3x3();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]","props":{"columnWidths":["3cm","2cm"]}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("3 column", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Width_alignment_padding_and_valign_reopen_verify()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [
              {"op":"set","path":"/body/table[1]","props":{"width":"100%","alignment":"center","cellPaddingCm":0.15}},
              {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"valign":"center"}}
            ]
            """);

        var table = Get(file, "/body/table[1]")["properties"]!;
        Assert.Equal("100%", table["width"]!.GetValue<string>());
        Assert.Equal("center", table["alignment"]!.GetValue<string>());
        Assert.Equal(0.15, table["cellPaddingCm"]!.GetValue<double>(), 2);
        Assert.Equal("center", Get(file, "/body/table[1]/tr[1]/tc[1]")["properties"]!["valign"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Fixed_width_table_in_cm_reports_back()
    {
        var file = CreateWith3x3();

        Edit(file, """[{"op":"set","path":"/body/table[1]","props":{"width":"12cm"}}]""");

        Assert.Equal("12cm", Get(file, "/body/table[1]")["properties"]!["width"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    // --------------------------------------------------------------- guards

    [Fact]
    public void Unknown_cell_prop_names_the_supported_set()
    {
        var file = CreateWith3x3();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"colour":"FF0000"}}]"""));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("mergeRight", ex.Candidates!);
    }

    [Fact]
    public void Unknown_table_prop_names_the_supported_set()
    {
        var file = CreateWith3x3();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]","props":{"border":"thick"}}]"""));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("borders", ex.Candidates!);
    }

    [Fact]
    public void Tracked_cell_merge_is_refused_honestly()
    {
        var file = CreateWith3x3();

        var ex = Assert.Throws<AiofficeException>(() => Edit(
            file,
            """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"mergeRight":2}}]""",
            new JsonObject { ["track"] = true }));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }

    [Fact]
    public void Set_text_still_works_alongside_cell_props()
    {
        var file = CreateWith3x3();

        Edit(file, """
            [{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"Total","shading":"D9E2F3","valign":"bottom"}}]
            """);

        var props = Get(file, "/body/table[1]/tr[1]/tc[1]")["properties"]!;
        Assert.Equal("Total", props["text"]!.GetValue<string>());
        Assert.Equal("D9E2F3", props["shading"]!.GetValue<string>());
        Assert.Equal("bottom", props["valign"]!.GetValue<string>());
        AssertValidatesClean(file);
    }
}
