using Xunit;

namespace AIOffice.Preview.Tests;

/// <summary>GET / must serve interactive HTML where addressable nodes carry data-aio-path.</summary>
public sealed class RenderRouteTests : PreviewTestBase
{
    [Fact]
    public async Task Docx_page_maps_blocks_to_canonical_body_paths()
    {
        using var server = StartServer(CreateDocx());

        var page = await GetStringAsync(server);

        Assert.Contains("<h1 data-aio-path=\"/body/p[1]\">Preview Heading</h1>", page, StringComparison.Ordinal);
        Assert.Contains("<p data-aio-path=\"/body/p[2]\">Hello preview</p>", page, StringComparison.Ordinal);
        Assert.Contains("<table data-aio-path=\"/body/table[1]\">", page, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Xlsx_page_maps_cells_to_sheet_cell_paths()
    {
        using var server = StartServer(CreateXlsx());

        var page = await GetStringAsync(server);

        // The contract: every rendered cell maps back to its canonical path.
        // (Whether empty cells inside the used range get a td is up to the
        // renderer in use — the handler's own render skips them, the preview
        // fallback grid emits them — so only the populated cells are pinned.)
        Assert.Contains("<td data-aio-path=\"/Sheet1/A1\">Name</td>", page, StringComparison.Ordinal);
        Assert.Contains("<td data-aio-path=\"/Sheet1/B2\">42</td>", page, StringComparison.Ordinal);
        Assert.Contains("data-sheet=\"Sheet1\"", page, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Pptx_page_wraps_each_shape_in_a_group_with_the_id_form_path()
    {
        using var server = StartServer(CreatePptx());

        var page = await GetStringAsync(server);

        Assert.Contains("<svg", page, StringComparison.Ordinal);
        Assert.Contains("<g data-aio-path=\"/slide[1]/shape[@id=", page, StringComparison.Ordinal);
        Assert.Contains("</g>", page, StringComparison.Ordinal);
        Assert.Contains("Preview Deck", page, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Page_ships_the_selection_and_live_reload_layer()
    {
        using var server = StartServer(CreateDocx());

        var page = await GetStringAsync(server);

        Assert.Contains("EventSource(\"/events\")", page, StringComparison.Ordinal);
        Assert.Contains("fetch(\"/selection\"", page, StringComparison.Ordinal);
        Assert.Contains(".aio-selected", page, StringComparison.Ordinal);
        Assert.Contains("2px solid", page, StringComparison.Ordinal); // the selection outline
    }
}
