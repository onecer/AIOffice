using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class QueryGetTests : WordTestBase
{
    private string CreateSample()
    {
        var file = CreateDoc(title: "Annual Report");
        Edit(file, """
            [
              {"op":"add","path":"/body","props":{"text":"Q3 revenue grew."}},
              {"op":"add","path":"/body","props":{"text":"Methods","style":"Heading2"}},
              {"op":"add","path":"/body","props":{"text":"Bold facts","bold":true}},
              {"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}}
            ]
            """);
        return file;
    }

    private JsonNode Query(string file, string selector) =>
        Data(Handler.Query(Ctx(file, new JsonObject { ["selector"] = selector })));

    [Fact]
    public void Query_by_style_returns_canonical_paths()
    {
        var file = CreateSample();

        var data = Query(file, "p[style=Heading1]");

        Assert.Equal(1, data["count"]!.GetValue<int>());
        var match = data["matches"]!.AsArray()[0]!;
        Assert.Equal("/body/p[1]", match["path"]!.GetValue<string>());
        Assert.Equal("Annual Report", match["snippet"]!.GetValue<string>());
    }

    [Fact]
    public void Query_contains_is_case_insensitive()
    {
        var file = CreateSample();

        var data = Query(file, "p:contains('q3')");

        Assert.Equal(1, data["count"]!.GetValue<int>());
        Assert.Equal("/body/p[3]", data["matches"]!.AsArray()[0]!["path"]!.GetValue<string>());
    }

    [Fact]
    public void Query_runs_by_bold_flag()
    {
        var file = CreateSample();

        var paths = Query(file, "run[bold=true]")["matches"]!.AsArray()
            .Select(m => m!["path"]!.GetValue<string>())
            .ToList();

        Assert.Contains("/body/p[5]/run[1]", paths);
    }

    [Fact]
    public void Query_star_covers_tables_and_cells()
    {
        var file = CreateSample();

        var types = Query(file, "*")["matches"]!.AsArray()
            .Select(m => m!["type"]!.GetValue<string>())
            .ToHashSet(StringComparer.Ordinal);

        Assert.Superset(new HashSet<string> { "p", "run", "table", "tr", "tc" }, types);
    }

    [Fact]
    public void Query_unknown_element_lists_the_valid_ones()
    {
        var file = CreateSample();

        var ex = Assert.Throws<AiofficeException>(() =>
            Handler.Query(Ctx(file, new JsonObject { ["selector"] = "shape:contains('x')" })));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.NotNull(ex.Candidates);
        Assert.Contains("p", ex.Candidates);
    }

    [Fact]
    public void Get_run_returns_its_formatting()
    {
        var file = CreateSample();

        var props = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[5]/run[1]" })))["properties"]!;

        Assert.Equal("Bold facts", props["text"]!.GetValue<string>());
        Assert.True(props["bold"]!.GetValue<bool>());
    }

    [Fact]
    public void Get_paragraph_reports_text_style_and_run_count()
    {
        var file = CreateSample();

        var props = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[4]" })))["properties"]!;

        Assert.Equal("Methods", props["text"]!.GetValue<string>());
        Assert.Equal("Heading2", props["style"]!.GetValue<string>());
        Assert.Equal(1, props["runs"]!.GetValue<int>());
    }

    [Fact]
    public void Get_table_cell_via_full_table_path()
    {
        var file = CreateSample();

        var data = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/table[1]/tr[1]/tc[2]" })));

        Assert.Equal("tc", data["type"]!.GetValue<string>());
    }

    [Fact]
    public void Out_of_range_index_offers_nearest_candidates()
    {
        var file = CreateSample();

        var ex = Assert.Throws<AiofficeException>(() =>
            Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[42]" })));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.NotNull(ex.Candidates);
        Assert.All(ex.Candidates!, c => Assert.StartsWith("/body/p[", c, StringComparison.Ordinal));
    }

    [Fact]
    public void Wrong_child_type_offers_existing_children()
    {
        var file = CreateSample();

        var ex = Assert.Throws<AiofficeException>(() =>
            Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[1]/tc[1]" })));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.NotNull(ex.Candidates);
    }
}
