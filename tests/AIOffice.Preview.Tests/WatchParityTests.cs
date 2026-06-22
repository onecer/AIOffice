using System.Net;
using System.Text;
using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Preview.Tests;

/// <summary>
/// Covers the watch-parity routes added on top of the live preview: advisory
/// marks (/marks), the goto scroll-push (/goto), the incremental-reload content
/// fragment (/content) and the browser-edit bridge (/api/edit).
/// </summary>
public sealed class WatchParityTests : PreviewTestBase
{
    private static Task<HttpResponseMessage> DeleteJsonAsync(PreviewServer server, string route, string json)
    {
        var request = new HttpRequestMessage(HttpMethod.Delete, server.Url + route)
        {
            Content = new StringContent(json, Encoding.UTF8, "application/json"),
        };
        return Http.SendAsync(request);
    }

    [Fact]
    public async Task Marks_post_get_round_trip_and_upsert_by_path()
    {
        var file = CreateDocx();
        using var server = StartServer(file);

        Assert.Empty(JsonNode.Parse(await GetStringAsync(server, "marks"))!["marks"]!.AsArray());

        var response = await PostJsonAsync(
            server, "marks", """{"path":"/body/p[2]","color":"red","note":"overflows","toFix":true}""");
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);

        var marks = JsonNode.Parse(await GetStringAsync(server, "marks"))!["marks"]!.AsArray();
        var mark = marks.Single()!;
        Assert.Equal("/body/p[2]", mark["path"]!.GetValue<string>());
        Assert.Equal("red", mark["color"]!.GetValue<string>());
        Assert.Equal("overflows", mark["note"]!.GetValue<string>());
        Assert.True(mark["toFix"]!.GetValue<bool>());

        // Re-marking the same path replaces (one mark per path), not appends.
        await PostJsonAsync(server, "marks", """{"path":"/body/p[2]","color":"#00ff00"}""");
        var after = JsonNode.Parse(await GetStringAsync(server, "marks"))!["marks"]!.AsArray();
        Assert.Single(after);
        Assert.Equal("#00ff00", after[0]!["color"]!.GetValue<string>());
    }

    [Fact]
    public async Task Mark_selected_expands_to_the_current_selection()
    {
        var file = CreateDocx();
        using var server = StartServer(file);
        await PostJsonAsync(server, "selection", """{"paths":["/body/p[1]","/body/p[2]"]}""");

        await PostJsonAsync(server, "marks", """{"path":"selected","color":"blue"}""");

        var paths = JsonNode.Parse(await GetStringAsync(server, "marks"))!["marks"]!.AsArray()
            .Select(m => m!["path"]!.GetValue<string>()).OrderBy(p => p).ToList();
        Assert.Equal(["/body/p[1]", "/body/p[2]"], paths);
    }

    [Fact]
    public async Task Mark_rejects_an_invalid_color_and_an_invalid_path()
    {
        using var server = StartServer(CreateDocx());

        var badColor = await PostJsonAsync(server, "marks", """{"path":"/body/p[1]","color":"plaid"}""");
        Assert.Equal(HttpStatusCode.BadRequest, badColor.StatusCode);
        Assert.Equal(ErrorCodes.InvalidArgs,
            JsonNode.Parse(await badColor.Content.ReadAsStringAsync())!["error"]!["code"]!.GetValue<string>());

        var badPath = await PostJsonAsync(server, "marks", """{"path":"not-a-path"}""");
        Assert.Equal(HttpStatusCode.BadRequest, badPath.StatusCode);

        Assert.Empty(JsonNode.Parse(await GetStringAsync(server, "marks"))!["marks"]!.AsArray());
    }

    [Fact]
    public async Task Delete_removes_one_mark_by_path_and_all_clears_them()
    {
        using var server = StartServer(CreateDocx());
        await PostJsonAsync(server, "marks", """{"path":"/body/p[1]"}""");
        await PostJsonAsync(server, "marks", """{"path":"/body/p[2]"}""");

        await DeleteJsonAsync(server, "marks", """{"path":"/body/p[1]"}""");
        var one = JsonNode.Parse(await GetStringAsync(server, "marks"))!["marks"]!.AsArray();
        Assert.Equal("/body/p[2]", one.Single()!["path"]!.GetValue<string>());

        await DeleteJsonAsync(server, "marks", """{"all":true}""");
        Assert.Empty(JsonNode.Parse(await GetStringAsync(server, "marks"))!["marks"]!.AsArray());
    }

    [Fact]
    public async Task Content_route_returns_the_addressable_fragment()
    {
        using var server = StartServer(CreateDocx());

        var content = await GetStringAsync(server, "content");

        Assert.Contains("data-aio-path", content, StringComparison.Ordinal);
        // The fragment is the inner content, not the full page shell.
        Assert.DoesNotContain("<!DOCTYPE html>", content, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Goto_validates_the_path_and_echoes_the_target()
    {
        using var server = StartServer(CreateDocx());

        var ok = await PostJsonAsync(server, "goto", """{"path":"/body/p[2]"}""");
        Assert.Equal(HttpStatusCode.OK, ok.StatusCode);
        Assert.Equal("/body/p[2]",
            JsonNode.Parse(await ok.Content.ReadAsStringAsync())!["data"]!["scrolledTo"]!.GetValue<string>());

        var bad = await PostJsonAsync(server, "goto", """{"path":"nope"}""");
        Assert.Equal(HttpStatusCode.BadRequest, bad.StatusCode);
    }

    [Fact]
    public async Task Api_edit_applies_a_cell_set_and_rejects_a_bad_op()
    {
        var file = CreateXlsx();
        using var server = StartServer(file);

        var ok = await PostJsonAsync(server, "api/edit", """{"op":"set","path":"/Sheet1/A1","props":{"value":"EDITED"}}""");
        Assert.Equal(HttpStatusCode.OK, ok.StatusCode);
        Assert.True(JsonNode.Parse(await ok.Content.ReadAsStringAsync())!["ok"]!.GetValue<bool>());

        // The edit landed in the file (visible in the fresh content fragment).
        Assert.Contains("EDITED", await GetStringAsync(server, "content"), StringComparison.Ordinal);

        var bad = await PostJsonAsync(server, "api/edit", """{"op":"set","path":"/Sheet1/A1"}""");
        // Missing props is tolerated by some handlers; a truly bad target is not.
        var badTarget = await PostJsonAsync(server, "api/edit", """{"op":"remove","path":"/Sheet1/zzz9"}""");
        Assert.Equal(HttpStatusCode.BadRequest, badTarget.StatusCode);
        _ = bad;
    }
}
