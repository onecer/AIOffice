using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Preview.Tests;

public sealed class SelectionRouteTests : PreviewTestBase
{
    [Fact]
    public async Task Initial_selection_is_empty_with_the_current_rev()
    {
        var file = CreateDocx();
        using var server = StartServer(file);

        var state = JsonNode.Parse(await GetStringAsync(server, "selection"))!;

        Assert.Empty(state["paths"]!.AsArray());
        Assert.Equal(Rev.OfFile(file), state["rev"]!.GetValue<string>());
        Assert.True(state["updatedAt"]!.GetValue<DateTimeOffset>() <= DateTimeOffset.UtcNow.AddMinutes(1));
    }

    [Fact]
    public async Task Post_replaces_the_selection_and_get_round_trips_it()
    {
        var file = CreateDocx();
        using var server = StartServer(file);

        var response = await PostJsonAsync(
            server, "selection",
            """{"paths":["/body/p[2]","/body/table[1]","/slide[1]/shape[@id=2]","/Sheet1/A1:C10"]}""");
        Assert.Equal(System.Net.HttpStatusCode.OK, response.StatusCode);

        var state = JsonNode.Parse(await GetStringAsync(server, "selection"))!;
        var paths = state["paths"]!.AsArray().Select(p => p!.GetValue<string>()).ToList();
        Assert.Equal(["/body/p[2]", "/body/table[1]", "/slide[1]/shape[@id=2]", "/Sheet1/A1:C10"], paths);
        Assert.Equal(Rev.OfFile(file), state["rev"]!.GetValue<string>());

        // A later POST replaces (not merges) the state.
        await PostJsonAsync(server, "selection", """{"paths":[]}""");
        var cleared = JsonNode.Parse(await GetStringAsync(server, "selection"))!;
        Assert.Empty(cleared["paths"]!.AsArray());
    }

    [Fact]
    public async Task Invalid_paths_are_rejected_and_the_state_is_unchanged()
    {
        using var server = StartServer(CreateDocx());
        await PostJsonAsync(server, "selection", """{"paths":["/body/p[1]"]}""");

        var response = await PostJsonAsync(server, "selection", """{"paths":["not-a-path"]}""");

        Assert.Equal(System.Net.HttpStatusCode.BadRequest, response.StatusCode);
        var error = JsonNode.Parse(await response.Content.ReadAsStringAsync())!["error"]!;
        Assert.Equal(ErrorCodes.InvalidPath, error["code"]!.GetValue<string>());
        Assert.False(string.IsNullOrWhiteSpace(error["suggestion"]!.GetValue<string>()));

        var state = JsonNode.Parse(await GetStringAsync(server, "selection"))!;
        Assert.Equal("/body/p[1]", state["paths"]!.AsArray().Single()!.GetValue<string>());
    }

    [Fact]
    public async Task Malformed_bodies_are_invalid_args()
    {
        using var server = StartServer(CreateDocx());

        var notJson = await PostJsonAsync(server, "selection", "{nope");
        Assert.Equal(System.Net.HttpStatusCode.BadRequest, notJson.StatusCode);
        var error = JsonNode.Parse(await notJson.Content.ReadAsStringAsync())!["error"]!;
        Assert.Equal(ErrorCodes.InvalidArgs, error["code"]!.GetValue<string>());

        var noPaths = await PostJsonAsync(server, "selection", """{"selection":[]}""");
        Assert.Equal(System.Net.HttpStatusCode.BadRequest, noPaths.StatusCode);

        var nonString = await PostJsonAsync(server, "selection", """{"paths":[42]}""");
        Assert.Equal(System.Net.HttpStatusCode.BadRequest, nonString.StatusCode);
    }
}
