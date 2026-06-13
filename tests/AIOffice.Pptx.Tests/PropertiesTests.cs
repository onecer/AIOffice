using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Pptx.Tests;

public sealed class PropertiesTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() => TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Q3"))));

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private JsonObject GetProperties() =>
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/properties"))));

    private static JsonObject Set(params (string Key, JsonNode? Value)[] pairs) => TestEnv.Props(pairs);

    [Fact]
    public void SetCoreProperties_ReopenVerify()
    {
        Create();
        Edit(TestEnv.Op("set", "/properties", props: Set(
            ("title", JsonValue.Create("Quarterly Review")),
            ("subject", JsonValue.Create("Sales")),
            ("author", JsonValue.Create("Dana")),
            ("keywords", JsonValue.Create("q3;sales")),
            ("category", JsonValue.Create("Reports")),
            ("comments", JsonValue.Create("draft")),
            ("lastModifiedBy", JsonValue.Create("Dana")),
            ("revision", JsonValue.Create("3")))));

        var core = GetProperties()["core"]!.AsObject();
        Assert.Equal("Quarterly Review", core["title"]!.GetValue<string>());
        Assert.Equal("Sales", core["subject"]!.GetValue<string>());
        Assert.Equal("Dana", core["author"]!.GetValue<string>());
        Assert.Equal("q3;sales", core["keywords"]!.GetValue<string>());
        Assert.Equal("Reports", core["category"]!.GetValue<string>());
        Assert.Equal("draft", core["comments"]!.GetValue<string>());
        Assert.Equal("Dana", core["lastModifiedBy"]!.GetValue<string>());
        Assert.Equal("3", core["revision"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetCustomProperties_TypedRoundTrip()
    {
        Create();
        Edit(TestEnv.Op("set", "/properties", props: Set(
            ("custom", new JsonObject
            {
                ["Project"] = "Acme",
                ["Reviewed"] = true,
                ["Budget"] = 1000,
                ["Due"] = "2026-07-01",
            }))));

        var custom = GetProperties()["custom"]!.AsObject();
        Assert.Equal("Acme", custom["Project"]!.GetValue<string>());
        Assert.True(custom["Reviewed"]!.GetValue<bool>());
        Assert.Equal(1000, custom["Budget"]!.GetValue<double>());
        Assert.Contains("2026-07-01", custom["Due"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetCreatedModified_AreIsoDates()
    {
        Create();
        Edit(TestEnv.Op("set", "/properties", props: Set(
            ("created", JsonValue.Create("2026-06-01")),
            ("modified", JsonValue.Create("2026-06-13T10:00:00Z")))));

        var core = GetProperties()["core"]!.AsObject();
        Assert.StartsWith("2026-06-01", core["created"]!.GetValue<string>());
        Assert.StartsWith("2026-06-13", core["modified"]!.GetValue<string>());
    }

    [Fact]
    public void EmptyStringClearsCoreProperty()
    {
        Create();
        Edit(TestEnv.Op("set", "/properties", props: Set(("author", JsonValue.Create("Dana")))));
        Assert.Equal("Dana", GetProperties()["core"]!["author"]!.GetValue<string>());

        Edit(TestEnv.Op("set", "/properties", props: Set(("author", JsonValue.Create("")))));
        Assert.Null(GetProperties()["core"]!["author"]);
    }

    [Fact]
    public void NullClearsCustomProperty()
    {
        Create();
        Edit(TestEnv.Op("set", "/properties", props: Set(("custom", new JsonObject { ["Project"] = "Acme" }))));
        Assert.True(GetProperties()["custom"]!.AsObject().ContainsKey("Project"));

        Edit(TestEnv.Op("set", "/properties", props: Set(("custom", new JsonObject { ["Project"] = null }))));
        Assert.False(GetProperties()["custom"]!.AsObject().ContainsKey("Project"));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ReadViewProperties_MatchesGet()
    {
        Create();
        Edit(TestEnv.Op("set", "/properties", props: Set(("author", JsonValue.Create("Dana")))));

        var view = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "properties"))));
        Assert.Equal("Dana", view["core"]!["author"]!.GetValue<string>());
        Assert.Equal("/properties", view["path"]!.GetValue<string>());
    }

    [Fact]
    public void UnknownCoreProperty_IsUnsupported()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/properties", props: Set(("manager", JsonValue.Create("Sam"))))]);
        TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
    }

    [Fact]
    public void RemoveProperties_IsRejectedWithSuggestion()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("remove", "/properties")]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }
}
