using System.IO.Compression;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
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

    // M10: properties now nest under data.properties.{core,custom} (unified
    // docx/xlsx/pptx contract shape); the helpers unwrap that one level.
    private JsonObject GetProperties() =>
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/properties"))))["properties"]!.AsObject();

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
        // M10: nested under data.properties.{core,custom} (unified contract shape).
        Assert.Equal("Dana", view["properties"]!["core"]!["author"]!.GetValue<string>());
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

    /// <summary>
    /// Migrate-on-read: a deck whose only metadata is the legacy non-standard
    /// <c>.psmdcp</c> PackageProperties part (as old builds wrote) still surfaces
    /// its title through <c>read --view properties</c>, and the next <c>set</c>
    /// standardizes the file — removing the <c>.psmdcp</c> part for a conventional
    /// core-properties part the SDK can open.
    /// </summary>
    [Fact]
    public void LegacyPackageProperties_ReadAsFallback_ThenStandardizedOnWrite()
    {
        var path = _ws.PathOf("legacy.pptx");

        // Build a legacy deck the old way: title only in PackageProperties, which
        // System.IO.Packaging stores in a .psmdcp part (never docProps/core.xml).
        using (var doc = PresentationDocument.Create(path, DocumentFormat.OpenXml.PresentationDocumentType.Presentation))
        {
            var pp = doc.AddPresentationPart();
            pp.Presentation = new DocumentFormat.OpenXml.Presentation.Presentation(
                new DocumentFormat.OpenXml.Presentation.SlideMasterIdList(),
                new DocumentFormat.OpenXml.Presentation.SlideIdList(),
                new DocumentFormat.OpenXml.Presentation.SlideSize { Cx = 12_192_000, Cy = 6_858_000 },
                new DocumentFormat.OpenXml.Presentation.NotesSize { Cx = 6_858_000L, Cy = 9_144_000L });
            doc.PackageProperties.Title = "Legacy Deck";
            doc.PackageProperties.Creator = "Old Author";
        }

        // The legacy .psmdcp part is the only metadata store at this point.
        using (var zip = ZipFile.OpenRead(path))
        {
            Assert.Null(zip.GetEntry("docProps/core.xml"));
            Assert.Contains(zip.Entries, e => e.FullName.Contains(".psmdcp", StringComparison.Ordinal));
        }

        // Migrate-on-read: the title surfaces despite the non-standard storage.
        var view = TestEnv.AssertOk(_handler.Read(_ws.Ctx("legacy.pptx", ("view", "properties"))))["properties"]!;
        Assert.Equal("Legacy Deck", view["core"]!["title"]!.GetValue<string>());
        Assert.Equal("Old Author", view["core"]!["author"]!.GetValue<string>());

        // A set standardizes the file: the legacy part is gone and the title round-trips.
        TestEnv.AssertOk(_handler.Edit(
            _ws.Ctx("legacy.pptx"),
            [TestEnv.Op("set", "/properties", props: Set(("subject", JsonValue.Create("Migrated"))))]));

        using (var zip = ZipFile.OpenRead(path))
        {
            Assert.DoesNotContain(zip.Entries, e => e.FullName.Contains(".psmdcp", StringComparison.Ordinal));
        }

        using (var doc = PresentationDocument.Open(path, false))
        {
            Assert.NotNull(doc.CoreFilePropertiesPart);
            Assert.Equal("Legacy Deck", doc.CoreFilePropertiesPart!.CoreFileProperties.Title);
        }

        var after = TestEnv.AssertOk(_handler.Read(_ws.Ctx("legacy.pptx", ("view", "properties"))))["properties"]!;
        Assert.Equal("Legacy Deck", after["core"]!["title"]!.GetValue<string>());
        Assert.Equal("Old Author", after["core"]!["author"]!.GetValue<string>());
        Assert.Equal("Migrated", after["core"]!["subject"]!.GetValue<string>());
    }

    /// <summary>
    /// The hardening contract: a set title must land in the conventional
    /// <c>docProps/core.xml</c> part (the SDK CoreFilePropertiesPart) — not the
    /// non-standard <c>.psmdcp</c> PackageProperties part — so Office and a plain
    /// <c>unzip</c> can see it. Opens the SAVED package and asserts the part and
    /// the <c>dc:title</c> element exist, beyond the envelope round-trip.
    /// </summary>
    [Fact]
    public void SetTitle_WritesStandardCorePropertiesPart()
    {
        Create();
        Edit(TestEnv.Op("set", "/properties", props: Set(("title", JsonValue.Create("Quarterly Review")))));

        var path = _ws.PathOf("deck.pptx");

        // The SDK sees the standard part at docProps/core.xml carrying the title.
        using (var doc = PresentationDocument.Open(path, false))
        {
            var part = doc.CoreFilePropertiesPart;
            Assert.NotNull(part);
            Assert.Equal("/docProps/core.xml", part!.Uri.ToString());
            Assert.Equal("Quarterly Review", part.CoreFileProperties.Title);
        }

        // A plain unzip finds docProps/core.xml with the dc:title element (and no
        // legacy .psmdcp part is the storage location).
        using (var zip = ZipFile.OpenRead(path))
        {
            var core = zip.GetEntry("docProps/core.xml");
            Assert.NotNull(core);
            using var reader = new StreamReader(core!.Open());
            var xml = reader.ReadToEnd();
            Assert.Contains("<dc:title>Quarterly Review</dc:title>", xml, StringComparison.Ordinal);
            Assert.DoesNotContain(".psmdcp", zip.Entries.Select(e => e.FullName).Aggregate(string.Empty, (a, b) => a + b), StringComparison.Ordinal);
        }
    }
}
