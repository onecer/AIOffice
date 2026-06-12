using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Pptx.Tests;

public sealed class RoundTripTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    /// <summary>
    /// The round-trip law for pptx: read-side verbs open the package from an
    /// in-memory copy and never save, so a deck that is read/queried/rendered/
    /// validated must remain byte-identical — every zip part included. (The
    /// handler only rewrites the file after a fully successful edit batch.)
    /// </summary>
    [Fact]
    public void ReadSideVerbs_LeaveEveryByteUntouched()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Immutable"))));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("title", "Second"))),
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "body text"))),
        ]));
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"))));
        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "shape"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/shape[1]"))));
        TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "html"))));
        TestEnv.AssertOk(_handler.Validate(_ws.Ctx("deck.pptx")));

        Assert.Equal(before, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }

    /// <summary>The round-trip law over the M2 surface: notes, background and picture parts included.</summary>
    [Fact]
    public void M2Features_ReadSideVerbs_LeaveEveryByteUntouched()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Immutable"))));
        File.WriteAllBytes(_ws.PathOf("logo.png"), TestImages.Png(40, 20));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(
                ("title", "Second"), ("background", "0F172A"))),
            TestEnv.Op("set", "/slide[1]/notes", props: TestEnv.Props(("text", "presenter cue"))),
            TestEnv.Op("add", "/slide[2]", type: "image", props: TestEnv.Props(
                ("src", "logo.png"), ("w", "8cm"))),
        ]));
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"))));
        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "text"))));
        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/notes"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[2]"))));
        TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "shape[kind=picture]"))));
        TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "html"))));
        TestEnv.AssertOk(_handler.Validate(_ws.Ctx("deck.pptx")));

        Assert.Equal(before, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }

    [Fact]
    public void FailedEditBatch_LeavesEveryByteUntouched()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Immutable"))));
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[2]", type: "slide"),
            TestEnv.Op("remove", "/slide[7]"),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
        Assert.Equal(before, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }

    /// <summary>
    /// Generates the human-check fixture: validator-clean is our automated
    /// proxy, opening fixtures/manual-check/pptx-showcase.pptx in PowerPoint
    /// or Keynote is the manual half of the law.
    /// </summary>
    [Fact]
    public void ManualCheckFixture_IsGeneratedAndValidatorClean()
    {
        var root = FindRepoRoot();
        var dir = root is null
            ? Path.Combine(_ws.Dir, "manual-check")
            : Path.Combine(root, "fixtures", "manual-check");
        Directory.CreateDirectory(dir);

        var file = Path.Combine(dir, "pptx-showcase.pptx");
        File.Delete(file);

        var ws = new Workspace(dir);
        var handler = new PptxHandler();
        var ctx = new CommandContext
        {
            Workspace = ws,
            File = ws.Resolve("pptx-showcase.pptx"),
            Args = new JsonObject { ["title"] = "AIOffice PPTX Showcase" },
        };
        TestEnv.AssertOk(handler.Create(ctx));

        var edit = new CommandContext { Workspace = ws, File = ctx.File, Args = [] };
        TestEnv.AssertOk(handler.Edit(edit, [
            TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("title", "Shapes & styles"))),
            TestEnv.Op("add", "/slide[2]", type: "shape", props: TestEnv.Props(
                ("x", JsonValue.Create(2)), ("y", JsonValue.Create(5)),
                ("w", JsonValue.Create(9)), ("h", JsonValue.Create(4)),
                ("text", JsonValue.Create("Solid yellow box\ncentered bold text")),
                ("fontSize", JsonValue.Create(20)), ("bold", JsonValue.Create(true)),
                ("align", JsonValue.Create("center")), ("fill", JsonValue.Create("FFEE00")))),
            TestEnv.Op("add", "/slide[2]", type: "shape", props: TestEnv.Props(
                ("x", JsonValue.Create(13)), ("y", JsonValue.Create(5)),
                ("w", JsonValue.Create(9)), ("h", JsonValue.Create(4)),
                ("text", JsonValue.Create("Red text on plain box")),
                ("color", JsonValue.Create("CC0000")))),
        ]));

        // M2 surface: a real p:bg background, speaker notes and an embedded picture.
        var logo = Path.Combine(dir, "showcase-logo.png");
        File.WriteAllBytes(logo, TestImages.Png(200, 100));
        TestEnv.AssertOk(handler.Edit(edit, [
            TestEnv.Op("add", "/slide[3]", type: "slide", props: TestEnv.Props(
                ("title", "M2: background, notes & image"),
                ("background", "E2E8F0"))),
            TestEnv.Op("set", "/slide[3]/notes", props: TestEnv.Props(
                ("text", "Speaker notes written by aioffice — check presenter view."))),
            TestEnv.Op("add", "/slide[3]", type: "image", props: TestEnv.Props(
                ("src", "showcase-logo.png"), ("x", "2cm"), ("y", "6cm"), ("w", "6cm"))),
        ]));
        File.Delete(logo);

        var validation = TestEnv.AssertOk(handler.Validate(
            new CommandContext { Workspace = ws, File = ctx.File, Args = [] }));
        Assert.True(
            validation["valid"]!.GetValue<bool>(),
            "showcase fixture has validator issues: " + validation["issues"]!.ToJsonString());
        Assert.True(File.Exists(file));
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
