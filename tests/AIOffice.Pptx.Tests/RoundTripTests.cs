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

    /// <summary>The round-trip law over the M3 surface: charts, transitions, geometries and z-order.</summary>
    [Fact]
    public void M3Features_ReadSideVerbs_LeaveEveryByteUntouched()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Immutable"))));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(
                ("transition", "fade"), ("transitionDuration", "0.5s"))),
            TestEnv.Op("add", "/slide[1]", type: "chart", props: TestEnv.Props(
                ("kind", "bar"),
                ("categories", new JsonArray("Q1", "Q2")),
                ("series", new JsonArray(new JsonObject
                {
                    ["name"] = "Sales",
                    ["values"] = new JsonArray(10, 20),
                })),
                ("title", "Revenue"))),
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("shape", "ellipse"), ("fill", "FFEE00"))),
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("shape", "line"), ("flip", "v"))),
            TestEnv.Op("move", "/slide[1]/shape[2]", position: "front"),
        ]));
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"))));
        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/chart[1]"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/shape[1]"))));
        TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "shape"))));
        TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "html"))));
        TestEnv.AssertOk(_handler.Validate(_ws.Ctx("deck.pptx")));

        Assert.Equal(before, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }

    /// <summary>The round-trip law over the M4 surface: embedded chart workbooks, animations, comments.</summary>
    [Fact]
    public void M4Features_ReadSideVerbs_LeaveEveryByteUntouched()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Immutable"))));
        var added = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "animate me"))),
        ]));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "chart", props: TestEnv.Props(
                ("kind", "line"),
                ("categories", new JsonArray("A", "B")),
                ("series", new JsonArray(new JsonObject
                {
                    ["name"] = "S1",
                    ["values"] = new JsonArray(1, 2),
                })))),
            TestEnv.Op("add", shapePath, type: "animation", props: TestEnv.Props(
                ("effect", "flyIn"), ("direction", "left"), ("duration", "0.5s"))),
            TestEnv.Op("add", "/slide[1]", type: "comment", props: TestEnv.Props(
                ("text", "reviewed"), ("author", "Dana"))),
            TestEnv.Op("replace", "/slide[1]", props: TestEnv.Props(
                ("find", "animate"), ("replace", "animated"))),
        ]));
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"))));
        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "comments"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/chart[1]"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[1]"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/comment[@id=1]"))));
        TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "shape"))));
        TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "html"))));
        TestEnv.AssertOk(_handler.Validate(_ws.Ctx("deck.pptx")));

        Assert.Equal(before, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }

    /// <summary>The round-trip law over the M5 surface: tables (with merges), replies, emphasis/exit animations.</summary>
    [Fact]
    public void M5Features_ReadSideVerbs_LeaveEveryByteUntouched()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Immutable"))));
        var added = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "pulse me"))),
            TestEnv.Op("add", "/slide[1]", type: "table", props: TestEnv.Props(
                ("rows", 3), ("cols", 3), ("headerRow", true), ("style", "medium"), ("w", "20cm"))),
        ]));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("text", "Region"))),
            TestEnv.Op("set", "/slide[1]/table[1]/tr[2]/tc[1]", props: TestEnv.Props(("mergeDown", 1))),
            TestEnv.Op("add", shapePath, type: "animation", props: TestEnv.Props(("effect", "pulse"))),
            TestEnv.Op("add", shapePath, type: "animation", props: TestEnv.Props(
                ("effect", "fadeOut"), ("trigger", "afterPrevious"))),
            TestEnv.Op("add", "/slide[1]", type: "comment", props: TestEnv.Props(
                ("text", "thread root"), ("author", "Dana"))),
        ]));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]/comment[@id=1]", type: "reply", props: TestEnv.Props(
                ("text", "threaded reply"), ("author", "Riley"))),
        ]));
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"))));
        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "comments"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]/tr[2]/tc[1]"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[2]"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/comment[@id=2]"))));
        TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "tc:contains('Region')"))));
        TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "table"))));
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

        // M3 surface: a native chart, preset geometries, a flipped line and a fade transition.
        TestEnv.AssertOk(handler.Edit(edit, [
            TestEnv.Op("add", "/slide[4]", type: "slide", props: TestEnv.Props(
                ("title", "M3: chart, geometries & transition"))),
            TestEnv.Op("set", "/slide[4]", props: TestEnv.Props(
                ("transition", "fade"), ("transitionDuration", "0.7s"))),
            TestEnv.Op("add", "/slide[4]", type: "chart", props: TestEnv.Props(
                ("kind", "bar"),
                ("title", "Quarterly revenue"),
                ("categories", new JsonArray("Q1", "Q2", "Q3", "Q4")),
                ("series", new JsonArray(
                    new JsonObject { ["name"] = "2025", ["values"] = new JsonArray(12, 18, 15, 22) },
                    new JsonObject { ["name"] = "2026", ["values"] = new JsonArray(14, 21, 19, 27) })),
                ("x", "2cm"), ("y", "4.5cm"), ("w", "16cm"), ("h", "11cm"))),
            TestEnv.Op("add", "/slide[4]", type: "shape", props: TestEnv.Props(
                ("shape", "ellipse"), ("x", "20cm"), ("y", "4.5cm"), ("w", "5cm"), ("h", "3cm"),
                ("fill", "70AD47"), ("text", "ellipse"), ("align", "center"))),
            TestEnv.Op("add", "/slide[4]", type: "shape", props: TestEnv.Props(
                ("shape", "arrow"), ("x", "26cm"), ("y", "4.5cm"), ("w", "5cm"), ("h", "3cm"),
                ("fill", "ED7D31"))),
            TestEnv.Op("add", "/slide[4]", type: "shape", props: TestEnv.Props(
                ("shape", "line"), ("x", "20cm"), ("y", "9cm"), ("w", "11cm"), ("h", "3cm"),
                ("flip", "v"), ("fill", "4472C4"))),
        ]));

        // M4 surface: an animated shape, a slide comment and a find/replace pass.
        TestEnv.AssertOk(handler.Edit(edit, [
            TestEnv.Op("add", "/slide[5]", type: "slide", props: TestEnv.Props(
                ("title", "M4: animations, comments & replace"))),
            TestEnv.Op("add", "/slide[5]", type: "shape", props: TestEnv.Props(
                ("x", "4cm"), ("y", "6cm"), ("w", "12cm"), ("h", "4cm"),
                ("text", "This box fades in on click PLACEHOLDER"), ("fill", "DCE6F2"))),
        ]));
        var slide5 = TestEnv.AssertOk(handler.Get(
            new CommandContext
            {
                Workspace = ws,
                File = ctx.File,
                Args = new JsonObject { ["path"] = "/slide[5]" },
            }));
        var animatedShape = slide5["shapes"]!.AsArray()
            .Single(s => s!["text"]!.GetValue<string>().Contains("fades in", StringComparison.Ordinal))!["path"]!
            .GetValue<string>();
        TestEnv.AssertOk(handler.Edit(edit, [
            TestEnv.Op("add", animatedShape, type: "animation", props: TestEnv.Props(
                ("effect", "fade"), ("trigger", "click"), ("duration", "0.5s"))),
            TestEnv.Op("add", "/slide[5]", type: "comment", props: TestEnv.Props(
                ("text", "Added by aioffice M4 — check the comments pane."), ("author", "AIOffice"))),
            TestEnv.Op("replace", "/slide[5]", props: TestEnv.Props(
                ("find", "PLACEHOLDER"), ("replace", "(replaced by aioffice)"))),
        ]));

        // M5 surface: a native styled table with merges, an emphasis + exit animation pair
        // and a threaded comment reply.
        TestEnv.AssertOk(handler.Edit(edit, [
            TestEnv.Op("add", "/slide[6]", type: "slide", props: TestEnv.Props(
                ("title", "M5: tables, replies & emphasis/exit"))),
            TestEnv.Op("add", "/slide[6]", type: "table", props: TestEnv.Props(
                ("rows", 4), ("cols", 4), ("headerRow", true), ("style", "medium"),
                ("x", "2cm"), ("y", "5cm"), ("w", "20cm"))),
            TestEnv.Op("set", "/slide[6]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("text", "Region"))),
            TestEnv.Op("set", "/slide[6]/table[1]/tr[1]/tc[2]", props: TestEnv.Props(("text", "Q1"))),
            TestEnv.Op("set", "/slide[6]/table[1]/tr[1]/tc[3]", props: TestEnv.Props(("text", "Q2"))),
            TestEnv.Op("set", "/slide[6]/table[1]/tr[1]/tc[4]", props: TestEnv.Props(("text", "Total"))),
            TestEnv.Op("set", "/slide[6]/table[1]/tr[2]/tc[1]", props: TestEnv.Props(
                ("text", "EMEA (merged down)"), ("mergeDown", 1))),
            TestEnv.Op("set", "/slide[6]/table[1]/tr[4]/tc[1]", props: TestEnv.Props(
                ("text", "Grand total (merged right)"), ("mergeRight", 2), ("align", "center"))),
            TestEnv.Op("add", "/slide[6]", type: "shape", props: TestEnv.Props(
                ("x", "23cm"), ("y", "5cm"), ("w", "8cm"), ("h", "4cm"),
                ("text", "Pulses, then fades out"), ("fill", "FFC000"))),
        ]));
        var slide6 = TestEnv.AssertOk(handler.Get(
            new CommandContext
            {
                Workspace = ws,
                File = ctx.File,
                Args = new JsonObject { ["path"] = "/slide[6]" },
            }));
        var pulsingShape = slide6["shapes"]!.AsArray()
            .Single(s => s!["text"]!.GetValue<string>().Contains("Pulses", StringComparison.Ordinal))!["path"]!
            .GetValue<string>();
        TestEnv.AssertOk(handler.Edit(edit, [
            TestEnv.Op("add", pulsingShape, type: "animation", props: TestEnv.Props(("effect", "pulse"))),
            TestEnv.Op("add", pulsingShape, type: "animation", props: TestEnv.Props(
                ("effect", "fadeOut"), ("trigger", "afterPrevious"), ("delay", "1s"))),
            TestEnv.Op("add", "/slide[6]", type: "comment", props: TestEnv.Props(
                ("text", "M5 thread root — expand to see the reply."), ("author", "AIOffice"))),
        ]));
        TestEnv.AssertOk(handler.Edit(edit, [
            TestEnv.Op("add", "/slide[6]/comment[@id=2]", type: "reply", props: TestEnv.Props(
                ("text", "Threaded reply added by aioffice M5."), ("author", "Reviewer"))),
        ]));

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
