using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class EditTests : WordTestBase
{
    [Fact]
    public void Set_text_persists_and_validates()
    {
        var file = CreateDoc(title: "Before");

        var envelope = Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"After"}}]""");

        Assert.Equal(1, Data(envelope)["applied"]!.GetValue<int>());
        Assert.Equal("After", BodyTexts(file)[0]);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Set_formatting_round_trips_through_get()
    {
        var file = CreateDoc(title: "Styled");

        Edit(file, """
            [{"op":"set","path":"/body/p[1]","props":
              {"bold":true,"italic":true,"color":"#1A2B3C","fontSize":14,"alignment":"center"}}]
            """);

        var props = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[1]" })))["properties"]!;
        Assert.True(props["bold"]!.GetValue<bool>());
        Assert.True(props["italic"]!.GetValue<bool>());
        Assert.Equal("1A2B3C", props["color"]!.GetValue<string>());
        Assert.Equal(14, props["fontSize"]!.GetValue<double>());
        Assert.Equal("center", props["alignment"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Failed_batch_is_atomic_and_leaves_the_file_untouched()
    {
        var file = CreateDoc(title: "Atomic");
        var before = File.ReadAllBytes(file);

        var ex = Assert.Throws<AiofficeException>(() => Edit(file, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"changed"}},
              {"op":"set","path":"/body/p[99]","props":{"text":"boom"}}
            ]
            """));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.StartsWith("ops[1]", ex.Message, StringComparison.Ordinal);
        Assert.Equal(before, File.ReadAllBytes(file));
        Assert.Empty(Snapshots.List(file)); // nothing written -> nothing snapshotted
    }

    [Fact]
    public void Expect_rev_mismatch_is_stale_address_before_any_write()
    {
        var file = CreateDoc(title: "Rev");
        var before = File.ReadAllBytes(file);

        var ex = Assert.Throws<AiofficeException>(() => Edit(
            file,
            """[{"op":"set","path":"/body/p[1]","props":{"text":"x"}}]""",
            new JsonObject { ["expectRev"] = "000000000000" }));

        Assert.Equal(ErrorCodes.StaleAddress, ex.Code);
        Assert.Equal(before, File.ReadAllBytes(file));
    }

    [Fact]
    public void Matching_expect_rev_lets_the_edit_through()
    {
        var file = CreateDoc(title: "Rev");

        var envelope = Edit(
            file,
            """[{"op":"set","path":"/body/p[1]","props":{"text":"fresh"}}]""",
            new JsonObject { ["expectRev"] = Rev.OfFile(file) });

        Assert.True(envelope.IsOk);
        Assert.Equal("fresh", BodyTexts(file)[0]);
    }

    [Fact]
    public void Every_successful_edit_snapshots_the_pre_image()
    {
        var file = CreateDoc(title: "Undo me");
        var preEditRev = Rev.OfFile(file);

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"edited"}}]""");

        var entry = Assert.Single(Snapshots.List(file));
        Assert.Equal(preEditRev, entry.Rev);
    }

    [Fact]
    public void Dry_run_applies_nothing()
    {
        var file = CreateDoc(title: "Dry");
        var before = File.ReadAllBytes(file);

        var envelope = Edit(
            file,
            """[{"op":"set","path":"/body/p[1]","props":{"text":"phantom"}}]""",
            new JsonObject { ["dryRun"] = true });

        Assert.True(Data(envelope)["dryRun"]!.GetValue<bool>());
        Assert.Equal(before, File.ReadAllBytes(file));
        Assert.Empty(Snapshots.List(file));
    }

    [Fact]
    public void Add_heading_via_style_sugar_lands_at_reported_path()
    {
        var file = CreateDoc(title: "Doc");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"p","position":"after",
              "props":{"text":"Background","style":"Heading2"}}]
            """);

        var added = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("/body/p[2]", added["path"]!.GetValue<string>());
        var props = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[2]" })))["properties"]!;
        Assert.Equal("Heading2", props["style"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Add_table_then_row_then_read_back()
    {
        var file = CreateDoc();

        Edit(file, """
            [
              {"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}},
              {"op":"add","path":"/body/table[1]","type":"tr","props":{"cells":["x","y"]}}
            ]
            """);

        var props = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/table[1]" })))["properties"]!;
        Assert.Equal(3, props["rows"]!.GetValue<int>());
        Assert.Equal(2, props["columns"]!.GetValue<int>());
        var row = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/table[1]/tr[3]" })))["properties"]!;
        Assert.Equal(new[] { "x", "y" }, row["cells"]!.AsArray().Select(c => c!.GetValue<string>()));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Remove_paragraph_shifts_following_paths()
    {
        var file = CreateDoc(title: "Keep");
        Edit(file, """[{"op":"add","path":"/body","props":{"text":"Drop me"}}]""");

        Edit(file, """[{"op":"remove","path":"/body/p[3]"}]""");

        Assert.DoesNotContain("Drop me", BodyTexts(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Move_paragraph_before_another()
    {
        var file = CreateDoc();
        Edit(file, """
            [
              {"op":"add","path":"/body","props":{"text":"A"}},
              {"op":"add","path":"/body","props":{"text":"B"}}
            ]
            """);

        Edit(file, """[{"op":"move","path":"/body/p[3]","position":"before /body/p[2]"}]""");

        Assert.Equal(new[] { string.Empty, "B", "A" }, BodyTexts(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Unsupported_prop_names_the_nearest_supported_one()
    {
        var file = CreateDoc(title: "Prop");

        // "fontSiz" is a typo of a real prop; the suggestion points at the nearest supported name.
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"fontSiz":"12"}}]"""));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("fontSize", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Removing_the_last_cell_paragraph_is_refused_with_a_workaround()
    {
        var file = CreateDoc();
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":1,"columns":1}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"remove","path":"/body/table[1]/tr[1]/tc[1]/p[1]"}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("set", ex.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Invalid_path_carries_nearest_match_candidates()
    {
        var file = CreateDoc(title: "Candidates");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/p[9]","props":{"text":"x"}}]"""));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.NotNull(ex.Candidates);
        Assert.Contains("/body/p[2]", ex.Candidates);
    }
}
