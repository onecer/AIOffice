using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// v1.24 docx underline widening: the 'underline' prop RELAXES from bool to also accept a
/// named style (double/thick/dotted/dash/dashLong/dotDash/dotDotDash/wave/wavyHeavy/wavyDouble/
/// words/single/none) on run set/add AND on style-definition set. The bool form stays FIRST +
/// byte-stable; the read is discriminated (bool for single/none, style STRING otherwise) mirroring
/// the shipped ReadOutline precedent. All sites (run get, style-def get, query) covered here.
/// </summary>
public sealed class UnderlineStyleTests : WordTestBase
{
    private JsonNode GetRun(string file, string path = "/body/p[1]/run[1]") =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })))["properties"]!;

    private JsonNode GetStyle(string file, string id) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = $"/style[@id={id}]" })))["properties"]!;

    /// <summary>The first run's saved w:u element (null when absent).</summary>
    private static Underline? FirstRunUnderline(string file)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        return doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().First().RunProperties?.Underline;
    }

    // ------------------------------------------------------------ 1. round-trip style

    [Fact]
    public void Run_underline_double_writes_val_double_and_round_trips()
    {
        var file = CreateDoc(title: "Double underline");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"heading","underline":"double"}}]""");

        Assert.Equal(UnderlineValues.Double, FirstRunUnderline(file)!.Val!.Value);
        Assert.Equal("double", GetRun(file)["underline"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Run_underline_wave_round_trips_through_the_run_path()
    {
        var file = CreateDoc(title: "Wavy underline");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"note","underline":"wave"}}]""");

        Assert.Equal(UnderlineValues.Wave, FirstRunUnderline(file)!.Val!.Value);
        Assert.Equal("wave", GetRun(file)["underline"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void StyleDefinition_underline_wave_round_trips_through_the_style_path()
    {
        var file = CreateDoc(title: "Wavy style");

        Edit(file, """[{"op":"add","path":"/styles","type":"style","props":{"id":"Wavy","kind":"character","underline":"wave"}}]""");
        Assert.Equal("wave", GetStyle(file, "Wavy")["underline"]!.GetValue<string>());
        AssertValidatesClean(file);

        // set on an existing style-def mirrors the same widening.
        Edit(file, """[{"op":"set","path":"/style[@id=Wavy]","props":{"underline":"dotDash"}}]""");
        Assert.Equal("dotDash", GetStyle(file, "Wavy")["underline"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    // ------------------------------------------------------------ 2. byte-stable bool form

    [Fact]
    public void Run_underline_true_still_writes_val_single_and_reads_bool_true()
    {
        var file = CreateDoc(title: "Single underline");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"link","underline":true}}]""");

        // The bool branch is byte-stable: exactly @val=single, exactly the same element old code produced.
        var u = FirstRunUnderline(file)!;
        Assert.Equal(UnderlineValues.Single, u.Val!.Value);
        Assert.Equal(
            new Underline { Val = UnderlineValues.Single }.OuterXml,
            u.OuterXml);

        // get STILL returns a JSON bool true (not the string "single").
        var read = GetRun(file)["underline"]!;
        Assert.Equal(JsonValueKind.True, read.GetValueKind());
        Assert.True(read.GetValue<bool>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Underline_true_run_properties_are_byte_identical_across_authorings()
    {
        // The load-bearing byte surface is the run's w:rPr (the whole package embeds timestamps,
        // so hash the rPr XML). Two independent single-underline runs hash-compare equal — the
        // bool path stayed byte-stable, single-underline unchanged.
        var a = CreateDoc(name: "a.docx", title: "Stable");
        var b = CreateDoc(name: "b.docx", title: "Stable");
        Edit(a, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"x","underline":true}}]""");
        Edit(b, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"x","underline":true}}]""");

        Assert.Equal(Sha(RunPrXml(a)), Sha(RunPrXml(b)));

        static string RunPrXml(string file)
        {
            using var doc = WordprocessingDocument.Open(file, isEditable: false);
            return doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().First().RunProperties!.OuterXml;
        }

        static string Sha(string xml) => Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(xml)));
    }

    // ------------------------------------------------------------ 3. clear semantics

    [Fact]
    public void Underline_false_and_none_both_clear_to_val_none_and_read_false()
    {
        var falseFile = CreateDoc(name: "f.docx", title: "Cleared bool");
        Edit(falseFile, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"x","underline":false}}]""");
        Assert.Equal(UnderlineValues.None, FirstRunUnderline(falseFile)!.Val!.Value);
        Assert.False(GetRun(falseFile)["underline"]!.GetValue<bool>());
        AssertValidatesClean(falseFile);

        var noneFile = CreateDoc(name: "n.docx", title: "Cleared string");
        Edit(noneFile, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"x","underline":"none"}}]""");
        Assert.Equal(UnderlineValues.None, FirstRunUnderline(noneFile)!.Val!.Value);
        Assert.False(GetRun(noneFile)["underline"]!.GetValue<bool>());
        AssertValidatesClean(noneFile);
    }

    // ------------------------------------------------------------ 4. negative

    [Fact]
    public void Unknown_underline_style_is_invalid_args_with_the_candidate_list()
    {
        var file = CreateDoc(title: "Bad underline");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"underline":"squiggle"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.NotNull(ex.Candidates);
        Assert.Contains("double", ex.Candidates!);
        Assert.Contains("wave", ex.Candidates!);
        Assert.Contains("none", ex.Candidates!);
    }

    // -------------------------------------------------- 5. read-side contract (foreign content)

    [Fact]
    public void ForeignSingleUnderline_StillReadsBoolTrue()
    {
        var file = CreateDoc(title: "Foreign single");
        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"imported"}}]""");

        // A single underline written by "someone else" straight into the OOXML.
        InjectFirstRunUnderline(file, UnderlineValues.Single);

        var read = GetRun(file)["underline"]!;
        Assert.Equal(JsonValueKind.True, read.GetValueKind());
        Assert.True(read.GetValue<bool>());
    }

    [Fact]
    public void ForeignDoubleUnderline_ReadsStyleString()
    {
        var file = CreateDoc(title: "Foreign double");
        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"imported"}}]""");

        // A double underline the tool never authored surfaces as its style STRING.
        InjectFirstRunUnderline(file, UnderlineValues.Double);

        Assert.Equal("double", GetRun(file)["underline"]!.GetValue<string>());
    }

    // ------------------------------------------------------------ 6. query selector

    [Fact]
    public void Query_surfaces_the_underline_style_and_matches_the_selector()
    {
        var file = CreateDoc(title: "Queryable underline");
        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"styled","underline":"double"}}]""");

        var matches = Data(Handler.Query(Ctx(file, new JsonObject { ["selector"] = "run[underline=double]" })))["matches"]!.AsArray();
        Assert.NotEmpty(matches);
    }

    [Fact]
    public void Query_underline_true_matches_single_but_not_a_non_single_style()
    {
        // The selector is discriminated with the read: run[underline=true] still matches a
        // (tool-authored or foreign) SINGLE underline — byte-stable — but a foreign DOUBLE
        // underline now surfaces as underline='double', so run[underline=true] does NOT match it
        // (query run[underline=double] to find those). Locks the value-shape change to non-single.
        var single = CreateDoc("single.docx", title: "Single");
        Edit(single, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"a","underline":true}}]""");
        Assert.NotEmpty(Data(Handler.Query(Ctx(single, new JsonObject { ["selector"] = "run[underline=true]" })))["matches"]!.AsArray());

        var dbl = CreateDoc("double.docx", title: "Double");
        Edit(dbl, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"b"}}]""");
        InjectFirstRunUnderline(dbl, UnderlineValues.Double);
        Assert.Empty(Data(Handler.Query(Ctx(dbl, new JsonObject { ["selector"] = "run[underline=true]" })))["matches"]!.AsArray());
        Assert.NotEmpty(Data(Handler.Query(Ctx(dbl, new JsonObject { ["selector"] = "run[underline=double]" })))["matches"]!.AsArray());
    }

    /// <summary>Writes a raw w:u @val into the first run, simulating foreign/non-tool content.</summary>
    private static void InjectFirstRunUnderline(string file, UnderlineValues val)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: true);
        var run = doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().First();
        var rPr = run.RunProperties ??= new RunProperties();
        rPr.Underline = new Underline { Val = val };
        doc.MainDocumentPart!.Document!.Save();
    }
}
