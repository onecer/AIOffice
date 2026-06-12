using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class TemplateTests : WordTestBase
{
    /// <summary>A document whose placeholders are deliberately split across formatted runs.</summary>
    private string CreateSplitRunTemplate(string name = "tpl.docx")
    {
        var file = Path.Combine(Dir, name);
        using (var doc = WordprocessingDocument.Create(file, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(
                    new Run(new Text("Hello {{na")),
                    new Run(new RunProperties(new Bold()), new Text("me}}")),
                    new Run(new Text(", welcome to {{place}}!")))));
        }

        return file;
    }

    [Fact]
    public void Placeholder_split_across_runs_is_merged()
    {
        var file = CreateSplitRunTemplate();

        var envelope = Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonObject { ["name"] = "Ada", ["place"] = "Rome" },
        }));

        Assert.Equal(2, Data(envelope)["replaced"]!.GetValue<int>());
        Assert.Equal("Hello Ada, welcome to Rome!", BodyTexts(file)[0]);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Formatting_of_split_runs_survives_the_merge()
    {
        var file = CreateSplitRunTemplate();

        Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonObject { ["name"] = "Ada", ["place"] = "Rome" },
        }));

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var runs = doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().ToList();
        Assert.Equal(3, runs.Count);
        // The replacement lands where the placeholder started; the bold run keeps its rPr.
        Assert.Equal("Hello Ada", runs[0].InnerText);
        Assert.NotNull(runs[1].RunProperties?.Bold);
        Assert.Equal(string.Empty, runs[1].InnerText);
        Assert.Equal(", welcome to Rome!", runs[2].InnerText);
    }

    [Fact]
    public void Unresolved_keys_stay_verbatim_and_are_reported()
    {
        var file = CreateSplitRunTemplate();

        var envelope = Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonObject { ["name"] = "Ada" },
        }));

        var data = Data(envelope);
        Assert.Equal(new[] { "place" }, data["unresolved"]!.AsArray().Select(n => n!.GetValue<string>()));
        Assert.Contains("{{place}}", BodyTexts(file)[0], StringComparison.Ordinal);
        Assert.Contains(
            JsonNode.Parse(envelope.ToJson())!["meta"]!["warnings"]!.AsArray(),
            w => w!["code"]!.GetValue<string>() == "unresolved_keys");
    }

    [Fact]
    public void Merging_to_an_output_file_leaves_the_source_untouched()
    {
        var file = CreateSplitRunTemplate();
        var before = File.ReadAllBytes(file);

        var envelope = Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonObject { ["name"] = "Ada", ["place"] = "Rome" },
            ["output"] = "merged.docx",
        }));

        Assert.Equal(before, File.ReadAllBytes(file));
        var merged = Path.Combine(Dir, "merged.docx");
        Assert.True(File.Exists(merged));
        Assert.Equal("Hello Ada, welcome to Rome!", BodyTexts(merged)[0]);
        Assert.EndsWith("merged.docx", Data(envelope)["written"]!.GetValue<string>(), StringComparison.Ordinal);
    }

    [Fact]
    public void In_place_merge_snapshots_the_pre_image()
    {
        var file = CreateSplitRunTemplate();
        var preRev = Rev.OfFile(file);

        Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonObject { ["name"] = "Ada", ["place"] = "Rome" },
        }));

        var entry = Assert.Single(Snapshots.List(file));
        Assert.Equal(preRev, entry.Rev);
    }

    [Fact]
    public void Missing_data_is_invalid_args()
    {
        var file = CreateSplitRunTemplate();

        var ex = Assert.Throws<AiofficeException>(() => Handler.Template(Ctx(file)));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Data_passed_as_a_json_string_is_accepted()
    {
        var file = CreateSplitRunTemplate();

        var envelope = Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = """{"name":"Grace","place":"Paris"}""",
        }));

        Assert.True(envelope.IsOk);
        Assert.Equal("Hello Grace, welcome to Paris!", BodyTexts(file)[0]);
    }
}
