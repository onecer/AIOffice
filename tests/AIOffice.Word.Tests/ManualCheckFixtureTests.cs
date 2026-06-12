using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// We cannot open real Word here, so this test regenerates a representative
/// fixture into fixtures/manual-check/ for a human to open, and holds the
/// OpenXmlValidator at 0 errors as the automated proxy.
/// </summary>
public sealed class ManualCheckFixtureTests : WordTestBase
{
    [Fact]
    public void Regenerate_word_sample_for_human_inspection()
    {
        var repoRoot = FindRepoRoot();
        if (repoRoot is null)
        {
            return; // running outside the repo (e.g. CI artifact stage); nothing to regenerate
        }

        var dir = Path.Combine(repoRoot, "fixtures", "manual-check");
        Directory.CreateDirectory(dir);
        var file = Path.Combine(dir, "word-sample.docx");
        if (File.Exists(file))
        {
            File.Delete(file);
        }

        var workspace = new Workspace(dir);
        var ctx = new CommandContext
        {
            Workspace = workspace,
            File = file,
            Args = new JsonObject { ["title"] = "AIOffice Manual Check" },
        };
        Handler.Create(ctx);

        Handler.Edit(
            new CommandContext { Workspace = workspace, File = file, Args = [] },
            EditOp.ParseBatch("""
                [
                  {"op":"set","path":"/body/p[2]","props":{"text":"Open this file in Word: heading, formatting and table below must look right."}},
                  {"op":"add","path":"/body","props":{"text":"Formatting","style":"Heading2"}},
                  {"op":"add","path":"/body","props":{"text":"Bold red 14pt centered.","bold":true,"color":"CC0000","fontSize":14,"alignment":"center"}},
                  {"op":"add","path":"/body","props":{"text":"Italic and underlined.","italic":true,"underline":true}},
                  {"op":"add","path":"/body","props":{"text":"Table","style":"Heading2"}},
                  {"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":3}},
                  {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"Quarter"}},
                  {"op":"set","path":"/body/table[1]/tr[1]/tc[2]","props":{"text":"Revenue"}},
                  {"op":"set","path":"/body/table[1]/tr[1]/tc[3]","props":{"text":"Growth"}},
                  {"op":"add","path":"/body/table[1]","type":"tr","props":{"cells":["Q3","1.2M","8%"]}},
                  {"op":"add","path":"/header[1]","type":"header","props":{"text":"AIOffice manual check — header","alignment":"center"}},
                  {"op":"add","path":"/footer[1]","type":"footer","props":{"text":"AIOffice manual check — footer"}}
                ]
                """));

        AssertValidatesClean(file);
    }

    private static string? FindRepoRoot()
    {
        for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir is not null; dir = dir.Parent)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AIOffice.sln")))
            {
                return dir.FullName;
            }
        }

        return null;
    }
}
