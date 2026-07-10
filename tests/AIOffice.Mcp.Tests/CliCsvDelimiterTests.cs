using System.Text.Json.Nodes;
using AIOffice.Cli;
using AIOffice.Core;
using AIOffice.Core.Cli;
using Xunit;

namespace AIOffice.Mcp.Tests;

/// <summary>
/// Guards the CLI arg-to-ctx plumbing for the 1.23 csv export delimiter: the
/// <c>read</c> verb must FORWARD <c>--delimiter</c> to the handler. The
/// handler-level CsvBridgeTests could not catch this — they inject the ctx
/// directly, so a missing <c>FileVerbs.Read</c> forward stayed invisible until
/// an end-to-end smoke. This drives the real <see cref="FileVerbs"/> path.
/// </summary>
public sealed class CliCsvDelimiterTests : IDisposable
{
    private readonly string _root;
    private readonly FileVerbs _verbs;

    public CliCsvDelimiterTests()
    {
        _root = Path.Combine(Path.GetTempPath(), "aioffice-cli-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_root);
        var snapshots = new SnapshotStore();
        // Fully qualified: AIOffice.Mcp also declares a HandlerDiscovery, and this
        // test lives under the AIOffice.Mcp.Tests namespace.
        _verbs = new FileVerbs(new Workspace(_root), AIOffice.Cli.HandlerDiscovery.Discover(snapshots), snapshots);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_root, recursive: true);
        }
        catch (IOException)
        {
            // Best-effort cleanup.
        }
    }

    private static string Content(Envelope env)
    {
        Assert.True(env.IsOk, env.ToJson());
        return JsonNode.Parse(env.ToJson())!["data"]!["content"]!.GetValue<string>();
    }

    private string ReadCsv(params string[] extra)
    {
        string[] argv = ["read", "m.xlsx", "--view", "csv", "--range", "A1:B2", .. extra];
        return Content(_verbs.Read(ArgParser.Parse(argv)));
    }

    [Fact]
    public void ReadCsv_forwards_the_delimiter_option_to_the_handler()
    {
        Assert.True(_verbs.Create(ArgParser.Parse(["create", "m.xlsx"])).IsOk);
        Assert.True(_verbs.Edit(ArgParser.Parse([
            "edit", "m.xlsx", "--ops",
            """[{"op":"set","path":"/Sheet1/A1","props":{"values":[["a","b,c"],["1","2"]]}}]""",
        ])).IsOk);

        // Default: comma, and the comma-bearing field is quoted (RFC 4180).
        Assert.Equal("a,\"b,c\"\r\n1,2\r\n", ReadCsv());
        // tab: the field carrying a comma is NOT quoted — quoting is delimiter-aware.
        Assert.Equal("a\tb,c\r\n1\t2\r\n", ReadCsv("--delimiter", "tab"));
        Assert.Equal("a;b,c\r\n1;2\r\n", ReadCsv("--delimiter", "semicolon"));

        // The delimiter vocabulary is not widened: pipe is rejected via the CLI too.
        var pipe = _verbs.Read(ArgParser.Parse(["read", "m.xlsx", "--view", "csv", "--delimiter", "pipe"]));
        Assert.False(pipe.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, pipe.Error!.Code);
    }
}
