using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>Temp-dir workspace + isolated snapshot ring around one WordHandler.</summary>
public abstract class WordTestBase : IDisposable
{
    protected WordTestBase()
    {
        var raw = Path.Combine(Path.GetTempPath(), "aioffice-word-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(raw);
        Workspace = new Workspace(raw);
        Dir = Workspace.Root; // canonical (symlinks resolved), so paths match handler output
        Snapshots = new SnapshotStore(Path.Combine(Dir, ".snapshots"));
        Handler = new WordHandler(Snapshots);
    }

    protected string Dir { get; }

    protected Workspace Workspace { get; }

    protected SnapshotStore Snapshots { get; }

    protected WordHandler Handler { get; }

    public void Dispose()
    {
        if (Directory.Exists(Dir))
        {
            Directory.Delete(Dir, recursive: true);
        }

        GC.SuppressFinalize(this);
    }

    protected CommandContext Ctx(string file, JsonObject? args = null) => new()
    {
        Workspace = Workspace,
        File = file,
        Args = args ?? [],
    };

    /// <summary>Creates a docx in the workspace and returns its absolute path.</summary>
    protected string CreateDoc(string name = "doc.docx", string? title = null)
    {
        var args = new JsonObject();
        if (title is not null)
        {
            args["title"] = title;
        }

        var envelope = Handler.Create(Ctx(name, args));
        Assert.True(envelope.IsOk, envelope.ToJson());
        var path = Path.Combine(Dir, name);
        Assert.True(File.Exists(path));
        return path;
    }

    protected Envelope Edit(string file, string opsJson, JsonObject? args = null) =>
        Handler.Edit(Ctx(file, args), EditOp.ParseBatch(opsJson));

    /// <summary>Envelope data through the JSON wire shape (what agents actually see).</summary>
    protected static JsonNode Data(Envelope envelope)
    {
        Assert.True(envelope.IsOk, envelope.ToJson());
        return JsonNode.Parse(envelope.ToJson())!["data"]!;
    }

    /// <summary>The validator oracle: after any mutation the file must report 0 errors.</summary>
    protected static void AssertValidatesClean(string file)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var errors = new OpenXmlValidator().Validate(doc).ToList();
        Assert.True(
            errors.Count == 0,
            "OpenXmlValidator errors:\n" + string.Join('\n', errors.Select(e => $"{e.Id}: {e.Description}")));
    }

    /// <summary>All paragraph texts of the body, via the handler's text view.</summary>
    protected IReadOnlyList<string> BodyTexts(string file)
    {
        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "text" })));
        return [.. data["lines"]!.AsArray().Select(l => l!["text"]!.GetValue<string>())];
    }
}
