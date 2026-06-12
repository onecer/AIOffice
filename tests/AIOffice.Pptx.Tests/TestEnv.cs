using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Pptx.Tests;

/// <summary>A throwaway sandboxed workspace per test.</summary>
public sealed class TempWorkspace : IDisposable
{
    public TempWorkspace()
    {
        Dir = Directory.CreateTempSubdirectory("aioffice-pptx-tests-").FullName;
        Workspace = new Workspace(Dir);
    }

    public string Dir { get; }

    public Workspace Workspace { get; }

    /// <summary>Sandbox-resolved absolute path for a file name inside the workspace.</summary>
    public string PathOf(string fileName) => Workspace.Resolve(fileName);

    public CommandContext Ctx(string fileName, params (string Key, JsonNode? Value)[] args)
    {
        var jsonArgs = new JsonObject();
        foreach (var (key, value) in args)
        {
            jsonArgs[key] = value;
        }

        return new CommandContext
        {
            Workspace = Workspace,
            File = Workspace.Resolve(fileName),
            Args = jsonArgs,
        };
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(Dir, recursive: true);
        }
        catch (IOException)
        {
            // Best-effort cleanup of the temp dir.
        }
    }
}

internal static class TestEnv
{
    /// <summary>The envelope data re-serialized to a JsonObject for assertions (camelCase wire shape).</summary>
    public static JsonObject Data(Envelope envelope)
    {
        Assert.NotNull(envelope.Data);
        return JsonSerializer.SerializeToNode(envelope.Data, JsonDefaults.Options)!.AsObject();
    }

    public static JsonObject AssertOk(Envelope envelope)
    {
        Assert.True(envelope.IsOk, envelope.Error is { } e ? $"{e.Code}: {e.Message}" : "envelope not ok");
        return Data(envelope);
    }

    public static ErrorBody AssertFail(Envelope envelope, string expectedCode)
    {
        Assert.False(envelope.IsOk, "expected a failure envelope");
        Assert.NotNull(envelope.Error);
        Assert.Equal(expectedCode, envelope.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(envelope.Error.Suggestion), "error without suggestion");
        return envelope.Error;
    }

    public static EditOp Op(string op, string path, string? type = null, string? position = null, JsonObject? props = null) =>
        new() { Op = op, Path = path, Type = type, Position = position, Props = props };

    public static JsonObject Props(params (string Key, JsonNode? Value)[] pairs)
    {
        var props = new JsonObject();
        foreach (var (key, value) in pairs)
        {
            props[key] = value;
        }

        return props;
    }

    /// <summary>Asserts the OpenXmlValidator reports zero issues via the validate verb.</summary>
    public static void AssertValid(TempWorkspace ws, string fileName)
    {
        var data = AssertOk(new PptxHandler().Validate(ws.Ctx(fileName)));
        Assert.True(
            data["valid"]!.GetValue<bool>(),
            "validator issues: " + data["issues"]!.ToJsonString());
    }
}
