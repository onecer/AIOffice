using System.Text.Json;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Core.Tests;

public class EnvelopeTests
{
    [Fact]
    public void Success_envelope_has_ok_data_and_meta_but_no_error()
    {
        var envelope = Envelope.Ok(new { Count = 3 }, new Meta { File = "a.docx", Rev = "abc123def456" });

        using var doc = JsonDocument.Parse(envelope.ToJson());
        var root = doc.RootElement;

        Assert.True(root.GetProperty("ok").GetBoolean());
        Assert.Equal(3, root.GetProperty("data").GetProperty("count").GetInt32());
        Assert.False(root.TryGetProperty("error", out _)); // nulls are omitted
        Assert.Equal("a.docx", root.GetProperty("meta").GetProperty("file").GetString());
        Assert.Equal("abc123def456", root.GetProperty("meta").GetProperty("rev").GetString());
        Assert.Equal(Meta.ToolVersion, root.GetProperty("meta").GetProperty("version").GetString());
        Assert.True(root.GetProperty("meta").TryGetProperty("elapsedMs", out _));
    }

    [Fact]
    public void Failure_envelope_carries_code_message_suggestion()
    {
        var envelope = Envelope.Fail(
            ErrorCodes.InvalidPath,
            "No paragraph at /body/p[99].",
            "The document has 4 paragraphs; try /body/p[4].",
            candidates: ["/body/p[4]"]);

        using var doc = JsonDocument.Parse(envelope.ToJson());
        var error = doc.RootElement.GetProperty("error");

        Assert.False(doc.RootElement.GetProperty("ok").GetBoolean());
        Assert.False(doc.RootElement.TryGetProperty("data", out _));
        Assert.Equal("invalid_path", error.GetProperty("code").GetString());
        Assert.Equal("/body/p[4]", error.GetProperty("candidates")[0].GetString());
        Assert.False(string.IsNullOrEmpty(error.GetProperty("suggestion").GetString()));
    }

    [Fact]
    public void Fail_requires_a_suggestion()
    {
        Assert.ThrowsAny<ArgumentException>(
            () => Envelope.Fail(ErrorCodes.InternalError, "boom", suggestion: ""));
    }

    [Fact]
    public void AiofficeException_requires_a_suggestion()
    {
        Assert.ThrowsAny<ArgumentException>(
            () => new AiofficeException(ErrorCodes.InvalidArgs, "bad", suggestion: " "));
    }

    [Fact]
    public void FromException_maps_aioffice_exception_fields()
    {
        var ex = new AiofficeException(
            ErrorCodes.SandboxDenied, "escape", "stay inside", candidates: ["a", "b"]);

        var envelope = Envelope.FromException(ex);

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.SandboxDenied, envelope.Error!.Code);
        Assert.Equal(["a", "b"], envelope.Error.Candidates);
        Assert.Equal(ExitCodes.SandboxDenied, envelope.ExitCode);
    }

    [Fact]
    public void FromException_wraps_unknown_exceptions_as_internal_error()
    {
        var envelope = Envelope.FromException(new InvalidOperationException("oops"));

        Assert.Equal(ErrorCodes.InternalError, envelope.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(envelope.Error.Suggestion));
        Assert.Equal(ExitCodes.InternalError, envelope.ExitCode);
    }

    [Fact]
    public void Warnings_serialize_under_meta()
    {
        var envelope = Envelope.Ok(
            data: null,
            new Meta { Warnings = [new Warning(ErrorCodes.FormulaNotEvaluated, "cached value used")] });

        using var doc = JsonDocument.Parse(envelope.ToJson());
        var warning = doc.RootElement.GetProperty("meta").GetProperty("warnings")[0];

        Assert.Equal("formula_not_evaluated", warning.GetProperty("code").GetString());
    }

    [Theory]
    [InlineData(ErrorCodes.InvalidArgs, ExitCodes.UserError)]
    [InlineData(ErrorCodes.FileNotFound, ExitCodes.UserError)]
    [InlineData(ErrorCodes.InvalidPath, ExitCodes.UserError)]
    [InlineData(ErrorCodes.StaleAddress, ExitCodes.UserError)]
    [InlineData(ErrorCodes.SandboxDenied, ExitCodes.SandboxDenied)]
    [InlineData(ErrorCodes.UnsupportedFeature, ExitCodes.UnsupportedFeature)]
    [InlineData(ErrorCodes.FormatCorrupt, ExitCodes.InternalError)]
    [InlineData(ErrorCodes.InternalError, ExitCodes.InternalError)]
    public void Error_codes_map_to_documented_exit_codes(string code, int exitCode)
    {
        Assert.Equal(exitCode, ExitCodes.ForErrorCode(code));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Wire_shape_has_only_ok_data_error_meta(bool ok)
    {
        var envelope = ok
            ? Envelope.Ok(new { X = 1 })
            : Envelope.Fail(ErrorCodes.InvalidArgs, "bad", "fix it");

        using var doc = JsonDocument.Parse(envelope.ToJson());
        var keys = doc.RootElement.EnumerateObject().Select(p => p.Name).ToArray();

        Assert.Subset(new HashSet<string> { "ok", "data", "error", "meta" }, new HashSet<string>(keys));
        Assert.Contains("ok", keys);
        Assert.Contains("meta", keys);
    }

    [Fact]
    public void Rev_is_first_12_hex_of_sha256()
    {
        // sha256("abc") = ba7816bf8f01cfea414140de5dae2223...
        Assert.Equal("ba7816bf8f01", Rev.OfString("abc"));
        Assert.Equal(12, Rev.OfBytes([1, 2, 3]).Length);
    }
}
