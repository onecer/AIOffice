using AIOffice.Core;
using Xunit;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// Guard wiring check: the pptx open path consults <see cref="FileSizeGuard"/>,
/// so an over-cap deck fails fast with file_too_large (sparse 51 MB fixture —
/// the guard fires on size alone, before the package is opened). M3 made the
/// cap OPT-IN (default unlimited, per the 功能第一 directive), so these tests
/// pin <c>AIOFFICE_MAX_FILE_MB=50</c> explicitly; without the env var the
/// guard must let the same file through.
/// </summary>
public sealed class FileSizeGuardTests : IDisposable
{
    private const int LimitMb = 50;

    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private sealed class EnvScope : IDisposable
    {
        private readonly string _key;
        private readonly string? _previous;

        public EnvScope(string key, string? value)
        {
            _key = key;
            _previous = Environment.GetEnvironmentVariable(key);
            Environment.SetEnvironmentVariable(key, value);
        }

        public void Dispose() => Environment.SetEnvironmentVariable(_key, _previous);
    }

    private void CreateOversizedFile(string name)
    {
        using var fs = File.Create(Path.Combine(_ws.Dir, name));
        fs.SetLength((LimitMb + 1L) * 1024 * 1024);
    }

    [Fact]
    public void Read_over_an_explicit_cap_is_file_too_large()
    {
        using var cap = new EnvScope(FileSizeGuard.EnvVar, "50");
        CreateOversizedFile("huge.pptx");

        var error = TestEnv.AssertFail(
            _handler.Read(_ws.Ctx("huge.pptx")),
            ErrorCodes.FileTooLarge);

        Assert.Contains(FileSizeGuard.EnvVar, error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Edit_over_an_explicit_cap_is_file_too_large_and_writes_nothing()
    {
        using var cap = new EnvScope(FileSizeGuard.EnvVar, "50");
        CreateOversizedFile("huge.pptx");
        var lengthBefore = new FileInfo(Path.Combine(_ws.Dir, "huge.pptx")).Length;

        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("huge.pptx"), [TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("background", "112233")))]),
            ErrorCodes.FileTooLarge);

        Assert.Equal(lengthBefore, new FileInfo(Path.Combine(_ws.Dir, "huge.pptx")).Length);
    }

    [Fact]
    public void Without_the_env_cap_an_oversized_file_passes_the_guard()
    {
        // M3 default = unlimited: the guard lets the file through; the honest
        // failure for a zero-filled sparse fixture is format_corrupt — never
        // file_too_large.
        using var cap = new EnvScope(FileSizeGuard.EnvVar, null);
        CreateOversizedFile("huge.pptx");

        var envelope = _handler.Read(_ws.Ctx("huge.pptx"));

        Assert.False(envelope.IsOk);
        Assert.NotEqual(ErrorCodes.FileTooLarge, envelope.Error!.Code);
    }
}
