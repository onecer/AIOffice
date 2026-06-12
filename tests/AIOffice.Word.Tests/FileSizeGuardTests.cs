using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// Guard wiring check: the docx open path consults <see cref="FileSizeGuard"/>,
/// so an over-cap file fails fast with file_too_large instead of being parsed.
/// The oversized fixture is a sparse 51 MB file (cheap on APFS) — the guard
/// must fire on size alone, before any package open. M3 made the cap OPT-IN
/// (default unlimited, per the 功能第一 directive), so these tests pin
/// <c>AIOFFICE_MAX_FILE_MB=50</c> explicitly; without the env var the same
/// file must open-attempt normally.
/// </summary>
public sealed class FileSizeGuardTests : WordTestBase
{
    private const int LimitMb = 50;

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

    private string CreateOversizedFile(string name = "huge.docx")
    {
        var path = Path.Combine(Dir, name);
        using var fs = File.Create(path);
        fs.SetLength((LimitMb + 1L) * 1024 * 1024);
        return path;
    }

    [Fact]
    public void Read_over_an_explicit_cap_is_file_too_large()
    {
        using var cap = new EnvScope(FileSizeGuard.EnvVar, "50");
        var file = CreateOversizedFile();

        var ex = Assert.Throws<AiofficeException>(
            () => Handler.Read(Ctx(file, new JsonObject { ["view"] = "text" })));

        Assert.Equal(ErrorCodes.FileTooLarge, ex.Code);
        Assert.Contains(FileSizeGuard.EnvVar, ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Edit_over_an_explicit_cap_is_file_too_large_and_writes_nothing()
    {
        using var cap = new EnvScope(FileSizeGuard.EnvVar, "50");
        var file = CreateOversizedFile();
        var lengthBefore = new FileInfo(file).Length;

        var ex = Assert.Throws<AiofficeException>(() => Handler.Edit(
            Ctx(file),
            EditOp.ParseBatch("""[{"op":"set","path":"/body/p[1]","props":{"text":"Hi"}}]""")));

        Assert.Equal(ErrorCodes.FileTooLarge, ex.Code);
        Assert.Equal(lengthBefore, new FileInfo(file).Length);
    }

    [Fact]
    public void Without_the_env_cap_an_oversized_file_passes_the_guard()
    {
        // M3 default = unlimited: the guard lets the file through, so the
        // failure (if any) is the honest format_corrupt from the zero-byte
        // sparse payload — NOT file_too_large.
        using var cap = new EnvScope(FileSizeGuard.EnvVar, null);
        var file = CreateOversizedFile();

        var ex = Record.Exception(() => Handler.Read(Ctx(file, new JsonObject { ["view"] = "text" })));

        var aio = Assert.IsType<AiofficeException>(ex);
        Assert.NotEqual(ErrorCodes.FileTooLarge, aio.Code);
    }
}
