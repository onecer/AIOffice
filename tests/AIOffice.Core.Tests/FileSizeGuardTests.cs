using AIOffice.Core;
using Xunit;

namespace AIOffice.Core.Tests;

/// <summary>
/// The M2 file-size guard: files over the limit (default 50 MB, overridable
/// via AIOFFICE_MAX_FILE_MB) fail fast with file_too_large; the suggestion
/// must name the env var so the agent can raise the limit deliberately.
/// </summary>
public sealed class FileSizeGuardTests : IDisposable
{
    private readonly string _dir = Directory.CreateTempSubdirectory("aio-guard-").FullName;

    public void Dispose() => Directory.Delete(_dir, recursive: true);

    private string WriteFile(int bytes)
    {
        var path = Path.Combine(_dir, "doc.docx");
        File.WriteAllBytes(path, new byte[bytes]);
        return path;
    }

    [Fact]
    public void File_under_the_limit_passes()
    {
        var path = WriteFile(1024);
        FileSizeGuard.Ensure(path, maxFileMb: 1); // no throw
    }

    [Fact]
    public void Missing_file_passes_so_not_found_handling_stays_authoritative()
    {
        FileSizeGuard.Ensure(Path.Combine(_dir, "nope.docx"), maxFileMb: 0); // no throw
    }

    [Fact]
    public void File_over_the_limit_is_file_too_large_with_env_var_suggestion()
    {
        var path = WriteFile(1024);

        var ex = Assert.Throws<AiofficeException>(() => FileSizeGuard.Ensure(path, maxFileMb: 0));

        Assert.Equal(ErrorCodes.FileTooLarge, ex.Code);
        Assert.Contains("0 MB limit", ex.Message, StringComparison.Ordinal);
        Assert.Contains(FileSizeGuard.EnvVar, ex.Suggestion, StringComparison.Ordinal);
        Assert.Contains("Split", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void File_exactly_at_the_limit_passes()
    {
        var path = WriteFile(1024 * 1024); // exactly 1 MB
        FileSizeGuard.Ensure(path, maxFileMb: 1); // no throw
    }

    [Fact]
    public void Env_var_overrides_the_default_limit()
    {
        var previous = Environment.GetEnvironmentVariable(FileSizeGuard.EnvVar);
        try
        {
            Environment.SetEnvironmentVariable(FileSizeGuard.EnvVar, "7");
            Assert.Equal(7, FileSizeGuard.MaxFileMb);

            Environment.SetEnvironmentVariable(FileSizeGuard.EnvVar, "not-a-number");
            Assert.Equal(FileSizeGuard.DefaultMaxFileMb, FileSizeGuard.MaxFileMb);

            Environment.SetEnvironmentVariable(FileSizeGuard.EnvVar, null);
            Assert.Equal(FileSizeGuard.DefaultMaxFileMb, FileSizeGuard.MaxFileMb);
        }
        finally
        {
            Environment.SetEnvironmentVariable(FileSizeGuard.EnvVar, previous);
        }
    }

    [Fact]
    public void File_too_large_is_a_registered_error_code_with_user_error_exit()
    {
        Assert.Contains(ErrorCodes.FileTooLarge, ErrorCodes.All);
        Assert.Equal(ExitCodes.UserError, ExitCodes.ForErrorCode(ErrorCodes.FileTooLarge));
    }
}
