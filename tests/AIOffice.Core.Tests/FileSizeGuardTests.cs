using AIOffice.Core;
using Xunit;

namespace AIOffice.Core.Tests;

/// <summary>
/// The file-size guard. Originally written for the M2 50 MB default; the M3
/// directive (功能第一 — features first) flipped the default to UNLIMITED, so
/// these tests now assert: no env var means no limit at all, and
/// AIOFFICE_MAX_FILE_MB is a purely opt-in cap. The suggestion must still
/// name the env var so the agent can adjust the cap deliberately.
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
        Assert.Contains("split", ex.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void File_exactly_at_the_limit_passes()
    {
        var path = WriteFile(1024 * 1024); // exactly 1 MB
        FileSizeGuard.Ensure(path, maxFileMb: 1); // no throw
    }

    [Fact]
    public void Env_var_is_an_opt_in_cap_and_the_default_is_unlimited()
    {
        // Rewritten for the M3 directive: the 50 MB default is gone; unset or
        // unparsable env means null (no limit), a parsable value is the cap.
        var previous = Environment.GetEnvironmentVariable(FileSizeGuard.EnvVar);
        try
        {
            Environment.SetEnvironmentVariable(FileSizeGuard.EnvVar, "7");
            Assert.Equal(7, FileSizeGuard.MaxFileMb);

            Environment.SetEnvironmentVariable(FileSizeGuard.EnvVar, "not-a-number");
            Assert.Null(FileSizeGuard.MaxFileMb);

            Environment.SetEnvironmentVariable(FileSizeGuard.EnvVar, null);
            Assert.Null(FileSizeGuard.MaxFileMb);
        }
        finally
        {
            Environment.SetEnvironmentVariable(FileSizeGuard.EnvVar, previous);
        }
    }

    [Fact]
    public void Without_an_env_cap_even_huge_files_pass()
    {
        // M3 directive (功能第一): no env var -> Ensure is a no-op, whatever the size.
        var previous = Environment.GetEnvironmentVariable(FileSizeGuard.EnvVar);
        try
        {
            Environment.SetEnvironmentVariable(FileSizeGuard.EnvVar, null);
            var path = Path.Combine(_dir, "big.docx");
            using (var fs = File.Create(path))
            {
                fs.SetLength(200L * 1024 * 1024); // sparse 200 MB
            }

            FileSizeGuard.Ensure(path); // no throw
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
