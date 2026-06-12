using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// Guard wiring check: the xlsx open path consults <see cref="FileSizeGuard"/>,
/// so an over-limit workbook fails fast with file_too_large (sparse fixture —
/// the guard fires on size alone, before anything opens the package). M3 made
/// the limit OPT-IN, so these tests pin <c>AIOFFICE_MAX_FILE_MB=50</c>
/// explicitly: they stay green with the old 50 MB default and with the new
/// unlimited default alike (the env override has always won).
/// </summary>
[Collection("FileSizeEnv")]
public sealed class FileSizeGuardTests : ExcelTestBase
{
    private const int LimitMb = 50;

    private string CreateOversizedFile(string name = "huge.xlsx")
    {
        var path = Path.Combine(Dir, name);
        using var fs = File.Create(path);
        fs.SetLength((LimitMb + 1L) * 1024 * 1024);
        return path;
    }

    [Fact]
    public void Read_over_an_explicit_limit_is_file_too_large()
    {
        using var limit = new EnvScope(FileSizeGuard.EnvVar, "50");
        var file = CreateOversizedFile();

        var envelope = Handler.Read(Ctx(file, ("view", "stats")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.FileTooLarge, envelope.Error!.Code);
        Assert.Contains(FileSizeGuard.EnvVar, envelope.Error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Edit_over_an_explicit_limit_is_file_too_large_and_writes_nothing()
    {
        using var limit = new EnvScope(FileSizeGuard.EnvVar, "50");
        var file = CreateOversizedFile();
        var lengthBefore = new FileInfo(file).Length;

        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("value", "x")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.FileTooLarge, envelope.Error!.Code);
        Assert.Equal(lengthBefore, new FileInfo(file).Length);
    }
}
