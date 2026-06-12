using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// M2 wiring check: the xlsx open path consults <see cref="FileSizeGuard"/>,
/// so an over-limit workbook fails fast with file_too_large (sparse 51 MB
/// fixture — the guard fires on size alone, before ClosedXML opens anything).
/// </summary>
public sealed class FileSizeGuardTests : ExcelTestBase
{
    private string CreateOversizedFile(string name = "huge.xlsx")
    {
        var path = Path.Combine(Dir, name);
        using var fs = File.Create(path);
        fs.SetLength((FileSizeGuard.DefaultMaxFileMb + 1L) * 1024 * 1024);
        return path;
    }

    [Fact]
    public void Read_on_an_oversized_xlsx_is_file_too_large()
    {
        var file = CreateOversizedFile();

        var envelope = Handler.Read(Ctx(file, ("view", "stats")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.FileTooLarge, envelope.Error!.Code);
        Assert.Contains(FileSizeGuard.EnvVar, envelope.Error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Edit_on_an_oversized_xlsx_is_file_too_large_and_writes_nothing()
    {
        var file = CreateOversizedFile();
        var lengthBefore = new FileInfo(file).Length;

        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("value", "x")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.FileTooLarge, envelope.Error!.Code);
        Assert.Equal(lengthBefore, new FileInfo(file).Length);
    }
}
