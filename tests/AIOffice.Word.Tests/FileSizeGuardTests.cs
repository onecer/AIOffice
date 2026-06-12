using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// M2 wiring check: the docx open path consults <see cref="FileSizeGuard"/>,
/// so an over-limit file fails fast with file_too_large instead of being
/// parsed. The oversized fixture is a sparse 51 MB file (cheap on APFS) —
/// the guard must fire on size alone, before any package open.
/// </summary>
public sealed class FileSizeGuardTests : WordTestBase
{
    private string CreateOversizedFile(string name = "huge.docx")
    {
        var path = Path.Combine(Dir, name);
        using var fs = File.Create(path);
        fs.SetLength((FileSizeGuard.DefaultMaxFileMb + 1L) * 1024 * 1024);
        return path;
    }

    [Fact]
    public void Read_on_an_oversized_docx_is_file_too_large()
    {
        var file = CreateOversizedFile();

        var ex = Assert.Throws<AiofficeException>(
            () => Handler.Read(Ctx(file, new JsonObject { ["view"] = "text" })));

        Assert.Equal(ErrorCodes.FileTooLarge, ex.Code);
        Assert.Contains(FileSizeGuard.EnvVar, ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Edit_on_an_oversized_docx_is_file_too_large_and_writes_nothing()
    {
        var file = CreateOversizedFile();
        var lengthBefore = new FileInfo(file).Length;

        var ex = Assert.Throws<AiofficeException>(() => Handler.Edit(
            Ctx(file),
            EditOp.ParseBatch("""[{"op":"set","path":"/body/p[1]","props":{"text":"Hi"}}]""")));

        Assert.Equal(ErrorCodes.FileTooLarge, ex.Code);
        Assert.Equal(lengthBefore, new FileInfo(file).Length);
    }
}
