using AIOffice.Core;
using Xunit;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// M2 wiring check: the pptx open path consults <see cref="FileSizeGuard"/>,
/// so an over-limit deck fails fast with file_too_large (sparse 51 MB fixture
/// — the guard fires on size alone, before the package is opened).
/// </summary>
public sealed class FileSizeGuardTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void CreateOversizedFile(string name)
    {
        using var fs = File.Create(Path.Combine(_ws.Dir, name));
        fs.SetLength((FileSizeGuard.DefaultMaxFileMb + 1L) * 1024 * 1024);
    }

    [Fact]
    public void Read_on_an_oversized_pptx_is_file_too_large()
    {
        CreateOversizedFile("huge.pptx");

        var error = TestEnv.AssertFail(
            _handler.Read(_ws.Ctx("huge.pptx")),
            ErrorCodes.FileTooLarge);

        Assert.Contains(FileSizeGuard.EnvVar, error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Edit_on_an_oversized_pptx_is_file_too_large_and_writes_nothing()
    {
        CreateOversizedFile("huge.pptx");
        var lengthBefore = new FileInfo(Path.Combine(_ws.Dir, "huge.pptx")).Length;

        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("huge.pptx"), [TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("background", "112233")))]),
            ErrorCodes.FileTooLarge);

        Assert.Equal(lengthBefore, new FileInfo(Path.Combine(_ws.Dir, "huge.pptx")).Length);
    }
}
