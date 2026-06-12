using AIOffice.Core;
using Xunit;

namespace AIOffice.Core.Tests;

public sealed class WorkspaceTests : IDisposable
{
    private readonly string _root;
    private readonly string _outside;

    public WorkspaceTests()
    {
        var stem = Path.Combine(Path.GetTempPath(), "aioffice-tests-" + Guid.NewGuid().ToString("N"));
        _root = Path.Combine(stem, "workspace");
        _outside = Path.Combine(stem, "outside");
        Directory.CreateDirectory(_root);
        Directory.CreateDirectory(_outside);
    }

    public void Dispose()
    {
        var stem = Path.GetDirectoryName(_root)!;
        if (Directory.Exists(stem))
        {
            Directory.Delete(stem, recursive: true);
        }
    }

    [Fact]
    public void Resolves_relative_path_inside_root()
    {
        var workspace = new Workspace(_root);

        var resolved = workspace.Resolve("report.docx");

        Assert.StartsWith(workspace.Root, resolved, StringComparison.Ordinal);
        Assert.EndsWith("report.docx", resolved, StringComparison.Ordinal);
    }

    [Fact]
    public void Resolves_nested_path_that_does_not_exist_yet()
    {
        var workspace = new Workspace(_root);

        var resolved = workspace.Resolve(Path.Combine("sub", "dir", "new.xlsx"));

        Assert.StartsWith(workspace.Root, resolved, StringComparison.Ordinal);
    }

    [Fact]
    public void Denies_dotdot_escape()
    {
        var workspace = new Workspace(_root);

        var ex = Assert.Throws<AiofficeException>(
            () => workspace.Resolve(Path.Combine("..", "outside", "secret.docx")));

        Assert.Equal(ErrorCodes.SandboxDenied, ex.Code);
        Assert.False(string.IsNullOrWhiteSpace(ex.Suggestion));
    }

    [Fact]
    public void Denies_dotdot_escape_that_reenters_nowhere()
    {
        var workspace = new Workspace(_root);

        var ex = Assert.Throws<AiofficeException>(
            () => workspace.Resolve("../../../../../../etc/passwd"));

        Assert.Equal(ErrorCodes.SandboxDenied, ex.Code);
    }

    [Fact]
    public void Allows_dotdot_that_stays_inside()
    {
        var workspace = new Workspace(_root);
        Directory.CreateDirectory(Path.Combine(_root, "a"));

        var resolved = workspace.Resolve(Path.Combine("a", "..", "file.docx"));

        Assert.Equal(Path.Combine(workspace.Root, "file.docx"), resolved);
    }

    [Fact]
    public void Denies_absolute_path_outside_root()
    {
        var workspace = new Workspace(_root);

        var ex = Assert.Throws<AiofficeException>(
            () => workspace.Resolve(Path.Combine(_outside, "x.docx")));

        Assert.Equal(ErrorCodes.SandboxDenied, ex.Code);
    }

    [Fact]
    public void Denies_symlinked_file_pointing_outside()
    {
        var target = Path.Combine(_outside, "real.docx");
        File.WriteAllText(target, "outside");
        var link = Path.Combine(_root, "sneaky.docx");
        File.CreateSymbolicLink(link, target);

        var workspace = new Workspace(_root);
        var ex = Assert.Throws<AiofficeException>(() => workspace.Resolve("sneaky.docx"));

        Assert.Equal(ErrorCodes.SandboxDenied, ex.Code);
    }

    [Fact]
    public void Denies_symlinked_directory_pointing_outside()
    {
        var link = Path.Combine(_root, "vault");
        Directory.CreateSymbolicLink(link, _outside);

        var workspace = new Workspace(_root);
        var ex = Assert.Throws<AiofficeException>(
            () => workspace.Resolve(Path.Combine("vault", "doc.docx")));

        Assert.Equal(ErrorCodes.SandboxDenied, ex.Code);
    }

    [Fact]
    public void Allows_symlink_that_stays_inside()
    {
        var target = Path.Combine(_root, "real.docx");
        File.WriteAllText(target, "inside");
        var link = Path.Combine(_root, "alias.docx");
        File.CreateSymbolicLink(link, target);

        var workspace = new Workspace(_root);
        var resolved = workspace.Resolve("alias.docx");

        Assert.Equal(Path.Combine(workspace.Root, "real.docx"), resolved);
    }

    [Fact]
    public void MustExist_throws_file_not_found()
    {
        var workspace = new Workspace(_root);

        var ex = Assert.Throws<AiofficeException>(
            () => workspace.Resolve("missing.docx", mustExist: true));

        Assert.Equal(ErrorCodes.FileNotFound, ex.Code);
    }

    [Fact]
    public void Second_root_is_also_allowed()
    {
        var workspace = new Workspace(_root, _outside);

        var resolved = workspace.Resolve(Path.Combine(_outside, "ok.docx"));

        Assert.StartsWith(workspace.Roots[1], resolved, StringComparison.Ordinal);
    }

    [Fact]
    public void Missing_root_is_invalid_args()
    {
        var ex = Assert.Throws<AiofficeException>(
            () => new Workspace(Path.Combine(_root, "does-not-exist")));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }
}
