namespace AIOffice.Core;

/// <summary>
/// The filesystem sandbox. Every user-supplied path must resolve — after
/// normalizing <c>..</c> and following symlinks — to a location inside one of
/// the allowed roots, otherwise <c>sandbox_denied</c> is thrown. Nothing in
/// aioffice touches the filesystem except through a Workspace.
/// </summary>
public sealed class Workspace
{
    private readonly string[] _roots;

    /// <param name="roots">Allowlisted root directories. They must exist; each is canonicalized (symlinks resolved).</param>
    public Workspace(params string[] roots)
    {
        if (roots.Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "Workspace requires at least one root directory.",
                "Pass --workspace <dir> or set AIOFFICE_WORKSPACE; the default is the current directory.");
        }

        _roots = new string[roots.Length];
        for (var i = 0; i < roots.Length; i++)
        {
            var full = Path.GetFullPath(roots[i]);
            if (!Directory.Exists(full))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Workspace root does not exist or is not a directory: {roots[i]}",
                    "Create the directory first, or point --workspace at an existing directory.");
            }

            _roots[i] = RealPath(full);
        }
    }

    /// <summary>The primary root (first allowlisted directory), canonicalized.</summary>
    public string Root => _roots[0];

    /// <summary>All canonicalized roots.</summary>
    public IReadOnlyList<string> Roots => _roots;

    /// <summary>
    /// Resolves a user-supplied path (relative paths are anchored at the primary
    /// root) to a canonical absolute path, then enforces the sandbox.
    /// </summary>
    /// <exception cref="AiofficeException">sandbox_denied when the real path escapes all roots; file_not_found when <paramref name="mustExist"/> is set and the file is missing.</exception>
    public string Resolve(string userPath, bool mustExist = false)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(userPath);

        var combined = Path.IsPathRooted(userPath) ? userPath : Path.Combine(Root, userPath);
        var normalized = Path.GetFullPath(combined); // lexical: collapses . and ..
        var real = RealPath(normalized); // semantic: follows symlinks on every existing component

        if (!IsInsideAnyRoot(real))
        {
            throw new AiofficeException(
                ErrorCodes.SandboxDenied,
                $"Path escapes the workspace sandbox: {userPath}",
                $"Use a path inside the workspace ({Root}), or widen the sandbox with --workspace.");
        }

        if (mustExist && !File.Exists(real) && !Directory.Exists(real))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"File not found: {userPath}",
                "Check the path spelling, or run 'aioffice create' to make a new document.");
        }

        return real;
    }

    /// <summary>True when <paramref name="path"/> (already canonical) is inside a root.</summary>
    public bool Contains(string path) => IsInsideAnyRoot(RealPath(Path.GetFullPath(path)));

    private bool IsInsideAnyRoot(string realPath)
    {
        foreach (var root in _roots)
        {
            if (realPath.Equals(root, StringComparison.Ordinal) ||
                realPath.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Canonicalizes a full path by following symlinks component-by-component.
    /// Non-existent trailing components are kept verbatim (so paths about to be
    /// created are still checked against the sandbox).
    /// </summary>
    internal static string RealPath(string fullPath)
    {
        var root = Path.GetPathRoot(fullPath) ?? Path.DirectorySeparatorChar.ToString();
        var rest = fullPath[root.Length..];
        var current = root;

        foreach (var segment in rest.Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries))
        {
            var candidate = Path.Combine(current, segment);
            var resolved = TryResolveLink(candidate);
            // A symlink target may itself contain ".." or further links; normalize again.
            current = resolved is null ? candidate : RealPath(Path.GetFullPath(resolved));
        }

        return current;
    }

    private static string? TryResolveLink(string path)
    {
        FileSystemInfo info = Directory.Exists(path) ? new DirectoryInfo(path) : new FileInfo(path);
        if (!info.Exists || info.LinkTarget is null)
        {
            return null;
        }

        return info.ResolveLinkTarget(returnFinalTarget: true)?.FullName;
    }
}
