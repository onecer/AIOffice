using Xunit;

// Locator tests mutate process-wide state (env vars, the probe cache), so the
// whole assembly runs sequentially.
[assembly: CollectionBehavior(DisableTestParallelization = true)]

namespace AIOffice.Render.Tests;

/// <summary>A throwaway temp dir per test, with stub-browser helpers.</summary>
public sealed class TempDir : IDisposable
{
    public TempDir()
    {
        Dir = Directory.CreateTempSubdirectory("aioffice-render-tests-").FullName;
    }

    public string Dir { get; }

    public string PathOf(string name) => Path.Combine(Dir, name);

    /// <summary>Writes a small HTML file and returns its absolute path.</summary>
    public string WriteHtml(string name = "input.html", string body = "<p>hello</p>")
    {
        var path = PathOf(name);
        File.WriteAllText(path, $"<!DOCTYPE html><html><body>{body}</body></html>");
        return path;
    }

    /// <summary>
    /// Writes an executable shell script that stands in for a browser binary.
    /// The body runs after a <c>#!/bin/sh</c> shebang with the original args.
    /// </summary>
    public string WriteStubBrowser(string name, string body)
    {
        var path = PathOf(name);
        File.WriteAllText(path, "#!/bin/sh\n" + body + "\n");
        if (!OperatingSystem.IsWindows())
        {
            File.SetUnixFileMode(
                path,
                UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute |
                UnixFileMode.GroupRead | UnixFileMode.GroupExecute |
                UnixFileMode.OtherRead | UnixFileMode.OtherExecute);
        }

        return path;
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(Dir, recursive: true);
        }
        catch (IOException)
        {
            // Best-effort cleanup of the temp dir.
        }
    }
}

/// <summary>Sets an environment variable for the scope and restores it after.</summary>
public sealed class EnvVarScope : IDisposable
{
    private readonly string _name;
    private readonly string? _previous;

    public EnvVarScope(string name, string? value)
    {
        _name = name;
        _previous = Environment.GetEnvironmentVariable(name);
        Environment.SetEnvironmentVariable(name, value);
    }

    public void Dispose() => Environment.SetEnvironmentVariable(_name, _previous);
}

/// <summary>
/// A fact that needs a REAL Chromium browser on the machine; skipped (not
/// failed) when <see cref="BrowserLocator.Probe"/> finds none, or on CI where
/// headless browser subprocesses are unreliable (the render logic is covered
/// deterministically by the stub-browser tests in this suite).
/// </summary>
public sealed class BrowserFactAttribute : FactAttribute
{
    public BrowserFactAttribute()
    {
        if (IsCi())
        {
            Skip = "Real-browser end-to-end render is skipped on CI (flaky headless subprocess); stub-browser tests cover the logic.";
        }
        else if (!BrowserLocator.Probe(refresh: true).Found)
        {
            Skip = "No Chromium-based browser found on this machine.";
        }
    }

    private static bool IsCi() =>
        string.Equals(Environment.GetEnvironmentVariable("CI"), "true", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(Environment.GetEnvironmentVariable("GITHUB_ACTIONS"), "true", StringComparison.OrdinalIgnoreCase);
}

/// <summary>
/// A fact that needs a REAL LibreOffice (<c>soffice</c>) — and, for PNG,
/// <c>pdftoppm</c> — on the machine; skipped (not failed) when
/// <see cref="SofficeLocator.Probe"/> finds none. CI (macos-14 + windows-latest)
/// has no LibreOffice, so these end-to-end soffice tests skip there exactly like
/// <see cref="BrowserFactAttribute"/> skips the real-browser test — the engine
/// plumbing (arg builders, locator, fallback, failure maps) is covered
/// deterministically by the non-soffice tests in this suite.
/// </summary>
public sealed class SofficeFactAttribute : FactAttribute
{
    /// <param name="needsPdftoppm">PNG tests also require pdftoppm (poppler).</param>
    public SofficeFactAttribute(bool needsPdftoppm = false)
    {
        if (IsCi())
        {
            // Symmetric with BrowserFactAttribute: never spawn a real soffice
            // subprocess on CI even if a runner happens to have LibreOffice —
            // the deterministic non-soffice tests carry the plumbing.
            Skip = "Real-soffice end-to-end render is skipped on CI; the non-soffice tests cover the logic.";
            return;
        }

        var info = SofficeLocator.Probe(refresh: true);
        if (!info.Found)
        {
            Skip = "No LibreOffice (soffice) found on this machine; the soffice engine is optional.";
        }
        else if (needsPdftoppm && !info.Pdftoppm)
        {
            Skip = "No pdftoppm (poppler) found on this machine; soffice PNG needs it.";
        }
    }

    private static bool IsCi() =>
        string.Equals(Environment.GetEnvironmentVariable("CI"), "true", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(Environment.GetEnvironmentVariable("GITHUB_ACTIONS"), "true", StringComparison.OrdinalIgnoreCase);
}
