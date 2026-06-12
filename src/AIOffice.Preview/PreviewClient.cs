using System.Globalization;
using System.Text.Json;
using AIOffice.Core;

namespace AIOffice.Preview;

/// <summary>
/// Talks to a running <see cref="PreviewServer"/> through its lockfile: reads
/// the current selection or shuts the server down. Used by the CLI/MCP wiring
/// (e.g. 'aioffice preview selection' / 'aioffice preview close').
/// </summary>
public static class PreviewClient
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(5) };

    /// <summary>
    /// The current selection of the preview for <paramref name="file"/>.
    /// </summary>
    /// <exception cref="AiofficeException"><c>preview_not_running</c> when no live server exists for the file.</exception>
    public static SelectionSnapshot GetSelection(string file, string? lockDirectory = null)
    {
        var (lockPath, lockfile) = RequireRunning(file, lockDirectory);
        try
        {
            var json = Http.GetStringAsync(RouteUrl(lockfile.Port, "selection")).GetAwaiter().GetResult();
            return JsonSerializer.Deserialize<SelectionSnapshot>(json, JsonDefaults.Options)
                ?? throw NotRunning(file);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
        {
            PreviewLock.Delete(lockPath); // the recorded server is gone; drop the stale lockfile
            throw NotRunning(file, ex);
        }
    }

    /// <summary>
    /// Gracefully stops the preview for <paramref name="file"/> via POST /shutdown.
    /// </summary>
    /// <exception cref="AiofficeException"><c>preview_not_running</c> when no live server exists for the file.</exception>
    public static void Close(string file, string? lockDirectory = null)
    {
        var (lockPath, lockfile) = RequireRunning(file, lockDirectory);
        try
        {
            using var response = Http.PostAsync(RouteUrl(lockfile.Port, "shutdown"), content: null)
                .GetAwaiter().GetResult();
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
        {
            PreviewLock.Delete(lockPath);
            throw NotRunning(file, ex);
        }
    }

    private static string RouteUrl(int port, string route) =>
        string.Create(CultureInfo.InvariantCulture, $"http://127.0.0.1:{port}/{route}");

    /// <summary>Finds the live lockfile for the file or throws <c>preview_not_running</c>.</summary>
    private static (string LockPath, PreviewLockfile Lockfile) RequireRunning(string file, string? lockDirectory)
    {
        var absolute = Path.GetFullPath(file);
        var lockPath = PreviewLock.PathFor(absolute, lockDirectory);
        var lockfile = PreviewLock.TryRead(lockPath) ?? FindByRecordedPath(absolute, lockDirectory, out lockPath);
        if (lockfile is null)
        {
            throw NotRunning(file);
        }

        if (!PreviewLock.IsPortAlive(lockfile.Port))
        {
            PreviewLock.Delete(lockPath);
            throw NotRunning(file);
        }

        return (lockPath, lockfile);
    }

    /// <summary>
    /// Fallback for symlinked paths: the server keys the lockfile by the
    /// sandbox-canonicalized path, which can differ from GetFullPath (e.g.
    /// /tmp vs /private/tmp on macOS). Scan the lock directory and match the
    /// recorded file path instead.
    /// </summary>
    private static PreviewLockfile? FindByRecordedPath(string absolute, string? lockDirectory, out string lockPath)
    {
        lockPath = string.Empty;
        var directory = lockDirectory ?? PreviewLock.DefaultDirectory;
        if (!Directory.Exists(directory))
        {
            return null;
        }

        foreach (var candidate in Directory.EnumerateFiles(directory, "*.json"))
        {
            if (PreviewLock.TryRead(candidate) is { } lockfile && PathsLookEqual(lockfile.File, absolute))
            {
                lockPath = candidate;
                return lockfile;
            }
        }

        return null;
    }

    private static bool PathsLookEqual(string recorded, string requested)
    {
        if (string.Equals(recorded, requested, StringComparison.Ordinal))
        {
            return true;
        }

        try
        {
            // Same file by identity once trailing symlinks are resolved.
            var recordedReal = new FileInfo(recorded).ResolveLinkTarget(returnFinalTarget: true)?.FullName ?? recorded;
            var requestedReal = new FileInfo(requested).ResolveLinkTarget(returnFinalTarget: true)?.FullName ?? requested;
            return string.Equals(recordedReal, requestedReal, StringComparison.Ordinal);
        }
        catch (IOException)
        {
            return false;
        }
    }

    private static AiofficeException NotRunning(string file, Exception? inner = null) => new(
        ErrorCodes.PreviewNotRunning,
        $"No preview server is running for: {file}",
        $"Run 'aioffice preview open {file}' first.",
        innerException: inner);
}
