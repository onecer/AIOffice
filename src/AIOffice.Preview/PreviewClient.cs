using System.Globalization;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Nodes;
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

    /// <summary>Reads the current advisory marks of the preview for <paramref name="file"/>.</summary>
    public static MarksSnapshot GetMarks(string file, string? lockDirectory = null)
    {
        var (lockPath, lockfile) = RequireRunning(file, lockDirectory);
        try
        {
            var json = Http.GetStringAsync(RouteUrl(lockfile.Port, "marks")).GetAwaiter().GetResult();
            return JsonSerializer.Deserialize<MarksSnapshot>(json, JsonDefaults.Options) ?? throw NotRunning(file);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
        {
            PreviewLock.Delete(lockPath);
            throw NotRunning(file, ex);
        }
    }

    /// <summary>Adds/replaces a mark (path may be the pseudo-path "selected").</summary>
    public static MarksSnapshot AddMark(string file, string path, string? color, string? note, string? find, bool toFix, string? lockDirectory = null) =>
        SendMarks(file, HttpMethod.Post, new MarkRequest(path, color, note, find, toFix), lockDirectory);

    /// <summary>Removes the mark on <paramref name="path"/>.</summary>
    public static MarksSnapshot RemoveMark(string file, string path, string? lockDirectory = null) =>
        SendMarks(file, HttpMethod.Delete, new MarkRequest(path, null, null, null, false), lockDirectory);

    /// <summary>Clears every mark.</summary>
    public static MarksSnapshot ClearMarks(string file, string? lockDirectory = null) =>
        SendMarks(file, HttpMethod.Delete, new ClearRequest(true), lockDirectory);

    /// <summary>Pushes a scroll-to-path command to every live viewer.</summary>
    public static void Goto(string file, string path, string? lockDirectory = null)
    {
        var (lockPath, lockfile) = RequireRunning(file, lockDirectory);
        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, RouteUrl(lockfile.Port, "goto"))
            {
                Content = JsonContent.Create(new { path }),
            };
            using var response = Http.Send(request);
            EnsureOk(response, response.Content.ReadAsStringAsync().GetAwaiter().GetResult());
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
        {
            PreviewLock.Delete(lockPath);
            throw NotRunning(file, ex);
        }
    }

    private sealed record MarkRequest(string Path, string? Color, string? Note, string? Find, bool ToFix);

    private sealed record ClearRequest(bool All);

    private static MarksSnapshot SendMarks(string file, HttpMethod method, object body, string? lockDirectory)
    {
        var (lockPath, lockfile) = RequireRunning(file, lockDirectory);
        try
        {
            using var request = new HttpRequestMessage(method, RouteUrl(lockfile.Port, "marks"))
            {
                Content = JsonContent.Create(body, options: JsonDefaults.Options),
            };
            using var response = Http.Send(request);
            var json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            EnsureOk(response, json);
            return JsonSerializer.Deserialize<MarksSnapshot>(json, JsonDefaults.Options) ?? throw NotRunning(file);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
        {
            PreviewLock.Delete(lockPath);
            throw NotRunning(file, ex);
        }
    }

    /// <summary>Surfaces a non-2xx preview response as its typed error envelope.</summary>
    private static void EnsureOk(HttpResponseMessage response, string body)
    {
        if (response.IsSuccessStatusCode)
        {
            return;
        }

        try
        {
            if (JsonNode.Parse(body)?["error"] is JsonObject error && error["code"] is { } code)
            {
                throw new AiofficeException(
                    code.GetValue<string>(),
                    error["message"]?.GetValue<string>() ?? "Preview request failed.",
                    error["suggestion"]?.GetValue<string>() ?? "Retry, or reopen the preview.");
            }
        }
        catch (JsonException)
        {
            // fall through to the generic error
        }

        throw new AiofficeException(
            ErrorCodes.InternalError,
            string.Create(CultureInfo.InvariantCulture, $"Preview request failed (HTTP {(int)response.StatusCode})."),
            "Reopen the preview ('aioffice preview open <file>') and retry.");
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
