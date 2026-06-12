using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using AIOffice.Preview;

namespace AIOffice.Mcp;

/// <summary>
/// The MCP side of live preview. <c>preview_open</c> spawns a detached
/// <c>aioffice preview open</c> child (the blocking server cannot live inside
/// the MCP process, whose stdio belongs to JSON-RPC), waits for the lockfile
/// and an HTTP health probe (max 10s) and returns the url.
/// <c>preview_selection</c> reads the human's click selection via
/// <see cref="PreviewClient"/>.
/// </summary>
internal static class PreviewTools
{
    /// <summary>Overrides own-executable resolution (used by tests and unusual hosts).</summary>
    public const string ExecutableEnvVar = "AIOFFICE_EXE";

    private static readonly TimeSpan StartupTimeout = TimeSpan.FromSeconds(10);
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(3) };

    // ---------------------------------------------------------- preview_open

    public static Envelope Open(Workspace workspace, JsonObject args)
    {
        var file = RequireFileArg(args, "Pass the document to preview, e.g. report.docx.");
        var resolved = workspace.Resolve(file, mustExist: true);
        var port = OptionalPort(args);

        // Idempotent open: a live preview for this file is an answer, not an error.
        var lockPath = PreviewLock.PathFor(resolved);
        if (PreviewLock.TryRead(lockPath) is { } existing && PreviewLock.IsPortAlive(existing.Port))
        {
            return Envelope.Ok(
                new { url = Url(existing.Port), port = existing.Port, pid = existing.Pid },
                new Meta
                {
                    File = resolved,
                    Warnings = [new Warning("already_running", "A preview for this file was already running; returning it.")],
                });
        }

        var executable = ResolveOwnExecutable();
        var child = Spawn(executable, resolved, workspace.Root, port);
        return AwaitStartup(child, lockPath, resolved);
    }

    /// <summary>Spawns the detached preview server child, stdio fully redirected away from JSON-RPC.</summary>
    private static Process Spawn(string executable, string resolvedFile, string workspaceRoot, int port)
    {
        var psi = new ProcessStartInfo(executable)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            WorkingDirectory = workspaceRoot,
        };
        psi.ArgumentList.Add("preview");
        psi.ArgumentList.Add("open");
        psi.ArgumentList.Add(resolvedFile);
        psi.ArgumentList.Add("--workspace");
        psi.ArgumentList.Add(workspaceRoot);
        psi.ArgumentList.Add("--json");
        if (port != 0)
        {
            psi.ArgumentList.Add("--port");
            psi.ArgumentList.Add(port.ToString(CultureInfo.InvariantCulture));
        }

        try
        {
            return Process.Start(psi) ?? throw new InvalidOperationException("Process.Start returned null.");
        }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException)
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                $"Failed to start the preview server child '{executable}': {ex.Message}",
                $"Check that the aioffice executable is runnable, or set {ExecutableEnvVar} to its path.",
                innerException: ex);
        }
    }

    /// <summary>Polls for the lockfile + a healthy HTTP root for up to 10s.</summary>
    private static Envelope AwaitStartup(Process child, string lockPath, string resolvedFile)
    {
        // Keep the pipes drained so the child can never block on a full buffer;
        // stdout is also our error channel when startup fails.
        var stdoutTask = child.StandardOutput.ReadToEndAsync();
        var stderrTask = child.StandardError.ReadToEndAsync();

        var clock = Stopwatch.StartNew();
        while (clock.Elapsed < StartupTimeout)
        {
            if (PreviewLock.TryRead(lockPath) is { } lockfile && PreviewLock.IsPortAlive(lockfile.Port) &&
                Healthy(lockfile.Port))
            {
                return Envelope.Ok(
                    new { url = Url(lockfile.Port), port = lockfile.Port, pid = lockfile.Pid },
                    new Meta { File = resolvedFile });
            }

            if (child.HasExited)
            {
                // The child printed exactly one envelope before dying — relay it.
                return ChildFailure(stdoutTask.Result, stderrTask.Result, child.ExitCode);
            }

            Thread.Sleep(100);
        }

        TryKill(child);
        throw new AiofficeException(
            ErrorCodes.InternalError,
            $"The preview server did not come up within {StartupTimeout.TotalSeconds:0}s.",
            "Re-run preview_open; if it keeps timing out, start it manually with 'aioffice preview open <file>' and read its output.");
    }

    private static bool Healthy(int port)
    {
        try
        {
            using var response = Http.GetAsync(Url(port)).GetAwaiter().GetResult();
            return response.IsSuccessStatusCode;
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
        {
            return false;
        }
    }

    private static Envelope ChildFailure(string stdout, string stderr, int exitCode)
    {
        try
        {
            if (JsonSerializer.Deserialize<Envelope>(stdout) is { IsOk: false } relayed)
            {
                return relayed; // the child's typed failure envelope is the real answer
            }
        }
        catch (JsonException)
        {
            // Fall through to the generic failure below.
        }

        var detail = (stderr.Length > 0 ? stderr : stdout).ReplaceLineEndings(" ").Trim();
        return Envelope.Fail(
            ErrorCodes.InternalError,
            $"The preview server child exited with code {exitCode} during startup." +
            (detail.Length > 0 ? $" Output: {detail[..Math.Min(detail.Length, 400)]}" : string.Empty),
            "Run 'aioffice preview open <file>' manually to see the full error.");
    }

    private static void TryKill(Process child)
    {
        try
        {
            child.Kill(entireProcessTree: true);
        }
        catch (Exception ex) when (ex is InvalidOperationException or SystemException)
        {
            // Already gone.
        }
    }

    /// <summary>
    /// Finds the aioffice executable hosting this very server: env override,
    /// then the current process (when it IS aioffice), then an aioffice binary
    /// sitting next to this assembly.
    /// </summary>
    internal static string ResolveOwnExecutable()
    {
        if (Environment.GetEnvironmentVariable(ExecutableEnvVar) is { Length: > 0 } overridePath &&
            File.Exists(overridePath))
        {
            return Path.GetFullPath(overridePath);
        }

        if (Environment.ProcessPath is { } processPath &&
            Path.GetFileNameWithoutExtension(processPath).Equals("aioffice", StringComparison.OrdinalIgnoreCase))
        {
            return processPath;
        }

        var sibling = Path.Combine(AppContext.BaseDirectory, OperatingSystem.IsWindows() ? "aioffice.exe" : "aioffice");
        if (File.Exists(sibling))
        {
            return sibling;
        }

        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            "Cannot locate the aioffice executable to spawn the preview server from this host.",
            $"Run 'aioffice preview open <file>' in a terminal instead, or set {ExecutableEnvVar} to the aioffice binary.");
    }

    // ----------------------------------------------------- preview_selection

    public static Envelope Selection(Workspace workspace, JsonObject args)
    {
        var file = RequireFileArg(args, "Pass the previewed document, e.g. report.docx.");
        var resolved = workspace.Resolve(file);
        var snapshot = PreviewClient.GetSelection(resolved);
        return Envelope.Ok(
            new { paths = snapshot.Paths, rev = snapshot.Rev, updatedAt = snapshot.UpdatedAt },
            new Meta { File = resolved, Rev = snapshot.Rev });
    }

    // -------------------------------------------------------------- plumbing

    private static string RequireFileArg(JsonObject args, string suggestion) =>
        args["file"] is JsonValue value && value.TryGetValue<string>(out var s) && s.Length > 0
            ? s
            : throw new AiofficeException(ErrorCodes.InvalidArgs, "'file' is required.", suggestion);

    private static int OptionalPort(JsonObject args)
    {
        var node = args["port"];
        if (node is null)
        {
            return 0;
        }

        if (node is JsonValue value && value.TryGetValue<int>(out var port) && port is >= 0 and <= 65535)
        {
            return port;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "'port' must be an integer in 0-65535.",
            "Omit port to auto-pick a free one in 26500-26600.");
    }

    private static string Url(int port) => string.Create(CultureInfo.InvariantCulture, $"http://127.0.0.1:{port}/");
}
