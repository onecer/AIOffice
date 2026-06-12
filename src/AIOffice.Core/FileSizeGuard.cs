using System.Globalization;

namespace AIOffice.Core;

/// <summary>
/// The large-file guard. M3 directive (功能第一 — features first): the default
/// is now UNLIMITED; the guard only fires when <c>AIOFFICE_MAX_FILE_MB</c> is
/// set as an explicit opt-in cap (read per call, so long-lived MCP servers
/// honour changes too). Reads of huge workbooks are served by the xlsx
/// streaming path; the cap remains available for callers that prefer a hard
/// <c>file_too_large</c> stop over a slow open.
/// </summary>
public static class FileSizeGuard
{
    public const string EnvVar = "AIOFFICE_MAX_FILE_MB";

    /// <summary>
    /// The effective limit in MB: the env var when it parses as a non-negative
    /// int, else null — which means unlimited (the M3 default).
    /// </summary>
    public static int? MaxFileMb =>
        Environment.GetEnvironmentVariable(EnvVar) is { } raw &&
        int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out var mb)
            ? mb
            : null;

    /// <summary>
    /// Throws <c>file_too_large</c> when an opt-in cap is configured and the
    /// file at <paramref name="resolvedPath"/> exists and exceeds it. With no
    /// cap (the default) this is a no-op. Missing files pass (the caller's
    /// own not-found handling stays authoritative).
    /// </summary>
    public static void Ensure(string resolvedPath)
    {
        if (MaxFileMb is { } limit)
        {
            Ensure(resolvedPath, limit);
        }
    }

    /// <summary>Same check against an explicit limit (exposed for tests).</summary>
    public static void Ensure(string resolvedPath, int maxFileMb)
    {
        var info = new FileInfo(resolvedPath);
        if (!info.Exists)
        {
            return;
        }

        var maxBytes = (long)maxFileMb * 1024 * 1024;
        if (info.Length <= maxBytes)
        {
            return;
        }

        throw new AiofficeException(
            ErrorCodes.FileTooLarge,
            $"File is {info.Length / (1024.0 * 1024.0):F1} MB, over the {maxFileMb} MB limit: {info.Name}",
            $"Raise or unset {EnvVar} (it is an opt-in cap; default is unlimited), or split the document " +
            $"into smaller files (current limit {maxFileMb} MB).");
    }
}
