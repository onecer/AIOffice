using System.Globalization;

namespace AIOffice.Core;

/// <summary>
/// The M2 large-file guard: opening any document bigger than the limit fails
/// fast with a typed <c>file_too_large</c> envelope instead of grinding (or
/// dying) inside a multi-hundred-MB OOXML parse. The default limit is
/// <see cref="DefaultMaxFileMb"/> MB; <c>AIOFFICE_MAX_FILE_MB</c> overrides it
/// (read per call, so long-lived MCP servers honour changes too). Real
/// streaming for huge workbooks is an M3 item — this guard is the honest
/// stop-gap.
/// </summary>
public static class FileSizeGuard
{
    public const int DefaultMaxFileMb = 50;

    public const string EnvVar = "AIOFFICE_MAX_FILE_MB";

    /// <summary>The effective limit in MB: env override when it parses as a non-negative int, else the default.</summary>
    public static int MaxFileMb =>
        Environment.GetEnvironmentVariable(EnvVar) is { } raw &&
        int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out var mb)
            ? mb
            : DefaultMaxFileMb;

    /// <summary>
    /// Throws <c>file_too_large</c> when the file at <paramref name="resolvedPath"/>
    /// exists and exceeds the effective limit. Missing files pass (the caller's
    /// own not-found handling stays authoritative).
    /// </summary>
    public static void Ensure(string resolvedPath) => Ensure(resolvedPath, MaxFileMb);

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
            $"Split the document into smaller files, or raise the limit with {EnvVar}=<mb> " +
            $"(current limit {maxFileMb}; default {DefaultMaxFileMb}).");
    }
}
