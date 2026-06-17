using System.Text.Json.Serialization;

namespace AIOffice.Render;

/// <summary>
/// Result of probing for a LibreOffice (<c>soffice</c>) binary and the
/// <c>pdftoppm</c> (poppler) helper used to rasterize PDF pages to PNG.
/// Serialized into doctor output, so property names are pinned to the wire
/// contract.
/// </summary>
public sealed record SofficeInfo(
    [property: JsonPropertyName("found")] bool Found,
    [property: JsonPropertyName("path")] string? Path,
    [property: JsonPropertyName("pdftoppm")] bool Pdftoppm,
    [property: JsonPropertyName("pdftoppmPath")] string? PdftoppmPath)
{
    /// <summary>The "nothing usable on this machine" value.</summary>
    public static SofficeInfo NotFound { get; } = new(false, null, false, null);
}

/// <summary>
/// Finds a LibreOffice <c>soffice</c> binary (and the companion
/// <c>pdftoppm</c> from poppler) for TRUE-fidelity rendering. Probe order:
/// <c>$AIOFFICE_SOFFICE</c> (explicit path), PATH (<c>soffice</c>,
/// <c>libreoffice</c>, on Windows also <c>soffice.com</c>), then well-known
/// per-platform install locations (macOS app bundle, linux <c>/usr/bin</c>,
/// Windows Program Files). The result is cached per process. This engine is
/// entirely optional — when nothing is found the default chromium engine is
/// 100% unaffected.
/// </summary>
public static class SofficeLocator
{
    /// <summary>Environment variable that pins an explicit soffice binary.</summary>
    public const string EnvVar = "AIOFFICE_SOFFICE";

    private static readonly string[] PathNames = ["soffice", "libreoffice", "soffice.com"];

    private static readonly string[] MacAppPaths =
    [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ];

    private static readonly string[] LinuxPaths =
    [
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/opt/libreoffice/program/soffice",
        "/snap/bin/libreoffice",
    ];

    private static readonly string[] PdftoppmNames = ["pdftoppm", "pdftoppm.exe"];

    private static readonly Lock Gate = new();
    private static SofficeInfo? _cached;

    /// <summary>
    /// Probes for a usable soffice binary (and pdftoppm), caching the answer for
    /// the process. Pass <paramref name="refresh"/> to re-probe (e.g. after the
    /// user installed LibreOffice or changed <c>$AIOFFICE_SOFFICE</c>).
    /// </summary>
    public static SofficeInfo Probe(bool refresh = false)
    {
        lock (Gate)
        {
            if (refresh || _cached is null)
            {
                _cached = ProbeNow();
            }

            return _cached;
        }
    }

    private static SofficeInfo ProbeNow()
    {
        var (pdftoppmFound, pdftoppmPath) = ProbePdftoppm();

        // 1. Explicit override wins when it points at a real file.
        if (Environment.GetEnvironmentVariable(EnvVar) is { } env && !string.IsNullOrWhiteSpace(env))
        {
            var full = System.IO.Path.GetFullPath(env);
            if (File.Exists(full))
            {
                return new SofficeInfo(true, full, pdftoppmFound, pdftoppmPath);
            }
        }

        // 2. Anything on PATH.
        foreach (var name in PathNames)
        {
            if (FindOnPath(name) is { } hit)
            {
                return new SofficeInfo(true, hit, pdftoppmFound, pdftoppmPath);
            }
        }

        // 3. Per-platform install locations.
        foreach (var candidate in PlatformPaths())
        {
            if (File.Exists(candidate))
            {
                return new SofficeInfo(true, candidate, pdftoppmFound, pdftoppmPath);
            }
        }

        // soffice missing, but still report whether pdftoppm exists.
        return new SofficeInfo(false, null, pdftoppmFound, pdftoppmPath);
    }

    private static (bool Found, string? Path) ProbePdftoppm()
    {
        if (Environment.GetEnvironmentVariable("AIOFFICE_PDFTOPPM") is { } env &&
            !string.IsNullOrWhiteSpace(env))
        {
            var full = System.IO.Path.GetFullPath(env);
            if (File.Exists(full))
            {
                return (true, full);
            }
        }

        foreach (var name in PdftoppmNames)
        {
            if (FindOnPath(name) is { } hit)
            {
                return (true, hit);
            }
        }

        // Common Homebrew / MacPorts location that may not be on a stripped PATH.
        string[] wellKnown =
        [
            "/opt/homebrew/bin/pdftoppm",
            "/usr/local/bin/pdftoppm",
            "/usr/bin/pdftoppm",
        ];
        foreach (var candidate in wellKnown)
        {
            if (File.Exists(candidate))
            {
                return (true, candidate);
            }
        }

        return (false, null);
    }

    private static IEnumerable<string> PlatformPaths()
    {
        if (OperatingSystem.IsMacOS())
        {
            foreach (var path in MacAppPaths)
            {
                yield return path;
            }
        }
        else if (OperatingSystem.IsWindows())
        {
            string?[] roots =
            [
                Environment.GetEnvironmentVariable("ProgramFiles"),
                Environment.GetEnvironmentVariable("ProgramFiles(x86)"),
            ];
            foreach (var root in roots)
            {
                if (string.IsNullOrEmpty(root))
                {
                    continue;
                }

                yield return System.IO.Path.Combine(root, "LibreOffice", "program", "soffice.com");
                yield return System.IO.Path.Combine(root, "LibreOffice", "program", "soffice.exe");
            }
        }
        else
        {
            foreach (var path in LinuxPaths)
            {
                yield return path;
            }
        }
    }

    private static string? FindOnPath(string name)
    {
        if (Environment.GetEnvironmentVariable("PATH") is not { } pathVar)
        {
            return null;
        }

        // The candidate name may already carry its extension (soffice.com); only
        // append .exe on Windows for the bare names.
        var hasExt = System.IO.Path.HasExtension(name);
        string[] suffixes = OperatingSystem.IsWindows() && !hasExt ? [".exe", ".com", ""] : [""];
        foreach (var dir in pathVar.Split(System.IO.Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            foreach (var suffix in suffixes)
            {
                var candidate = System.IO.Path.Combine(dir, name + suffix);
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
        }

        return null;
    }
}
