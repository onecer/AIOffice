using System.Text.Json.Serialization;

namespace AIOffice.Render;

/// <summary>
/// Result of probing for a headless-capable Chromium browser. Serialized into
/// doctor output, so property names are pinned to the wire contract.
/// </summary>
public sealed record BrowserInfo(
    [property: JsonPropertyName("found")] bool Found,
    [property: JsonPropertyName("path")] string? Path,
    [property: JsonPropertyName("kind")] string? Kind)
{
    /// <summary>The single "nothing usable on this machine" value.</summary>
    public static BrowserInfo NotFound { get; } = new(false, null, null);
}

/// <summary>
/// Finds a Chromium-based browser for PNG rendering. Probe order:
/// <c>$AIOFFICE_BROWSER</c> (explicit path), PATH (chrome, google-chrome,
/// chromium, msedge), well-known macOS app bundles, well-known Windows install
/// paths. The result is cached per process.
/// </summary>
public static class BrowserLocator
{
    /// <summary>Environment variable that pins an explicit browser binary.</summary>
    public const string EnvVar = "AIOFFICE_BROWSER";

    private static readonly string[] PathNames = ["chrome", "google-chrome", "chromium", "msedge"];

    private static readonly string[] MacAppPaths =
    [
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
        "/Applications/Chromium.app/Contents/MacOS/Chromium",
    ];

    private static readonly Lock Gate = new();
    private static BrowserInfo? _cached;

    /// <summary>
    /// Probes for a usable browser, caching the answer for the process.
    /// Pass <paramref name="refresh"/> to re-probe (e.g. after the user
    /// installed a browser or changed <c>$AIOFFICE_BROWSER</c>).
    /// </summary>
    public static BrowserInfo Probe(bool refresh = false)
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

    private static BrowserInfo ProbeNow()
    {
        // 1. Explicit override wins when it points at a real file.
        if (Environment.GetEnvironmentVariable(EnvVar) is { } env && !string.IsNullOrWhiteSpace(env))
        {
            var full = System.IO.Path.GetFullPath(env);
            if (File.Exists(full))
            {
                return Found(full);
            }
        }

        // 2. Anything on PATH.
        foreach (var name in PathNames)
        {
            if (FindOnPath(name) is { } hit)
            {
                return Found(hit);
            }
        }

        // 3. macOS app bundles.
        foreach (var candidate in MacAppPaths)
        {
            if (File.Exists(candidate))
            {
                return Found(candidate);
            }
        }

        // 4. Windows install locations.
        foreach (var candidate in WindowsPaths())
        {
            if (File.Exists(candidate))
            {
                return Found(candidate);
            }
        }

        return BrowserInfo.NotFound;
    }

    private static BrowserInfo Found(string path) => new(true, path, ClassifyKind(path));

    /// <summary>chrome | edge | chromium | custom, judged from the binary name.</summary>
    private static string ClassifyKind(string path)
    {
        var name = System.IO.Path.GetFileNameWithoutExtension(path).ToLowerInvariant();
        if (name.Contains("edge"))
        {
            return "edge";
        }

        if (name.Contains("chromium"))
        {
            return "chromium";
        }

        if (name.Contains("chrome"))
        {
            return "chrome";
        }

        return "custom";
    }

    private static string? FindOnPath(string name)
    {
        if (Environment.GetEnvironmentVariable("PATH") is not { } pathVar)
        {
            return null;
        }

        string[] suffixes = OperatingSystem.IsWindows() ? [".exe", ""] : [""];
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

    private static IEnumerable<string> WindowsPaths()
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

            yield return System.IO.Path.Combine(root, "Google", "Chrome", "Application", "chrome.exe");
            yield return System.IO.Path.Combine(root, "Microsoft", "Edge", "Application", "msedge.exe");
        }
    }
}
