using System.ComponentModel;
using System.Diagnostics;
using AIOffice.Core;

namespace AIOffice.Render;

/// <summary>
/// Renders HTML/SVG to PNG by screenshotting with a headless Chromium browser
/// (no Office, no native rendering stack). All failures surface as
/// <see cref="AiofficeException"/> so they map onto the standard envelope.
/// </summary>
public static class PngRenderer
{
    /// <summary>How long one browser screenshot run may take.</summary>
    public static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(30);

    /// <summary>Viewport height used when the caller asks for auto height.</summary>
    public const int DefaultHeightPx = 720;

    private const string NoBrowserSuggestion =
        "Install Google Chrome/Edge, set AIOFFICE_BROWSER to a Chromium binary, or use --to svg|html instead.";

    private const string RetrySuggestion =
        "Re-run the command; if it keeps failing, set AIOFFICE_BROWSER to a different Chromium binary or use --to svg|html instead.";

    /// <summary>
    /// Screenshots an HTML file into a PNG. <paramref name="heightPx"/> null
    /// means auto, which currently uses a <see cref="DefaultHeightPx"/>-tall
    /// viewport (full-page height detection is a later milestone). Returns the
    /// absolute path of the written PNG.
    /// </summary>
    public static string HtmlFileToPng(
        string htmlPath,
        string outPng,
        int widthPx = 1280,
        int? heightPx = null,
        BrowserInfo? browser = null,
        TimeSpan? timeout = null)
    {
        var html = Path.GetFullPath(htmlPath);
        if (!File.Exists(html))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"HTML input not found: {html}",
                "Render the document to HTML first, or pass the path of an existing .html file.");
        }

        if (widthPx <= 0 || heightPx is <= 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"PNG dimensions must be positive (got {widthPx}x{heightPx?.ToString() ?? "auto"}).",
                "Pass widthPx > 0 and either omit heightPx (auto) or pass a positive value.");
        }

        var info = browser ?? BrowserLocator.Probe();
        if (!info.Found || info.Path is null)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "PNG rendering needs a Chromium-based browser and none was found on this machine.",
                NoBrowserSuggestion);
        }

        var output = Path.GetFullPath(outPng);
        if (Path.GetDirectoryName(output) is { Length: > 0 } outDir)
        {
            Directory.CreateDirectory(outDir);
        }

        if (File.Exists(output))
        {
            File.Delete(output); // stale output must never pass the "browser wrote it" check
        }

        // A private profile dir avoids first-run prompts and locks held by a
        // running desktop browser instance.
        var profileDir = Directory.CreateTempSubdirectory("aioffice-render-profile-").FullName;
        try
        {
            var args = BuildScreenshotArgs(html, output, widthPx, heightPx ?? DefaultHeightPx, profileDir);
            var (outcome, exitCode, stderrTail) = RunBrowser(info.Path, args, output, timeout ?? DefaultTimeout);

            if (outcome == RunOutcome.TimedOut)
            {
                throw new AiofficeException(
                    ErrorCodes.InternalError,
                    $"Browser timed out after {(timeout ?? DefaultTimeout).TotalSeconds:0}s while rendering PNG.",
                    RetrySuggestion);
            }

            if (outcome == RunOutcome.Exited && exitCode != 0)
            {
                var detail = stderrTail.Length > 0 ? $" stderr: {stderrTail}" : string.Empty;
                throw new AiofficeException(
                    ErrorCodes.InternalError,
                    $"Browser exited with code {exitCode} while rendering PNG.{detail}",
                    RetrySuggestion);
            }

            if (!File.Exists(output) || new FileInfo(output).Length == 0)
            {
                throw new AiofficeException(
                    ErrorCodes.InternalError,
                    "Browser reported success but produced no PNG output.",
                    RetrySuggestion);
            }

            return output;
        }
        finally
        {
            TryDeleteDirectory(profileDir);
        }
    }

    /// <summary>Screenshots an HTML string (via a temp file) into a PNG.</summary>
    public static string HtmlStringToPng(
        string html,
        string outPng,
        int widthPx = 1280,
        int? heightPx = null,
        BrowserInfo? browser = null,
        TimeSpan? timeout = null)
    {
        var tempDir = Directory.CreateTempSubdirectory("aioffice-render-html-").FullName;
        var tempHtml = Path.Combine(tempDir, "input.html");
        try
        {
            File.WriteAllText(tempHtml, html);
            return HtmlFileToPng(tempHtml, outPng, widthPx, heightPx, browser, timeout);
        }
        finally
        {
            TryDeleteDirectory(tempDir);
        }
    }

    /// <summary>
    /// Screenshots an SVG document into a PNG by wrapping it in a minimal HTML
    /// page with a white background.
    /// </summary>
    public static string SvgToPng(
        string svg,
        string outPng,
        int widthPx = 1280,
        int? heightPx = null,
        BrowserInfo? browser = null,
        TimeSpan? timeout = null)
    {
        if (!svg.Contains("<svg", StringComparison.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "SvgToPng input does not contain an <svg> element.",
                "Pass a complete <svg ...>...</svg> document, e.g. the output of render --to svg.");
        }

        var page =
            "<!DOCTYPE html><html><head><meta charset=\"utf-8\">" +
            "<style>html,body{margin:0;padding:0;background:#ffffff}</style>" +
            $"</head><body>{svg}</body></html>";
        return HtmlStringToPng(page, outPng, widthPx, heightPx, browser, timeout);
    }

    /// <summary>
    /// The exact headless-screenshot argument list: <c>--headless=new
    /// --screenshot=&lt;out&gt; --window-size=W,H --hide-scrollbars
    /// --disable-gpu ... file://&lt;abs&gt;</c> (URL last). Public so the
    /// doctor command can show the command line it would run.
    /// </summary>
    public static IReadOnlyList<string> BuildScreenshotArgs(
        string htmlAbsPath,
        string outPngAbsPath,
        int widthPx,
        int heightPx,
        string? profileDir = null)
    {
        var args = new List<string>
        {
            "--headless=new",
            $"--screenshot={outPngAbsPath}",
            $"--window-size={widthPx},{heightPx}",
            "--hide-scrollbars",
            "--disable-gpu",
            "--no-first-run",
        };
        if (profileDir is not null)
        {
            args.Add($"--user-data-dir={profileDir}");
        }

        args.Add(new Uri(Path.GetFullPath(htmlAbsPath)).AbsoluteUri);
        return args;
    }

    private enum RunOutcome
    {
        /// <summary>The browser exited on its own; trust its exit code.</summary>
        Exited,

        /// <summary>The PNG is structurally complete; the lingering browser was killed.</summary>
        OutputComplete,

        TimedOut,
    }

    /// <summary>
    /// Runs the browser and waits for EITHER process exit OR a structurally
    /// complete PNG (signature + IEND trailer). Some Chrome builds keep the
    /// process alive for 20s+ after writing the screenshot (GCM registration /
    /// updater keep-alives), so a finished file counts as success and the
    /// straggler is killed.
    /// </summary>
    private static (RunOutcome Outcome, int ExitCode, string StderrTail) RunBrowser(
        string browserPath,
        IReadOnlyList<string> args,
        string outputPng,
        TimeSpan timeout)
    {
        var psi = new ProcessStartInfo(browserPath)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
        };
        foreach (var arg in args)
        {
            psi.ArgumentList.Add(arg);
        }

        Process process;
        try
        {
            process = Process.Start(psi) ?? throw new InvalidOperationException("Process.Start returned null.");
        }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException or PlatformNotSupportedException)
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                $"Failed to start browser '{browserPath}': {ex.Message}",
                "Check that AIOFFICE_BROWSER points to an executable Chromium binary, or unset it to auto-detect.",
                innerException: ex);
        }

        using (process)
        {
            var stderrTask = process.StandardError.ReadToEndAsync();
            var stdoutTask = process.StandardOutput.ReadToEndAsync();

            var clock = Stopwatch.StartNew();
            while (clock.Elapsed < timeout)
            {
                if (process.WaitForExit(100))
                {
                    process.WaitForExit(); // flush redirected streams
                    _ = stdoutTask.Result;
                    return (RunOutcome.Exited, process.ExitCode, Tail(stderrTask.Result));
                }

                if (IsCompletePng(outputPng))
                {
                    KillTree(process);
                    return (RunOutcome.OutputComplete, 0, string.Empty);
                }
            }

            KillTree(process);
            return (RunOutcome.TimedOut, -1, string.Empty);
        }
    }

    private static void KillTree(Process process)
    {
        try
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit();
        }
        catch (Exception ex) when (ex is InvalidOperationException or SystemException)
        {
            // The process raced us to exit; nothing left to kill.
        }
    }

    private static readonly byte[] PngSignature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
    private static readonly byte[] PngIendTrailer = [0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82];

    /// <summary>True when the file starts with the PNG signature and ends with the IEND chunk.</summary>
    private static bool IsCompletePng(string path)
    {
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            if (fs.Length < PngSignature.Length + PngIendTrailer.Length)
            {
                return false;
            }

            Span<byte> head = stackalloc byte[8];
            Span<byte> tail = stackalloc byte[8];
            fs.ReadExactly(head);
            fs.Seek(-PngIendTrailer.Length, SeekOrigin.End);
            fs.ReadExactly(tail);
            return head.SequenceEqual(PngSignature) && tail.SequenceEqual(PngIendTrailer);
        }
        catch (IOException)
        {
            return false; // not there yet, or still being written
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>Last 500 chars of stderr, single-line, for error messages.</summary>
    private static string Tail(string stderr)
    {
        var flat = stderr.ReplaceLineEndings(" ").Trim();
        return flat.Length <= 500 ? flat : flat[^500..];
    }

    private static void TryDeleteDirectory(string dir)
    {
        try
        {
            Directory.Delete(dir, recursive: true);
        }
        catch (IOException)
        {
            // Best-effort cleanup of temp dirs.
        }
        catch (UnauthorizedAccessException)
        {
            // Best-effort cleanup of temp dirs.
        }
    }
}
