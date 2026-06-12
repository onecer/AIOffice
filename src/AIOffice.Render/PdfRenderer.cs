using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using AIOffice.Core;

namespace AIOffice.Render;

/// <summary>
/// Renders HTML to paged PDF with a headless Chromium browser
/// (<c>--headless=new --print-to-pdf=&lt;out&gt; --no-pdf-header-footer</c>) —
/// no Office, no native PDF stack. Page geometry is controlled by the input
/// document's print CSS (<c>@page</c> rules). All failures surface as
/// <see cref="AiofficeException"/> so they map onto the standard envelope.
/// </summary>
public static class PdfRenderer
{
    /// <summary>How long one browser print run may take.</summary>
    public static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(60);

    private const string NoBrowserSuggestion =
        "Install Google Chrome/Edge, set AIOFFICE_BROWSER to a Chromium binary, " +
        "or render --to html|svg and print/convert externally.";

    private const string RetrySuggestion =
        "Re-run the command; if it keeps failing, set AIOFFICE_BROWSER to a different " +
        "Chromium binary or render --to html|svg and convert externally.";

    /// <summary>Prints an HTML file into a PDF. Returns the absolute path of the written PDF.</summary>
    public static string HtmlFileToPdf(
        string htmlPath,
        string outPdf,
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

        var info = browser ?? BrowserLocator.Probe();
        if (!info.Found || info.Path is null)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "PDF rendering needs a Chromium-based browser and none was found on this machine.",
                NoBrowserSuggestion);
        }

        var output = Path.GetFullPath(outPdf);
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
            var args = BuildPrintToPdfArgs(html, output, profileDir);
            var (outcome, exitCode, stderrTail) = RunBrowser(info.Path, args, output, timeout ?? DefaultTimeout);

            if (outcome == RunOutcome.TimedOut)
            {
                throw new AiofficeException(
                    ErrorCodes.InternalError,
                    $"Browser timed out after {(timeout ?? DefaultTimeout).TotalSeconds:0}s while rendering PDF.",
                    RetrySuggestion);
            }

            if (outcome == RunOutcome.Exited && exitCode != 0)
            {
                var detail = stderrTail.Length > 0 ? $" stderr: {stderrTail}" : string.Empty;
                throw new AiofficeException(
                    ErrorCodes.InternalError,
                    $"Browser exited with code {exitCode} while rendering PDF.{detail}",
                    RetrySuggestion);
            }

            if (!File.Exists(output) || new FileInfo(output).Length == 0)
            {
                throw new AiofficeException(
                    ErrorCodes.InternalError,
                    "Browser reported success but produced no PDF output.",
                    RetrySuggestion);
            }

            return output;
        }
        finally
        {
            TryDeleteDirectory(profileDir);
        }
    }

    /// <summary>Prints an HTML string (via a temp file) into a PDF.</summary>
    public static string HtmlStringToPdf(
        string html,
        string outPdf,
        BrowserInfo? browser = null,
        TimeSpan? timeout = null)
    {
        var tempDir = Directory.CreateTempSubdirectory("aioffice-render-html-").FullName;
        var tempHtml = Path.Combine(tempDir, "input.html");
        try
        {
            File.WriteAllText(tempHtml, html);
            return HtmlFileToPdf(tempHtml, outPdf, browser, timeout);
        }
        finally
        {
            TryDeleteDirectory(tempDir);
        }
    }

    /// <summary>
    /// The exact headless print argument list: <c>--headless=new
    /// --print-to-pdf=&lt;out&gt; --no-pdf-header-footer ... file://&lt;abs&gt;</c>
    /// (URL last). Public so the doctor command can show the command line it
    /// would run.
    /// </summary>
    public static IReadOnlyList<string> BuildPrintToPdfArgs(
        string htmlAbsPath,
        string outPdfAbsPath,
        string? profileDir = null)
    {
        var args = new List<string>
        {
            "--headless=new",
            $"--print-to-pdf={outPdfAbsPath}",
            "--no-pdf-header-footer",
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

        /// <summary>The PDF is structurally complete; the lingering browser was killed.</summary>
        OutputComplete,

        TimedOut,
    }

    /// <summary>
    /// Runs the browser and waits for EITHER process exit OR a structurally
    /// complete PDF (%PDF- header + %%EOF trailer). Some Chrome builds keep
    /// the process alive long after writing the file (GCM registration /
    /// updater keep-alives), so a finished file counts as success and the
    /// straggler is killed.
    /// </summary>
    private static (RunOutcome Outcome, int ExitCode, string StderrTail) RunBrowser(
        string browserPath,
        IReadOnlyList<string> args,
        string outputPdf,
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

                if (IsCompletePdf(outputPdf))
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

    private static readonly byte[] PdfSignature = "%PDF-"u8.ToArray();

    /// <summary>True when the file starts with %PDF- and the tail carries the %%EOF marker.</summary>
    public static bool IsCompletePdf(string path)
    {
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            if (fs.Length < PdfSignature.Length + 5)
            {
                return false;
            }

            Span<byte> head = stackalloc byte[5];
            fs.ReadExactly(head);
            if (!head.SequenceEqual(PdfSignature))
            {
                return false;
            }

            var tailLength = (int)Math.Min(1024, fs.Length);
            var tail = new byte[tailLength];
            fs.Seek(-tailLength, SeekOrigin.End);
            fs.ReadExactly(tail);
            return Encoding.ASCII.GetString(tail).Contains("%%EOF", StringComparison.Ordinal);
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
