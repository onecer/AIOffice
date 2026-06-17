using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using AIOffice.Core;

namespace AIOffice.Render;

/// <summary>
/// TRUE-fidelity render engine that drives a headless LibreOffice
/// (<c>soffice --headless --convert-to pdf</c>) to turn a whole Office document
/// into a PDF, then optionally rasterizes one page to PNG with
/// <c>pdftoppm</c> (poppler). Unlike the chromium engine — which screenshots
/// aioffice's own HTML/SVG projection — this hands the original .docx/.xlsx/.pptx
/// to LibreOffice, so the output matches Office's own layout closely. Entirely
/// optional: selected only when the caller passes <c>--engine soffice</c> (or
/// <c>auto</c> with soffice present). All failures surface as
/// <see cref="AiofficeException"/> so they map onto the standard envelope.
/// </summary>
public static class SofficeRenderer
{
    /// <summary>How long one soffice / pdftoppm run may take.</summary>
    public static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(120);

    private const string NoSofficeSuggestion =
        "Install LibreOffice (macOS: the LibreOffice.app bundle; linux: the libreoffice package; " +
        "windows: the LibreOffice installer), set AIOFFICE_SOFFICE to a soffice binary, " +
        "or drop --engine (the default chromium engine needs no LibreOffice).";

    private const string NoPdftoppmSuggestion =
        "PNG at soffice fidelity needs pdftoppm from poppler (macOS/linux: 'brew install poppler' " +
        "or your package manager's poppler-utils; windows: a poppler build on PATH). " +
        "Use --to pdf for a soffice render without poppler, or --engine chromium for a screenshot PNG.";

    /// <summary>
    /// Renders the whole document to PDF with LibreOffice. Returns the absolute
    /// path of the written PDF (moved onto <paramref name="outPdf"/>).
    /// </summary>
    public static string DocumentToPdf(
        string sourceFile,
        string outPdf,
        SofficeInfo? soffice = null,
        TimeSpan? timeout = null)
    {
        var source = RequireSource(sourceFile);
        var info = RequireSoffice(soffice);

        var output = Path.GetFullPath(outPdf);
        EnsureOutputDir(output);
        if (File.Exists(output))
        {
            File.Delete(output); // stale output must never pass the "soffice wrote it" check
        }

        // soffice --convert-to writes "<basename>.pdf" into --outdir; it gives us
        // no control over the leaf name, so convert into a private temp dir and
        // move the single result onto the requested path.
        var workDir = Directory.CreateTempSubdirectory("aioffice-soffice-").FullName;
        var profileDir = Directory.CreateTempSubdirectory("aioffice-soffice-profile-").FullName;
        try
        {
            var args = BuildConvertToPdfArgs(source, workDir, profileDir);
            RunTool(info.Path!, args, timeout ?? DefaultTimeout, "PDF");

            var produced = Path.Combine(workDir, Path.GetFileNameWithoutExtension(source) + ".pdf");
            if (!File.Exists(produced) || new FileInfo(produced).Length == 0)
            {
                // Fall back to whatever single .pdf landed in the work dir.
                produced = Directory.EnumerateFiles(workDir, "*.pdf").FirstOrDefault() ?? produced;
            }

            if (!File.Exists(produced) || new FileInfo(produced).Length == 0)
            {
                throw new AiofficeException(
                    ErrorCodes.InternalError,
                    "LibreOffice reported success but produced no PDF output.",
                    "Re-run; if it persists, open the document in LibreOffice manually to confirm it converts, " +
                    "or use --engine chromium.");
            }

            File.Move(produced, output, overwrite: true);
            return output;
        }
        finally
        {
            TryDeleteDirectory(workDir);
            TryDeleteDirectory(profileDir);
        }
    }

    /// <summary>
    /// Renders one page of the document to PNG: soffice → PDF, then
    /// <c>pdftoppm</c> rasterizes the 1-based <paramref name="page"/> at the
    /// requested pixel width. Returns the absolute path of the written PNG.
    /// </summary>
    public static string DocumentToPng(
        string sourceFile,
        string outPng,
        int page,
        int widthPx,
        SofficeInfo? soffice = null,
        TimeSpan? timeout = null)
    {
        var info = RequireSoffice(soffice);
        if (!info.Pdftoppm || info.PdftoppmPath is null)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "PNG rendering with the soffice engine needs 'pdftoppm' (poppler) and none was found.",
                NoPdftoppmSuggestion);
        }

        if (page < 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Page must be a positive 1-based index (got {page}).",
                "Pass --scope /slide[N] for a pptx, or omit it to render page 1.");
        }

        if (widthPx <= 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"PNG width must be positive (got {widthPx}).",
                "Pass a positive width or omit it for the default.");
        }

        var output = Path.GetFullPath(outPng);
        EnsureOutputDir(output);
        if (File.Exists(output))
        {
            File.Delete(output);
        }

        var workDir = Directory.CreateTempSubdirectory("aioffice-soffice-png-").FullName;
        try
        {
            var pdf = Path.Combine(workDir, "render.pdf");
            DocumentToPdf(sourceFile, pdf, info, timeout);

            // pdftoppm writes "<prefix>-<page>.png" (or "<prefix>.png" with -singlefile).
            // -singlefile drops the page suffix so we know the exact leaf name.
            var prefix = Path.Combine(workDir, "page");
            var args = BuildPdftoppmArgs(pdf, prefix, page, widthPx);
            RunTool(info.PdftoppmPath, args, timeout ?? DefaultTimeout, "PNG", pageOutOfRange: page);

            var produced = prefix + ".png";
            if (!File.Exists(produced) || new FileInfo(produced).Length == 0)
            {
                produced = Directory.EnumerateFiles(workDir, "page*.png").FirstOrDefault() ?? produced;
            }

            if (!File.Exists(produced) || new FileInfo(produced).Length == 0)
            {
                throw new AiofficeException(
                    ErrorCodes.InternalError,
                    $"pdftoppm reported success but produced no PNG for page {page}.",
                    "Check that the page exists (the deck/document may have fewer pages), or use --to pdf.");
            }

            File.Move(produced, output, overwrite: true);
            return output;
        }
        finally
        {
            TryDeleteDirectory(workDir);
        }
    }

    /// <summary>
    /// The soffice headless-convert argument list:
    /// <c>--headless --convert-to pdf --outdir &lt;dir&gt; -env:UserInstallation=&lt;profile&gt; &lt;file&gt;</c>.
    /// Public so the doctor command can show the command line it would run.
    /// </summary>
    public static IReadOnlyList<string> BuildConvertToPdfArgs(string sourceAbs, string outDir, string? profileDir = null)
    {
        var args = new List<string>
        {
            "--headless",
            "--norestore",
            "--convert-to",
            "pdf",
            "--outdir",
            outDir,
        };
        if (profileDir is not null)
        {
            // A private user profile avoids colliding with a running desktop
            // LibreOffice (which otherwise refuses headless conversion).
            args.Add("-env:UserInstallation=" + new Uri(Path.GetFullPath(profileDir)).AbsoluteUri);
        }

        args.Add(Path.GetFullPath(sourceAbs));
        return args;
    }

    /// <summary>
    /// The pdftoppm argument list for one page at a fixed width:
    /// <c>-png -singlefile -scale-to-x &lt;w&gt; -scale-to-y -1 -f &lt;page&gt; -l &lt;page&gt; &lt;pdf&gt; &lt;prefix&gt;</c>.
    /// Public so the doctor command can show the command line it would run.
    /// </summary>
    public static IReadOnlyList<string> BuildPdftoppmArgs(string pdfAbs, string outPrefix, int page, int widthPx)
    {
        var pageStr = page.ToString(CultureInfo.InvariantCulture);
        return
        [
            "-png",
            "-singlefile",
            "-scale-to-x",
            widthPx.ToString(CultureInfo.InvariantCulture),
            "-scale-to-y",
            "-1", // preserve aspect ratio
            "-f",
            pageStr,
            "-l",
            pageStr,
            Path.GetFullPath(pdfAbs),
            outPrefix,
        ];
    }

    private static string RequireSource(string sourceFile)
    {
        var source = Path.GetFullPath(sourceFile);
        if (!File.Exists(source))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"Document not found: {source}",
                "Pass the path of an existing document to render.");
        }

        return source;
    }

    private static SofficeInfo RequireSoffice(SofficeInfo? soffice)
    {
        var info = soffice ?? SofficeLocator.Probe();
        if (!info.Found || info.Path is null)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "The soffice render engine needs LibreOffice and none was found on this machine.",
                NoSofficeSuggestion);
        }

        return info;
    }

    private static void EnsureOutputDir(string output)
    {
        if (Path.GetDirectoryName(output) is { Length: > 0 } outDir)
        {
            Directory.CreateDirectory(outDir);
        }
    }

    private static void RunTool(
        string toolPath, IReadOnlyList<string> args, TimeSpan timeout, string what, int? pageOutOfRange = null)
    {
        var psi = new ProcessStartInfo(toolPath)
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
                $"Failed to start '{toolPath}': {ex.Message}",
                "Check that AIOFFICE_SOFFICE points to an executable LibreOffice binary, or unset it to auto-detect.",
                innerException: ex);
        }

        using (process)
        {
            var stderrTask = process.StandardError.ReadToEndAsync();
            var stdoutTask = process.StandardOutput.ReadToEndAsync();

            if (!process.WaitForExit((int)timeout.TotalMilliseconds))
            {
                KillTree(process);
                throw new AiofficeException(
                    ErrorCodes.InternalError,
                    $"LibreOffice/pdftoppm timed out after {timeout.TotalSeconds:0}s while rendering {what}.",
                    "Re-run; large documents can be slow. If it persists, use --engine chromium.");
            }

            process.WaitForExit(); // flush redirected streams
            _ = stdoutTask.Result;
            if (process.ExitCode != 0)
            {
                var detail = Tail(stderrTask.Result);

                // A pdftoppm "page range" failure means the requested slide/page
                // is past the end of the document — that is a user error, not ours.
                if (pageOutOfRange is { } requested &&
                    detail.Contains("page range", StringComparison.OrdinalIgnoreCase))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Page {requested} is past the end of the document.",
                        "Pass a --scope /slide[N] within the deck, or render --to pdf for the whole document.");
                }

                var suffix = detail.Length > 0 ? $" stderr: {detail}" : string.Empty;
                throw new AiofficeException(
                    ErrorCodes.InternalError,
                    $"LibreOffice/pdftoppm exited with code {process.ExitCode} while rendering {what}.{suffix}",
                    "Re-run; if it persists, open the document in LibreOffice manually, or use --engine chromium.");
            }
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
