using AIOffice.Core;
using Xunit;

namespace AIOffice.Render.Tests;

public sealed class PdfRendererTests : IDisposable
{
    private readonly TempDir _tmp = new();

    public void Dispose() => _tmp.Dispose();

    private BrowserInfo Stub(string body, string name = "stub-browser") =>
        new(true, _tmp.WriteStubBrowser(name, body), "custom");

    // ------------------------------------------------------- argument contract

    [Fact]
    public void Print_args_follow_the_headless_contract_with_url_last()
    {
        var html = _tmp.WriteHtml();
        var output = _tmp.PathOf("out.pdf");

        var args = PdfRenderer.BuildPrintToPdfArgs(html, output, profileDir: "/tmp/profile");

        Assert.Equal("--headless=new", args[0]);
        Assert.Contains($"--print-to-pdf={output}", args);
        Assert.Contains("--no-pdf-header-footer", args);
        Assert.Contains("--disable-gpu", args);
        Assert.Contains("--user-data-dir=/tmp/profile", args);
        Assert.StartsWith("file://", args[^1]);
        Assert.EndsWith("input.html", args[^1]);
    }

    // ------------------------------------------------------------ failure map

    [Fact]
    public void No_browser_maps_to_unsupported_feature_with_the_workaround()
    {
        var ex = Assert.Throws<AiofficeException>(() =>
            PdfRenderer.HtmlFileToPdf(_tmp.WriteHtml(), _tmp.PathOf("out.pdf"), browser: BrowserInfo.NotFound));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("AIOFFICE_BROWSER", ex.Suggestion);
        Assert.Contains("--to html|svg", ex.Suggestion);

        var envelope = Envelope.FromException(ex);
        Assert.Equal(ExitCodes.UnsupportedFeature, envelope.ExitCode);
    }

    [Fact]
    public void Missing_html_input_maps_to_file_not_found()
    {
        var ex = Assert.Throws<AiofficeException>(() =>
            PdfRenderer.HtmlFileToPdf(_tmp.PathOf("missing.html"), _tmp.PathOf("out.pdf"), browser: Stub("exit 0")));

        Assert.Equal(ErrorCodes.FileNotFound, ex.Code);
        Assert.False(string.IsNullOrWhiteSpace(ex.Suggestion));
    }

    [Fact]
    public void Nonzero_exit_maps_to_internal_error_carrying_the_stderr_tail()
    {
        if (OperatingSystem.IsWindows())
        {
            return; // stub browsers are shell scripts
        }

        var browser = Stub("echo 'pdf boom' >&2\nexit 5");
        var ex = Assert.Throws<AiofficeException>(() =>
            PdfRenderer.HtmlFileToPdf(_tmp.WriteHtml(), _tmp.PathOf("out.pdf"), browser: browser));

        Assert.Equal(ErrorCodes.InternalError, ex.Code);
        Assert.Contains("exited with code 5", ex.Message);
        Assert.Contains("pdf boom", ex.Message);
    }

    [Fact]
    public void Success_without_output_maps_to_internal_error()
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        var ex = Assert.Throws<AiofficeException>(() =>
            PdfRenderer.HtmlFileToPdf(_tmp.WriteHtml(), _tmp.PathOf("out.pdf"), browser: Stub("exit 0")));

        Assert.Equal(ErrorCodes.InternalError, ex.Code);
        Assert.Contains("no PDF output", ex.Message);
    }

    [Fact]
    public void Timeout_kills_the_browser_and_maps_to_internal_error()
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        var ex = Assert.Throws<AiofficeException>(() =>
            PdfRenderer.HtmlFileToPdf(
                _tmp.WriteHtml(), _tmp.PathOf("out.pdf"),
                browser: Stub("sleep 10"), timeout: TimeSpan.FromSeconds(1)));

        Assert.Equal(ErrorCodes.InternalError, ex.Code);
        Assert.Contains("timed out", ex.Message);
    }

    // ------------------------------------------------- completeness detection

    [Fact]
    public void A_lingering_browser_with_a_complete_pdf_counts_as_success()
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        // Writes a structurally complete PDF, then never exits: the renderer
        // must accept the file and kill the straggler instead of timing out.
        var browser = Stub(
            """
            for a in "$@"; do
              case "$a" in
                --print-to-pdf=*) printf '%%PDF-1.7 fake body %%%%EOF' > "${a#--print-to-pdf=}" ;;
              esac
            done
            sleep 30
            """);

        var written = PdfRenderer.HtmlFileToPdf(
            _tmp.WriteHtml(), _tmp.PathOf("out.pdf"),
            browser: browser, timeout: TimeSpan.FromSeconds(10));

        Assert.StartsWith("%PDF-", File.ReadAllText(written));
    }

    [Fact]
    public void Pdf_completeness_requires_signature_and_eof_trailer()
    {
        var path = _tmp.PathOf("probe.pdf");

        File.WriteAllText(path, "%PDF-1.7 body %%EOF");
        Assert.True(PdfRenderer.IsCompletePdf(path));

        File.WriteAllText(path, "%PDF-1.7 body without trailer");
        Assert.False(PdfRenderer.IsCompletePdf(path));

        File.WriteAllText(path, "not a pdf %%EOF");
        Assert.False(PdfRenderer.IsCompletePdf(path));
    }

    // ------------------------------------------------------------- happy path

    [Fact]
    public void Html_string_round_trips_through_a_temp_file()
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        var browser = Stub(
            """
            for a in "$@"; do
              case "$a" in
                --print-to-pdf=*) printf '%%PDF-1.7 stub %%%%EOF' > "${a#--print-to-pdf=}" ;;
                file://*) printf '%s' "$a" > "$(dirname "$0")/seen-url.txt" ;;
              esac
            done
            exit 0
            """);

        var written = PdfRenderer.HtmlStringToPdf("<h1>hi</h1>", _tmp.PathOf("out.pdf"), browser: browser);

        Assert.StartsWith("%PDF-", File.ReadAllText(written));
        var seenUrl = File.ReadAllText(_tmp.PathOf("seen-url.txt"));
        Assert.StartsWith("file://", seenUrl);
        Assert.EndsWith(".html", seenUrl);
    }
}

/// <summary>End-to-end against a REAL Chromium browser; skipped when none is installed.</summary>
public sealed class RealBrowserPdfTests : IDisposable
{
    private readonly TempDir _tmp = new();

    public void Dispose() => _tmp.Dispose();

    [BrowserFact]
    public void Html_prints_to_a_real_pdf_file()
    {
        var browser = BrowserLocator.Probe(refresh: true);
        var output = PdfRenderer.HtmlStringToPdf(
            "<style>@page{size:A4;margin:14mm}</style><h1>AIOffice pdf smoke test</h1>",
            _tmp.PathOf("smoke.pdf"),
            browser: browser);

        var bytes = File.ReadAllBytes(output);
        Assert.True(bytes.Length > 100, "PDF output is suspiciously small");
        Assert.Equal("%PDF-"u8.ToArray(), bytes[..5]);
    }
}
