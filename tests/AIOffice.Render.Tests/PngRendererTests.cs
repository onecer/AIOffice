using AIOffice.Core;
using Xunit;

namespace AIOffice.Render.Tests;

public sealed class PngRendererTests : IDisposable
{
    private readonly TempDir _tmp = new();

    public void Dispose() => _tmp.Dispose();

    private BrowserInfo Stub(string body, string name = "stub-browser") =>
        new(true, _tmp.WriteStubBrowser(name, body), "custom");

    // ------------------------------------------------------- argument contract

    [Fact]
    public void Screenshot_args_follow_the_headless_contract_with_url_last()
    {
        var html = _tmp.WriteHtml();
        var output = _tmp.PathOf("out.png");

        var args = PngRenderer.BuildScreenshotArgs(html, output, 1280, 720, profileDir: "/tmp/profile");

        Assert.Equal("--headless=new", args[0]);
        Assert.Contains($"--screenshot={output}", args);
        Assert.Contains("--window-size=1280,720", args);
        Assert.Contains("--hide-scrollbars", args);
        Assert.Contains("--disable-gpu", args);
        Assert.Contains("--user-data-dir=/tmp/profile", args);
        Assert.StartsWith("file://", args[^1]);
        Assert.EndsWith("input.html", args[^1]);
    }

    [Fact]
    public void Auto_height_uses_the_documented_default_viewport()
    {
        var args = PngRenderer.BuildScreenshotArgs(
            _tmp.WriteHtml(), _tmp.PathOf("out.png"), 800, PngRenderer.DefaultHeightPx);
        Assert.Contains($"--window-size=800,{PngRenderer.DefaultHeightPx}", args);
    }

    // ------------------------------------------------------------ failure map

    [Fact]
    public void No_browser_maps_to_unsupported_feature_with_the_workaround()
    {
        var ex = Assert.Throws<AiofficeException>(() =>
            PngRenderer.HtmlFileToPng(_tmp.WriteHtml(), _tmp.PathOf("out.png"), browser: BrowserInfo.NotFound));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Equal(
            "Install Google Chrome/Edge, set AIOFFICE_BROWSER to a Chromium binary, or use --to svg|html instead.",
            ex.Suggestion);
    }

    [Fact]
    public void No_browser_failure_maps_onto_the_envelope_and_exit_code()
    {
        var ex = Assert.Throws<AiofficeException>(() =>
            PngRenderer.HtmlFileToPng(_tmp.WriteHtml(), _tmp.PathOf("out.png"), browser: BrowserInfo.NotFound));

        var envelope = Envelope.FromException(ex);
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(envelope.Error.Suggestion));
        Assert.Equal(ExitCodes.UnsupportedFeature, envelope.ExitCode);
    }

    [Fact]
    public void Missing_html_input_maps_to_file_not_found()
    {
        var ex = Assert.Throws<AiofficeException>(() =>
            PngRenderer.HtmlFileToPng(_tmp.PathOf("missing.html"), _tmp.PathOf("out.png"), browser: Stub("exit 0")));

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

        var browser = Stub("echo 'boom from stub' >&2\nexit 3");
        var ex = Assert.Throws<AiofficeException>(() =>
            PngRenderer.HtmlFileToPng(_tmp.WriteHtml(), _tmp.PathOf("out.png"), browser: browser));

        Assert.Equal(ErrorCodes.InternalError, ex.Code);
        Assert.Contains("exited with code 3", ex.Message);
        Assert.Contains("boom from stub", ex.Message);
        Assert.Contains("Re-run", ex.Suggestion);
    }

    [Fact]
    public void Success_without_output_maps_to_internal_error()
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        var ex = Assert.Throws<AiofficeException>(() =>
            PngRenderer.HtmlFileToPng(_tmp.WriteHtml(), _tmp.PathOf("out.png"), browser: Stub("exit 0")));

        Assert.Equal(ErrorCodes.InternalError, ex.Code);
        Assert.Contains("no PNG output", ex.Message);
    }

    [Fact]
    public void Timeout_kills_the_browser_and_maps_to_internal_error()
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        var ex = Assert.Throws<AiofficeException>(() =>
            PngRenderer.HtmlFileToPng(
                _tmp.WriteHtml(), _tmp.PathOf("out.png"),
                browser: Stub("sleep 10"), timeout: TimeSpan.FromSeconds(1)));

        Assert.Equal(ErrorCodes.InternalError, ex.Code);
        Assert.Contains("timed out", ex.Message);
    }

    [Fact]
    public void Unstartable_browser_binary_maps_to_internal_error()
    {
        var notExecutable = _tmp.PathOf("not-a-browser.txt");
        File.WriteAllText(notExecutable, "plain text");

        var ex = Assert.Throws<AiofficeException>(() =>
            PngRenderer.HtmlFileToPng(
                _tmp.WriteHtml(), _tmp.PathOf("out.png"),
                browser: new BrowserInfo(true, notExecutable, "custom")));

        Assert.Equal(ErrorCodes.InternalError, ex.Code);
        Assert.Contains("AIOFFICE_BROWSER", ex.Suggestion);
    }

    [Fact]
    public void Svg_input_without_svg_markup_maps_to_invalid_args()
    {
        var ex = Assert.Throws<AiofficeException>(() =>
            PngRenderer.SvgToPng("<p>not svg</p>", _tmp.PathOf("out.png"), browser: Stub("exit 0")));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Zero_width_maps_to_invalid_args()
    {
        var ex = Assert.Throws<AiofficeException>(() =>
            PngRenderer.HtmlFileToPng(_tmp.WriteHtml(), _tmp.PathOf("out.png"), widthPx: 0, browser: Stub("exit 0")));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    // ------------------------------------------------------------- happy path

    [Fact]
    public void Stub_browser_that_writes_the_screenshot_file_succeeds()
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        var browser = Stub(
            """
            for a in "$@"; do
              case "$a" in
                --screenshot=*) printf 'FAKEPNG' > "${a#--screenshot=}" ;;
              esac
            done
            exit 0
            """);

        var written = PngRenderer.HtmlFileToPng(_tmp.WriteHtml(), _tmp.PathOf("out.png"), browser: browser);

        Assert.Equal(Path.GetFullPath(_tmp.PathOf("out.png")), written);
        Assert.Equal("FAKEPNG", File.ReadAllText(written));
    }

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
                --screenshot=*) printf 'FAKEPNG' > "${a#--screenshot=}" ;;
                file://*) printf '%s' "$a" > "$(dirname "$0")/seen-url.txt" ;;
              esac
            done
            exit 0
            """);

        var written = PngRenderer.HtmlStringToPng("<h1>hi</h1>", _tmp.PathOf("out.png"), browser: browser);

        Assert.Equal("FAKEPNG", File.ReadAllText(written));
        var seenUrl = File.ReadAllText(_tmp.PathOf("seen-url.txt"));
        Assert.StartsWith("file://", seenUrl);
        Assert.EndsWith(".html", seenUrl);
    }
}
