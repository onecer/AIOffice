using Xunit;

namespace AIOffice.Render.Tests;

/// <summary>
/// The one end-to-end test against a REAL Chromium browser. Skipped (not
/// failed) on machines without one; everything else in this suite runs on
/// stub scripts.
/// </summary>
public sealed class RealBrowserTests : IDisposable
{
    private static readonly byte[] PngMagic = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    private readonly TempDir _tmp = new();

    public void Dispose() => _tmp.Dispose();

    [BrowserFact]
    public void Html_renders_to_a_real_png_file()
    {
        var browser = BrowserLocator.Probe(refresh: true);
        var output = PngRenderer.HtmlStringToPng(
            "<h1 style=\"font-family:sans-serif\">AIOffice render smoke test</h1>",
            _tmp.PathOf("smoke.png"),
            widthPx: 640,
            heightPx: 400,
            browser: browser);

        var bytes = File.ReadAllBytes(output);
        Assert.True(bytes.Length > PngMagic.Length, "PNG output is empty");
        Assert.Equal(PngMagic, bytes[..PngMagic.Length]);
    }
}
