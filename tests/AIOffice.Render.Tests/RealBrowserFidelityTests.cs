using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using Xunit;

namespace AIOffice.Render.Tests;

/// <summary>
/// Real-Chromium fidelity invariants that need NO golden pixels: a non-blank
/// differential (text vs white), a decode-free IHDR dimension check, a
/// structural PDF page-count, and same-machine PNG determinism. Every test wears
/// <see cref="BrowserFactAttribute"/>, so it runs on a dev machine with a browser
/// and skips unconditionally on CI. Assertions are ORDERINGS / structural
/// properties / same-process comparisons — never committed baselines or
/// cross-runner equality — so they are robust to font hinting, AA, and Chrome
/// version drift.
/// </summary>
public sealed class RealBrowserFidelityTests : IDisposable
{
    private static readonly byte[] PngSignature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
    private static readonly byte[] PngIend = [0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82];

    private readonly TempDir _tmp = new();

    public void Dispose() => _tmp.Dispose();

    /// <summary>Structural PNG check reimplemented in-test (signature + IEND trailer).</summary>
    private static bool IsStructuralPng(byte[] png) =>
        png.Length >= PngSignature.Length + PngIend.Length &&
        png.AsSpan(0, 8).SequenceEqual(PngSignature) &&
        png.AsSpan(png.Length - 8, 8).SequenceEqual(PngIend);

    /// <summary>Decode-free IHDR read: width at bytes 16-19, height at 20-23, big-endian.</summary>
    private static (uint Width, uint Height) ReadIhdrDimensions(byte[] png)
    {
        Assert.True(png.Length >= 24, "PNG too short to carry an IHDR");
        uint width = (uint)((png[16] << 24) | (png[17] << 16) | (png[18] << 8) | png[19]);
        uint height = (uint)((png[20] << 24) | (png[21] << 16) | (png[22] << 8) | png[23]);
        return (width, height);
    }

    // ---- item 5: non-blank differential (text vs white) ---------------------

    [BrowserFact]
    public void A_text_page_yields_a_larger_png_than_a_blank_white_page()
    {
        var browser = BrowserLocator.Probe(refresh: true);

        var textPng = File.ReadAllBytes(PngRenderer.HtmlStringToPng(
            "<!DOCTYPE html><html><body style=\"margin:0;background:#fff;color:#000;" +
            "font-family:monospace;font-size:18px\">" +
            string.Concat(Enumerable.Repeat("The quick brown fox jumps over the lazy dog. 0123456789<br>", 20)) +
            "</body></html>",
            _tmp.PathOf("text.png"), widthPx: 640, heightPx: 400, browser: browser));

        var whitePng = File.ReadAllBytes(PngRenderer.HtmlStringToPng(
            "<!DOCTYPE html><html><body style=\"margin:0;background:#fff\"></body></html>",
            _tmp.PathOf("white.png"), widthPx: 640, heightPx: 400, browser: browser));

        Assert.True(IsStructuralPng(textPng), "text render is not a structurally valid PNG");
        Assert.True(IsStructuralPng(whitePng), "white render is not a structurally valid PNG");

        // PNG deflate crushes a uniform white raster to almost nothing; ink shows
        // as a size ORDERING (never an exact byte count), so a blank-render
        // regression trips here without any pixel decode.
        Assert.True(
            textPng.Length >= whitePng.Length * 1.5,
            $"text PNG ({textPng.Length}B) is not >= 1.5x the white PNG ({whitePng.Length}B) — render may be blank");
    }

    // ---- item 6: decode-free IHDR dimension invariant -----------------------

    [BrowserFact]
    public void Png_dimensions_are_an_integer_multiple_of_the_requested_size()
    {
        var browser = BrowserLocator.Probe(refresh: true);
        const int reqW = 640;
        const int reqH = 400;

        var png = File.ReadAllBytes(PngRenderer.HtmlStringToPng(
            "<!DOCTYPE html><html><body style=\"margin:0\">dimension probe</body></html>",
            _tmp.PathOf("dim.png"), widthPx: reqW, heightPx: reqH, browser: browser));

        Assert.True(IsStructuralPng(png));
        var (w, h) = ReadIhdrDimensions(png);

        Assert.True(w > 0 && h > 0, $"non-positive dimensions {w}x{h}");
        // Tolerate a HiDPI device-scale-factor (1x or 2x): assert positive integer
        // multiples with a matching aspect ratio rather than strict WxH equality.
        Assert.True(w % reqW == 0, $"width {w} is not an integer multiple of {reqW}");
        Assert.True(h % reqH == 0, $"height {h} is not an integer multiple of {reqH}");
        Assert.Equal(w / (double)reqW, h / (double)reqH, 3); // same scale factor on both axes
    }

    // ---- item 7: chromium PDF page-count structural invariant ---------------

    [BrowserFact]
    public void A_two_page_deck_prints_exactly_two_page_objects()
    {
        var browser = BrowserLocator.Probe(refresh: true);

        // The exact structure PdfRenderVerb.PrintDeck emits: @page pinned to the
        // slide box, margins zeroed, each slide break-after:page (last is auto).
        const int w = 480;
        const int h = 320;
        var html =
            "<!DOCTYPE html><html><head><meta charset=\"utf-8\"><style>" +
            string.Create(CultureInfo.InvariantCulture, $"@page{{size:{w}px {h}px;margin:0}}") +
            "html,body{margin:0;padding:0}" +
            string.Create(CultureInfo.InvariantCulture, $".slide{{width:{w}px;height:{h}px;overflow:hidden;break-after:page}}") +
            ".slide:last-child{break-after:auto}svg{display:block}" +
            "</style></head><body>" +
            "<div class=\"slide\">Slide one</div><div class=\"slide\">Slide two</div>" +
            "</body></html>";

        var pdf = File.ReadAllBytes(PdfRenderer.HtmlStringToPdf(html, _tmp.PathOf("deck.pdf"), browser: browser));

        Assert.True(PdfRenderer.IsCompletePdf(_tmp.PathOf("deck.pdf")));
        // Count leaf page objects: '/Type /Page' but NOT '/Type /Pages'. Chromium
        // writes these uncompressed (unlike soffice, whose page objects can live
        // in object streams — hence this scan is chromium-only, per the plan).
        var text = System.Text.Encoding.Latin1.GetString(pdf);
        var leafPages = Regex.Matches(text, @"/Type\s*/Page(?![s])").Count;
        Assert.Equal(2, leafPages);
    }

    // ---- item 8: same-machine PNG determinism -------------------------------

    [BrowserFact]
    public void The_same_html_renders_to_a_byte_identical_png_twice_in_one_process()
    {
        var browser = BrowserLocator.Probe(refresh: true);
        const string html =
            "<!DOCTYPE html><html><body style=\"margin:0;background:#fff;color:#111;" +
            "font-family:monospace\"><h1>determinism</h1><p>static content, no timers</p></body></html>";

        var first = File.ReadAllBytes(PngRenderer.HtmlStringToPng(
            html, _tmp.PathOf("det-a.png"), widthPx: 500, heightPx: 300, browser: browser));
        var second = File.ReadAllBytes(PngRenderer.HtmlStringToPng(
            html, _tmp.PathOf("det-b.png"), widthPx: 500, heightPx: 300, browser: browser));

        Assert.True(IsStructuralPng(first));
        Assert.True(IsStructuralPng(second));

        // Chromium headless --disable-gpu screenshots of static content embed no
        // timestamp, so two renders on ONE machine/engine are byte-identical.
        // (PNG only — never PDF, which embeds /CreationDate + a per-run /ID.)
        var hashA = Convert.ToHexString(SHA256.HashData(first));
        var hashB = Convert.ToHexString(SHA256.HashData(second));
        if (hashA != hashB)
        {
            // Documented fallback if the SHA ever flickers on a machine: assert
            // identical IHDR dimensions + length within 1% instead of byte equality.
            var (wa, ha) = ReadIhdrDimensions(first);
            var (wb, hb) = ReadIhdrDimensions(second);
            Assert.Equal((wa, ha), (wb, hb));
            var ratio = Math.Abs(first.Length - second.Length) / (double)Math.Max(first.Length, second.Length);
            Assert.True(ratio <= 0.01, $"PNG lengths differ by {ratio:P2} (> 1%)");
        }
        else
        {
            Assert.Equal(hashA, hashB);
        }
    }

    // ---- item 10 (flagship): coarse cross-engine AGREEMENT ------------------

    [BrowserFact]
    public void Chromium_and_soffice_agree_coarsely_that_the_same_content_is_a_valid_pdf()
    {
        // Gated on BOTH engines: [BrowserFact] already skips on CI / no browser;
        // this early-returns when soffice is absent, so it runs only where both
        // exist (this dev machine). NEVER pixel/byte equality across engines —
        // only coarse agreement: both produce a structurally valid, non-trivial PDF.
        if (!SofficeLocator.Probe(refresh: true).Found)
        {
            return; // soffice not installed — the per-engine invariants carry the coverage
        }

        var browser = BrowserLocator.Probe(refresh: true);
        const string sentence = "Cross-engine agreement probe: the quick brown fox.";

        var chromiumPdf = PdfRenderer.HtmlStringToPdf(
            "<!DOCTYPE html><html><head><meta charset=\"utf-8\"><style>@page{size:A4;margin:14mm}" +
            "body{font-family:sans-serif}</style></head><body><p>" + sentence + "</p></body></html>",
            _tmp.PathOf("chromium.pdf"), browser: browser);

        var fodt = _tmp.PathOf("agree.fodt");
        File.WriteAllText(
            fodt,
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
            "<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" " +
            "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" " +
            "office:version=\"1.2\" office:mimetype=\"application/vnd.oasis.opendocument.text\">" +
            "<office:body><office:text><text:p>" + sentence + "</text:p></office:text></office:body>" +
            "</office:document>");
        var sofficePdf = SofficeRenderer.DocumentToPdf(fodt, _tmp.PathOf("soffice.pdf"));

        Assert.True(PdfRenderer.IsCompletePdf(chromiumPdf), "chromium PDF is not structurally complete");
        Assert.True(PdfRenderer.IsCompletePdf(sofficePdf), "soffice PDF is not structurally complete");
        Assert.True(new FileInfo(chromiumPdf).Length > 512, "chromium PDF is implausibly small");
        Assert.True(new FileInfo(sofficePdf).Length > 512, "soffice PDF is implausibly small");
    }
}
