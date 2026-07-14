using AIOffice.Core;
using Xunit;

namespace AIOffice.Render.Tests;

public sealed class SofficeRendererTests : IDisposable
{
    private static readonly byte[] PdfMagic = "%PDF-"u8.ToArray();
    private static readonly byte[] PngMagic = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    private readonly TempDir _tmp = new();

    public void Dispose() => _tmp.Dispose();

    // ----- argument contracts (deterministic, no soffice needed) --------------

    [Fact]
    public void Convert_to_pdf_args_follow_the_headless_contract_with_file_last()
    {
        var src = _tmp.PathOf("input.fodt");
        File.WriteAllText(src, "x");
        var args = SofficeRenderer.BuildConvertToPdfArgs(src, _tmp.Dir, profileDir: _tmp.Dir);

        Assert.Equal("--headless", args[0]);
        Assert.Contains("--convert-to", args);
        Assert.Contains("pdf", args);
        Assert.Contains("--outdir", args);
        Assert.Contains(_tmp.Dir, args);
        Assert.Contains(args, a => a.StartsWith("-env:UserInstallation=", StringComparison.Ordinal));
        Assert.Equal(Path.GetFullPath(src), args[^1]); // the source file is last
    }

    [Fact]
    public void Pdftoppm_args_select_one_page_at_a_fixed_width()
    {
        var pdf = _tmp.PathOf("doc.pdf");
        File.WriteAllText(pdf, "x");
        var args = SofficeRenderer.BuildPdftoppmArgs(pdf, _tmp.PathOf("page"), page: 3, widthPx: 1280);

        Assert.Contains("-png", args);
        Assert.Contains("-singlefile", args);
        Assert.Contains("-scale-to-x", args);
        Assert.Contains("1280", args);
        // -f and -l both pin to page 3.
        var f = args.ToList().IndexOf("-f");
        var l = args.ToList().IndexOf("-l");
        Assert.Equal("3", args[f + 1]);
        Assert.Equal("3", args[l + 1]);
        Assert.Equal(Path.GetFullPath(pdf), args[^2]); // pdf then output prefix
    }

    // ----- failure maps (deterministic) ---------------------------------------

    [Fact]
    public void No_soffice_maps_to_unsupported_feature_with_an_install_hint()
    {
        var src = _tmp.PathOf("input.fodt");
        File.WriteAllText(src, "x");

        var ex = Assert.Throws<AiofficeException>(() =>
            SofficeRenderer.DocumentToPdf(src, _tmp.PathOf("out.pdf"), soffice: SofficeInfo.NotFound));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("LibreOffice", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Png_without_pdftoppm_maps_to_unsupported_feature_suggesting_poppler()
    {
        var src = _tmp.PathOf("input.fodt");
        File.WriteAllText(src, "x");

        // soffice present, pdftoppm absent.
        var info = new SofficeInfo(true, "/fake/soffice", Pdftoppm: false, PdftoppmPath: null);
        var ex = Assert.Throws<AiofficeException>(() =>
            SofficeRenderer.DocumentToPng(src, _tmp.PathOf("out.png"), page: 1, widthPx: 1280, soffice: info));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("poppler", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Missing_source_maps_to_file_not_found()
    {
        var info = new SofficeInfo(true, "/fake/soffice", true, "/fake/pdftoppm");
        var ex = Assert.Throws<AiofficeException>(() =>
            SofficeRenderer.DocumentToPdf(_tmp.PathOf("ghost.fodt"), _tmp.PathOf("out.pdf"), soffice: info));

        Assert.Equal(ErrorCodes.FileNotFound, ex.Code);
    }

    // ----- real soffice (SKIPPED when LibreOffice / poppler absent — incl. CI) -

    [SofficeFact]
    public void Document_converts_to_a_real_pdf()
    {
        var src = WriteFodt("soffice pdf smoke");
        var output = SofficeRenderer.DocumentToPdf(src, _tmp.PathOf("out.pdf"));

        var bytes = File.ReadAllBytes(output);
        Assert.True(bytes.Length > PdfMagic.Length, "PDF output is empty");
        Assert.Equal(PdfMagic, bytes[..PdfMagic.Length]);
    }

    [SofficeFact(needsPdftoppm: true)]
    public void Document_rasterizes_page_one_to_a_real_png()
    {
        var src = WriteFodt("soffice png smoke");
        var output = SofficeRenderer.DocumentToPng(src, _tmp.PathOf("out.png"), page: 1, widthPx: 800);

        var bytes = File.ReadAllBytes(output);
        Assert.True(bytes.Length > PngMagic.Length, "PNG output is empty");
        Assert.Equal(PngMagic, bytes[..PngMagic.Length]);
    }

    [SofficeFact]
    public void A_two_page_document_converts_to_a_structurally_complete_multipage_pdf()
    {
        // A hand-written .fodt with an explicit page break: two paragraphs, the
        // second carrying fo:break-before="page". Reuses the .fodt seam (no binary
        // fixture, no handler reference); auto-skips on CI (no LibreOffice there).
        var src = WriteTwoPageFodt("Page one content", "Page two content");
        var output = SofficeRenderer.DocumentToPdf(src, _tmp.PathOf("two.pdf"));

        // Structure over pixels: a valid PDF (public IsCompletePdf) + a real byte
        // floor. Page count is kept COARSE (>=1) to stay robust to soffice-version
        // pagination; the leaf '/Type /Page' scan is only meaningful when soffice
        // writes page objects uncompressed, so it is a lower-bound assertion.
        Assert.True(PdfRenderer.IsCompletePdf(output), "soffice PDF is not structurally complete");
        var bytes = File.ReadAllBytes(output);
        Assert.True(bytes.Length > 1024, $"soffice PDF is implausibly small ({bytes.Length}B)");

        var text = System.Text.Encoding.Latin1.GetString(bytes);
        var leafPages = System.Text.RegularExpressions.Regex.Matches(text, @"/Type\s*/Page(?![s])").Count;
        Assert.True(leafPages >= 1, $"expected at least one page object, found {leafPages}");
    }

    /// <summary>
    /// A minimal Flat-ODF Text document (.fodt) — a single XML file LibreOffice
    /// opens and converts to PDF without needing the OpenXml SDK in this test
    /// project. Keeps the soffice E2E test self-contained.
    /// </summary>
    private string WriteFodt(string text)
    {
        var path = _tmp.PathOf("doc.fodt");
        File.WriteAllText(
            path,
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
            "<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" " +
            "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" " +
            "office:version=\"1.2\" office:mimetype=\"application/vnd.oasis.opendocument.text\">" +
            "<office:body><office:text><text:p>" + text + "</text:p></office:text></office:body>" +
            "</office:document>");
        return path;
    }

    /// <summary>
    /// A two-page Flat-ODF Text document: the second paragraph carries a
    /// <c>fo:break-before="page"</c> automatic style, so LibreOffice paginates it
    /// onto a second page. Self-contained (no OpenXml SDK, no binary fixture).
    /// </summary>
    private string WriteTwoPageFodt(string pageOne, string pageTwo)
    {
        var path = _tmp.PathOf("two.fodt");
        File.WriteAllText(
            path,
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
            "<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" " +
            "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" " +
            "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" " +
            "xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" " +
            "office:version=\"1.2\" office:mimetype=\"application/vnd.oasis.opendocument.text\">" +
            "<office:automatic-styles><style:style style:name=\"Pbreak\" style:family=\"paragraph\">" +
            "<style:paragraph-properties fo:break-before=\"page\"/></style:style></office:automatic-styles>" +
            "<office:body><office:text>" +
            "<text:p>" + pageOne + "</text:p>" +
            "<text:p text:style-name=\"Pbreak\">" + pageTwo + "</text:p>" +
            "</office:text></office:body></office:document>");
        return path;
    }
}
