using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Render.Tests;

/// <summary>
/// Drives <see cref="PngRenderVerb.Execute"/> / <see cref="PdfRenderVerb.Execute"/>
/// end-to-end against a POSIX stub browser (pointed to via <c>AIOFFICE_BROWSER</c>),
/// so the cross-format verb ORCHESTRATION is asserted deterministically without a
/// real Chromium: pages==slideCount, scope_defaulted, default <c>-o</c> naming,
/// sizeBytes floor, and non-ok pass-through. The stub is a <c>#!/bin/sh</c> script,
/// so the whole class no-ops on Windows (the assembly is already
/// <c>DisableTestParallelization</c>, so mutating the process-wide probe cache is
/// safe under <see cref="ProbeCacheReset"/> + <see cref="EnvVarScope"/>).
/// </summary>
public sealed class VerbOrchestrationTests : IDisposable
{
    private readonly TempDir _tmp = new();

    public void Dispose() => _tmp.Dispose();

    private const string SlideSvg =
        "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"960\" height=\"540\">" +
        "<rect width=\"960\" height=\"540\" fill=\"#ffffff\"/></svg>";

    /// <summary>A pptx-shaped stub returning <paramref name="slideCount"/> slides.</summary>
    private static StubHandler DeckHandler(int slideCount) =>
        new(DocumentKind.Pptx, _ => Envelope.Ok(new
        {
            slides = Enumerable.Range(1, slideCount)
                .Select(i => new { path = $"/slide[{i}]", svg = SlideSvg })
                .ToArray(),
        }));

    /// <summary>A stub browser that fulfils the --screenshot / --print-to-pdf contract.</summary>
    private string StubBrowser() => _tmp.WriteStubBrowser(
        "stub-browser",
        """
        for a in "$@"; do
          case "$a" in
            --screenshot=*) printf 'FAKEPNG-BYTES-0123456789' > "${a#--screenshot=}" ;;
            --print-to-pdf=*) { echo '%PDF-1.4'; echo 'stub deck body'; echo '%%EOF'; } > "${a#--print-to-pdf=}" ;;
          esac
        done
        exit 0
        """);

    private CommandContext Ctx(IFormatHandler handler, string? scope = null, string? output = null)
    {
        var args = new JsonObject();
        if (scope is not null)
        {
            args["scope"] = scope;
        }

        if (output is not null)
        {
            args["output"] = output;
        }

        var ext = handler.Kind == DocumentKind.Pptx ? ".pptx" : ".docx";
        return new CommandContext
        {
            Workspace = new Workspace(_tmp.Dir),
            File = _tmp.PathOf("deck" + ext),
            Args = args,
        };
    }

    private static JsonObject Data(Envelope env) =>
        JsonNode.Parse(env.ToJson())!["data"]!.AsObject();

    private static string[] WarningCodes(Envelope env) =>
        env.Meta.Warnings is null ? [] : [.. env.Meta.Warnings.Select(w => w.Code)];

    /// <summary>Points the browser probe at the stub for the scope of the action.</summary>
    private IDisposable UseStubBrowser()
    {
        var reset = new ProbeCacheReset();
        var env = new EnvVarScope(BrowserLocator.EnvVar, StubBrowser());
        BrowserLocator.Probe(refresh: true); // prime the process cache off the stub
        return new Combined(env, reset);
    }

    private sealed class Combined(IDisposable a, IDisposable b) : IDisposable
    {
        public void Dispose()
        {
            a.Dispose();
            b.Dispose();
        }
    }

    // ---- pdf orchestration ---------------------------------------------------

    [Fact]
    public void Pdf_deck_reports_pages_equal_to_slide_count_and_the_default_output_path()
    {
        if (OperatingSystem.IsWindows())
        {
            return; // stub browser is a shell script
        }

        var handler = DeckHandler(3);
        using (UseStubBrowser())
        {
            var env = PdfRenderVerb.Execute(handler, Ctx(handler));

            Assert.True(env.IsOk, env.ToJson());
            var data = Data(env);
            Assert.Equal(3, data["pages"]!.GetValue<int>()); // pages == slides.Count, pre-render
            Assert.Equal("deck.pdf", Path.GetFileName(data["written"]!.GetValue<string>()));
            Assert.True(data["sizeBytes"]!.GetValue<long>() > 0);
            Assert.True(File.Exists(data["written"]!.GetValue<string>()));
        }
    }

    [Fact]
    public void Pdf_deck_honors_an_explicit_output_path()
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        var handler = DeckHandler(2);
        var outPdf = _tmp.PathOf("custom.pdf");
        using (UseStubBrowser())
        {
            var env = PdfRenderVerb.Execute(handler, Ctx(handler, output: outPdf));

            Assert.Equal(outPdf, Data(env)["written"]!.GetValue<string>());
        }
    }

    [Fact]
    public void Pdf_docx_leaves_pages_null_and_defaults_the_output_name()
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        // A native (html) handler: PdfRenderVerb deliberately reports pages=null
        // (the browser paginator is content-dependent), but still writes the file.
        var handler = new StubHandler(DocumentKind.Docx);
        using (UseStubBrowser())
        {
            var env = PdfRenderVerb.Execute(handler, Ctx(handler));

            var data = Data(env);
            Assert.Null(data["pages"]); // WhenWritingNull drops it from the envelope
            Assert.Equal("deck.pdf", Path.GetFileName(data["written"]!.GetValue<string>()));
            Assert.True(data["sizeBytes"]!.GetValue<long>() > 0);
        }
    }

    [Fact]
    public void Pdf_passes_a_non_ok_inner_render_through_untouched()
    {
        // No browser needed: the failure short-circuits before any screenshot.
        var handler = new StubHandler(DocumentKind.Pptx, _ =>
            Envelope.Fail(ErrorCodes.InvalidArgs, "bad scope", "Pass a valid --scope."));

        var env = PdfRenderVerb.Execute(handler, Ctx(handler, scope: "/slide[99]"));

        Assert.False(env.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, env.Error!.Code);
    }

    // ---- png orchestration ---------------------------------------------------

    [Fact]
    public void Png_deck_without_scope_warns_scope_defaulted_and_defaults_the_output_name()
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        var handler = DeckHandler(4);
        using (UseStubBrowser())
        {
            var env = PngRenderVerb.Execute(handler, Ctx(handler));

            Assert.True(env.IsOk, env.ToJson());
            Assert.Contains("scope_defaulted", WarningCodes(env)); // multi-slide, no scope
            var data = Data(env);
            Assert.Equal("deck.png", Path.GetFileName(data["written"]!.GetValue<string>()));
            Assert.True(data["sizeBytes"]!.GetValue<long>() > 0);
        }
    }

    [Fact]
    public void Png_deck_with_an_explicit_scope_does_not_warn_scope_defaulted()
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        var handler = DeckHandler(4);
        using (UseStubBrowser())
        {
            var env = PngRenderVerb.Execute(handler, Ctx(handler, scope: "/slide[1]"));

            Assert.True(env.IsOk, env.ToJson());
            Assert.DoesNotContain("scope_defaulted", WarningCodes(env));
        }
    }

    [Fact]
    public void Png_passes_a_non_ok_inner_render_through_untouched()
    {
        var handler = new StubHandler(DocumentKind.Pptx, _ =>
            Envelope.Fail(ErrorCodes.FormatCorrupt, "broken deck", "Restore a snapshot."));

        var env = PngRenderVerb.Execute(handler, Ctx(handler));

        Assert.False(env.IsOk);
        Assert.Equal(ErrorCodes.FormatCorrupt, env.Error!.Code);
    }
}
