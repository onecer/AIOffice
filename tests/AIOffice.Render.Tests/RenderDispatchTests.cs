using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Render.Tests;

/// <summary>
/// A stub handler whose <see cref="Render"/> records the context it saw and
/// returns a marker envelope — enough to prove the dispatch path without a real
/// docx/xlsx/pptx handler (which this project does not reference).
/// </summary>
internal sealed class StubHandler : IFormatHandler
{
    public StubHandler(DocumentKind kind) => Kind = kind;

    public DocumentKind Kind { get; }

    public int RenderCalls { get; private set; }

    public Envelope Create(CommandContext ctx) => throw new NotSupportedException();

    public Envelope Read(CommandContext ctx) => throw new NotSupportedException();

    public Envelope Get(CommandContext ctx) => throw new NotSupportedException();

    public Envelope Query(CommandContext ctx) => throw new NotSupportedException();

    public Envelope Edit(CommandContext ctx, IReadOnlyList<EditOp> ops) => throw new NotSupportedException();

    public Envelope Render(CommandContext ctx)
    {
        RenderCalls++;
        return Envelope.Ok(new { content = "<p>native</p>" });
    }

    public Envelope Validate(CommandContext ctx) => throw new NotSupportedException();

    public Envelope Template(CommandContext ctx) => throw new NotSupportedException();
}

public sealed class RenderDispatchTests
{
    private static CommandContext Ctx(string? to, string? engine = null, string? scope = null)
    {
        var args = new JsonObject { ["to"] = to };
        if (engine is not null)
        {
            args["engine"] = engine;
        }

        if (scope is not null)
        {
            args["scope"] = scope;
        }

        return new CommandContext { Workspace = new Workspace(Path.GetTempPath()), File = "x.docx", Args = args };
    }

    [Fact]
    public void Svg_with_no_engine_goes_straight_to_the_native_handler_unchanged()
    {
        var handler = new StubHandler(DocumentKind.Docx);
        var result = RenderDispatch.Execute(handler, Ctx("svg"), "svg");

        Assert.True(result.IsOk);
        Assert.Equal(1, handler.RenderCalls);
        // The default (chromium) path must add NO warning to the native render.
        Assert.Null(result.Meta.Warnings);
    }

    [Fact]
    public void Html_with_chromium_engine_is_byte_for_byte_the_native_render()
    {
        var handler = new StubHandler(DocumentKind.Xlsx);
        var native = handler.Render(Ctx("html"));
        var dispatched = RenderDispatch.Execute(new StubHandler(DocumentKind.Xlsx), Ctx("html", engine: "chromium"), "html");

        Assert.True(dispatched.IsOk);
        Assert.Null(dispatched.Meta.Warnings);
        Assert.Equal(native.Data!.ToString(), dispatched.Data!.ToString());
    }

    [Fact]
    public void Soffice_engine_on_svg_falls_back_to_native_with_an_engine_fallback_warning()
    {
        var handler = new StubHandler(DocumentKind.Pptx);
        var result = RenderDispatch.Execute(handler, Ctx("svg", engine: "soffice"), "svg");

        Assert.True(result.IsOk);
        Assert.Equal(1, handler.RenderCalls); // native engine still produced the svg
        var warning = Assert.Single(result.Meta.Warnings!);
        Assert.Equal(RenderDispatch.EngineFallbackCode, warning.Code);
    }

    [Fact]
    public void Soffice_engine_on_text_also_warns_and_falls_back()
    {
        var handler = new StubHandler(DocumentKind.Docx);
        var result = RenderDispatch.Execute(handler, Ctx("text", engine: "soffice"), "text");

        Assert.True(result.IsOk);
        Assert.Contains(result.Meta.Warnings!, w => w.Code == RenderDispatch.EngineFallbackCode);
    }

    [Fact]
    public void An_unknown_engine_is_invalid_args_even_on_a_native_target()
    {
        // The dispatch validates --engine up-front, so a bad value fails fast
        // even for svg/html/text (which otherwise skip the engine entirely).
        var handler = new StubHandler(DocumentKind.Docx);
        var ex = Assert.Throws<AiofficeException>(() =>
            RenderDispatch.Execute(handler, Ctx("svg", engine: "ms-word"), "svg"));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }
}
