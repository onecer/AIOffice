using AIOffice.Core;
using Xunit;

namespace AIOffice.Render.Tests;

public sealed class RenderEngineSelectorTests
{
    [Theory]
    [InlineData(null, RenderEngine.Chromium)]
    [InlineData("", RenderEngine.Chromium)]
    [InlineData("chromium", RenderEngine.Chromium)]
    [InlineData("CHROMIUM", RenderEngine.Chromium)]
    [InlineData("soffice", RenderEngine.Soffice)]
    [InlineData("Soffice", RenderEngine.Soffice)]
    [InlineData("auto", RenderEngine.Auto)]
    public void Parse_maps_the_engine_values(string? value, RenderEngine expected)
    {
        Assert.Equal(expected, RenderEngineSelector.Parse(value));
    }

    [Fact]
    public void Parse_rejects_an_unknown_engine_as_invalid_args()
    {
        var ex = Assert.Throws<AiofficeException>(() => RenderEngineSelector.Parse("libreoffice"));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("chromium", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Resolve_passes_chromium_and_soffice_through_unchanged()
    {
        Assert.Equal(RenderEngine.Chromium, RenderEngineSelector.Resolve(RenderEngine.Chromium));
        Assert.Equal(RenderEngine.Soffice, RenderEngineSelector.Resolve(RenderEngine.Soffice));
    }

    [Fact]
    public void Resolve_auto_picks_soffice_when_present_else_chromium()
    {
        Assert.Equal(
            RenderEngine.Soffice,
            RenderEngineSelector.Resolve(RenderEngine.Auto, new SofficeInfo(true, "/x/soffice", true, "/x/pdftoppm")));

        Assert.Equal(
            RenderEngine.Chromium,
            RenderEngineSelector.Resolve(RenderEngine.Auto, SofficeInfo.NotFound));
    }

    [Theory]
    [InlineData(null, 1)]
    [InlineData("", 1)]
    [InlineData("/slide[1]", 1)]
    [InlineData("/slide[3]", 3)]
    [InlineData("/SLIDE[7]", 7)]
    [InlineData("/body/p[2]", 1)] // non-slide scope -> page 1
    [InlineData("/Sheet1/A1:F20", 1)]
    public void PageFromScope_reads_the_slide_index_or_defaults_to_one(string? scope, int expected)
    {
        Assert.Equal(expected, RenderEngineSelector.PageFromScope(scope));
    }
}
