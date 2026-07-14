using Xunit;

namespace AIOffice.Render.Tests;

/// <summary>
/// Pure-unit coverage for <see cref="PngRenderVerb.SvgPixelSize"/> — the internal
/// parser that feeds BOTH the png viewport and the pdf <c>@page</c> box. No
/// subprocess, no browser, no flake: runs identically on every OS/runner. It is
/// the cheapest guard on dimension fidelity for every raster path.
/// </summary>
public sealed class SvgPixelSizeTests
{
    [Fact]
    public void Reads_width_and_height_off_the_svg_root()
    {
        var (w, h) = PngRenderVerb.SvgPixelSize(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"1280\" height=\"720\"><rect/></svg>");

        Assert.Equal(1280, w);
        Assert.Equal(720, h);
    }

    [Fact]
    public void Non_default_dimensions_survive_verbatim()
    {
        var (w, h) = PngRenderVerb.SvgPixelSize("<svg width=\"960\" height=\"540\"></svg>");

        Assert.Equal(960, w);
        Assert.Equal(540, h);
    }

    [Fact]
    public void Fractional_dimensions_round_up_via_ceiling()
    {
        // Math.Ceiling: a 100.2px slide must never clip to 100 and lose a column.
        var (w, h) = PngRenderVerb.SvgPixelSize("<svg width=\"100.2\" height=\"200.9\"></svg>");

        Assert.Equal(101, w);
        Assert.Equal(201, h);
    }

    [Fact]
    public void Missing_attributes_fall_back_to_the_documented_default()
    {
        var (w, h) = PngRenderVerb.SvgPixelSize("<svg xmlns=\"http://www.w3.org/2000/svg\"></svg>");

        Assert.Equal(1280, w);
        Assert.Equal(720, h);
    }

    [Theory]
    [InlineData("<svg width=\"nope\" height=\"garbage\"></svg>")]   // non-numeric
    [InlineData("<svg width=\"0\" height=\"0\"></svg>")]             // must be > 0
    [InlineData("<svg width=\"-5\" height=\"-9\"></svg>")]           // negative
    [InlineData("<svg width=\"100000\" height=\"100000\"></svg>")]  // must be < 100000
    [InlineData("<svg width=\"999999\" height=\"999999\"></svg>")]  // out of range
    public void Garbage_or_out_of_range_dimensions_fall_back(string svg)
    {
        var (w, h) = PngRenderVerb.SvgPixelSize(svg);

        Assert.Equal(1280, w);
        Assert.Equal(720, h);
    }

    [Fact]
    public void A_width_only_svg_keeps_the_default_height()
    {
        var (w, h) = PngRenderVerb.SvgPixelSize("<svg width=\"640\"></svg>");

        Assert.Equal(640, w);
        Assert.Equal(720, h); // height falls back independently
    }
}
