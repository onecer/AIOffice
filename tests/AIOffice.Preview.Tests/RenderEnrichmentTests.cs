using Xunit;

namespace AIOffice.Preview.Tests;

/// <summary>
/// Pins the enrichment law: until the format renderers emit data-aio-path
/// themselves, the preview injects it deterministically — and steps aside the
/// moment upstream output already carries the attribute.
/// </summary>
public sealed class RenderEnrichmentTests
{
    [Fact]
    public void Word_blocks_get_sequential_paragraph_and_table_paths()
    {
        const string html = "<h2>Title</h2>\n<ul>\n<li>first</li>\n<li>second</li>\n</ul>\n" +
                            "<table>\n<tr><td>x</td><td>y</td></tr>\n</table>\n<p>tail</p>";

        var enriched = PreviewRenderer.InjectWordPaths(html);

        Assert.Contains("<h2 data-aio-path=\"/body/p[1]\">Title</h2>", enriched, StringComparison.Ordinal);
        Assert.Contains("<li data-aio-path=\"/body/p[2]\">first</li>", enriched, StringComparison.Ordinal);
        Assert.Contains("<li data-aio-path=\"/body/p[3]\">second</li>", enriched, StringComparison.Ordinal);
        Assert.Contains("<table data-aio-path=\"/body/table[1]\">", enriched, StringComparison.Ordinal);
        Assert.Contains("<p data-aio-path=\"/body/p[4]\">tail</p>", enriched, StringComparison.Ordinal);
        // Table internals are not addressable blocks in this milestone.
        Assert.Contains("<td>x</td>", enriched, StringComparison.Ordinal);
    }

    [Fact]
    public void Word_html_that_already_carries_the_attribute_is_untouched()
    {
        const string html = "<p data-aio-path=\"/body/p[9]\">already done upstream</p>";

        Assert.Same(html, PreviewRenderer.InjectWordPaths(html));
    }

    [Fact]
    public void Pptx_shapes_are_grouped_with_their_text_lines()
    {
        const string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 100 100\" data-slide=\"1\">\n" +
                           "  <rect x=\"0\" y=\"0\" width=\"100\" height=\"100\" fill=\"#ffffff\" stroke=\"#cccccc\"/>\n" +
                           "  <rect x=\"1\" y=\"1\" width=\"10\" height=\"10\" data-path=\"/slide[1]/shape[@id=2]\" data-name=\"Title\"/>\n" +
                           "  <text x=\"2\" y=\"2\">Hello</text>\n" +
                           "  <rect x=\"5\" y=\"5\" width=\"10\" height=\"10\" data-path=\"/slide[1]/shape[@id=3]\" data-name=\"Body\"/>\n" +
                           "</svg>";

        var wrapped = PreviewRenderer.WrapPptxShapes(svg);

        Assert.Contains("<g data-aio-path=\"/slide[1]/shape[@id=2]\">", wrapped, StringComparison.Ordinal);
        Assert.Contains("<g data-aio-path=\"/slide[1]/shape[@id=3]\">", wrapped, StringComparison.Ordinal);
        Assert.Equal(2, wrapped.Split("</g>").Length - 1);
        // The text line lands inside the first group, before its closing tag.
        Assert.True(
            wrapped.IndexOf(">Hello</text>", StringComparison.Ordinal) < wrapped.IndexOf("</g>", StringComparison.Ordinal),
            "shape text must live inside the shape group");
        // The background rect (no data-path) stays outside any group.
        Assert.True(
            wrapped.IndexOf("stroke=\"#cccccc\"", StringComparison.Ordinal) < wrapped.IndexOf("<g ", StringComparison.Ordinal),
            "the slide background must not be wrapped");
    }

    [Fact]
    public void Pptx_svg_that_already_carries_the_attribute_is_untouched()
    {
        const string svg = "<svg><g data-aio-path=\"/slide[1]/shape[@id=2]\"><rect/></g></svg>";

        Assert.Same(svg, PreviewRenderer.WrapPptxShapes(svg));
    }
}
