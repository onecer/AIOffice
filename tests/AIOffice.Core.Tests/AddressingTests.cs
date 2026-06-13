using AIOffice.Core;
using Xunit;

namespace AIOffice.Core.Tests;

public class AddressingTests
{
    // --- docx paths ---

    [Fact]
    public void Parses_body_paragraph_with_index()
    {
        var path = DocPath.Parse("/body/p[3]");

        Assert.Equal(2, path.Segments.Count);
        Assert.Equal(PathSegmentKind.Element, path.Segments[0].Kind);
        Assert.Equal("body", path.Segments[0].Name);
        Assert.Null(path.Segments[0].Index);
        Assert.Equal("p", path.Segments[1].Name);
        Assert.Equal(3, path.Segments[1].Index);
    }

    [Fact]
    public void Parses_nested_table_cell_path()
    {
        var path = DocPath.Parse("/body/table[1]/tr[2]/tc[1]");

        Assert.Equal(4, path.Segments.Count);
        Assert.Equal(["body", "table", "tr", "tc"], path.Segments.Select(s => s.Name!).ToArray());
        Assert.Equal([null, 1, 2, 1], path.Segments.Select(s => s.Index).ToArray());
    }

    [Fact]
    public void Parses_run_inside_paragraph()
    {
        var path = DocPath.Parse("/body/p[3]/run[2]");

        Assert.Equal("run", path.Segments[2].Name);
        Assert.Equal(2, path.Segments[2].Index);
        Assert.Equal("/body/p[3]/run[2]", path.ToCanonicalString());
    }

    [Fact]
    public void Parses_header_paragraph()
    {
        var path = DocPath.Parse("/header[1]/p[1]");

        Assert.Equal("header", path.Segments[0].Name);
        Assert.Equal(1, path.Segments[0].Index);
        Assert.Equal("/header[1]/p[1]", path.ToCanonicalString());
    }

    // --- xlsx paths ---

    [Fact]
    public void Parses_sheet_and_cell()
    {
        var path = DocPath.Parse("/Sheet1/A1");

        Assert.Equal(PathSegmentKind.Element, path.Segments[0].Kind); // bare sheet name
        Assert.Equal("Sheet1", path.Segments[0].Name);
        Assert.Equal(PathSegmentKind.Cell, path.Segments[1].Kind);
        Assert.Equal(new CellRef("A", 1), path.Segments[1].Start);
    }

    [Fact]
    public void Parses_cell_range()
    {
        var path = DocPath.Parse("/Sheet1/A1:C10");

        var range = path.Segments[1];
        Assert.Equal(PathSegmentKind.Range, range.Kind);
        Assert.Equal(new CellRef("A", 1), range.Start);
        Assert.Equal(new CellRef("C", 10), range.End);
        Assert.Equal("/Sheet1/A1:C10", path.ToCanonicalString());
    }

    [Fact]
    public void Parses_quoted_sheet_name_with_space()
    {
        var path = DocPath.Parse("/'Q3 Data'/B2");

        Assert.Equal(PathSegmentKind.Name, path.Segments[0].Kind);
        Assert.Equal("Q3 Data", path.Segments[0].Name);
        Assert.Equal(new CellRef("B", 2), path.Segments[1].Start);
        Assert.Equal("/'Q3 Data'/B2", path.ToCanonicalString());
    }

    [Fact]
    public void Quoted_sheet_name_doubles_single_quotes()
    {
        var path = DocPath.Parse("/'Bob''s Sheet'/A1");

        Assert.Equal("Bob's Sheet", path.Segments[0].Name);
        Assert.Equal("/'Bob''s Sheet'/A1", path.ToCanonicalString());
    }

    [Fact]
    public void Parses_row_element_on_sheet()
    {
        var path = DocPath.Parse("/Sheet1/row[3]");

        Assert.Equal("row", path.Segments[1].Name);
        Assert.Equal(3, path.Segments[1].Index);
        Assert.Equal(PathSegmentKind.Element, path.Segments[1].Kind);
    }

    [Fact]
    public void Multi_letter_columns_parse_and_number_correctly()
    {
        var path = DocPath.Parse("/Sheet1/AA10:AB20");

        Assert.Equal(27, path.Segments[1].Start!.Value.ColumnNumber);
        Assert.Equal(28, path.Segments[1].End!.Value.ColumnNumber);
    }

    // --- M6: element spans (xlsx outline groups) ---

    [Fact]
    public void Parses_row_span_segment()
    {
        var path = DocPath.Parse("/Sheet1/row[2]:row[6]");

        Assert.Equal(2, path.Segments.Count);
        var span = path.Segments[1];
        Assert.Equal(PathSegmentKind.ElementSpan, span.Kind);
        Assert.Equal("row", span.Name);
        Assert.Equal("2", span.SpanFrom);
        Assert.Equal("6", span.SpanTo);
        Assert.Equal("/Sheet1/row[2]:row[6]", path.ToCanonicalString());
    }

    [Fact]
    public void Parses_column_span_segment()
    {
        var path = DocPath.Parse("/Sheet1/col[B]:col[E]");

        var span = path.Segments[1];
        Assert.Equal(PathSegmentKind.ElementSpan, span.Kind);
        Assert.Equal("col", span.Name);
        Assert.Equal("B", span.SpanFrom);
        Assert.Equal("E", span.SpanTo);
        Assert.Equal("/Sheet1/col[B]:col[E]", path.ToCanonicalString());
    }

    [Theory]
    [InlineData("/Sheet1/row[2]:col[6]")] // mismatched element names
    [InlineData("/Sheet1/row[2]:row[B]")] // numeric : letter bounds
    [InlineData("/Sheet1/col[B]:col[6]")] // letter : numeric bounds
    public void Mismatched_span_bounds_are_rejected(string text)
    {
        Assert.False(DocPath.TryParse(text, out _, out var error));
        Assert.False(string.IsNullOrWhiteSpace(error));
    }

    // --- M6: the document root "/" ---

    [Fact]
    public void Parses_document_root()
    {
        var path = DocPath.Parse("/");

        Assert.True(path.IsRoot);
        Assert.Empty(path.Segments);
        Assert.Equal("/", path.ToCanonicalString());
    }

    [Fact]
    public void Document_root_round_trips()
    {
        var once = DocPath.Parse("/");
        var twice = DocPath.Parse(once.ToCanonicalString());

        Assert.True(twice.IsRoot);
        Assert.Equal("/", twice.ToCanonicalString());
    }

    [Fact]
    public void Non_root_paths_are_not_marked_root()
    {
        Assert.False(DocPath.Parse("/slide[2]").IsRoot);
    }

    // --- pptx paths ---

    [Fact]
    public void Parses_slide_shape_paragraph_chain()
    {
        var path = DocPath.Parse("/slide[2]/shape[3]/p[1]");

        Assert.Equal(3, path.Segments.Count);
        Assert.Equal(["slide", "shape", "p"], path.Segments.Select(s => s.Name!).ToArray());
        Assert.Equal([2, 3, 1], path.Segments.Select(s => s.Index).ToArray());
        Assert.Equal("/slide[2]/shape[3]/p[1]", path.ToCanonicalString());
    }

    [Fact]
    public void Parses_bare_slide()
    {
        var path = DocPath.Parse("/slide[2]");

        Assert.Single(path.Segments);
        Assert.Equal("slide", path.Segments[0].Name);
    }

    // --- canonicalization round-trips ---

    [Theory]
    [InlineData("/body/p[3]")]
    [InlineData("/body/table[1]/tr[2]/tc[1]")]
    [InlineData("/Sheet1/A1:C10")]
    [InlineData("/'Q3 Data'/B2")]
    [InlineData("/slide[2]/shape[3]/p[1]")]
    [InlineData("/header[1]/p[1]")]
    public void Canonical_string_round_trips(string text)
    {
        var once = DocPath.Parse(text);
        var twice = DocPath.Parse(once.ToCanonicalString());

        Assert.Equal(text, once.ToCanonicalString());
        Assert.Equal(once.ToCanonicalString(), twice.ToCanonicalString());
    }

    // --- rejection cases ---

    [Theory]
    [InlineData("")]                  // empty
    [InlineData("body/p[1]")]         // missing leading slash
    [InlineData("/body/p[0]")]        // 0 is not a valid 1-based index
    [InlineData("/body//p[1]")]       // empty segment
    [InlineData("/body/p[1]/")]       // trailing slash
    [InlineData("/'Q3 Data/B2")]      // unterminated quote
    [InlineData("/Sheet1/C10:A1")]    // inverted range
    [InlineData("/body/p[abc]")]      // non-numeric index is not a valid bare name either
    public void Invalid_paths_are_rejected(string text)
    {
        Assert.False(DocPath.TryParse(text, out _, out var error));
        Assert.False(string.IsNullOrWhiteSpace(error));
    }

    [Fact]
    public void Parse_throws_invalid_path_with_suggestion()
    {
        var ex = Assert.Throws<AiofficeException>(() => DocPath.Parse("not-a-path"));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.False(string.IsNullOrWhiteSpace(ex.Suggestion));
    }
}
