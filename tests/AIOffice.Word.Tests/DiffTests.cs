using AIOffice.Core;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class DiffTests : WordTestBase
{
    /// <summary>Diffs <paramref name="current"/> against <paramref name="baseline"/> (the OLD file).</summary>
    private DiffResult Diff(string current, string baseline) =>
        Handler.Diff(Ctx(current), baseline);

    /// <summary>A two-paragraph fixture: p1 "Intro" (Heading1), p2 "First body line".</summary>
    private string BaseDoc(string name)
    {
        var file = CreateDoc(name);
        Edit(file, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"Intro","style":"Heading1"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"First body line."}}
            ]
            """);
        return file;
    }

    // ---------------------------------------------------------- identical files

    [Fact]
    public void Identical_documents_produce_an_empty_diff()
    {
        var baseline = BaseDoc("base.docx");
        var current = CopyOf(baseline, "current.docx");

        var result = Diff(current, baseline);
        Assert.Empty(result.Changes);
        Assert.Equal((0, 0, 0, 0),
            (result.Summary.Added, result.Summary.Removed, result.Summary.Modified, result.Summary.Moved));
    }

    // ------------------------------------------------------------- added block

    [Fact]
    public void Added_paragraph_is_reported_as_added()
    {
        var baseline = BaseDoc("base.docx");
        var current = CopyOf(baseline, "current.docx");
        Edit(current, """[{"op":"add","path":"/body","type":"p","props":{"text":"Brand new paragraph."}}]""");

        var result = Diff(current, baseline);
        var added = Assert.Single(result.Changes, c => c.Kind == "added");
        Assert.Equal("/body/p[3]", added.Path);
        Assert.Equal(1, result.Summary.Added);
    }

    // ----------------------------------------------------------- removed block

    [Fact]
    public void Removed_paragraph_is_reported_as_removed_at_baseline_path()
    {
        var baseline = BaseDoc("base.docx");
        var current = CopyOf(baseline, "current.docx");
        Edit(current, """[{"op":"remove","path":"/body/p[2]"}]""");

        var result = Diff(current, baseline);
        var removed = Assert.Single(result.Changes, c => c.Kind == "removed");
        Assert.Equal("/body/p[2]", removed.Path);
        Assert.Equal(1, result.Summary.Removed);
    }

    // ---------------------------------------------------------- modified text

    [Fact]
    public void Modified_paragraph_text_carries_before_and_after()
    {
        var baseline = BaseDoc("base.docx");
        var current = CopyOf(baseline, "current.docx");
        Edit(current, """[{"op":"set","path":"/body/p[2]","props":{"text":"Edited body line."}}]""");

        var result = Diff(current, baseline);
        var modified = Assert.Single(result.Changes, c => c.Kind == "modified" && c.Detail == "text");
        Assert.Equal("/body/p[2]", modified.Path);
        Assert.Equal("First body line.", modified.Before);
        Assert.Equal("Edited body line.", modified.After);
    }

    // --------------------------------------------------------- modified style

    [Fact]
    public void Heading_style_change_is_reported_with_before_and_after_styles()
    {
        var baseline = BaseDoc("base.docx");
        var current = CopyOf(baseline, "current.docx");
        // Promote the body line to a Heading2 (text unchanged -> only a style change).
        Edit(current, """[{"op":"set","path":"/body/p[2]","props":{"style":"Heading2"}}]""");

        var result = Diff(current, baseline);
        var styleChange = Assert.Single(result.Changes, c => c.Kind == "modified" && c.Detail == "style");
        Assert.Equal("/body/p[2]", styleChange.Path);
        Assert.Equal("Normal", styleChange.Before);
        Assert.Equal("Heading2", styleChange.After);
    }

    // ------------------------------------------------------------ moved block

    [Fact]
    public void Reordered_paragraph_is_reported_as_moved()
    {
        var baseline = CreateDoc("base.docx");
        Edit(baseline, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"Alpha"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Bravo"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Charlie"}}
            ]
            """);
        var current = CopyOf(baseline, "current.docx");
        // Move "Alpha" to after "Charlie": same content, new position.
        Edit(current, """[{"op":"move","path":"/body/p[1]","position":"after /body/p[3]"}]""");

        var result = Diff(current, baseline);
        var moved = Assert.Single(result.Changes, c => c.Kind == "moved");
        Assert.StartsWith("/body/p[", moved.Path);
        Assert.Contains("moved from", moved.Detail);
        Assert.Equal(0, result.Summary.Added);
        Assert.Equal(0, result.Summary.Removed);
    }

    // ------------------------------------------------------- table cell change

    [Fact]
    public void Table_cell_text_change_is_reported_at_the_cell_path()
    {
        var baseline = CreateDoc("base.docx");
        Edit(baseline, """
            [
              {"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}},
              {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"old"}}
            ]
            """);
        var current = CopyOf(baseline, "current.docx");
        Edit(current, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"new"}}]""");

        var result = Diff(current, baseline);
        var cell = Assert.Single(result.Changes, c => c.Detail == "cell");
        Assert.Equal("modified", cell.Kind);
        Assert.Equal("/body/table[1]/tr[1]/tc[1]", cell.Path);
        Assert.Equal("old", cell.Before);
        Assert.Equal("new", cell.After);
    }

    // -------------------------------------------------------- header change

    [Fact]
    public void Header_text_change_is_reported_at_the_header_paragraph_path()
    {
        var baseline = CreateDoc("base.docx");
        Edit(baseline, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"Old header"}}]""");
        var current = CopyOf(baseline, "current.docx");
        Edit(current, """[{"op":"set","path":"/header[1]/p[1]","props":{"text":"New header"}}]""");

        var result = Diff(current, baseline);
        var header = Assert.Single(result.Changes, c => c.Detail == "headerFooter");
        Assert.Equal("modified", header.Kind);
        Assert.Equal("/header[1]/p[1]", header.Path);
        Assert.Equal("Old header", header.Before);
        Assert.Equal("New header", header.After);
    }

    // ------------------------------------------------------ property change

    [Fact]
    public void Document_property_change_is_reported_at_properties()
    {
        var baseline = CreateDoc("base.docx");
        // --title only adds a heading; set the CORE property the diff reads.
        Edit(baseline, """[{"op":"set","path":"/properties","props":{"title":"Original Title"}}]""");
        var current = CopyOf(baseline, "current.docx");
        Edit(current, """[{"op":"set","path":"/properties","props":{"title":"Revised Title","author":"Alex"}}]""");

        var result = Diff(current, baseline);
        var titleChange = Assert.Single(result.Changes, c => c.Detail == "title");
        Assert.Equal("/properties", titleChange.Path);
        Assert.Equal("Original Title", titleChange.Before);
        Assert.Equal("Revised Title", titleChange.After);

        var authorChange = Assert.Single(result.Changes, c => c.Detail == "author");
        Assert.Null(authorChange.Before);
        Assert.Equal("Alex", authorChange.After);
    }

    // ------------------------------------------------------- deterministic order

    [Fact]
    public void Changes_are_sorted_by_path_then_kind_deterministically()
    {
        var baseline = CreateDoc("base.docx", title: "T");
        Edit(baseline, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"Keep one"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Edit me"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Delete me"}}
            ]
            """);
        var current = CopyOf(baseline, "current.docx");
        Edit(current, """
            [
              {"op":"set","path":"/body/p[2]","props":{"text":"Edited"}},
              {"op":"remove","path":"/body/p[3]"},
              {"op":"add","path":"/body","type":"p","props":{"text":"Appended"}},
              {"op":"set","path":"/properties","props":{"title":"T2"}}
            ]
            """);

        // The diff must be sorted by (Path, Kind) ordinal, identically every run.
        var first = Diff(current, baseline).Changes;
        var second = Diff(current, baseline).Changes;

        var firstKeys = first.Select(c => $"{c.Path}|{c.Kind}|{c.Detail}").ToList();
        var secondKeys = second.Select(c => $"{c.Path}|{c.Kind}|{c.Detail}").ToList();
        Assert.Equal(firstKeys, secondKeys);

        // Explicitly assert ordinal (Path, Kind) ordering.
        var sortedKeys = first
            .OrderBy(c => c.Path, StringComparer.Ordinal)
            .ThenBy(c => c.Kind, StringComparer.Ordinal)
            .ThenBy(c => c.Detail ?? string.Empty, StringComparer.Ordinal)
            .Select(c => $"{c.Path}|{c.Kind}|{c.Detail}")
            .ToList();
        Assert.Equal(sortedKeys, firstKeys);
    }

    // ----------------------------------------------------------- guard rails

    [Fact]
    public void Diffing_against_a_non_docx_baseline_is_invalid_args()
    {
        var current = BaseDoc("current.docx");
        File.WriteAllText(Path.Combine(Dir, "data.xlsx"), "not really a workbook");

        var ex = Assert.Throws<AiofficeException>(() => Handler.Diff(Ctx(current), "data.xlsx"));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Diffing_against_a_missing_baseline_is_file_not_found()
    {
        var current = BaseDoc("current.docx");
        // A baseline path that resolves inside the sandbox but does not exist.
        var ex = Assert.Throws<AiofficeException>(() => Handler.Diff(Ctx(current), "ghost.docx"));
        Assert.Equal(ErrorCodes.FileNotFound, ex.Code);
    }

    // --------------------------------------------------------------- helpers

    /// <summary>Byte-copies a baseline doc into a second workspace file.</summary>
    private string CopyOf(string source, string name)
    {
        var dest = Path.Combine(Dir, name);
        File.Copy(source, dest, overwrite: true);
        return dest;
    }
}
