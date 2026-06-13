using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class AuditTests : WordTestBase
{
    private static bool Has(AuditResult result, string code) => result.Findings.Any(f => f.Code == code);

    private string CleanDoc()
    {
        // A core title, a properly headed table, default-coloured body text — no findings.
        var file = CreateDoc(title: "Accessible Report");
        Edit(file, """
            [
              {"op":"set","path":"/properties","props":{"title":"Accessible Report"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Section","style":"Heading1"}},
              {"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}},
              {"op":"set","path":"/body/table[1]","props":{"headerRow":true}}
            ]
            """);
        return file;
    }

    // ----------------------------------------------------- clean doc is silent

    [Fact]
    public void Clean_document_reports_no_findings()
    {
        var file = CleanDoc();
        var result = Audit(file);
        Assert.Empty(result.Findings);
        Assert.Equal((0, 0, 0), (result.Summary.Errors, result.Summary.Warnings, result.Summary.Infos));
    }

    // --------------------------------------------------------- a11y_no_doc_title

    [Fact]
    public void Missing_title_fires_and_is_silent_when_present()
    {
        var file = CreateDoc(); // no title
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"Intro","style":"Heading1"}}]""");
        Assert.True(Has(Audit(file), "a11y_no_doc_title"));

        Edit(file, """[{"op":"set","path":"/properties","props":{"title":"Now Titled"}}]""");
        Assert.False(Has(Audit(file), "a11y_no_doc_title"));
    }

    // ---------------------------------------------------------- a11y_no_alt_text

    [Fact]
    public void Image_without_alt_fires_and_is_silent_with_alt()
    {
        var file = CreateDoc(title: "Pics");
        WritePng("logo.png");
        Edit(file, """[{"op":"add","path":"/body","type":"image","props":{"src":"logo.png"}}]""");
        Assert.True(Has(Audit(file), "a11y_no_alt_text"));

        WritePng("logo2.png");
        var clean = CreateDoc("withalt.docx", title: "WithAlt");
        Edit(clean, """[{"op":"add","path":"/body","type":"image","props":{"src":"logo2.png","alt":"Company logo"}}]""");
        Assert.False(Has(Audit(clean), "a11y_no_alt_text"));
    }

    // ----------------------------------------------------------- heading checks

    [Fact]
    public void Heading_skip_fires_on_h1_to_h3()
    {
        var file = CreateDoc(title: "Skips");
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"Top","style":"Heading1"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Deep","style":"Heading3"}}
            ]
            """);

        var result = Audit(file);
        Assert.True(Has(result, "a11y_heading_skip"));
        Assert.False(result.Findings.First(f => f.Code == "a11y_heading_skip").Autofixable);
    }

    [Fact]
    public void Consecutive_levels_do_not_fire_heading_skip()
    {
        var file = CreateDoc(title: "Ordered");
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"A","style":"Heading1"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"B","style":"Heading2"}}
            ]
            """);
        Assert.False(Has(Audit(file), "a11y_heading_skip"));
    }

    [Fact]
    public void Empty_heading_fires_quality_finding()
    {
        var file = CreateDoc(title: "Hollow");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"style":"Heading1"}}]""");
        Assert.True(Has(Audit(file), "quality_empty_heading"));
    }

    // ----------------------------------------------------- a11y_no_table_header

    [Fact]
    public void Table_without_header_row_fires_and_is_silent_with_header()
    {
        var file = CreateDoc(title: "Grids");
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}}]""");
        Assert.True(Has(Audit(file), "a11y_no_table_header"));

        Edit(file, """[{"op":"set","path":"/body/table[1]","props":{"headerRow":true}}]""");
        Assert.False(Has(Audit(file), "a11y_no_table_header"));
    }

    // ------------------------------------------------------- a11y_low_contrast

    [Fact]
    public void Low_contrast_text_fires_and_good_contrast_is_silent()
    {
        var file = CreateDoc(title: "Contrast");
        // Light yellow on the default white page: ratio well under 4.5:1.
        // (Created docs are p1=title heading + p2=empty, so the new run lands at p3.)
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"Faint","color":"FFFF00"}}]""");
        Assert.True(Has(Audit(file), "a11y_low_contrast"));

        // Black on white is the maximum ratio.
        Edit(file, """[{"op":"set","path":"/body/p[3]","props":{"color":"000000"}}]""");
        Assert.False(Has(Audit(file), "a11y_low_contrast"));
    }

    [Fact]
    public void Contrast_ratio_matches_known_wcag_pairs()
    {
        var black = (0, 0, 0);
        var white = (255, 255, 255);

        // Black on white is exactly 21:1.
        Assert.Equal(21.0, WordHandler.ContrastRatio(black, white), 1);

        // A colour against itself is 1:1.
        Assert.Equal(1.0, WordHandler.ContrastRatio(white, white), 3);

        // #767676 on white is the canonical AA threshold (~4.54:1).
        var grey = (0x76, 0x76, 0x76);
        Assert.True(WordHandler.ContrastRatio(grey, white) >= 4.5);

        // Pure white luminance is 1.0, pure black is 0.0.
        Assert.Equal(1.0, WordHandler.RelativeLuminance(white), 3);
        Assert.Equal(0.0, WordHandler.RelativeLuminance(black), 3);
    }

    // ---------------------------------------------------- quality_broken_link

    [Fact]
    public void Broken_internal_link_fires_and_valid_anchor_is_silent()
    {
        var file = CreateDoc(title: "Links");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"Body"}}]""");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Target"}}]""");

        // Link to the real bookmark — fine.
        Edit(file, """[{"op":"add","path":"/body/p[2]","type":"link","props":{"text":"Go","anchor":"Target"}}]""");
        Assert.False(Has(Audit(file), "quality_broken_link"));

        // Manually point a hyperlink at a missing anchor (the add path guards this,
        // so we craft the broken state directly).
        using (var doc = WordprocessingDocument.Open(file, isEditable: true))
        {
            doc.MainDocumentPart!.Document!.Body!.Descendants<Hyperlink>().First().Anchor = "Missing";
        }

        Assert.True(Has(Audit(file), "quality_broken_link"));
    }

    // ------------------------------------------------ quality_orphan_bookmark

    [Fact]
    public void Orphan_bookmark_fires_as_info_and_referenced_one_is_silent()
    {
        var file = CreateDoc(title: "Marks");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Lonely"}}]""");

        var result = Audit(file);
        Assert.True(Has(result, "quality_orphan_bookmark"));
        Assert.Equal("info", result.Findings.First(f => f.Code == "quality_orphan_bookmark").Severity);

        // Add a link that references it -> no longer orphaned.
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"See"}}]""");
        Edit(file, """[{"op":"add","path":"/body/p[2]","type":"link","props":{"text":"Jump","anchor":"Lonely"}}]""");
        Assert.False(Has(Audit(file), "quality_orphan_bookmark"));
    }

    // -------------------------------------------------------- filtering knobs

    [Fact]
    public void Category_filter_limits_to_one_family()
    {
        var file = CreateDoc(title: "Mixed");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Orphan"}}]""");
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":1,"columns":1}}]""");

        var a11y = Audit(file, new AuditOptions { Category = "accessibility" });
        Assert.All(a11y.Findings, f => Assert.Equal("accessibility", f.Category));
        Assert.Contains(a11y.Findings, f => f.Code == "a11y_no_table_header");

        var quality = Audit(file, new AuditOptions { Category = "quality" });
        Assert.All(quality.Findings, f => Assert.Equal("quality", f.Category));
        Assert.Contains(quality.Findings, f => f.Code == "quality_orphan_bookmark");
    }

    [Fact]
    public void Min_severity_filters_out_lower_findings()
    {
        var file = CreateDoc(title: "Severities");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Info"}}]""");

        // The orphan bookmark is info-level; raising the floor to warning drops it.
        Assert.True(Has(Audit(file, new AuditOptions { MinSeverity = "info" }), "quality_orphan_bookmark"));
        Assert.False(Has(Audit(file, new AuditOptions { MinSeverity = "warning" }), "quality_orphan_bookmark"));
    }

    // ----------------------------------------------------------------- --fix

    [Fact]
    public void Fix_sets_placeholder_alt_text_and_reaudit_is_clean()
    {
        var file = CreateDoc(title: "FixAlt");
        WritePng("p.png");
        Edit(file, """[{"op":"add","path":"/body","type":"image","props":{"src":"p.png"}}]""");

        var finding = Audit(file).Findings.Single(f => f.Code == "a11y_no_alt_text");
        var fixedCount = Handler.Fix(Ctx(file), [finding.Id]);
        Assert.Equal(1, fixedCount);

        Assert.False(Has(Audit(file), "a11y_no_alt_text"));
        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = finding.Path! })))["properties"]!;
        Assert.Equal("(describe this image)", got["alt"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Fix_marks_a_table_header_row()
    {
        var file = CreateDoc(title: "FixTable");
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}}]""");

        var finding = Audit(file).Findings.Single(f => f.Code == "a11y_no_table_header");
        Assert.Equal(1, Handler.Fix(Ctx(file), [finding.Id]));

        Assert.False(Has(Audit(file), "a11y_no_table_header"));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Fix_sets_title_from_the_first_heading()
    {
        var file = CreateDoc(); // no title
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"Annual Review","style":"Heading1"}}]""");

        var finding = Audit(file).Findings.Single(f => f.Code == "a11y_no_doc_title");
        Assert.Equal(1, Handler.Fix(Ctx(file), [finding.Id]));

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        Assert.Equal("Annual Review", doc.PackageProperties.Title);
    }

    [Fact]
    public void Fix_removes_an_orphan_bookmark()
    {
        var file = CreateDoc(title: "FixOrphan");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Dead"}}]""");

        var finding = Audit(file).Findings.Single(f => f.Code == "quality_orphan_bookmark");
        Assert.Equal(1, Handler.Fix(Ctx(file), [finding.Id]));

        Assert.False(Has(Audit(file), "quality_orphan_bookmark"));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Default_fix_applies_every_autofixable_finding()
    {
        var file = CreateDoc(); // no title
        WritePng("d.png");
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"Heading","style":"Heading1"}},
              {"op":"add","path":"/body","type":"image","props":{"src":"d.png"}},
              {"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}},
              {"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Orphan"}}
            ]
            """);

        // Empty id list = fix all autofixable.
        var fixedCount = Handler.Fix(Ctx(file), []);
        Assert.Equal(4, fixedCount); // title, alt, header, orphan

        var result = Audit(file);
        Assert.DoesNotContain(result.Findings, f => f.Autofixable);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Fix_never_touches_non_autofixable_findings()
    {
        var file = CreateDoc(title: "Stubborn");
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"H1","style":"Heading1"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"H3","style":"Heading3"}}
            ]
            """);

        var skip = Audit(file).Findings.Single(f => f.Code == "a11y_heading_skip");
        Assert.Equal(0, Handler.Fix(Ctx(file), [skip.Id])); // not autofixable
        Assert.True(Has(Audit(file), "a11y_heading_skip")); // still there
    }

    // ------------------------------------------------------------ finding id

    [Fact]
    public void Finding_id_combines_code_and_path()
    {
        var file = CreateDoc(title: "Ids");
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":1,"columns":1}}]""");

        var finding = Audit(file).Findings.Single(f => f.Code == "a11y_no_table_header");
        Assert.Equal("a11y_no_table_header#/body/table[1]", finding.Id);
        Assert.Equal("/body/table[1]", finding.Path);
    }
}
