using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class RoundTripTests : WordTestBase
{
    /// <summary>
    /// The round-trip law: opening a docx and saving it WITHOUT edits must leave
    /// every zip part byte-identical. This pins the editing pipeline (in-memory
    /// copy, DOM load, save-on-dispose) as non-destructive for untouched content.
    /// </summary>
    [Fact]
    public void Open_then_save_without_edits_keeps_every_part_byte_identical()
    {
        var file = CreateDoc(title: "Round trip");
        // Enrich the fixture so the law covers styles, tables and run formatting.
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"Bold lead","bold":true,"color":"FF0000"}},
              {"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":3}},
              {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"cell"}}
            ]
            """);

        var before = File.ReadAllBytes(file);

        var ms = new MemoryStream();
        ms.Write(before);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, isEditable: true))
        {
            // Load the same DOM roots the edit pipeline touches, then save with no changes.
            _ = doc.MainDocumentPart!.Document!.Body;
            _ = doc.MainDocumentPart.StyleDefinitionsPart?.Styles;
        }

        var after = ms.ToArray();

        AssertZipPartsIdentical(before, after);
    }

    /// <summary>
    /// The round-trip law extended over the M2 surface: tracked changes,
    /// comments, custom styles and embedded images must all survive a no-edit
    /// open+save byte-identically.
    /// </summary>
    [Fact]
    public void Open_then_save_with_m2_features_keeps_every_part_byte_identical()
    {
        var file = CreateDoc(title: "M2 round trip");
        File.WriteAllBytes(
            Path.Combine(Dir, "dot.png"),
            Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="));

        Edit(file, """
            [
              {"op":"add","path":"/styles","type":"style","props":{"id":"Callout","bold":true,"color":"1F4E79"}},
              {"op":"add","path":"/body","props":{"text":"Styled","style":"Callout"}},
              {"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"please check"}},
              {"op":"add","path":"/body","type":"image","props":{"src":"dot.png","width":"2cm"}}
            ]
            """);
        Edit(
            file,
            """[{"op":"set","path":"/body/p[3]","props":{"text":"Styled and tracked"}}]""",
            new System.Text.Json.Nodes.JsonObject { ["track"] = true });

        var before = File.ReadAllBytes(file);

        var ms = new MemoryStream();
        ms.Write(before);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, isEditable: true))
        {
            _ = doc.MainDocumentPart!.Document!.Body;
            _ = doc.MainDocumentPart.StyleDefinitionsPart?.Styles;
            _ = doc.MainDocumentPart.WordprocessingCommentsPart?.Comments;
        }

        var after = ms.ToArray();

        AssertZipPartsIdentical(before, after);
    }

    /// <summary>
    /// The round-trip law over the M3 surface: lists, links, bookmarks,
    /// footnotes, page setup and comment threads must survive a no-edit
    /// open+save byte-identically.
    /// </summary>
    [Fact]
    public void Open_then_save_with_m3_features_keeps_every_part_byte_identical()
    {
        var file = CreateDoc(title: "M3 round trip");
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"Item one","list":"number"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Sub item","list":"number","level":1}},
              {"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Top"}},
              {"op":"add","path":"/body/p[1]","type":"link","props":{"text":"site","url":"https://example.com"}},
              {"op":"add","path":"/body/p[1]","type":"link","props":{"text":"top","anchor":"Top"}},
              {"op":"add","path":"/body/p[3]","type":"footnote","props":{"text":"a note"}},
              {"op":"set","path":"/section[1]","props":{"pageSize":"A4","orientation":"landscape","marginTop":"2cm"}},
              {"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"root comment"}}
            ]
            """);
        Edit(file, """[{"op":"add","path":"/comment[@id=1]","type":"reply","props":{"text":"a reply"}}]""");

        var before = File.ReadAllBytes(file);

        var ms = new MemoryStream();
        ms.Write(before);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, isEditable: true))
        {
            _ = doc.MainDocumentPart!.Document!.Body;
            _ = doc.MainDocumentPart.StyleDefinitionsPart?.Styles;
            _ = doc.MainDocumentPart.NumberingDefinitionsPart?.Numbering;
            _ = doc.MainDocumentPart.FootnotesPart?.Footnotes;
            _ = doc.MainDocumentPart.WordprocessingCommentsPart?.Comments;
            _ = doc.MainDocumentPart.WordprocessingCommentsExPart?.CommentsEx;
        }

        var after = ms.ToArray();

        AssertZipPartsIdentical(before, after);
    }

    /// <summary>
    /// The round-trip law over the M4 surface: find/replace rewrites, a TOC
    /// (sdt + field + bookmarks), a VML watermark in the header, endnotes and
    /// a multi-section body must survive a no-edit open+save byte-identically.
    /// </summary>
    [Fact]
    public void Open_then_save_with_m4_features_keeps_every_part_byte_identical()
    {
        var file = CreateDoc(title: "M4 round trip");
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"Chapter one","style":"Heading1"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Original body text."}},
              {"op":"add","path":"/body","type":"toc","props":{"levels":"1-3","title":"Contents","position":"before /body/p[1]"}},
              {"op":"add","path":"/body","type":"watermark","props":{"text":"DRAFT"}},
              {"op":"add","path":"/body/p[3]","type":"endnote","props":{"text":"an endnote"}},
              {"op":"add","path":"/body/p[2]","type":"sectionBreak","props":{"kind":"continuous"}},
              {"op":"replace","path":"/body","props":{"find":"Original","replace":"Replaced"}}
            ]
            """);

        var before = File.ReadAllBytes(file);

        var ms = new MemoryStream();
        ms.Write(before);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, isEditable: true))
        {
            _ = doc.MainDocumentPart!.Document!.Body;
            _ = doc.MainDocumentPart.StyleDefinitionsPart?.Styles;
            _ = doc.MainDocumentPart.EndnotesPart?.Endnotes;
            foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
            {
                _ = headerPart.Header; // loads the VML watermark DOM
            }
        }

        var after = ms.ToArray();

        AssertZipPartsIdentical(before, after);
    }

    /// <summary>
    /// The round-trip law over the M5 surface: merged/styled tables, fldSimple
    /// fields, first-page/even headers (titlePg + settings) must survive a
    /// no-edit open+save byte-identically.
    /// </summary>
    [Fact]
    public void Open_then_save_with_m5_features_keeps_every_part_byte_identical()
    {
        var file = CreateDoc(title: "M5 round trip");
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"table","props":{"rows":3,"columns":3}},
              {"op":"set","path":"/body/table[1]","props":{"borders":"outer","shading":"F7F7F7","headerRow":true,"columnWidths":["3cm","3cm","3cm"],"width":"100%","alignment":"center","cellPaddingCm":0.15}},
              {"op":"set","path":"/body/table[1]/tr[2]/tc[1]","props":{"mergeRight":2}},
              {"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"mergeDown":2}},
              {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"valign":"center"}},
              {"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"date","format":"yyyy-MM-dd"}},
              {"op":"add","path":"/header[firstPage]","type":"header","props":{"text":"First"}},
              {"op":"add","path":"/footer[even]","type":"footer","props":{"text":"Even"}},
              {"op":"add","path":"/footer[1]","type":"footer","props":{"text":"Page "}},
              {"op":"add","path":"/footer[default]/p[1]","type":"field","props":{"kind":"pageNumber"}},
              {"op":"add","path":"/footer[default]/p[1]","type":"field","props":{"kind":"numPages","leadingText":" of "}}
            ]
            """);

        var before = File.ReadAllBytes(file);

        var ms = new MemoryStream();
        ms.Write(before);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, isEditable: true))
        {
            _ = doc.MainDocumentPart!.Document!.Body;
            _ = doc.MainDocumentPart.StyleDefinitionsPart?.Styles;
            _ = doc.MainDocumentPart.DocumentSettingsPart?.Settings;
            foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
            {
                _ = headerPart.Header;
            }

            foreach (var footerPart in doc.MainDocumentPart.FooterParts)
            {
                _ = footerPart.Footer;
            }
        }

        var after = ms.ToArray();

        AssertZipPartsIdentical(before, after);
    }

    /// <summary>
    /// The round-trip law over the M6 surface: inline + display LaTeX equations
    /// (with their stored mc:Ignorable LaTeX source), RTL paragraphs/tables and a
    /// multi-column section must survive a no-edit open+save byte-identically.
    /// </summary>
    [Fact]
    public void Open_then_save_with_m6_features_keeps_every_part_byte_identical()
    {
        var file = CreateDoc(title: "M6 round trip");
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"Body text."}},
              {"op":"add","path":"/body/p[2]","type":"equation","props":{"latex":"E = mc^2"}},
              {"op":"add","path":"/body","type":"equation","props":{"latex":"\\sum_{i=1}^{n} i^2","display":true}},
              {"op":"set","path":"/body/p[1]","props":{"rtl":true}},
              {"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}},
              {"op":"set","path":"/body/table[1]","props":{"rtl":true}},
              {"op":"set","path":"/section[1]","props":{"columns":2,"columnGap":"1cm"}}
            ]
            """);

        var before = File.ReadAllBytes(file);

        var ms = new MemoryStream();
        ms.Write(before);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, isEditable: true))
        {
            _ = doc.MainDocumentPart!.Document!.Body; // loads the OMML equation DOM
            _ = doc.MainDocumentPart.StyleDefinitionsPart?.Styles;
        }

        var after = ms.ToArray();

        AssertZipPartsIdentical(before, after);
    }

    /// <summary>
    /// The round-trip law over the v1.1.0 surface: w14 text effects (shadow /
    /// glow / reflection / outline) on a run, plus the bibliography store
    /// (a customXml Sources part), a CITATION field and a BIBLIOGRAPHY block,
    /// must survive a no-edit open+save byte-identically.
    /// </summary>
    [Fact]
    public void Open_then_save_with_v110_features_keeps_every_part_byte_identical()
    {
        var file = CreateDoc(title: "v1.1.0 round trip");
        Edit(file, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"Effected lead","shadow":true,"glow":{"color":"4F81BD","radius":5},"reflection":true,"outline":{"color":"000000","width":1}}},
              {"op":"add","path":"/sources","type":"source","props":{"tag":"Smith2020","kind":"book","author":"Smith, John","title":"A Great Book","year":2020,"publisher":"Acme Press"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"As shown."}},
              {"op":"add","path":"/body/p[2]","type":"citation","props":{"source":"Smith2020","pages":"42"}},
              {"op":"add","path":"/body","type":"bibliography","props":{"style":"APA"}}
            ]
            """);

        var before = File.ReadAllBytes(file);

        var ms = new MemoryStream();
        ms.Write(before);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, isEditable: true))
        {
            _ = doc.MainDocumentPart!.Document!.Body; // loads the w14 effect + field DOM
            _ = doc.MainDocumentPart.StyleDefinitionsPart?.Styles;
            foreach (var customXml in doc.MainDocumentPart.CustomXmlParts)
            {
                using var reader = new StreamReader(customXml.GetStream());
                _ = reader.ReadToEnd(); // touches the bibliography Sources store
            }
        }

        var after = ms.ToArray();

        AssertZipPartsIdentical(before, after);
    }

    private static void AssertZipPartsIdentical(byte[] before, byte[] after)
    {
        using var zipBefore = new ZipArchive(new MemoryStream(before), ZipArchiveMode.Read);
        using var zipAfter = new ZipArchive(new MemoryStream(after), ZipArchiveMode.Read);

        var namesBefore = zipBefore.Entries.Select(e => e.FullName).OrderBy(n => n, StringComparer.Ordinal).ToList();
        var namesAfter = zipAfter.Entries.Select(e => e.FullName).OrderBy(n => n, StringComparer.Ordinal).ToList();
        Assert.Equal(namesBefore, namesAfter);

        foreach (var name in namesBefore)
        {
            var a = ReadEntry(zipBefore, name);
            var b = ReadEntry(zipAfter, name);
            Assert.True(a.SequenceEqual(b), $"Zip part '{name}' changed on a no-edit open+save.");
        }
    }

    private static byte[] ReadEntry(ZipArchive zip, string name)
    {
        using var stream = zip.GetEntry(name)!.Open();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }
}
