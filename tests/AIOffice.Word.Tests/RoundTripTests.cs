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
