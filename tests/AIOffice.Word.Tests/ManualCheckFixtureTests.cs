using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// We cannot open real Word here, so this test regenerates a representative
/// fixture into fixtures/manual-check/ for a human to open, and holds the
/// OpenXmlValidator at 0 errors as the automated proxy.
/// </summary>
public sealed class ManualCheckFixtureTests : WordTestBase
{
    [Fact]
    public void Regenerate_word_sample_for_human_inspection()
    {
        var repoRoot = FindRepoRoot();
        if (repoRoot is null)
        {
            return; // running outside the repo (e.g. CI artifact stage); nothing to regenerate
        }

        var dir = Path.Combine(repoRoot, "fixtures", "manual-check");
        Directory.CreateDirectory(dir);
        var file = Path.Combine(dir, "word-sample.docx");
        if (File.Exists(file))
        {
            File.Delete(file);
        }

        var workspace = new Workspace(dir);
        var ctx = new CommandContext
        {
            Workspace = workspace,
            File = file,
            Args = new JsonObject { ["title"] = "AIOffice Manual Check" },
        };
        Handler.Create(ctx);

        Handler.Edit(
            new CommandContext { Workspace = workspace, File = file, Args = [] },
            EditOp.ParseBatch("""
                [
                  {"op":"set","path":"/body/p[2]","props":{"text":"Open this file in Word: heading, formatting and table below must look right."}},
                  {"op":"add","path":"/body","props":{"text":"Formatting","style":"Heading2"}},
                  {"op":"add","path":"/body","props":{"text":"Bold red 14pt centered.","bold":true,"color":"CC0000","fontSize":14,"alignment":"center"}},
                  {"op":"add","path":"/body","props":{"text":"Italic and underlined.","italic":true,"underline":true}},
                  {"op":"add","path":"/body","props":{"text":"Table","style":"Heading2"}},
                  {"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":3}},
                  {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"Quarter"}},
                  {"op":"set","path":"/body/table[1]/tr[1]/tc[2]","props":{"text":"Revenue"}},
                  {"op":"set","path":"/body/table[1]/tr[1]/tc[3]","props":{"text":"Growth"}},
                  {"op":"add","path":"/body/table[1]","type":"tr","props":{"cells":["Q3","1.2M","8%"]}},
                  {"op":"add","path":"/header[1]","type":"header","props":{"text":"AIOffice manual check — header","alignment":"center"}},
                  {"op":"add","path":"/footer[1]","type":"footer","props":{"text":"AIOffice manual check — footer"}}
                ]
                """));

        // M2 surface: a custom style, an embedded picture, a comment, and one
        // tracked change deliberately left pending for the human to see.
        var logo = Path.Combine(dir, "manual-check-logo.png");
        File.WriteAllBytes(logo, OpaquePng(120, 40));
        try
        {
            var m2 = Handler.Edit(
                new CommandContext { Workspace = workspace, File = file, Args = [] },
                EditOp.ParseBatch("""
                    [
                      {"op":"add","path":"/styles","type":"style","props":
                        {"id":"Callout","bold":true,"color":"1F4E79","fontSize":12,"alignment":"center","spacingBefore":6,"spacingAfter":6}},
                      {"op":"add","path":"/body","props":{"text":"Styles (M2)","style":"Heading2"}},
                      {"op":"add","path":"/body","props":{"text":"This paragraph uses the custom Callout style.","style":"Callout"}},
                      {"op":"add","path":"/body","props":{"text":"Image (M2)","style":"Heading2"}},
                      {"op":"add","path":"/body","type":"image","props":{"src":"manual-check-logo.png","width":"4cm"}},
                      {"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"Comment written by aioffice — visible in the review pane.","author":"Reviewer"}}
                    ]
                    """));
            Assert.True(m2.IsOk, m2.ToJson());
        }
        finally
        {
            File.Delete(logo);
        }

        // M3 surface: lists with nesting, external + internal links over a
        // bookmark, a footnote, a threaded comment reply, landscape A4 pages.
        var m3 = Handler.Edit(
            new CommandContext { Workspace = workspace, File = file, Args = [] },
            EditOp.ParseBatch("""
                [
                  {"op":"add","path":"/body","props":{"text":"Lists (M3)","style":"Heading2"}},
                  {"op":"add","path":"/body","type":"p","props":{"text":"First step","list":"number"}},
                  {"op":"add","path":"/body","type":"p","props":{"text":"Nested detail","list":"number","level":1}},
                  {"op":"add","path":"/body","type":"p","props":{"text":"Second step","list":"number"}},
                  {"op":"add","path":"/body","type":"p","props":{"text":"An unordered point","list":"bullet"}},
                  {"op":"add","path":"/body","props":{"text":"Links, bookmark, footnote (M3)","style":"Heading2"}},
                  {"op":"add","path":"/body","type":"p","props":{"text":"This sentence carries two links and a footnote: "}},
                  {"op":"add","path":"/body/p[11]","type":"bookmark","props":{"name":"Lists"}},
                  {"op":"add","path":"/body/p[17]","type":"link","props":{"text":"external site","url":"https://example.com/aioffice"}},
                  {"op":"add","path":"/body/p[17]","type":"link","props":{"text":" / jump to Lists","anchor":"Lists"}},
                  {"op":"add","path":"/body/p[17]","type":"footnote","props":{"text":"Footnote written by aioffice — must appear at the page bottom."}},
                  {"op":"add","path":"/comment[@id=1]","type":"reply","props":{"text":"Reply written by aioffice — must thread under the first comment.","author":"Author"}},
                  {"op":"set","path":"/section[1]","props":{"pageSize":"A4","orientation":"landscape","marginTop":"2cm","marginBottom":"2cm"}}
                ]
                """));
        Assert.True(m3.IsOk, m3.ToJson());

        // The pending tracked change: the human must see an insertion/deletion
        // attributed to "Reviewer" in Word's review pane.
        var tracked = Handler.Edit(
            new CommandContext
            {
                Workspace = workspace,
                File = file,
                Args = new JsonObject { ["track"] = true, ["author"] = "Reviewer" },
            },
            EditOp.ParseBatch("""
                [{"op":"set","path":"/body/p[2]","props":{"text":"Open this file in Word: heading, formatting, table, styles, image, comment + threaded reply, numbered/bulleted lists, links, footnote, landscape A4 pages and ONE PENDING tracked change must look right."}}]
                """));
        Assert.True(tracked.IsOk, tracked.ToJson());

        var revisions = Handler.Read(new CommandContext
        {
            Workspace = workspace,
            File = file,
            Args = new JsonObject { ["view"] = "revisions" },
        });
        Assert.True(revisions.IsOk, revisions.ToJson());

        AssertValidatesClean(file);
    }

    /// <summary>A real, decodable PNG with opaque dark-blue pixels (zlib IDAT), so the human actually sees it.</summary>
    private static byte[] OpaquePng(int width, int height)
    {
        using var png = new MemoryStream();
        png.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);

        var ihdr = new byte[13];
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(ihdr.AsSpan(0, 4), width);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(ihdr.AsSpan(4, 4), height);
        ihdr[8] = 8; // bit depth
        ihdr[9] = 6; // RGBA
        WriteChunk(png, "IHDR", ihdr);

        var raw = new byte[height * (1 + (width * 4))];
        for (var y = 0; y < height; y++)
        {
            var row = y * (1 + (width * 4)) + 1; // skip the filter byte
            for (var x = 0; x < width; x++)
            {
                raw[row + (x * 4) + 0] = 0x1F; // R
                raw[row + (x * 4) + 1] = 0x4E; // G
                raw[row + (x * 4) + 2] = 0x79; // B
                raw[row + (x * 4) + 3] = 0xFF; // A opaque
            }
        }

        using var idat = new MemoryStream();
        using (var zlib = new System.IO.Compression.ZLibStream(idat, System.IO.Compression.CompressionMode.Compress, leaveOpen: true))
        {
            zlib.Write(raw);
        }

        WriteChunk(png, "IDAT", idat.ToArray());
        WriteChunk(png, "IEND", []);
        return png.ToArray();
    }

    private static void WriteChunk(MemoryStream png, string type, byte[] payload)
    {
        var header = new byte[4];
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(header, payload.Length);
        png.Write(header);
        var typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        png.Write(typeBytes);
        png.Write(payload);
        var crcInput = new byte[typeBytes.Length + payload.Length];
        typeBytes.CopyTo(crcInput, 0);
        payload.CopyTo(crcInput, typeBytes.Length);
        var crc = new byte[4];
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32BigEndian(crc, Crc32(crcInput));
        png.Write(crc);
    }

    private static uint Crc32(byte[] data)
    {
        var crc = 0xFFFFFFFFu;
        foreach (var b in data)
        {
            crc ^= b;
            for (var i = 0; i < 8; i++)
            {
                crc = (crc >> 1) ^ (0xEDB88320u & (uint)(-(crc & 1)));
            }
        }

        return crc ^ 0xFFFFFFFFu;
    }

    private static string? FindRepoRoot()
    {
        for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir is not null; dir = dir.Parent)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AIOffice.sln")))
            {
                return dir.FullName;
            }
        }

        return null;
    }
}
