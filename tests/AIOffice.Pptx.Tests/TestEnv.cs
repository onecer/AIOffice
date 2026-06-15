using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Pptx.Tests;

/// <summary>A throwaway sandboxed workspace per test.</summary>
public sealed class TempWorkspace : IDisposable
{
    public TempWorkspace()
    {
        Dir = Directory.CreateTempSubdirectory("aioffice-pptx-tests-").FullName;
        Workspace = new Workspace(Dir);
    }

    public string Dir { get; }

    public Workspace Workspace { get; }

    /// <summary>Sandbox-resolved absolute path for a file name inside the workspace.</summary>
    public string PathOf(string fileName) => Workspace.Resolve(fileName);

    public CommandContext Ctx(string fileName, params (string Key, JsonNode? Value)[] args)
    {
        var jsonArgs = new JsonObject();
        foreach (var (key, value) in args)
        {
            jsonArgs[key] = value;
        }

        return new CommandContext
        {
            Workspace = Workspace,
            File = Workspace.Resolve(fileName),
            Args = jsonArgs,
        };
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(Dir, recursive: true);
        }
        catch (IOException)
        {
            // Best-effort cleanup of the temp dir.
        }
    }
}

internal static class TestEnv
{
    /// <summary>The envelope data re-serialized to a JsonObject for assertions (camelCase wire shape).</summary>
    public static JsonObject Data(Envelope envelope)
    {
        Assert.NotNull(envelope.Data);
        return JsonSerializer.SerializeToNode(envelope.Data, JsonDefaults.Options)!.AsObject();
    }

    public static JsonObject AssertOk(Envelope envelope)
    {
        Assert.True(envelope.IsOk, envelope.Error is { } e ? $"{e.Code}: {e.Message}" : "envelope not ok");
        return Data(envelope);
    }

    public static ErrorBody AssertFail(Envelope envelope, string expectedCode)
    {
        Assert.False(envelope.IsOk, "expected a failure envelope");
        Assert.NotNull(envelope.Error);
        Assert.Equal(expectedCode, envelope.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(envelope.Error.Suggestion), "error without suggestion");
        return envelope.Error;
    }

    public static EditOp Op(string op, string path, string? type = null, string? position = null, JsonObject? props = null) =>
        new() { Op = op, Path = path, Type = type, Position = position, Props = props };

    public static JsonObject Props(params (string Key, JsonNode? Value)[] pairs)
    {
        var props = new JsonObject();
        foreach (var (key, value) in pairs)
        {
            props[key] = value;
        }

        return props;
    }

    /// <summary>Asserts the OpenXmlValidator reports zero issues via the validate verb.</summary>
    public static void AssertValid(TempWorkspace ws, string fileName)
    {
        var data = AssertOk(new PptxHandler().Validate(ws.Ctx(fileName)));
        Assert.True(
            data["valid"]!.GetValue<bool>(),
            "validator issues: " + data["issues"]!.ToJsonString());
    }
}

/// <summary>Generates small but genuinely valid image files for picture tests.</summary>
internal static class TestImages
{
    /// <summary>A valid RGBA PNG (black pixels) of the given pixel size.</summary>
    public static byte[] Png(int width, int height)
    {
        using var png = new MemoryStream();
        png.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);

        var ihdr = new byte[13];
        BinaryPrimitives.WriteInt32BigEndian(ihdr.AsSpan(0, 4), width);
        BinaryPrimitives.WriteInt32BigEndian(ihdr.AsSpan(4, 4), height);
        ihdr[8] = 8; // bit depth
        ihdr[9] = 6; // RGBA
        WriteChunk(png, "IHDR", ihdr);

        var raw = new byte[height * (1 + (width * 4))]; // filter byte + RGBA per scanline, all zero
        using var idat = new MemoryStream();
        using (var zlib = new ZLibStream(idat, CompressionMode.Compress, leaveOpen: true))
        {
            zlib.Write(raw);
        }

        WriteChunk(png, "IDAT", idat.ToArray());
        WriteChunk(png, "IEND", []);
        return png.ToArray();
    }

    /// <summary>A minimal JPEG (SOI + JFIF APP0 + SOF0 + EOI) carrying the given pixel size.</summary>
    public static byte[] Jpeg(int width, int height)
    {
        using var jpeg = new MemoryStream();
        jpeg.Write([0xFF, 0xD8]); // SOI
        jpeg.Write([
            0xFF, 0xE0, 0x00, 0x10, // APP0, length 16
            (byte)'J', (byte)'F', (byte)'I', (byte)'F', 0x00,
            0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00,
        ]);
        jpeg.Write([
            0xFF, 0xC0, 0x00, 0x11, 0x08, // SOF0, length 17, 8-bit precision
            (byte)(height >> 8), (byte)height,
            (byte)(width >> 8), (byte)width,
            0x03, 0x01, 0x11, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01,
        ]);
        jpeg.Write([0xFF, 0xD9]); // EOI
        return jpeg.ToArray();
    }

    private static void WriteChunk(MemoryStream png, string type, byte[] data)
    {
        Span<byte> length = stackalloc byte[4];
        BinaryPrimitives.WriteInt32BigEndian(length, data.Length);
        png.Write(length);

        var typeAndData = new byte[4 + data.Length];
        Encoding.ASCII.GetBytes(type).CopyTo(typeAndData, 0);
        data.CopyTo(typeAndData, 4);
        png.Write(typeAndData);

        Span<byte> crc = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(crc, Crc32(typeAndData));
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
                crc = (crc >> 1) ^ (0xEDB88320u & (uint)-(crc & 1));
            }
        }

        return crc ^ 0xFFFFFFFF;
    }
}

/// <summary>Generates small but well-formed 3D-model files for model3d tests.</summary>
internal static class TestModels
{
    /// <summary>A minimal binary glTF (.glb): the 12-byte header ("glTF" magic + version + length) plus a small JSON chunk.</summary>
    public static byte[] Glb()
    {
        var json = Encoding.ASCII.GetBytes("{\"asset\":{\"version\":\"2.0\"}}");
        while (json.Length % 4 != 0)
        {
            json = [.. json, (byte)' ']; // chunks are 4-byte aligned
        }

        using var glb = new MemoryStream();
        glb.Write(Encoding.ASCII.GetBytes("glTF"));
        glb.Write(BitConverter.GetBytes(2u)); // version
        glb.Write(BitConverter.GetBytes((uint)(12 + 8 + json.Length))); // total length
        glb.Write(BitConverter.GetBytes((uint)json.Length)); // JSON chunk length
        glb.Write(Encoding.ASCII.GetBytes("JSON"));
        glb.Write(json);
        return glb.ToArray();
    }

    /// <summary>A minimal glTF (.gltf) JSON document.</summary>
    public static byte[] Gltf() =>
        Encoding.ASCII.GetBytes("{\"asset\":{\"version\":\"2.0\"},\"scenes\":[],\"nodes\":[]}");
}
