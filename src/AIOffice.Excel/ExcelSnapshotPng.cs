using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;

namespace AIOffice.Excel;

/// <summary>
/// (1.7) A tiny, dependency-free PNG snapshot renderer for the linked-picture
/// (Excel "camera tool") feature. Paints a cell-text grid into a 24-bit RGB PNG
/// using a built-in 5x7 ASCII bitmap font, then encodes it with the BCL's
/// <see cref="ZLibStream"/> (a plain 8-bit-per-channel non-interlaced PNG — the
/// format <see cref="ExcelImages.Sniff"/> already accepts). Fully deterministic
/// across platforms: same input grid → identical bytes (no fonts, no GDI, no
/// floating point), which keeps the CI image-bytes reproducible.
///
/// This is an HONEST static snapshot, not a live link: it captures the range's
/// values at edit time. The caller raises a warning saying so.
/// </summary>
internal static class ExcelSnapshotPng
{
    private const int CellWidth = 96;   // px per column block
    private const int CellHeight = 20;  // px per row block
    private const int Padding = 4;      // text inset inside a cell
    private const int GlyphW = 5;
    private const int GlyphH = 7;
    private const int MaxWidth = 2048;  // clamp the bitmap so a huge range stays sane
    private const int MaxHeight = 2048;

    // Colors (RGB).
    private static readonly byte[] Background = [0xFF, 0xFF, 0xFF];
    private static readonly byte[] Gridline = [0xD0, 0xD0, 0xD0];
    private static readonly byte[] TextColor = [0x20, 0x20, 0x20];

    /// <summary>Renders a row-major grid of cell text to PNG bytes (at least 1x1 cell).</summary>
    public static byte[] Render(IReadOnlyList<IReadOnlyList<string>> grid)
    {
        var rows = Math.Max(1, grid.Count);
        var cols = Math.Max(1, grid.Count > 0 ? grid.Max(r => r.Count) : 1);

        var width = Math.Min(MaxWidth, cols * CellWidth);
        var height = Math.Min(MaxHeight, rows * CellHeight);
        var pixels = new byte[width * height * 3];

        // Fill background.
        for (var i = 0; i < pixels.Length; i += 3)
        {
            pixels[i] = Background[0];
            pixels[i + 1] = Background[1];
            pixels[i + 2] = Background[2];
        }

        // Gridlines (right/bottom edges of every cell block).
        for (var c = 1; c <= cols; c++)
        {
            var x = (c * CellWidth) - 1;
            if (x < width)
            {
                for (var y = 0; y < height; y++)
                {
                    SetPixel(pixels, width, x, y, Gridline);
                }
            }
        }

        for (var r = 1; r <= rows; r++)
        {
            var y = (r * CellHeight) - 1;
            if (y < height)
            {
                for (var x = 0; x < width; x++)
                {
                    SetPixel(pixels, width, x, y, Gridline);
                }
            }
        }

        // Cell text.
        for (var r = 0; r < grid.Count; r++)
        {
            var row = grid[r];
            for (var c = 0; c < row.Count; c++)
            {
                DrawText(pixels, width, height, c, r, row[c]);
            }
        }

        return Encode(pixels, width, height);
    }

    private static void DrawText(byte[] pixels, int width, int height, int col, int row, string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return;
        }

        var originX = (col * CellWidth) + Padding;
        var originY = (row * CellHeight) + Padding;
        var maxGlyphs = (CellWidth - (2 * Padding)) / (GlyphW + 1);
        if (maxGlyphs < 1)
        {
            return;
        }

        var clipped = text.Length > maxGlyphs ? text[..maxGlyphs] : text;
        for (var i = 0; i < clipped.Length; i++)
        {
            DrawGlyph(pixels, width, height, originX + (i * (GlyphW + 1)), originY, clipped[i]);
        }
    }

    private static void DrawGlyph(byte[] pixels, int width, int height, int x0, int y0, char ch)
    {
        var glyph = Font.Glyph(ch);
        for (var gy = 0; gy < GlyphH; gy++)
        {
            var bits = glyph[gy];
            for (var gx = 0; gx < GlyphW; gx++)
            {
                if ((bits & (1 << (GlyphW - 1 - gx))) == 0)
                {
                    continue;
                }

                var x = x0 + gx;
                var y = y0 + gy;
                if (x >= 0 && x < width && y >= 0 && y < height)
                {
                    SetPixel(pixels, width, x, y, TextColor);
                }
            }
        }
    }

    private static void SetPixel(byte[] pixels, int width, int x, int y, byte[] rgb)
    {
        var offset = ((y * width) + x) * 3;
        pixels[offset] = rgb[0];
        pixels[offset + 1] = rgb[1];
        pixels[offset + 2] = rgb[2];
    }

    // ----- PNG encoding -------------------------------------------------------

    private static byte[] Encode(byte[] rgb, int width, int height)
    {
        // Filter each scanline with filter type 0 (None), prefixed per row.
        var stride = width * 3;
        var raw = new byte[(stride + 1) * height];
        for (var y = 0; y < height; y++)
        {
            raw[y * (stride + 1)] = 0; // filter type None
            Array.Copy(rgb, y * stride, raw, (y * (stride + 1)) + 1, stride);
        }

        byte[] compressed;
        using (var output = new MemoryStream())
        {
            using (var deflate = new ZLibStream(output, CompressionLevel.Optimal, leaveOpen: true))
            {
                deflate.Write(raw, 0, raw.Length);
            }

            compressed = output.ToArray();
        }

        using var png = new MemoryStream();
        png.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A], 0, 8); // signature

        // IHDR
        var ihdr = new byte[13];
        BinaryPrimitives.WriteInt32BigEndian(ihdr.AsSpan(0), width);
        BinaryPrimitives.WriteInt32BigEndian(ihdr.AsSpan(4), height);
        ihdr[8] = 8;  // bit depth
        ihdr[9] = 2;  // color type 2 = truecolor RGB
        ihdr[10] = 0; // compression
        ihdr[11] = 0; // filter
        ihdr[12] = 0; // interlace
        WriteChunk(png, "IHDR", ihdr);
        WriteChunk(png, "IDAT", compressed);
        WriteChunk(png, "IEND", []);
        return png.ToArray();
    }

    private static void WriteChunk(Stream stream, string type, byte[] data)
    {
        Span<byte> length = stackalloc byte[4];
        BinaryPrimitives.WriteInt32BigEndian(length, data.Length);
        stream.Write(length);

        var typeBytes = Encoding.ASCII.GetBytes(type);
        stream.Write(typeBytes, 0, typeBytes.Length);
        stream.Write(data, 0, data.Length);

        var crcInput = new byte[typeBytes.Length + data.Length];
        Array.Copy(typeBytes, crcInput, typeBytes.Length);
        Array.Copy(data, 0, crcInput, typeBytes.Length, data.Length);
        Span<byte> crc = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(crc, Crc32(crcInput));
        stream.Write(crc);
    }

    private static readonly uint[] CrcTable = BuildCrcTable();

    private static uint[] BuildCrcTable()
    {
        var table = new uint[256];
        for (uint n = 0; n < 256; n++)
        {
            var c = n;
            for (var k = 0; k < 8; k++)
            {
                c = (c & 1) != 0 ? 0xEDB88320 ^ (c >> 1) : c >> 1;
            }

            table[n] = c;
        }

        return table;
    }

    private static uint Crc32(byte[] data)
    {
        var crc = 0xFFFFFFFFu;
        foreach (var b in data)
        {
            crc = CrcTable[(crc ^ b) & 0xFF] ^ (crc >> 8);
        }

        return crc ^ 0xFFFFFFFFu;
    }
}
