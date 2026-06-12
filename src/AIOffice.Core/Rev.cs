using System.Security.Cryptography;
using System.Text;

namespace AIOffice.Core;

/// <summary>
/// Content revision stamps: the first 12 hex chars (lowercase) of SHA-256.
/// Used in envelope meta and for optimistic concurrency (edit --expect-rev).
/// </summary>
public static class Rev
{
    public const int Length = 12;

    public static string OfBytes(ReadOnlySpan<byte> bytes)
    {
        Span<byte> hash = stackalloc byte[SHA256.HashSizeInBytes];
        SHA256.HashData(bytes, hash);
        return Convert.ToHexStringLower(hash)[..Length];
    }

    public static string OfString(string text) => OfBytes(Encoding.UTF8.GetBytes(text));

    public static string OfFile(string path)
    {
        if (!File.Exists(path))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"File not found: {path}",
                "Check the path spelling, or run 'aioffice create' to make a new document.");
        }

        using var stream = File.OpenRead(path);
        Span<byte> hash = stackalloc byte[SHA256.HashSizeInBytes];
        SHA256.HashData(stream, hash);
        return Convert.ToHexStringLower(hash)[..Length];
    }
}
