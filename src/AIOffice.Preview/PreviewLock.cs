using System.Net;
using System.Net.Sockets;
using System.Text.Json;
using System.Text.Json.Serialization;
using AIOffice.Core;

namespace AIOffice.Preview;

/// <summary>
/// The on-disk record of one running preview server. Lives at
/// <c>~/.aioffice/preview/&lt;sha12-of-absolute-file-path&gt;.json</c> so the CLI/MCP
/// can find the server for a file without any registry.
/// </summary>
public sealed record PreviewLockfile(
    [property: JsonPropertyName("port")] int Port,
    [property: JsonPropertyName("pid")] int Pid,
    [property: JsonPropertyName("file")] string File,
    [property: JsonPropertyName("startedAt")] DateTimeOffset StartedAt);

/// <summary>Lockfile naming, IO and liveness probes shared by server and client.</summary>
public static class PreviewLock
{
    /// <summary>Default lockfile directory: <c>~/.aioffice/preview</c>.</summary>
    public static string DefaultDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        ".aioffice", "preview");

    /// <summary>
    /// Lockfile path for a document: the first 12 hex chars of the SHA-256 of
    /// its absolute path (the same stamp <see cref="Rev"/> uses), plus .json.
    /// </summary>
    public static string PathFor(string absoluteFile, string? directory = null) =>
        Path.Combine(directory ?? DefaultDirectory, Rev.OfString(absoluteFile) + ".json");

    /// <summary>Reads a lockfile; null when missing or unparseable (treated as stale).</summary>
    public static PreviewLockfile? TryRead(string lockPath)
    {
        try
        {
            if (!System.IO.File.Exists(lockPath))
            {
                return null;
            }

            return JsonSerializer.Deserialize<PreviewLockfile>(
                System.IO.File.ReadAllText(lockPath), JsonDefaults.Options);
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            return null;
        }
    }

    public static void Write(string lockPath, PreviewLockfile lockfile)
    {
        if (Path.GetDirectoryName(lockPath) is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
        }

        System.IO.File.WriteAllText(lockPath, JsonSerializer.Serialize(lockfile, JsonDefaults.Options));
    }

    /// <summary>Best-effort delete; a missing or busy lockfile is not an error.</summary>
    public static void Delete(string lockPath)
    {
        // Windows can briefly report a sharing violation (antivirus/indexer or a
        // handle still settling) right after the server closes, so File.Delete can
        // throw transiently. Retry for a short window so the "lockfile is gone after
        // Close" invariant holds on every platform; if it still fails, the
        // stale-port overwrite on the next start handles the leftover file.
        for (var attempt = 0; ; attempt++)
        {
            try
            {
                System.IO.File.Delete(lockPath);
                return;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                if (attempt >= 9)
                {
                    return;
                }

                System.Threading.Thread.Sleep(20);
            }
        }
    }

    /// <summary>True when something accepts TCP connections on 127.0.0.1:<paramref name="port"/>.</summary>
    public static bool IsPortAlive(int port)
    {
        using var client = new TcpClient();
        try
        {
            return client.ConnectAsync(IPAddress.Loopback, port).Wait(250);
        }
        catch (AggregateException)
        {
            return false; // connection refused -> dead
        }
    }
}
