using System.Globalization;

namespace AIOffice.Core;

/// <summary>One entry in a file's snapshot ring.</summary>
public sealed record SnapshotEntry(int Number, string Path, long SizeBytes, DateTimeOffset CreatedUtc, string Rev);

/// <summary>
/// Automatic pre-edit snapshot ring. Snapshots live OUTSIDE the workspace
/// (default <c>~/.aioffice/snapshots/&lt;path-hash&gt;/NNN.snap</c>) so edits can
/// never clobber their own undo history. The ring keeps the most recent
/// <see cref="Capacity"/> snapshots per file; numbers grow monotonically.
/// </summary>
public sealed class SnapshotStore
{
    public const int Capacity = 20;

    private readonly string _baseDir;

    /// <param name="baseDir">Override for tests; defaults to <c>~/.aioffice/snapshots</c>.</param>
    public SnapshotStore(string? baseDir = null)
    {
        _baseDir = baseDir ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".aioffice", "snapshots");
    }

    /// <summary>Directory holding the ring for one original file path.</summary>
    public string RingDirectory(string filePath) =>
        Path.Combine(_baseDir, Rev.OfString(Path.GetFullPath(filePath)));

    /// <summary>Saves a snapshot of the file's current bytes; evicts the oldest beyond capacity.</summary>
    public SnapshotEntry Save(string filePath)
    {
        var full = Path.GetFullPath(filePath);
        if (!File.Exists(full))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"Cannot snapshot a missing file: {filePath}",
                "Snapshots capture the pre-edit state of an existing file; create the file first.");
        }

        var dir = RingDirectory(full);
        Directory.CreateDirectory(dir);

        var existing = ListNumbers(dir);
        var next = existing.Count == 0 ? 1 : existing[^1] + 1;
        var snapPath = Path.Combine(dir, FormatName(next));
        File.Copy(full, snapPath, overwrite: false);

        // Evict oldest entries beyond capacity.
        existing.Add(next);
        for (var i = 0; i < existing.Count - Capacity; i++)
        {
            File.Delete(Path.Combine(dir, FormatName(existing[i])));
        }

        var info = new FileInfo(snapPath);
        return new SnapshotEntry(next, snapPath, info.Length, info.LastWriteTimeUtc, Rev.OfFile(snapPath));
    }

    /// <summary>Lists the ring for a file, oldest first. Empty when no snapshots exist.</summary>
    public IReadOnlyList<SnapshotEntry> List(string filePath)
    {
        var dir = RingDirectory(Path.GetFullPath(filePath));
        if (!Directory.Exists(dir))
        {
            return [];
        }

        var entries = new List<SnapshotEntry>();
        foreach (var n in ListNumbers(dir))
        {
            var snapPath = Path.Combine(dir, FormatName(n));
            var info = new FileInfo(snapPath);
            entries.Add(new SnapshotEntry(n, snapPath, info.Length, info.LastWriteTimeUtc, Rev.OfFile(snapPath)));
        }

        return entries;
    }

    /// <summary>
    /// Restores snapshot <paramref name="number"/> (or the newest when null)
    /// over the original file. The current state is snapshotted first so a
    /// restore is itself undoable.
    /// </summary>
    public SnapshotEntry Restore(string filePath, int? number = null)
    {
        var full = Path.GetFullPath(filePath);
        var ring = List(full);
        if (ring.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"No snapshots exist for: {filePath}",
                "Snapshots are taken automatically before every successful edit; edit the file first or check the path.");
        }

        SnapshotEntry? target = number is null
            ? ring[^1]
            : ring.FirstOrDefault(e => e.Number == number);
        if (target is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Snapshot {number} does not exist for: {filePath}",
                $"Run 'aioffice snapshot list {filePath}' to see available snapshot numbers.",
                candidates: [.. ring.Select(e => e.Number.ToString(CultureInfo.InvariantCulture))]);
        }

        if (File.Exists(full))
        {
            Save(full); // make the restore undoable
        }

        File.Copy(target.Path, full, overwrite: true);
        return target;
    }

    private static string FormatName(int number) =>
        number.ToString("D3", CultureInfo.InvariantCulture) + ".snap";

    private static List<int> ListNumbers(string dir)
    {
        var numbers = new List<int>();
        foreach (var file in Directory.EnumerateFiles(dir, "*.snap"))
        {
            var stem = Path.GetFileNameWithoutExtension(file);
            if (int.TryParse(stem, NumberStyles.None, CultureInfo.InvariantCulture, out var n))
            {
                numbers.Add(n);
            }
        }

        numbers.Sort();
        return numbers;
    }
}
