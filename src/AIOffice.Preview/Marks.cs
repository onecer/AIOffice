using System.Text.Json.Serialization;
using AIOffice.Core;

namespace AIOffice.Preview;

/// <summary>
/// One advisory mark: an in-memory highlight an agent (or human) puts on an
/// addressable element to direct attention — "this overflows, fix it". Marks
/// never touch the document; they live only in the running preview server and
/// are pushed to every viewer over SSE. Mirrors OfficeCLI's <c>watch mark</c>:
/// a path, an optional color/note, an optional <c>find</c> substring to
/// highlight inside the element, and a <c>toFix</c> flag.
/// </summary>
public sealed record Mark(
    [property: JsonPropertyName("path")] string Path,
    [property: JsonPropertyName("color")] string? Color = null,
    [property: JsonPropertyName("note")] string? Note = null,
    [property: JsonPropertyName("find")] string? Find = null,
    [property: JsonPropertyName("toFix")] bool ToFix = false);

/// <summary>
/// The wire shape of GET/POST/DELETE <c>/marks</c>:
/// <c>{"marks":[...],"rev":"&lt;rev&gt;","updatedAt":iso}</c>.
/// </summary>
public sealed record MarksSnapshot(
    [property: JsonPropertyName("marks")] IReadOnlyList<Mark> Marks,
    [property: JsonPropertyName("rev")] string? Rev,
    [property: JsonPropertyName("updatedAt")] DateTimeOffset UpdatedAt);

/// <summary>Thread-safe set of advisory marks for one preview server, keyed (upsert) by path.</summary>
internal sealed class MarkStore
{
    private readonly object _gate = new();
    private readonly List<Mark> _marks = [];
    private DateTimeOffset _updatedAt = DateTimeOffset.UtcNow;

    public (IReadOnlyList<Mark> Marks, DateTimeOffset UpdatedAt) Read()
    {
        lock (_gate)
        {
            return ([.. _marks], _updatedAt);
        }
    }

    /// <summary>Adds or replaces the mark on <paramref name="mark"/>'s path (one mark per path).</summary>
    public void Add(Mark mark)
    {
        lock (_gate)
        {
            _marks.RemoveAll(m => string.Equals(m.Path, mark.Path, StringComparison.Ordinal));
            _marks.Add(mark);
            _updatedAt = DateTimeOffset.UtcNow;
        }
    }

    /// <summary>Removes the mark on <paramref name="path"/>. Returns whether one was removed.</summary>
    public bool Remove(string path)
    {
        lock (_gate)
        {
            var removed = _marks.RemoveAll(m => string.Equals(m.Path, path, StringComparison.Ordinal)) > 0;
            if (removed)
            {
                _updatedAt = DateTimeOffset.UtcNow;
            }

            return removed;
        }
    }

    /// <summary>Drops every mark. Returns how many were cleared.</summary>
    public int Clear()
    {
        lock (_gate)
        {
            var count = _marks.Count;
            _marks.Clear();
            if (count > 0)
            {
                _updatedAt = DateTimeOffset.UtcNow;
            }

            return count;
        }
    }
}

/// <summary>Validates and normalizes a mark color: <c>#RGB/#RRGGBB/#RRGGBBAA</c> hex (bare hex auto-prefixed) or a named color.</summary>
internal static class MarkColor
{
    public const string Default = "#ffeb3b";

    private static readonly HashSet<string> Named = new(StringComparer.OrdinalIgnoreCase)
    {
        "red", "orange", "yellow", "green", "teal", "blue", "indigo", "purple",
        "pink", "brown", "gray", "grey", "black", "white", "cyan", "magenta",
        "lime", "amber", "coral", "gold", "salmon", "violet",
    };

    /// <summary>Returns a CSS-safe color, or throws <c>invalid_args</c>. Null/empty → the default highlight.</summary>
    public static string Normalize(string? color)
    {
        if (string.IsNullOrWhiteSpace(color))
        {
            return Default;
        }

        var value = color.Trim();
        if (Named.Contains(value))
        {
            return value.ToLowerInvariant();
        }

        var hex = value.StartsWith('#') ? value[1..] : value;
        if ((hex.Length is 3 or 6 or 8) && hex.All(Uri.IsHexDigit))
        {
            return "#" + hex;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Not a valid mark color: '{color}'.",
            "Use a hex color (#ffeb3b, #f33, #11223344) or a named color (red, blue, yellow, …).");
    }
}
