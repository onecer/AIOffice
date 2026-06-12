using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using AIOffice.Core;

namespace AIOffice.Preview;

/// <summary>
/// The wire shape of GET/POST <c>/selection</c>:
/// <c>{"paths":[...canonical...],"rev":"&lt;rev&gt;","updatedAt":iso}</c>.
/// </summary>
public sealed record SelectionSnapshot(
    [property: JsonPropertyName("paths")] IReadOnlyList<string> Paths,
    [property: JsonPropertyName("rev")] string? Rev,
    [property: JsonPropertyName("updatedAt")] DateTimeOffset UpdatedAt);

/// <summary>Thread-safe current selection of one preview server.</summary>
internal sealed class SelectionStore
{
    private readonly object _gate = new();
    private IReadOnlyList<string> _paths = [];
    private DateTimeOffset _updatedAt = DateTimeOffset.UtcNow;

    public (IReadOnlyList<string> Paths, DateTimeOffset UpdatedAt) Read()
    {
        lock (_gate)
        {
            return (_paths, _updatedAt);
        }
    }

    /// <summary>Replaces the whole selection (POST /selection semantics).</summary>
    public void Replace(IReadOnlyList<string> paths)
    {
        lock (_gate)
        {
            _paths = paths;
            _updatedAt = DateTimeOffset.UtcNow;
        }
    }
}

/// <summary>
/// Syntactic validation for selection paths. Accepts the shared
/// <see cref="DocPath"/> grammar plus the pptx id-form
/// (<c>/slide[2]/shape[@id=7]</c>), which is the canonical form the pptx
/// renderer emits but which the core grammar does not cover.
/// </summary>
internal static partial class PreviewPaths
{
    [GeneratedRegex(@"^/slide\[[0-9]+\]/shape\[@id=[0-9]+\](?:/p\[[0-9]+\](?:/run\[[0-9]+\])?)?$")]
    private static partial Regex PptxIdForm();

    /// <summary>Throws <c>invalid_path</c> when the path is not syntactically valid.</summary>
    public static void Validate(string path)
    {
        if (DocPath.TryParse(path, out _, out var error))
        {
            return;
        }

        if (PptxIdForm().IsMatch(path))
        {
            return;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"Not a valid document path: '{path}'. {error}",
            "Selection paths are canonical document paths like /body/p[3], /Sheet1/A1:C10 " +
            "or /slide[2]/shape[@id=7]; indices are 1-based.");
    }
}
