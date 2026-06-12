using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;

namespace AIOffice.Mcp;

/// <summary>
/// The M4 document-wide find/replace sugar, shared by the CLI
/// (<c>aioffice edit f --find X --replace Y</c>) and MCP (<c>office_edit</c>
/// with a replace op whose path is <c>"/"</c>). A root-scoped replace op is
/// expanded — BEFORE the rev guard and snapshot, like
/// <see cref="CrossDocDataFrom"/> — into one replace op per default scope:
/// <list type="bullet">
/// <item>docx: <c>/body</c> + every <c>/header[i]</c> + <c>/footer[i]</c>
/// (body only under track=true — tracked replace is body-scoped by contract).</item>
/// <item>xlsx: every sheet.</item>
/// <item>pptx: every <c>/slide[i]</c>; slide scope includes the notes by default.</item>
/// </list>
/// After the edit, <see cref="Aggregate"/> folds the per-scope results into one
/// document-level <c>{replacements, locations}</c> pair and collapses the
/// per-scope <c>find_no_match</c> warnings into a single document-level one
/// (or none, when something matched anywhere).
/// </summary>
public static class ReplaceSugar
{
    /// <summary>What a root-scope expansion did — feeds <see cref="Aggregate"/>.</summary>
    public sealed record Expansion(string Find, int ScopeCount);

    private const int MaxLocations = 20;

    /// <summary>
    /// Expands every <c>{"op":"replace","path":"/"}</c> in the batch into the
    /// format's default scopes. Returns the ops unchanged (and a null
    /// <paramref name="expansion"/>) when the batch has no root-scoped replace.
    /// </summary>
    public static IReadOnlyList<EditOp> ExpandDocumentScopes(
        IReadOnlyList<EditOp> ops, DocumentKind kind, string resolvedFile, bool track, out Expansion? expansion)
    {
        expansion = null;
        if (!ops.Any(IsRootReplace))
        {
            return ops;
        }

        var scopes = DefaultScopes(kind, resolvedFile, track);
        if (scopes.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "Document-wide replace found nothing to search: the deck has no slides.",
                "Add a slide first ({\"op\":\"add\",\"path\":\"/\",\"type\":\"slide\"}), or scope the replace explicitly.");
        }

        var expanded = new List<EditOp>(ops.Count + scopes.Count - 1);
        string find = string.Empty;
        var scopeCount = 0;
        foreach (var op in ops)
        {
            if (!IsRootReplace(op))
            {
                expanded.Add(op);
                continue;
            }

            if (find.Length == 0 &&
                op.Props?["find"] is JsonValue findValue && findValue.TryGetValue<string>(out var findText))
            {
                find = findText;
            }

            foreach (var scope in scopes)
            {
                expanded.Add(op with { Path = scope, Props = op.Props?.DeepClone().AsObject() });
                scopeCount++;
            }
        }

        expansion = new Expansion(find, scopeCount);
        return expanded;
    }

    /// <summary>
    /// Folds the per-scope replace results of an expanded batch into one
    /// document-level aggregate: <c>data.replacements</c> (total) and
    /// <c>data.locations</c> (first 20 canonical paths). Per-scope
    /// <c>find_no_match</c> warnings collapse into a single document-level
    /// warning when NOTHING matched, and disappear when anything did.
    /// </summary>
    public static Envelope Aggregate(Envelope envelope, Expansion expansion)
    {
        if (!envelope.IsOk || envelope.Data is null)
        {
            return envelope;
        }

        if (JsonSerializer.SerializeToNode(envelope.Data, JsonDefaults.Options) is not JsonObject data)
        {
            return envelope;
        }

        var total = 0;
        var locations = new JsonArray();
        foreach (var (_, value) in data)
        {
            if (value is not JsonArray results)
            {
                continue;
            }

            foreach (var item in results)
            {
                // Per-op replace results are the only objects carrying both keys
                // ({replacements:N, locations:[paths]}), across all three formats.
                if (item is not JsonObject result ||
                    result["replacements"] is not JsonValue count || !count.TryGetValue<int>(out var n) ||
                    result["locations"] is not JsonArray scopeLocations)
                {
                    continue;
                }

                total += n;
                foreach (var location in scopeLocations)
                {
                    if (locations.Count < MaxLocations && location is not null)
                    {
                        locations.Add(location.DeepClone());
                    }
                }
            }
        }

        data["replacements"] = total;
        data["locations"] = locations;

        var warnings = (envelope.Meta.Warnings ?? [])
            .Where(w => w.Code != ErrorCodes.FindNoMatch)
            .ToList();
        if (total == 0)
        {
            warnings.Add(new Warning(
                ErrorCodes.FindNoMatch,
                $"'{expansion.Find}' matched nothing in any of the {expansion.ScopeCount} document scope(s) searched; 0 replacements made."));
        }

        return envelope with
        {
            Data = data,
            Meta = envelope.Meta with { Warnings = warnings.Count > 0 ? warnings : null },
        };
    }

    // "/" never parses as a DocPath (paths need >= 1 segment); ParseBatch
    // carves it out for replace ops so this marker can reach the expansion.
    private static bool IsRootReplace(EditOp op) =>
        op.Op == "replace" && op.Path == "/";

    /// <summary>The format's default replace scopes, enumerated read-only and cheaply.</summary>
    private static IReadOnlyList<string> DefaultScopes(DocumentKind kind, string resolvedFile, bool track)
    {
        try
        {
            return kind switch
            {
                DocumentKind.Docx => DocxScopes(resolvedFile, track),
                DocumentKind.Xlsx => XlsxScopes(resolvedFile),
                DocumentKind.Pptx => PptxScopes(resolvedFile),
                _ => throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"Document-wide replace is not available for '{kind}'.",
                    "Scope the replace op to an explicit container path instead."),
            };
        }
        catch (Exception ex) when (ex is not AiofficeException)
        {
            throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                $"Could not enumerate the document's replace scopes: {ex.Message}",
                "Run 'aioffice validate' on the file; restore a snapshot if it is corrupt.",
                innerException: ex);
        }
    }

    private static List<string> DocxScopes(string file, bool track)
    {
        var scopes = new List<string> { "/body" };
        if (track)
        {
            // Tracked replace is body-only by the handler contract
            // (headers/footers cannot carry w:ins/w:del here).
            return scopes;
        }

        using var doc = WordprocessingDocument.Open(file, false);
        var main = doc.MainDocumentPart;
        var headers = main?.HeaderParts.Count() ?? 0;
        var footers = main?.FooterParts.Count() ?? 0;
        for (var i = 1; i <= headers; i++)
        {
            scopes.Add($"/header[{i}]");
        }

        for (var i = 1; i <= footers; i++)
        {
            scopes.Add($"/footer[{i}]");
        }

        return scopes;
    }

    private static List<string> XlsxScopes(string file)
    {
        using var doc = SpreadsheetDocument.Open(file, false);
        var sheets = doc.WorkbookPart?.Workbook?.Sheets?
            .Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .Select(s => s.Name?.Value)
            .OfType<string>()
            .ToList() ?? [];

        // Always-quoted parse form; DocPath canonicalization in the handlers
        // drops the quotes again for names that are safe bare.
        return [.. sheets.Select(name => "/'" + name.Replace("'", "''", StringComparison.Ordinal) + "'")];
    }

    private static List<string> PptxScopes(string file)
    {
        using var doc = PresentationDocument.Open(file, false);
        var slides = doc.PresentationPart?.Presentation?.SlideIdList?
            .Elements<DocumentFormat.OpenXml.Presentation.SlideId>()
            .Count() ?? 0;
        return [.. Enumerable.Range(1, slides).Select(i => $"/slide[{i}]")];
    }
}
