using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    [GeneratedRegex(@"\{\{\s*([A-Za-z0-9_.-]+)\s*\}\}")]
    private static partial Regex PlaceholderRegex();

    public Envelope Template(CommandContext ctx)
    {
        var file = RequireFile(ctx, mustExist: true);
        var data = TemplateData(ctx.Args);

        var ms = new MemoryStream();
        var originalBytes = File.ReadAllBytes(file);
        ms.Write(originalBytes);
        ms.Position = 0;

        var replaced = 0;
        var resolvedKeys = new SortedSet<string>(StringComparer.Ordinal);
        var unresolvedKeys = new SortedSet<string>(StringComparer.Ordinal);

        using (var doc = OpenPackage(ms, file, editable: true))
        {
            var paragraphs = new List<Paragraph>();
            if (doc.MainDocumentPart?.Document is { } document)
            {
                paragraphs.AddRange(document.Descendants<Paragraph>());
            }

            foreach (var header in doc.MainDocumentPart?.HeaderParts ?? [])
            {
                if (header.Header is { } h)
                {
                    paragraphs.AddRange(h.Descendants<Paragraph>());
                }
            }

            foreach (var footer in doc.MainDocumentPart?.FooterParts ?? [])
            {
                if (footer.Footer is { } f)
                {
                    paragraphs.AddRange(f.Descendants<Paragraph>());
                }
            }

            foreach (var paragraph in paragraphs)
            {
                replaced += MergeParagraph(paragraph, data, resolvedKeys, unresolvedKeys);
            }

            // The same data map fills MERGEFIELD complex fields by name, alongside
            // the {{key}} markers; unresolved field names join the same report.
            replaced += MergeMergeFields(doc, data, resolvedKeys, unresolvedKeys);
        }

        var newBytes = ms.ToArray();

        string target;
        if (StringArg(ctx.Args, "output") is { } outArg)
        {
            target = ctx.Workspace.Resolve(outArg);
            var dir = Path.GetDirectoryName(target);
            if (dir is { Length: > 0 })
            {
                Directory.CreateDirectory(dir);
            }
        }
        else
        {
            target = file;
            _snapshots.Save(file); // in-place merge is undoable
        }

        File.WriteAllBytes(target, newBytes);

        var warnings = unresolvedKeys.Count > 0
            ? new List<Warning>
            {
                new("unresolved_keys", $"No data for: {string.Join(", ", unresolvedKeys)}; left as-is."),
            }
            : null;

        return Envelope.Ok(
            new
            {
                replaced,
                keys = resolvedKeys.ToList(),
                unresolved = unresolvedKeys.ToList(),
                written = target,
            },
            MetaFor(target, Rev.OfBytes(newBytes), warnings));
    }

    private static Dictionary<string, string> TemplateData(JsonObject args)
    {
        var node = args["data"];
        if (node is JsonValue v && v.TryGetValue<string>(out var raw))
        {
            try
            {
                node = JsonNode.Parse(raw);
            }
            catch (System.Text.Json.JsonException ex)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"--data is not valid JSON: {ex.Message}",
                    "Pass a JSON object like '{\"name\":\"Ada\"}' or @data.json.",
                    innerException: ex);
            }
        }

        if (node is not JsonObject obj || obj.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "template requires --data with a non-empty JSON object.",
                "Pass key/value pairs matching the {{placeholders}}, e.g. --data '{\"name\":\"Ada\"}'.");
        }

        return obj.ToDictionary(kv => kv.Key, kv => NodeToString(kv.Value), StringComparer.Ordinal);
    }

    /// <summary>
    /// Replaces {{key}} placeholders in one paragraph, correctly handling
    /// placeholders split across runs: the replacement lands in the text node
    /// where the placeholder starts; the rest of its characters are cut from
    /// the following text nodes, so all other formatting survives.
    /// </summary>
    private static int MergeParagraph(
        Paragraph paragraph,
        IReadOnlyDictionary<string, string> data,
        ISet<string> resolvedKeys,
        ISet<string> unresolvedKeys)
    {
        var texts = paragraph.Descendants<Text>().ToList();
        if (texts.Count == 0)
        {
            return 0;
        }

        var full = string.Concat(texts.Select(t => t.Text));
        if (!full.Contains("{{", StringComparison.Ordinal))
        {
            return 0;
        }

        var replaced = 0;

        // Right-to-left so earlier match offsets stay valid after replacement.
        foreach (var match in PlaceholderRegex().Matches(full).Reverse())
        {
            var key = match.Groups[1].Value;
            if (!data.TryGetValue(key, out var value))
            {
                unresolvedKeys.Add(key);
                continue;
            }

            ReplaceTextRange(texts, match.Index, match.Length, value);
            resolvedKeys.Add(key);
            replaced++;
        }

        return replaced;
    }

    /// <summary>Splices <paramref name="replacement"/> over [start, start+length) of the concatenated text nodes.</summary>
    private static void ReplaceTextRange(IReadOnlyList<Text> texts, int start, int length, string replacement)
    {
        var end = start + length;
        var offset = 0;
        var inserted = false;

        foreach (var text in texts)
        {
            var tStart = offset;
            var tEnd = offset + text.Text.Length;
            offset = tEnd;

            if (tEnd <= start || tStart >= end)
            {
                continue; // not touched by this placeholder
            }

            var cutFrom = Math.Max(start, tStart) - tStart;
            var cutTo = Math.Min(end, tEnd) - tStart;
            var insert = inserted ? string.Empty : replacement;
            inserted = true;

            var updated = text.Text[..cutFrom] + insert + text.Text[cutTo..];
            text.Text = updated;
            if (updated.Length > 0 && (char.IsWhiteSpace(updated[0]) || char.IsWhiteSpace(updated[^1])))
            {
                text.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
            }
        }
    }
}
