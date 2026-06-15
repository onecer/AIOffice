using System.Globalization;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    [GeneratedRegex(@"\{\{\s*([A-Za-z0-9_.-]+)\s*\}\}")]
    private static partial Regex PlaceholderRegex();

    /// <summary>{n} (1-based record index) and {FieldName} substitutions in an --output pattern.</summary>
    [GeneratedRegex(@"\{([A-Za-z0-9_.-]+)\}")]
    private static partial Regex OutputPatternRegex();

    public Envelope Template(CommandContext ctx)
    {
        var node = TemplateDataNode(ctx.Args);

        // An array of records runs a mail merge (v1.4.0); a single object keeps the
        // original one-document fill unchanged.
        if (node is JsonArray records)
        {
            return MailMerge(ctx, records);
        }

        var data = TemplateData(node);
        return FillOne(ctx, data);
    }

    // ----------------------------------------------------------- single object

    /// <summary>Fills {{key}}, MERGEFIELD and «IF» fields in one document from a single data map.</summary>
    private Envelope FillOne(CommandContext ctx, Dictionary<string, string> data)
    {
        var file = RequireFile(ctx, mustExist: true);
        var originalBytes = File.ReadAllBytes(file);

        var (newBytes, _, unresolvedKeys, resolvedKeys) = MergeDocument(originalBytes, file, data);

        string target;
        if (StringArg(ctx.Args, "output") is { } outArg)
        {
            target = ctx.Workspace.Resolve(outArg);
            EnsureDirectory(target);
        }
        else
        {
            target = file;
            _snapshots.Save(file); // in-place merge is undoable
        }

        File.WriteAllBytes(target, newBytes);

        var replaced = resolvedKeys.Count;
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

    // --------------------------------------------------------------- mail merge

    /// <summary>
    /// Mail-merge execution (v1.4.0): one output document per record when
    /// <c>--output</c> gives a path pattern ({n}=1-based index, {Field}=record
    /// value), or a single combined document (the source body repeated per record,
    /// separated by next-page section breaks) when no pattern is given. Every
    /// output path is sandbox-resolved. Returns {records, produced:[paths]}; any
    /// unresolved fields raise a <c>template_unresolved</c> warning.
    /// </summary>
    private Envelope MailMerge(CommandContext ctx, JsonArray records)
    {
        var file = RequireFile(ctx, mustExist: true);
        var originalBytes = File.ReadAllBytes(file);

        var data = records
            .Select((r, i) => RecordData(r, i))
            .ToList();
        if (data.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "template --data was an empty array; a mail merge needs at least one record.",
                "Pass an array of record objects, e.g. --data '[{\"Name\":\"Ada\"},{\"Name\":\"Grace\"}]'.");
        }

        return StringArg(ctx.Args, "output") is { } pattern
            ? MergePerRecord(ctx, file, originalBytes, data, pattern)
            : MergeCombined(ctx, file, originalBytes, data);
    }

    /// <summary>One merged document per record, named from the --output pattern.</summary>
    private Envelope MergePerRecord(
        CommandContext ctx, string file, byte[] originalBytes,
        List<Dictionary<string, string>> records, string pattern)
    {
        var produced = new List<string>();
        var unresolvedAll = new SortedSet<string>(StringComparer.Ordinal);
        byte[]? lastBytes = null;
        var seenPaths = new HashSet<string>(StringComparer.Ordinal);

        for (var i = 0; i < records.Count; i++)
        {
            var record = records[i];
            var relative = ExpandOutputPattern(pattern, record, i + 1);

            // Sandbox-resolve EVERY output path; an escaping pattern is denied.
            var target = ctx.Workspace.Resolve(relative);
            if (!seenPaths.Add(target))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"The --output pattern '{pattern}' produced the same path twice ({relative}).",
                    "Include {n} (the 1-based record index) or a unique record field so each record gets its own file.");
            }

            EnsureDirectory(target);

            var (bytes, _, unresolved, _) = MergeDocument(originalBytes, file, record);
            File.WriteAllBytes(target, bytes);
            produced.Add(target);
            lastBytes = bytes;
            foreach (var key in unresolved)
            {
                unresolvedAll.Add(key);
            }
        }

        return MailMergeEnvelope(produced, records.Count, unresolvedAll, lastBytes!);
    }

    /// <summary>A single combined document: the source body repeated per record, split by next-page breaks.</summary>
    private Envelope MergeCombined(
        CommandContext ctx, string file, byte[] originalBytes, List<Dictionary<string, string>> records)
    {
        var unresolvedAll = new SortedSet<string>(StringComparer.Ordinal);

        // Merge each record into its own copy, then concatenate the bodies into the
        // first copy with a next-page section break between successive records.
        var ms = new MemoryStream();
        ms.Write(MergeDocument(originalBytes, file, records[0]).Bytes);
        ms.Position = 0;

        using (var combined = OpenPackage(ms, file, editable: true))
        {
            var combinedBody = combined.MainDocumentPart?.Document?.Body
                ?? throw new AiofficeException(
                    ErrorCodes.FormatCorrupt,
                    $"The document has no body: {file}",
                    "The main document part is missing or empty. Re-export the file from Word.");

            // Record the first record's unresolved fields.
            CollectUnresolved(originalBytes, file, records[0], unresolvedAll);

            for (var i = 1; i < records.Count; i++)
            {
                var (recordBytes, _, unresolved, _) = MergeDocument(originalBytes, file, records[i]);
                foreach (var key in unresolved)
                {
                    unresolvedAll.Add(key);
                }

                AppendRecordSection(combinedBody, recordBytes, file);
            }
        }

        var newBytes = ms.ToArray();

        // Default destination: back to the source file (snapshotted, undoable).
        _snapshots.Save(file);
        File.WriteAllBytes(file, newBytes);

        return MailMergeEnvelope([file], records.Count, unresolvedAll, newBytes);
    }

    /// <summary>
    /// Appends a merged record's body content to the combined body, ending the
    /// previous section with a next-page break first so each record starts on a
    /// fresh page. The trailing body-level sectPr of the appended content is dropped
    /// (only the combined document keeps the final sectPr).
    /// </summary>
    private static void AppendRecordSection(Body combinedBody, byte[] recordBytes, string file)
    {
        using var ms = new MemoryStream(recordBytes, writable: false);
        using var recordDoc = OpenPackage(ms, file, editable: false);
        var recordBody = recordDoc.MainDocumentPart?.Document?.Body;
        if (recordBody is null)
        {
            return;
        }

        // The combined body's trailing sectPr (or a materialized one) becomes a
        // paragraph-level next-page break that ends the preceding record's section.
        var bodySectPr = combinedBody.Elements<SectionProperties>().FirstOrDefault();
        if (bodySectPr is null)
        {
            bodySectPr = new SectionProperties();
            combinedBody.AppendChild(bodySectPr);
        }

        var breakSectPr = (SectionProperties)bodySectPr.CloneNode(true);
        breakSectPr.RemoveAllChildren<SectionType>(); // nextPage is the default (no w:type)
        var breakParagraph = new Paragraph(new ParagraphProperties(breakSectPr));
        bodySectPr.InsertBeforeSelf(breakParagraph);

        // Import every block of the record's body except its own trailing sectPr,
        // inserting before the combined body's final sectPr so it stays last.
        foreach (var child in recordBody.ChildElements)
        {
            if (child is SectionProperties)
            {
                continue; // the record's final sectPr is not carried over
            }

            bodySectPr.InsertBeforeSelf(child.CloneNode(true));
        }
    }

    private static Envelope MailMergeEnvelope(
        List<string> produced, int records, SortedSet<string> unresolved, byte[] revBytes)
    {
        var warnings = unresolved.Count > 0
            ? new List<Warning>
            {
                new(WarningCodes.TemplateUnresolved,
                    $"Unresolved field(s) across the merge: {string.Join(", ", unresolved)}; left as-is."),
            }
            : null;

        // meta.file is the first produced document (the combined doc, or letter 1).
        return Envelope.Ok(
            new
            {
                records,
                produced,
                unresolved = unresolved.ToList(),
            },
            MetaFor(produced[0], Rev.OfBytes(revBytes), warnings));
    }

    /// <summary>Runs a record merge purely to capture which fields it leaves unresolved.</summary>
    private static void CollectUnresolved(
        byte[] originalBytes, string file, Dictionary<string, string> record, ISet<string> into)
    {
        var (_, _, unresolved, _) = MergeDocument(originalBytes, file, record);
        foreach (var key in unresolved)
        {
            into.Add(key);
        }
    }

    // -------------------------------------------------------------- core merge

    /// <summary>The result of merging one data map into a document copy.</summary>
    private readonly record struct MergeResult(
        byte[] Bytes, int Replaced, SortedSet<string> Unresolved, SortedSet<string> Resolved);

    /// <summary>
    /// Merges one data map into an in-memory copy of <paramref name="originalBytes"/>:
    /// {{key}} markers, MERGEFIELD complex fields and «IF» conditional fields are all
    /// resolved from the same map. Returns the merged bytes and the resolved /
    /// unresolved key sets. The source bytes are never mutated.
    /// </summary>
    private static MergeResult MergeDocument(byte[] originalBytes, string file, IReadOnlyDictionary<string, string> data)
    {
        var ms = new MemoryStream();
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

            // The same data map fills MERGEFIELD complex fields by name alongside the
            // {{key}} markers, then resolves «IF» fields; unresolved field names join
            // the same report.
            replaced += MergeMergeFields(doc, data, resolvedKeys, unresolvedKeys);
            replaced += MergeIfFields(doc, data, resolvedKeys, unresolvedKeys);
        }

        return new MergeResult(ms.ToArray(), replaced, unresolvedKeys, resolvedKeys);
    }

    // -------------------------------------------------------------- data plumbing

    /// <summary>Creates the parent directory of an output target when it does not exist.</summary>
    private static void EnsureDirectory(string target)
    {
        var dir = Path.GetDirectoryName(target);
        if (dir is { Length: > 0 })
        {
            Directory.CreateDirectory(dir);
        }
    }

    /// <summary>
    /// Expands an --output pattern for one record: <c>{n}</c> is the 1-based record
    /// index, any other <c>{Field}</c> is that record's value (empty when absent).
    /// </summary>
    private static string ExpandOutputPattern(string pattern, IReadOnlyDictionary<string, string> record, int index) =>
        OutputPatternRegex().Replace(pattern, match =>
        {
            var key = match.Groups[1].Value;
            if (key.Equals("n", StringComparison.Ordinal))
            {
                return index.ToString(CultureInfo.InvariantCulture);
            }

            return record.TryGetValue(key, out var value) ? value : string.Empty;
        });

    /// <summary>The --data node parsed: a JsonObject (single fill) or JsonArray (mail merge).</summary>
    private static JsonNode TemplateDataNode(JsonObject args)
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
                    "Pass a JSON object like '{\"name\":\"Ada\"}', an array of records, or @data.json.",
                    innerException: ex);
            }
        }

        if (node is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "template requires --data with a JSON object or an array of record objects.",
                "Pass key/value pairs matching the {{placeholders}}, e.g. --data '{\"name\":\"Ada\"}', " +
                "or an array for a mail merge.");
        }

        return node;
    }

    /// <summary>Coerces a single-object --data node into a string map, rejecting empty/non-object input.</summary>
    private static Dictionary<string, string> TemplateData(JsonNode node)
    {
        if (node is not JsonObject obj || obj.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "template requires --data with a non-empty JSON object (or an array of records for a mail merge).",
                "Pass key/value pairs matching the {{placeholders}}, e.g. --data '{\"name\":\"Ada\"}'.");
        }

        return obj.ToDictionary(kv => kv.Key, kv => NodeToString(kv.Value), StringComparer.Ordinal);
    }

    /// <summary>Coerces one mail-merge record into a string map; each record must be an object.</summary>
    private static Dictionary<string, string> RecordData(JsonNode? record, int index)
    {
        if (record is not JsonObject obj)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"template --data record [{index}] is not a JSON object.",
                "Every element of the records array must be an object of field/value pairs.");
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
