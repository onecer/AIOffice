using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// Mail-merge fields (v1.2.0). A MERGEFIELD is a Word merge placeholder distinct
/// from our <c>{{template}}</c> markers: it is a complex field
/// (<c>MERGEFIELD FirstName</c>) whose cached result reads <c>«FirstName»</c>
/// until data is merged in. The <c>template</c> verb now fills MERGEFIELDs by
/// name from the same <c>--data</c> map that resolves <c>{{key}}</c> markers, and
/// <c>read --view fields</c> lists every merge field with its current value.
/// </summary>
public sealed partial class WordHandler
{
    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[2]","type":"mergeField","props":{"name":"FirstName"}}</c>:
    /// appends (or, with props.find, places at matched text) a MERGEFIELD complex
    /// field on the target paragraph. The cached result is Word's «name» chevron
    /// form so the document reads sensibly before a merge fills it.
    /// </summary>
    private static object ApplyAddMergeField(WordprocessingDocument doc, EditOp op)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        props.Remove("author");

        var name = props["name"] is { } nameNode ? NodeToString(nameNode).Trim() : null;
        if (string.IsNullOrEmpty(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type mergeField needs props.name (the merge field's name).",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[2]\",\"type\":\"mergeField\",\"props\":{\"name\":\"FirstName\"}}.");
        }

        // A MERGEFIELD name is a single Word field-name token: letters, digits,
        // underscores. Quotes/backslashes/spaces would corrupt the instruction.
        if (name.AsSpan().ContainsAnyExcept(MergeFieldNameChars))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.name '{name}' is not a valid merge field name.",
                "Use letters, digits and underscores only (e.g. \"FirstName\" or \"Order_Id\").");
        }

        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (anchor.Element is not Paragraph paragraph)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Merge fields are placed on paragraphs, not '{anchor.Type}'.",
                anchor.Type is "tc" or "header" or "footer"
                    ? $"Target the paragraph inside it: {anchor.CanonicalPath}/p[1]."
                    : "Pick a paragraph path, e.g. /body/p[2].");
        }

        var find = props["find"] is { } findNode ? NodeToString(findNode) : null;
        var cached = MergeFieldChevron(name);
        var instruction = $" MERGEFIELD {name} ";

        if (find is { Length: > 0 })
        {
            PlaceComplexFieldAtText(paragraph, find, instruction, cached);
        }
        else
        {
            AppendComplexField(paragraph, instruction, cached);
        }

        return new
        {
            op = "add",
            type = "mergeField",
            path = anchor.CanonicalPath,
            name,
            cached,
            note = "Fill it with 'aioffice template <file> --data {\"" + name + "\":\"…\"}'; Word merges it from a data source too.",
        };
    }

    /// <summary>Word's display form for an unmerged merge field: «name».</summary>
    private static string MergeFieldChevron(string name) => "«" + name + "»";

    private static readonly System.Buffers.SearchValues<char> MergeFieldNameChars =
        System.Buffers.SearchValues.Create(
            "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_");

    /// <summary>
    /// Inserts a complex field immediately before the first occurrence of
    /// <paramref name="find"/> in the paragraph's text, splitting a run at the
    /// match boundary so the field lands exactly there. If the text is not found,
    /// the field is appended (so it is never lost).
    /// </summary>
    private static void PlaceComplexFieldAtText(Paragraph paragraph, string find, string instruction, string cached) =>
        PlaceFieldAtText(paragraph, find, () => ComplexFieldRuns(instruction, cached), () => AppendComplexField(paragraph, instruction, cached));

    /// <summary>The five runs of a complex field (begin / instr / separate / cached / end), unattached.</summary>
    private static List<Run> ComplexFieldRuns(string instruction, string cached) =>
    [
        new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
        new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
        new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
        new Run(NewText(cached)),
        new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
    ];

    /// <summary>
    /// Shared field placement at matched text: finds the first occurrence of
    /// <paramref name="find"/> in the paragraph's concatenated text, splits the
    /// runs at that boundary (reusing the replace machinery), and inserts the
    /// field runs before the run that now starts with the match. Falls back to
    /// <paramref name="append"/> when the text is absent.
    /// </summary>
    private static void PlaceFieldAtText(Paragraph paragraph, string find, Func<List<Run>> buildRuns, Action append)
    {
        var full = string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));
        var offset = full.IndexOf(find, StringComparison.Ordinal);
        if (offset < 0)
        {
            append();
            return;
        }

        EnsureRunBoundary(paragraph, offset);

        // The run that begins exactly at the match offset is our insertion point.
        var running = 0;
        foreach (var run in paragraph.Elements<Run>())
        {
            if (running == offset)
            {
                var anchor = (OpenXmlElement)run;
                foreach (var fieldRun in buildRuns())
                {
                    anchor.InsertBeforeSelf(fieldRun);
                }

                return;
            }

            running += run.InnerText.Length;
        }

        append();
    }

    // ------------------------------------------------------------------- read

    /// <summary>A merge field resolved from the document: its name, current cached value and the run holding the instruction.</summary>
    internal sealed record MergeFieldInfo(string Name, string Value, FieldCode Code);

    /// <summary>Every MERGEFIELD complex field across body, headers and footers, in document order.</summary>
    private static List<MergeFieldInfo> EnumerateMergeFields(WordprocessingDocument doc)
    {
        var fields = new List<MergeFieldInfo>();
        foreach (var root in MergeFieldRoots(doc))
        {
            foreach (var code in root.Descendants<FieldCode>())
            {
                if (MergeFieldName(code.Text) is not { } name)
                {
                    continue;
                }

                fields.Add(new MergeFieldInfo(name, ComplexFieldResult(code), code));
            }
        }

        return fields;
    }

    private static IEnumerable<OpenXmlElement> MergeFieldRoots(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart?.Document?.Body is { } body)
        {
            yield return body;
        }

        foreach (var header in doc.MainDocumentPart?.HeaderParts ?? [])
        {
            if (header.Header is { } h)
            {
                yield return h;
            }
        }

        foreach (var footer in doc.MainDocumentPart?.FooterParts ?? [])
        {
            if (footer.Footer is { } f)
            {
                yield return f;
            }
        }
    }

    /// <summary>The field name of a MERGEFIELD instruction ("MERGEFIELD FirstName \* MERGEFORMAT" → "FirstName"), or null.</summary>
    internal static string? MergeFieldName(string? instruction)
    {
        if (instruction is null)
        {
            return null;
        }

        var tokens = instruction.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return tokens.Length >= 2 && tokens[0].Equals("MERGEFIELD", StringComparison.OrdinalIgnoreCase)
            ? tokens[1]
            : null;
    }

    /// <summary>read --view fields shape for a merge field: {path, kind:"mergeField", name, value}.</summary>
    private static List<object> MergeFieldsView(WordprocessingDocument doc)
    {
        var seen = new Dictionary<string, int>(StringComparer.Ordinal);
        var result = new List<object>();
        foreach (var field in EnumerateMergeFields(doc))
        {
            seen.TryGetValue(field.Name, out var ordinal);
            ordinal++;
            seen[field.Name] = ordinal;
            result.Add(new
            {
                path = MergeFieldPath(field.Name, ordinal),
                kind = "mergeField",
                name = field.Name,
                value = field.Value,
            });
        }

        return result;
    }

    /// <summary>get on a merge field path /mergeField[@name=X] (or /mergeField[@name=X][i]) → {name, value}.</summary>
    private static (string Path, Dictionary<string, object?> Properties) GetMergeFieldProperties(
        WordprocessingDocument doc, string pathArg)
    {
        var (name, index) = ParseMergeFieldPath(pathArg);
        var matches = EnumerateMergeFields(doc).Where(f => f.Name == name).ToList();
        if (matches.Count == 0 || index > matches.Count)
        {
            var available = EnumerateMergeFields(doc).Select(f => f.Name).Distinct(StringComparer.Ordinal).ToList();
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                matches.Count == 0
                    ? $"No merge field named '{name}' exists."
                    : $"/mergeField[@name={name}][{index}] does not exist; there are {matches.Count} with that name.",
                "Run 'aioffice read <file> --view fields' to list merge fields with their names and values.",
                candidates: [.. available.Take(5).Select(n => MergeFieldPath(n, 1))]);
        }

        var field = matches[index - 1];
        return (MergeFieldPath(name, index), new Dictionary<string, object?>
        {
            ["name"] = field.Name,
            ["value"] = field.Value,
        });
    }

    private static string MergeFieldPath(string name, int ordinal) =>
        string.Create(CultureInfo.InvariantCulture, $"/mergeField[@name={name}][{ordinal}]");

    private static (string Name, int Index) ParseMergeFieldPath(string pathArg)
    {
        var match = System.Text.RegularExpressions.Regex.Match(
            pathArg.Trim(), @"^/?mergeField\[@name=([A-Za-z0-9_]+)\](?:\[([0-9]+)\])?$");
        if (!match.Success)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"'{pathArg}' is not a merge field path.",
                "Address a merge field as /mergeField[@name=FirstName] (optionally [@name=FirstName][2] for repeats).");
        }

        var index = match.Groups[2].Success
            ? int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture)
            : 1;
        return (match.Groups[1].Value, index);
    }

    /// <summary>
    /// Fills MERGEFIELD complex fields by name from the template data, mirroring
    /// the <c>{{key}}</c> merge: the cached result run between separate and end is
    /// replaced with the data value. Unresolved field names are recorded so the
    /// template verb can report them; resolved ones count toward the replacement
    /// total. Returns the number of merge fields filled.
    /// </summary>
    private static int MergeMergeFields(
        WordprocessingDocument doc,
        IReadOnlyDictionary<string, string> data,
        ISet<string> resolvedKeys,
        ISet<string> unresolvedKeys)
    {
        var replaced = 0;
        foreach (var field in EnumerateMergeFields(doc))
        {
            if (!data.TryGetValue(field.Name, out var value))
            {
                unresolvedKeys.Add(field.Name);
                continue;
            }

            if (SetComplexFieldResult(field.Code, value))
            {
                resolvedKeys.Add(field.Name);
                replaced++;
            }
        }

        return replaced;
    }

    /// <summary>
    /// Replaces a complex field's cached result (the runs between separate and
    /// end) with one text run carrying <paramref name="value"/>. Returns false if
    /// the field has no separate/end frame (it is left untouched).
    /// </summary>
    private static bool SetComplexFieldResult(FieldCode code, string value)
    {
        var startRun = code.Ancestors<Run>().FirstOrDefault() ?? code.Parent as Run;
        if (startRun?.Parent is not { } parent)
        {
            return false;
        }

        var siblings = parent.ChildElements.ToList();
        var startIndex = siblings.IndexOf(startRun);
        var separateIndex = -1;
        var endIndex = -1;
        for (var i = startIndex + 1; i < siblings.Count; i++)
        {
            if (siblings[i] is not Run run)
            {
                continue;
            }

            var type = run.GetFirstChild<FieldChar>()?.FieldCharType?.Value;
            if (type == FieldCharValues.Separate && separateIndex < 0)
            {
                separateIndex = i;
            }
            else if (type == FieldCharValues.End)
            {
                endIndex = i;
                break;
            }
        }

        if (separateIndex < 0 || endIndex < 0)
        {
            return false;
        }

        // Drop the cached result runs strictly between separate and end, then
        // insert one run with the merged value right after the separate marker.
        for (var i = endIndex - 1; i > separateIndex; i--)
        {
            siblings[i].Remove();
        }

        siblings[separateIndex].InsertAfterSelf(new Run(NewText(value)));
        return true;
    }
}
