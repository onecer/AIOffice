using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    /// <summary>Field kinds and their OOXML instruction heads.</summary>
    private static readonly Dictionary<string, string> FieldInstructions = new(StringComparer.Ordinal)
    {
        ["pageNumber"] = "PAGE",
        ["numPages"] = "NUMPAGES",
        ["date"] = "DATE",
        ["docTitle"] = "TITLE",
    };

    /// <summary>
    /// <c>{"op":"add","path":"/footer[1]/p[1]","type":"field","props":{"kind":"pageNumber"}}</c>:
    /// appends a w:fldSimple (PAGE | NUMPAGES | DATE | TITLE) with sensible cached
    /// result text to the target paragraph, in the body or any header/footer.
    /// props.format adds a date picture (<c>\@</c>) for date fields or a general
    /// switch (<c>\*</c>, e.g. roman) for the others. props.leadingText inserts a
    /// plain run before the field, which is how composites are built — the
    /// "Page X of Y" pattern is: set the paragraph text to "Page ", add a
    /// pageNumber field, then add a numPages field with leadingText " of ".
    /// </summary>
    private static object ApplyAddField(WordprocessingDocument doc, EditOp op)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        props.Remove("author");

        var kind = props["kind"] is { } kindNode ? NodeToString(kindNode) : null;
        if (kind is null || !FieldInstructions.TryGetValue(kind, out var instructionHead))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                kind is null
                    ? "add --type field needs props.kind."
                    : $"Unknown field kind '{kind}'.",
                "Use kind pageNumber, numPages, date or docTitle, e.g. " +
                "{\"op\":\"add\",\"path\":\"/footer[1]/p[1]\",\"type\":\"field\",\"props\":{\"kind\":\"pageNumber\"}}.",
                candidates: [.. FieldInstructions.Keys]);
        }

        var format = props["format"] is { } formatNode ? NodeToString(formatNode) : null;
        if (format is not null && format.AsSpan().ContainsAny('"', '\\'))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.format must not contain quotes or backslashes, got '{format}'.",
                "Pass a plain picture like \"yyyy-MM-dd\" for date fields or \"roman\" for page numbers.");
        }

        var target = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (target.Element is not Paragraph paragraph)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Fields are appended to paragraphs, not '{target.Type}'.",
                target.Type is "tc" or "header" or "footer"
                    ? $"Target the paragraph inside it: {target.CanonicalPath}/p[1]."
                    : "Pick a paragraph path, e.g. /footer[1]/p[1].");
        }

        if (props["leadingText"] is { } leadingNode && NodeToString(leadingNode) is { Length: > 0 } leadingText)
        {
            paragraph.AppendChild(new Run(NewText(leadingText)));
        }

        var instruction = format is null
            ? $" {instructionHead} \\* MERGEFORMAT "
            : kind == "date"
                ? $" {instructionHead} \\@ \"{format}\" "
                : $" {instructionHead} \\* {format} ";

        var field = new SimpleField(new Run(NewText(CachedFieldText(doc, kind, format))))
        {
            Instruction = instruction,
        };
        paragraph.AppendChild(field);

        return new
        {
            op = "add",
            type = "field",
            path = target.CanonicalPath,
            kind,
            cached = field.InnerText,
            note = "Word refreshes the value when fields update; the cached text is what shows until then.",
        };
    }

    /// <summary>A plausible cached result so the document reads sensibly before Word updates fields.</summary>
    private static string CachedFieldText(WordprocessingDocument doc, string kind, string? format)
    {
        switch (kind)
        {
            case "date":
                try
                {
                    return DateTime.Now.ToString(format ?? "yyyy-MM-dd", CultureInfo.InvariantCulture);
                }
                catch (FormatException)
                {
                    return DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                }

            case "docTitle":
                var title = doc.PackageProperties.Title;
                return title is { Length: > 0 } ? title : "(document title)";

            default: // pageNumber, numPages: 1 is the safest placeholder
                return "1";
        }
    }

    /// <summary>The field kind behind a w:fldSimple instruction ("PAGE \* MERGEFORMAT" → pageNumber).</summary>
    internal static string FieldKindName(string? instruction)
    {
        var head = instruction?.Trim().Split(' ', 2)[0] ?? string.Empty;
        return FieldInstructions.FirstOrDefault(kv => kv.Value.Equals(head, StringComparison.OrdinalIgnoreCase)).Key
            ?? (head.Length > 0 ? head : "unknown");
    }
}
