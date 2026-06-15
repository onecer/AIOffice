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
        // (1.7) reference/insertion fields that take their own props (styleRef, charCode/symbolFont, quoteText).
        ["styleRef"] = "STYLEREF",
        ["symbol"] = "SYMBOL",
        ["quote"] = "QUOTE",
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
                "Use kind pageNumber, numPages, date, docTitle, styleRef (props.styleRef), " +
                "symbol (props.charCode, props.symbolFont?) or quote (props.quoteText), e.g. " +
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

        var (instruction, cached) = kind switch
        {
            "styleRef" => BuildStyleRefField(props),
            "symbol" => BuildSymbolField(props),
            "quote" => BuildQuoteField(props),
            _ => (
                format is null
                    ? $" {instructionHead} \\* MERGEFORMAT "
                    : kind == "date"
                        ? $" {instructionHead} \\@ \"{format}\" "
                        : $" {instructionHead} \\* {format} ",
                CachedFieldText(doc, kind, format)),
        };

        // SYMBOL renders its glyph with no separate cached-result run (the glyph IS
        // the field's display); the others carry a cached result run Word repaints.
        var field = kind == "symbol"
            ? new SimpleField { Instruction = instruction }
            : new SimpleField(new Run(NewText(cached))) { Instruction = instruction };
        paragraph.AppendChild(field);

        return new
        {
            op = "add",
            type = "field",
            path = target.CanonicalPath,
            kind,
            cached = kind == "symbol" ? cached : field.InnerText,
            note = "Word refreshes the value when fields update; the cached text is what shows until then.",
        };
    }

    /// <summary>STYLEREF "Style Name": the text of the nearest paragraph in that style (running headers, etc.).</summary>
    private static (string Instruction, string Cached) BuildStyleRefField(System.Text.Json.Nodes.JsonObject props)
    {
        var style = props["styleRef"] is { } styleNode ? NodeToString(styleNode)
            : props["style"] is { } altNode ? NodeToString(altNode)
            : null;
        if (string.IsNullOrWhiteSpace(style))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type field kind styleRef needs props.styleRef (the style name to reference).",
                "Example: {\"op\":\"add\",\"path\":\"/header[1]/p[1]\",\"type\":\"field\"," +
                "\"props\":{\"kind\":\"styleRef\",\"styleRef\":\"Heading 1\"}}.");
        }

        if (style.Contains('"', StringComparison.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.styleRef must not contain quotes, got '{style}'.",
                "Pass a plain style name like \"Heading 1\".");
        }

        // The cached result is a placeholder; Word fills in the referenced style's text on refresh.
        return ($" STYLEREF \"{style}\" \\* MERGEFORMAT ", $"(text of {style})");
    }

    /// <summary>SYMBOL charCode \f "Font": one symbol glyph by character code (decimal or 0x-hex) and optional font.</summary>
    private static (string Instruction, string Cached) BuildSymbolField(System.Text.Json.Nodes.JsonObject props)
    {
        var codeText = props["charCode"] is { } codeNode ? NodeToString(codeNode)
            : props["code"] is { } altNode ? NodeToString(altNode)
            : null;
        if (string.IsNullOrWhiteSpace(codeText))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type field kind symbol needs props.charCode (a decimal or 0x-hex character code).",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[1]\",\"type\":\"field\"," +
                "\"props\":{\"kind\":\"symbol\",\"charCode\":169}} for ©, optionally with symbolFont.");
        }

        var trimmed = codeText.Trim();
        int code;
        if (trimmed.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(trimmed.AsSpan(2), System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out var hexValue))
        {
            code = hexValue;
        }
        else if (int.TryParse(trimmed, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var decValue))
        {
            code = decValue;
        }
        else
        {
            code = -1;
        }
        if (code is < 0 or > 0x10FFFF)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.charCode '{codeText}' is not a valid character code.",
                "Pass a decimal (169) or 0x-hex (0xA9) code in the Unicode range.");
        }

        var font = props["symbolFont"] is { } fontNode ? NodeToString(fontNode)
            : props["font"] is { } altFontNode ? NodeToString(altFontNode)
            : null;
        if (font is not null && font.Contains('"', StringComparison.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.symbolFont must not contain quotes, got '{font}'.",
                "Pass a plain font name like \"Symbol\" or \"Wingdings\".");
        }

        // Word's SYMBOL field takes a decimal code; \f names the font, \u uses Unicode.
        var fontSwitch = font is { Length: > 0 } ? $" \\f \"{font}\"" : string.Empty;
        var instruction = $" SYMBOL {code.ToString(System.Globalization.CultureInfo.InvariantCulture)}{fontSwitch} \\u ";

        // A best-effort cached glyph: render the code point directly when it is one
        // (font-mapped symbol fonts still show the right glyph in Word on refresh).
        var cached = code <= 0x10FFFF && !char.IsSurrogate((char)Math.Min(code, 0xFFFF))
            ? char.ConvertFromUtf32(code)
            : "□";
        return (instruction, cached);
    }

    /// <summary>QUOTE "literal text": inserts a literal text field whose result is the quoted string.</summary>
    private static (string Instruction, string Cached) BuildQuoteField(System.Text.Json.Nodes.JsonObject props)
    {
        var quote = props["quoteText"] is { } quoteNode ? NodeToString(quoteNode)
            : props["text"] is { } textNode ? NodeToString(textNode)
            : null;
        if (quote is null || quote.Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type field kind quote needs props.quoteText (the literal text to quote).",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[1]\",\"type\":\"field\"," +
                "\"props\":{\"kind\":\"quote\",\"quoteText\":\"Confidential\"}}.");
        }

        if (quote.Contains('"', StringComparison.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.quoteText must not contain quotes, got '{quote}'.",
                "Pass plain text without double quotes.");
        }

        return ($" QUOTE \"{quote}\" ", quote);
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
                var title = ReadCoreTitle(doc);
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
