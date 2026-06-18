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
        // (1.12) document-info & reference fields, each with a cached value computed
        // headlessly (file name, body counts, core-property author/dates, bookmark text).
        ["fileName"] = "FILENAME",
        ["numWords"] = "NUMWORDS",
        ["numChars"] = "NUMCHARS",
        ["author"] = "AUTHOR",
        ["createDate"] = "CREATEDATE",
        ["saveDate"] = "SAVEDATE",
        ["printDate"] = "PRINTDATE",
        ["ref"] = "REF",
        ["hyperlink"] = "HYPERLINK",
        ["fillIn"] = "FILLIN",
    };

    /// <summary>The (1.12) date-property field kinds whose instruction takes a <c>\@</c> date picture.</summary>
    private static readonly HashSet<string> DateFieldKinds = new(StringComparer.Ordinal)
    {
        "date", "createDate", "saveDate", "printDate",
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
    private static object ApplyAddField(WordprocessingDocument doc, string file, EditOp op)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];

        var kind = props["kind"] is { } kindNode ? NodeToString(kindNode) : null;
        // props.author is the batch-author noise (the edit's revision author) that the
        // generic prop path strips. The AUTHOR field reads the document's own author
        // (core property) at cache time and Word recomputes it on open, so an inline
        // override is not supported — drop it for every kind.
        props.Remove("author");

        if (kind is null || !FieldInstructions.TryGetValue(kind, out var instructionHead))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                kind is null
                    ? "add --type field needs props.kind."
                    : $"Unknown field kind '{kind}'.",
                "Use kind pageNumber, numPages, date, docTitle, styleRef (props.styleRef), " +
                "symbol (props.charCode, props.symbolFont?), quote (props.quoteText), " +
                "fileName (props.includePath?), numWords, numChars, author, " +
                "createDate/saveDate/printDate (props.format?), ref (props.bookmark, props.mode?), " +
                "hyperlink (props.url, props.linkText?) or fillIn (props.prompt, props.default?), e.g. " +
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
            "fileName" => BuildFileNameField(file, props),
            "ref" => BuildRefField(doc, props),
            "hyperlink" => BuildHyperlinkField(props),
            "fillIn" => BuildFillInField(props),
            _ => (
                format is null
                    ? $" {instructionHead} \\* MERGEFORMAT "
                    : DateFieldKinds.Contains(kind)
                        ? $" {instructionHead} \\@ \"{format}\" "
                        : $" {instructionHead} \\* {format} ",
                CachedFieldText(doc, file, kind, format)),
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
    private static string CachedFieldText(WordprocessingDocument doc, string file, string kind, string? format)
    {
        switch (kind)
        {
            case "date":
                return FormatFieldDate(DateTime.Now, format);

            // CREATEDATE/SAVEDATE come from the core properties (created / modified);
            // PRINTDATE has no core analogue, so cache "now" as the best headless value.
            case "createDate":
                return FormatFieldDate(ReadCoreProps(doc).Created ?? DateTime.Now, format);

            case "saveDate":
                return FormatFieldDate(ReadCoreProps(doc).Modified ?? DateTime.Now, format);

            case "printDate":
                return FormatFieldDate(DateTime.Now, format);

            case "docTitle":
                var title = ReadCoreTitle(doc);
                return title is { Length: > 0 } ? title : "(document title)";

            case "author":
                var creator = ReadCoreProps(doc).Creator;
                return creator is { Length: > 0 } ? creator : "(author)";

            case "numWords":
                return BodyWordCount(doc).ToString(CultureInfo.InvariantCulture);

            case "numChars":
                return BodyCharacterCount(doc).ToString(CultureInfo.InvariantCulture);

            default: // pageNumber, numPages: 1 is the safest placeholder
                return "1";
        }
    }

    /// <summary>Formats a field's cached date with props.format (a .NET picture), falling back to ISO on a bad picture.</summary>
    private static string FormatFieldDate(DateTime when, string? format)
    {
        try
        {
            return when.ToString(format ?? "yyyy-MM-dd", CultureInfo.InvariantCulture);
        }
        catch (FormatException)
        {
            return when.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }
    }

    /// <summary>NUMWORDS cached value: body word count, matching the read --view stats counter.</summary>
    private static int BodyWordCount(WordprocessingDocument doc) =>
        BodyText(doc).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;

    /// <summary>NUMCHARS cached value: body character count (newlines excluded), matching read --view stats.</summary>
    private static int BodyCharacterCount(WordprocessingDocument doc) =>
        BodyText(doc).Replace("\n", string.Empty, StringComparison.Ordinal).Length;

    /// <summary>The body's paragraph text joined by newlines (the basis for word/char counts).</summary>
    private static string BodyText(WordprocessingDocument doc) =>
        doc.MainDocumentPart?.Document?.Body is { } body
            ? string.Join('\n', WordAddress.EnumerateBody(body).Where(n => n.Type == "p").Select(n => n.Element.InnerText))
            : string.Empty;

    /// <summary>FILENAME [\p]: the document's file name, optionally the full path (props.includePath).</summary>
    private static (string Instruction, string Cached) BuildFileNameField(string file, System.Text.Json.Nodes.JsonObject props)
    {
        var includePath = props["includePath"] is { } pathNode
            && NodeToString(pathNode) is "true" or "1" or "yes" or "on";
        var cached = includePath ? Path.GetFullPath(file) : Path.GetFileName(file);
        var pathSwitch = includePath ? " \\p" : string.Empty;
        return ($" FILENAME{pathSwitch} \\* MERGEFORMAT ", cached);
    }

    /// <summary>REF bookmark [\* switch]: a reference to a bookmark; cached = the bookmark's text (mode text, the default).</summary>
    private static (string Instruction, string Cached) BuildRefField(WordprocessingDocument doc, System.Text.Json.Nodes.JsonObject props)
    {
        var bookmark = props["bookmark"] is { } bmNode ? NodeToString(bmNode)
            : props["name"] is { } altNode ? NodeToString(altNode)
            : null;
        if (string.IsNullOrWhiteSpace(bookmark))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type field kind ref needs props.bookmark (the bookmark name to reference).",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[1]\",\"type\":\"field\"," +
                "\"props\":{\"kind\":\"ref\",\"bookmark\":\"Results\"}}.");
        }

        if (!BookmarkNamePattern().IsMatch(bookmark))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{bookmark}' is not a valid bookmark name.",
                "Bookmark names start with a letter or underscore, use letters/digits/underscores only, max 40 chars (no spaces).");
        }

        var mode = props["mode"] is { } modeNode ? NodeToString(modeNode).Trim().ToLowerInvariant() : "text";
        var modeSwitch = mode switch
        {
            "page" => " \\p",                  // the page number of the bookmark
            "abovebelow" or "above-below" => " \\p \\h", // "above"/"below" relative position
            _ => string.Empty,                  // text (default): the bookmark's text
        };

        // Cached = the bookmarked range's text for mode text (Word recomputes page refs on open).
        var cached = mode is "page" or "abovebelow" or "above-below"
            ? "1"
            : BookmarkText(doc, bookmark) ?? $"(text of {bookmark})";

        return ($" REF {bookmark}{modeSwitch} \\h ", cached);
    }

    /// <summary>The text spanning a bookmark's start→end markers, or null when the bookmark is missing/empty.</summary>
    private static string? BookmarkText(WordprocessingDocument doc, string name)
    {
        var start = EnumerateBookmarks(doc)
            .FirstOrDefault(b => string.Equals(b.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
        if (start is null)
        {
            return null;
        }

        // The bookmark commonly wraps a paragraph (our add type:bookmark form); use that
        // paragraph's text, which is what a mode-text REF resolves to.
        var paragraph = start.Ancestors<Paragraph>().FirstOrDefault();
        var text = paragraph?.InnerText;
        return string.IsNullOrEmpty(text) ? null : text;
    }

    /// <summary>HYPERLINK "url" [\o tip]: a field hyperlink; cached = the link text (or the url).</summary>
    private static (string Instruction, string Cached) BuildHyperlinkField(System.Text.Json.Nodes.JsonObject props)
    {
        var url = props["url"] is { } urlNode ? NodeToString(urlNode)
            : props["href"] is { } altNode ? NodeToString(altNode)
            : null;
        if (string.IsNullOrWhiteSpace(url))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type field kind hyperlink needs props.url (the link target).",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[1]\",\"type\":\"field\"," +
                "\"props\":{\"kind\":\"hyperlink\",\"url\":\"https://example.com\",\"linkText\":\"site\"}}.");
        }

        if (url.AsSpan().ContainsAny('"', '\\'))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.url must not contain quotes or backslashes, got '{url}'.",
                "Pass a plain URL like \"https://example.com\".");
        }

        var linkText = props["linkText"] is { } textNode ? NodeToString(textNode)
            : props["text"] is { } altTextNode ? NodeToString(altTextNode)
            : null;
        if (linkText is not null && linkText.Contains('"', StringComparison.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.linkText must not contain quotes, got '{linkText}'.",
                "Pass plain display text without double quotes.");
        }

        var cached = linkText is { Length: > 0 } ? linkText : url;
        return ($" HYPERLINK \"{url}\" ", cached);
    }

    /// <summary>FILLIN "prompt" [\d default]: an interactive prompt; cached = the default/placeholder text.</summary>
    private static (string Instruction, string Cached) BuildFillInField(System.Text.Json.Nodes.JsonObject props)
    {
        var prompt = props["prompt"] is { } promptNode ? NodeToString(promptNode) : null;
        if (string.IsNullOrWhiteSpace(prompt))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type field kind fillIn needs props.prompt (the prompt Word shows the user).",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[1]\",\"type\":\"field\"," +
                "\"props\":{\"kind\":\"fillIn\",\"prompt\":\"Enter your name\",\"default\":\"Name\"}}.");
        }

        if (prompt.Contains('"', StringComparison.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.prompt must not contain quotes, got '{prompt}'.",
                "Pass plain prompt text without double quotes.");
        }

        var def = props["default"] is { } defNode ? NodeToString(defNode)
            : props["defaultText"] is { } altNode ? NodeToString(altNode)
            : null;
        if (def is not null && def.Contains('"', StringComparison.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.default must not contain quotes, got '{def}'.",
                "Pass plain default text without double quotes.");
        }

        var defaultSwitch = def is { Length: > 0 } ? $" \\d \"{def}\"" : string.Empty;
        return ($" FILLIN \"{prompt}\"{defaultSwitch} ", def ?? string.Empty);
    }

    /// <summary>The field kind behind a w:fldSimple instruction ("PAGE \* MERGEFORMAT" → pageNumber).</summary>
    internal static string FieldKindName(string? instruction)
    {
        var head = instruction?.Trim().Split(' ', 2)[0] ?? string.Empty;
        return FieldInstructions.FirstOrDefault(kv => kv.Value.Equals(head, StringComparison.OrdinalIgnoreCase)).Key
            ?? (head.Length > 0 ? head : "unknown");
    }
}
