using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    /// <summary>
    /// The (1.7) drop-cap props an agent can <c>set</c> on a body paragraph:
    /// <c>dropCap</c> ("drop" inside the text column, "margin" out in the margin),
    /// optional <c>dropCapLines</c> (how many lines tall, default 3) and optional
    /// <c>dropCapFont</c> (a font name for the dropped letter). They restructure the
    /// paragraph into Word's drop-cap shape: the first letter moves to a framed
    /// paragraph (<c>w:framePr w:dropCap</c>) and the remaining text stays put.
    /// </summary>
    private static readonly IReadOnlyList<string> DropCapProps = ["dropCap", "dropCapLines", "dropCapFont"];

    /// <summary>
    /// Detaches the drop-cap keys from <paramref name="props"/> (mutating it) so the
    /// stringly-typed formatting loop never sees them. Returns null when no drop-cap
    /// key is present. The companion remover (<c>dropCap:"none"|false</c>) is folded
    /// in so an agent can turn a drop cap off the same way it turned it on.
    /// </summary>
    private static JsonObject? ExtractDropCapProps(JsonObject props)
    {
        var requested = DropCapProps.Where(props.ContainsKey).ToList();
        if (requested.Count == 0)
        {
            return null;
        }

        var detached = new JsonObject();
        foreach (var key in requested)
        {
            var node = props[key];
            props.Remove(key);
            detached[key] = node?.DeepClone();
        }

        return detached;
    }

    /// <summary>
    /// Builds the drop cap on <paramref name="paragraph"/>: the first character is
    /// lifted into a new framed paragraph carrying <c>w:framePr</c> with the dropCap
    /// location, line span and wrap, sized so the glyph spans the requested number of
    /// lines. dropCap:"none"/false removes an existing drop cap (rejoining the letter).
    /// </summary>
    private static object ApplyDropCap(Paragraph paragraph, JsonObject props)
    {
        var location = props["dropCap"] is { } locationNode ? NodeToString(locationNode).Trim() : "drop";
        if (location is "none" or "false" or "off" or "0" or "")
        {
            return RemoveDropCap(paragraph);
        }

        var dropCap = location.ToLowerInvariant() switch
        {
            "drop" => DropCapLocationValues.Drop,
            "margin" => DropCapLocationValues.Margin,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"dropCap must be \"drop\", \"margin\" or \"none\", got '{location}'.",
                "Use dropCap:\"drop\" (inside the text), dropCap:\"margin\" (in the margin) or dropCap:\"none\" to clear it.",
                candidates: ["drop", "margin", "none"]),
        };

        var lines = props["dropCapLines"] is { } linesNode && int.TryParse(NodeToString(linesNode), out var n)
            ? Math.Clamp(n, 1, 10)
            : 3;
        var font = props["dropCapFont"] is { } fontNode && NodeToString(fontNode) is { Length: > 0 } f ? f : null;

        // If a drop cap is already present, clear it first so re-applying is idempotent.
        RemoveDropCap(paragraph);

        var firstLetter = FirstLetter(paragraph);
        if (firstLetter is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "The paragraph has no text to drop-cap.",
                "Set the paragraph text first, then apply dropCap.");
        }

        RemoveFirstLetter(paragraph);

        // The dropped glyph is roughly lines × the body size; ~2 half-points per line
        // is a sensible default Word-like cap height when the body is the default 11pt.
        var capHalfPoints = (lines * 22).ToString(CultureInfo.InvariantCulture);
        var letterRunProps = new RunProperties(
            new Position { Val = "-6" },
            new FontSize { Val = capHalfPoints },
            new FontSizeComplexScript { Val = capHalfPoints });
        if (font is not null)
        {
            letterRunProps.RunFonts = new RunFonts { Ascii = font, HighAnsi = font, ComplexScript = font };
        }

        var framePr = new FrameProperties
        {
            DropCap = dropCap,
            Lines = lines,
            Wrap = TextWrappingValues.Around,
            VerticalSpace = "0",
            HorizontalSpace = "0",
            VerticalPosition = VerticalAnchorValues.Text,
            HorizontalPosition = dropCap == DropCapLocationValues.Margin
                ? HorizontalAnchorValues.Margin
                : HorizontalAnchorValues.Text,
        };

        var framedParagraph = new Paragraph(
            new ParagraphProperties(framePr) { SpacingBetweenLines = new SpacingBetweenLines { LineRule = LineSpacingRuleValues.Exact, Line = "1" } },
            new Run(letterRunProps, NewText(firstLetter)));

        paragraph.InsertBeforeSelf(framedParagraph);

        return new { location = location.ToLowerInvariant(), lines, letter = firstLetter, font };
    }

    /// <summary>
    /// Removes a drop cap that precedes <paramref name="paragraph"/>: the framed
    /// drop-cap paragraph (if any) is folded back, its letter rejoining the head of
    /// the body paragraph. Returns a small summary either way (a no-op is reported).
    /// </summary>
    private static object RemoveDropCap(Paragraph paragraph)
    {
        if (paragraph.PreviousSibling() is Paragraph previous &&
            previous.ParagraphProperties?.FrameProperties?.DropCap is not null)
        {
            var letter = previous.InnerText;
            previous.Remove();
            if (letter.Length > 0)
            {
                PrependLetter(paragraph, letter);
            }

            return new { removed = true, letter };
        }

        return new { removed = false };
    }

    /// <summary>The first non-whitespace character of the paragraph's text (the glyph to drop), or null when empty.</summary>
    private static string? FirstLetter(Paragraph paragraph)
    {
        foreach (var run in paragraph.ChildElements.OfType<Run>())
        {
            foreach (var text in run.ChildElements.OfType<Text>())
            {
                if (text.Text is { Length: > 0 } value)
                {
                    return value[..1];
                }
            }
        }

        return null;
    }

    /// <summary>Drops the first character from the paragraph's first text run (used when lifting it into the frame).</summary>
    private static void RemoveFirstLetter(Paragraph paragraph)
    {
        foreach (var run in paragraph.ChildElements.OfType<Run>())
        {
            foreach (var text in run.ChildElements.OfType<Text>())
            {
                if (text.Text is { Length: > 0 } value)
                {
                    text.Text = value[1..];
                    text.Space = SpaceProcessingModeValues.Preserve;
                    return;
                }
            }
        }
    }

    /// <summary>Re-inserts a letter at the head of the paragraph's first text run (drop-cap removal).</summary>
    private static void PrependLetter(Paragraph paragraph, string letter)
    {
        var firstRun = paragraph.ChildElements.OfType<Run>().FirstOrDefault();
        if (firstRun is null)
        {
            paragraph.AppendChild(new Run(NewText(letter)));
            return;
        }

        var firstText = firstRun.ChildElements.OfType<Text>().FirstOrDefault();
        if (firstText is null)
        {
            firstRun.InsertAt(NewText(letter), 0);
        }
        else
        {
            firstText.Text = letter + firstText.Text;
            firstText.Space = SpaceProcessingModeValues.Preserve;
        }
    }

    /// <summary>get reporting: the drop-cap shape preceding a paragraph, or null when it has none.</summary>
    private static Dictionary<string, object?>? DropCapShape(Paragraph paragraph)
    {
        if (paragraph.PreviousSibling() is not Paragraph previous ||
            previous.ParagraphProperties?.FrameProperties is not { DropCap: { } dropCap } framePr)
        {
            return null;
        }

        return new Dictionary<string, object?>
        {
            ["location"] = dropCap.Value == DropCapLocationValues.Margin ? "margin" : "drop",
            ["lines"] = framePr.Lines?.Value ?? 1,
            ["letter"] = previous.InnerText,
            ["font"] = previous.Descendants<RunFonts>().FirstOrDefault()?.Ascii?.Value,
        };
    }
}
