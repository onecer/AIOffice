using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>The list-ness of one paragraph: bullet|number, 0-based level, its w:numId.</summary>
internal sealed record ListInfo(string Kind, int Level, int NumId);

/// <summary>Parsed list props off an add op: kind, level and the restart flag.</summary>
internal sealed record ListRequest(string Kind, int Level, bool Restart);

public sealed partial class WordHandler
{
    private const int MaxListLevel = 8;

    // ------------------------------------------------------------ add props

    /// <summary>
    /// Pops props.list / props.level / props.listRestart from an add-paragraph
    /// op so generic prop application never sees them. Returns null when the
    /// paragraph is not a list item.
    /// </summary>
    private static ListRequest? PopListProps(JsonObject props)
    {
        string? kind = null;
        if (props.TryGetPropertyValue("list", out var listNode))
        {
            props.Remove("list");
            kind = NodeToString(listNode).ToLowerInvariant();
            if (kind is not ("bullet" or "number"))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown list kind '{kind}'.",
                    "Use list \"bullet\" or \"number\", e.g. {\"op\":\"add\",\"path\":\"/body\",\"type\":\"p\",\"props\":{\"text\":\"item\",\"list\":\"bullet\"}}.",
                    candidates: ["bullet", "number"]);
            }
        }

        var level = 0;
        if (props.TryGetPropertyValue("level", out var levelNode))
        {
            props.Remove("level");
            if (kind is null)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "props.level only applies to list paragraphs.",
                    "Add props.list \"bullet\" or \"number\" alongside level, or drop level.");
            }

            if (!int.TryParse(NodeToString(levelNode), NumberStyles.Integer, CultureInfo.InvariantCulture, out level)
                || level is < 0 or > MaxListLevel)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"props.level must be 0..{MaxListLevel}, got '{NodeToString(levelNode)}'.",
                    "Level 0 is the outermost list level; each deeper level indents once more.");
            }
        }

        var restart = false;
        if (props.TryGetPropertyValue("listRestart", out var restartNode))
        {
            props.Remove("listRestart");
            if (kind is null)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "props.listRestart only applies to list paragraphs.",
                    "Add props.list \"number\" alongside listRestart, or drop it.");
            }

            restart = NodeToString(restartNode) is "true" or "1";
        }

        return kind is null ? null : new ListRequest(kind, level, restart);
    }

    /// <summary>
    /// Wires a freshly inserted paragraph into list numbering. Consecutive
    /// same-kind list paragraphs share a w:numId (the sequence continues);
    /// a non-list neighbor, a different kind, or listRestart starts a fresh
    /// numbering instance (numbers restart at 1 via startOverride).
    /// </summary>
    private static void ApplyListNumbering(WordprocessingDocument doc, Paragraph paragraph, ListRequest request)
    {
        int? numId = null;
        if (!request.Restart)
        {
            numId = NeighborListNumId(doc, paragraph.PreviousSibling<Paragraph>(), request.Kind)
                ?? NeighborListNumId(doc, paragraph.NextSibling<Paragraph>(), request.Kind);
        }

        numId ??= NewNumberingInstance(doc, request.Kind);

        var pPr = paragraph.ParagraphProperties ??= new ParagraphProperties();
        pPr.NumberingProperties = new NumberingProperties(
            new NumberingLevelReference { Val = request.Level },
            new NumberingId { Val = numId.Value });
    }

    /// <summary>The neighbor's numId when it is a list paragraph of the same kind.</summary>
    private static int? NeighborListNumId(WordprocessingDocument doc, Paragraph? neighbor, string kind) =>
        neighbor is not null && ListInfoOf(doc, neighbor) is { } info && info.Kind == kind
            ? info.NumId
            : null;

    // -------------------------------------------------------- numbering part

    private static Numbering EnsureNumberingRoot(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The document has no main part.",
            "Re-export the file from Word.");

        var part = main.NumberingDefinitionsPart;
        if (part is null)
        {
            part = main.AddNewPart<NumberingDefinitionsPart>();
            part.Numbering = new Numbering();
        }

        return part.Numbering ??= new Numbering();
    }

    /// <summary>One shared abstract definition per list kind, created on demand.</summary>
    private static int GetOrCreateAbstractNum(WordprocessingDocument doc, string kind)
    {
        var numbering = EnsureNumberingRoot(doc);
        var wantBullet = kind == "bullet";

        foreach (var abstractNum in numbering.Elements<AbstractNum>())
        {
            var format = abstractNum.GetFirstChild<Level>()?.NumberingFormat?.Val?.Value;
            if (format is not null && (format == NumberFormatValues.Bullet) == wantBullet)
            {
                return abstractNum.AbstractNumberId?.Value ?? 0;
            }
        }

        var id = numbering.Elements<AbstractNum>()
            .Select(a => a.AbstractNumberId?.Value ?? 0)
            .DefaultIfEmpty(-1)
            .Max() + 1;

        var created = new AbstractNum { AbstractNumberId = id };
        created.AppendChild(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });
        for (var level = 0; level <= MaxListLevel; level++)
        {
            created.AppendChild(new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = wantBullet ? NumberFormatValues.Bullet : NumberFormatValues.Decimal },
                new LevelText { Val = wantBullet ? "•" : $"%{level + 1}." },
                new LevelJustification { Val = LevelJustificationValues.Left },
                new PreviousParagraphProperties(new Indentation
                {
                    Left = ((level + 1) * 720).ToString(CultureInfo.InvariantCulture),
                    Hanging = "360",
                }))
            {
                LevelIndex = level,
            });
        }

        // Schema order inside w:numbering: abstractNum* before num*.
        var firstNum = numbering.Elements<NumberingInstance>().FirstOrDefault();
        if (firstNum is null)
        {
            numbering.AppendChild(created);
        }
        else
        {
            numbering.InsertBefore(created, firstNum);
        }

        return id;
    }

    /// <summary>
    /// A fresh w:num for one sequence. Numbered instances carry startOverride 1
    /// on every level so a new sequence restarts even though instances share
    /// the abstract definition.
    /// </summary>
    private static int NewNumberingInstance(WordprocessingDocument doc, string kind)
    {
        var abstractId = GetOrCreateAbstractNum(doc, kind);
        var numbering = EnsureNumberingRoot(doc);
        var numId = numbering.Elements<NumberingInstance>()
            .Select(n => n.NumberID?.Value ?? 0)
            .DefaultIfEmpty(0)
            .Max() + 1;

        var instance = new NumberingInstance(new AbstractNumId { Val = abstractId }) { NumberID = numId };
        if (kind == "number")
        {
            for (var level = 0; level <= MaxListLevel; level++)
            {
                instance.AppendChild(new LevelOverride(new StartOverrideNumberingValue { Val = 1 })
                {
                    LevelIndex = level,
                });
            }
        }

        numbering.AppendChild(instance);
        return numId;
    }

    // ------------------------------------------------------------------ read

    /// <summary>The list info of a paragraph, or null when it is not a (resolvable) list item.</summary>
    private static ListInfo? ListInfoOf(WordprocessingDocument doc, Paragraph paragraph)
    {
        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        if (numPr?.NumberingId?.Val?.Value is not { } numId || numId == 0)
        {
            return null;
        }

        var level = numPr.NumberingLevelReference?.Val?.Value ?? 0;
        var numbering = doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        var instance = numbering?.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == numId);
        var abstractId = instance?.GetFirstChild<AbstractNumId>()?.Val?.Value;
        var abstractNum = numbering?.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == abstractId);
        var format = abstractNum?.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == level)?.NumberingFormat?.Val
            ?? abstractNum?.GetFirstChild<Level>()?.NumberingFormat?.Val;

        if (abstractNum is null)
        {
            return null; // dangling numId: treat as not-a-list rather than guessing
        }

        var kind = format is not null && format.Value == NumberFormatValues.Bullet ? "bullet" : "number";
        return new ListInfo(kind, level, numId);
    }

    /// <summary>
    /// The 1-based ordinal of a numbered list item: items of the same instance
    /// and level count up in document order; a shallower item of the same
    /// instance resets deeper counters (standard multi-level behavior).
    /// </summary>
    private static int ComputeListNumber(WordprocessingDocument doc, Paragraph target, ListInfo info)
    {
        var counters = new Dictionary<int, int>(); // level -> count, scoped to info.NumId
        foreach (var paragraph in doc.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>())
        {
            if (ListInfoOf(doc, paragraph) is not { } current || current.NumId != info.NumId)
            {
                continue;
            }

            counters[current.Level] = counters.GetValueOrDefault(current.Level) + 1;
            foreach (var deeper in counters.Keys.Where(l => l > current.Level).ToList())
            {
                counters[deeper] = 0;
            }

            if (ReferenceEquals(paragraph, target))
            {
                return counters[current.Level];
            }
        }

        return counters.GetValueOrDefault(info.Level) + 1;
    }

    /// <summary>"• " / "3. " for the text view, indented two spaces per level; empty for non-list paragraphs.</summary>
    private static string ListMarker(WordprocessingDocument doc, Paragraph paragraph)
    {
        if (ListInfoOf(doc, paragraph) is not { } info)
        {
            return string.Empty;
        }

        var indent = new string(' ', info.Level * 2);
        return info.Kind == "bullet"
            ? indent + "• "
            : string.Create(
                CultureInfo.InvariantCulture,
                $"{indent}{ComputeListNumber(doc, paragraph, info)}. ");
    }
}
