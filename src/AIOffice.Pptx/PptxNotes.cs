using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Speaker notes (/slide[i]/notes): read, create, replace, append and clear.
/// A slide's notes live in a NotesSlidePart whose body placeholder holds the
/// text; the part is wired to the deck's single notes master (created on
/// demand, with its own theme, and registered in p:notesMasterIdLst).
/// </summary>
internal static class PptxNotes
{
    // ---- read side ----------------------------------------------------------

    /// <summary>The paragraph texts of a slide's notes; empty when the slide has no notes part.</summary>
    public static List<string> Paragraphs(SlidePart slidePart)
    {
        var body = BodyShape(slidePart.NotesSlidePart)?.TextBody;
        if (body is null)
        {
            return [];
        }

        return [.. body.Elements<A.Paragraph>().Select(PptxDoc.ParagraphText)];
    }

    /// <summary>Newline-joined notes text; empty when the slide has no notes.</summary>
    public static string Text(SlidePart slidePart) => string.Join('\n', Paragraphs(slidePart));

    /// <summary>The `get` projection for /slide[i]/notes.</summary>
    public static object NotesDetail(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var paragraphs = Paragraphs(slidePart);
        return new
        {
            Path = address.CanonicalNotesPath,
            Slide = address.SlideIndex,
            Exists = slidePart.NotesSlidePart is not null,
            Text = string.Join('\n', paragraphs),
            Paragraphs = paragraphs,
        };
    }

    // ---- edit ops -----------------------------------------------------------

    /// <summary>set /slide[i]/notes {text}: creates the notes part if absent, else replaces the whole text.</summary>
    public static string Set(PresentationPart presentation, PptxAddress address, JsonObject? props)
    {
        var text = RequireText(props, "set");
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var body = EnsureBodyShape(presentation, slidePart);

        foreach (var paragraph in body.TextBody!.Elements<A.Paragraph>().ToList())
        {
            paragraph.Remove();
        }

        foreach (var line in text.Split('\n'))
        {
            body.TextBody.Append(PptxEditor.BuildParagraph(line, fontSizeHundredths: null, bold: null, colorHex: null, align: null));
        }

        return address.CanonicalNotesPath;
    }

    /// <summary>add /slide[i]/notes {text}: appends one paragraph (creating the notes part if absent).</summary>
    public static string Append(PresentationPart presentation, PptxAddress address, EditOp op)
    {
        if (op.Type is { } type && type.Trim().ToLowerInvariant() is not ("p" or "paragraph" or "notes"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Cannot add '{op.Type}' to speaker notes.",
                "Omit the type (or use \"p\") — add on /slide[i]/notes appends one paragraph: " +
                "{\"op\":\"add\",\"path\":\"/slide[2]/notes\",\"props\":{\"text\":\"...\"}}.",
                candidates: ["p"]);
        }

        var text = RequireText(op.Props, "add");
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var body = EnsureBodyShape(presentation, slidePart);

        // A freshly created notes body carries one empty seed paragraph; an
        // append replaces it so the first added line is paragraph 1.
        var existing = body.TextBody!.Elements<A.Paragraph>().ToList();
        if (existing.Count == 1 && PptxDoc.ParagraphText(existing[0]).Length == 0)
        {
            existing[0].Remove();
        }

        foreach (var line in text.Split('\n'))
        {
            body.TextBody.Append(PptxEditor.BuildParagraph(line, fontSizeHundredths: null, bold: null, colorHex: null, align: null));
        }

        return address.CanonicalNotesPath;
    }

    /// <summary>remove /slide[i]/notes: deletes the notes part (idempotent — clearing absent notes succeeds).</summary>
    public static string Clear(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        if (slidePart.NotesSlidePart is { } notesPart)
        {
            slidePart.DeletePart(notesPart);
        }

        return address.CanonicalNotesPath;
    }

    // ---- part plumbing ------------------------------------------------------

    /// <summary>The notes body placeholder shape, or any text shape as fallback; null when no notes part exists.</summary>
    private static P.Shape? BodyShape(NotesSlidePart? notesPart)
    {
        var tree = notesPart?.NotesSlide?.CommonSlideData?.ShapeTree;
        if (tree is null)
        {
            return null;
        }

        var shapes = tree.Elements<P.Shape>().Where(s => s.TextBody is not null).ToList();
        return shapes.FirstOrDefault(s => PptxDoc.PlaceholderType(s) == "body") ?? shapes.FirstOrDefault();
    }

    /// <summary>Returns the notes body shape, creating the notes part (and notes master wiring) when needed.</summary>
    private static P.Shape EnsureBodyShape(PresentationPart presentation, SlidePart slidePart)
    {
        var notesPart = slidePart.NotesSlidePart ?? CreateNotesPart(presentation, slidePart);
        if (BodyShape(notesPart) is { } body)
        {
            body.TextBody ??= new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph());
            return body;
        }

        // Foreign notes part without a usable text shape: add our own body placeholder.
        var tree = notesPart.NotesSlide?.CommonSlideData?.ShapeTree ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The notes part has no shape tree (p:spTree).",
            "Clear the notes first ({\"op\":\"remove\",\"path\":\"/slide[i]/notes\"}), then set them again.");
        var shape = BuildNotesBodyShape(PptxDoc.NextShapeId(tree));
        tree.Append(shape);
        return shape;
    }

    private static NotesSlidePart CreateNotesPart(PresentationPart presentation, SlidePart slidePart)
    {
        var masterPart = EnsureNotesMaster(presentation);
        var notesPart = slidePart.AddNewPart<NotesSlidePart>();

        var tree = PptxFactory.EmptyShapeTree();
        tree.Append(BuildNotesBodyShape(id: 2));
        notesPart.NotesSlide = new P.NotesSlide(
            new P.CommonSlideData(tree),
            new P.ColorMapOverride(new A.MasterColorMapping()));

        // A notes slide references both its slide and the notes master (the
        // same bidirectional wiring PowerPoint writes).
        notesPart.AddPart(slidePart);
        notesPart.AddPart(masterPart);
        return notesPart;
    }

    /// <summary>The deck's notes master, created and registered in p:notesMasterIdLst on first use.</summary>
    private static NotesMasterPart EnsureNotesMaster(PresentationPart presentation)
    {
        var masterPart = presentation.NotesMasterPart;
        if (masterPart is null)
        {
            masterPart = presentation.AddNewPart<NotesMasterPart>();
            masterPart.NotesMaster = new P.NotesMaster(
                new P.CommonSlideData(PptxFactory.EmptyShapeTree()),
                PptxFactory.BuildColorMap(),
                new P.NotesStyle());
            masterPart.AddNewPart<ThemePart>().Theme = PptxFactory.BuildTheme();
        }

        var root = presentation.Presentation ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The presentation is malformed: p:presentation is missing.",
            "Re-export the file from PowerPoint/Keynote, or restore a snapshot.");
        root.NotesMasterIdList ??= new P.NotesMasterIdList();
        if (!root.NotesMasterIdList.Elements<P.NotesMasterId>().Any())
        {
            // NotesMasterId carries only r:id (the SDK names it Id) — there is no numeric id.
            root.NotesMasterIdList.Append(new P.NotesMasterId { Id = presentation.GetIdOfPart(masterPart) });
        }

        return masterPart;
    }

    private static P.Shape BuildNotesBodyShape(uint id) => new(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = id, Name = "Notes Placeholder" },
            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            new P.ApplicationNonVisualDrawingProperties(
                new P.PlaceholderShape { Type = P.PlaceholderValues.Body, Index = (UInt32Value)1U })),
        new P.ShapeProperties(),
        new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph()));

    private static string RequireText(JsonObject? props, string opName)
    {
        foreach (var (key, _) in props ?? [])
        {
            if (!string.Equals(key, "text", StringComparison.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown notes prop '{key}'.",
                    "Speaker notes accept only the text prop; styling notes is not supported yet.",
                    candidates: ["text"]);
            }
        }

        if (props is null || !props.TryGetPropertyValue("text", out var node) || node is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"{opName} on notes requires props.text.",
                "Pass the notes text, e.g. {\"op\":\"" + opName + "\",\"path\":\"/slide[2]/notes\",\"props\":{\"text\":\"Remember the demo\"}}.");
        }

        return J.ScalarText(node);
    }
}
