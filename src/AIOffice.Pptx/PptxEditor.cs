using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// What one applied op reports back: the canonical target, plus replace counters
/// for replace ops. <see cref="Mutated"/> is false for producing-only ops (extract)
/// that read the document but must not rewrite it.
/// </summary>
internal sealed record PptxOpOutcome(
    string Target,
    int? Replacements = null,
    IReadOnlyList<string>? Locations = null,
    bool Mutated = true,
    IReadOnlyList<Warning>? Warnings = null);

/// <summary>
/// Applies one edit op to an open presentation. Callers batch ops over an
/// in-memory stream and only write the file when every op succeeded (atomic).
/// </summary>
internal static class PptxEditor
{
    private static readonly IReadOnlyList<string> AddTypes =
        ["slide", "shape", "textbox", "image", "chart", "table", "row", "animation", "zoom", "comment", "reply", "embed", "media", "model3d", "equation",
         "smartart", "connector", "actionButton", "group", "ungroup"];

    private static readonly IReadOnlyList<string> ShapePropKeys =
        ["text", "x", "y", "w", "h", "fill", "gradient", "image", "fontSize", "bold", "color", "align", "name", "title", "altText", "altTitle",
         "shadow", "glow", "reflection", "outline", "autofit"];

    /// <summary>add shape additionally accepts a preset geometry and a flip.</summary>
    private static readonly IReadOnlyList<string> AddShapePropKeys = [.. ShapePropKeys, "shape", "flip"];

    /// <summary>The preset geometries add shape understands ("line" builds a connector instead).</summary>
    private static readonly IReadOnlyDictionary<string, A.ShapeTypeValues> GeometryPresets =
        new Dictionary<string, A.ShapeTypeValues>(StringComparer.OrdinalIgnoreCase)
        {
            ["rect"] = A.ShapeTypeValues.Rectangle,
            ["roundRect"] = A.ShapeTypeValues.RoundRectangle,
            ["ellipse"] = A.ShapeTypeValues.Ellipse,
            ["triangle"] = A.ShapeTypeValues.Triangle,
            ["diamond"] = A.ShapeTypeValues.Diamond,
            ["arrow"] = A.ShapeTypeValues.RightArrow,
        };

    private static readonly IReadOnlyList<string> GeometryTokens =
        ["rect", "roundRect", "ellipse", "triangle", "diamond", "arrow", "line"];

    private static readonly IReadOnlyList<string> ZOrderPositions = ["front", "back", "forward", "backward"];

    /// <summary>Default stroke width for line shapes (1.5pt).</summary>
    private const int LineWidthEmu = 19_050;

    /// <summary>Applies the op and returns the canonical path of the affected node (plus replace counters).</summary>
    public static PptxOpOutcome Apply(PresentationDocument document, PresentationPart presentation, EditOp op, Workspace workspace)
    {
        var parsed = PptxAddress.Parse(op.Path);
        if (parsed.IsSmartArt)
        {
            throw PptxSmartArt.EditUnsupported(op.Path);
        }

        // /properties is a package-level node: core + custom document metadata.
        if (parsed.IsProperties)
        {
            return new PptxOpOutcome(ApplyProperties(document, op));
        }

        switch (op.Op)
        {
            case "add" when string.Equals(op.Type?.Trim(), "equation", StringComparison.OrdinalIgnoreCase):
                return ApplyAddEquation(presentation, op);
            case "add" when string.Equals(op.Type?.Trim(), "model3d", StringComparison.OrdinalIgnoreCase):
                return ApplyAddModel3D(document, presentation, op, workspace);
            case "add":
                return new PptxOpOutcome(ApplyAdd(document, presentation, op, workspace));
            case "set":
                return new PptxOpOutcome(ApplySet(presentation, op, workspace));
            case "remove":
                return new PptxOpOutcome(ApplyRemove(presentation, op));
            case "move":
                return new PptxOpOutcome(ApplyMove(presentation, op));
            case "replace":
                var result = PptxReplace.Apply(presentation, op);
                return new PptxOpOutcome(result.Target, result.Replacements, result.Locations);
            case "extract":
                return new PptxOpOutcome(ApplyExtract(presentation, op, workspace), Mutated: false);
            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown op '{op.Op}' for pptx.",
                    "Use set, add, remove, move, replace or extract (accept/reject apply to docx tracked changes).",
                    candidates: ["set", "add", "remove", "move", "replace", "extract"]);
        }
    }

    private static string ApplyAdd(PresentationDocument document, PresentationPart presentation, EditOp op, Workspace workspace)
    {
        var address = PptxAddress.Parse(op.Path);
        if (address.IsNotes)
        {
            return PptxNotes.Append(presentation, address, op);
        }

        if (address.IsFonts)
        {
            var fontType = op.Type?.Trim().ToLowerInvariant();
            if (fontType == "font")
            {
                if (address.FontName is not null)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"add font targets the /fonts list, not '{op.Path}'.",
                        "Use {\"op\":\"add\",\"path\":\"/fonts\",\"type\":\"font\",\"props\":{\"src\":\"MyFont.ttf\"}}.");
                }

                return PptxFonts.Add(presentation, op.Props, workspace);
            }

            throw new AiofficeException(
                op.Type is null ? ErrorCodes.InvalidArgs : ErrorCodes.UnsupportedFeature,
                op.Type is null ? "add on /fonts requires a type." : $"Cannot add '{op.Type}' on '{op.Path}'.",
                "The only thing you can add on /fonts is a font " +
                "({\"op\":\"add\",\"path\":\"/fonts\",\"type\":\"font\",\"props\":{\"src\":\"MyFont.ttf\"}}).",
                candidates: ["font"]);
        }

        if (address.IsNotesMaster || address.IsHandoutMaster)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{op.Path}' is a whole-part root; nothing is added to it.",
                "Edit it with set, e.g. {\"op\":\"set\",\"path\":\"" + op.Path + "\",\"props\":{\"background\":\"0F172A\"}}.");
        }

        if (address.IsMaster)
        {
            return PptxMasters.Add(presentation, address, op);
        }

        if (address.IsPresentation || address.IsSection)
        {
            var sectionType = op.Type?.Trim().ToLowerInvariant();
            if (sectionType == "section")
            {
                return PptxSections.Add(presentation, address, op.Props);
            }

            throw new AiofficeException(
                op.Type is null ? ErrorCodes.InvalidArgs : ErrorCodes.UnsupportedFeature,
                op.Type is null ? "add on '/' requires a type." : $"Cannot add '{op.Type}' on '{op.Path}'.",
                "The only thing you can add at the presentation root is a section " +
                "({\"op\":\"add\",\"path\":\"/\",\"type\":\"section\",\"props\":{\"name\":\"Intro\"}}); " +
                "add slides via /slide[i].",
                candidates: ["section"]);
        }

        var type = op.Type?.Trim().ToLowerInvariant();
        switch (type)
        {
            case "slide":
                return AddSlide(presentation, address, op.Position, op.Props);

            case "section":
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"add section targets the presentation root, not '{op.Path}'.",
                    "Use the root path: {\"op\":\"add\",\"path\":\"/\",\"type\":\"section\",\"props\":{\"name\":\"Intro\",\"afterSlide\":0}}.");

            case "shape" or "textbox":
            {
                var slidePart = ResolveAddTargetSlide(presentation, address, op.Path, "shape");
                var id = AddTextBox(slidePart, op.Props, workspace);
                return Units.Inv($"/slide[{address.SlideIndex}]/shape[@id={id}]");
            }

            case "image" or "picture":
            {
                var slidePart = ResolveAddTargetSlide(presentation, address, op.Path, "image");
                var id = PptxImages.AddImage(slidePart, op.Props, workspace);
                return Units.Inv($"/slide[{address.SlideIndex}]/shape[@id={id}]");
            }

            case "chart":
            {
                var slidePart = ResolveAddTargetSlide(presentation, address, op.Path, "chart");
                var id = PptxCharts.Add(slidePart, op.Props);
                return Units.Inv($"/slide[{address.SlideIndex}]/shape[@id={id}]");
            }

            case "smartart":
            {
                var slidePart = ResolveAddTargetSlide(presentation, address, op.Path, "smartart");
                return PptxDiagrams.Add(slidePart, address.SlideIndex, op.Props);
            }

            case "connector":
            {
                var slidePart = ResolveAddTargetSlide(presentation, address, op.Path, "connector");
                var id = PptxConnectors.Add(slidePart, op.Props);
                return Units.Inv($"/slide[{address.SlideIndex}]/shape[@id={id}]");
            }

            case "actionbutton":
            {
                var slidePart = ResolveAddTargetSlide(presentation, address, op.Path, "actionButton");
                return PptxActionButtons.Add(presentation, slidePart, address.SlideIndex, op.Props);
            }

            case "group":
            {
                if (address.IsNotes)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"add group targets a slide, not '{op.Path}'.",
                        "Use the slide path with a shapes list: {\"op\":\"add\",\"path\":\"/slide[1]\"," +
                        "\"type\":\"group\",\"props\":{\"shapes\":[\"@id1\",\"@id2\"]}}.");
                }

                var slidePart = ResolveAddTargetSlide(presentation, address, op.Path, "group");
                return PptxGroups.Group(slidePart, address.SlideIndex, op.Props);
            }

            case "ungroup":
            {
                if (!address.IsGroup || address.HasShape)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"add ungroup targets a group, not '{op.Path}'.",
                        "Use a group path: {\"op\":\"add\",\"path\":\"/slide[1]/group[@id=5]\",\"type\":\"ungroup\"} — " +
                        "it dissolves the group, promoting its children back onto the slide with absolute coordinates.");
                }

                if (op.Props is { Count: > 0 })
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        "ungroup takes no props.",
                        "Send it without props: {\"op\":\"add\",\"path\":\"" + op.Path + "\",\"type\":\"ungroup\"}.");
                }

                var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, op.Path);
                return PptxGroups.Ungroup(slidePart, address);
            }

            case "embed":
            {
                var slidePart = ResolveAddTargetSlide(presentation, address, op.Path, "embed");
                return PptxEmbeds.Add(slidePart, address.SlideIndex, op.Props, workspace);
            }

            case "media":
            {
                var slidePart = ResolveAddTargetSlide(presentation, address, op.Path, "media");
                return PptxMedia.Add(document, slidePart, address.SlideIndex, op.Props, workspace);
            }

            case "table":
            {
                var slidePart = ResolveAddTargetSlide(presentation, address, op.Path, "table");
                var index = PptxTables.Add(slidePart, op.Props);
                return Units.Inv($"/slide[{address.SlideIndex}]/table[{index}]");
            }

            case "row":
            {
                if (!address.IsTable)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"add row targets a table, not '{op.Path}'.",
                        "Use {\"op\":\"add\",\"path\":\"/slide[2]/table[1]/tr[2]\",\"type\":\"row\"} — " +
                        "the new row becomes tr[2]; omit /tr[r] to append.");
                }

                return PptxTables.AddRow(presentation, address, op.Position);
            }

            case "animation":
                return PptxAnimations.Add(presentation, address, op.Props);

            case "zoom":
            {
                if (address.HasShape || address.IsChart || address.IsTable || address.IsAnimation ||
                    address.IsZoom || address.IsComment || address.IsGroup || address.IsSmartArt)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"add zoom targets a slide, not '{op.Path}'.",
                        "Use the slide path: {\"op\":\"add\",\"path\":\"/slide[2]\",\"type\":\"zoom\"," +
                        "\"props\":{\"kind\":\"slide\",\"target\":\"slide 3\"}}.");
                }

                return PptxZoom.Add(presentation, address, op.Props);
            }

            case "comment":
            {
                if (address.IsComment)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"add comment targets a slide; '{op.Path}' is a comment.",
                        "To reply to it, use type reply: {\"op\":\"add\",\"path\":\"" + op.Path +
                        "\",\"type\":\"reply\",\"props\":{\"text\":\"...\"}}.");
                }

                if (address.HasShape || address.IsChart || address.IsTable || address.IsAnimation)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"add comment targets a slide, not '{op.Path}'.",
                        "Use the slide path: {\"op\":\"add\",\"path\":\"/slide[2]\",\"type\":\"comment\"," +
                        "\"props\":{\"text\":\"...\"}}.");
                }

                return PptxComments.Add(presentation, address, op.Props);
            }

            case "reply":
            {
                if (!address.IsComment)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"add reply targets a comment, not '{op.Path}'.",
                        "Use the parent comment's path: {\"op\":\"add\",\"path\":\"/slide[2]/comment[@id=1]\"," +
                        "\"type\":\"reply\",\"props\":{\"text\":\"...\"}}.");
                }

                return PptxComments.AddReply(presentation, address, op.Props);
            }

            default:
                throw new AiofficeException(
                    op.Type is null ? ErrorCodes.InvalidArgs : ErrorCodes.UnsupportedFeature,
                    op.Type is null ? "add requires a type." : $"Cannot add '{op.Type}' yet.",
                    "Addable types today: slide, shape (textbox or preset geometry), image (PNG/JPEG picture), " +
                    "chart (bar/line/pie/doughnut/radar/bubble/stackedBar/percentStackedBar/stackedArea/combo), " +
                    "table (with rows/cols), row (on a table path), " +
                    "animation (on a shape path), zoom (slide/section/summary navigation, on a slide path), " +
                    "comment and reply (on a comment path), embed (a file as an OLE object), " +
                    "media (an mp4/mov video or m4a/mp3/wav audio clip), " +
                    "model3d (a glb/gltf 3D model behind a poster fallback), " +
                    "equation (LaTeX -> OMML math in a text box), " +
                    "smartart (a list/process/hierarchy/orgChart/cycle diagram), connector (a line/elbow/curved link between two shapes), " +
                    "actionButton (a navigation button: first/last/next/prev/home/end/slide/url), " +
                    "group (wrap shapes in a group), ungroup (dissolve a group, on a group path).",
                    candidates: AddTypes);
        }
    }

    /// <summary>
    /// extract op: writes an embedded object's payload to props.to (sandbox-resolved).
    /// It does not modify the source deck — the resolved address is returned so the
    /// op echoes a stable target, but the bytes flow out, not in.
    /// </summary>
    private static string ApplyExtract(PresentationPart presentation, EditOp op, Workspace workspace)
    {
        var address = PptxAddress.Parse(op.Path);
        if (!address.IsEmbed)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"extract targets an embedded object, not '{op.Path}'.",
                "Use an embed path, e.g. {\"op\":\"extract\",\"path\":\"/slide[2]/embed[1]\"," +
                "\"props\":{\"to\":\"out/data.xlsx\"}} — run 'aioffice read <file> --view embeds' to list embeds.");
        }

        if (op.Props is null ||
            !op.Props.TryGetPropertyValue("to", out var toNode) || toNode is null || J.ScalarText(toNode).Trim().Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "extract requires props.to (the destination path).",
                "Pass a workspace destination, e.g. " +
                "{\"op\":\"extract\",\"path\":\"" + op.Path + "\",\"props\":{\"to\":\"out/data.xlsx\"}}.");
        }

        foreach (var (key, _) in op.Props)
        {
            if (!string.Equals(key, "to", StringComparison.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown extract prop '{key}'.",
                    "extract takes only 'to' (the sandbox destination path).",
                    candidates: ["to"]);
            }
        }

        // Sandbox the destination: a path outside the workspace is sandbox_denied,
        // never written. The parent need not exist yet (Extract creates it).
        var dest = workspace.Resolve(J.ScalarText(toNode).Trim());
        return PptxEmbeds.Extract(presentation, address, dest);
    }

    /// <summary>
    /// add equation: inserts a LaTeX equation as OMML into a slide text body and
    /// echoes its canonical /omath[k] path. Unknown LaTeX commands degrade to
    /// literal runs and raise an equation_partial warning — the docx contract.
    /// </summary>
    private static PptxOpOutcome ApplyAddEquation(PresentationPart presentation, EditOp op)
    {
        var address = PptxAddress.Parse(op.Path);
        var (path, unknown) = PptxEquations.Add(presentation, address, op.Props);

        IReadOnlyList<Warning>? warnings = null;
        if (unknown.Count > 0)
        {
            warnings =
            [
                new Warning(
                    "equation_partial",
                    "The equation rendered, but these LaTeX tokens were not recognized and appear literally: " +
                    string.Join(", ", unknown.Distinct()) + ". " +
                    "Check the spelling, or split the unsupported part into a \\text{…} run."),
            ];
        }

        return new PptxOpOutcome(path, Warnings: warnings);
    }

    /// <summary>
    /// add model3d: embeds a glb/gltf 3D model behind a poster picture fallback and
    /// echoes its canonical /model3d[@id=N] path with an honest media-backed note.
    /// </summary>
    private static PptxOpOutcome ApplyAddModel3D(PresentationDocument document, PresentationPart presentation, EditOp op, Workspace workspace)
    {
        var address = PptxAddress.Parse(op.Path);
        var slidePart = ResolveAddTargetSlide(presentation, address, op.Path, "model3d");
        var (path, note) = PptxModels.Add(document, slidePart, address.SlideIndex, op.Props, workspace);
        return new PptxOpOutcome(path, Warnings: [note]);
    }

    private static SlidePart ResolveAddTargetSlide(PresentationPart presentation, PptxAddress address, string path, string what)
    {
        if (address.HasShape || address.IsChart || address.IsTable || address.IsAnimation || address.IsComment ||
            address.IsEmbed || address.IsMedia || address.IsModel3D || address.IsGroup || address.IsSmartArt)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add {what} targets a slide, not '{path}'.",
                Units.Inv($"Use the slide path as the target, e.g. {{\"op\":\"add\",\"path\":\"/slide[2]\",\"type\":\"{what}\"}}."));
        }

        return PptxDoc.ResolveSlide(presentation, address.SlideIndex, path);
    }

    private static string AddSlide(PresentationPart presentation, PptxAddress address, string? position, JsonObject? props)
    {
        if (address.HasShape || address.IsChart)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add slide targets a slide position, not '{address.Raw}'.",
                "Use e.g. {\"op\":\"add\",\"path\":\"/slide[3]\",\"type\":\"slide\"} — the new slide becomes slide 3.");
        }

        var slideIdList = presentation.Presentation?.SlideIdList
            ?? throw CorruptPresentation("p:sldIdLst is missing");
        var count = slideIdList.Elements<P.SlideId>().Count();

        var target = position?.Trim().ToLowerInvariant() switch
        {
            null or "" or "at" or "before" => address.SlideIndex,
            "after" => address.SlideIndex + 1,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown position '{position}' for add slide.",
                "Use \"at\"/\"before\" (new slide takes the path's index) or \"after\".",
                candidates: ["at", "before", "after"]),
        };

        if (target < 1 || target > count + 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"Cannot insert a slide at position {target}; the deck has {count} slide(s).",
                Units.Inv($"Valid positions are 1..{count + 1} (use /slide[{count + 1}] to append)."),
                candidates: [.. Enumerable.Range(1, Math.Min(count + 1, 10)).Select(i => Units.Inv($"/slide[{i}]"))]);
        }

        var layoutPart = PickLayout(presentation, props);

        var slidePart = presentation.AddNewPart<SlidePart>();
        slidePart.Slide = PptxFactory.BuildBlankSlide();
        slidePart.AddPart(layoutPart);

        var slideId = new P.SlideId
        {
            Id = PptxDoc.NextSlideId(slideIdList),
            RelationshipId = presentation.GetIdOfPart(slidePart),
        };
        slideIdList.InsertAt(slideId, target - 1);

        if (props is not null && props.TryGetPropertyValue("title", out var titleNode) && titleNode is not null)
        {
            AddTitleShape(slidePart, J.ScalarText(titleNode));
        }

        if (props is not null && props.TryGetPropertyValue("background", out var backgroundNode))
        {
            SetBackground(slidePart, backgroundNode);
        }

        return Units.Inv($"/slide[{target}]");
    }

    /// <summary>
    /// The layout a new slide binds to: props.layout is a 1-based index into the
    /// first master's layouts (read --view structure lists them); the default
    /// stays the master's first layout.
    /// </summary>
    private static SlideLayoutPart PickLayout(PresentationPart presentation, JsonObject? props)
    {
        var masters = PptxDoc.Masters(presentation);
        if (masters.Count == 0)
        {
            throw CorruptPresentation("no slide master part exists");
        }

        var layouts = PptxDoc.Layouts(masters[0].Part);
        if (layouts.Count == 0)
        {
            throw CorruptPresentation("no slide layout part exists");
        }

        // A layout name binds a slide to a layout by its display name ("My Layout");
        // it is a distinct prop from the 1-based layout index so the index path keeps
        // its strict numeric contract. layoutName wins when both are present.
        if (props is not null && props.TryGetPropertyValue("layoutName", out var nameNode) && nameNode is not null)
        {
            var wantedName = J.ScalarText(nameNode).Trim();
            foreach (var (_, layoutPart) in layouts)
            {
                if (string.Equals(PptxDoc.LayoutName(layoutPart), wantedName, StringComparison.OrdinalIgnoreCase))
                {
                    return layoutPart;
                }
            }

            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"No layout named '{wantedName}' on master 1.",
                "Use a layout name from the deck (run 'aioffice read <file> --view structure'), or a 1-based 'layout' index.",
                candidates: [.. layouts.Select(l => PptxDoc.LayoutName(l.Part)).Where(n => !string.IsNullOrEmpty(n)).Take(10).Select(n => n!)]);
        }

        if (props is null || !props.TryGetPropertyValue("layout", out var layoutNode))
        {
            return layouts[0].Part;
        }

        // Props arrive string-valued through the CLI sugar and the MCP schema
        // ({"layout":"2"}), and as JSON numbers from hand-written ops — accept both.
        double number = 0;
        var numeric = layoutNode is JsonValue value &&
            (Units.TryNumber(value, out number) ||
             (value.TryGetValue<string>(out var raw) &&
              double.TryParse(raw, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out number)));
        if (!numeric || number != Math.Floor(number) || number < 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.layout is not a valid layout index: {layoutNode?.ToJsonString() ?? "null"}",
                "Use a 1-based integer index into the master's layouts (or 'layoutName' to bind by name); " +
                "run 'aioffice read <file> --view structure' to list them.");
        }

        var index = (int)number;
        if (index > layouts.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"props.layout is {index} but master 1 has only {layouts.Count} layout(s)."),
                "Run 'aioffice read <file> --view structure' to list the master's layouts.",
                candidates: [.. layouts.Take(10).Select(l => Units.Inv($"/master[1]/layout[{l.Index}]"))]);
        }

        return layouts[index - 1].Part;
    }

    /// <summary>Routes a /properties op: only <c>set</c> is meaningful (write core/custom metadata).</summary>
    private static string ApplyProperties(PresentationDocument document, EditOp op)
    {
        if (op.Op != "set")
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{op.Op}' is not supported on /properties.",
                "Use set to write metadata: {\"op\":\"set\",\"path\":\"/properties\"," +
                "\"props\":{\"title\":\"Q3 Deck\",\"custom\":{\"Reviewed\":true}}}. " +
                "Clear a value by setting it to \"\" (core) or null (custom).");
        }

        if (op.Props is null || op.Props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "set /properties has no props.",
                "Pass core props and/or a custom object, e.g. " +
                "{\"op\":\"set\",\"path\":\"/properties\",\"props\":{\"author\":\"Dana\",\"custom\":{\"Project\":\"Acme\"}}}.");
        }

        return PptxProperties.Set(document, op.Props);
    }

    private static string ApplySet(PresentationPart presentation, EditOp op, Workspace workspace)
    {
        var address = PptxAddress.Parse(op.Path);
        if (op.Props is null || op.Props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"set on '{op.Path}' has no props.",
                "Pass props, e.g. {\"op\":\"set\",\"path\":\"/slide[1]/shape[2]\",\"props\":{\"text\":\"Hello\"}}.");
        }

        if (address.IsNotes)
        {
            return PptxNotes.Set(presentation, address, op.Props);
        }

        if (address.IsFonts)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "Embedded fonts cannot be edited in place.",
                "Remove the font ({\"op\":\"remove\",\"path\":\"/fonts/font[@name=...]\"}) and add it again with the new file; " +
                "add a font with {\"op\":\"add\",\"path\":\"/fonts\",\"type\":\"font\",\"props\":{\"src\":\"MyFont.ttf\"}}.");
        }

        if (address.IsPresentation)
        {
            return PptxSlideSize.Set(presentation, op.Props);
        }

        if (address.IsSection)
        {
            return PptxSections.Set(presentation, address, op.Props);
        }

        if (address.IsNotesMaster || address.IsHandoutMaster)
        {
            return PptxNotesHandoutMasters.Set(presentation, address, op.Props);
        }

        if (address.IsMaster)
        {
            return PptxMasters.Set(presentation, address, op.Props, workspace);
        }

        if (address.IsChart)
        {
            if (op.Props.ContainsKey("embedData"))
            {
                return SetChartEmbedData(presentation, address, op.Props);
            }

            // title/categories/series edit the chart's data in place (v1.12); the
            // chart-polish props (dataLabels/legend/axisTitles/trendline/errorBars/
            // gridlines/secondaryAxis) edit its presentation in place. Both may
            // appear in one op: apply the data edit first, then the polish.
            var handlesData = PptxCharts.HandlesData(op.Props);
            var handlesPolish = PptxChartPolish.Handles(op.Props);
            if (handlesData || handlesPolish)
            {
                var data = SplitChartProps(op.Props, PptxCharts.DataProps, out var afterData);
                var polish = SplitChartProps(afterData, PptxChartPolish.PropKeys, out var unknown);
                if (unknown.Count > 0)
                {
                    throw new AiofficeException(
                        ErrorCodes.UnsupportedFeature,
                        Units.Inv($"Chart prop(s) {string.Join(", ", unknown.Select(kv => kv.Key))} cannot be set in place."),
                        "Editable in place: data props (title, categories, series) and polish props " +
                        "(dataLabels, legend, axisTitles, trendline, errorBars, gridlines, secondaryAxis); " +
                        "position/size/name sets target the chart's shape path.");
                }

                string target;
                if (data.Count > 0)
                {
                    target = PptxCharts.SetData(presentation, address, data);
                }
                else
                {
                    target = address.CanonicalChartPath;
                }

                if (polish.Count > 0)
                {
                    target = PptxCharts.SetPolish(presentation, address, polish);
                }

                return target;
            }

            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                Units.Inv($"Chart prop(s) {string.Join(", ", op.Props.Select(kv => kv.Key))} cannot be set in place."),
                "Editable in place: data props (title, categories, series) and polish props " +
                "(dataLabels, legend, axisTitles, trendline, errorBars, gridlines, secondaryAxis); " +
                "{\"embedData\":true} retrofits an editable workbook; position/size/name sets target " +
                "the chart's shape path (" + address.CanonicalChartPath + ").");
        }

        if (address.IsTable)
        {
            return PptxTables.Set(presentation, address, op.Props);
        }

        if (address.IsAnimation)
        {
            return PptxAnimations.Set(presentation, address, op.Props);
        }

        if (address.IsZoom)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "A zoom's kind and target cannot be edited in place.",
                "Remove the zoom ({\"op\":\"remove\",\"path\":\"" + address.CanonicalZoomPath + "\"}) and add it again " +
                "with the new kind/target; position/size/name sets target its shape path (/slide[i]/shape[@id=N]).");
        }

        if (address.IsComment)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "Comments cannot be edited in place yet.",
                "Remove the comment ({\"op\":\"remove\",\"path\":\"" + address.CanonicalCommentPath + "\"}) " +
                "and add a new one with the corrected text.");
        }

        if (address.IsEmbed)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "Embedded objects cannot be edited in place.",
                "Remove the embed ({\"op\":\"remove\",\"path\":\"" + op.Path + "\"}) and add it again with the new file, " +
                "or extract it ({\"op\":\"extract\",...}) to pull its bytes out.");
        }

        if (address.IsMedia)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "Embedded media cannot be edited in place.",
                "Remove the media ({\"op\":\"remove\",\"path\":\"" + op.Path + "\"}) and add it again with the new clip; " +
                "position/size sets target its shape path (/slide[i]/shape[@id=N]).");
        }

        if (address.IsModel3D)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "Embedded 3D models cannot be edited in place.",
                "Remove the model ({\"op\":\"remove\",\"path\":\"" + op.Path + "\"}) and add it again with the new file; " +
                "position/size sets target its shape path (/slide[i]/shape[@id=N]).");
        }

        if (address.IsGroup)
        {
            return SetGroupProps(presentation, address, op.Props, workspace);
        }

        if (!address.HasShape)
        {
            return SetSlideProps(presentation, address, op.Props, workspace);
        }

        if (address.RunIndex is not null)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "Run-level set is not supported yet.",
                "Set the paragraph instead: target /slide[i]/shape[j]/p[k] with a text prop.");
        }

        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, op.Path);
        var view = PptxDoc.ResolveShape(slidePart, address);

        if (address.ParagraphIndex is not null)
        {
            SetParagraphProps(view, address, op.Props);
            return Units.Inv($"{view.CanonicalPath(address.SlideIndex)}/p[{address.ParagraphIndex}]");
        }

        // M8 shape hyperlinks/actions live on the cNvPr and need the presentation
        // (slide-jump targets resolve against the slide list), so they are split
        // out here and the remaining props flow to the format-agnostic setter.
        var rest = op.Props;
        if (PptxHyperlinks.Handles(op.Props))
        {
            rest = SplitHyperlinkProps(op.Props, out var linkProps);
            PptxHyperlinks.Apply(presentation, slidePart, view, linkProps);
        }

        if (rest.Count > 0)
        {
            SetShapeProps(view, rest, slidePart, workspace);
        }

        return view.CanonicalPath(address.SlideIndex);
    }

    /// <summary>Peels hyperlink/linkText out of a props object, returning the rest (the original is not mutated).</summary>
    private static JsonObject SplitHyperlinkProps(JsonObject props, out JsonObject linkProps)
    {
        linkProps = [];
        var rest = new JsonObject();
        foreach (var (key, value) in props)
        {
            if (key is "hyperlink" or "linkText")
            {
                linkProps[key] = value?.DeepClone();
            }
            else
            {
                rest[key] = value?.DeepClone();
            }
        }

        return rest;
    }

    /// <summary>Pulls the named keys out of a chart props object, returning them and (via out) the rest.</summary>
    private static JsonObject SplitChartProps(JsonObject props, IReadOnlyList<string> keys, out JsonObject rest)
    {
        var taken = new JsonObject();
        rest = new JsonObject();
        foreach (var (key, value) in props)
        {
            if (keys.Contains(key, StringComparer.Ordinal))
            {
                taken[key] = value?.DeepClone();
            }
            else
            {
                rest[key] = value?.DeepClone();
            }
        }

        return taken;
    }

    /// <summary>set /slide[i]/chart[k] {embedData:true}: retrofit an embedded, editable data workbook.</summary>
    private static string SetChartEmbedData(PresentationPart presentation, PptxAddress address, JsonObject props)
    {
        if (props.Count > 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "embedData cannot be combined with other chart props.",
                "Send {\"op\":\"set\",\"path\":\"" + address.CanonicalChartPath + "\",\"props\":{\"embedData\":true}} on its own.");
        }

        if (!props.TryGetPropertyValue("embedData", out var node) || node is null || !AsBool("embedData", node))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "embedData only supports the value true (charts cannot be un-embedded).",
                "Use {\"embedData\":true} to retrofit the workbook; to drop it, re-add the chart in PowerPoint.");
        }

        return PptxCharts.EmbedData(presentation, address);
    }

    /// <summary>
    /// set on a group path: a group child (/group[@id=N]/shape[...]) takes normal shape
    /// props; the group itself (/group[@id=N]) takes name, altText and altTitle (resize a
    /// group by ungrouping first, then move/resize the shapes).
    /// </summary>
    private static string SetGroupProps(PresentationPart presentation, PptxAddress address, JsonObject props, Workspace workspace)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var tree = PptxDoc.RequireShapeTree(slidePart);
        var group = PptxGroups.ResolveGroup(tree, address);

        if (address.HasShape)
        {
            var child = PptxGroups.ResolveChild(group, address);
            SetShapeProps(child, props, slidePart, workspace);
            return Units.Inv($"{address.CanonicalGroupPath}/shape[@id={child.Id}]");
        }

        var groupView = new ShapeView(
            group,
            group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? address.GroupId ?? 0,
            0,
            "group",
            group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty);

        foreach (var (key, value) in props)
        {
            switch (key)
            {
                case "name":
                    var nonVisual = group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties
                        ?? throw CorruptPresentation("group has no p:cNvPr");
                    nonVisual.Name = value is null ? string.Empty : J.ScalarText(value);
                    break;
                case "altText":
                    SetAltText(groupView, value);
                    break;
                case "altTitle":
                    SetAltTitle(groupView, value);
                    break;
                default:
                    throw new AiofficeException(
                        ErrorCodes.UnsupportedFeature,
                        $"Prop '{key}' cannot be set on a group.",
                        "A group takes name, altText and altTitle; set geometry/fill/text on its child shapes " +
                        "(/slide[i]/group[@id=N]/shape[...]), or ungroup first to move/resize the shapes freely.",
                        candidates: ["name", "altText", "altTitle"]);
            }
        }

        return address.CanonicalGroupPath;
    }

    /// <summary>Slide-level set: solid/gradient/image background, transition and transition duration.</summary>
    private static string SetSlideProps(PresentationPart presentation, PptxAddress address, JsonObject props, Workspace workspace)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var slideData = slidePart.Slide?.CommonSlideData ?? throw CorruptPresentation("the slide has no p:cSld");
        JsonNode? transitionNode = null, durationNode = null;
        var hasTransition = false;
        var hasDuration = false;

        foreach (var (key, value) in props)
        {
            switch (key)
            {
                case "background":
                    SetBackground(slidePart, value);
                    break;
                case "gradient":
                    PptxFill.ApplyToBackground(slideData, PptxFill.BuildGradientFill(value));
                    break;
                case "image":
                    PptxFill.ApplyToBackground(slideData, PptxFill.BuildImageFill(value, slidePart, workspace));
                    break;
                case "transition":
                    transitionNode = value;
                    hasTransition = true;
                    break;
                case "transitionDuration":
                    durationNode = value;
                    hasDuration = true;
                    break;
                default:
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Prop '{key}' does not apply to a slide.",
                        "Slide props: background, gradient, image, transition, transitionDuration. " +
                        "Shape props (text, fill, geometry, …) target /slide[i]/shape[j].",
                        candidates: ["background", "gradient", "image", "transition", "transitionDuration"]);
            }
        }

        if (hasTransition || hasDuration)
        {
            PptxTransitions.Set(slidePart, transitionNode, hasTransition, durationNode, hasDuration);
        }

        return address.CanonicalSlidePath;
    }

    /// <summary>Sets a proper p:bg solid fill on a slide (replacing any previous background).</summary>
    internal static void SetBackground(SlidePart slidePart, JsonNode? value)
    {
        var slideData = slidePart.Slide?.CommonSlideData ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The slide has no common slide data (p:cSld).",
            "The slide part is malformed; re-export the file or restore a snapshot.");
        SetBackground(slideData, value);
    }

    /// <summary>Sets a proper p:bg solid fill on any p:cSld (slide, master or layout), replacing any previous one.</summary>
    internal static void SetBackground(P.CommonSlideData slideData, JsonNode? value)
    {
        if (value is JsonObject or JsonArray ||
            (value is JsonValue v && v.TryGetValue<string>(out var raw) && LooksLikeGradientOrImage(raw)))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "Gradient and image backgrounds are not supported.",
                "Use a solid color, e.g. {\"background\":\"0F172A\"} — or add a full-bleed picture: " +
                "{\"op\":\"add\",\"type\":\"image\",\"props\":{\"src\":\"bg.png\",\"x\":0,\"y\":0,\"w\":\"33.87cm\",\"h\":\"19.05cm\"}}.");
        }

        var hex = Units.ParseColorHex("background", value);
        slideData.Background = new P.Background(
            new P.BackgroundProperties(
                new A.SolidFill(new A.RgbColorModelHex { Val = hex }),
                new A.EffectList()));
    }

    private static bool LooksLikeGradientOrImage(string raw)
    {
        var text = raw.Trim().ToLowerInvariant();
        return text.Contains("gradient", StringComparison.Ordinal)
            || text.StartsWith("image", StringComparison.Ordinal)
            || text.StartsWith("url(", StringComparison.Ordinal)
            || text.EndsWith(".png", StringComparison.Ordinal)
            || text.EndsWith(".jpg", StringComparison.Ordinal)
            || text.EndsWith(".jpeg", StringComparison.Ordinal);
    }

    private static string ApplyRemove(PresentationPart presentation, EditOp op)
    {
        var address = PptxAddress.Parse(op.Path);
        if (address.IsNotes)
        {
            return PptxNotes.Clear(presentation, address);
        }

        if (address.IsFonts)
        {
            if (address.FontName is null)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "remove targets one embedded font, not the whole /fonts list.",
                    "Name the font: {\"op\":\"remove\",\"path\":\"/fonts/font[@name=MyFont]\"} — " +
                    "run 'aioffice get <file> /fonts' to list them.");
            }

            return PptxFonts.Remove(presentation, address);
        }

        if (address.IsPresentation)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "The presentation root cannot be removed.",
                "Remove a slide (/slide[i]), a section (/section[i]) or a layout (/master[m]/layout[l]) instead.");
        }

        if (address.IsSection)
        {
            return PptxSections.Remove(presentation, address);
        }

        if (address.IsNotesMaster || address.IsHandoutMaster)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"The {(address.IsNotesMaster ? "notes" : "handout")} master cannot be removed (every deck keeps one).",
                "Reset its look with set, e.g. {\"op\":\"set\",\"path\":\"" + op.Path +
                "\",\"props\":{\"background\":\"FFFFFF\"}}.");
        }

        if (address.IsMaster)
        {
            return PptxMasters.Remove(presentation, address);
        }

        if (address.IsChart)
        {
            return PptxCharts.Remove(presentation, address);
        }

        if (address.IsTable)
        {
            return PptxTables.Remove(presentation, address);
        }

        if (address.IsAnimation)
        {
            return PptxAnimations.Remove(presentation, address);
        }

        if (address.IsZoom)
        {
            return PptxZoom.Remove(presentation, address);
        }

        if (address.IsComment)
        {
            return PptxComments.Remove(presentation, address);
        }

        if (address.IsEmbed)
        {
            return PptxEmbeds.Remove(presentation, address);
        }

        if (address.IsMedia)
        {
            return PptxMedia.Remove(presentation, address);
        }

        if (address.IsModel3D)
        {
            return PptxModels.Remove(presentation, address);
        }

        if (address.IsOMath)
        {
            return PptxEquations.Remove(presentation, address);
        }

        if (address.RunIndex is not null)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "Run-level remove is not supported yet.",
                "Remove the paragraph (/slide[i]/shape[j]/p[k]) or set the paragraph text instead.");
        }

        if (address.IsGroup)
        {
            var groupSlidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, op.Path);
            var groupTree = PptxDoc.RequireShapeTree(groupSlidePart);
            var group = PptxGroups.ResolveGroup(groupTree, address);

            // remove on a group child deletes just that child; on the group itself it
            // deletes the whole group (use type ungroup to dissolve and keep the children).
            if (address.HasShape)
            {
                var child = PptxGroups.ResolveChild(group, address);
                var childPath = Units.Inv($"{address.CanonicalGroupPath}/shape[@id={child.Id}]");
                child.Element.Remove();
                return childPath;
            }

            group.Remove();
            return address.CanonicalGroupPath;
        }

        if (!address.HasShape)
        {
            RemoveSlide(presentation, address);
            return address.CanonicalSlidePath;
        }

        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, op.Path);
        var view = PptxDoc.ResolveShape(slidePart, address);

        if (address.ParagraphIndex is not null)
        {
            PptxDoc.ResolveParagraph(view, address).Remove();
            return Units.Inv($"{view.CanonicalPath(address.SlideIndex)}/p[{address.ParagraphIndex}]");
        }

        var canonical = view.CanonicalPath(address.SlideIndex);
        PptxCharts.DeletePartFor(slidePart, view.Element); // a chart frame must not orphan its part
        if (view.Element is P.Picture mediaPicture) // a media/model picture must not orphan its data part / rels
        {
            PptxMedia.DeleteMediaPartsFor(slidePart, mediaPicture);
            PptxModels.DeleteModelPartsFor(slidePart, mediaPicture);
        }

        view.Element.Remove();
        return canonical;
    }

    private static void RemoveSlide(PresentationPart presentation, PptxAddress address)
    {
        var slides = PptxDoc.Slides(presentation);
        if (address.SlideIndex < 1 || address.SlideIndex > slides.Count)
        {
            _ = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw); // throws invalid_path with candidates
        }

        var (slideId, slidePart) = slides[address.SlideIndex - 1];
        var relId = slideId.RelationshipId?.Value;
        slideId.Remove();
        if (relId is not null)
        {
            presentation.DeletePart(relId);
        }
        else
        {
            presentation.DeletePart(slidePart);
        }
    }

    private static string ApplyMove(PresentationPart presentation, EditOp op)
    {
        var address = PptxAddress.Parse(op.Path);
        if (address.IsNotes)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "Speaker notes cannot be moved; they belong to their slide.",
                "Move the slide itself ({\"op\":\"move\",\"path\":\"/slide[3]\",\"position\":\"1\"}), or set/remove the notes text.");
        }

        if (address.IsAnimation)
        {
            return PptxAnimations.Move(presentation, address, op.Position);
        }

        if (address.IsComment)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{op.Path}' cannot be moved.",
                "Comments are anchored by x/y; reorder animations with move (before/after another animation).");
        }

        if (address.IsEmbed)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{op.Path}' cannot be moved.",
                "Embedded objects are anchored by x/y; remove and re-add the embed to reposition it.");
        }

        if (address.IsMedia)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{op.Path}' cannot be moved by media path.",
                "Reorder the host picture by its shape path (/slide[i]/shape[@id=N]) with a z-order position.");
        }

        if (address.IsModel3D)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{op.Path}' cannot be moved by model3d path.",
                "Reorder the host picture by its shape path (/slide[i]/shape[@id=N]) with a z-order position.");
        }

        if (address.IsZoom)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{op.Path}' cannot be moved by zoom path.",
                "Reorder the host graphic frame by its shape path (/slide[i]/shape[@id=N]) with a z-order position.");
        }

        if (address.IsPresentation || address.IsSection || address.IsMaster || address.IsFonts)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{op.Path}' cannot be moved.",
                "Move slides (/slide[i]) or reorder animations; sections follow slide order, " +
                "masters/layouts have no order to change, and embedded fonts are unordered.");
        }

        if (address.IsTable && (address.TableRowIndex is not null || address.TableCellIndex is not null))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{op.Path}' cannot be moved; rows and cells stay in grid order.",
                "Add/remove rows to restructure, or move the whole table: " +
                "{\"op\":\"move\",\"path\":\"" + address.CanonicalTablePath + "\",\"position\":\"front\"}.");
        }

        if (address.IsGroup)
        {
            return MoveGroupZOrder(presentation, address, op.Position);
        }

        if (address.HasShape || address.IsChart || address.IsTable)
        {
            return MoveShapeZOrder(presentation, address, op.Position);
        }

        var slideIdList = presentation.Presentation?.SlideIdList
            ?? throw CorruptPresentation("p:sldIdLst is missing");
        var ids = slideIdList.Elements<P.SlideId>().ToList();
        var count = ids.Count;

        if (address.SlideIndex < 1 || address.SlideIndex > count)
        {
            _ = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        }

        var target = ParseMoveTarget(op.Position, address.SlideIndex, count);
        var moving = ids[address.SlideIndex - 1];
        moving.Remove();
        slideIdList.InsertAt(moving, target - 1);
        return Units.Inv($"/slide[{target}]");
    }

    /// <summary>
    /// Z-order move: reorders the shape (or chart frame) among the slide's
    /// spTree drawing children. Paint order is document order, so "front" is
    /// last and "back" is first.
    /// </summary>
    private static string MoveShapeZOrder(PresentationPart presentation, PptxAddress address, string? position)
    {
        if (address.ParagraphIndex is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "Paragraphs cannot be moved; move targets a shape or a slide.",
                "Reorder text by setting paragraph texts, or move the whole shape: " +
                "{\"op\":\"move\",\"path\":\"/slide[1]/shape[@id=5]\",\"position\":\"front\"}.");
        }

        var token = position?.Trim().ToLowerInvariant();
        var readingOrderTarget = ParseReadingOrder(position);
        if (readingOrderTarget is null && (token is null || !ZOrderPositions.Contains(token, StringComparer.Ordinal)))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown z-order position '{position ?? "(none)"}'.",
                "Use front (paint last/topmost), back (paint first/bottom), forward, backward, " +
                "or \"readingOrder N\" to set the 1-based narration/tab order.",
                candidates: ZOrderPositions);
        }

        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var view = address.IsChart
            ? PptxCharts.Resolve(slidePart, address).View
            : address.IsTable
                ? PptxTables.Resolve(slidePart, address).View
                : PptxDoc.ResolveShape(slidePart, address);

        var tree = PptxDoc.RequireShapeTree(slidePart);
        var siblings = PptxDoc.Shapes(tree).Select(s => s.Element).ToList();
        var element = view.Element;
        var index = siblings.FindIndex(s => ReferenceEquals(s, element));

        // "readingOrder N" places the shape at absolute 1-based document order N.
        // Document order IS the narration/tab order and the z-order (paint order),
        // so this is the single lever the auditor's reading-order fix targets.
        if (readingOrderTarget is { } target)
        {
            if (target < 1 || target > siblings.Count)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    Units.Inv($"readingOrder {target} is out of range 1..{siblings.Count} on slide {address.SlideIndex}."),
                    "Reading order is 1-based over the slide's top-level shapes; " +
                    "run 'aioffice read <file> --view structure' to see the current order.");
            }

            if (target - 1 != index)
            {
                element.Remove();
                // Re-snapshot the siblings without the moved element, then splice at target-1.
                var remaining = siblings.Where(s => !ReferenceEquals(s, element)).ToList();
                if (target - 1 >= remaining.Count)
                {
                    tree.InsertAfter(element, remaining[^1]);
                }
                else
                {
                    tree.InsertBefore(element, remaining[target - 1]);
                }
            }

            return view.CanonicalPath(address.SlideIndex);
        }

        switch (token)
        {
            case "front" when index < siblings.Count - 1:
                element.Remove();
                tree.InsertAfter(element, siblings[^1]);
                break;
            case "back" when index > 0:
                element.Remove();
                tree.InsertBefore(element, siblings[0]);
                break;
            case "forward" when index < siblings.Count - 1:
                element.Remove();
                tree.InsertAfter(element, siblings[index + 1]);
                break;
            case "backward" when index > 0:
                element.Remove();
                tree.InsertBefore(element, siblings[index - 1]);
                break;
            default:
                break; // already at the requested extreme: a no-op, not an error
        }

        return view.CanonicalPath(address.SlideIndex);
    }

    /// <summary>
    /// Z-order move for a top-level group: reorders the p:grpSp among the slide's spTree
    /// children (front/back/forward/backward). A group child cannot be reordered on its
    /// own — ungroup first.
    /// </summary>
    private static string MoveGroupZOrder(PresentationPart presentation, PptxAddress address, string? position)
    {
        if (address.HasShape)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "A group's child shapes cannot be reordered on their own.",
                "Move the whole group ({\"op\":\"move\",\"path\":\"" + address.CanonicalGroupPath + "\",\"position\":\"front\"}), " +
                "or ungroup first to reorder the shapes individually.");
        }

        var token = position?.Trim().ToLowerInvariant();
        if (token is null || !ZOrderPositions.Contains(token, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown z-order position '{position ?? "(none)"}'.",
                "Use front (paint last/topmost), back (paint first/bottom), forward or backward.",
                candidates: ZOrderPositions);
        }

        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var tree = PptxDoc.RequireShapeTree(slidePart);
        var group = PptxGroups.ResolveGroup(tree, address);
        var siblings = PptxDoc.Shapes(tree).Select(s => s.Element).ToList();
        var index = siblings.FindIndex(s => ReferenceEquals(s, group));

        switch (token)
        {
            case "front" when index < siblings.Count - 1:
                group.Remove();
                tree.InsertAfter(group, siblings[^1]);
                break;
            case "back" when index > 0:
                group.Remove();
                tree.InsertBefore(group, siblings[0]);
                break;
            case "forward" when index < siblings.Count - 1:
                group.Remove();
                tree.InsertAfter(group, siblings[index + 1]);
                break;
            case "backward" when index > 0:
                group.Remove();
                tree.InsertBefore(group, siblings[index - 1]);
                break;
            default:
                break;
        }

        return address.CanonicalGroupPath;
    }

    /// <summary>Parses a "readingOrder N" position into its 1-based target, or null when not that form.</summary>
    private static int? ParseReadingOrder(string? position)
    {
        if (position is null)
        {
            return null;
        }

        var text = position.Trim();
        if (!text.StartsWith("readingOrder", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var rest = text["readingOrder".Length..].Trim();
        if (int.TryParse(rest, NumberStyles.None, CultureInfo.InvariantCulture, out var value))
        {
            return value;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"readingOrder needs a 1-based index, got '{position}'.",
            "Use \"readingOrder 1\" (narrate first) … \"readingOrder N\"; " +
            "run 'aioffice read <file> --view structure' to see the current order.");
    }

    /// <summary>Move targets: a 1-based final index ("3"), or "before:/slide[k]" / "after:/slide[k]".</summary>
    private static int ParseMoveTarget(string? position, int fromIndex, int count)
    {
        const string usage =
            "Use a 1-based destination index (\"1\"), or \"before:/slide[k]\" / \"after:/slide[k]\".";

        if (string.IsNullOrWhiteSpace(position))
        {
            throw new AiofficeException(ErrorCodes.InvalidArgs, "move requires a position.", usage);
        }

        var text = position.Trim();
        int target;
        if (int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var index))
        {
            target = index;
        }
        else if (TryParseAnchor(text, "before:", out var beforeAnchor))
        {
            RequireAnchorInRange(beforeAnchor, count, usage);
            target = fromIndex < beforeAnchor ? beforeAnchor - 1 : beforeAnchor;
        }
        else if (TryParseAnchor(text, "after:", out var afterAnchor))
        {
            RequireAnchorInRange(afterAnchor, count, usage);
            target = fromIndex < afterAnchor ? afterAnchor : afterAnchor + 1;
        }
        else
        {
            throw new AiofficeException(ErrorCodes.InvalidArgs, $"Unknown move position '{position}'.", usage);
        }

        if (target < 1 || target > count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"Move destination {target} is out of range 1..{count}."),
                usage);
        }

        return target;
    }

    private static bool TryParseAnchor(string text, string prefix, out int anchor)
    {
        anchor = 0;
        if (!text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var path = PptxAddress.Parse(text[prefix.Length..].Trim());
        if (path.HasShape || path.IsMaster)
        {
            return false;
        }

        anchor = path.SlideIndex;
        return true;
    }

    private static void RequireAnchorInRange(int anchor, int count, string usage)
    {
        if (anchor < 1 || anchor > count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"Move anchor slide {anchor} is out of range 1..{count}."),
                usage);
        }
    }

    /// <summary>Adds a styled title placeholder with explicit geometry (used by create and add slide).</summary>
    internal static P.Shape AddTitleShape(SlidePart slidePart, string title)
    {
        var tree = PptxDoc.RequireShapeTree(slidePart);
        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = PptxDoc.NextShapeId(tree), Name = "Title" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 831_850L, Y = 365_125L },
                    new A.Extents { Cx = 10_515_600L, Cy = 1_325_563L }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                BuildParagraph(title, fontSizeHundredths: 4000, bold: null, colorHex: null, align: null)));
        tree.Append(shape);
        return shape;
    }

    /// <summary>
    /// Adds a shape from props: the default is a rectangle textbox; props.shape
    /// picks a preset geometry (rect/roundRect/ellipse/triangle/diamond/arrow)
    /// or "line" (a straight connector honoring props.flip).
    /// </summary>
    private static uint AddTextBox(SlidePart slidePart, JsonObject? props, Workspace workspace)
    {
        var tree = PptxDoc.RequireShapeTree(slidePart);
        var id = AddTextBox(tree, props);

        // An image fill at add time needs the host part (where the picture embeds);
        // apply it once the shape is in the tree so the blip resolves against it. A
        // line (connector) has no area to fill, so image is a no-op there.
        if (props is not null && props.TryGetPropertyValue("image", out var imageNode))
        {
            var shape = PptxDoc.Shapes(tree)
                .Select(v => v.Element)
                .OfType<P.Shape>()
                .FirstOrDefault(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == id);
            if (shape is not null)
            {
                PptxFill.ApplyToShape(shape, PptxFill.BuildImageFill(imageNode, slidePart, workspace));
            }
        }

        return id;
    }

    internal static uint AddTextBox(P.ShapeTree tree, JsonObject? props)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            RequireKnownPropKey(key, AddShapePropKeys);
        }

        var geometry = props.TryGetPropertyValue("shape", out var shapeNode)
            ? ParseGeometryToken(shapeNode)
            : null;
        if (string.Equals(geometry, "line", StringComparison.Ordinal))
        {
            return AddLine(tree, props);
        }

        var id = PptxDoc.NextShapeId(tree);

        var x = props.TryGetPropertyValue("x", out var xNode) ? Units.ParseLengthEmu("x", xNode) : Units.CmToEmu(2.5);
        var y = props.TryGetPropertyValue("y", out var yNode) ? Units.ParseLengthEmu("y", yNode) : Units.CmToEmu(2.5);
        var w = props.TryGetPropertyValue("w", out var wNode) ? Units.ParseLengthEmu("w", wNode) : Units.CmToEmu(10);
        var h = props.TryGetPropertyValue("h", out var hNode) ? Units.ParseLengthEmu("h", hNode) : Units.CmToEmu(3);

        var fontSize = props.TryGetPropertyValue("fontSize", out var sizeNode)
            ? Units.ParseFontSizeHundredths("fontSize", sizeNode)
            : (int?)null;
        var bold = props.TryGetPropertyValue("bold", out var boldNode) ? AsBool("bold", boldNode) : (bool?)null;
        var color = props.TryGetPropertyValue("color", out var colorNode) ? Units.ParseColorHex("color", colorNode) : null;
        A.TextAlignmentTypeValues? align = null;
        if (props.TryGetPropertyValue("align", out var alignNode))
        {
            align = ParseAlign(alignNode) ?? throw InvalidAlign(alignNode);
        }
        var text = props.TryGetPropertyValue("text", out var textNode) && textNode is not null
            ? J.ScalarText(textNode)
            : string.Empty;
        var name = props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null
            ? J.ScalarText(nameNode)
            : geometry is null
                ? Units.Inv($"TextBox {id}")
                : Units.Inv($"{char.ToUpperInvariant(geometry[0])}{geometry[1..]} {id}");

        var transform = new A.Transform2D(
            new A.Offset { X = x, Y = y },
            new A.Extents { Cx = w, Cy = h });
        ApplyFlip(transform, props);

        var shapeProperties = new P.ShapeProperties(
            transform,
            new A.PresetGeometry(new A.AdjustValueList())
            {
                Preset = geometry is null ? A.ShapeTypeValues.Rectangle : GeometryPresets[geometry],
            });

        // A gradient fill at add time wins over a solid fill (gradient needs no host
        // part); an image fill at add time only works on the slide overload, which
        // applies it after the shape is in the tree (see AddTextBox(SlidePart, …)).
        if (props.TryGetPropertyValue("gradient", out var gradientNode))
        {
            shapeProperties.Append(PptxFill.BuildGradientFill(gradientNode));
        }
        else if (props.TryGetPropertyValue("fill", out var fillNode))
        {
            shapeProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = Units.ParseColorHex("fill", fillNode) }));
        }

        var body = new P.TextBody(new A.BodyProperties { Wrap = A.TextWrappingValues.Square }, new A.ListStyle());
        foreach (var line in text.Split('\n'))
        {
            body.Append(BuildParagraph(line, fontSize, bold, color, align));
        }

        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                geometry is null
                    ? new P.NonVisualShapeDrawingProperties { TextBox = true }
                    : new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            shapeProperties,
            body);

        // Autofit lives on the bodyPr, which exists now that the body is built.
        if (props.TryGetPropertyValue("autofit", out var autofitNode))
        {
            ApplyAutofit(shape, autofitNode);
        }

        tree.Append(shape);
        return id;
    }

    /// <summary>
    /// Adds a straight connector-style line spanning the x/y/w/h box. The
    /// default runs top-left to bottom-right; flip "v"/"h" mirrors it. props.fill
    /// sets the stroke color (a line has no area).
    /// </summary>
    private static uint AddLine(P.ShapeTree tree, JsonObject props)
    {
        foreach (var key in new[] { "text", "title", "fontSize", "bold", "color", "align" })
        {
            if (props.ContainsKey(key))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Prop '{key}' does not apply to a line.",
                    "Lines take x, y, w, h, fill (the stroke color), flip and name; put text in a separate shape.");
            }
        }

        var id = PptxDoc.NextShapeId(tree);

        var x = props.TryGetPropertyValue("x", out var xNode) ? Units.ParseLengthEmu("x", xNode) : Units.CmToEmu(2.5);
        var y = props.TryGetPropertyValue("y", out var yNode) ? Units.ParseLengthEmu("y", yNode) : Units.CmToEmu(2.5);
        var w = props.TryGetPropertyValue("w", out var wNode) ? Units.ParseLengthEmu("w", wNode) : Units.CmToEmu(10);
        var h = props.TryGetPropertyValue("h", out var hNode) ? Units.ParseLengthEmu("h", hNode) : 0L;
        var stroke = props.TryGetPropertyValue("fill", out var fillNode)
            ? Units.ParseColorHex("fill", fillNode)
            : "000000";
        var name = props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null
            ? J.ScalarText(nameNode)
            : Units.Inv($"Line {id}");

        var transform = new A.Transform2D(
            new A.Offset { X = x, Y = y },
            new A.Extents { Cx = w, Cy = h });
        ApplyFlip(transform, props);

        tree.Append(new P.ConnectionShape(
            new P.NonVisualConnectionShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualConnectorShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                transform,
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Line },
                new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = stroke })) { Width = LineWidthEmu })));
        return id;
    }

    /// <summary>Resolves a props.shape token to its canonical form or throws unsupported_feature.</summary>
    private static string ParseGeometryToken(JsonNode? node)
    {
        var raw = node is null ? string.Empty : J.ScalarText(node).Trim();
        foreach (var token in GeometryTokens)
        {
            if (string.Equals(token, raw, StringComparison.OrdinalIgnoreCase))
            {
                return token;
            }
        }

        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"Shape geometry '{raw}' is not supported.",
            "Supported geometries: rect, roundRect, ellipse, triangle, diamond, arrow, line. " +
            "Pick the closest one and refine it in PowerPoint.",
            candidates: GeometryTokens);
    }

    /// <summary>Applies props.flip ("v", "h" or "hv") to a transform.</summary>
    private static void ApplyFlip(A.Transform2D transform, JsonObject props)
    {
        if (!props.TryGetPropertyValue("flip", out var node))
        {
            return;
        }

        var token = (node is null ? string.Empty : J.ScalarText(node)).Trim().ToLowerInvariant();
        var (flipH, flipV) = token switch
        {
            "h" => (true, false),
            "v" => (false, true),
            "hv" or "vh" => (true, true),
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a valid flip value: {node?.ToJsonString() ?? "null"}",
                "Use \"h\" (mirror horizontally), \"v\" (mirror vertically) or \"hv\".",
                candidates: ["h", "v", "hv"]),
        };

        if (flipH)
        {
            transform.HorizontalFlip = true;
        }

        if (flipV)
        {
            transform.VerticalFlip = true;
        }
    }

    internal static void SetShapeProps(ShapeView view, JsonObject props, OpenXmlPartContainer? host = null, Workspace? workspace = null)
    {
        long? x = null, y = null, w = null, h = null;

        foreach (var (key, value) in props)
        {
            switch (RequireKnownPropKey(key, ShapePropKeys))
            {
                case "gradient" when view.Element is P.Shape gradientShape:
                    PptxFill.ApplyToShape(gradientShape, PptxFill.BuildGradientFill(value));
                    break;
                case "image" when view.Element is P.Shape imageShape:
                    PptxFill.ApplyToShape(imageShape, BuildShapeImageFill(value, host, workspace));
                    break;
                case "text" or "title" when view.Element is P.Shape shape:
                    ReplaceText(shape, value is null ? string.Empty : J.ScalarText(value));
                    break;
                case "x":
                    x = Units.ParseLengthEmu(key, value);
                    break;
                case "y":
                    y = Units.ParseLengthEmu(key, value);
                    break;
                case "w":
                    w = Units.ParseLengthEmu(key, value);
                    break;
                case "h":
                    h = Units.ParseLengthEmu(key, value);
                    break;
                case "fill" when view.Element is P.Shape shape:
                    SetFill(shape, Units.ParseColorHex(key, value));
                    break;
                case "fill" when view.Element is P.ConnectionShape connector:
                    // A line has no area; its visible color is the stroke.
                    SetLineColor(connector, Units.ParseColorHex(key, value));
                    break;
                case "fontSize" when view.Element is P.Shape shape:
                    ApplyRunProps(shape, rPr => rPr.FontSize = Units.ParseFontSizeHundredths(key, value));
                    break;
                case "bold" when view.Element is P.Shape shape:
                    ApplyRunProps(shape, rPr => rPr.Bold = AsBool(key, value));
                    break;
                case "color" when view.Element is P.Shape shape:
                    var hex = Units.ParseColorHex(key, value);
                    ApplyRunProps(shape, rPr => SetRunColor(rPr, hex));
                    break;
                case "align" when view.Element is P.Shape shape:
                    var alignment = ParseAlign(value) ?? throw InvalidAlign(value);
                    foreach (var paragraph in shape.TextBody?.Elements<A.Paragraph>() ?? [])
                    {
                        SetAlignment(paragraph, alignment);
                    }

                    break;
                case "name":
                    var nonVisual = PptxDoc.NonVisualProps(view.Element) ?? throw CorruptPresentation("shape has no p:cNvPr");
                    nonVisual.Name = value is null ? string.Empty : J.ScalarText(value);
                    break;
                case "altText":
                    SetAltText(view, value);
                    break;
                case "altTitle":
                    SetAltTitle(view, value);
                    break;
                case "shadow":
                    PptxEffects.SetShadow(view, value);
                    break;
                case "glow":
                    PptxEffects.SetGlow(view, value);
                    break;
                case "reflection":
                    PptxEffects.SetReflection(view, value);
                    break;
                case "outline":
                    PptxEffects.SetOutline(view, value);
                    break;
                case "autofit" when view.Element is P.Shape autofitShape:
                    ApplyAutofit(autofitShape, value);
                    break;
                default:
                    throw new AiofficeException(
                        ErrorCodes.UnsupportedFeature,
                        $"Prop '{key}' does not apply to a '{view.Kind}'.",
                        "Pictures, charts, lines and groups take x, y, w, h, name, altText, altTitle, shadow, glow, " +
                        "reflection and outline (lines also fill for the stroke color); text and styling props " +
                        "(including autofit) target text shapes.");
            }
        }

        if (x is not null || y is not null || w is not null || h is not null)
        {
            SetGeometry(view, x, y, w, h);
        }
    }

    /// <summary>
    /// Sets the accessibility description (a:cNvPr/@descr) of any shape; an empty
    /// value clears it. This is the alt text screen readers announce.
    /// </summary>
    private static void SetAltText(ShapeView view, JsonNode? value)
    {
        var nonVisual = PptxDoc.NonVisualProps(view.Element) ?? throw CorruptPresentation("shape has no p:cNvPr");
        var text = value is null ? string.Empty : J.ScalarText(value);
        nonVisual.Description = text.Length == 0 ? null : text;
    }

    /// <summary>Sets the accessibility title (a:cNvPr/@title) of any shape; an empty value clears it.</summary>
    private static void SetAltTitle(ShapeView view, JsonNode? value)
    {
        var nonVisual = PptxDoc.NonVisualProps(view.Element) ?? throw CorruptPresentation("shape has no p:cNvPr");
        var text = value is null ? string.Empty : J.ScalarText(value);
        nonVisual.Title = text.Length == 0 ? null : text;
    }

    private static void SetParagraphProps(ShapeView view, PptxAddress address, JsonObject props)
    {
        if (view.Element is not P.Shape)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{view.Kind}' shapes have no addressable paragraphs.",
                "Paragraph sets target text shapes; check the path with 'aioffice get'.");
        }

        var paragraph = PptxDoc.ResolveParagraph(view, address);
        foreach (var (key, value) in props)
        {
            switch (RequireKnownPropKey(key, ShapePropKeys))
            {
                case "text" or "title":
                    ReplaceParagraphText(paragraph, value is null ? string.Empty : J.ScalarText(value));
                    break;
                case "fontSize":
                    ApplyParagraphRunProps(paragraph, rPr => rPr.FontSize = Units.ParseFontSizeHundredths(key, value));
                    break;
                case "bold":
                    ApplyParagraphRunProps(paragraph, rPr => rPr.Bold = AsBool(key, value));
                    break;
                case "color":
                    var hex = Units.ParseColorHex(key, value);
                    ApplyParagraphRunProps(paragraph, rPr => SetRunColor(rPr, hex));
                    break;
                case "align":
                    SetAlignment(paragraph, ParseAlign(value) ?? throw InvalidAlign(value));
                    break;
                default:
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Prop '{key}' does not apply to a paragraph.",
                        "Paragraph props: text, fontSize, bold, color, align. Geometry and fill belong to the shape.",
                        candidates: ["text", "fontSize", "bold", "color", "align"]);
            }
        }
    }

    internal static A.Paragraph BuildParagraph(string text, int? fontSizeHundredths, bool? bold, string? colorHex, A.TextAlignmentTypeValues? align)
    {
        var paragraph = new A.Paragraph();
        if (align is { } alignment)
        {
            paragraph.Append(new A.ParagraphProperties { Alignment = alignment });
        }

        var runProperties = new A.RunProperties { Language = "en-US" };
        if (fontSizeHundredths is { } size)
        {
            runProperties.FontSize = size;
        }

        if (bold is { } isBold)
        {
            runProperties.Bold = isBold;
        }

        if (colorHex is { } hex)
        {
            SetRunColor(runProperties, hex);
        }

        paragraph.Append(new A.Run(runProperties, new A.Text(text)));
        return paragraph;
    }

    /// <summary>Replaces the whole text body, keeping the first run's formatting as the prototype.</summary>
    internal static void ReplaceText(P.Shape shape, string text)
    {
        var body = shape.TextBody;
        if (body is null)
        {
            body = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
            shape.Append(body);
        }

        var runPrototype = body.Descendants<A.RunProperties>().FirstOrDefault();
        var paragraphPrototype = body.Elements<A.Paragraph>().FirstOrDefault()?.ParagraphProperties;

        foreach (var paragraph in body.Elements<A.Paragraph>().ToList())
        {
            paragraph.Remove();
        }

        foreach (var line in text.Split('\n'))
        {
            var paragraph = new A.Paragraph();
            if (paragraphPrototype is not null)
            {
                paragraph.Append((A.ParagraphProperties)paragraphPrototype.CloneNode(true));
            }

            var run = new A.Run();
            if (runPrototype is not null)
            {
                run.Append((A.RunProperties)runPrototype.CloneNode(true));
            }

            run.Append(new A.Text(line));
            paragraph.Append(run);
            body.Append(paragraph);
        }
    }

    /// <summary>
    /// Sets the text-autofit behaviour on a shape's a:bodyPr. A bodyPr holds exactly
    /// one autofit child, so this replaces any existing one. Accepts a bare mode token
    /// ("shrink" -> a:normAutofit, "resize" -> a:spAutoFit, "none" -> a:noAutofit) or,
    /// for "shrink", an object {mode:"shrink", fontScale:90, lineSpaceReduction:10}
    /// that writes the explicit a:normAutofit @fontScale/@lnSpcReduction percentages
    /// (90 -> "90000"); a bare "shrink" leaves them off so PowerPoint computes the
    /// scale when the deck opens.
    /// </summary>
    internal static void ApplyAutofit(P.Shape shape, JsonNode? value)
    {
        var body = shape.TextBody;
        if (body is null)
        {
            body = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
            shape.Append(body);
        }

        var bodyPr = body.GetFirstChild<A.BodyProperties>();
        if (bodyPr is null)
        {
            bodyPr = new A.BodyProperties();
            body.InsertAt(bodyPr, 0);
        }

        var autofit = BuildAutofit(value);

        // A bodyPr has at most one autofit; drop any existing one before inserting.
        foreach (var existing in bodyPr.Elements<OpenXmlElement>()
                     .Where(e => e is A.NormalAutoFit or A.ShapeAutoFit or A.NoAutoFit).ToList())
        {
            existing.Remove();
        }

        // In a:bodyPr the autofit child follows a:prstTxWarp and precedes the 3D /
        // extLst children; insert right after prstTxWarp when present, else first.
        if (bodyPr.PresetTextWarp is { } warp)
        {
            bodyPr.InsertAfter<OpenXmlElement>(autofit, warp);
        }
        else
        {
            bodyPr.InsertAt(autofit, 0);
        }
    }

    /// <summary>Builds the autofit child element for an autofit prop value.</summary>
    private static OpenXmlElement BuildAutofit(JsonNode? value)
    {
        if (value is JsonObject obj)
        {
            var mode = J.Str(obj, "mode")?.Trim().ToLowerInvariant();
            if (mode is not null && !string.Equals(mode, "shrink", StringComparison.Ordinal))
            {
                // resize/none carry no parameters, so route the object's mode through
                // the same scalar path (it just yields a parameterless element).
                return BuildAutofitFromMode(mode);
            }

            // Object form configures the shrink (normAutofit) percentages.
            var normAutofit = new A.NormalAutoFit();
            if (TryAutofitPercent(obj, "fontScale", out var fontScale))
            {
                normAutofit.FontScale = fontScale;
            }

            if (TryAutofitPercent(obj, "lineSpaceReduction", out var lineSpaceReduction))
            {
                normAutofit.LineSpaceReduction = lineSpaceReduction;
            }

            return normAutofit;
        }

        return BuildAutofitFromMode(value is null ? null : J.ScalarText(value).Trim().ToLowerInvariant());
    }

    private static OpenXmlElement BuildAutofitFromMode(string? mode) => mode switch
    {
        "shrink" => new A.NormalAutoFit(),
        "resize" => new A.ShapeAutoFit(),
        "none" => new A.NoAutoFit(),
        _ => throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Unknown autofit mode: {mode ?? "null"}.",
            "Use \"shrink\" (normAutofit, shrink text to fit), \"resize\" (spAutoFit, grow the shape) " +
            "or \"none\" (noAutofit); \"shrink\" also accepts {mode:\"shrink\", fontScale, lineSpaceReduction}."),
    };

    /// <summary>
    /// Reads an autofit percentage field (e.g. fontScale: 90) and converts it to the
    /// OOXML thousandths-of-a-percent the bodyPr stores (90 -> 90000). Percent inputs
    /// are 0..100; bare-thousandths inputs are rejected as out of range.
    /// </summary>
    private static bool TryAutofitPercent(JsonObject obj, string key, out int thousandths)
    {
        thousandths = 0;
        if (!obj.TryGetPropertyValue(key, out var node) || node is not JsonValue value || !Units.TryNumber(value, out var percent))
        {
            return false;
        }

        if (percent is < 0 or > 100)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"autofit '{key}' is out of range: {percent.ToString(CultureInfo.InvariantCulture)}.",
                "Use a percentage between 0 and 100 (e.g. 90 means 90%).");
        }

        thousandths = (int)Math.Round(percent * 1000);
        return true;
    }

    private static void ReplaceParagraphText(A.Paragraph paragraph, string text)
    {
        var runPrototype = paragraph.Elements<A.Run>().FirstOrDefault()?.RunProperties;
        var run = new A.Run();
        if (runPrototype is not null)
        {
            run.Append((A.RunProperties)runPrototype.CloneNode(true));
        }

        run.Append(new A.Text(text));

        foreach (var child in paragraph.ChildElements.Where(c => c is A.Run or A.Break or A.Field).ToList())
        {
            child.Remove();
        }

        paragraph.Append(run);
    }

    private static void SetGeometry(ShapeView view, long? x, long? y, long? w, long? h)
    {
        // Graphic frames (charts, tables) carry p:xfrm directly, not inside spPr.
        if (view.Element is P.GraphicFrame frame)
        {
            frame.Transform ??= new P.Transform(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = Units.CmToEmu(10), Cy = Units.CmToEmu(3) });
            ApplyOffsetExtents(
                frame.Transform.Offset ??= new A.Offset { X = 0L, Y = 0L },
                frame.Transform.Extents ??= new A.Extents { Cx = Units.CmToEmu(10), Cy = Units.CmToEmu(3) },
                x, y, w, h);
            return;
        }

        var properties = view.Element switch
        {
            P.Shape s => s.ShapeProperties ?? InsertShapeProperties(s),
            P.Picture p => p.ShapeProperties ??= new P.ShapeProperties(),
            P.ConnectionShape c => c.ShapeProperties ??= new P.ShapeProperties(),
            _ => throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Geometry sets are not supported on a '{view.Kind}'.",
                "Move/resize shapes, pictures, lines and charts; ungroup grouped content in PowerPoint first."),
        };

        var transform = properties.Transform2D;
        if (transform is null)
        {
            transform = new A.Transform2D(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = Units.CmToEmu(10), Cy = Units.CmToEmu(3) });
            properties.InsertAt(transform, 0);
        }

        ApplyOffsetExtents(
            transform.Offset ??= new A.Offset { X = 0L, Y = 0L },
            transform.Extents ??= new A.Extents { Cx = Units.CmToEmu(10), Cy = Units.CmToEmu(3) },
            x, y, w, h);
    }

    private static P.ShapeProperties InsertShapeProperties(P.Shape shape)
    {
        var properties = new P.ShapeProperties();
        shape.InsertAfter(properties, shape.NonVisualShapeProperties);
        return properties;
    }

    private static void ApplyOffsetExtents(A.Offset offset, A.Extents extents, long? x, long? y, long? w, long? h)
    {
        if (x is not null)
        {
            offset.X = x;
        }

        if (y is not null)
        {
            offset.Y = y;
        }

        if (w is not null)
        {
            extents.Cx = w;
        }

        if (h is not null)
        {
            extents.Cy = h;
        }
    }

    /// <summary>Replaces a connector's outline color (the visible color of a line shape).</summary>
    private static void SetLineColor(P.ConnectionShape connector, string hex)
    {
        var properties = connector.ShapeProperties ??= new P.ShapeProperties();
        var outline = properties.GetFirstChild<A.Outline>();
        if (outline is null)
        {
            outline = new A.Outline { Width = LineWidthEmu };
            properties.Append(outline);
        }

        foreach (var fill in outline.ChildElements.Where(c => c is A.SolidFill or A.NoFill or A.GradientFill).ToList())
        {
            fill.Remove();
        }

        outline.InsertAt(new A.SolidFill(new A.RgbColorModelHex { Val = hex }), 0);
    }

    private static void SetFill(P.Shape shape, string hex)
    {
        var properties = shape.ShapeProperties;
        if (properties is null)
        {
            properties = new P.ShapeProperties();
            shape.InsertAfter(properties, shape.NonVisualShapeProperties);
        }

        foreach (var fill in properties.ChildElements
            .Where(c => c is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill)
            .ToList())
        {
            fill.Remove();
        }

        var solidFill = new A.SolidFill(new A.RgbColorModelHex { Val = hex });
        OpenXmlElement? anchor = (OpenXmlElement?)properties.GetFirstChild<A.PresetGeometry>()
            ?? (OpenXmlElement?)properties.GetFirstChild<A.CustomGeometry>()
            ?? properties.Transform2D;
        if (anchor is not null)
        {
            properties.InsertAfter(solidFill, anchor);
        }
        else
        {
            properties.InsertAt(solidFill, 0);
        }
    }

    /// <summary>
    /// Builds an image (blip) fill for a shape, requiring the host part (where the
    /// picture embeds) and the workspace (which sandbox-resolves the source path).
    /// </summary>
    private static A.BlipFill BuildShapeImageFill(JsonNode? value, OpenXmlPartContainer? host, Workspace? workspace)
    {
        if (host is null || workspace is null)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "An image fill cannot be applied here (no host part to embed the picture).",
                "Set the image fill on a slide shape, e.g. " +
                "{\"op\":\"set\",\"path\":\"/slide[1]/shape[@id=5]\",\"props\":{\"image\":\"banner.jpg\"}}.");
        }

        return PptxFill.BuildImageFill(value, host, workspace);
    }

    private static void ApplyRunProps(P.Shape shape, Action<A.RunProperties> mutate)
    {
        foreach (var paragraph in shape.TextBody?.Elements<A.Paragraph>() ?? [])
        {
            ApplyParagraphRunProps(paragraph, mutate);
        }
    }

    private static void ApplyParagraphRunProps(A.Paragraph paragraph, Action<A.RunProperties> mutate)
    {
        foreach (var run in paragraph.Elements<A.Run>())
        {
            var runProperties = run.RunProperties;
            if (runProperties is null)
            {
                runProperties = new A.RunProperties { Language = "en-US" };
                run.InsertAt(runProperties, 0);
            }

            mutate(runProperties);
        }
    }

    private static void SetRunColor(A.RunProperties runProperties, string hex)
    {
        foreach (var fill in runProperties.ChildElements.Where(c => c is A.SolidFill or A.NoFill or A.GradientFill).ToList())
        {
            fill.Remove();
        }

        runProperties.InsertAt(new A.SolidFill(new A.RgbColorModelHex { Val = hex }), 0);
    }

    private static void SetAlignment(A.Paragraph paragraph, A.TextAlignmentTypeValues alignment)
    {
        var properties = paragraph.ParagraphProperties;
        if (properties is null)
        {
            properties = new A.ParagraphProperties();
            paragraph.InsertAt(properties, 0);
        }

        properties.Alignment = alignment;
    }

    internal static A.TextAlignmentTypeValues? ParseAlign(JsonNode? node)
    {
        var text = node is null ? null : J.ScalarText(node).Trim().ToLowerInvariant();
        return text switch
        {
            "left" => A.TextAlignmentTypeValues.Left,
            "center" => A.TextAlignmentTypeValues.Center,
            "right" => A.TextAlignmentTypeValues.Right,
            "justify" => A.TextAlignmentTypeValues.Justified,
            _ => null,
        };
    }

    private static AiofficeException InvalidAlign(JsonNode? node) => new(
        ErrorCodes.InvalidArgs,
        $"Not a valid align value: {node?.ToJsonString() ?? "null"}",
        "Use left, center, right or justify.",
        candidates: ["left", "center", "right", "justify"]);

    private static bool AsBool(string key, JsonNode? node)
    {
        if (node is JsonValue value)
        {
            if (value.TryGetValue<bool>(out var flag))
            {
                return flag;
            }

            if (value.TryGetValue<string>(out var text) && bool.TryParse(text, out var parsed))
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Property '{key}' is not a boolean: {node?.ToJsonString() ?? "null"}",
            "Use true or false.");
    }

    private static string RequireKnownPropKey(string key, IReadOnlyList<string> allowed)
    {
        if (allowed.Contains(key, StringComparer.Ordinal))
        {
            return key;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Unknown shape prop '{key}'.",
            "Run 'aioffice help properties' for the per-type prop list.",
            candidates: allowed);
    }

    private static AiofficeException CorruptPresentation(string detail) => new(
        ErrorCodes.FormatCorrupt,
        $"The presentation is malformed: {detail}.",
        "Re-export the file from PowerPoint/Keynote, or restore a snapshot.");
}
