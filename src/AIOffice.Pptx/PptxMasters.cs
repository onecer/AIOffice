using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// M6 master/layout editing (the read-only debt M1 deferred):
/// <list type="bullet">
/// <item>set /master[m] {background, accent1..accent6} — slide-master background and theme color scheme.</item>
/// <item>set /master[m]/layout[l] {background} — per-layout background.</item>
/// <item>set/add/remove on /master[m]/shape[i] and /master[m]/layout[l]/shape[i] — reuses the slide shape ops.</item>
/// <item>add a layout (clone-from an existing one) and remove a non-referenced layout.</item>
/// </list>
/// Every edit keeps the package validator-clean and leaves existing slides valid.
/// </summary>
internal static class PptxMasters
{
    /// <summary>The theme accent slots that set /master[m] can recolor.</summary>
    private static readonly IReadOnlyList<string> AccentKeys =
        ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];

    // ---- set ---------------------------------------------------------------

    /// <summary>Routes a set op on a master/layout path to the right editor; returns the canonical target.</summary>
    public static string Set(PresentationPart presentation, PptxAddress address, JsonObject props)
    {
        if (address.HasShape)
        {
            return SetShape(presentation, address, props);
        }

        return address.LayoutIndex is null
            ? SetMaster(presentation, address, props)
            : SetLayout(presentation, address, props);
    }

    /// <summary>set /master[m]: background plus theme accent colors (the color scheme in the theme part).</summary>
    private static string SetMaster(PresentationPart presentation, PptxAddress address, JsonObject props)
    {
        var masterPart = PptxDoc.ResolveMaster(presentation, address.MasterIndex, address.Raw);
        var accents = new List<(int Slot, string Hex)>();

        foreach (var (key, value) in props)
        {
            if (string.Equals(key, "background", StringComparison.Ordinal))
            {
                var slideData = masterPart.SlideMaster?.CommonSlideData ?? throw Corrupt("the master has no p:cSld");
                PptxEditor.SetBackground(slideData, value);
                continue;
            }

            var lower = key.ToLowerInvariant();
            var slot = -1;
            for (var i = 0; i < AccentKeys.Count; i++)
            {
                if (string.Equals(AccentKeys[i], lower, StringComparison.Ordinal))
                {
                    slot = i;
                    break;
                }
            }

            if (slot >= 0)
            {
                accents.Add((slot + 1, Units.ParseColorHex(key, value)));
                continue;
            }

            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Prop '{key}' does not apply to a master.",
                "Master props: background, accent1..accent6 (theme colors). " +
                "Edit master shapes via /master[m]/shape[i]; layout props target /master[m]/layout[l].",
                candidates: [.. new[] { "background" }.Concat(AccentKeys)]);
        }

        if (accents.Count > 0)
        {
            SetThemeAccents(masterPart, accents);
        }

        return address.CanonicalMasterPath;
    }

    /// <summary>set /master[m]/layout[l]: the layout's background (shapes target the shape path).</summary>
    private static string SetLayout(PresentationPart presentation, PptxAddress address, JsonObject props)
    {
        var masterPart = PptxDoc.ResolveMaster(presentation, address.MasterIndex, address.Raw);
        var layoutPart = PptxDoc.ResolveLayout(masterPart, address.MasterIndex, address.LayoutIndex!.Value, address.Raw);

        foreach (var (key, value) in props)
        {
            if (string.Equals(key, "background", StringComparison.Ordinal))
            {
                var slideData = layoutPart.SlideLayout?.CommonSlideData ?? throw Corrupt("the layout has no p:cSld");
                PptxEditor.SetBackground(slideData, value);
                continue;
            }

            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Prop '{key}' does not apply to a layout.",
                "Layout props: background. Edit layout shapes via /master[m]/layout[l]/shape[i]; " +
                "theme colors target the master (/master[m]).",
                candidates: ["background"]);
        }

        return address.CanonicalLayoutPath;
    }

    /// <summary>set on a master/layout shape: delegates to the shared slide shape-prop setter.</summary>
    private static string SetShape(PresentationPart presentation, PptxAddress address, JsonObject props)
    {
        var (tree, containerPath, label) = ResolveShapeTree(presentation, address);
        var view = PptxDoc.ResolveShape(PptxDoc.Shapes(tree), address, containerPath, label);
        PptxEditor.SetShapeProps(view, props);
        return view.CanonicalPathIn(containerPath);
    }

    /// <summary>Recolors the master theme part's accent slots (a:accent1..a:accent6 in the color scheme).</summary>
    private static void SetThemeAccents(SlideMasterPart masterPart, List<(int Slot, string Hex)> accents)
    {
        var scheme = masterPart.ThemePart?.Theme?.ThemeElements?.ColorScheme ?? throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            "The master has no theme color scheme to recolor.",
            "Theme accents are stored in the master's theme part; re-export the deck from PowerPoint to add one, " +
            "or set shape fills directly instead.");

        foreach (var (slot, hex) in accents)
        {
            OpenXmlElement accent = slot switch
            {
                1 => scheme.Accent1Color ??= new A.Accent1Color(),
                2 => scheme.Accent2Color ??= new A.Accent2Color(),
                3 => scheme.Accent3Color ??= new A.Accent3Color(),
                4 => scheme.Accent4Color ??= new A.Accent4Color(),
                5 => scheme.Accent5Color ??= new A.Accent5Color(),
                _ => scheme.Accent6Color ??= new A.Accent6Color(),
            };

            // Replace whatever color child it had (srgbClr/sysClr/...) with an explicit RGB hex.
            accent.RemoveAllChildren();
            accent.AppendChild(new A.RgbColorModelHex { Val = hex });
        }
    }

    // ---- add ---------------------------------------------------------------

    /// <summary>add layout (clone an existing one) or add shape, on a master/layout path.</summary>
    public static string Add(PresentationPart presentation, PptxAddress address, EditOp op)
    {
        var type = op.Type?.Trim().ToLowerInvariant();
        return type switch
        {
            "layout" => AddLayout(presentation, address, op.Props),
            "shape" or "textbox" => AddShape(presentation, address, op.Props),
            "image" or "picture" or "chart" or "table" => throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Adding a '{op.Type}' to a master/layout is not supported.",
                "Add pictures/charts/tables to a slide instead; masters/layouts take shape (textbox/geometry) and layout."),
            "slide" => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add slide targets the slide list, not a master.",
                "Use a /slide[i] path, e.g. {\"op\":\"add\",\"path\":\"/slide[2]\",\"type\":\"slide\",\"props\":{\"layout\":2}}."),
            _ => throw new AiofficeException(
                op.Type is null ? ErrorCodes.InvalidArgs : ErrorCodes.UnsupportedFeature,
                op.Type is null ? "add on a master/layout path requires a type." : $"Cannot add '{op.Type}' to a master/layout.",
                "On a master you can add a layout (clone-from an existing one) or a shape; " +
                "on a layout you can add a shape.",
                candidates: ["layout", "shape"]),
        };
    }

    /// <summary>
    /// add /master[m] {type:layout}: clones an existing layout (props.basedOn, 1-based; default the first)
    /// into a new layout part, renames it (props.name) and appends it to the master's layout list.
    /// </summary>
    private static string AddLayout(PresentationPart presentation, PptxAddress address, JsonObject? props)
    {
        if (address.LayoutIndex is not null || address.HasShape)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add layout targets a master, not '{address.Raw}'.",
                "Use the master path, e.g. {\"op\":\"add\",\"path\":\"/master[1]\",\"type\":\"layout\"," +
                "\"props\":{\"name\":\"My Layout\",\"basedOn\":2}}.");
        }

        props ??= [];
        foreach (var (key, _) in props)
        {
            if (key is not ("name" or "basedOn"))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown add-layout prop '{key}'.",
                    "add layout takes name (the layout's display name) and basedOn (a 1-based layout index to clone).",
                    candidates: ["name", "basedOn"]);
            }
        }

        var masterPart = PptxDoc.ResolveMaster(presentation, address.MasterIndex, address.Raw);
        var layouts = PptxDoc.Layouts(masterPart);
        if (layouts.Count == 0)
        {
            throw Corrupt("the master has no layouts to clone from");
        }

        var source = ResolveBasedOn(masterPart, address.MasterIndex, layouts, props);
        var clonedXml = (P.SlideLayout)source.SlideLayout!.CloneNode(true);

        if (props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null)
        {
            (clonedXml.CommonSlideData ??= new P.CommonSlideData(PptxFactory.EmptyShapeTree())).Name =
                J.ScalarText(nameNode);
        }

        var newLayout = masterPart.AddNewPart<SlideLayoutPart>();
        newLayout.SlideLayout = clonedXml;
        newLayout.AddPart(masterPart); // a layout must reference its master

        var idList = masterPart.SlideMaster!.SlideLayoutIdList
            ?? masterPart.SlideMaster.AppendChild(new P.SlideLayoutIdList());
        idList.Append(new P.SlideLayoutId
        {
            Id = NextLayoutId(idList),
            RelationshipId = masterPart.GetIdOfPart(newLayout),
        });

        return Units.Inv($"/master[{address.MasterIndex}]/layout[{layouts.Count + 1}]");
    }

    /// <summary>The layout add shapes/clones from: props.basedOn (1-based) or the master's first layout.</summary>
    private static SlideLayoutPart ResolveBasedOn(
        SlideMasterPart masterPart, int masterIndex, List<(int Index, SlideLayoutPart Part)> layouts, JsonObject props)
    {
        if (!props.TryGetPropertyValue("basedOn", out var node) || node is null)
        {
            return layouts[0].Part;
        }

        double number = 0;
        var numeric = node is JsonValue value &&
            (Units.TryNumber(value, out number) ||
             (value.TryGetValue<string>(out var raw) &&
              double.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out number)));
        if (!numeric || number != Math.Floor(number) || number < 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.basedOn is not a valid layout index: {node.ToJsonString()}",
                "Use a 1-based integer index into the master's layouts; run 'aioffice read <file> --view structure' to list them.");
        }

        var index = (int)number;
        if (index > layouts.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"props.basedOn is {index} but master {masterIndex} has only {layouts.Count} layout(s)."),
                "Run 'aioffice read <file> --view structure' to list the master's layouts.",
                candidates: [.. layouts.Take(10).Select(l => Units.Inv($"/master[{masterIndex}]/layout[{l.Index}]"))]);
        }

        return layouts[index - 1].Part;
    }

    /// <summary>add shape (or textbox/geometry/line) onto a master or layout shape tree.</summary>
    private static string AddShape(PresentationPart presentation, PptxAddress address, JsonObject? props)
    {
        if (address.HasShape)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add shape targets a master/layout, not '{address.Raw}'.",
                "Use the container path, e.g. {\"op\":\"add\",\"path\":\"/master[1]/layout[2]\",\"type\":\"shape\"}.");
        }

        var (tree, containerPath, _) = ResolveShapeTree(presentation, address);
        var id = PptxEditor.AddTextBox(tree, props);
        return Units.Inv($"{containerPath}/shape[@id={id}]");
    }

    // ---- remove ------------------------------------------------------------

    /// <summary>remove a layout (when no slide uses it) or a master/layout shape.</summary>
    public static string Remove(PresentationPart presentation, PptxAddress address)
    {
        if (address.HasShape)
        {
            return RemoveShape(presentation, address);
        }

        if (address.LayoutIndex is not null)
        {
            return RemoveLayout(presentation, address);
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "A slide master cannot be removed (every deck needs at least one).",
            "Remove a layout (/master[m]/layout[l]) or a shape (/master[m]/shape[i]) instead.");
    }

    /// <summary>remove /master[m]/layout[l]: refuses when any slide uses it, naming those slides.</summary>
    private static string RemoveLayout(PresentationPart presentation, PptxAddress address)
    {
        var masterPart = PptxDoc.ResolveMaster(presentation, address.MasterIndex, address.Raw);
        var layouts = PptxDoc.Layouts(masterPart);
        var layoutIndex = address.LayoutIndex!.Value;
        if (layoutIndex < 1 || layoutIndex > layouts.Count)
        {
            _ = PptxDoc.ResolveLayout(masterPart, address.MasterIndex, layoutIndex, address.Raw); // throws with candidates
        }

        var layoutPart = layouts[layoutIndex - 1].Part;
        var usedBy = new List<int>();
        var slides = PptxDoc.Slides(presentation);
        for (var i = 0; i < slides.Count; i++)
        {
            if (slides[i].Part.SlideLayoutPart?.Uri == layoutPart.Uri)
            {
                usedBy.Add(i + 1);
            }
        }

        if (usedBy.Count > 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"Layout {layoutIndex} is used by slide(s) {string.Join(", ", usedBy)} and cannot be removed."),
                "Re-bind those slides to another layout in PowerPoint first, or remove the slides, then drop the layout.",
                candidates: [.. usedBy.Take(10).Select(s => Units.Inv($"/slide[{s}]"))]);
        }

        if (layouts.Count == 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "A master must keep at least one layout; this is its only one.",
                "Add another layout first ({\"op\":\"add\",\"path\":\"" + address.CanonicalMasterPath +
                "\",\"type\":\"layout\"}), then remove this one.");
        }

        // Drop the p:sldLayoutId entry, then the part itself.
        var idList = masterPart.SlideMaster!.SlideLayoutIdList!;
        var relId = masterPart.GetIdOfPart(layoutPart);
        idList.Elements<P.SlideLayoutId>().FirstOrDefault(e => e.RelationshipId?.Value == relId)?.Remove();
        masterPart.DeletePart(layoutPart);

        return address.CanonicalLayoutPath;
    }

    /// <summary>remove a master/layout shape from its tree.</summary>
    private static string RemoveShape(PresentationPart presentation, PptxAddress address)
    {
        var (tree, containerPath, label) = ResolveShapeTree(presentation, address);
        var view = PptxDoc.ResolveShape(PptxDoc.Shapes(tree), address, containerPath, label);
        var canonical = view.CanonicalPathIn(containerPath);
        view.Element.Remove();
        return canonical;
    }

    // ---- shared ------------------------------------------------------------

    /// <summary>Resolves the shape tree (and labels) behind a /master[m]/shape or /master[m]/layout[l]/shape path.</summary>
    private static (P.ShapeTree Tree, string ContainerPath, string Label) ResolveShapeTree(
        PresentationPart presentation, PptxAddress address)
    {
        var masterPart = PptxDoc.ResolveMaster(presentation, address.MasterIndex, address.Raw);
        if (address.LayoutIndex is { } layoutIndex)
        {
            var layoutPart = PptxDoc.ResolveLayout(masterPart, address.MasterIndex, layoutIndex, address.Raw);
            return (
                PptxDoc.RequireShapeTree(layoutPart),
                address.CanonicalLayoutPath,
                Units.Inv($"on layout {layoutIndex} of master {address.MasterIndex}"));
        }

        return (
            PptxDoc.RequireShapeTree(masterPart),
            address.CanonicalMasterPath,
            Units.Inv($"on master {address.MasterIndex}"));
    }

    private static uint NextLayoutId(P.SlideLayoutIdList idList)
    {
        // Layout ids share the slide-master id space; the spec floor is 2147483648.
        uint max = 2_147_483_647;
        foreach (var id in idList.Elements<P.SlideLayoutId>())
        {
            if (id.Id?.Value is { } value && value > max)
            {
                max = value;
            }
        }

        return max + 1;
    }

    private static AiofficeException Corrupt(string detail) => new(
        ErrorCodes.FormatCorrupt,
        $"The presentation is malformed: {detail}.",
        "Re-export the file from PowerPoint/Keynote, or restore a snapshot.");
}
