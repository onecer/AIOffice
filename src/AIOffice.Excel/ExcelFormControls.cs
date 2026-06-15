using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace AIOffice.Excel;

/// <summary>The interactive form-control kinds aioffice can author.</summary>
internal enum FormControlKind
{
    CheckBox,
    OptionButton,
    Spinner,
    ComboBox,
    ListBox,
    Button,
}

/// <summary>
/// A fully validated form-control add, captured from the in-memory workbook
/// while the edit batch is applied. The raw OpenXml write happens later, in the
/// post-save pass (<see cref="ExcelFormControls.Apply"/>), because ClosedXML
/// cannot author legacy form-control parts (VML drawing + ctrlProp +
/// worksheet controls) itself — and, measured here, preserves them
/// byte-identical across its own saves once present (exactly like charts and
/// slicers).
/// </summary>
internal sealed record FormControlAddSpec(
    string SheetName,
    FormControlKind Kind,
    string Cell,
    int AnchorColumn, // 0-based
    int AnchorRow,    // 0-based
    string? LinkedCell,
    string? Text,
    IReadOnlyList<string>? Items,
    int? Min,
    int? Max,
    int? Increment,
    string? ListFillRange);

/// <summary>A validated form-control removal (1-based per-sheet index at apply time).</summary>
internal sealed record FormControlRemoveSpec(string SheetName, int Index);

/// <summary>A form control as read back from the raw package (for get / read --view structure).</summary>
internal sealed record FormControlInfo(
    string SheetName,
    int Index,
    string Path,
    string Kind,
    string Cell,
    string? LinkedCell,
    string? Text,
    IReadOnlyList<string>? Items,
    int? Min,
    int? Max,
    int? Increment);

/// <summary>
/// Excel form controls (v1.2 interactive feature): checkbox, option button,
/// spinner, combo box, list box and button. Authored entirely on raw OpenXml as
/// LEGACY form controls — the same layout Excel writes when you drop a control
/// from the Developer ribbon, which the OpenXmlValidator (Office2019) accepts
/// with zero errors and which opens in Excel:
///
/// <list type="bullet">
/// <item>A <c>ControlPropertiesPart</c> (ctrlProp) on the worksheet holds the
/// <c>x14:formControlPr</c> with the control's object type and behavior
/// (checked state, fmlaLink, min/max/inc, list items, …).</item>
/// <item>A <c>VmlDrawingPart</c> holds the legacy shape (a <c>v:shape</c> of the
/// matching <c>ObjectType</c> with an <c>x:ClientData</c> anchored at the cell);
/// the worksheet references it via a <c>legacyDrawing</c> element.</item>
/// <item>The worksheet <c>controls</c> element carries an
/// <c>mc:AlternateContent</c> with an <c>x14:control</c> Choice (linkedCell,
/// listFillRange and the cell anchor) plus a plain <c>control</c> Fallback, both
/// pointing at the ctrlProp part by relationship id.</item>
/// </list>
///
/// Where a control kind has no <c>linkedCell</c> in Excel (a button), the prop is
/// ignored; an unsupported kind returns <c>unsupported_feature</c> naming the
/// supported set.
/// </summary>
internal static class ExcelFormControls
{
    /// <summary>
    /// Collects form-control ops during an edit batch so they can be validated up
    /// front (atomicity: a bad op aborts before any byte is written) and applied
    /// after ClosedXML has saved. Tracks projected per-sheet control counts so
    /// indices in a multi-op batch stay consistent (precedent: ChartOpBatch).
    /// </summary>
    public sealed class Batch
    {
        private readonly string _file;
        private Dictionary<string, int>? _projectedCounts;

        public Batch(string file) => _file = file;

        public List<object> Ops { get; } = [];

        public bool IsEmpty => Ops.Count == 0;

        /// <summary>Queues an add and returns the control's projected 1-based index on its sheet.</summary>
        public int Add(FormControlAddSpec spec)
        {
            var counts = ProjectedCounts();
            var next = counts.GetValueOrDefault(spec.SheetName) + 1;
            counts[spec.SheetName] = next;
            Ops.Add(spec);
            return next;
        }

        /// <summary>Queues a removal, validating the index against the projected state.</summary>
        public void Remove(string sheetName, string sheetPath, int index, int opIndex)
        {
            var counts = ProjectedCounts();
            var current = counts.GetValueOrDefault(sheetName);
            if (index > current)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidPath,
                    $"ops[{opIndex}]: no formControl[{index}] on sheet '{sheetName}' ({current} control(s) exist).",
                    current > 0
                        ? "Form-control indices are 1-based per sheet; run 'aioffice read --view structure' to list them."
                        : "This sheet has no form controls; add one with {op:add, type:formControl, path:" + sheetPath +
                          ", props:{kind:\"checkbox\", cell:\"E2\", linkedCell:\"F2\"}}.",
                    candidates: current > 0
                        ? [.. Enumerable.Range(1, current).Select(i => $"{sheetPath}/formControl[{i}]")]
                        : [sheetPath]);
            }

            counts[sheetName] = current - 1;
            Ops.Add(new FormControlRemoveSpec(sheetName, index));
        }

        private Dictionary<string, int> ProjectedCounts()
        {
            if (_projectedCounts is null)
            {
                _projectedCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                using var document = SpreadsheetDocument.Open(_file, isEditable: false);
                foreach (var info in Read(document))
                {
                    _projectedCounts[info.SheetName] = info.Index; // indices are sequential per sheet
                }
            }

            return _projectedCounts;
        }
    }

    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string X14Ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    private const string XmNs = "http://schemas.microsoft.com/office/excel/2006/main";
    private const string RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string XdrNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private const string VmlNs = "urn:schemas-microsoft-com:vml";
    private const string ONs = "urn:schemas-microsoft-com:office:office";
    private const string XNs = "urn:schemas-microsoft-com:office:excel";

    private const int MaxColumns = 16384;
    private const int MaxRows = 1048576;

    public static readonly IReadOnlyList<string> Kinds =
        ["checkbox", "optionButton", "spinner", "comboBox", "listBox", "button"];

    private static readonly IReadOnlyList<string> AddProps =
        ["kind", "cell", "linkedCell", "text", "items", "min", "max", "increment", "listFillRange"];

    // ----- op-time validation & capture -----------------------------------------

    /// <summary>
    /// Validates an <c>add formControl</c> op against the in-memory workbook and
    /// captures the spec the post-save writer needs. The op path names the host
    /// sheet (<c>/Sheet1</c>); props.cell is the anchor cell.
    /// </summary>
    public static FormControlAddSpec ParseAdd(ExcelTarget target, EditOp op, int opIndex)
    {
        if (target.Kind != ExcelTargetKind.Sheet)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add formControl targets a sheet path like /Sheet1; the anchor goes in props.cell.",
                "Use {op:add, type:formControl, path:/Sheet1, props:{kind:\"checkbox\", cell:\"E2\", linkedCell:\"F2\"}}.");
        }

        var props = op.Props;
        if (props is null || props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add formControl needs props.",
                "Pass props like {\"kind\":\"checkbox\",\"cell\":\"E2\",\"linkedCell\":\"F2\"}.");
        }

        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown formControl prop '{key}'.",
                    "Supported props: " + string.Join(", ", AddProps) + ".",
                    candidates: AddProps);
            }
        }

        var kindText = RequiredString(props, "kind", opIndex);
        var kind = kindText switch
        {
            "checkbox" => FormControlKind.CheckBox,
            "optionButton" => FormControlKind.OptionButton,
            "spinner" => FormControlKind.Spinner,
            "comboBox" => FormControlKind.ComboBox,
            "listBox" => FormControlKind.ListBox,
            "button" => FormControlKind.Button,
            _ => throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"ops[{opIndex}]: form-control kind '{kindText}' is not supported.",
                "Supported kinds: " + string.Join(", ", Kinds) + ".",
                candidates: Kinds),
        };

        var cellText = RequiredString(props, "cell", opIndex);
        var (anchorColumn, anchorRow) = ParseCell(cellText, opIndex);

        var linkedCell = OptionalString(props, "linkedCell");
        if (linkedCell is not null)
        {
            _ = ParseCell(linkedCell, opIndex); // validate format up front
        }

        var items = OptionalStrings(props, "items", opIndex);
        if (kind is FormControlKind.ComboBox or FormControlKind.ListBox)
        {
            var listFill = OptionalString(props, "listFillRange");
            if ((items is null || items.Count == 0) && listFill is null)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: a {kindText} needs items (a list) or listFillRange (a worksheet range).",
                    "Pass {items:[\"Red\",\"Green\",\"Blue\"]} or {listFillRange:\"H1:H3\"}.");
            }
        }
        else if (items is { Count: > 0 })
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'items' only applies to a comboBox or listBox, not a {kindText}.",
                "Drop the items prop for this control kind.");
        }

        int? min = null, max = null, increment = null;
        if (kind == FormControlKind.Spinner)
        {
            min = OptionalInt(props, "min", opIndex) ?? 0;
            max = OptionalInt(props, "max", opIndex) ?? 100;
            increment = OptionalInt(props, "increment", opIndex) ?? 1;
            if (min < 0 || max < 0 || increment < 1 || min >= max)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: spinner needs 0 <= min < max and increment >= 1 (got min={min}, max={max}, increment={increment}).",
                    "Use e.g. {min:0, max:10, increment:1}; the spinner stores its value in linkedCell.");
            }
        }
        else if (OptionalInt(props, "min", opIndex) is not null ||
                 OptionalInt(props, "max", opIndex) is not null ||
                 OptionalInt(props, "increment", opIndex) is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: min/max/increment only apply to a spinner, not a {kindText}.",
                "Drop those props for this control kind.");
        }

        return new FormControlAddSpec(
            SheetName: target.Sheet.Name,
            Kind: kind,
            Cell: cellText.ToUpperInvariant(),
            AnchorColumn: anchorColumn,
            AnchorRow: anchorRow,
            LinkedCell: linkedCell is null ? null : AbsRef(linkedCell),
            Text: OptionalString(props, "text"),
            Items: items,
            Min: min,
            Max: max,
            Increment: increment,
            ListFillRange: OptionalString(props, "listFillRange"));
    }

    // ----- post-save apply ------------------------------------------------------

    /// <summary>
    /// Applies queued form-control ops to the file ClosedXML just saved. All
    /// semantic validation already happened at op time, so this pass is
    /// mechanical raw-OpenXml writing.
    /// </summary>
    public static void Apply(string file, IReadOnlyList<object> ops)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        foreach (var op in ops)
        {
            switch (op)
            {
                case FormControlAddSpec add:
                    AddControl(document, add);
                    break;
                case FormControlRemoveSpec remove:
                    RemoveControl(document, remove);
                    break;
            }
        }
    }

    private static WorksheetPart SheetPartOrThrow(SpreadsheetDocument document, string sheetName)
    {
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart?.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id?.Value is { } relationshipId &&
            workbookPart!.GetPartById(relationshipId) is WorksheetPart worksheetPart)
        {
            return worksheetPart;
        }

        throw new AiofficeException(
            ErrorCodes.InternalError,
            $"Sheet '{sheetName}' disappeared between validation and the form-control write pass.",
            "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
    }

    private static void AddControl(SpreadsheetDocument document, FormControlAddSpec spec)
    {
        var worksheetPart = SheetPartOrThrow(document, spec.SheetName);
        var worksheet = worksheetPart.Worksheet ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            $"Sheet '{spec.SheetName}' has no worksheet XML.",
            "Restore a snapshot ('aioffice snapshot list') or re-export the file from its source.");

        // 1) ctrlProp part (the form-control properties).
        var ctrlPart = worksheetPart.AddNewPart<ControlPropertiesPart>();
        ctrlPart.FormControlProperties = BuildFormControlProperties(spec);
        var ctrlId = worksheetPart.GetIdOfPart(ctrlPart);

        // 2) a fresh shape id (max existing + 1, base 1025 like Excel).
        var shapeId = NextShapeId(worksheet);

        // 3) VML drawing shape, appended to the worksheet's single VML drawing part.
        var vmlPart = EnsureVmlPart(worksheetPart, out var vmlRelId);
        AppendVmlShape(vmlPart, spec, shapeId);
        EnsureLegacyDrawing(worksheet, vmlRelId);

        // 4) the worksheet controls entry (x14 Choice + Fallback).
        AppendControlEntry(worksheet, spec, ctrlId, shapeId);

        worksheet.Save();
    }

    private static X14.FormControlProperties BuildFormControlProperties(FormControlAddSpec spec)
    {
        var fcp = new X14.FormControlProperties { ObjectType = ObjectTypeOf(spec.Kind) };
        fcp.AddNamespaceDeclaration("xdr", XdrNs);
        fcp.AddNamespaceDeclaration("x14", X14Ns);

        switch (spec.Kind)
        {
            case FormControlKind.CheckBox:
            case FormControlKind.OptionButton:
                fcp.Checked = X14.CheckedValues.Unchecked;
                fcp.LockText = true;
                fcp.NoThreeD = true;
                if (spec.LinkedCell is { } link)
                {
                    fcp.FmlaLink = link;
                }

                break;

            case FormControlKind.Spinner:
                fcp.Min = (uint)spec.Min!.Value;
                fcp.Max = (uint)spec.Max!.Value;
                fcp.Incremental = (uint)spec.Increment!.Value;
                fcp.Page = 1u;
                fcp.Horizontal = false;
                if (spec.LinkedCell is { } spinLink)
                {
                    fcp.FmlaLink = spinLink;
                }

                break;

            case FormControlKind.ComboBox:
                fcp.DropStyle = X14.DropStyleValues.Combo;
                fcp.DropLines = 8u;
                fcp.NoThreeD2 = true;
                ApplyListProps(fcp, spec);
                break;

            case FormControlKind.ListBox:
                fcp.SelectionType = X14.SelectionTypeValues.Single;
                fcp.NoThreeD2 = true;
                ApplyListProps(fcp, spec);
                break;

            case FormControlKind.Button:
                fcp.LockText = true;
                break;
        }

        return fcp;
    }

    /// <summary>combo/list share fmlaRange (the source range) + fmlaLink (the picked index/value).</summary>
    private static void ApplyListProps(X14.FormControlProperties fcp, FormControlAddSpec spec)
    {
        if (spec.ListFillRange is { } range)
        {
            fcp.FmlaRange = AbsRangeRef(range);
        }

        if (spec.LinkedCell is { } link)
        {
            fcp.FmlaLink = link;
        }
    }

    private static X14.ObjectTypeValues ObjectTypeOf(FormControlKind kind) => kind switch
    {
        FormControlKind.CheckBox => X14.ObjectTypeValues.CheckBox,
        FormControlKind.OptionButton => X14.ObjectTypeValues.Radio,
        FormControlKind.Spinner => X14.ObjectTypeValues.Spin,
        FormControlKind.ComboBox => X14.ObjectTypeValues.Drop,
        FormControlKind.ListBox => X14.ObjectTypeValues.List,
        _ => X14.ObjectTypeValues.Button,
    };

    private static uint NextShapeId(S.Worksheet worksheet)
    {
        // Shape ids are per-worksheet; start at 1025 like Excel and bump past any
        // control already present (read from the controls entries).
        var existing = worksheet.Descendants<S.Control>()
            .Select(c => c.ShapeId?.Value ?? 0u)
            .DefaultIfEmpty(1024u)
            .Max();
        return Math.Max(existing + 1, 1025u);
    }

    private static VmlDrawingPart EnsureVmlPart(WorksheetPart worksheetPart, out string relId)
    {
        var existing = worksheetPart.VmlDrawingParts.FirstOrDefault();
        if (existing is not null)
        {
            relId = worksheetPart.GetIdOfPart(existing);
            return existing;
        }

        var part = worksheetPart.AddNewPart<VmlDrawingPart>();
        relId = worksheetPart.GetIdOfPart(part);
        using var writer = new StreamWriter(part.GetStream(FileMode.Create), Encoding.UTF8);
        writer.Write(
            "<xml xmlns:v=\"" + VmlNs + "\" xmlns:o=\"" + ONs + "\" xmlns:x=\"" + XNs + "\">" +
            "<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>" +
            "<v:shapetype id=\"_x0000_t201\" coordsize=\"21600,21600\" o:spt=\"201\" " +
            "path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/>" +
            "<v:path shadowok=\"f\" o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>" +
            "</xml>");
        return part;
    }

    private static void AppendVmlShape(VmlDrawingPart part, FormControlAddSpec spec, uint shapeId)
    {
        // Read the current VML, splice a new shape before </xml>, rewrite.
        string text;
        using (var reader = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read)))
        {
            text = reader.ReadToEnd();
        }

        var shape = BuildVmlShape(spec, shapeId);
        var closeIndex = text.LastIndexOf("</xml>", StringComparison.Ordinal);
        text = closeIndex >= 0 ? text[..closeIndex] + shape + "</xml>" : text + shape;

        using var writer = new StreamWriter(part.GetStream(FileMode.Create), Encoding.UTF8);
        writer.Write(text);
    }

    private static string BuildVmlShape(FormControlAddSpec spec, uint shapeId)
    {
        // Anchor the shape one cell wide/tall from the cell, in the legacy
        // 8-number ClientData anchor form (fromCol,fromColOff,fromRow,fromRowOff,
        // toCol,toColOff,toRow,toRowOff).
        var fromCol = spec.AnchorColumn;
        var fromRow = spec.AnchorRow;
        var toCol = Math.Min(fromCol + 2, MaxColumns - 1);
        var toRow = Math.Min(fromRow + 1, MaxRows - 1);
        var anchor = string.Create(
            CultureInfo.InvariantCulture,
            $"{fromCol}, 0, {fromRow}, 0, {toCol}, 0, {toRow}, 0");

        var text = HtmlEscape(spec.Text ?? DefaultText(spec.Kind));
        var clientType = ClientDataType(spec.Kind);
        var clientExtra = ClientDataExtra(spec);

        // Position is approximate (Excel recomputes from the anchor); the anchor
        // is authoritative. width/height in points: ~2 cells wide, 1 tall.
        var marginLeft = (fromCol * 48 + 1).ToString(CultureInfo.InvariantCulture);
        var marginTop = (fromRow * 15 + 1).ToString(CultureInfo.InvariantCulture);
        var style =
            "position:absolute;margin-left:" + marginLeft + "pt;margin-top:" + marginTop +
            "pt;width:96pt;height:15pt;mso-wrap-style:tight";

        return
            "<v:shape id=\"_x0000_s" + shapeId.ToString(CultureInfo.InvariantCulture) + "\" " +
            "type=\"#_x0000_t201\" style=\"" + style + "\" o:connecttype=\"none\" o:button=\"t\" " +
            "fillcolor=\"window [65]\" strokecolor=\"windowText [64]\">" +
            "<v:fill color2=\"window [65]\"/><o:lock v:ext=\"edit\" rotation=\"t\"/>" +
            "<v:textbox style=\"mso-direction-alt:auto\" o:singleclick=\"f\">" +
            "<div style=\"text-align:left\"><font face=\"Calibri\" size=\"220\" color=\"auto\">" +
            text + "</font></div></v:textbox>" +
            "<x:ClientData ObjectType=\"" + clientType + "\">" +
            "<x:MoveWithCells/><x:SizeWithCells/>" +
            "<x:Anchor>" + anchor + "</x:Anchor>" +
            "<x:AutoFill>False</x:AutoFill>" + clientExtra +
            "</x:ClientData></v:shape>";
    }

    private static string ClientDataType(FormControlKind kind) => kind switch
    {
        FormControlKind.CheckBox => "Checkbox",
        FormControlKind.OptionButton => "Radio",
        FormControlKind.Spinner => "Spin",
        FormControlKind.ComboBox => "Drop",
        FormControlKind.ListBox => "List",
        _ => "Button",
    };

    /// <summary>The ClientData fields beyond the shared anchor/autofill, per kind.</summary>
    private static string ClientDataExtra(FormControlAddSpec spec)
    {
        var sb = new StringBuilder();
        switch (spec.Kind)
        {
            case FormControlKind.CheckBox:
            case FormControlKind.OptionButton:
                if (spec.LinkedCell is { } link)
                {
                    sb.Append("<x:FmlaLink>").Append(HtmlEscape(link)).Append("</x:FmlaLink>");
                }

                sb.Append("<x:Checked>0</x:Checked><x:TextVAlign>Center</x:TextVAlign>");
                if (spec.Kind == FormControlKind.OptionButton)
                {
                    sb.Append("<x:FirstButton/>");
                }

                break;

            case FormControlKind.Spinner:
                if (spec.LinkedCell is { } spinLink)
                {
                    sb.Append("<x:FmlaLink>").Append(HtmlEscape(spinLink)).Append("</x:FmlaLink>");
                }

                sb.Append(string.Create(
                    CultureInfo.InvariantCulture,
                    $"<x:Val>{spec.Min}</x:Val><x:Min>{spec.Min}</x:Min><x:Max>{spec.Max}</x:Max>" +
                    $"<x:Inc>{spec.Increment}</x:Inc><x:Page>1</x:Page>"));
                break;

            case FormControlKind.ComboBox:
            case FormControlKind.ListBox:
                if (spec.ListFillRange is { } range)
                {
                    sb.Append("<x:FmlaRange>").Append(HtmlEscape(AbsRangeRef(range))).Append("</x:FmlaRange>");
                }

                if (spec.LinkedCell is { } listLink)
                {
                    sb.Append("<x:FmlaLink>").Append(HtmlEscape(listLink)).Append("</x:FmlaLink>");
                }

                sb.Append("<x:Sel>0</x:Sel><x:NoThreeD2/>");
                if (spec.Kind == FormControlKind.ComboBox)
                {
                    sb.Append("<x:DropStyle>Combo</x:DropStyle><x:DropLines>8</x:DropLines>");
                }
                else
                {
                    sb.Append("<x:SelType>Single</x:SelType>");
                }

                break;
        }

        return sb.ToString();
    }

    private static string DefaultText(FormControlKind kind) => kind switch
    {
        FormControlKind.CheckBox => "Check Box",
        FormControlKind.OptionButton => "Option Button",
        FormControlKind.Button => "Button",
        _ => string.Empty,
    };

    private static void EnsureLegacyDrawing(S.Worksheet worksheet, string vmlRelId)
    {
        var existing = worksheet.Elements<S.LegacyDrawing>().FirstOrDefault();
        if (existing is not null)
        {
            existing.Id = vmlRelId; // keep it pointed at the (single) VML part
            return;
        }

        var ld = new S.LegacyDrawing { Id = vmlRelId };
        var successor = FirstOf(
            worksheet,
            typeof(S.LegacyDrawingHeaderFooter), typeof(S.DrawingHeaderFooter), typeof(S.Picture),
            typeof(S.OleObjects), typeof(S.Controls), typeof(S.WebPublishItems),
            typeof(S.TableParts), typeof(S.ExtensionList));
        if (successor is null)
        {
            worksheet.Append(ld);
        }
        else
        {
            worksheet.InsertBefore(ld, successor);
        }
    }

    private static void AppendControlEntry(S.Worksheet worksheet, FormControlAddSpec spec, string ctrlId, uint shapeId)
    {
        var controls = worksheet.Elements<S.Controls>().FirstOrDefault();
        if (controls is null)
        {
            controls = new S.Controls();
            var successor = FirstOf(
                worksheet, typeof(S.WebPublishItems), typeof(S.TableParts), typeof(S.ExtensionList));
            if (successor is null)
            {
                worksheet.Append(controls);
            }
            else
            {
                worksheet.InsertBefore(controls, successor);
            }
        }

        controls.Append(BuildControlAlternateContent(spec, ctrlId, shapeId));
    }

    private static OpenXmlElement BuildControlAlternateContent(FormControlAddSpec spec, string ctrlId, uint shapeId)
    {
        var name = ControlName(spec.Kind, shapeId);
        var fromCol = spec.AnchorColumn;
        var fromRow = spec.AnchorRow;
        var toCol = Math.Min(fromCol + 2, MaxColumns - 1);
        var toRow = Math.Min(fromRow + 1, MaxRows - 1);
        var anchor =
            "<xdr:from><xdr:col>" + fromCol + "</xdr:col><xdr:colOff>0</xdr:colOff>" +
            "<xdr:row>" + fromRow + "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>" +
            "<xdr:to><xdr:col>" + toCol + "</xdr:col><xdr:colOff>0</xdr:colOff>" +
            "<xdr:row>" + toRow + "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>";

        var controlPrAttrs = new StringBuilder("defaultSize=\"0\" autoFill=\"0\" autoLine=\"0\"");
        if (spec.LinkedCell is { } link)
        {
            controlPrAttrs.Append(" linkedCell=\"").Append(XmlEscape(link)).Append('"');
        }

        if (spec.ListFillRange is { } range && spec.Kind is FormControlKind.ComboBox or FormControlKind.ListBox)
        {
            controlPrAttrs.Append(" listFillRange=\"").Append(XmlEscape(AbsRangeRef(range))).Append('"');
        }

        controlPrAttrs.Append(" r:id=\"").Append(ctrlId).Append('"');

        var xml =
            "<mc:AlternateContent xmlns:mc=\"" + Mc + "\" xmlns:x14=\"" + X14Ns + "\">" +
            "<mc:Choice Requires=\"x14\">" +
            "<x14:control shapeId=\"" + shapeId.ToString(CultureInfo.InvariantCulture) + "\" " +
            "r:id=\"" + ctrlId + "\" name=\"" + XmlEscape(name) + "\" xmlns:r=\"" + RNs + "\">" +
            "<x14:controlPr " + controlPrAttrs + ">" +
            "<xm:anchor xmlns:xm=\"" + XmNs + "\" xmlns:xdr=\"" + XdrNs + "\">" + anchor + "</xm:anchor>" +
            "</x14:controlPr></x14:control></mc:Choice>" +
            "<mc:Fallback>" +
            "<control xmlns=\"" + Main + "\" shapeId=\"" + shapeId.ToString(CultureInfo.InvariantCulture) + "\" " +
            "r:id=\"" + ctrlId + "\" name=\"" + XmlEscape(name) + "\" xmlns:r=\"" + RNs + "\"/>" +
            "</mc:Fallback></mc:AlternateContent>";
        return new AlternateContent(xml);
    }

    private static string ControlName(FormControlKind kind, uint shapeId)
    {
        var n = (shapeId - 1024).ToString(CultureInfo.InvariantCulture);
        return kind switch
        {
            FormControlKind.CheckBox => "Check Box " + n,
            FormControlKind.OptionButton => "Option Button " + n,
            FormControlKind.Spinner => "Spinner " + n,
            FormControlKind.ComboBox => "Drop Down " + n,
            FormControlKind.ListBox => "List Box " + n,
            _ => "Button " + n,
        };
    }

    private static void RemoveControl(SpreadsheetDocument document, FormControlRemoveSpec spec)
    {
        var worksheetPart = SheetPartOrThrow(document, spec.SheetName);
        var worksheet = worksheetPart.Worksheet ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            $"Sheet '{spec.SheetName}' has no worksheet XML.",
            "Restore a snapshot ('aioffice snapshot list') or re-export the file from its source.");
        var controls = worksheet.Elements<S.Controls>().FirstOrDefault();
        var entries = controls?.Elements<AlternateContent>().ToList() ?? [];
        if (controls is null || spec.Index > entries.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                $"formControl[{spec.Index}] on '{spec.SheetName}' disappeared between validation and the write pass.",
                "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
        }

        var entry = entries[spec.Index - 1];
        var (ctrlRelId, shapeId) = ControlRefs(entry);
        entry.Remove();

        // Drop the ctrlProp part this entry referenced.
        if (ctrlRelId is not null &&
            worksheetPart.GetPartById(ctrlRelId) is ControlPropertiesPart ctrlPart)
        {
            worksheetPart.DeletePart(ctrlPart);
        }

        // Strip the matching VML shape so Excel does not show a ghost control.
        if (shapeId is { } removedShape)
        {
            RemoveVmlShape(worksheetPart, removedShape);
        }

        // If no controls remain, drop the controls element (the VML part stays if
        // notes still use it; an empty VML part is harmless and validator-clean).
        if (!controls.Elements<AlternateContent>().Any())
        {
            controls.Remove();
        }

        worksheet.Save();
    }

    private static (string? RelId, uint? ShapeId) ControlRefs(AlternateContent entry)
    {
        var choice = entry.GetFirstChild<AlternateContentChoice>();
        var control = choice?.ChildElements.FirstOrDefault(e => e.LocalName == "control");
        var attrs = control?.GetAttributes().ToList() ?? [];
        var relId = attrs.FirstOrDefault(a => a.LocalName == "id").Value;
        uint? shapeId = uint.TryParse(
            attrs.FirstOrDefault(a => a.LocalName == "shapeId").Value,
            NumberStyles.None, CultureInfo.InvariantCulture, out var s)
            ? s
            : null;
        return (relId, shapeId);
    }

    private static void RemoveVmlShape(WorksheetPart worksheetPart, uint shapeId)
    {
        foreach (var vml in worksheetPart.VmlDrawingParts)
        {
            string text;
            using (var reader = new StreamReader(vml.GetStream(FileMode.Open, FileAccess.Read)))
            {
                text = reader.ReadToEnd();
            }

            var pattern = "<v:shape\\b[^>]*\\bid=\"_x0000_s" +
                          shapeId.ToString(CultureInfo.InvariantCulture) + "\"[^>]*>.*?</v:shape>";
            var updated = System.Text.RegularExpressions.Regex.Replace(
                text, pattern, string.Empty, System.Text.RegularExpressions.RegexOptions.Singleline);
            if (!ReferenceEquals(updated, text) && updated.Length != text.Length)
            {
                using var writer = new StreamWriter(vml.GetStream(FileMode.Create), Encoding.UTF8);
                writer.Write(updated);
                return;
            }
        }
    }

    // ----- raw read-back --------------------------------------------------------

    /// <summary>All form controls in the package, in sheet order then control order (1-based per sheet).</summary>
    public static List<FormControlInfo> Read(SpreadsheetDocument document)
    {
        var result = new List<FormControlInfo>();
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is null)
        {
            return result;
        }

        foreach (var sheet in workbookPart.Workbook.Descendants<S.Sheet>())
        {
            if (sheet.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart ||
                worksheetPart.Worksheet?.Elements<S.Controls>().FirstOrDefault() is not { } controls)
            {
                continue;
            }

            var sheetName = sheet.Name?.Value ?? string.Empty;
            // ClosedXML re-serializes the foreign xm:anchor with empty markers on
            // its own saves, so the anchor cell is read from the VML <x:Anchor>
            // (which it preserves byte-identical), keyed by shape id.
            var anchorsByShape = ReadVmlAnchors(worksheetPart);
            var index = 0;
            foreach (var entry in controls.Elements<AlternateContent>())
            {
                index++;
                result.Add(Describe(sheetName, index, entry, worksheetPart, anchorsByShape));
            }
        }

        return result;
    }

    /// <summary>Maps each VML control shape's id to its anchor cell, read from <c>x:Anchor</c>.</summary>
    private static Dictionary<uint, string> ReadVmlAnchors(WorksheetPart worksheetPart)
    {
        var result = new Dictionary<uint, string>();
        foreach (var vml in worksheetPart.VmlDrawingParts)
        {
            string text;
            using (var reader = new StreamReader(vml.GetStream(FileMode.Open, FileAccess.Read)))
            {
                text = reader.ReadToEnd();
            }

            // Each <v:shape id="_x0000_sNNNN" ...> carries an <x:Anchor> with the
            // 8-number legacy anchor (fromCol,fromColOff,fromRow,...). Pair them up.
            foreach (System.Text.RegularExpressions.Match shape in
                     System.Text.RegularExpressions.Regex.Matches(
                         text, "<v:shape\\b[^>]*\\bid=\"_x0000_s(?<id>\\d+)\"[^>]*>(?<body>.*?)</v:shape>",
                         System.Text.RegularExpressions.RegexOptions.Singleline))
            {
                if (!uint.TryParse(shape.Groups["id"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var id))
                {
                    continue;
                }

                var anchorMatch = System.Text.RegularExpressions.Regex.Match(
                    shape.Groups["body"].Value, "<x:Anchor>(?<a>[^<]*)</x:Anchor>");
                if (!anchorMatch.Success)
                {
                    continue;
                }

                var parts = anchorMatch.Groups["a"].Value.Split(',', StringSplitOptions.TrimEntries);
                if (parts.Length >= 3 &&
                    int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out var col) &&
                    int.TryParse(parts[2], NumberStyles.None, CultureInfo.InvariantCulture, out var row))
                {
                    result[id] = ExcelCharts.ColumnLetters(col + 1) +
                                 (row + 1).ToString(CultureInfo.InvariantCulture);
                }
            }
        }

        return result;
    }

    private static FormControlInfo Describe(
        string sheetName, int index, AlternateContent entry, WorksheetPart worksheetPart,
        IReadOnlyDictionary<uint, string> anchorsByShape)
    {
        var choice = entry.GetFirstChild<AlternateContentChoice>();
        var control = choice?.ChildElements.FirstOrDefault(e => e.LocalName == "control");
        var attrs = control?.GetAttributes().ToList() ?? [];
        var linkedCell = attrs.FirstOrDefault(a => a.LocalName == "linkedCell").Value;
        var ctrlRelId = attrs.FirstOrDefault(a => a.LocalName == "id").Value;
        var shapeIdText = attrs.FirstOrDefault(a => a.LocalName == "shapeId").Value;

        string kindName = "unknown";
        int? min = null, max = null, increment = null;
        if (ctrlRelId is not null && worksheetPart.GetPartById(ctrlRelId) is ControlPropertiesPart ctrlPart &&
            ctrlPart.FormControlProperties is { } fcp)
        {
            kindName = KindNameOf(fcp.ObjectType);
            if (fcp.FmlaLink?.Value is { } fl && linkedCell is null)
            {
                linkedCell = fl;
            }

            if (fcp.Min?.Value is { } mn)
            {
                min = (int)mn;
            }

            if (fcp.Max?.Value is { } mx)
            {
                max = (int)mx;
            }

            if (fcp.Incremental?.Value is { } inc)
            {
                increment = (int)inc;
            }
        }

        var cell = shapeIdText is not null &&
                   uint.TryParse(shapeIdText, NumberStyles.None, CultureInfo.InvariantCulture, out var shapeId) &&
                   anchorsByShape.TryGetValue(shapeId, out var anchorCell)
            ? anchorCell
            : "A1";

        var path = string.Create(
            CultureInfo.InvariantCulture, $"/{ExcelPaths.QuoteSheet(sheetName)}/formControl[{index}]");
        return new FormControlInfo(
            sheetName, index, path, kindName, cell, linkedCell, null, null, min, max, increment);
    }

    private static string KindNameOf(EnumValue<X14.ObjectTypeValues>? objectType)
    {
        if (objectType is null)
        {
            return "unknown";
        }

        var v = objectType.Value;
        if (v == X14.ObjectTypeValues.CheckBox) return "checkbox";
        if (v == X14.ObjectTypeValues.Radio) return "optionButton";
        if (v == X14.ObjectTypeValues.Spin) return "spinner";
        if (v == X14.ObjectTypeValues.Drop) return "comboBox";
        if (v == X14.ObjectTypeValues.List) return "listBox";
        if (v == X14.ObjectTypeValues.Button) return "button";
        return "unknown";
    }

    // ----- small helpers --------------------------------------------------------

    private static OpenXmlElement? FirstOf(S.Worksheet worksheet, params Type[] types) =>
        worksheet.ChildElements.FirstOrDefault(e => types.Any(t => t.IsInstanceOfType(e)));

    private static (int Column, int Row) ParseCell(string text, int opIndex)
    {
        var upper = text.ToUpperInvariant();
        var i = 0;
        while (i < upper.Length && char.IsLetter(upper[i]))
        {
            i++;
        }

        if (i == 0 || i == upper.Length || i > 3 ||
            !int.TryParse(upper[i..], NumberStyles.None, CultureInfo.InvariantCulture, out var row) || row < 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{text}' is not a usable cell reference.",
                "Use a single cell like E2 (no sheet prefix, no range).");
        }

        var column = 0;
        foreach (var ch in upper[..i])
        {
            column = column * 26 + (ch - 'A' + 1);
        }

        // 0-based for the anchor markers.
        return (column - 1, row - 1);
    }

    /// <summary>Absolute single-cell ref ($F$2) for fmlaLink.</summary>
    private static string AbsRef(string cell)
    {
        var upper = cell.ToUpperInvariant();
        var i = 0;
        while (i < upper.Length && char.IsLetter(upper[i]))
        {
            i++;
        }

        return "$" + upper[..i] + "$" + upper[i..];
    }

    /// <summary>Absolute range ref ($H$1:$H$3) for fmlaRange/listFillRange.</summary>
    private static string AbsRangeRef(string range)
    {
        var parts = range.Split(':');
        return parts.Length == 2 ? AbsRef(parts[0]) + ":" + AbsRef(parts[1]) : AbsRef(range);
    }

    private static string RequiredString(JsonObject props, string key, int opIndex)
    {
        if (props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
            value.GetValueKind() == JsonValueKind.String &&
            value.GetValue<string>() is { Length: > 0 } text)
        {
            return text;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: add formControl needs the '{key}' prop.",
            $"Pass it as a string, e.g. {{\"{key}\":\"…\"}}.");
    }

    private static string? OptionalString(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String
            ? value.GetValue<string>()
            : null;

    private static int? OptionalInt(JsonObject props, string key, int opIndex)
    {
        if (!props.TryGetPropertyValue(key, out var node) || node is null)
        {
            return null;
        }

        if (node is JsonValue value)
        {
            if (value.GetValueKind() == JsonValueKind.Number && value.TryGetValue<int>(out var number))
            {
                return number;
            }

            if (value.GetValueKind() == JsonValueKind.String &&
                int.TryParse(value.GetValue<string>(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: '{key}' must be a whole number.",
            $"Pass e.g. {{\"{key}\":1}}.");
    }

    private static IReadOnlyList<string>? OptionalStrings(JsonObject props, string key, int opIndex)
    {
        if (!props.TryGetPropertyValue(key, out var node) || node is null)
        {
            return null;
        }

        if (node is not JsonArray array)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{key}' must be an array of strings.",
                "Pass e.g. {\"items\":[\"Red\",\"Green\",\"Blue\"]}.");
        }

        var result = new List<string>(array.Count);
        foreach (var element in array)
        {
            if (element is JsonValue v && v.GetValueKind() == JsonValueKind.String)
            {
                result.Add(v.GetValue<string>());
            }
            else
            {
                result.Add(element?.ToString() ?? string.Empty);
            }
        }

        return result;
    }

    private static string HtmlEscape(string text) => XmlEscape(text);

    private static string XmlEscape(string text) => text
        .Replace("&", "&amp;", StringComparison.Ordinal)
        .Replace("<", "&lt;", StringComparison.Ordinal)
        .Replace(">", "&gt;", StringComparison.Ordinal)
        .Replace("\"", "&quot;", StringComparison.Ordinal);
}
