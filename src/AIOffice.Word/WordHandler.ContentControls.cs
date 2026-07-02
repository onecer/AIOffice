using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using WordLock = DocumentFormat.OpenXml.Wordprocessing.Lock; // 'Lock' alone clashes with System.Threading.Lock

namespace AIOffice.Word;

/// <summary>
/// Content controls (w:sdt — structured document tags) let agents fill
/// templates: a tagged, titled placeholder that holds text, a dropdown choice,
/// a date or a checkbox. They are added at block level next to a paragraph,
/// addressed by <c>/sdt[@tag=clientName]</c> (or positionally <c>/sdt[i]</c>),
/// their value set with <c>set</c>, listed by <c>read --view fields</c>, and
/// removed (content kept, by default).
/// </summary>
public sealed partial class WordHandler
{
    private static readonly string[] ContentControlKinds = ["text", "dropdown", "date", "checkbox"];

    private static readonly string[] ContentControlLockTokens =
        ["sdtLocked", "contentLocked", "sdtContentLocked", "unlocked"];

    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[2]","type":"contentControl","props":{"kind":…,"tag":…,"title":…}}</c>:
    /// inserts a block-level w:sdt after the anchor paragraph. kind picks the
    /// sdtPr child (text/dropdown/date/checkbox); dropdown needs props.items.
    /// </summary>
    private static object ApplyAddContentControl(WordprocessingDocument doc, EditOp op)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        var kind = props["kind"] is { } kindNode ? NodeToString(kindNode) : "text";
        if (!ContentControlKinds.Contains(kind, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Content-control kind '{kind}' is not supported.",
                "Use kind text, dropdown, date or checkbox.",
                candidates: ContentControlKinds);
        }

        var tag = props["tag"] is { } tagNode ? NodeToString(tagNode) : null;
        if (string.IsNullOrWhiteSpace(tag))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type contentControl needs props.tag (the field's stable identifier).",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[2]\",\"type\":\"contentControl\",\"props\":{\"kind\":\"text\",\"tag\":\"clientName\",\"title\":\"Client Name\"}}.");
        }

        if (EnumerateContentControls(doc).Any(s => string.Equals(TagOf(s), tag, StringComparison.OrdinalIgnoreCase)))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"A content control with tag '{tag}' already exists.",
                "Tags are unique. Remove the old one ({\"op\":\"remove\",\"path\":\"" + ContentControlPath(tag) +
                "\"}) or pick a different tag.");
        }

        var title = props["title"] is { } titleNode ? NodeToString(titleNode) : null;

        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (anchor.Element is not Paragraph)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Content controls are placed next to paragraphs, not '{anchor.Type}'.",
                anchor.Type is "tc"
                    ? $"Target the paragraph inside the cell: {anchor.CanonicalPath}/p[1]."
                    : "Pick a paragraph path, e.g. /body/p[2].");
        }

        var sdt = BuildContentControl(doc, kind, tag, title, props);
        var position = op.Position ?? "after";
        if (position is not ("before" or "after"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add contentControl position '{position}' is not valid.",
                "Use position before or after a paragraph (default: after).",
                candidates: ["before", "after"]);
        }

        if (position == "before")
        {
            anchor.Element.InsertBeforeSelf(sdt);
        }
        else
        {
            anchor.Element.InsertAfterSelf(sdt);
        }

        return new
        {
            op = "add",
            type = "contentControl",
            path = ContentControlPath(tag),
            kind,
            tag,
            title,
            anchor = anchor.CanonicalPath,
        };
    }

    /// <summary>Builds a schema-valid block-level w:sdt for one kind.</summary>
    private static SdtBlock BuildContentControl(
        WordprocessingDocument doc, string kind, string tag, string? title, JsonObject props)
    {
        var sdtPr = new SdtProperties(
            new SdtId { Val = NextContentControlId(doc) },
            new Tag { Val = tag });
        if (title is { Length: > 0 })
        {
            sdtPr.AppendChild(new SdtAlias { Val = title });
        }

        // Optional editing lock (w:lock): placed in the ECMA-canonical slot after
        // w:alias and before w:dataBinding — pinned by the SdtPr_ChildOrdering test
        // and validator-clean both alone and combined with a dataBinding.
        if (BuildContentControlLock(props["lock"]) is { } sdtLock)
        {
            sdtPr.AppendChild(sdtLock);
        }

        // Optional external XML data binding (w:dataBinding): after w:alias, before the kind child.
        if (BuildDataBinding(props["dataBinding"]) is { } binding)
        {
            sdtPr.AppendChild(binding);
        }

        string displayText;
        switch (kind)
        {
            case "dropdown":
            {
                var ddl = new SdtContentDropDownList();
                var items = props["items"] as JsonArray
                    ?? throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        "A dropdown content control needs props.items (the choices).",
                        "Example: {\"kind\":\"dropdown\",\"tag\":\"region\",\"items\":[\"North\",\"South\"]}.");
                if (items.Count == 0)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        "A dropdown content control needs at least one item in props.items.",
                        "Pass props.items like [\"A\",\"B\",\"C\"].");
                }

                foreach (var item in items)
                {
                    var label = NodeToString(item);
                    ddl.AppendChild(new ListItem { DisplayText = label, Value = label });
                }

                sdtPr.AppendChild(ddl);
                var selected = props["text"] is { } sel ? NodeToString(sel) : null;
                displayText = selected is { Length: > 0 } ? selected : NodeToString(items[0]);
                break;
            }

            case "date":
            {
                var date = new SdtContentDate();
                if (props["text"] is { } dn && DateTime.TryParse(NodeToString(dn), CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var parsed))
                {
                    date.FullDate = parsed;
                    displayText = parsed.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                }
                else
                {
                    displayText = props["text"] is { } pt ? NodeToString(pt) : "(choose a date)";
                }

                date.AppendChild(new DateFormat { Val = "yyyy-MM-dd" });
                date.AppendChild(new LanguageId { Val = "en-US" });
                sdtPr.AppendChild(date);
                break;
            }

            case "checkbox":
            {
                var isChecked = props["checked"] is { } c && WordFormatting.ParseBool("checked", NodeToString(c));
                var checkbox = new W14.SdtContentCheckBox(
                    new W14.Checked { Val = isChecked ? W14.OnOffValues.One : W14.OnOffValues.Zero },
                    new W14.CheckedState { Font = "MS Gothic", Val = "2612" },
                    new W14.UncheckedState { Font = "MS Gothic", Val = "2610" });
                sdtPr.AppendChild(checkbox);
                DeclareDocumentW14Ignorable(doc); // w14:checkbox lives in the 2010 namespace
                displayText = isChecked ? "☒" : "☐";
                break;
            }

            default: // text
                sdtPr.AppendChild(new SdtContentText());
                displayText = props["text"] is { } t ? NodeToString(t) : (title is { Length: > 0 } ? title : tag);
                break;
        }

        var run = new Run(NewText(displayText));
        var content = new SdtContentBlock(new Paragraph(run));
        return new SdtBlock(sdtPr, content);
    }

    /// <summary>
    /// Parses props.dataBinding = {xpath (required, non-empty), storeItemId?, prefixMappings?}
    /// into a w:dataBinding, or returns null when no binding was requested.
    /// </summary>
    private static DataBinding? BuildDataBinding(JsonNode? node)
    {
        if (node is null)
        {
            return null;
        }

        if (node is not JsonObject obj)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "props.dataBinding must be an object with an xpath.",
                "Example: {\"dataBinding\":{\"xpath\":\"/root/client\"}}.");
        }

        var xpath = obj["xpath"] is { } xp ? NodeToString(xp) : null;
        if (string.IsNullOrWhiteSpace(xpath))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "props.dataBinding needs a non-empty xpath (the XML path the control binds to).",
                "Example: {\"dataBinding\":{\"xpath\":\"/root/client\",\"storeItemId\":\"{GUID}\"}}.",
                candidates: ["xpath"]);
        }

        // w:storeItemID is REQUIRED by the WordprocessingML schema, so it is always
        // emitted; when the caller omits it we write an empty value (a valid binding
        // that is not yet attached to a custom-XML store item).
        var storeItemId = obj["storeItemId"] is { } sid ? NodeToString(sid) : string.Empty;
        var binding = new DataBinding { XPath = xpath, StoreItemId = storeItemId };

        if (obj["prefixMappings"] is { } pm && NodeToString(pm) is { Length: > 0 } prefixMappings)
        {
            binding.PrefixMappings = prefixMappings;
        }

        return binding;
    }

    /// <summary>
    /// Parses props.lock (sdtLocked = Word won't delete, contentLocked = Word won't
    /// edit contents, sdtContentLocked = both, unlocked = explicitly none) into a
    /// w:lock, or returns null when no lock was requested. HONEST semantics: the
    /// lock is Word-UI metadata only — Word greys out the matching commands, but
    /// AIOffice set/remove (and any other OOXML tool) still edit the control.
    /// </summary>
    private static WordLock? BuildContentControlLock(JsonNode? node)
    {
        if (node is null)
        {
            return null;
        }

        var value = NodeToString(node);
        return new WordLock
        {
            Val = value switch
            {
                "sdtLocked" => LockingValues.SdtLocked,
                "contentLocked" => LockingValues.ContentLocked,
                "sdtContentLocked" => LockingValues.SdtContentLocked,
                "unlocked" => LockingValues.Unlocked,
                _ => throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Content-control lock '{value}' is not supported.",
                    "Use lock sdtLocked (Word won't delete), contentLocked (Word won't edit contents), sdtContentLocked (both) or unlocked.",
                    candidates: ContentControlLockTokens),
            },
        };
    }

    // -------------------------------------------------------------------- set

    /// <summary>
    /// <c>{"op":"set","path":"/sdt[@tag=clientName]","props":{"text":"Acme"}}</c>
    /// or <c>{"checked":true}</c>: writes the control's current value.
    /// </summary>
    private static object ApplySetContentControl(WordprocessingDocument doc, EditOp op)
    {
        var (sdt, _) = ResolveContentControl(doc, DocPath.Parse(op.Path));
        var props = RequireProps(op);
        var kind = ContentControlKind(sdt);
        var tag = TagOf(sdt) ?? string.Empty;

        if (kind == "checkbox")
        {
            var checkedNode = props["checked"] ?? props["value"] ?? throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "set on a checkbox content control needs props.checked (true/false).",
                "Example: {\"op\":\"set\",\"path\":\"" + ContentControlPath(tag) + "\",\"props\":{\"checked\":true}}.");
            var isChecked = WordFormatting.ParseBool("checked", NodeToString(checkedNode));
            SetCheckboxState(sdt, isChecked);
            SetContentControlDisplay(sdt, isChecked ? "☒" : "☐");
            return new { op = "set", path = ContentControlPath(tag), kind, @checked = isChecked };
        }

        var textNode = props["text"] ?? props["value"] ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"set on a {kind} content control needs props.text (the value).",
            "Example: {\"op\":\"set\",\"path\":\"" + ContentControlPath(tag) + "\",\"props\":{\"text\":\"Acme\"}}.");
        var value = NodeToString(textNode);

        if (kind == "dropdown")
        {
            var allowed = DropdownItems(sdt);
            if (allowed.Count > 0 && !allowed.Contains(value, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"'{value}' is not one of the dropdown's items.",
                    "Pick one of: " + string.Join(", ", allowed) + ".",
                    candidates: allowed);
            }
        }

        if (kind == "date" &&
            DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var parsed))
        {
            value = parsed.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            var dateProps = sdt.SdtProperties?.GetFirstChild<SdtContentDate>();
            if (dateProps is not null)
            {
                dateProps.FullDate = parsed;
            }
        }

        SetContentControlDisplay(sdt, value);
        return new { op = "set", path = ContentControlPath(tag), kind, value };
    }

    private static void SetCheckboxState(SdtBlock sdt, bool isChecked)
    {
        var checkbox = sdt.SdtProperties?.GetFirstChild<W14.SdtContentCheckBox>();
        if (checkbox is null)
        {
            return;
        }

        checkbox.Checked ??= new W14.Checked();
        checkbox.Checked.Val = isChecked ? W14.OnOffValues.One : W14.OnOffValues.Zero;
    }

    /// <summary>Replaces the sdt's content with a single run carrying the new display text.</summary>
    private static void SetContentControlDisplay(SdtBlock sdt, string text)
    {
        var content = sdt.GetFirstChild<SdtContentBlock>();
        if (content is null)
        {
            content = new SdtContentBlock();
            sdt.AppendChild(content);
        }

        var paragraph = content.GetFirstChild<Paragraph>();
        if (paragraph is null)
        {
            paragraph = new Paragraph();
            content.AppendChild(paragraph);
        }

        WordFormatting.ReplaceParagraphText(paragraph, text);

        // A block sdt must keep exactly one paragraph; drop any extras left over.
        foreach (var extra in content.Elements<Paragraph>().Skip(1).ToList())
        {
            extra.Remove();
        }
    }

    // ----------------------------------------------------------------- remove

    /// <summary>
    /// remove /sdt[@tag=X]: unwraps the control, keeping its content paragraph in
    /// place (so the filled value survives). props.keepContent=false drops it whole.
    /// </summary>
    private static object ApplyRemoveContentControl(WordprocessingDocument doc, EditOp op)
    {
        var (sdt, _) = ResolveContentControl(doc, DocPath.Parse(op.Path));
        var tag = TagOf(sdt) ?? string.Empty;
        var keepContent = op.Props?["keepContent"] is { } k
            ? WordFormatting.ParseBool("keepContent", NodeToString(k))
            : true;

        var parent = sdt.Parent;
        if (keepContent && sdt.GetFirstChild<SdtContentBlock>() is { } content && parent is not null)
        {
            foreach (var child in content.ChildElements.ToList())
            {
                child.Remove();
                sdt.InsertBeforeSelf(child);
            }
        }

        sdt.Remove();
        return new { op = "remove", path = ContentControlPath(tag), type = "contentControl", keptContent = keepContent };
    }

    // ------------------------------------------------------------------- read

    /// <summary>read --view fields: every content control AND merge field with its current value.</summary>
    private static object FieldsView(WordprocessingDocument doc)
    {
        var fields = EnumerateContentControls(doc)
            .Select((sdt, i) => (object)ContentControlShape(doc, sdt, i + 1))
            .ToList();

        // Merge fields share the fields view: each lists its name and current value.
        var mergeFields = MergeFieldsView(doc);
        fields.AddRange(mergeFields);

        // IF merge fields (v1.4.0) join the same view, distinguished by kind:"ifField".
        var ifFields = IfFieldsView(doc);
        fields.AddRange(ifFields);

        // Legacy form fields (v1.3.0) join the same view, distinguished by kind:"formField".
        var formFields = FormFieldsView(doc);
        fields.AddRange(formFields);

        return new { view = "fields", count = fields.Count, fields };
    }

    /// <summary>All block-level content controls (body first, then headers/footers), document order.</summary>
    private static List<SdtBlock> EnumerateContentControls(WordprocessingDocument doc) =>
        [.. ContentControlRoots(doc).SelectMany(root => root.Descendants<SdtBlock>())];

    private static IEnumerable<OpenXmlElement> ContentControlRoots(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart?.Document?.Body is { } body)
        {
            yield return body;
        }

        foreach (var root in WordAddress.HeaderFooterRoots(doc))
        {
            yield return root.Element;
        }
    }

    private static int NextContentControlId(WordprocessingDocument doc) =>
        EnumerateContentControls(doc)
            .Select(s => s.SdtProperties?.GetFirstChild<SdtId>()?.Val?.Value ?? 0)
            .DefaultIfEmpty(0)
            .Max() + 1;

    /// <summary>Resolves /sdt[@tag=X] or positional /sdt[i] to its block + shape, or throws invalid_path.</summary>
    private static (SdtBlock Sdt, Dictionary<string, object?> Shape) ResolveContentControl(WordprocessingDocument doc, DocPath path)
    {
        var controls = EnumerateContentControls(doc);
        var segment = path.Segments[0];

        if (path.Segments.Count != 1 || segment.Name != "sdt")
        {
            throw ContentControlNotFound($"'{path.ToCanonicalString()}' is not a content-control path.", controls);
        }

        SdtBlock? match = null;
        var index = 0;
        if (segment.Id is { } wanted)
        {
            if (segment.IdAttribute is not (null or "tag"))
            {
                throw ContentControlNotFound(
                    $"Content controls are addressed by their tag, not @{segment.IdAttribute}.",
                    controls);
            }

            match = controls.FirstOrDefault(s => string.Equals(TagOf(s), wanted, StringComparison.OrdinalIgnoreCase));
            if (match is null)
            {
                throw ContentControlNotFound($"No content control has tag '{wanted}'.", controls);
            }

            index = controls.IndexOf(match) + 1;
        }
        else
        {
            index = segment.Index ?? 1;
            if (index > controls.Count)
            {
                throw ContentControlNotFound(
                    controls.Count == 0
                        ? "This document has no content controls."
                        : $"/sdt[{index}] does not exist; there are {controls.Count} content control(s).",
                    controls);
            }

            match = controls[index - 1];
        }

        return (match, ContentControlShape(doc, match, index));
    }

    private static AiofficeException ContentControlNotFound(string message, List<SdtBlock> controls) => new(
        ErrorCodes.InvalidPath,
        message,
        "Run 'aioffice read <file> --view fields' to list content controls with their tags.",
        candidates: [.. controls.Take(5).Select(s => ContentControlPath(TagOf(s) ?? string.Empty))]);

    private static Dictionary<string, object?> ContentControlShape(WordprocessingDocument doc, SdtBlock sdt, int index)
    {
        _ = doc;
        var kind = ContentControlKind(sdt);
        var tag = TagOf(sdt);
        var path = tag is { Length: > 0 } ? ContentControlPath(tag) : $"/sdt[{index}]";
        var shape = new Dictionary<string, object?>
        {
            ["path"] = path,
            ["kind"] = kind,
            ["tag"] = tag,
            ["title"] = sdt.SdtProperties?.GetFirstChild<SdtAlias>()?.Val?.Value,
            ["value"] = ContentControlValue(sdt, kind),
        };

        if (kind == "dropdown")
        {
            shape["items"] = DropdownItems(sdt);
        }

        // External XML data binding (w:dataBinding), only when the control carries one —
        // legacy controls stay byte-identical (no key emitted when absent).
        if (sdt.SdtProperties?.GetFirstChild<DataBinding>() is { } binding)
        {
            var dataBinding = new Dictionary<string, object?> { ["xpath"] = binding.XPath?.Value };
            if (binding.StoreItemId?.Value is { Length: > 0 } storeItemId)
            {
                dataBinding["storeItemId"] = storeItemId;
            }

            if (binding.PrefixMappings?.Value is { Length: > 0 } prefixMappings)
            {
                dataBinding["prefixMappings"] = prefixMappings;
            }

            shape["dataBinding"] = dataBinding;
        }

        // Editing lock (w:lock), only when the control carries one — legacy controls
        // emit no key. A w:lock without @w:val means unlocked (ECMA-376 §17.5.2.23).
        if (sdt.SdtProperties?.GetFirstChild<WordLock>() is { } sdtLock)
        {
            shape["lock"] = sdtLock.Val?.InnerText ?? "unlocked";
        }

        return shape;
    }

    /// <summary>The kind name from the sdtPr child (text/dropdown/date/checkbox).</summary>
    private static string ContentControlKind(SdtBlock sdt)
    {
        var pr = sdt.SdtProperties;
        if (pr is null)
        {
            return "text";
        }

        if (pr.GetFirstChild<W14.SdtContentCheckBox>() is not null)
        {
            return "checkbox";
        }

        if (pr.GetFirstChild<SdtContentDropDownList>() is not null)
        {
            return "dropdown";
        }

        if (pr.GetFirstChild<SdtContentDate>() is not null)
        {
            return "date";
        }

        return "text";
    }

    private static object? ContentControlValue(SdtBlock sdt, string kind)
    {
        if (kind == "checkbox")
        {
            var checkbox = sdt.SdtProperties?.GetFirstChild<W14.SdtContentCheckBox>();
            return checkbox?.Checked?.Val is { } v && v == W14.OnOffValues.One;
        }

        return sdt.GetFirstChild<SdtContentBlock>()?.InnerText ?? string.Empty;
    }

    private static List<string> DropdownItems(SdtBlock sdt) =>
        [.. sdt.SdtProperties?.GetFirstChild<SdtContentDropDownList>()?.Elements<ListItem>()
            .Select(i => i.Value?.Value ?? i.DisplayText?.Value ?? string.Empty) ?? []];

    private static string? TagOf(SdtBlock sdt) => sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;

    private static string ContentControlPath(string tag) => $"/sdt[@tag={tag}]";

    /// <summary>Declares w14 (and mc) on the document root so w14:checkbox validates.</summary>
    private static void DeclareDocumentW14Ignorable(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart?.Document is not { } document)
        {
            return;
        }

        if (document.LookupNamespace("w14") is null)
        {
            document.AddNamespaceDeclaration("w14", W14Namespace);
        }

        if (document.LookupNamespace("mc") is null)
        {
            document.AddNamespaceDeclaration("mc", McNamespace);
        }

        var mc = document.MCAttributes ??= new MarkupCompatibilityAttributes();
        var ignorable = mc.Ignorable?.Value ?? string.Empty;
        if (!ignorable.Split(' ', StringSplitOptions.RemoveEmptyEntries).Contains("w14", StringComparer.Ordinal))
        {
            mc.Ignorable = ignorable.Length == 0 ? "w14" : ignorable + " w14";
        }
    }
}
