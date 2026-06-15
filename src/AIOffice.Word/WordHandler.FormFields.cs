using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// v1.3.0 legacy form fields — the classic FORMTEXT / FORMCHECKBOX / FORMDROPDOWN
/// fields (a <c>w:fldChar</c> complex field carrying <c>w:ffData</c>) that work
/// with Word's legacy "Restrict Editing → filling in forms" protection. They are
/// distinct from M7 content controls (<c>w:sdt</c>) and v1.2 mail-merge fields
/// (<c>MERGEFIELD</c>). Added on a paragraph, addressed by
/// <c>/formField[@name=clientName]</c>, value set with <c>set</c>, listed by
/// <c>read --view fields</c> (kind:formField), and removed.
/// </summary>
public sealed partial class WordHandler
{
    private static readonly string[] FormFieldKinds = ["text", "checkbox", "dropdown"];

    private static readonly System.Buffers.SearchValues<char> FormFieldNameChars =
        System.Buffers.SearchValues.Create(
            "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_");

    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[2]","type":"formField","props":{"kind":"text","name":"clientName","default"?,"items"?,"checked"?,"maxLength"?}}</c>:
    /// appends a legacy form field (begin fldChar+ffData / instr / separate /
    /// result / end) to the target paragraph.
    /// </summary>
    private static object ApplyAddFormField(WordprocessingDocument doc, EditOp op)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        props.Remove("author");

        var kind = props["kind"] is { } kindNode ? NodeToString(kindNode) : "text";
        if (!FormFieldKinds.Contains(kind, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Form-field kind '{kind}' is not supported.",
                "Use kind text, checkbox or dropdown.",
                candidates: FormFieldKinds);
        }

        var name = props["name"] is { } nameNode ? NodeToString(nameNode).Trim() : null;
        if (string.IsNullOrEmpty(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type formField needs props.name (the field's identifier).",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[2]\",\"type\":\"formField\",\"props\":{\"kind\":\"text\",\"name\":\"clientName\"}}.");
        }

        // A bookmark/form-field name is a single token: letters, digits, underscores.
        if (name.AsSpan().ContainsAnyExcept(FormFieldNameChars))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.name '{name}' is not a valid form-field name.",
                "Use letters, digits and underscores only (e.g. \"clientName\" or \"order_id\").");
        }

        if (EnumerateFormFields(doc).Any(f => string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase)))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"A form field named '{name}' already exists.",
                "Names are unique. Remove the old one ({\"op\":\"remove\",\"path\":\"" + FormFieldPath(name) +
                "\"}) or pick a different name.");
        }

        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (anchor.Element is not Paragraph paragraph)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Form fields are placed on paragraphs, not '{anchor.Type}'.",
                anchor.Type is "tc" or "header" or "footer"
                    ? $"Target the paragraph inside it: {anchor.CanonicalPath}/p[1]."
                    : "Pick a paragraph path, e.g. /body/p[2].");
        }

        var (ffData, instruction, result) = BuildFormFieldData(doc, kind, name, props);
        AppendFormField(paragraph, ffData, instruction, result);

        return new
        {
            op = "add",
            type = "formField",
            path = FormFieldPath(name),
            kind,
            name,
        };
    }

    /// <summary>Builds the w:ffData for one kind plus the matching field instruction and cached result text.</summary>
    private static (FormFieldData FfData, string Instruction, string Result) BuildFormFieldData(
        WordprocessingDocument doc, string kind, string name, JsonObject props)
    {
        var ffData = new FormFieldData(new FormFieldName { Val = name });
        ffData.AppendChild(new Enabled());
        ffData.AppendChild(new CalculateOnExit { Val = false });

        switch (kind)
        {
            case "checkbox":
            {
                var isChecked = props["checked"] is { } c && WordFormatting.ParseBool("checked", NodeToString(c));
                var checkbox = new CheckBox(
                    new FormFieldSize { Val = "20" },
                    new DefaultCheckBoxFormFieldState { Val = isChecked },
                    new Checked { Val = isChecked });
                ffData.AppendChild(checkbox);
                return (ffData, " FORMCHECKBOX ", string.Empty);
            }

            case "dropdown":
            {
                var items = props["items"] as JsonArray
                    ?? throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        "A dropdown form field needs props.items (the choices).",
                        "Example: {\"kind\":\"dropdown\",\"name\":\"region\",\"items\":[\"North\",\"South\"]}.");
                if (items.Count == 0)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        "A dropdown form field needs at least one item in props.items.",
                        "Pass props.items like [\"A\",\"B\",\"C\"].");
                }

                var ddl = new DropDownListFormField();
                var labels = items.Select(NodeToString).ToList();

                // Schema order: the selection (w:result) precedes the list entries.
                var selected = props["default"] is { } d ? NodeToString(d) : null;
                var selectedIndex = selected is { Length: > 0 } ? labels.IndexOf(selected) : 0;
                if (selectedIndex < 0)
                {
                    selectedIndex = 0;
                }

                if (selectedIndex > 0)
                {
                    ddl.AppendChild(new DropDownListSelection { Val = selectedIndex });
                }

                foreach (var label in labels)
                {
                    ddl.AppendChild(new ListEntryFormField { Val = label });
                }

                ffData.AppendChild(ddl);
                return (ffData, " FORMDROPDOWN ", labels[selectedIndex]);
            }

            default: // text
            {
                var def = props["default"] is { } dn ? NodeToString(dn) : null;
                var textInput = new TextInput();
                textInput.AppendChild(new TextBoxFormFieldType { Val = TextBoxFormFieldValues.Regular });
                if (def is { Length: > 0 })
                {
                    textInput.AppendChild(new DefaultTextBoxFormFieldString { Val = def });
                }

                if (props["maxLength"] is { } ml && short.TryParse(NodeToString(ml), NumberStyles.Integer, CultureInfo.InvariantCulture, out var max) && max > 0)
                {
                    textInput.AppendChild(new MaxLength { Val = max });
                }

                ffData.AppendChild(textInput);
                _ = doc;
                return (ffData, " FORMTEXT ", def ?? string.Empty);
            }
        }
    }

    /// <summary>Appends the five-run complex field (begin+ffData / instr / separate / result / end) to a paragraph.</summary>
    private static void AppendFormField(Paragraph paragraph, FormFieldData ffData, string instruction, string result)
    {
        paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin, FormFieldData = ffData }));
        paragraph.AppendChild(new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }));
        paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
        paragraph.AppendChild(new Run(NewText(result)));
        paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
    }

    // ------------------------------------------------------------------- set

    /// <summary>
    /// <c>{"op":"set","path":"/formField[@name=clientName]","props":{"value":"Acme"}}</c>
    /// (text/dropdown) or <c>{"checked":true}</c> (checkbox): updates the field's value.
    /// </summary>
    private static object ApplySetFormField(WordprocessingDocument doc, EditOp op)
    {
        var field = ResolveFormField(doc, op.Path);
        var props = RequireProps(op);
        var kind = field.Kind;

        if (kind == "checkbox")
        {
            var checkedNode = props["checked"] ?? props["value"] ?? throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "set on a checkbox form field needs props.checked (true/false).",
                "Example: {\"op\":\"set\",\"path\":\"" + FormFieldPath(field.Name) + "\",\"props\":{\"checked\":true}}.");
            var isChecked = WordFormatting.ParseBool("checked", NodeToString(checkedNode));
            var checkbox = field.FfData.GetFirstChild<CheckBox>()!;
            (checkbox.GetFirstChild<Checked>() ?? checkbox.AppendChild(new Checked())).Val = isChecked;
            return new { op = "set", path = FormFieldPath(field.Name), kind, @checked = isChecked };
        }

        var valueNode = props["value"] ?? props["text"] ?? props["default"] ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"set on a {kind} form field needs props.value.",
            "Example: {\"op\":\"set\",\"path\":\"" + FormFieldPath(field.Name) + "\",\"props\":{\"value\":\"Acme\"}}.");
        var value = NodeToString(valueNode);

        if (kind == "dropdown")
        {
            var items = DropdownEntries(field.FfData);
            var idx = items.IndexOf(value);
            if (idx < 0)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"'{value}' is not one of the dropdown's items.",
                    "Pick one of: " + string.Join(", ", items) + ".",
                    candidates: items);
            }

            var ddl = field.FfData.GetFirstChild<DropDownListFormField>()!;
            var selection = ddl.GetFirstChild<DropDownListSelection>();
            if (idx == 0)
            {
                selection?.Remove();
            }
            else if (selection is not null)
            {
                selection.Val = idx;
            }
            else
            {
                // The selection (w:result) must precede the list entries.
                ddl.InsertAt(new DropDownListSelection { Val = idx }, 0);
            }
        }

        SetFormFieldResult(field, value);
        return new { op = "set", path = FormFieldPath(field.Name), kind, value };
    }

    /// <summary>Replaces the cached result run (between separate and end) of a legacy form field with new text.</summary>
    private static void SetFormFieldResult(FormFieldRef field, string value)
    {
        var runs = field.Runs;
        var separateIndex = runs.FindIndex(r => r.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Separate);
        var endIndex = runs.FindIndex(r => r.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.End);
        if (separateIndex < 0 || endIndex < 0 || endIndex <= separateIndex)
        {
            return;
        }

        for (var i = endIndex - 1; i > separateIndex; i--)
        {
            runs[i].Remove();
        }

        runs[separateIndex].InsertAfterSelf(new Run(NewText(value)));
    }

    // ---------------------------------------------------------------- remove

    /// <summary>remove /formField[@name=X]: drops the field's run chain entirely.</summary>
    private static object ApplyRemoveFormField(WordprocessingDocument doc, EditOp op)
    {
        var field = ResolveFormField(doc, op.Path);
        foreach (var run in field.Runs)
        {
            run.Remove();
        }

        return new { op = "remove", path = FormFieldPath(field.Name), type = "formField" };
    }

    // ------------------------------------------------------------------- read

    /// <summary>A resolved legacy form field: its name, the begin-carrying ffData, kind and the full run chain.</summary>
    private sealed record FormFieldRef(string Name, string Kind, FormFieldData FfData, List<Run> Runs);

    /// <summary>Every legacy form field across body, headers and footers, in document order.</summary>
    private static List<FormFieldRef> EnumerateFormFields(WordprocessingDocument doc)
    {
        var result = new List<FormFieldRef>();
        foreach (var root in MergeFieldRoots(doc))
        {
            foreach (var paragraph in root.Descendants<Paragraph>())
            {
                CollectFormFields(paragraph, result);
            }
        }

        return result;
    }

    /// <summary>Pairs each ffData-bearing begin run with the matching separate/end runs in its paragraph.</summary>
    private static void CollectFormFields(Paragraph paragraph, List<FormFieldRef> sink)
    {
        var runs = paragraph.Elements<Run>().ToList();
        for (var i = 0; i < runs.Count; i++)
        {
            var begin = runs[i].GetFirstChild<FieldChar>();
            if (begin?.FieldCharType?.Value != FieldCharValues.Begin || begin.FormFieldData is not { } ffData)
            {
                continue;
            }

            var name = ffData.GetFirstChild<FormFieldName>()?.Val?.Value;
            if (string.IsNullOrEmpty(name))
            {
                continue;
            }

            var chain = new List<Run> { runs[i] };
            var j = i + 1;
            for (; j < runs.Count; j++)
            {
                chain.Add(runs[j]);
                if (runs[j].GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.End)
                {
                    break;
                }
            }

            sink.Add(new FormFieldRef(name, FormFieldKind(ffData), ffData, chain));
            i = j;
        }
    }

    private static string FormFieldKind(FormFieldData ffData) =>
        ffData.GetFirstChild<CheckBox>() is not null ? "checkbox"
        : ffData.GetFirstChild<DropDownListFormField>() is not null ? "dropdown"
        : "text";

    private static List<string> DropdownEntries(FormFieldData ffData) =>
        [.. ffData.GetFirstChild<DropDownListFormField>()?.Elements<ListEntryFormField>()
            .Select(e => e.Val?.Value ?? string.Empty) ?? []];

    /// <summary>read --view fields shape for a legacy form field: {path, kind:"formField", fieldKind, name, value}.</summary>
    private static List<object> FormFieldsView(WordprocessingDocument doc) =>
        [.. EnumerateFormFields(doc).Select(f => (object)new
        {
            path = FormFieldPath(f.Name),
            kind = "formField",
            fieldKind = f.Kind,
            name = f.Name,
            value = FormFieldValue(f),
        })];

    /// <summary>The current value: checkbox checked-state, otherwise the cached result text.</summary>
    private static object? FormFieldValue(FormFieldRef field)
    {
        if (field.Kind == "checkbox")
        {
            return field.FfData.GetFirstChild<CheckBox>()?.GetFirstChild<Checked>()?.Val?.Value ?? false;
        }

        var separateIndex = field.Runs.FindIndex(r => r.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Separate);
        var endIndex = field.Runs.FindIndex(r => r.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.End);
        if (separateIndex < 0 || endIndex <= separateIndex)
        {
            return string.Empty;
        }

        return string.Concat(field.Runs.Skip(separateIndex + 1).Take(endIndex - separateIndex - 1).Select(r => r.InnerText));
    }

    /// <summary>get /formField[@name=X] -> {name, fieldKind, value, items?}.</summary>
    private static (string Path, Dictionary<string, object?> Properties) GetFormFieldProperties(
        WordprocessingDocument doc, string pathArg)
    {
        var field = ResolveFormField(doc, pathArg);
        var properties = new Dictionary<string, object?>
        {
            ["name"] = field.Name,
            ["fieldKind"] = field.Kind,
            ["value"] = FormFieldValue(field),
        };

        if (field.Kind == "dropdown")
        {
            properties["items"] = DropdownEntries(field.FfData);
        }

        return (FormFieldPath(field.Name), properties);
    }

    // ------------------------------------------------------------ addressing

    /// <summary>Resolves /formField[@name=X] (or positional /formField[i]) to its run chain, or invalid_path.</summary>
    private static FormFieldRef ResolveFormField(WordprocessingDocument doc, string pathArg)
    {
        var fields = EnumerateFormFields(doc);
        var path = DocPath.Parse(pathArg);
        if (path.Segments.Count != 1 || path.Segments[0].Name != "formField")
        {
            throw FormFieldNotFound($"'{pathArg}' is not a form-field path.", fields);
        }

        var segment = path.Segments[0];
        if (segment.Id is { } wanted)
        {
            if (segment.IdAttribute is not (null or "name"))
            {
                throw FormFieldNotFound($"Form fields are addressed by their name, not @{segment.IdAttribute}.", fields);
            }

            return fields.FirstOrDefault(f => string.Equals(f.Name, wanted, StringComparison.OrdinalIgnoreCase))
                ?? throw FormFieldNotFound($"No form field named '{wanted}' exists.", fields);
        }

        var index = segment.Index ?? 1;
        if (index > fields.Count)
        {
            throw FormFieldNotFound(
                fields.Count == 0
                    ? "This document has no form fields."
                    : $"/formField[{index}] does not exist; there are {fields.Count}.",
                fields);
        }

        return fields[index - 1];
    }

    private static AiofficeException FormFieldNotFound(string message, List<FormFieldRef> fields) => new(
        ErrorCodes.InvalidPath,
        message,
        "Run 'aioffice read <file> --view fields' to list form fields with their names.",
        candidates: [.. fields.Take(5).Select(f => FormFieldPath(f.Name))]);

    private static string FormFieldPath(string name) => $"/formField[@name={name}]";
}
