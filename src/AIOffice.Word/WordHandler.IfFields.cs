using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// Conditional merge fields (v1.4.0). An «IF» field is a Word complex field whose
/// instruction compares a record field against a value and shows one of two texts:
/// <c>IF Country = "US" "Domestic" "International"</c>. Like a MERGEFIELD it lives
/// between begin/separate/end FieldChars; its cached result is the branch that the
/// current data selects (the false branch until a merge supplies the field). The
/// <c>template</c> verb evaluates it from the same <c>--data</c> map / record that
/// fills {{key}} markers and MERGEFIELDs, <c>read --view fields</c> lists it and
/// <c>get</c> reports its condition.
/// </summary>
public sealed partial class WordHandler
{
    /// <summary>The comparison operators an IF field understands.</summary>
    private static readonly string[] IfFieldOperators = ["=", "!=", "<>", ">", "<", ">=", "<="];

    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[2]","type":"ifField","props":{"field":"Country","operator":"=","value":"US","trueText":"Domestic","falseText":"International"}}</c>:
    /// appends (or, with props.find, places at matched text) an «IF» complex field
    /// that resolves during template/merge from the named record field. The cached
    /// result is the branch the empty-merge state selects (the false branch), so the
    /// document reads sensibly before a merge fills it.
    /// </summary>
    private static object ApplyAddIfField(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        var author = session.ResolveAuthor(props);

        var field = props["field"] is { } fieldNode ? NodeToString(fieldNode).Trim() : null;
        if (string.IsNullOrEmpty(field))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type ifField needs props.field (the record field the condition tests).",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[2]\",\"type\":\"ifField\",\"props\":{\"field\":\"Country\",\"operator\":\"=\",\"value\":\"US\",\"trueText\":\"Domestic\",\"falseText\":\"International\"}}.");
        }

        // The field reference is a single Word merge-field token; quotes/spaces/
        // backslashes would corrupt the IF instruction.
        if (field.AsSpan().ContainsAnyExcept(MergeFieldNameChars))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.field '{field}' is not a valid merge field name.",
                "Use letters, digits and underscores only (e.g. \"Country\" or \"Order_Id\").");
        }

        var op2 = props["operator"] is { } opNode ? NodeToString(opNode).Trim() : "=";
        if (op2.Length == 0)
        {
            op2 = "=";
        }

        if (!IfFieldOperators.Contains(op2, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.operator '{op2}' is not a supported comparison.",
                "Use one of: = , != , > , < , >= , <= .",
                candidates: IfFieldOperators);
        }

        var value = props["value"] is { } valueNode ? NodeToString(valueNode) : string.Empty;
        var trueText = props["trueText"] is { } trueNode ? NodeToString(trueNode) : string.Empty;
        var falseText = props["falseText"] is { } falseNode ? NodeToString(falseNode) : string.Empty;

        // Quotes in the operands would break the instruction's quoting; reject them
        // with a clear suggestion rather than emitting a corrupt field.
        foreach (var (label, text) in new[] { ("value", value), ("trueText", trueText), ("falseText", falseText) })
        {
            if (text.Contains('"', StringComparison.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"props.{label} must not contain a double-quote character.",
                    "Remove the quotes; IF-field operands are quoted automatically.");
            }
        }

        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (anchor.Element is not Paragraph paragraph)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"IF fields are placed on paragraphs, not '{anchor.Type}'.",
                anchor.Type is "tc" or "header" or "footer"
                    ? $"Target the paragraph inside it: {anchor.CanonicalPath}/p[1]."
                    : "Pick a paragraph path, e.g. /body/p[2].");
        }

        var instruction = IfFieldInstruction(field, op2, value, trueText, falseText);

        // Empty-merge state: no data yet, so the false branch is what shows.
        var cached = falseText;

        var find = props["find"] is { } findNode ? NodeToString(findNode) : null;
        if (find is { Length: > 0 })
        {
            // Mid-paragraph placement splices runs into an existing paragraph — a
            // larger change than CT_Ins can wrap, so it stays tracked-unsupported.
            if (session.Track)
            {
                throw TrackedStructureUnsupported("ifField");
            }

            PlaceFieldAtText(
                paragraph,
                find,
                () => ComplexFieldRuns(instruction, cached),
                () => AppendComplexField(paragraph, instruction, cached));
        }
        else
        {
            // Remember the anchor's last pre-existing child so a tracked add wraps
            // exactly the five runs AppendComplexField is about to append.
            var lastExisting = paragraph.LastChild;
            AppendComplexField(paragraph, instruction, cached);
            if (session.Track)
            {
                WrapAppendedFieldRuns(doc, paragraph, lastExisting, author);
            }
        }

        return new
        {
            op = "add",
            type = "ifField",
            path = anchor.CanonicalPath,
            field,
            @operator = op2,
            value,
            trueText,
            falseText,
            cached,
            note = "Resolves during 'aioffice template' from the record's '" + field + "' field; Word evaluates it on field refresh too.",
        };
    }

    /// <summary>
    /// Builds the IF field's instruction text:
    /// <c>IF Field Operator "value" "trueText" "falseText"</c>. The field token is
    /// bare (a merge-field name); the value and branch texts are quoted.
    /// </summary>
    private static string IfFieldInstruction(string field, string op, string value, string trueText, string falseText) =>
        string.Create(
            CultureInfo.InvariantCulture,
            $" IF {field} {op} \"{value}\" \"{trueText}\" \"{falseText}\" ");

    // ------------------------------------------------------------------- read

    /// <summary>An IF field resolved from the document: its parsed condition and the run holding the instruction.</summary>
    internal sealed record IfFieldInfo(
        string Field, string Operator, string Value, string TrueText, string FalseText, string Result, FieldCode Code);

    /// <summary>Every «IF» complex field across body, headers and footers, in document order.</summary>
    private static List<IfFieldInfo> EnumerateIfFields(WordprocessingDocument doc)
    {
        var fields = new List<IfFieldInfo>();
        foreach (var root in MergeFieldRoots(doc))
        {
            foreach (var code in root.Descendants<FieldCode>())
            {
                if (ParseIfFieldInstruction(code.Text) is not { } parsed)
                {
                    continue;
                }

                fields.Add(new IfFieldInfo(
                    parsed.Field, parsed.Operator, parsed.Value, parsed.TrueText, parsed.FalseText,
                    ComplexFieldResult(code), code));
            }
        }

        return fields;
    }

    /// <summary>
    /// Parses an <c>IF Field Operator "value" "true" "false"</c> instruction into
    /// its parts, or null when the instruction is not an IF field of this shape.
    /// </summary>
    internal static (string Field, string Operator, string Value, string TrueText, string FalseText)?
        ParseIfFieldInstruction(string? instruction)
    {
        if (instruction is null)
        {
            return null;
        }

        var trimmed = instruction.Trim();
        if (!trimmed.StartsWith("IF ", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        // Tokenize: field and operator are bare; the three operands are "quoted".
        var rest = trimmed[3..].TrimStart();
        var fieldEnd = rest.IndexOf(' ', StringComparison.Ordinal);
        if (fieldEnd < 0)
        {
            return null;
        }

        var field = rest[..fieldEnd];
        rest = rest[(fieldEnd + 1)..].TrimStart();

        var opToken = IfFieldOperators.FirstOrDefault(o =>
            rest.StartsWith(o + " ", StringComparison.Ordinal) ||
            rest.StartsWith(o + "\"", StringComparison.Ordinal));
        if (opToken is null)
        {
            return null;
        }

        rest = rest[opToken.Length..].TrimStart();

        var quoted = ReadQuotedOperands(rest, 3);
        if (quoted is null)
        {
            return null;
        }

        return (field, opToken, quoted[0], quoted[1], quoted[2]);
    }

    /// <summary>Reads <paramref name="count"/> consecutive "quoted" tokens (space-separated), or null if too few.</summary>
    private static List<string>? ReadQuotedOperands(string text, int count)
    {
        var operands = new List<string>(count);
        var i = 0;
        while (operands.Count < count)
        {
            while (i < text.Length && text[i] == ' ')
            {
                i++;
            }

            if (i >= text.Length || text[i] != '"')
            {
                return null;
            }

            i++; // opening quote
            var start = i;
            while (i < text.Length && text[i] != '"')
            {
                i++;
            }

            if (i >= text.Length)
            {
                return null; // unterminated
            }

            operands.Add(text[start..i]);
            i++; // closing quote
        }

        return operands;
    }

    /// <summary>read --view fields shape for an IF field: {path, kind:"ifField", field, operator, value, trueText, falseText, value:result}.</summary>
    private static List<object> IfFieldsView(WordprocessingDocument doc)
    {
        var seen = new Dictionary<string, int>(StringComparer.Ordinal);
        var result = new List<object>();
        foreach (var field in EnumerateIfFields(doc))
        {
            seen.TryGetValue(field.Field, out var ordinal);
            ordinal++;
            seen[field.Field] = ordinal;
            result.Add(new
            {
                path = IfFieldPath(field.Field, ordinal),
                kind = "ifField",
                field = field.Field,
                @operator = field.Operator,
                condition = field.Value,
                trueText = field.TrueText,
                falseText = field.FalseText,
                value = field.Result,
            });
        }

        return result;
    }

    /// <summary>get on an IF-field path /ifField[@field=X] (or [@field=X][i]) → its condition and current branch.</summary>
    private static (string Path, Dictionary<string, object?> Properties) GetIfFieldProperties(
        WordprocessingDocument doc, string pathArg)
    {
        var (name, index) = ParseIfFieldPath(pathArg);
        var matches = EnumerateIfFields(doc).Where(f => f.Field == name).ToList();
        if (matches.Count == 0 || index > matches.Count)
        {
            var available = EnumerateIfFields(doc).Select(f => f.Field).Distinct(StringComparer.Ordinal).ToList();
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                matches.Count == 0
                    ? $"No IF field testing '{name}' exists."
                    : $"/ifField[@field={name}][{index}] does not exist; there are {matches.Count} with that field.",
                "Run 'aioffice read <file> --view fields' to list IF fields with their conditions.",
                candidates: [.. available.Take(5).Select(n => IfFieldPath(n, 1))]);
        }

        var field = matches[index - 1];
        return (IfFieldPath(name, index), new Dictionary<string, object?>
        {
            ["field"] = field.Field,
            ["operator"] = field.Operator,
            ["value"] = field.Value,
            ["trueText"] = field.TrueText,
            ["falseText"] = field.FalseText,
            ["result"] = field.Result,
        });
    }

    private static string IfFieldPath(string field, int ordinal) =>
        string.Create(CultureInfo.InvariantCulture, $"/ifField[@field={field}][{ordinal}]");

    private static (string Field, int Index) ParseIfFieldPath(string pathArg)
    {
        var match = System.Text.RegularExpressions.Regex.Match(
            pathArg.Trim(), @"^/?ifField\[@field=([A-Za-z0-9_]+)\](?:\[([0-9]+)\])?$");
        if (!match.Success)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"'{pathArg}' is not an IF-field path.",
                "Address an IF field as /ifField[@field=Country] (optionally [@field=Country][2] for repeats).");
        }

        var index = match.Groups[2].Success
            ? int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture)
            : 1;
        return (match.Groups[1].Value, index);
    }

    // ------------------------------------------------------------------- merge

    /// <summary>
    /// Evaluates every «IF» field against the merge data: the named field is looked
    /// up, the comparison run, and the cached result set to the winning branch.
    /// Fields whose tested name has no data are recorded as unresolved (the empty
    /// string compares as the field value). Returns the number of IF fields evaluated.
    /// </summary>
    private static int MergeIfFields(
        WordprocessingDocument doc,
        IReadOnlyDictionary<string, string> data,
        ISet<string> resolvedKeys,
        ISet<string> unresolvedKeys)
    {
        var evaluated = 0;
        foreach (var field in EnumerateIfFields(doc))
        {
            var present = data.TryGetValue(field.Field, out var left);
            var branch = EvaluateIfCondition(left ?? string.Empty, field.Operator, field.Value)
                ? field.TrueText
                : field.FalseText;

            if (SetComplexFieldResult(field.Code, branch))
            {
                evaluated++;
                if (present)
                {
                    resolvedKeys.Add(field.Field);
                }
                else
                {
                    unresolvedKeys.Add(field.Field);
                }
            }
        }

        return evaluated;
    }

    /// <summary>
    /// Compares <paramref name="left"/> (the record value) against
    /// <paramref name="right"/> with <paramref name="op"/>. Ordering comparisons use
    /// numeric semantics when both sides parse as numbers, else ordinal string order;
    /// equality is ordinal string comparison.
    /// </summary>
    private static bool EvaluateIfCondition(string left, string op, string right)
    {
        if (op is "=" )
        {
            return string.Equals(left, right, StringComparison.Ordinal);
        }

        if (op is "!=" or "<>")
        {
            return !string.Equals(left, right, StringComparison.Ordinal);
        }

        int comparison;
        if (double.TryParse(left, NumberStyles.Float, CultureInfo.InvariantCulture, out var l) &&
            double.TryParse(right, NumberStyles.Float, CultureInfo.InvariantCulture, out var r))
        {
            comparison = l.CompareTo(r);
        }
        else
        {
            comparison = string.CompareOrdinal(left, right);
        }

        return op switch
        {
            ">" => comparison > 0,
            "<" => comparison < 0,
            ">=" => comparison >= 0,
            "<=" => comparison <= 0,
            _ => false,
        };
    }
}
