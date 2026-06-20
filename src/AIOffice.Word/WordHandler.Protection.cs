using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// 1.13 document-level protection on the document root (<c>set /</c> / <c>get /</c>).
/// These are ENFORCEMENT-FLAG protections written into <c>word/settings.xml</c>,
/// not password encryption (CONTRACT §10 keeps strong encryption out of scope):
/// <list type="bullet">
///   <item><c>w:documentProtection</c> with <c>@w:edit</c> + <c>@w:enforcement="1"</c>
///   restricts editing to a mode (readOnly / comments / trackedChanges / forms).</item>
///   <item><c>w:writeProtection</c> with <c>@w:recommended="1"</c> makes Word open the
///   document read-only-recommended.</item>
/// </list>
/// Mirrors how xlsx <c>set /</c> carries workbook protection/calc props onto the
/// document root. The settings part is created on demand.
/// </summary>
public sealed partial class WordHandler
{
    /// <summary>The edit-mode spellings accepted by props.protection.edit -> their w:edit enum.</summary>
    private static readonly IReadOnlyDictionary<string, DocumentProtectionValues> ProtectionEditModes =
        new Dictionary<string, DocumentProtectionValues>(StringComparer.Ordinal)
        {
            ["readOnly"] = DocumentProtectionValues.ReadOnly,
            ["comments"] = DocumentProtectionValues.Comments,
            ["trackedChanges"] = DocumentProtectionValues.TrackedChanges,
            ["forms"] = DocumentProtectionValues.Forms,
            ["none"] = DocumentProtectionValues.None,
        };

    private static readonly IReadOnlyList<string> ProtectionRootProps = ["protection", "readOnlyRecommended"];

    // ------------------------------------------------------------------- set

    /// <summary>
    /// <c>{"op":"set","path":"/","props":{"protection":{"edit":"readOnly"},"readOnlyRecommended":true}}</c>:
    /// writes document-level enforcement-flag protection into settings.xml.
    /// </summary>
    private static object ApplySetDocumentProtection(WordprocessingDocument doc, EditOp op)
    {
        var props = RequireProps(op);
        foreach (var (key, _) in props)
        {
            if (!ProtectionRootProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown document-root prop '{key}'.",
                    "The document root ('/') supports document-level protection: "
                    + string.Join(", ", ProtectionRootProps)
                    + ". Example: {op:set, path:/, props:{protection:{edit:\"readOnly\"}}}.",
                    candidates: ProtectionRootProps);
            }
        }

        var settings = EnsureSettingsRoot(doc);
        var applied = new List<string>();

        if (props.TryGetPropertyValue("protection", out var protNode))
        {
            applied.Add(ApplyDocumentProtection(settings, protNode));
        }

        if (props.TryGetPropertyValue("readOnlyRecommended", out var recNode))
        {
            applied.Add(ApplyReadOnlyRecommended(settings, recNode));
        }

        return new { op = "set", path = "/", type = "documentProtection", applied };
    }

    /// <summary>Writes (or clears) w:documentProtection from the protection prop object.</summary>
    private static string ApplyDocumentProtection(Settings settings, JsonNode? protNode)
    {
        if (protNode is not JsonObject prot)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "protection must be an object, e.g. {edit:\"readOnly\", enforce:true}.",
                "Set the edit mode (readOnly | comments | trackedChanges | forms | none); pass enforce:false to lift it.");
        }

        var editRaw = prot.TryGetPropertyValue("edit", out var editNode) ? NodeToString(editNode) : null;
        if (string.IsNullOrEmpty(editRaw))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "protection needs an 'edit' mode.",
                "Use edit: readOnly | comments | trackedChanges | forms | none.",
                candidates: [.. ProtectionEditModes.Keys]);
        }

        if (!ProtectionEditModes.TryGetValue(editRaw, out var editValue))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown protection edit mode '{editRaw}'.",
                "Use edit: readOnly | comments | trackedChanges | forms | none.",
                candidates: [.. ProtectionEditModes.Keys]);
        }

        // enforce defaults to true; edit:"none" (or enforce:false) lifts protection.
        var enforce = !prot.TryGetPropertyValue("enforce", out var enforceNode) || enforceNode is null || IsTrue(enforceNode);

        // A password buys nothing here: this is flag enforcement, not encryption.
        // Accept and ignore it (CONTRACT §10) rather than erroring, so an agent that
        // habitually passes one still gets the intended enforcement flag.

        var existing = settings.GetFirstChild<DocumentProtection>();
        if (editValue == DocumentProtectionValues.None || !enforce)
        {
            existing?.Remove();
            return "protectionRemoved";
        }

        if (existing is null)
        {
            existing = new DocumentProtection();
            InsertSettingByRank(settings, existing);
        }

        existing.Edit = editValue;
        existing.Enforcement = OnOffValue.FromBoolean(true);
        return "protection:" + editRaw;
    }

    /// <summary>Writes (or clears) w:writeProtection @w:recommended from the readOnlyRecommended prop.</summary>
    private static string ApplyReadOnlyRecommended(Settings settings, JsonNode? recNode)
    {
        var recommended = recNode is not null && IsTrue(recNode);
        var existing = settings.GetFirstChild<WriteProtection>();

        if (!recommended)
        {
            existing?.Remove();
            return "readOnlyRecommendedRemoved";
        }

        if (existing is null)
        {
            existing = new WriteProtection();
            InsertSettingByRank(settings, existing);
        }

        existing.Recommended = OnOffValue.FromBoolean(true);
        return "readOnlyRecommended";
    }

    /// <summary>Inserts a settings child at its CT_Settings schema position (unknown children sort first).</summary>
    private static void InsertSettingByRank(Settings settings, OpenXmlElement child)
    {
        var rank = Array.IndexOf(SettingsOrder, child.GetType());
        var before = settings.ChildElements.FirstOrDefault(existing =>
        {
            var existingRank = Array.IndexOf(SettingsOrder, existing.GetType());
            return existingRank > rank; // unknown (-1) children sort first and never push us back
        });

        if (before is null)
        {
            settings.AppendChild(child);
        }
        else
        {
            settings.InsertBefore(child, before);
        }
    }

    // ------------------------------------------------------------------- get

    /// <summary>
    /// The document-root protection shape reported by <c>get /</c>:
    /// <c>{protection:{edit, enforced}, readOnlyRecommended}</c>. With no protection
    /// set, edit is "none", enforced is false and readOnlyRecommended is false.
    /// </summary>
    private static Dictionary<string, object?> DocumentProtectionShape(WordprocessingDocument doc)
    {
        var settings = doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
        var prot = settings?.GetFirstChild<DocumentProtection>();
        var write = settings?.GetFirstChild<WriteProtection>();

        // The @w:edit wire spelling (readOnly / comments / trackedChanges / forms /
        // none) is exactly the prop spelling we accept, so InnerText round-trips.
        var edit = prot?.Edit is { } editVal && !string.IsNullOrEmpty(editVal.InnerText)
            ? editVal.InnerText
            : "none";
        var enforced = prot is not null
            && edit != "none"
            && (prot.Enforcement is null || OnOffValue.ToBoolean(prot.Enforcement));

        return new Dictionary<string, object?>
        {
            ["protection"] = new Dictionary<string, object?>
            {
                ["edit"] = enforced ? edit : "none",
                ["enforced"] = enforced,
            },
            ["readOnlyRecommended"] = write?.Recommended is { } rec && OnOffValue.ToBoolean(rec),
        };
    }
}
