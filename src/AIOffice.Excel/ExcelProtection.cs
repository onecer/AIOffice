using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>One validated sheet-protection change, queued for the post-save pass.</summary>
internal sealed record SheetProtectionSpec(
    string SheetName,
    bool Protect,
    string? Password,
    IReadOnlyDictionary<string, bool> AllowFlags);

/// <summary>One validated workbook-structure-protection change, queued for the post-save pass.</summary>
internal sealed record WorkbookProtectionSpec(
    bool ProtectStructure,
    bool ProtectWindows,
    bool Clear,
    string? Password);

/// <summary>
/// Sheet, range and workbook protection (v1.2, additive). Per-cell <c>locked</c>
/// is ClosedXML-native (it round-trips cleanly); sheet and workbook protection
/// are authored on raw OpenXml in a post-save pass — ClosedXML cannot UNprotect a
/// password-protected sheet without the password, and this is documented LIGHT
/// protection that aioffice can always lift, so it owns the
/// <c>sheetProtection</c> / <c>workbookProtection</c> bytes directly. The
/// password is hashed with the ISO/OOXML SHA-512 + salt + spin-count scheme Excel
/// uses (matched exactly against ClosedXML's own output).
///
/// HONEST SCOPE: this is Excel's "light" protection — it controls which UI
/// actions Excel allows, not encryption. A protected sheet/workbook with a
/// password still has fully readable contents in the zip; the password gates
/// editing in Excel, it does NOT encrypt the file. Anyone (including aioffice)
/// can read or rewrite the bytes. For confidentiality use real encryption, which
/// AIOffice does not do. <c>get</c> reflects the protection state.
/// </summary>
internal static class ExcelProtection
{
    private const uint SpinCount = 100000;
    private const int SaltBytes = 16;

    /// <summary>Sheet-level set props this layer owns (range form owns only <c>locked</c>).</summary>
    public static readonly IReadOnlyList<string> SheetProps =
    [
        "protected", "password", "allowFormatCells", "allowFormatColumns", "allowFormatRows",
        "allowInsertColumns", "allowInsertRows", "allowDeleteColumns", "allowDeleteRows",
        "allowSort", "allowAutoFilter", "allowPivotTables", "allowSelectLockedCells", "allowSelectUnlockedCells",
    ];

    /// <summary>Workbook-root set props this layer owns.</summary>
    public static readonly IReadOnlyList<string> WorkbookProps = ["protectStructure", "protectWindows", "password"];

    /// <summary>
    /// Maps each <c>allow*</c> prop to the sheetProtection attribute it relaxes.
    /// In OOXML each attribute is <c>1</c> = LOCKED, <c>0</c>/absent = allowed; an
    /// allow* flag of true clears the lock. The defaults below are Excel's fresh
    /// protect (only cell selection allowed).
    /// </summary>
    private static readonly IReadOnlyList<(string Prop, string Attribute, bool DefaultLocked)> SheetAttributes =
    [
        ("allowFormatCells", "formatCells", true),
        ("allowFormatColumns", "formatColumns", true),
        ("allowFormatRows", "formatRows", true),
        ("allowInsertColumns", "insertColumns", true),
        ("allowInsertRows", "insertRows", true),
        ("allowDeleteColumns", "deleteColumns", true),
        ("allowDeleteRows", "deleteRows", true),
        ("allowSort", "sort", true),
        ("allowAutoFilter", "autoFilter", true),
        ("allowPivotTables", "pivotTables", true),
        ("allowSelectLockedCells", "selectLockedCells", false),
        ("allowSelectUnlockedCells", "selectUnlockedCells", false),
    ];

    // ----- per-cell locked ------------------------------------------------------

    /// <summary>
    /// Applies <c>{locked:true|false}</c> to a cell or range. A locked cell stays
    /// read-only under sheet protection; unlocking a range is how an agent leaves
    /// an editable window in an otherwise-protected sheet. (Cells default to
    /// locked, but the lock only bites once the sheet is protected.)
    /// </summary>
    public static void ApplyLocked(ExcelTarget target, bool locked, int index, List<string> applied)
    {
        var style = target.Kind switch
        {
            ExcelTargetKind.Cell => target.Cell!.Style,
            ExcelTargetKind.Range => target.Range!.Style,
            ExcelTargetKind.Row => target.Sheet.Row(target.RowNumber!.Value).Style,
            ExcelTargetKind.Column => target.Sheet.Column(target.ColumnNumber!.Value).Style,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: 'locked' targets a cell, range, row or col, not a {target.Kind.ToString().ToLowerInvariant()}.",
                "Unlock an editable window with {op:set, path:/Sheet1/A1:B10, props:{locked:false}}; " +
                "protect the sheet with {op:set, path:/Sheet1, props:{protected:true}}."),
        };

        style.Protection.Locked = locked;
        applied.Add("locked");
    }

    /// <summary>True when the cell's style marks it locked (the OOXML default is locked).</summary>
    public static bool IsLocked(IXLCell cell) => cell.Style.Protection.Locked;

    // ----- spec building (op-time) ----------------------------------------------

    /// <summary>
    /// Validates the sheet-protection props and builds a queued spec. The actual
    /// element write happens in the post-save pass.
    /// </summary>
    public static SheetProtectionSpec BuildSheetSpec(
        ExcelTarget target, bool protect, string? password, IReadOnlyDictionary<string, bool> flags, int index)
    {
        if (target.Kind != ExcelTargetKind.Sheet)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: sheet protection targets a sheet path like /Sheet1, not '{target.Kind.ToString().ToLowerInvariant()}'.",
                "Use {op:set, path:/Sheet1, props:{protected:true, allowFormatCells:true}}; " +
                "unlock individual cells with {op:set, path:/Sheet1/A1:B10, props:{locked:false}}.");
        }

        return new SheetProtectionSpec(target.Sheet.Name, protect, password, flags);
    }

    /// <summary>Validates and builds the workbook-structure-protection spec.</summary>
    public static WorkbookProtectionSpec BuildWorkbookSpec(
        bool? protectStructure, bool? protectWindows, string? password, int index)
    {
        if (protectStructure == false && protectWindows is null)
        {
            return new WorkbookProtectionSpec(false, false, Clear: true, Password: null);
        }

        if (protectStructure != true && protectWindows != true)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: workbook protection needs protectStructure (and/or protectWindows) set to true.",
                "Use {op:set, path:/, props:{protectStructure:true, password:\"secret\"}} to lock the sheet structure.");
        }

        return new WorkbookProtectionSpec(protectStructure == true, protectWindows == true, Clear: false, password);
    }

    // ----- post-save apply ------------------------------------------------------

    /// <summary>Writes queued sheet/workbook protection specs to the saved file.</summary>
    public static void Apply(string file, IReadOnlyList<SheetProtectionSpec> sheets, WorkbookProtectionSpec? workbook)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart ?? throw new AiofficeException(
            ErrorCodes.InternalError,
            "Protection write failed: the package has no workbook part.",
            "The edit was rolled back; re-run it. If this persists, report a bug with the workbook.");

        foreach (var spec in sheets)
        {
            ApplySheetRaw(workbookPart, spec);
        }

        if (workbook is not null)
        {
            ApplyWorkbookRaw(workbookPart, workbook);
        }
    }

    private static void ApplySheetRaw(WorkbookPart workbookPart, SheetProtectionSpec spec)
    {
        var sheetElement = workbookPart.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, spec.SheetName, StringComparison.OrdinalIgnoreCase));
        if (sheetElement?.Id?.Value is not { } relationshipId ||
            workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart ||
            worksheetPart.Worksheet is not { } worksheet)
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                $"Sheet '{spec.SheetName}' disappeared between validation and the protection write pass.",
                "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
        }

        // Always drop any existing protection first (unprotect is unconditional —
        // this is light protection aioffice can always lift).
        foreach (var existing in worksheet.Elements<S.SheetProtection>().ToList())
        {
            existing.Remove();
        }

        if (!spec.Protect)
        {
            worksheet.Save();
            return;
        }

        var protection = new S.SheetProtection { Sheet = true, Objects = true, Scenarios = true };
        foreach (var (prop, attribute, defaultLocked) in SheetAttributes)
        {
            var allowed = spec.AllowFlags.TryGetValue(prop, out var v) ? v : !defaultLocked;
            // Only emit the attribute when it differs from the OOXML default
            // (which is "locked"); an allowed action is attribute=false.
            SetSheetAttribute(protection, attribute, locked: !allowed);
        }

        if (!string.IsNullOrEmpty(spec.Password))
        {
            var salt = RandomNumberGenerator.GetBytes(SaltBytes);
            protection.AlgorithmName = "SHA-512";
            protection.SaltValue = Convert.ToBase64String(salt);
            protection.HashValue = HashPassword(spec.Password, salt, SpinCount);
            protection.SpinCount = SpinCount;
        }

        // Schema order: sheetProtection sits after sheetData and before any of
        // protectedRanges/autoFilter/.../drawing. Insert before the first known
        // later sibling; otherwise after sheetData.
        InsertSheetProtection(worksheet, protection);
        worksheet.Save();
    }

    private static void SetSheetAttribute(S.SheetProtection protection, string attribute, bool locked)
    {
        switch (attribute)
        {
            case "formatCells": protection.FormatCells = locked; break;
            case "formatColumns": protection.FormatColumns = locked; break;
            case "formatRows": protection.FormatRows = locked; break;
            case "insertColumns": protection.InsertColumns = locked; break;
            case "insertRows": protection.InsertRows = locked; break;
            case "deleteColumns": protection.DeleteColumns = locked; break;
            case "deleteRows": protection.DeleteRows = locked; break;
            case "sort": protection.Sort = locked; break;
            case "autoFilter": protection.AutoFilter = locked; break;
            case "pivotTables": protection.PivotTables = locked; break;
            case "selectLockedCells": protection.SelectLockedCells = locked; break;
            case "selectUnlockedCells": protection.SelectUnlockedCells = locked; break;
        }
    }

    private static void InsertSheetProtection(S.Worksheet worksheet, S.SheetProtection protection)
    {
        var sheetData = worksheet.Elements<S.SheetData>().FirstOrDefault();
        // The first element that legally FOLLOWS sheetProtection in CT_Worksheet.
        var successor = worksheet.ChildElements.FirstOrDefault(e =>
            e is S.ProtectedRanges or S.Scenarios or S.AutoFilter or S.SortState or S.DataConsolidate
                or S.CustomSheetViews or S.MergeCells or S.PhoneticProperties or S.ConditionalFormatting
                or S.DataValidations or S.Hyperlinks or S.PrintOptions or S.PageMargins or S.PageSetup
                or S.HeaderFooter or S.RowBreaks or S.ColumnBreaks or S.CustomProperties or S.CellWatches
                or S.IgnoredErrors or S.Drawing or S.LegacyDrawing or S.Picture or S.OleObjects or S.Controls
                or S.TableParts or S.ExtensionList);
        if (successor is not null)
        {
            worksheet.InsertBefore(protection, successor);
        }
        else if (sheetData is not null)
        {
            worksheet.InsertAfter(protection, sheetData);
        }
        else
        {
            worksheet.AppendChild(protection);
        }
    }

    private static void ApplyWorkbookRaw(WorkbookPart workbookPart, WorkbookProtectionSpec spec)
    {
        var workbook = workbookPart.Workbook!;
        foreach (var existing in workbook.Elements<S.WorkbookProtection>().ToList())
        {
            existing.Remove();
        }

        if (spec.Clear)
        {
            workbook.Save();
            return;
        }

        var protection = new S.WorkbookProtection();
        if (spec.ProtectStructure)
        {
            protection.LockStructure = true;
        }

        if (spec.ProtectWindows)
        {
            protection.LockWindows = true;
        }

        if (!string.IsNullOrEmpty(spec.Password))
        {
            var salt = RandomNumberGenerator.GetBytes(SaltBytes);
            protection.WorkbookAlgorithmName = "SHA-512";
            protection.WorkbookSaltValue = Convert.ToBase64String(salt);
            protection.WorkbookHashValue = HashPassword(spec.Password, salt, SpinCount);
            protection.WorkbookSpinCount = SpinCount;
        }

        // Schema order: workbookProtection sits after fileSharing/workbookPr and
        // before bookViews/sheets. Insert before the first known later sibling.
        var successor = workbook.ChildElements.FirstOrDefault(e =>
            e is S.BookViews or S.Sheets or S.FunctionGroups or S.ExternalReferences
                or S.DefinedNames or S.CalculationProperties or S.WorkbookExtensionList);
        if (successor is not null)
        {
            workbook.InsertBefore(protection, successor);
        }
        else
        {
            workbook.AppendChild(protection);
        }

        workbook.Save();
    }

    /// <summary>
    /// The ISO/OOXML SHA-512 password hash (salt prepended, then spin-count
    /// iterations folding in the little-endian iteration counter), matched
    /// byte-for-byte against ClosedXML's own output.
    /// </summary>
    internal static string HashPassword(string password, byte[] salt, uint spinCount)
    {
        var passwordBytes = Encoding.Unicode.GetBytes(password);
        using var sha = SHA512.Create();
        var hash = sha.ComputeHash([.. salt, .. passwordBytes]);
        for (uint i = 0; i < spinCount; i++)
        {
            hash = sha.ComputeHash([.. hash, .. BitConverter.GetBytes(i)]);
        }

        return Convert.ToBase64String(hash);
    }

    // ----- get reflection -------------------------------------------------------

    /// <summary>
    /// The protection block for a sheet get: whether it is protected, whether a
    /// password was set, and the allow* flags still permitted. Null (omitted)
    /// when the sheet is unprotected.
    /// </summary>
    public static object? SheetInfo(IXLWorksheet sheet)
    {
        var protection = sheet.Protection;
        if (!protection.IsProtected)
        {
            return null;
        }

        var allowed = protection.AllowedElements;
        return new
        {
            @protected = true,
            passwordProtected = protection.IsPasswordProtected,
            allowFormatCells = allowed.HasFlag(XLSheetProtectionElements.FormatCells),
            allowFormatColumns = allowed.HasFlag(XLSheetProtectionElements.FormatColumns),
            allowFormatRows = allowed.HasFlag(XLSheetProtectionElements.FormatRows),
            allowInsertColumns = allowed.HasFlag(XLSheetProtectionElements.InsertColumns),
            allowInsertRows = allowed.HasFlag(XLSheetProtectionElements.InsertRows),
            allowDeleteColumns = allowed.HasFlag(XLSheetProtectionElements.DeleteColumns),
            allowDeleteRows = allowed.HasFlag(XLSheetProtectionElements.DeleteRows),
            allowSort = allowed.HasFlag(XLSheetProtectionElements.Sort),
            allowAutoFilter = allowed.HasFlag(XLSheetProtectionElements.AutoFilter),
            allowPivotTables = allowed.HasFlag(XLSheetProtectionElements.PivotTables),
            allowSelectLockedCells = allowed.HasFlag(XLSheetProtectionElements.SelectLockedCells),
            allowSelectUnlockedCells = allowed.HasFlag(XLSheetProtectionElements.SelectUnlockedCells),
        };
    }

    /// <summary>
    /// The workbook structure-protection block for the workbook get. Null
    /// (omitted) when the workbook structure is not protected.
    /// </summary>
    public static object? WorkbookInfo(XLWorkbook workbook)
    {
        var protection = ((IXLWorkbook)workbook).Protection;
        if (!protection.IsProtected)
        {
            return null;
        }

        // AllowedElements names what stays editable; an aspect is PROTECTED when
        // it is absent from the allowed set.
        var allowed = protection.AllowedElements;
        return new
        {
            protectStructure = !allowed.HasFlag(XLWorkbookProtectionElements.Structure),
            protectWindows = !allowed.HasFlag(XLWorkbookProtectionElements.Windows),
            passwordProtected = protection.IsPasswordProtected,
        };
    }
}
