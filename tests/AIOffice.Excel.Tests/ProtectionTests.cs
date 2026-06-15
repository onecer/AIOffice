using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The v1.2 protection slice (additive, ClosedXML-native, validator-clean):
/// per-cell <c>locked</c>, sheet protection (password hashed per OOXML), and
/// workbook structure protection. Raw assertions reopen the file with the OpenXml
/// SDK to read the actual sheetProtection / workbookProtection elements.
/// </summary>
public sealed class ProtectionTests : ExcelTestBase
{
    private static S.SheetProtection? RawSheetProtection(string file, string sheetName)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return worksheetPart.Worksheet!.Descendants<S.SheetProtection>().FirstOrDefault();
    }

    private static S.WorkbookProtection? RawWorkbookProtection(string file)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        return document.WorkbookPart!.Workbook!.Descendants<S.WorkbookProtection>().FirstOrDefault();
    }

    [Fact]
    public void Protect_sheet_with_password_writes_hashed_sheet_protection()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1", ("protected", true), ("password", "secret")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var protection = RawSheetProtection(file, "Sheet1");
        Assert.NotNull(protection);
        Assert.True(protection!.Sheet?.Value);
        // Password is stored as a SHA-512 salted hash, never as plaintext.
        Assert.NotNull(protection.HashValue);
        Assert.Equal("SHA-512", protection.AlgorithmName?.Value);
        Assert.NotNull(protection.SaltValue);

        AssertValidatorClean(file);
    }

    [Fact]
    public void Protect_sheet_without_password_is_still_protected()
    {
        var file = CreateWorkbook();

        EditOps(file, SetOp("/Sheet1", ("protected", true)));

        var protection = RawSheetProtection(file, "Sheet1");
        Assert.NotNull(protection);
        Assert.True(protection!.Sheet?.Value);
        // No password → no real hash (ClosedXML emits an empty hashValue, never
        // a populated one) and the protection is not password-protected.
        Assert.True(string.IsNullOrEmpty(protection.HashValue?.Value));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Unprotect_removes_sheet_protection()
    {
        var file = CreateWorkbook();
        EditOps(file, SetOp("/Sheet1", ("protected", true), ("password", "secret")));
        Assert.NotNull(RawSheetProtection(file, "Sheet1"));

        var envelope = EditOps(file, SetOp("/Sheet1", ("protected", false)));
        Assert.True(envelope.IsOk, envelope.ToJson());

        Assert.Null(RawSheetProtection(file, "Sheet1"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Allow_flags_relax_the_protection()
    {
        var file = CreateWorkbook();

        EditOps(file, SetOp("/Sheet1",
            ("protected", true), ("allowFormatCells", true), ("allowInsertRows", true)));

        // In the OOXML sheetProtection, an attribute of "0" (or absent default 0)
        // means the action is ALLOWED; "1" means it is locked. So an allowed
        // action must not be locked.
        var protection = RawSheetProtection(file, "Sheet1");
        Assert.NotNull(protection);
        Assert.NotEqual(true, protection!.FormatCells?.Value); // formatting cells is allowed
        Assert.NotEqual(true, protection.InsertRows?.Value);   // inserting rows is allowed
        // A flag we did not allow stays locked (true).
        Assert.True(protection.DeleteRows?.Value ?? true);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Unlock_range_then_protect_leaves_an_editable_window()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1:B10", ("locked", false)),
            SetOp("/Sheet1", ("protected", true), ("password", "secret")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        // Reopen-verify: the unlocked range stays editable under protection, while
        // a cell outside it keeps the default lock.
        using (var workbook = new XLWorkbook(file))
        {
            var sheet = workbook.Worksheet("Sheet1");
            Assert.False(sheet.Cell("A1").Style.Protection.Locked);
            Assert.False(sheet.Cell("B10").Style.Protection.Locked);
            Assert.True(sheet.Cell("C1").Style.Protection.Locked);
            Assert.True(sheet.Protection.IsProtected);
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Sheet_get_reflects_protection_state()
    {
        var file = CreateWorkbook();
        EditOps(file, SetOp("/Sheet1", ("protected", true), ("password", "pw"), ("allowFormatCells", true)));

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        var protection = data["protection"]!;

        Assert.True(protection["protected"]!.GetValue<bool>());
        Assert.True(protection["passwordProtected"]!.GetValue<bool>());
        Assert.True(protection["allowFormatCells"]!.GetValue<bool>());
        Assert.False(protection["allowDeleteRows"]!.GetValue<bool>());
    }

    [Fact]
    public void Unprotected_sheet_get_omits_the_protection_block()
    {
        var file = CreateWorkbook();

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        // Null-omitted on the wire — the protection key is absent.
        Assert.Null(data["protection"]);
    }

    [Fact]
    public void Cell_get_surfaces_an_explicitly_unlocked_cell()
    {
        var file = CreateWorkbook();
        EditOps(file, SetOp("/Sheet1/A1", ("locked", false)));

        var unlocked = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        Assert.False(unlocked["locked"]!.GetValue<bool>());

        // A default (locked) cell omits the key — only deviations are reported.
        var defaultCell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))));
        Assert.Null(defaultCell["locked"]);
    }

    [Fact]
    public void Protect_workbook_structure_writes_workbook_protection()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/", ("protectStructure", true), ("password", "wb")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var protection = RawWorkbookProtection(file);
        Assert.NotNull(protection);
        Assert.True(protection!.LockStructure?.Value);
        Assert.NotNull(protection.WorkbookHashValue);
        Assert.Equal("SHA-512", protection.WorkbookAlgorithmName?.Value);

        AssertValidatorClean(file);
    }

    [Fact]
    public void Workbook_structure_protection_reflected_and_liftable()
    {
        var file = CreateWorkbook();
        EditOps(file, SetOp("/", ("protectStructure", true)));

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        Assert.True(data["workbookProtection"]!["protectStructure"]!.GetValue<bool>());

        var lift = EditOps(file, SetOp("/", ("protectStructure", false)));
        Assert.True(lift.IsOk, lift.ToJson());
        Assert.Null(RawWorkbookProtection(file));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Allow_props_without_protected_flag_is_invalid_args()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1", ("allowFormatCells", true)));
        Assert.False(envelope.IsOk);
        var error = Json(envelope)["error"]!;
        Assert.Equal(ErrorCodes.InvalidArgs, error["code"]!.GetValue<string>());
        Assert.False(string.IsNullOrEmpty(error["suggestion"]!.GetValue<string>()));
    }

    [Fact]
    public void Locked_on_a_sheet_target_is_invalid_args()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1", ("locked", false)));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, Json(envelope)["error"]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void Protection_survives_a_later_unrelated_edit()
    {
        var file = CreateWorkbook();
        EditOps(file, SetOp("/Sheet1", ("protected", true), ("password", "pw")));

        // A later value edit must not disturb the protection (ClosedXML preserves it).
        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("value", "hello")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        Assert.NotNull(RawSheetProtection(file, "Sheet1"));
        AssertValidatorClean(file);
    }
}
