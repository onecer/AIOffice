using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// Per-test temp sandbox plus the two oracles every mutating test must consult:
/// the OpenXml schema validator and raw-part inspection (reading the file back
/// with the OpenXml SDK, not with ClosedXML, so ClosedXML cannot grade its own
/// homework).
/// </summary>
public abstract class ExcelTestBase : IDisposable
{
    protected ExcelTestBase()
    {
        Dir = Path.Combine(Path.GetTempPath(), "aioffice-xlsx-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Dir);
        Workspace = new Workspace(Dir);
        Snapshots = new SnapshotStore(Path.Combine(Dir, ".snapshots"));
        Handler = new ExcelHandler(Snapshots);
    }

    protected string Dir { get; }

    protected Workspace Workspace { get; }

    protected SnapshotStore Snapshots { get; }

    protected ExcelHandler Handler { get; }

    public void Dispose()
    {
        if (Directory.Exists(Dir))
        {
            Directory.Delete(Dir, recursive: true);
        }

        GC.SuppressFinalize(this);
    }

    protected CommandContext Ctx(string file, params (string Key, JsonNode? Value)[] args)
    {
        var obj = new JsonObject();
        foreach (var (key, value) in args)
        {
            obj[key] = value;
        }

        return new CommandContext { Workspace = Workspace, File = file, Args = obj };
    }

    protected string NewFile(string name = "book.xlsx") => Path.Combine(Dir, name);

    /// <summary>Creates a workbook through the handler and asserts success.</summary>
    protected string CreateWorkbook(string name = "book.xlsx", string? title = null)
    {
        var file = NewFile(name);
        var envelope = title is null
            ? Handler.Create(Ctx(file))
            : Handler.Create(Ctx(file, ("title", title)));
        Assert.True(envelope.IsOk, envelope.ToJson());
        return file;
    }

    protected Envelope EditOps(string file, params EditOp[] ops) => Handler.Edit(Ctx(file), ops);

    protected static EditOp SetOp(string path, params (string Key, JsonNode? Value)[] props)
    {
        var obj = new JsonObject();
        foreach (var (key, value) in props)
        {
            obj[key] = value;
        }

        return new EditOp { Op = "set", Path = path, Props = obj };
    }

    protected static EditOp AddOp(string path, string type, params (string Key, JsonNode? Value)[] props)
    {
        var obj = new JsonObject();
        foreach (var (key, value) in props)
        {
            obj[key] = value;
        }

        return new EditOp { Op = "add", Path = path, Type = type, Props = obj.Count > 0 ? obj : null };
    }

    protected static EditOp RemoveOp(string path) => new() { Op = "remove", Path = path };

    /// <summary>Parses an envelope's canonical JSON wire form for assertions.</summary>
    protected static JsonNode Json(Envelope envelope) => JsonNode.Parse(envelope.ToJson())!;

    protected static JsonNode OkData(Envelope envelope)
    {
        Assert.True(envelope.IsOk, envelope.ToJson());
        return Json(envelope)["data"]!;
    }

    /// <summary>Oracle: the OpenXml schema validator must report zero errors.</summary>
    protected static void AssertValidatorClean(string file)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
        Assert.True(
            errors.Count == 0,
            "OpenXmlValidator errors: " + string.Join("; ", errors.Select(e => $"{e.Id}: {e.Description}")));
    }

    /// <summary>
    /// Oracle: reads one cell raw with the OpenXml SDK — formula text, cached
    /// value text (<c>&lt;v&gt;</c>) and data type, exactly as stored on disk.
    /// </summary>
    protected static (string? Formula, string? CachedValue, string? DataType) RawCell(
        string file, string sheetName, string address)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        var cell = worksheetPart.Worksheet!.Descendants<Cell>()
            .FirstOrDefault(c => c.CellReference?.Value == address);
        return cell is null
            ? (null, null, null)
            : (cell.CellFormula?.Text, cell.CellValue?.Text, cell.DataType?.Value.ToString());
    }
}
