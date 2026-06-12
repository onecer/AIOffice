using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// The xlsx format handler, built on ClosedXML (OpenXml SDK for raw fix-ups).
///
/// Honest capability notes (v0):
/// <list type="bullet">
/// <item>Formula evaluation uses ClosedXML's engine. Verified working: SUM,
/// AVERAGE, IF, VLOOKUP, INDEX, MATCH, TEXT, DATE, COUNTIF, CONCATENATE,
/// TEXTJOIN. Verified NOT evaluated (saved without a cached value, Excel
/// computes on open, warning <c>formula_not_evaluated</c> raised): SEQUENCE,
/// XLOOKUP, LET and other dynamic-array/365 functions.</item>
/// <item>A no-edit open/save through ClosedXML rewrites
/// <c>xl/worksheets/sheetN.xml</c> (namespace attribute order on the root
/// element only); every other zip part stays byte-identical. The round-trip
/// law test pins this exact set.</item>
/// <item>Charts (bar | line | pie | scatter | area) are authored on raw
/// OpenXml in a post-save pass, because ClosedXML cannot create them.
/// Measured: ClosedXML 0.105 preserves existing chart/drawing parts
/// byte-identical across its own saves, so charts survive later edits. Other
/// chart kinds return <c>unsupported_feature</c> naming the supported set.</item>
/// <item>Pivot tables, conditional formats (cellIs | colorScale | dataBar |
/// containsText), images (png | jpeg, header-sniffed, sandbox-resolved),
/// defined names, freeze panes, autofilter and page setup are
/// ClosedXML-native; a post-save pass corrects ClosedXML's data-bar GUID
/// casing so files stay validator-clean (see ExcelConditionalFormats).</item>
/// <item>Big workbooks (M3): <c>read --view stats|text</c> and <c>get</c> of a
/// cell/range stream the raw XML without loading the DOM — see
/// <see cref="ExcelStreaming"/>. Mutating ops on huge files still load the
/// full workbook through ClosedXML, with ONE exception: bulk 2D writes over
/// 50k cells into a bare sheet stream through a SAX writer — see
/// <see cref="ExcelBulkWrites"/>. In-place streaming edits remain M5.</item>
/// <item>M4: find/replace (<see cref="ApplyReplace"/>; text cells + optional
/// formula text, regex with 2s timeout, zero matches = <c>find_no_match</c>
/// warning), bulk 2D writes (anchor + strict range forms), row/column
/// insert/delete/size/hide (<c>/Sheet1/col[C]</c> letter addressing; ClosedXML
/// shifts formula references and tests assert it), and cell notes (always
/// addressed via their cell; no <c>note[i]</c> index form).</item>
/// <item>M5 — the csv bridge and form/annotation surface:
/// <see cref="CreateFrom"/> imports csv (RFC 4180, sniffed delimiter, typed
/// values, leading-zero strings stay text) and <c>read --view csv</c> exports
/// one sheet window (formulas as cached values); data validations
/// (list/wholeNumber/decimal/date/textLength — <see cref="ExcelDataValidations"/>),
/// THREADED comments with replies + persons dedup + legacy note shadows
/// (<see cref="ExcelComments"/>, <c>read --view comments</c>), cell hyperlinks
/// (external/internal, render emits <c>&lt;a&gt;</c>) and sparklines
/// (line/column/winLoss in x14 extLst — <see cref="ExcelSparklines"/>).</item>
/// <item><c>move</c> ops and png rendering are not implemented; they return
/// typed <c>unsupported_feature</c> envelopes with workarounds.</item>
/// </list>
/// </summary>
public sealed partial class ExcelHandler : IFormatHandler
{
    private readonly SnapshotStore _snapshots;

    /// <param name="snapshots">Snapshot ring override (tests); defaults to <c>~/.aioffice/snapshots</c>.</param>
    public ExcelHandler(SnapshotStore? snapshots = null) => _snapshots = snapshots ?? new SnapshotStore();

    public DocumentKind Kind => DocumentKind.Xlsx;

    public Envelope Create(CommandContext ctx) => Run(ctx, sw =>
    {
        var file = RequireFile(ctx, mustExist: false);
        if (File.Exists(file))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"File already exists: {file}",
                "Pick a new file name, or use 'aioffice edit' to change the existing workbook.");
        }

        var title = ArgString(ctx, "title") ?? "Sheet1";
        using var workbook = new XLWorkbook();
        AddSheetOrThrow(workbook, title);

        var directory = Path.GetDirectoryName(file);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        workbook.SaveAs(file);
        return Envelope.Ok(
            new { file, kind = "xlsx", sheets = new[] { title } },
            MetaFor(file, sw));
    });

    // ----- shared plumbing -------------------------------------------------

    /// <summary>Runs a verb body, converting any exception into a failure envelope (never a crash).</summary>
    private static Envelope Run(CommandContext ctx, Func<Stopwatch, Envelope> body)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            return body(sw);
        }
        catch (Exception exception)
        {
            return Envelope.FromException(exception, MetaFor(ctx.File, sw));
        }
    }

    private static Meta MetaFor(string? file, Stopwatch sw, IReadOnlyList<Warning>? warnings = null) => new()
    {
        File = file,
        Rev = file is not null && File.Exists(file) ? Core.Rev.OfFile(file) : null,
        ElapsedMs = sw.ElapsedMilliseconds,
        Warnings = warnings is { Count: > 0 } ? warnings : null,
    };

    private static string RequireFile(CommandContext ctx, bool mustExist)
    {
        if (string.IsNullOrWhiteSpace(ctx.File))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "This command needs a target file.",
                "Pass the workbook path, e.g. aioffice read book.xlsx.");
        }

        if (mustExist && !File.Exists(ctx.File))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"File not found: {ctx.File}",
                "Check the path spelling, or run 'aioffice create' to make a new workbook.");
        }

        FileSizeGuard.Ensure(ctx.File); // file_too_large before any expensive open
        return ctx.File;
    }

    private static XLWorkbook OpenWorkbook(string file)
    {
        try
        {
            return new XLWorkbook(file);
        }
        catch (Exception exception)
        {
            throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                $"The file could not be opened as an xlsx workbook: {exception.Message}",
                "Verify the file is a real .xlsx (a zip container); re-export it from its source application.",
                innerException: exception);
        }
    }

    private static void AddSheetOrThrow(XLWorkbook workbook, string name)
    {
        try
        {
            workbook.AddWorksheet(name);
        }
        catch (ArgumentException exception)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{name}' is not a usable sheet name: {exception.Message}",
                @"Sheet names are 1-31 characters and cannot contain : \ / ? * [ ].",
                innerException: exception);
        }
    }

    private static string? ArgString(CommandContext ctx, string key) =>
        ctx.Args.TryGetPropertyValue(key, out var node) && node is JsonValue value
            ? value.GetValueKind() == JsonValueKind.String
                ? value.GetValue<string>()
                : value.ToJsonString(JsonDefaults.Options)
            : null;

    private static bool ArgBool(CommandContext ctx, string key)
    {
        if (!ctx.Args.TryGetPropertyValue(key, out var node) || node is not JsonValue value)
        {
            return false;
        }

        return value.GetValueKind() switch
        {
            JsonValueKind.True => true,
            JsonValueKind.String => string.Equals(value.GetValue<string>(), "true", StringComparison.OrdinalIgnoreCase),
            _ => false,
        };
    }

    private static int? ArgInt(CommandContext ctx, string key)
    {
        if (!ctx.Args.TryGetPropertyValue(key, out var node) || node is not JsonValue value)
        {
            return null;
        }

        return value.GetValueKind() switch
        {
            JsonValueKind.Number => value.GetValue<int>(),
            JsonValueKind.String when int.TryParse(
                value.GetValue<string>(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) => parsed,
            _ => null,
        };
    }
}
