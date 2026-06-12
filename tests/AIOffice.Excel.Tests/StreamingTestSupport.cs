using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// Tests that mutate <c>AIOFFICE_MAX_FILE_MB</c> (process-wide) must not run
/// in parallel with anything else; this collection serializes them.
/// </summary>
[CollectionDefinition("FileSizeEnv", DisableParallelization = true)]
public sealed class FileSizeEnvCollection;

/// <summary>Scoped env-var override (restores the previous value on dispose).</summary>
internal sealed class EnvScope : IDisposable
{
    private readonly string _key;
    private readonly string? _previous;

    public EnvScope(string key, string? value)
    {
        _key = key;
        _previous = Environment.GetEnvironmentVariable(key);
        Environment.SetEnvironmentVariable(key, value);
    }

    public void Dispose() => Environment.SetEnvironmentVariable(_key, _previous);
}

/// <summary>
/// Deterministic workbook generator for the streaming tests, written with raw
/// part streams (NOT ClosedXML — a 300k-row DOM save would dominate the suite).
/// Every cell value is a pure function of its row, so tests can assert exact
/// values without keeping the data in memory:
/// <list type="bullet">
/// <item>A: the row number;</item>
/// <item>B: a pseudo-random double (high-entropy digits defeat zip compression);</item>
/// <item>C: <c>row * 1.5 + 0.25</c>;</item>
/// <item>D: shared string <c>Label_{row % 1000}</c> (exercises the sst path);</item>
/// <item>E: formula <c>A{row}*2</c> with cached value <c>2*row</c>;</item>
/// <item>F: a 110-char pseudo-random inline string (the size filler).</item>
/// </list>
/// At 310k rows the package lands around 39 MB — above the 20 MB streaming
/// threshold, below the old 50 MB guard default, so the suite stays green
/// whether or not the integrator has flipped the guard default yet.
/// </summary>
internal static class BigWorkbookGenerator
{
    public const int Columns = 6;
    public const int SharedStringPool = 1000;
    private const int NoiseLength = 110;

    public static string ExpectedSharedString(int row) =>
        "Label_" + (row % SharedStringPool).ToString(CultureInfo.InvariantCulture);

    public static double ExpectedA(int row) => row;

    public static double ExpectedB(int row) => SplitMix((ulong)row * 2654435761ul) / (double)ulong.MaxValue;

    public static double ExpectedC(int row) => (row * 1.5) + 0.25;

    public static double ExpectedCachedE(int row) => 2.0 * row;

    public static string ExpectedNoise(int row)
    {
        const string alphabet = "0123456789abcdefghijklmnopqrstuvwxyz";
        var sb = new StringBuilder(NoiseLength);
        var state = SplitMix(((ulong)row << 1) | 1ul);
        for (var i = 0; i < NoiseLength; i++)
        {
            state = SplitMix(state);
            sb.Append(alphabet[(int)(state % 36)]);
        }

        return sb.ToString();
    }

    public static void Generate(string path, int rows)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sharedStringsPart = workbookPart.AddNewPart<SharedStringTablePart>();

        using (var stream = sharedStringsPart.GetStream(FileMode.Create))
        using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            writer.Write(
                "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                $"count=\"{SharedStringPool}\" uniqueCount=\"{SharedStringPool}\">");
            for (var i = 0; i < SharedStringPool; i++)
            {
                writer.Write($"<si><t>Label_{i}</t></si>");
            }

            writer.Write("</sst>");
        }

        using (var stream = worksheetPart.GetStream(FileMode.Create))
        using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
        {
            var ci = CultureInfo.InvariantCulture;
            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            writer.Write($"<dimension ref=\"A1:F{rows}\"/>");
            writer.Write("<sheetData>");
            for (var r = 1; r <= rows; r++)
            {
                writer.Write($"<row r=\"{r}\">");
                writer.Write($"<c r=\"A{r}\"><v>{r}</v></c>");
                writer.Write($"<c r=\"B{r}\"><v>{ExpectedB(r).ToString("R", ci)}</v></c>");
                writer.Write($"<c r=\"C{r}\"><v>{ExpectedC(r).ToString("R", ci)}</v></c>");
                writer.Write($"<c r=\"D{r}\" t=\"s\"><v>{r % SharedStringPool}</v></c>");
                writer.Write($"<c r=\"E{r}\"><f>A{r}*2</f><v>{ExpectedCachedE(r).ToString("R", ci)}</v></c>");
                writer.Write($"<c r=\"F{r}\" t=\"inlineStr\"><is><t>{ExpectedNoise(r)}</t></is></c>");
                writer.Write("</row>");
            }

            writer.Write("</sheetData></worksheet>");
        }

        workbookPart.Workbook = new S.Workbook(new S.Sheets(new S.Sheet
        {
            Name = "Sheet1",
            SheetId = 1U,
            Id = workbookPart.GetIdOfPart(worksheetPart),
        }));
        workbookPart.Workbook.Save();
    }

    private static ulong SplitMix(ulong x)
    {
        x += 0x9E3779B97F4A7C15ul;
        x = (x ^ (x >> 30)) * 0xBF58476D1CE4E5B9ul;
        x = (x ^ (x >> 27)) * 0x94D049BB133111EBul;
        return x ^ (x >> 31);
    }
}

/// <summary>
/// One ~39 MB / 310k-row workbook shared by every large-file test (generated
/// once per run; xUnit class fixtures dispose it afterwards).
/// </summary>
public sealed class LargeWorkbookFixture : IDisposable
{
    public const int Rows = 310_000;

    public LargeWorkbookFixture()
    {
        Directory.CreateDirectory(Dir);
        File = Path.Combine(Dir, "large.xlsx");
        BigWorkbookGenerator.Generate(File, Rows);
        SizeBytes = new FileInfo(File).Length;
    }

    public string Dir { get; } =
        Path.Combine(Path.GetTempPath(), "aioffice-xlsx-large-" + Guid.NewGuid().ToString("N"));

    public string File { get; }

    public long SizeBytes { get; }

    public void Dispose()
    {
        if (Directory.Exists(Dir))
        {
            Directory.Delete(Dir, recursive: true);
        }
    }
}
