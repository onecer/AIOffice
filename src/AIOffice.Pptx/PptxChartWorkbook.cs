using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Pptx;

/// <summary>
/// Builds the minimal real workbook embedded inside a ChartPart so PowerPoint's
/// "Edit Data" opens a live sheet instead of prompting to create one. Layout
/// matches what PowerPoint itself writes: Sheet1 with categories in column A
/// (rows 2..N+1), one series per column from B (name in row 1, values below) —
/// exactly the ranges the chart's reference caches point at.
/// </summary>
internal static class PptxChartWorkbook
{
    /// <summary>The content type of the embedded chart workbook package part.</summary>
    public const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /// <summary>The c:f range of the category labels: Sheet1!$A$2:$A$N+1.</summary>
    public static string CategoriesRange(int categoryCount) => Range("A", 2, categoryCount);

    /// <summary>The c:f single-cell reference of a series name (0-based series ordinal): Sheet1!$B$1, $C$1, …</summary>
    public static string SeriesNameReference(int seriesOrdinal) =>
        Units.Inv($"Sheet1!${ColumnName(seriesOrdinal + 2)}$1");

    /// <summary>The c:f range of a series' values (0-based series ordinal): Sheet1!$B$2:$B$N+1, …</summary>
    public static string ValuesRange(int seriesOrdinal, int valueCount) =>
        Range(ColumnName(seriesOrdinal + 2), 2, valueCount);

    private static string Range(string column, int firstRow, int count)
    {
        var lastRow = firstRow + Math.Max(count, 1) - 1;
        return lastRow == firstRow
            ? Units.Inv($"Sheet1!${column}${firstRow}")
            : Units.Inv($"Sheet1!${column}${firstRow}:${column}${lastRow}");
    }

    /// <summary>1-based spreadsheet column name (1 = A, 2 = B, 27 = AA …).</summary>
    private static string ColumnName(int index1)
    {
        var name = string.Empty;
        var n = index1;
        while (n > 0)
        {
            n--;
            name = (char)('A' + (n % 26)) + name;
            n /= 26;
        }

        return name;
    }

    /// <summary>The complete xlsx package bytes for one chart's data.</summary>
    public static byte[] Build(PptxChartData data)
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            var sheetData = new S.SheetData();
            var header = new S.Row { RowIndex = 1u };
            for (var s = 0; s < data.Series.Count; s++)
            {
                header.Append(TextCell(Units.Inv($"{ColumnName(s + 2)}1"), data.Series[s].Name));
            }

            sheetData.Append(header);

            for (var c = 0; c < data.Categories.Count; c++)
            {
                var rowIndex = (uint)(c + 2);
                var row = new S.Row { RowIndex = rowIndex };
                row.Append(TextCell(Units.Inv($"A{rowIndex}"), data.Categories[c]));
                for (var s = 0; s < data.Series.Count; s++)
                {
                    if (c < data.Series[s].Values.Count && data.Series[s].Values[c] is { } value)
                    {
                        row.Append(NumberCell(Units.Inv($"{ColumnName(s + 2)}{rowIndex}"), value));
                    }
                }

                sheetData.Append(row);
            }

            worksheetPart.Worksheet = new S.Worksheet(sheetData);
            workbookPart.Workbook = new S.Workbook(new S.Sheets(new S.Sheet
            {
                Name = "Sheet1",
                SheetId = 1u,
                Id = workbookPart.GetIdOfPart(worksheetPart),
            }));
        }

        return stream.ToArray();
    }

    private static S.Cell TextCell(string reference, string text) => new(new S.InlineString(new S.Text(text)))
    {
        CellReference = reference,
        DataType = S.CellValues.InlineString,
    };

    private static S.Cell NumberCell(string reference, double value) =>
        new(new S.CellValue(value.ToString(CultureInfo.InvariantCulture)))
        {
            CellReference = reference,
            DataType = S.CellValues.Number,
        };
}
