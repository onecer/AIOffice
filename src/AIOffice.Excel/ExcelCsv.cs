using System.Globalization;
using System.Text;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// The M5 CSV bridge: a hand-rolled RFC 4180 parser (quoted fields, embedded
/// delimiters/newlines, <c>""</c> escapes, CRLF records, UTF-8 with BOM
/// tolerance) plus the RFC-4180-safe exporter behind <c>read --view csv</c>.
///
/// Documented choices:
/// <list type="bullet">
/// <item>Delimiter is sniffed among <c>,</c> <c>;</c> and tab (most hits
/// outside quotes in the first 16 lines; comma wins ties) and can be forced
/// with the <c>delimiter</c> prop.</item>
/// <item>Newlines INSIDE quoted fields are normalized to <c>\n</c> (what Excel
/// stores in cell text); record separators may be CRLF, LF or CR.</item>
/// <item>Field typing follows the bulk-write heuristic — see
/// <see cref="ExcelValues.ParseCsvField"/> (leading-zero strings stay text).</item>
/// <item>Blank lines import as empty rows (row alignment is preserved);
/// trailing blank records are dropped, so a final newline adds no row.</item>
/// <item>Export writes CRLF records and quotes a field only when it contains
/// the delimiter, a quote or a newline. Formula cells export their CACHED
/// VALUE (csv has no formulas); dates export as ISO 8601, booleans as
/// <c>true</c>/<c>false</c>, so a re-import types them back identically.
/// Round-trip exceptions (documented): error values and timespans come back
/// as text, and a number-looking string forced with <c>valueType:"text"</c>
/// re-imports as a number unless it has a leading zero.</item>
/// </list>
/// </summary>
internal static class ExcelCsv
{
    private static readonly IReadOnlyList<char> SniffCandidates = [',', ';', '\t'];

    public static readonly IReadOnlyList<string> DelimiterNames = [",", ";", "\\t (or 'tab')"];

    /// <summary>Parses a delimiter prop value; null input means "sniff".</summary>
    public static char? ParseDelimiterArg(string? arg)
    {
        switch (arg)
        {
            case null:
                return null;
            case ",":
            case "comma":
                return ',';
            case ";":
            case "semicolon":
                return ';';
            case "\t":
            case "\\t":
            case "tab":
                return '\t';
            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"'{arg}' is not a usable csv delimiter.",
                    "Pass \",\", \";\" or \"\\t\" (the word 'tab' also works); omit it to auto-detect.",
                    candidates: DelimiterNames);
        }
    }

    /// <summary>Most frequent candidate delimiter outside quotes in the first 16 lines (comma wins ties).</summary>
    public static char Sniff(string text)
    {
        var counts = new int[SniffCandidates.Count];
        var inQuotes = false;
        var lines = 0;
        foreach (var c in text)
        {
            if (inQuotes)
            {
                inQuotes = c != '"';
                continue;
            }

            switch (c)
            {
                case '"':
                    inQuotes = true;
                    break;
                case '\n':
                    lines++;
                    break;
                default:
                    for (var d = 0; d < SniffCandidates.Count; d++)
                    {
                        if (c == SniffCandidates[d])
                        {
                            counts[d]++;
                        }
                    }

                    break;
            }

            if (lines >= 16)
            {
                break;
            }
        }

        var best = 0;
        for (var d = 1; d < SniffCandidates.Count; d++)
        {
            if (counts[d] > counts[best])
            {
                best = d; // strict > keeps the , ; \t precedence on ties
            }
        }

        return SniffCandidates[best];
    }

    /// <summary>Parses csv text into a typed grid (rows may be ragged, like anchor-form bulk writes).</summary>
    public static List<List<ExcelValues.ParsedValue>> Parse(string text, char delimiter)
    {
        var records = ParseRecords(text, delimiter);
        var grid = new List<List<ExcelValues.ParsedValue>>(records.Count);
        foreach (var record in records)
        {
            var row = new List<ExcelValues.ParsedValue>(record.Count);
            foreach (var field in record)
            {
                row.Add(ExcelValues.ParseCsvField(field));
            }

            grid.Add(row);
        }

        return grid;
    }

    /// <summary>The raw RFC 4180 state machine; returns one string list per record.</summary>
    public static List<List<string>> ParseRecords(string text, char delimiter)
    {
        var records = new List<List<string>>();
        var fields = new List<string>();
        var sb = new StringBuilder();
        var inQuotes = false;
        var fieldWasQuoted = false; // a quoted empty field ("" alone) still counts as content

        var start = text.Length > 0 && text[0] == '\uFEFF' ? 1 : 0; // BOM tolerance
        for (var i = start; i < text.Length; i++)
        {
            var c = text[i];
            if (inQuotes)
            {
                switch (c)
                {
                    case '"' when i + 1 < text.Length && text[i + 1] == '"':
                        sb.Append('"');
                        i++;
                        break;
                    case '"':
                        inQuotes = false;
                        break;
                    case '\r': // newline inside quotes: normalize CRLF / CR to \n
                        sb.Append('\n');
                        if (i + 1 < text.Length && text[i + 1] == '\n')
                        {
                            i++;
                        }

                        break;
                    default:
                        sb.Append(c);
                        break;
                }

                continue;
            }

            if (c == '"' && sb.Length == 0)
            {
                inQuotes = true;
                fieldWasQuoted = true;
                continue;
            }

            if (c == delimiter)
            {
                fields.Add(sb.ToString());
                sb.Clear();
                fieldWasQuoted = false;
                continue;
            }

            if (c is '\n' or '\r')
            {
                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                {
                    i++;
                }

                fields.Add(sb.ToString());
                records.Add(fields);
                fields = [];
                sb.Clear();
                fieldWasQuoted = false;
                continue;
            }

            sb.Append(c);
        }

        if (inQuotes)
        {
            throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                "The csv ends inside a quoted field (unbalanced \").",
                "Close the quote (RFC 4180 escapes a literal quote as \"\"), or re-export the csv from its source.");
        }

        if (fields.Count > 0 || sb.Length > 0 || fieldWasQuoted)
        {
            fields.Add(sb.ToString());
            records.Add(fields);
        }

        // A final newline (or trailing blank lines) must not become rows.
        while (records.Count > 0 && records[^1] is [{ Length: 0 }])
        {
            records.RemoveAt(records.Count - 1);
        }

        return records;
    }

    // ----- export -------------------------------------------------------------

    /// <summary>
    /// One sheet window as RFC 4180 csv (CRLF records, minimal quoting). The
    /// window is rectangular: blank cells inside it become empty fields.
    /// </summary>
    public static (string Content, bool Truncated) Build(IXLRange? range, int maxBytes)
    {
        if (range is null)
        {
            return (string.Empty, false);
        }

        var sb = new StringBuilder();
        var address = range.RangeAddress;
        var columns = address.LastAddress.ColumnNumber - address.FirstAddress.ColumnNumber + 1;
        foreach (var row in range.Rows())
        {
            for (var column = 1; column <= columns; column++)
            {
                if (column > 1)
                {
                    sb.Append(',');
                }

                sb.Append(ExcelValues.CsvEscape(FieldText(row.Cell(column))));
            }

            sb.Append("\r\n");
            if (sb.Length > maxBytes)
            {
                return (sb.ToString(0, Math.Min(sb.Length, maxBytes)), true);
            }
        }

        return (sb.ToString(), false);
    }

    /// <summary>
    /// A cell's csv field text: typed so a re-import recreates the same value.
    /// Formulas contribute their cached/evaluated value, never formula text.
    /// </summary>
    public static string FieldText(IXLCell cell)
    {
        XLCellValue value;
        try
        {
            value = cell.Value; // evaluates dirty formulas in memory
        }
        catch (Exception)
        {
            value = cell.CachedValue;
        }

        return value.Type switch
        {
            XLDataType.Blank => string.Empty,
            XLDataType.Boolean => value.GetBoolean() ? "true" : "false",
            XLDataType.Number => value.GetNumber().ToString(CultureInfo.InvariantCulture),
            XLDataType.Text => value.GetText(),
            XLDataType.DateTime => value.GetDateTime() is { TimeOfDay.Ticks: 0 } d
                ? d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                : value.GetDateTime().ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture),
            XLDataType.TimeSpan => value.GetTimeSpan().ToString("c", CultureInfo.InvariantCulture),
            _ => value.ToString(CultureInfo.InvariantCulture), // errors keep their display text
        };
    }
}
