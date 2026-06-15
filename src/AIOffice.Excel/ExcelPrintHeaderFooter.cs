using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// (1.7) Print header/footer field-code strings, authored raw post-save.
///
/// ClosedXML 0.105's <c>IXLHFItem.AddText</c> does not persist header/footer text
/// to the saved bytes (measured: <c>oddHeader</c>/<c>oddFooter</c> come back empty),
/// so aioffice owns the <c>headerFooter</c> element directly — exactly like charts,
/// embeds and slicers. ClosedXML always emits an (empty) <c>&lt;headerFooter/&gt;</c>
/// after <c>pageSetup</c>; this pass POPULATES that element in place (a second one
/// would fail the schema's worksheet child-order). The OOXML header/footer encoding
/// is a single string per occurrence with section markers: <c>&amp;L</c> left,
/// <c>&amp;C</c> center, <c>&amp;R</c> right, followed by the field codes
/// (<c>&amp;P</c> page, <c>&amp;N</c> pages, <c>&amp;D</c> date, <c>&amp;T</c> time,
/// <c>&amp;F</c> file, <c>&amp;A</c> sheet) verbatim. Only the odd-page (= every
/// page, no differentFirst/differentOddEven) header/footer is written.
/// </summary>
internal static class ExcelPrintHeaderFooter
{
    /// <summary>
    /// One validated header/footer change queued at op time. Each section is null
    /// when the op did not mention it (leave it as-is) or empty when it cleared it.
    /// </summary>
    internal sealed record Spec(
        string SheetName,
        bool IsHeader,
        string? Left,
        string? Center,
        string? Right);

    /// <summary>Builds the OOXML occurrence string (<c>&amp;L…&amp;C…&amp;R…</c>) from three sections.</summary>
    public static string BuildOccurrence(string? left, string? center, string? right)
    {
        var result = new System.Text.StringBuilder();
        if (!string.IsNullOrEmpty(left))
        {
            result.Append("&L").Append(left);
        }

        if (!string.IsNullOrEmpty(center))
        {
            result.Append("&C").Append(center);
        }

        if (!string.IsNullOrEmpty(right))
        {
            result.Append("&R").Append(right);
        }

        return result.ToString();
    }

    /// <summary>
    /// Applies queued header/footer specs to the file ClosedXML just saved. Per-op
    /// sections are merged onto whatever is already in <c>oddHeader</c>/<c>oddFooter</c>
    /// so a later op can change one section without clearing the others.
    /// </summary>
    public static void Apply(string file, IReadOnlyList<Spec> specs)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is null)
        {
            return;
        }

        foreach (var group in specs.GroupBy(s => s.SheetName, StringComparer.OrdinalIgnoreCase))
        {
            var sheetElement = workbookPart.Workbook.Descendants<S.Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, group.Key, StringComparison.OrdinalIgnoreCase));
            if (sheetElement?.Id?.Value is not { } relId ||
                workbookPart.GetPartById(relId) is not WorksheetPart worksheetPart ||
                worksheetPart.Worksheet is not { } worksheet)
            {
                continue;
            }

            var headerFooter = EnsureHeaderFooter(worksheet);
            foreach (var spec in group)
            {
                ApplyOne(headerFooter, spec);
            }

            // Drop an all-empty headerFooter so a cleared sheet round-trips to nothing.
            if (string.IsNullOrEmpty(headerFooter.OddHeader?.Text) &&
                string.IsNullOrEmpty(headerFooter.OddFooter?.Text))
            {
                headerFooter.OddHeader?.Remove();
                headerFooter.OddFooter?.Remove();
            }

            worksheet.Save();
        }
    }

    private static void ApplyOne(S.HeaderFooter headerFooter, Spec spec)
    {
        var occurrence = spec.IsHeader
            ? (DocumentFormat.OpenXml.OpenXmlLeafTextElement?)headerFooter.OddHeader
            : headerFooter.OddFooter;
        var (left, center, right) = Parse(occurrence?.Text);

        // A null section is left untouched; a present (possibly empty) section replaces.
        left = spec.Left ?? left;
        center = spec.Center ?? center;
        right = spec.Right ?? right;

        var text = BuildOccurrence(left, center, right);
        if (spec.IsHeader)
        {
            SetOdd(headerFooter, text, isHeader: true);
        }
        else
        {
            SetOdd(headerFooter, text, isHeader: false);
        }
    }

    private static void SetOdd(S.HeaderFooter headerFooter, string text, bool isHeader)
    {
        if (isHeader)
        {
            if (headerFooter.OddHeader is null)
            {
                headerFooter.PrependChild(new S.OddHeader(text));
            }
            else
            {
                headerFooter.OddHeader.Text = text;
            }
        }
        else
        {
            if (headerFooter.OddFooter is null)
            {
                headerFooter.Append(new S.OddFooter(text));
            }
            else
            {
                headerFooter.OddFooter.Text = text;
            }
        }
    }

    /// <summary>
    /// ClosedXML always writes an (empty) headerFooter after pageSetup; reuse it.
    /// If absent, insert a fresh one in document-order (after pageSetup, before
    /// rowBreaks/colBreaks/tableParts) so the worksheet child sequence stays valid.
    /// </summary>
    private static S.HeaderFooter EnsureHeaderFooter(S.Worksheet worksheet)
    {
        var existing = worksheet.Elements<S.HeaderFooter>().FirstOrDefault();
        if (existing is not null)
        {
            return existing;
        }

        var headerFooter = new S.HeaderFooter();
        var pageSetup = worksheet.Elements<S.PageSetup>().FirstOrDefault();
        if (pageSetup is not null)
        {
            worksheet.InsertAfter(headerFooter, pageSetup);
        }
        else if (worksheet.Elements<S.PageMargins>().FirstOrDefault() is { } margins)
        {
            worksheet.InsertAfter(headerFooter, margins);
        }
        else
        {
            // Before any rowBreaks/colBreaks/tableParts that may already exist.
            var before = worksheet.Elements<S.RowBreaks>().Cast<DocumentFormat.OpenXml.OpenXmlElement>()
                .Concat(worksheet.Elements<S.ColumnBreaks>())
                .Concat(worksheet.Elements<S.TableParts>())
                .FirstOrDefault();
            if (before is not null)
            {
                worksheet.InsertBefore(headerFooter, before);
            }
            else
            {
                worksheet.Append(headerFooter);
            }
        }

        return headerFooter;
    }

    /// <summary>Splits an OOXML occurrence string into its (left, center, right) sections.</summary>
    public static (string? Left, string? Center, string? Right) Parse(string? occurrence)
    {
        if (string.IsNullOrEmpty(occurrence))
        {
            return (null, null, null);
        }

        string? left = null, center = null, right = null;
        var current = 'C'; // OOXML default section when no marker leads is center
        var buffer = new System.Text.StringBuilder();

        void Flush()
        {
            var text = buffer.ToString();
            buffer.Clear();
            if (text.Length == 0)
            {
                return;
            }

            switch (current)
            {
                case 'L': left = (left ?? string.Empty) + text; break;
                case 'C': center = (center ?? string.Empty) + text; break;
                case 'R': right = (right ?? string.Empty) + text; break;
            }
        }

        for (var i = 0; i < occurrence.Length; i++)
        {
            var ch = occurrence[i];
            if (ch == '&' && i + 1 < occurrence.Length && occurrence[i + 1] is 'L' or 'C' or 'R')
            {
                Flush();
                current = occurrence[i + 1];
                i++;
                continue;
            }

            buffer.Append(ch);
        }

        Flush();
        return (left, center, right);
    }

    /// <summary>Reads a sheet's odd header/footer sections back for <c>get</c> (raw; ClosedXML cannot see them).</summary>
    public static (object? Header, object? Footer) Read(string file, string sheetName)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var workbookPart = document.WorkbookPart;
        var sheetElement = workbookPart?.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheetElement?.Id?.Value is not { } relId ||
            workbookPart!.GetPartById(relId) is not WorksheetPart worksheetPart ||
            worksheetPart.Worksheet?.Elements<S.HeaderFooter>().FirstOrDefault() is not { } headerFooter)
        {
            return (null, null);
        }

        return (Section(headerFooter.OddHeader?.Text), Section(headerFooter.OddFooter?.Text));
    }

    private static object? Section(string? occurrence)
    {
        var (left, center, right) = Parse(occurrence);
        if (left is null && center is null && right is null)
        {
            return null;
        }

        return new
        {
            left = string.IsNullOrEmpty(left) ? null : left,
            center = string.IsNullOrEmpty(center) ? null : center,
            right = string.IsNullOrEmpty(right) ? null : right,
        };
    }
}
