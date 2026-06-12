using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace AIOffice.Pptx;

/// <summary>
/// {{key}} merge across slide text runs. PowerPoint freely splits text into
/// multiple runs, so when a placeholder spans runs the paragraph's runs are
/// coalesced into the first run (keeping its formatting) before replacing.
/// </summary>
internal static class PptxTemplater
{
    public static int Apply(PresentationPart presentation, IReadOnlyDictionary<string, string> data)
    {
        var replacements = 0;
        foreach (var (_, slidePart) in PptxDoc.Slides(presentation))
        {
            if (slidePart.Slide is null)
            {
                continue;
            }

            foreach (var paragraph in slidePart.Slide.Descendants<A.Paragraph>())
            {
                replacements += MergeParagraph(paragraph, data);
            }
        }

        return replacements;
    }

    private static int MergeParagraph(A.Paragraph paragraph, IReadOnlyDictionary<string, string> data)
    {
        var runs = paragraph.Elements<A.Run>().Where(r => r.Text is not null).ToList();
        if (runs.Count == 0)
        {
            return 0;
        }

        var combined = string.Concat(runs.Select(r => r.Text!.Text));
        if (!combined.Contains("{{", StringComparison.Ordinal))
        {
            return 0;
        }

        // A placeholder that exists in the combined text but in no single run
        // spans run boundaries: coalesce the paragraph into its first run.
        var spanning = data.Keys.Any(key =>
        {
            var placeholder = Placeholder(key);
            return combined.Contains(placeholder, StringComparison.Ordinal) &&
                !runs.Any(r => r.Text!.Text.Contains(placeholder, StringComparison.Ordinal));
        });

        if (spanning)
        {
            runs[0].Text!.Text = combined;
            foreach (var run in runs.Skip(1))
            {
                run.Remove();
            }

            runs = [runs[0]];
        }

        var replacements = 0;
        foreach (var run in runs)
        {
            var text = run.Text!.Text;
            foreach (var (key, value) in data)
            {
                var placeholder = Placeholder(key);
                var count = CountOccurrences(text, placeholder);
                if (count > 0)
                {
                    text = text.Replace(placeholder, value, StringComparison.Ordinal);
                    replacements += count;
                }
            }

            run.Text!.Text = text;
        }

        return replacements;
    }

    private static string Placeholder(string key) => "{{" + key + "}}";

    private static int CountOccurrences(string text, string needle)
    {
        var count = 0;
        var index = 0;
        while ((index = text.IndexOf(needle, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += needle.Length;
        }

        return count;
    }
}
