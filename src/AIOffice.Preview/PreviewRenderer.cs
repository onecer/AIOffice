using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using AIOffice.Excel;
using AIOffice.Pptx;
using AIOffice.Word;

namespace AIOffice.Preview;

/// <summary>
/// Turns a document into the interactive preview page served at <c>GET /</c>.
///
/// The data-aio-path contract: every rendered element that maps to an
/// addressable node carries <c>data-aio-path="&lt;canonical path&gt;"</c> so a
/// browser click maps back to a document path. When the format handler's own
/// render already emits the attribute this class uses it untouched; until the
/// handlers do, the attributes are injected here (docx blocks, pptx shape
/// groups) or the grid is built from the handler's structured output (xlsx,
/// whose HTML render cannot be enriched after the fact because sparse rows
/// drop empty cells).
/// </summary>
internal static partial class PreviewRenderer
{
    private const string PathAttribute = "data-aio-path";

    [GeneratedRegex("^([A-Z]{1,3})([0-9]{1,7})(?::([A-Z]{1,3})([0-9]{1,7}))?$")]
    private static partial Regex CellOrRange();

    public static IReadOnlyList<string> SupportedExtensions { get; } = [".docx", ".xlsx", ".pptx"];

    /// <summary>Renders the full HTML page (shell + content + selection/reload script).</summary>
    public static string RenderPage(string file, Workspace workspace)
    {
        var content = RenderContent(file, workspace);
        return Shell(Path.GetFileName(file), content);
    }

    /// <summary>The format-specific content fragment, with data-aio-path attributes guaranteed.</summary>
    public static string RenderContent(string file, Workspace workspace)
    {
        var extension = Path.GetExtension(file).ToLowerInvariant();
        return extension switch
        {
            ".docx" => RenderDocx(file, workspace),
            ".xlsx" => RenderXlsx(file, workspace),
            ".pptx" => RenderPptx(file, workspace),
            _ => throw UnsupportedExtension(extension),
        };
    }

    internal static AiofficeException UnsupportedExtension(string extension) => new(
        ErrorCodes.UnsupportedFeature,
        $"Preview cannot render '{(extension.Length == 0 ? "(no extension)" : extension)}' files.",
        "Preview supports .docx, .xlsx and .pptx. Convert the file to one of these first.",
        candidates: SupportedExtensions);

    // ------------------------------------------------------------------- docx

    private static string RenderDocx(string file, Workspace workspace)
    {
        var data = UnwrapData(new WordHandler().Render(Ctx(file, workspace, ("to", "html"))));
        var html = data["content"]!.GetValue<string>();
        return InjectWordPaths(html);
    }

    /// <summary>
    /// Adds <c>data-aio-path</c> to every top-level block of the docx HTML
    /// render. The M0 renderer emits attribute-less <c>&lt;p&gt;</c>,
    /// <c>&lt;h1..6&gt;</c>, <c>&lt;li&gt;</c> (one per body paragraph, in body
    /// order) and <c>&lt;table&gt;</c> (one per body table), so sequential
    /// counters reproduce exactly the canonical indices WordAddress assigns.
    /// Output that already carries the attribute is returned untouched.
    /// </summary>
    internal static string InjectWordPaths(string html)
    {
        if (html.Contains(PathAttribute, StringComparison.Ordinal))
        {
            return html;
        }

        var sb = new StringBuilder(html.Length + 512);
        var paragraph = 0;
        var table = 0;
        var insideTable = false;
        var i = 0;

        while (i < html.Length)
        {
            if (html[i] != '<')
            {
                sb.Append(html[i]);
                i++;
                continue;
            }

            var close = html.IndexOf('>', i);
            if (close < 0)
            {
                sb.Append(html, i, html.Length - i);
                break;
            }

            var tag = html[i..(close + 1)];
            if (insideTable)
            {
                sb.Append(tag);
                insideTable = tag != "</table>";
            }
            else if (tag == "<table>")
            {
                table++;
                insideTable = true;
                sb.Append(CultureInfo.InvariantCulture, $"<table {PathAttribute}=\"/body/table[{table}]\">");
            }
            else if (IsParagraphBlockOpen(tag, out var name))
            {
                paragraph++;
                sb.Append(CultureInfo.InvariantCulture, $"<{name} {PathAttribute}=\"/body/p[{paragraph}]\">");
            }
            else
            {
                sb.Append(tag);
            }

            i = close + 1;
        }

        return sb.ToString();
    }

    /// <summary>Matches the exact attribute-less block tags the docx renderer emits for body paragraphs.</summary>
    private static bool IsParagraphBlockOpen(string tag, out string name)
    {
        name = string.Empty;
        if (tag is "<p>" or "<li>")
        {
            name = tag[1..^1];
            return true;
        }

        if (tag.Length == 4 && tag[1] == 'h' && tag[2] is >= '1' and <= '6' && tag[0] == '<' && tag[3] == '>')
        {
            name = tag[1..3];
            return true;
        }

        return false;
    }

    // ------------------------------------------------------------------- pptx

    private static string RenderPptx(string file, Workspace workspace)
    {
        var data = UnwrapData(new PptxHandler().Render(Ctx(file, workspace, ("to", "svg"))));
        var sb = new StringBuilder();
        foreach (var slide in data["slides"]!.AsArray())
        {
            sb.Append("<section class=\"aio-slide\">\n");
            sb.Append(WrapPptxShapes(slide!["svg"]!.GetValue<string>()));
            sb.Append("\n</section>\n");
        }

        return sb.ToString();
    }

    /// <summary>
    /// Wraps each shape of a slide SVG in <c>&lt;g data-aio-path="…"&gt;</c>.
    /// The M0 renderer emits, per shape, one <c>&lt;rect data-path="…"&gt;</c>
    /// line followed by that shape's <c>&lt;text&gt;</c> lines, so grouping a
    /// data-path rect with the lines up to the next data-path rect (or the svg
    /// end) reconstructs the shape boundary. SVG that already carries
    /// data-aio-path is returned untouched.
    /// </summary>
    internal static string WrapPptxShapes(string svg)
    {
        if (svg.Contains(PathAttribute, StringComparison.Ordinal))
        {
            return svg;
        }

        const string marker = "data-path=\"";
        var lines = svg.Split('\n');
        var sb = new StringBuilder(svg.Length + 256);
        var groupOpen = false;

        foreach (var line in lines)
        {
            var markerAt = line.IndexOf(marker, StringComparison.Ordinal);
            if (markerAt >= 0)
            {
                if (groupOpen)
                {
                    sb.Append("  </g>\n");
                }

                var start = markerAt + marker.Length;
                var end = line.IndexOf('"', start);
                var path = line[start..end]; // already XML-escaped by the renderer
                sb.Append(CultureInfo.InvariantCulture, $"  <g {PathAttribute}=\"{path}\">\n");
                sb.Append(line).Append('\n');
                groupOpen = true;
            }
            else if (line.StartsWith("</svg>", StringComparison.Ordinal))
            {
                if (groupOpen)
                {
                    sb.Append("  </g>\n");
                    groupOpen = false;
                }

                sb.Append(line).Append('\n');
            }
            else
            {
                sb.Append(line).Append('\n');
            }
        }

        return sb.ToString().TrimEnd('\n');
    }

    // ------------------------------------------------------------------- xlsx

    private static string RenderXlsx(string file, Workspace workspace)
    {
        var handler = new ExcelHandler();

        // Prefer the handler's own render once it carries the attributes.
        var rendered = UnwrapData(handler.Render(Ctx(file, workspace, ("to", "html"))));
        var content = rendered["content"]!.GetValue<string>();
        if (content.Contains(PathAttribute, StringComparison.Ordinal))
        {
            return content;
        }

        // Otherwise build the grid from structured output: the HTML render
        // skips empty cells inside sparse rows, so positions cannot be
        // recovered from it after the fact.
        var outline = UnwrapData(handler.Read(Ctx(file, workspace, ("view", "outline"))));
        var sb = new StringBuilder();
        foreach (var sheetNode in outline["sheets"]!.AsArray())
        {
            var name = sheetNode!["name"]!.GetValue<string>();
            var usedRange = sheetNode["usedRange"]?.GetValue<string>();
            var sheetPath = SheetPath(name);

            sb.Append("<table data-sheet=\"").Append(EscapeHtml(name)).Append("\">\n<caption>")
              .Append(EscapeHtml(name)).Append("</caption>\n");

            if (usedRange is not null)
            {
                var grid = UnwrapData(handler.Get(Ctx(
                    file, workspace,
                    ("path", sheetPath + "/" + usedRange),
                    ("maxCells", 20000))));
                AppendGridRows(sb, sheetPath, grid);
            }

            sb.Append("</table>\n");
        }

        return sb.ToString();
    }

    private static void AppendGridRows(StringBuilder sb, string sheetPath, JsonNode grid)
    {
        var kind = grid["kind"]!.GetValue<string>();
        if (kind == "cell")
        {
            // A 1x1 used range resolves to a single cell.
            sb.Append("<tr><td ").Append(PathAttribute).Append("=\"")
              .Append(EscapeHtml(grid["path"]!.GetValue<string>())).Append("\">")
              .Append(EscapeHtml(grid["text"]?.GetValue<string>() ?? CellText(grid["value"])))
              .Append("</td></tr>\n");
            return;
        }

        var range = grid["range"]!.GetValue<string>();
        var match = CellOrRange().Match(range);
        if (!match.Success)
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                $"Unrecognized range address from the xlsx handler: '{range}'.",
                "This is a bug in aioffice preview. Re-run and report the issue with this message.");
        }

        var startColumn = ColumnNumber(match.Groups[1].Value);
        var startRow = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);

        var values = grid["values"]!.AsArray();
        for (var r = 0; r < values.Count; r++)
        {
            sb.Append("<tr>");
            var row = values[r]!.AsArray();
            for (var c = 0; c < row.Count; c++)
            {
                var cellRef = ColumnLetters(startColumn + c) +
                              (startRow + r).ToString(CultureInfo.InvariantCulture);
                sb.Append("<td ").Append(PathAttribute).Append("=\"")
                  .Append(EscapeHtml(sheetPath + "/" + cellRef)).Append("\">")
                  .Append(EscapeHtml(CellText(row[c])))
                  .Append("</td>");
            }

            sb.Append("</tr>\n");
        }

        if (grid["truncated"]?.GetValue<bool>() == true)
        {
            sb.Append("<tr><td class=\"aio-truncated\">… output truncated; the sheet has more rows</td></tr>\n");
        }
    }

    /// <summary>Canonical sheet path with the core grammar's quoting rules.</summary>
    private static string SheetPath(string sheetName) =>
        new DocPath
        {
            Segments = [new PathSegment { Kind = PathSegmentKind.Name, Name = sheetName }],
        }.ToCanonicalString();

    private static string CellText(JsonNode? value) => value switch
    {
        null => string.Empty,
        JsonValue v => v.ToString(),
        _ => value.ToJsonString(JsonDefaults.Options),
    };

    private static int ColumnNumber(string letters)
    {
        var n = 0;
        foreach (var c in letters)
        {
            n = (n * 26) + (c - 'A' + 1);
        }

        return n;
    }

    private static string ColumnLetters(int number)
    {
        var sb = new StringBuilder(3);
        while (number > 0)
        {
            number--;
            sb.Insert(0, (char)('A' + (number % 26)));
            number /= 26;
        }

        return sb.ToString();
    }

    // ---------------------------------------------------------------- shared

    private static CommandContext Ctx(string file, Workspace workspace, params (string Key, JsonNode? Value)[] args)
    {
        var jsonArgs = new JsonObject();
        foreach (var (key, value) in args)
        {
            jsonArgs[key] = value;
        }

        return new CommandContext { Workspace = workspace, File = file, Args = jsonArgs };
    }

    /// <summary>Envelope data as a JsonNode; failure envelopes resurface as the typed exception.</summary>
    private static JsonNode UnwrapData(Envelope envelope)
    {
        if (!envelope.IsOk)
        {
            var error = envelope.Error!;
            throw new AiofficeException(error.Code, error.Message, error.Suggestion, error.Candidates);
        }

        return JsonSerializer.SerializeToNode(envelope.Data, JsonDefaults.Options)!;
    }

    internal static string EscapeHtml(string text) => text
        .Replace("&", "&amp;", StringComparison.Ordinal)
        .Replace("<", "&lt;", StringComparison.Ordinal)
        .Replace(">", "&gt;", StringComparison.Ordinal)
        .Replace("\"", "&quot;", StringComparison.Ordinal);

    /// <summary>
    /// The page shell: minimal styling plus the click-to-select and live-reload
    /// layer. Elements with data-aio-path are clickable (click = select one,
    /// shift/cmd-click = toggle), selection auto-POSTs to /selection, and
    /// /events reloads the page when the file changes on disk.
    /// </summary>
    private static string Shell(string title, string content) => $$"""
        <!DOCTYPE html>
        <html lang="en">
        <head>
        <meta charset="utf-8">
        <title>aioffice preview — {{EscapeHtml(title)}}</title>
        <style>
        body { margin: 0; padding: 24px; background: #f3f4f6; font-family: -apple-system, "Segoe UI", Helvetica, Arial, sans-serif; color: #111827; }
        main { max-width: 960px; margin: 0 auto; background: #ffffff; padding: 24px; box-shadow: 0 1px 4px rgba(0,0,0,0.12); }
        table { border-collapse: collapse; margin: 12px 0; }
        td, th { border: 1px solid #d1d5db; padding: 2px 8px; min-width: 2em; }
        caption { text-align: left; font-weight: 600; padding: 4px 0; }
        .aio-slide { margin: 0 0 24px; }
        [data-aio-path] { cursor: pointer; }
        [data-aio-path]:hover { outline: 1px dashed #93c5fd; }
        .aio-selected { outline: 2px solid #2563eb !important; outline-offset: 1px; }
        g.aio-selected rect { stroke: #2563eb; stroke-width: 2; }
        .aio-truncated { color: #6b7280; font-style: italic; }
        </style>
        </head>
        <body>
        <main>
        {{content}}
        </main>
        <script>
        (() => {
          "use strict";
          const selected = new Set();
          const apply = () => {
            document.querySelectorAll("[data-aio-path]").forEach((el) => {
              el.classList.toggle("aio-selected", selected.has(el.getAttribute("data-aio-path")));
            });
          };
          const push = () => {
            fetch("/selection", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ paths: Array.from(selected) })
            }).catch(() => {});
          };
          document.addEventListener("click", (ev) => {
            const el = ev.target instanceof Element ? ev.target.closest("[data-aio-path]") : null;
            if (!el) return;
            ev.preventDefault();
            const path = el.getAttribute("data-aio-path");
            if (ev.shiftKey || ev.metaKey || ev.ctrlKey) {
              if (!selected.delete(path)) selected.add(path);
            } else {
              selected.clear();
              selected.add(path);
            }
            apply();
            push();
          });
          fetch("/selection").then((r) => r.json()).then((s) => {
            (s.paths || []).forEach((p) => selected.add(p));
            apply();
          }).catch(() => {});
          new EventSource("/events").addEventListener("reload", () => location.reload());
        })();
        </script>
        </body>
        </html>
        """;
}
