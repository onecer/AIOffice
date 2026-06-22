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
    /// The page shell: styling plus the full interactive layer — click /
    /// box-drag select (auto-POSTed to /selection), agent-pushed mark highlights
    /// and goto-scroll over SSE, double-click inline edit (POST /api/edit), and a
    /// soft per-path live reload that swaps only the changed nodes when the file
    /// changes on disk (preserving scroll, selection, marks and the SSE stream).
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
        .aio-mark { outline: 3px solid var(--aio-mark-color, #ffeb3b) !important; outline-offset: 1px; position: relative; }
        .aio-mark.aio-tofix { outline-style: dashed !important; }
        .aio-mark[data-aio-note]:hover::after { content: attr(data-aio-note); position: absolute; left: 0; top: -1.7em; background: #111827; color: #fff; font-size: 11px; line-height: 1.4; padding: 2px 6px; border-radius: 4px; white-space: nowrap; z-index: 20; pointer-events: none; }
        .aio-flash { animation: aio-flash 1.3s ease-out; }
        @keyframes aio-flash { 0% { background: #fde68a; } 100% { background: transparent; } }
        #aio-dragbox { position: fixed; border: 1px solid #2563eb; background: rgba(37,99,235,0.12); pointer-events: none; z-index: 30; display: none; }
        [data-aio-path][contenteditable="true"] { outline: 2px solid #059669 !important; background: #ecfdf5; cursor: text; }
        .aio-toast { position: fixed; bottom: 16px; left: 50%; transform: translateX(-50%); background: #111827; color: #fff; padding: 6px 12px; border-radius: 6px; font-size: 12px; z-index: 40; opacity: 0; transition: opacity .2s; }
        .aio-toast.show { opacity: 1; }
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
          let marks = [];
          let suppressClick = false;

          const esc = (p) => (window.CSS && CSS.escape) ? CSS.escape(p) : p.replace(/["\\]/g, "\\$&");
          const byPath = (p) => document.querySelector('[data-aio-path="' + esc(p) + '"]');
          const toast = (msg) => { let t = document.querySelector(".aio-toast"); if (!t) { t = document.createElement("div"); t.className = "aio-toast"; document.body.appendChild(t); } t.textContent = msg; t.classList.add("show"); setTimeout(() => t.classList.remove("show"), 1800); };

          const applySelection = () => {
            document.querySelectorAll("[data-aio-path]").forEach((el) => {
              el.classList.toggle("aio-selected", selected.has(el.getAttribute("data-aio-path")));
            });
          };
          const pushSelection = () => {
            fetch("/selection", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ paths: Array.from(selected) }) }).catch(() => {});
          };

          const applyMarks = () => {
            document.querySelectorAll(".aio-mark").forEach((el) => {
              el.classList.remove("aio-mark", "aio-tofix");
              el.removeAttribute("data-aio-note");
              el.style.removeProperty("--aio-mark-color");
            });
            marks.forEach((m) => {
              const el = byPath(m.path); if (!el) return;
              el.classList.add("aio-mark");
              if (m.toFix) el.classList.add("aio-tofix");
              if (m.color) el.style.setProperty("--aio-mark-color", m.color);
              const note = [m.note, m.find ? ("find: " + m.find) : null, m.toFix ? "⚑ fix" : null].filter(Boolean).join("  ·  ");
              if (note) el.setAttribute("data-aio-note", note);
            });
          };

          // Compare a decorated live node against a fresh one: strip our injected
          // classes/attrs so only real content changes count as a diff.
          const cleanHtml = (el) => {
            const c = el.cloneNode(true);
            c.classList.remove("aio-selected", "aio-mark", "aio-tofix", "aio-flash");
            c.removeAttribute("data-aio-note");
            c.style.removeProperty("--aio-mark-color");
            if (!c.getAttribute("style")) c.removeAttribute("style");
            if (c.getAttribute("class") === "") c.removeAttribute("class");
            return c.outerHTML;
          };

          // Soft live reload: fetch fresh content, replace only changed nodes
          // (no location.reload) — scroll, selection, marks and SSE all survive.
          const softReload = async () => {
            let html; try { html = await (await fetch("/content")).text(); } catch { return; }
            const main = document.querySelector("main"); if (!main) return;
            const tmp = document.createElement("div"); tmp.innerHTML = html;
            const oldNodes = new Map(); main.querySelectorAll("[data-aio-path]").forEach((e) => oldNodes.set(e.getAttribute("data-aio-path"), e));
            const newNodes = new Map(); tmp.querySelectorAll("[data-aio-path]").forEach((e) => newNodes.set(e.getAttribute("data-aio-path"), e));
            const sameKeys = oldNodes.size === newNodes.size && [...oldNodes.keys()].every((k) => newNodes.has(k));
            if (!sameKeys) {
              main.innerHTML = html;
            } else {
              oldNodes.forEach((oldEl, p) => { const ne = newNodes.get(p); if (ne && cleanHtml(oldEl) !== ne.outerHTML) oldEl.replaceWith(ne); });
            }
            applySelection(); applyMarks();
          };

          // ---- click select ----
          document.addEventListener("click", (ev) => {
            if (suppressClick) { suppressClick = false; return; }
            const el = ev.target instanceof Element ? ev.target.closest("[data-aio-path]") : null;
            if (!el || el.getAttribute("contenteditable") === "true") return;
            ev.preventDefault();
            const path = el.getAttribute("data-aio-path");
            if (ev.shiftKey || ev.metaKey || ev.ctrlKey) { if (!selected.delete(path)) selected.add(path); }
            else { selected.clear(); selected.add(path); }
            applySelection(); pushSelection();
          });

          // ---- box-drag (rubber-band) multi-select ----
          let drag = null;
          const box = document.createElement("div"); box.id = "aio-dragbox"; document.body.appendChild(box);
          document.addEventListener("mousedown", (ev) => {
            if (ev.button !== 0 || (ev.target instanceof Element && ev.target.closest("[contenteditable='true']"))) return;
            drag = { x: ev.clientX, y: ev.clientY, moved: false, add: ev.shiftKey || ev.metaKey || ev.ctrlKey };
          });
          document.addEventListener("mousemove", (ev) => {
            if (!drag) return;
            if (!drag.moved && Math.abs(ev.clientX - drag.x) + Math.abs(ev.clientY - drag.y) < 6) return;
            drag.moved = true;
            box.style.display = "block";
            box.style.left = Math.min(ev.clientX, drag.x) + "px";
            box.style.top = Math.min(ev.clientY, drag.y) + "px";
            box.style.width = Math.abs(ev.clientX - drag.x) + "px";
            box.style.height = Math.abs(ev.clientY - drag.y) + "px";
          });
          document.addEventListener("mouseup", (ev) => {
            if (!drag) return;
            const d = drag; drag = null; box.style.display = "none";
            if (!d.moved) return;
            suppressClick = true;
            const x1 = Math.min(ev.clientX, d.x), y1 = Math.min(ev.clientY, d.y), x2 = Math.max(ev.clientX, d.x), y2 = Math.max(ev.clientY, d.y);
            if (!d.add) selected.clear();
            document.querySelectorAll("[data-aio-path]").forEach((el) => {
              const b = el.getBoundingClientRect();
              if (b.left < x2 && b.right > x1 && b.top < y2 && b.bottom > y1) selected.add(el.getAttribute("data-aio-path"));
            });
            applySelection(); pushSelection();
          });

          // ---- inline edit (double-click an xlsx cell or docx paragraph) ----
          document.addEventListener("dblclick", (ev) => {
            const el = ev.target instanceof Element ? ev.target.closest("[data-aio-path]") : null;
            if (!el) return;
            const path = el.getAttribute("data-aio-path");
            const editable = /\/[^/]+\/[A-Z]{1,3}[0-9]{1,7}$/.test(path) || /^\/body\/p\[[0-9]+\]$/.test(path);
            if (!editable) return;
            ev.preventDefault();
            const original = el.textContent;
            let done = false;
            el.setAttribute("contenteditable", "true"); el.focus();
            const sel = window.getSelection(); sel.removeAllRanges();
            const rg = document.createRange(); rg.selectNodeContents(el); sel.addRange(rg);
            const finish = (commit) => {
              if (done) return; done = true;
              el.removeAttribute("contenteditable");
              const text = el.textContent;
              if (!commit || text === original) { el.textContent = original; return; }
              const props = path.startsWith("/body/") ? { text: text } : { value: text };
              fetch("/api/edit", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ op: "set", path: path, props: props }) })
                .then((r) => { if (!r.ok) { el.textContent = original; toast("Edit rejected"); } })
                .catch(() => { el.textContent = original; });
            };
            el.addEventListener("keydown", (k) => { if (k.key === "Enter") { k.preventDefault(); finish(true); } else if (k.key === "Escape") { finish(false); } });
            el.addEventListener("blur", () => finish(true), { once: true });
          });

          // ---- initial state + SSE wiring ----
          fetch("/selection").then((r) => r.json()).then((s) => { (s.paths || []).forEach((p) => selected.add(p)); applySelection(); }).catch(() => {});
          fetch("/marks").then((r) => r.json()).then((s) => { marks = s.marks || []; applyMarks(); }).catch(() => {});
          const es = new EventSource("/events");
          es.addEventListener("reload", softReload);
          es.addEventListener("marks", (e) => { try { marks = JSON.parse(e.data).marks || []; applyMarks(); } catch (_) {} });
          es.addEventListener("scroll", (e) => {
            try {
              const el = byPath(JSON.parse(e.data).path);
              if (el) { el.scrollIntoView({ behavior: "smooth", block: "center" }); el.classList.add("aio-flash"); setTimeout(() => el.classList.remove("aio-flash"), 1300); }
            } catch (_) {}
          });
        })();
        </script>
        </body>
        </html>
        """;
}
