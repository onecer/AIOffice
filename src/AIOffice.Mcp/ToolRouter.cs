using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ModelContextProtocol.Protocol;

namespace AIOffice.Mcp;

/// <summary>
/// Routes MCP <c>tools/call</c> requests to the shared <see cref="CommandService"/>.
/// Every result carries the canonical JSON envelope as a text content block —
/// byte-identical in shape to CLI stdout. <c>office_render</c> additionally
/// attaches an SVG image block when an inline SVG artifact was produced.
/// Failure envelopes set <c>isError</c> so MCP clients surface them, but the
/// envelope (with its mandatory <c>error.suggestion</c>) is always the payload.
/// </summary>
internal static class ToolRouter
{
    public static CallToolResult Call(CommandService service, CallToolRequestParams? request)
    {
        var name = request?.Name ?? string.Empty;
        Envelope envelope;
        try
        {
            envelope = Dispatch(service, name, ToJsonObject(request?.Arguments));
        }
        catch (Exception ex)
        {
            // Defense in depth: CommandService already converts exceptions, but a
            // tools/call must never bubble a raw exception to the protocol layer.
            envelope = Envelope.FromException(ex);
        }

        return ToResult(name, envelope);
    }

    private static Envelope Dispatch(CommandService service, string name, JsonObject args) => name switch
    {
        "office_create" => service.Create(args),
        "office_read" => service.Read(args),
        "office_query" => service.Query(args),
        "office_get" => service.Get(args),
        "office_edit" => service.Edit(args),
        "office_render" => service.Render(args),
        "office_validate" => service.Validate(args),
        "office_template" => service.Template(args),
        "file_snapshot" => service.Snapshot(args),
        "office_status" => service.Status(),
        "office_help" => service.Help(args),
        "office_schema" => service.Schema(args),
        "preview_open" => service.PreviewOpen(args),
        "preview_selection" => service.PreviewSelection(args),
        _ => Envelope.Fail(
            ErrorCodes.InvalidArgs,
            $"Unknown tool: '{name}'.",
            "Call tools/list for the available tools.",
            candidates: ToolCatalog.Names),
    };

    private static JsonObject ToJsonObject(IDictionary<string, JsonElement>? arguments)
    {
        var result = new JsonObject();
        if (arguments is null)
        {
            return result;
        }

        foreach (var (key, value) in arguments)
        {
            result[key] = JsonNode.Parse(value.GetRawText());
        }

        return result;
    }

    private static CallToolResult ToResult(string toolName, Envelope envelope)
    {
        var content = new List<ContentBlock> { new TextContentBlock { Text = envelope.ToJson() } };
        if (toolName == "office_render" && envelope.IsOk)
        {
            if (TryGetInlineSvg(envelope) is { } svg)
            {
                content.Add(new ImageContentBlock
                {
                    MimeType = "image/svg+xml",
                    Data = Encoding.UTF8.GetBytes(svg), // serialized base64 on the wire by the SDK
                });
            }
            else if (TryReadWrittenPng(envelope) is { } png)
            {
                content.Add(new ImageContentBlock
                {
                    MimeType = "image/png",
                    Data = png, // the file's bytes verbatim — no downscaling
                });
            }
        }

        return new CallToolResult
        {
            Content = content,
            IsError = envelope.IsOk ? null : true,
        };
    }

    /// <summary>Reads the PNG a png-render wrote (<c>data.format == "png"</c> + <c>data.written</c>), if any.</summary>
    private static byte[]? TryReadWrittenPng(Envelope envelope)
    {
        try
        {
            if (JsonSerializer.SerializeToNode(envelope.Data, JsonDefaults.Options) is not JsonObject data ||
                data["format"] is not JsonValue fv || !fv.TryGetValue<string>(out var format) || format != "png" ||
                data["written"] is not JsonValue wv || !wv.TryGetValue<string>(out var written))
            {
                return null;
            }

            return File.Exists(written) ? File.ReadAllBytes(written) : null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
        {
            return null; // the envelope (with data.written) still answers; only the inline block is skipped
        }
    }

    /// <summary>Extracts an inlined SVG render artifact (<c>data.content</c> + a .svg output), if any.</summary>
    private static string? TryGetInlineSvg(Envelope envelope)
    {
        try
        {
            if (JsonSerializer.SerializeToNode(envelope.Data, JsonDefaults.Options) is not JsonObject data)
            {
                return null;
            }

            var content = data["content"] is JsonValue cv && cv.TryGetValue<string>(out var c) ? c : null;
            var firstOutput = data["outputs"] is JsonArray { Count: > 0 } outputs &&
                              outputs[0] is JsonValue ov && ov.TryGetValue<string>(out var o) ? o : null;
            var isSvg = firstOutput?.EndsWith(".svg", StringComparison.OrdinalIgnoreCase) == true ||
                        content?.TrimStart().StartsWith("<svg", StringComparison.OrdinalIgnoreCase) == true;
            return isSvg ? content : null;
        }
        catch (NotSupportedException)
        {
            return null; // Unserializable handler payload: skip the enrichment, keep the envelope.
        }
    }
}
