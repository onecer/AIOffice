using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using AIOffice.Core.Cli;

namespace AIOffice.Cli;

/// <summary>
/// The <c>plugin</c> verb: installs aioffice into AI coding hosts (Claude Code,
/// Codex, opencode, TonoBraid) by registering the MCP server, dropping the agent
/// guide/skill, and adding the <c>/aioffice</c> command where supported. Every
/// write is idempotent (keyed maps + marker blocks) and reversible
/// (<c>plugin uninstall</c>). No network, no host-CLI dependency — pure local
/// config-file writes plus file drops. The binary is the source of truth: it
/// embeds one canonical payload and resolves its own absolute path at runtime.
/// </summary>
internal static class PluginVerbs
{
    private static readonly string[] Actions = ["install", "uninstall", "list", "status"];
    private static readonly string[] KnownHosts = ["claude", "codex", "opencode", "tonobraid"];

    private const string PluginId = "aioffice";
    private const string MarkerBegin = "<!-- BEGIN aioffice -->";
    private const string MarkerEnd = "<!-- END aioffice -->";

    private const string SkillDescription =
        "Create or edit polished Word/Excel/PowerPoint files (.docx/.xlsx/.pptx) with the aioffice CLI/MCP. " +
        "Use whenever producing or modifying an Office document, slide deck, spreadsheet, dashboard, or report. " +
        "Enforces the design method + render→look→fix loop so output looks designed (and varied per brand/brief), " +
        "not default-white and not one house style. " +
        "Triggers: \"make a deck/slides/presentation\", \"build a spreadsheet/dashboard\", \"write a .docx/report\", any .pptx/.xlsx/.docx deliverable.";

    private const string CommandDescription =
        "Build a designed Office deliverable (deck/dashboard/report) via aioffice — enforces brief → direction → render→look→fix.";

    private const string AgentDescription =
        "Office-document specialist — build/edit .docx/.xlsx/.pptx via the aioffice MCP tools, enforcing the design + render→look→fix method.";

    private const string GuidePointer =
        "## Office documents (.docx / .xlsx / .pptx) → use aioffice\n" +
        "Build or edit any Word/Excel/PowerPoint file with the **aioffice** MCP tools (`office_*`) or the `aioffice` CLI — " +
        "never hand-roll OOXML or use python-pptx/pptxgenjs/SheetJS. aioffice ships a self-contained binary that reads/edits/renders real OOXML.\n" +
        "Method: brief → pick a non-default visual direction → load design props via `office_help` → build (theme + filled shapes + labeled charts) " +
        "→ render with `--engine soffice` → LOOK → fix → `validate`(0) → `audit`. " +
        "See the bundled aioffice skill / AGENT-GUIDE for the palette menu and per-format mechanics.";

    // ------------------------------------------------------------------ entry

    public static Envelope Run(ParsedArgs parsed)
    {
        var action = parsed.Positionals.Count > 0 ? parsed.Positionals[0] : null;
        if (action is null || !Actions.Contains(action, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                action is null ? "plugin needs an action." : $"Unknown plugin action: '{action}'.",
                "Use one of: install, uninstall, list, status. E.g. 'aioffice plugin install --host all'.",
                candidates: Actions);
        }

        var scope = parsed.GetOption("scope") ?? "user";
        if (scope is not ("user" or "project"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs, $"Unknown scope: '{scope}'.",
                "Use --scope user (home dirs) or --scope project (workspace root).",
                candidates: ["user", "project"]);
        }

        var (hosts, explicitHosts) = ResolveHosts(parsed.GetOption("host"));
        var dryRun = parsed.HasFlag("dry-run");
        var force = parsed.HasFlag("force");

        var home = Home();
        var ws = parsed.GetOption("workspace")
            ?? Environment.GetEnvironmentVariable("AIOFFICE_WORKSPACE")
            ?? (scope == "project" ? Directory.GetCurrentDirectory() : home);
        ws = Path.GetFullPath(ws);

        var projectRoot = scope == "project"
            ? Path.GetFullPath(parsed.GetOption("workspace") ?? Directory.GetCurrentDirectory())
            : ws;

        // The recorded MCP command must be the aioffice binary itself. Only the
        // install action writes it, so resolution failures only matter there.
        var exe = TryResolveExe();
        if (action == "install" && exe is null)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "Cannot locate the aioffice executable to record in host configs.",
                "Run the installed 'aioffice' binary (npm/Homebrew), or set AIOFFICE_EXE to its absolute path, then re-run.");
        }

        var ctx = new InstallContext(action, exe, ws, scope, dryRun, force, explicitHosts, home, projectRoot);

        var results = new List<HostOutcome>();
        foreach (var host in hosts)
        {
            try
            {
                results.Add(host switch
                {
                    "claude" => RunClaude(ctx),
                    "codex" => RunCodex(ctx),
                    "opencode" => RunOpencode(ctx),
                    "tonobraid" => RunTonoBraid(ctx),
                    _ => throw new AiofficeException(ErrorCodes.InternalError, $"Unhandled host {host}.", "Report this bug."),
                });
            }
            catch (Exception ex)
            {
                // One host's bad/locked config must not abort the others under --host all.
                var message = ex is AiofficeException ax ? ax.Message : ex.Message;
                results.Add(HostOutcome.Errored(host, message));
            }
        }

        return Envelope.Ok(new
        {
            action,
            scope,
            dryRun,
            workspace = ws,
            exe,
            hosts = results,
        });
    }

    private static (IReadOnlyList<string> hosts, bool explicitHosts) ResolveHosts(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw) || raw == "all")
        {
            return (KnownHosts, false);
        }

        var requested = raw.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var resolved = new List<string>();
        foreach (var h in requested)
        {
            if (h == "all")
            {
                return (KnownHosts, false);
            }

            if (!KnownHosts.Contains(h, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs, $"Unknown host: '{h}'.",
                    "Use claude, codex, opencode, tonobraid, or all (comma-separate for several).",
                    candidates: [.. KnownHosts, "all"]);
            }

            if (!resolved.Contains(h))
            {
                resolved.Add(h);
            }
        }

        return (resolved, true);
    }

    // ----------------------------------------------------------- Claude Code

    private static HostOutcome RunClaude(InstallContext ctx)
    {
        var claudeHome = Path.Combine(ctx.Home, ".claude");
        var detected = Directory.Exists(claudeHome);
        if (ctx.Action is "install" && !detected && !ctx.Explicit)
        {
            return HostOutcome.NotDetected("claude");
        }

        // MCP: user scope -> ~/.claude.json top-level mcpServers (proven location);
        // project scope -> <root>/.mcp.json.
        var mcpPath = ctx.Scope == "project"
            ? Path.Combine(ctx.ProjectRoot, ".mcp.json")
            : Path.Combine(ctx.Home, ".claude.json");
        var skillPath = ctx.Scope == "project"
            ? Path.Combine(ctx.ProjectRoot, ".claude", "skills", PluginId, "SKILL.md")
            : Path.Combine(claudeHome, "skills", PluginId, "SKILL.md");
        var commandPath = ctx.Scope == "project"
            ? Path.Combine(ctx.ProjectRoot, ".claude", "commands", "aioffice.md")
            : Path.Combine(claudeHome, "commands", "aioffice.md");

        var surfaces = new List<SurfaceChange>();
        if (ctx.Action is "list" or "status")
        {
            surfaces.Add(StatusMapKey(mcpPath, "mcpServers", PluginId, "mcp"));
            surfaces.Add(StatusFile(skillPath, "skill"));
            surfaces.Add(StatusFile(commandPath, "command"));
            return HostOutcome.ForStatus("claude", detected, surfaces);
        }

        if (ctx.Action == "uninstall")
        {
            surfaces.Add(RemoveMapKey(mcpPath, "mcpServers", PluginId, "mcp", ctx.DryRun));
            surfaces.Add(RemoveDir(Path.GetDirectoryName(skillPath)!, "skill", ctx.DryRun));
            surfaces.Add(RemoveFile(commandPath, "command", ctx.DryRun));
            return HostOutcome.ForUninstall("claude", detected, surfaces);
        }

        var node = StdioMcpNode(ctx.Exe!, ctx.Ws);
        surfaces.Add(UpsertMapKey(mcpPath, "mcpServers", PluginId, node, "mcp", ctx.DryRun, ctx.Force));
        surfaces.Add(WriteTextFile(skillPath, SkillFile(), "skill", ctx.DryRun, ctx.Force));
        surfaces.Add(WriteTextFile(commandPath, CommandFile(includeName: true), "command", ctx.DryRun, ctx.Force));
        return HostOutcome.ForInstall("claude", detected, surfaces);
    }

    // ----------------------------------------------------------------- Codex

    private static HostOutcome RunCodex(InstallContext ctx)
    {
        var codexHome = Environment.GetEnvironmentVariable("CODEX_HOME") ?? Path.Combine(ctx.Home, ".codex");
        var detected = Directory.Exists(codexHome);
        if (ctx.Action is "install" && !detected && !ctx.Explicit)
        {
            return HostOutcome.NotDetected("codex");
        }

        var configPath = Path.Combine(codexHome, "config.toml");
        var skillPath = Path.Combine(codexHome, "skills", PluginId, "SKILL.md");
        var agentsPath = Path.Combine(codexHome, "AGENTS.md");

        var notes = new List<string>();
        if (ctx.Scope == "project")
        {
            notes.Add("Codex has no project scope; installed at user scope (~/.codex).");
        }

        var surfaces = new List<SurfaceChange>();
        if (ctx.Action is "list" or "status")
        {
            surfaces.Add(StatusCodexToml(configPath));
            surfaces.Add(StatusFile(skillPath, "skill"));
            surfaces.Add(StatusMarker(agentsPath, "guide"));
            return HostOutcome.ForStatus("codex", detected, surfaces, notes);
        }

        if (ctx.Action == "uninstall")
        {
            surfaces.Add(UninstallCodexToml(configPath, ctx.DryRun));
            surfaces.Add(RemoveDir(Path.GetDirectoryName(skillPath)!, "skill", ctx.DryRun));
            surfaces.Add(RemoveMarker(agentsPath, "guide", ctx.DryRun));
            return HostOutcome.ForUninstall("codex", detected, surfaces, notes);
        }

        surfaces.Add(InstallCodexToml(configPath, ctx.Exe!, ctx.Ws, ctx.DryRun, ctx.Force));
        surfaces.Add(WriteTextFile(skillPath, SkillFile(), "skill", ctx.DryRun, ctx.Force));
        surfaces.Add(UpsertMarker(agentsPath, GuidePointer, "guide", ctx.DryRun));
        return HostOutcome.ForInstall("codex", detected, surfaces, notes);
    }

    // -------------------------------------------------------------- opencode

    private static HostOutcome RunOpencode(InstallContext ctx)
    {
        var configHome = Environment.GetEnvironmentVariable("XDG_CONFIG_HOME") ?? Path.Combine(ctx.Home, ".config");
        var ocHome = Path.Combine(configHome, "opencode");
        var detected = Directory.Exists(ocHome);
        if (ctx.Action is "install" && !detected && !ctx.Explicit)
        {
            return HostOutcome.NotDetected("opencode");
        }

        var mcpPath = ctx.Scope == "project"
            ? Path.Combine(ctx.ProjectRoot, "opencode.json")
            : Path.Combine(ocHome, "opencode.json");
        var agentPath = ctx.Scope == "project"
            ? Path.Combine(ctx.ProjectRoot, ".opencode", "agents", "aioffice.md")
            : Path.Combine(ocHome, "agents", "aioffice.md");
        var commandPath = ctx.Scope == "project"
            ? Path.Combine(ctx.ProjectRoot, ".opencode", "command", "aioffice.md")
            : Path.Combine(ocHome, "command", "aioffice.md");
        var agentsMdPath = ctx.Scope == "project"
            ? Path.Combine(ctx.ProjectRoot, "AGENTS.md")
            : Path.Combine(ocHome, "AGENTS.md");

        var surfaces = new List<SurfaceChange>();
        if (ctx.Action is "list" or "status")
        {
            surfaces.Add(StatusMapKey(mcpPath, "mcp", PluginId, "mcp"));
            surfaces.Add(StatusFile(agentPath, "agent"));
            surfaces.Add(StatusFile(commandPath, "command"));
            surfaces.Add(StatusMarker(agentsMdPath, "guide"));
            return HostOutcome.ForStatus("opencode", detected, surfaces);
        }

        if (ctx.Action == "uninstall")
        {
            surfaces.Add(RemoveMapKey(mcpPath, "mcp", PluginId, "mcp", ctx.DryRun));
            surfaces.Add(RemoveFile(agentPath, "agent", ctx.DryRun));
            surfaces.Add(RemoveFile(commandPath, "command", ctx.DryRun));
            surfaces.Add(RemoveMarker(agentsMdPath, "guide", ctx.DryRun));
            return HostOutcome.ForUninstall("opencode", detected, surfaces);
        }

        var node = OpencodeMcpNode(ctx.Exe!, ctx.Ws);
        surfaces.Add(UpsertMapKey(mcpPath, "mcp", PluginId, node, "mcp", ctx.DryRun, ctx.Force, schema: "https://opencode.ai/config.json"));
        surfaces.Add(WriteTextFile(agentPath, AgentFile(), "agent", ctx.DryRun, ctx.Force));
        surfaces.Add(WriteTextFile(commandPath, CommandFile(includeName: false), "command", ctx.DryRun, ctx.Force));
        surfaces.Add(UpsertMarker(agentsMdPath, GuidePointer, "guide", ctx.DryRun));
        return HostOutcome.ForInstall("opencode", detected, surfaces);
    }

    // ------------------------------------------------------------- TonoBraid

    private static HostOutcome RunTonoBraid(InstallContext ctx)
    {
        var tonoDir = Path.Combine(ctx.Home, ".tonoagent");
        var detected = Directory.Exists(tonoDir);
        if (ctx.Action is "install" && !detected && !ctx.Explicit)
        {
            return HostOutcome.NotDetected("tonobraid");
        }

        var pluginDir = Path.Combine(tonoDir, "plugins", PluginId);
        var pluginsJson = Path.Combine(tonoDir, "plugins.json");

        var notes = new List<string>();
        if (ctx.Scope == "project")
        {
            notes.Add("TonoBraid has no project scope; installed at user scope (~/.tonoagent).");
        }

        var surfaces = new List<SurfaceChange>();
        if (ctx.Action is "list" or "status")
        {
            surfaces.Add(new SurfaceChange("plugin", Directory.Exists(pluginDir) ? "present" : "absent", pluginDir));
            surfaces.Add(StatusRegistry(pluginsJson));
            return HostOutcome.ForStatus("tonobraid", detected, surfaces, notes);
        }

        if (ctx.Action == "uninstall")
        {
            surfaces.Add(RemoveDir(pluginDir, "plugin", ctx.DryRun));
            surfaces.Add(RemoveRegistryEntry(pluginsJson, ctx.DryRun));
            return HostOutcome.ForUninstall("tonobraid", detected, surfaces, notes);
        }

        // install
        var exists = Directory.Exists(pluginDir) && RegistryHasEntry(pluginsJson);
        if (exists && !ctx.Force)
        {
            surfaces.Add(new SurfaceChange("plugin", "skipped:exists", pluginDir));
            surfaces.Add(new SurfaceChange("registry", "skipped:exists", pluginsJson));
        }
        else if (ctx.DryRun)
        {
            surfaces.Add(new SurfaceChange("plugin", "would-write", pluginDir));
            surfaces.Add(new SurfaceChange("registry", "would-write", pluginsJson));
        }
        else
        {
            if (Directory.Exists(pluginDir))
            {
                Directory.Delete(pluginDir, recursive: true);
            }

            WritePluginTree(pluginDir, ctx.Exe!, ctx.Ws);
            var digest = ComputeTreeDigest(pluginDir);
            UpsertRegistry(pluginsJson, pluginDir, digest);
            surfaces.Add(new SurfaceChange("plugin", "written", pluginDir));
            surfaces.Add(new SurfaceChange("registry", "written", pluginsJson, $"trusted, digest {digest[..12]}"));
        }

        // Heads-up about the independent flat mcp.json registration path.
        var mcpJson = Path.Combine(tonoDir, "mcp.json");
        if (File.Exists(mcpJson))
        {
            try
            {
                var m = ReadJsonObject(mcpJson);
                if (m["mcpServers"] is JsonArray servers &&
                    servers.Any(s => s?["name"]?.GetValue<string>() == PluginId))
                {
                    notes.Add("~/.tonoagent/mcp.json already registers aioffice (flat). The plugin adds the namespaced " +
                              "aioffice.aioffice server; remove the mcp.json entry to avoid a duplicate.");
                }
            }
            catch
            {
                // best-effort heads-up only
            }
        }

        notes.Add("Restart TonoBraid to activate the plugin's MCP server.");
        return HostOutcome.ForInstall("tonobraid", detected, surfaces, notes);
    }

    private static void WritePluginTree(string pluginDir, string exe, string ws)
    {
        var manifest = new JsonObject
        {
            ["name"] = PluginId,
            ["version"] = Meta.ToolVersion,
            ["description"] = "Create/edit polished .docx/.xlsx/.pptx via the aioffice MCP server — design method + render→look→fix.",
            ["author"] = new JsonObject { ["name"] = "aioffice" },
            ["skills"] = new JsonArray("./skills/aioffice/SKILL.md"),
            ["commands"] = new JsonArray("./commands/aioffice.md"),
        };
        WriteText(Path.Combine(pluginDir, ".claude-plugin", "plugin.json"), Pretty(manifest));

        var mcp = new JsonObject
        {
            ["mcpServers"] = new JsonObject { [PluginId] = StdioMcpNode(exe, ws) },
        };
        WriteText(Path.Combine(pluginDir, ".mcp.json"), Pretty(mcp));

        WriteText(Path.Combine(pluginDir, "skills", PluginId, "SKILL.md"), SkillFile());
        WriteText(Path.Combine(pluginDir, "commands", "aioffice.md"), CommandFile(includeName: true));
    }

    // ----------------------------------------------------- payload synthesis

    private static string Body() => LoadPayload("AGENT-GUIDE.md").Replace("{{AIOFFICE_VERSION}}", Meta.ToolVersion);

    private static string SkillFile() =>
        $"---\nname: {PluginId}\ndescription: >-\n  {SkillDescription}\n---\n\n{Body()}";

    private static string AgentFile() =>
        $"---\ndescription: >-\n  {AgentDescription}\nmode: subagent\n---\n\n{Body()}";

    private static string CommandFile(bool includeName)
    {
        var name = includeName ? $"name: {PluginId}\n" : string.Empty;
        var command = LoadPayload("command-aioffice.md");
        return $"---\n{name}description: >-\n  {CommandDescription}\n---\n\n{command}";
    }

    private static string LoadPayload(string fileName)
    {
        var assembly = typeof(PluginVerbs).Assembly;
        var suffix = $".Plugin.{fileName}";
        var resourceName = assembly.GetManifestResourceNames()
            .FirstOrDefault(n => n.EndsWith(suffix, StringComparison.Ordinal));
        if (resourceName is null)
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                $"Embedded plugin payload missing: {fileName}.",
                "This is a packaging bug in aioffice; reinstall the binary.");
        }

        using var stream = assembly.GetManifestResourceStream(resourceName)!;
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    // ----------------------------------------------------------- MCP nodes

    private static JsonObject StdioMcpNode(string exe, string ws) => new()
    {
        ["type"] = "stdio",
        ["command"] = exe,
        ["args"] = new JsonArray("mcp", "--workspace", ws),
        ["env"] = new JsonObject { ["AIOFFICE_WORKSPACE"] = ws },
    };

    private static JsonObject OpencodeMcpNode(string exe, string ws) => new()
    {
        ["type"] = "local",
        ["command"] = new JsonArray(exe, "mcp", "--workspace", ws),
        ["enabled"] = true,
    };

    // -------------------------------------------------- JSON map-key surfaces

    private static SurfaceChange UpsertMapKey(
        string path, string mapKey, string entryKey, JsonObject value, string surface,
        bool dryRun, bool force, string? schema = null)
    {
        var root = ReadJsonObject(path);
        if (root[mapKey] is not JsonObject map)
        {
            map = new JsonObject();
            root[mapKey] = map;
        }

        var exists = map.ContainsKey(entryKey);
        if (exists && !force)
        {
            return new SurfaceChange(surface, dryRun ? "would-skip:exists" : "skipped:exists", path);
        }

        if (dryRun)
        {
            return new SurfaceChange(surface, "would-write", path);
        }

        if (schema is not null && root["$schema"] is null)
        {
            // keep $schema as the first key
            var rebuilt = new JsonObject { ["$schema"] = schema };
            foreach (var kv in root.ToList())
            {
                root.Remove(kv.Key);
                rebuilt[kv.Key] = kv.Value;
            }

            root = rebuilt;
            map = (JsonObject)root[mapKey]!;
        }

        map[entryKey] = value;
        WriteText(path, Pretty(root));
        return new SurfaceChange(surface, "written", path);
    }

    private static SurfaceChange RemoveMapKey(string path, string mapKey, string entryKey, string surface, bool dryRun)
    {
        if (!File.Exists(path))
        {
            return new SurfaceChange(surface, "absent", path);
        }

        var root = ReadJsonObject(path);
        if (root[mapKey] is not JsonObject map || !map.ContainsKey(entryKey))
        {
            return new SurfaceChange(surface, "absent", path);
        }

        if (dryRun)
        {
            return new SurfaceChange(surface, "would-remove", path);
        }

        map.Remove(entryKey);
        WriteText(path, Pretty(root));
        return new SurfaceChange(surface, "removed", path);
    }

    private static SurfaceChange StatusMapKey(string path, string mapKey, string entryKey, string surface)
    {
        var present = File.Exists(path)
            && ReadJsonObject(path)[mapKey] is JsonObject map
            && map.ContainsKey(entryKey);
        return new SurfaceChange(surface, present ? "present" : "absent", path);
    }

    // ----------------------------------------------------- text-file surfaces

    private static SurfaceChange WriteTextFile(string path, string content, string surface, bool dryRun, bool force)
    {
        var exists = File.Exists(path);
        if (exists && !force)
        {
            return new SurfaceChange(surface, dryRun ? "would-skip:exists" : "skipped:exists", path);
        }

        if (dryRun)
        {
            return new SurfaceChange(surface, "would-write", path);
        }

        WriteText(path, content);
        return new SurfaceChange(surface, "written", path);
    }

    private static SurfaceChange StatusFile(string path, string surface) =>
        new(surface, File.Exists(path) ? "present" : "absent", path);

    private static SurfaceChange RemoveFile(string path, string surface, bool dryRun)
    {
        if (!File.Exists(path))
        {
            return new SurfaceChange(surface, "absent", path);
        }

        if (dryRun)
        {
            return new SurfaceChange(surface, "would-remove", path);
        }

        File.Delete(path);
        return new SurfaceChange(surface, "removed", path);
    }

    private static SurfaceChange RemoveDir(string dir, string surface, bool dryRun)
    {
        if (!Directory.Exists(dir))
        {
            return new SurfaceChange(surface, "absent", dir);
        }

        if (dryRun)
        {
            return new SurfaceChange(surface, "would-remove", dir);
        }

        Directory.Delete(dir, recursive: true);
        return new SurfaceChange(surface, "removed", dir);
    }

    // --------------------------------------------------- marker-block surfaces

    private static SurfaceChange UpsertMarker(string path, string content, string surface, bool dryRun)
    {
        var existing = File.Exists(path) ? File.ReadAllText(path) : string.Empty;
        var block = $"{MarkerBegin}\n{content}\n{MarkerEnd}";
        var (text, existed) = ApplyMarker(existing, block);

        if (text == existing)
        {
            // block already present and identical — a real no-op, keep install idempotent
            return new SurfaceChange(surface, "skipped:exists", path);
        }

        if (dryRun)
        {
            return new SurfaceChange(surface, existed ? "would-update" : "would-write", path);
        }

        WriteText(path, text);
        return new SurfaceChange(surface, existed ? "updated" : "written", path);
    }

    private static (string text, bool existed) ApplyMarker(string existing, string block)
    {
        var begin = existing.IndexOf(MarkerBegin, StringComparison.Ordinal);
        if (begin >= 0)
        {
            var end = existing.IndexOf(MarkerEnd, begin, StringComparison.Ordinal);
            if (end >= 0)
            {
                end += MarkerEnd.Length;
                return (existing[..begin] + block + existing[end..], true);
            }
        }

        var prefix = existing.Length == 0
            ? string.Empty
            : existing.EndsWith("\n\n", StringComparison.Ordinal) ? string.Empty
            : existing.EndsWith("\n", StringComparison.Ordinal) ? "\n" : "\n\n";
        return (existing + prefix + block + "\n", false);
    }

    private static SurfaceChange RemoveMarker(string path, string surface, bool dryRun)
    {
        if (!File.Exists(path))
        {
            return new SurfaceChange(surface, "absent", path);
        }

        var existing = File.ReadAllText(path);
        var begin = existing.IndexOf(MarkerBegin, StringComparison.Ordinal);
        var end = begin >= 0 ? existing.IndexOf(MarkerEnd, begin, StringComparison.Ordinal) : -1;
        if (begin < 0 || end < 0)
        {
            return new SurfaceChange(surface, "absent", path);
        }

        if (dryRun)
        {
            return new SurfaceChange(surface, "would-remove", path);
        }

        end += MarkerEnd.Length;
        var after = existing[end..].TrimStart('\n');
        var before = existing[..begin].TrimEnd('\n');
        var result = before.Length == 0
            ? (after.Length == 0 ? string.Empty : after + "\n")
            : (after.Length == 0 ? before + "\n" : before + "\n\n" + after);
        WriteText(path, result);
        return new SurfaceChange(surface, "removed", path);
    }

    private static SurfaceChange StatusMarker(string path, string surface)
    {
        var present = File.Exists(path) && File.ReadAllText(path).Contains(MarkerBegin, StringComparison.Ordinal);
        return new SurfaceChange(surface, present ? "present" : "absent", path);
    }

    // -------------------------------------------------------- Codex TOML

    private static SurfaceChange InstallCodexToml(string path, string exe, string ws, bool dryRun, bool force)
    {
        var existing = File.Exists(path) ? File.ReadAllText(path) : string.Empty;
        var has = existing.Contains("[mcp_servers.aioffice]", StringComparison.Ordinal);
        if (has && !force)
        {
            return new SurfaceChange("mcp", dryRun ? "would-skip:exists" : "skipped:exists", path);
        }

        if (dryRun)
        {
            return new SurfaceChange("mcp", has ? "would-update" : "would-write", path);
        }

        if (File.Exists(path))
        {
            File.Copy(path, path + ".aioffice.bak", overwrite: true);
        }

        var baseText = has ? StripCodexBlock(existing) : existing;
        var sep = baseText.Length == 0
            ? string.Empty
            : baseText.EndsWith("\n\n", StringComparison.Ordinal) ? string.Empty
            : baseText.EndsWith("\n", StringComparison.Ordinal) ? "\n" : "\n\n";
        WriteText(path, baseText + sep + CodexBlock(exe, ws));
        return new SurfaceChange("mcp", has ? "updated" : "written", path);
    }

    private static SurfaceChange UninstallCodexToml(string path, bool dryRun)
    {
        if (!File.Exists(path))
        {
            return new SurfaceChange("mcp", "absent", path);
        }

        var existing = File.ReadAllText(path);
        if (!existing.Contains("[mcp_servers.aioffice]", StringComparison.Ordinal))
        {
            return new SurfaceChange("mcp", "absent", path);
        }

        if (dryRun)
        {
            return new SurfaceChange("mcp", "would-remove", path);
        }

        File.Copy(path, path + ".aioffice.bak", overwrite: true);
        WriteText(path, StripCodexBlock(existing));
        return new SurfaceChange("mcp", "removed", path);
    }

    private static SurfaceChange StatusCodexToml(string path)
    {
        var present = File.Exists(path) && File.ReadAllText(path).Contains("[mcp_servers.aioffice]", StringComparison.Ordinal);
        return new SurfaceChange("mcp", present ? "present" : "absent", path);
    }

    private static string CodexBlock(string exe, string ws)
    {
        var sb = new StringBuilder();
        sb.Append("[mcp_servers.aioffice]\n");
        sb.Append("command = ").Append(TomlStr(exe)).Append('\n');
        sb.Append("args = [").Append(TomlStr("mcp")).Append(", ").Append(TomlStr("--workspace")).Append(", ").Append(TomlStr(ws)).Append("]\n");
        sb.Append('\n');
        sb.Append("[mcp_servers.aioffice.env]\n");
        sb.Append("AIOFFICE_WORKSPACE = ").Append(TomlStr(ws)).Append('\n');
        return sb.ToString();
    }

    /// <summary>Removes the <c>[mcp_servers.aioffice]</c> table and its subtables, preserving everything else.</summary>
    private static string StripCodexBlock(string toml)
    {
        var lines = toml.Replace("\r\n", "\n").Split('\n');
        var output = new List<string>();
        var skipping = false;
        foreach (var line in lines)
        {
            var trimmed = line.TrimStart();
            if (trimmed.StartsWith('['))
            {
                skipping = trimmed.StartsWith("[mcp_servers.aioffice]", StringComparison.Ordinal)
                    || trimmed.StartsWith("[mcp_servers.aioffice.", StringComparison.Ordinal);
                if (skipping)
                {
                    continue;
                }
            }

            if (skipping)
            {
                continue;
            }

            output.Add(line);
        }

        return string.Join('\n', output).TrimEnd('\n') + "\n";
    }

    private static string TomlStr(string value) =>
        "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";

    // -------------------------------------------------- TonoBraid registry

    private static bool RegistryHasEntry(string pluginsJson)
    {
        if (!File.Exists(pluginsJson))
        {
            return false;
        }

        var root = ReadJsonObject(pluginsJson);
        return root["plugins"] is JsonArray arr
            && arr.Any(p => p?["id"]?.GetValue<string>() == PluginId);
    }

    private static SurfaceChange StatusRegistry(string pluginsJson) =>
        new("registry", RegistryHasEntry(pluginsJson) ? "present" : "absent", pluginsJson);

    private static void UpsertRegistry(string pluginsJson, string installPath, string digest)
    {
        var root = ReadJsonObject(pluginsJson);
        if (root["version"] is null)
        {
            root["version"] = 1;
        }

        if (root["plugins"] is not JsonArray arr)
        {
            arr = new JsonArray();
            root["plugins"] = arr;
        }

        for (var i = arr.Count - 1; i >= 0; i--)
        {
            if (arr[i]?["id"]?.GetValue<string>() == PluginId)
            {
                arr.RemoveAt(i);
            }
        }

        arr.Add(new JsonObject
        {
            ["id"] = PluginId,
            ["name"] = PluginId,
            ["format"] = "cc",
            ["sourceKind"] = "local",
            ["spec"] = installPath,
            ["installPath"] = installPath,
            ["version"] = Meta.ToolVersion,
            ["digest"] = digest,
            ["trust"] = "trusted",
            ["enabled"] = true,
            ["installedAt"] = DateTimeOffset.UtcNow.ToString("o"),
        });

        WriteText(pluginsJson, Pretty(root));
    }

    private static SurfaceChange RemoveRegistryEntry(string pluginsJson, bool dryRun)
    {
        if (!RegistryHasEntry(pluginsJson))
        {
            return new SurfaceChange("registry", "absent", pluginsJson);
        }

        if (dryRun)
        {
            return new SurfaceChange("registry", "would-remove", pluginsJson);
        }

        var root = ReadJsonObject(pluginsJson);
        if (root["plugins"] is JsonArray arr)
        {
            for (var i = arr.Count - 1; i >= 0; i--)
            {
                if (arr[i]?["id"]?.GetValue<string>() == PluginId)
                {
                    arr.RemoveAt(i);
                }
            }
        }

        WriteText(pluginsJson, Pretty(root));
        return new SurfaceChange("registry", "removed", pluginsJson);
    }

    /// <summary>
    /// SHA-256 of the installed tree, byte-for-byte matching TonoBraid's
    /// computeTreeDigest: forward-slash relative paths sorted lexicographically,
    /// each contributing <c>relPath \0 fileBytes \0</c>; .git and node_modules excluded.
    /// </summary>
    private static string ComputeTreeDigest(string root)
    {
        var files = Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)
            .Select(p => (Rel: Path.GetRelativePath(root, p).Replace('\\', '/'), Abs: p))
            .Where(t => !t.Rel.Split('/').Any(s => s is ".git" or "node_modules"))
            .OrderBy(t => t.Rel, StringComparer.Ordinal)
            .ToList();

        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var nul = new byte[] { 0 };
        foreach (var (rel, abs) in files)
        {
            hash.AppendData(Encoding.UTF8.GetBytes(rel));
            hash.AppendData(nul);
            try
            {
                hash.AppendData(File.ReadAllBytes(abs));
            }
            catch
            {
                // unreadable file contributes its path only, matching the reference
            }

            hash.AppendData(nul);
        }

        return Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant();
    }

    // ------------------------------------------------------------- IO helpers

    private static JsonObject ReadJsonObject(string path)
    {
        if (!File.Exists(path))
        {
            return new JsonObject();
        }

        var text = File.ReadAllText(path);
        if (string.IsNullOrWhiteSpace(text))
        {
            return new JsonObject();
        }

        JsonNode? node;
        try
        {
            node = JsonNode.Parse(text, documentOptions: new JsonDocumentOptions
            {
                AllowTrailingCommas = true,
                CommentHandling = JsonCommentHandling.Skip,
            });
        }
        catch (JsonException ex)
        {
            // Never overwrite a config we cannot parse — leave the file untouched.
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"{path} is not valid JSON, so aioffice won't edit it.",
                "Fix or remove the file, then re-run; aioffice never overwrites a config it can't parse.",
                innerException: ex);
        }

        if (node is not JsonObject obj)
        {
            // A non-object root (array / scalar) would be clobbered by a merge — refuse.
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"{path} is not a JSON object at the top level, so aioffice won't edit it.",
                "aioffice only merges into object-shaped config files; refusing to overwrite it.");
        }

        return obj;
    }

    private static string Pretty(JsonNode node) =>
        node.ToJsonString(new JsonSerializerOptions
        {
            WriteIndented = true,
            // Don't \uXXXX-escape non-ASCII when re-serializing the user's own
            // config files (~/.claude.json, opencode.json) — keep their content
            // byte-similar. The TonoBraid tree digest is computed over whatever
            // bytes we write, so this stays self-consistent.
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        });

    /// <summary>Creates parent dirs and writes atomically (temp + move).</summary>
    private static void WriteText(string path, string content)
    {
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir))
        {
            Directory.CreateDirectory(dir);
        }

        var tmp = path + ".aioffice.tmp";
        File.WriteAllText(tmp, content);
        File.Move(tmp, path, overwrite: true);
    }

    private static string Home() =>
        Environment.GetEnvironmentVariable("HOME")
        ?? Environment.GetEnvironmentVariable("USERPROFILE")
        ?? Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

    private static string? TryResolveExe()
    {
        var overridePath = Environment.GetEnvironmentVariable("AIOFFICE_EXE");
        if (!string.IsNullOrEmpty(overridePath) && File.Exists(overridePath))
        {
            return Path.GetFullPath(overridePath);
        }

        if (Environment.ProcessPath is { } pp &&
            Path.GetFileNameWithoutExtension(pp).Equals("aioffice", StringComparison.OrdinalIgnoreCase))
        {
            return pp;
        }

        var sibling = Path.Combine(AppContext.BaseDirectory, OperatingSystem.IsWindows() ? "aioffice.exe" : "aioffice");
        return File.Exists(sibling) ? sibling : null;
    }

    // --------------------------------------------------------------- records

    private sealed record InstallContext(
        string Action,
        string? Exe,
        string Ws,
        string Scope,
        bool DryRun,
        bool Force,
        bool Explicit,
        string Home,
        string ProjectRoot);
}

/// <summary>One install/uninstall/status change to a single host surface (mcp / skill / command / guide / plugin / registry).</summary>
internal sealed record SurfaceChange(string Surface, string Status, string? Path = null, string? Note = null);

/// <summary>The per-host result reported in the envelope data.</summary>
internal sealed record HostOutcome
{
    public required string Host { get; init; }

    public required bool Detected { get; init; }

    public required string Status { get; init; }

    public IReadOnlyList<SurfaceChange>? Surfaces { get; init; }

    public IReadOnlyList<string>? Notes { get; init; }

    public static HostOutcome NotDetected(string host) =>
        new() { Host = host, Detected = false, Status = "skipped:not-detected" };

    public static HostOutcome Errored(string host, string message) =>
        new() { Host = host, Detected = false, Status = "error", Notes = [message] };

    public static HostOutcome ForInstall(string host, bool detected, IReadOnlyList<SurfaceChange> surfaces, List<string>? notes = null) =>
        new()
        {
            Host = host,
            Detected = detected,
            Status = surfaces.Any(s => s.Status.StartsWith("would", StringComparison.Ordinal)) ? "dry-run"
                : surfaces.Any(s => s.Status is "written" or "updated") ? "installed"
                : "up-to-date",
            Surfaces = surfaces,
            Notes = NullIfEmpty(notes),
        };

    public static HostOutcome ForUninstall(string host, bool detected, IReadOnlyList<SurfaceChange> surfaces, List<string>? notes = null) =>
        new()
        {
            Host = host,
            Detected = detected,
            Status = surfaces.Any(s => s.Status.StartsWith("would", StringComparison.Ordinal)) ? "dry-run"
                : surfaces.Any(s => s.Status == "removed") ? "uninstalled"
                : "absent",
            Surfaces = surfaces,
            Notes = NullIfEmpty(notes),
        };

    public static HostOutcome ForStatus(string host, bool detected, IReadOnlyList<SurfaceChange> surfaces, List<string>? notes = null) =>
        new()
        {
            Host = host,
            Detected = detected,
            Status = surfaces.Any(s => s.Status == "present") ? "installed" : "absent",
            Surfaces = surfaces,
            Notes = NullIfEmpty(notes),
        };

    private static IReadOnlyList<string>? NullIfEmpty(List<string>? notes) =>
        notes is { Count: > 0 } ? notes : null;
}
