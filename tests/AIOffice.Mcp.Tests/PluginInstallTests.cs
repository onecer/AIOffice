using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Xunit;

namespace AIOffice.Mcp.Tests;

/// <summary>
/// Integration coverage for <c>aioffice plugin install/uninstall/status</c>.
/// Runs the real CLI binary (copied into this test's output dir via the
/// AIOffice.Cli project reference) in a throwaway <c>$HOME</c>, so the user's
/// live host configs are never touched and tests stay parallel-safe (the child
/// process gets its own environment — no global mutation). Guards the invariants
/// that matter most: existing host config is preserved, the TonoBraid plugin
/// digest matches TonoBraid's own algorithm (so trust stays "trusted"), install
/// is idempotent, and uninstall is surgical.
/// </summary>
public sealed class PluginInstallTests : IDisposable
{
    private readonly string _home;
    private readonly string _exe;
    private readonly string _dll = Path.Combine(AppContext.BaseDirectory, "aioffice.dll");

    public PluginInstallTests()
    {
        _home = Path.Combine(Path.GetTempPath(), "aioffice-plugin-it-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(_home, ".claude", "skills"));
        Directory.CreateDirectory(Path.Combine(_home, ".codex"));
        Directory.CreateDirectory(Path.Combine(_home, ".config", "opencode"));
        Directory.CreateDirectory(Path.Combine(_home, ".tonoagent"));

        // A fake binary the installer records as the MCP command (AIOFFICE_EXE).
        _exe = Path.Combine(_home, "bin", "aioffice");
        Directory.CreateDirectory(Path.GetDirectoryName(_exe)!);
        File.WriteAllText(_exe, "#!/bin/sh\n");
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_home, recursive: true);
        }
        catch
        {
            // best-effort temp cleanup
        }
    }

    [Fact]
    public void Install_registers_mcp_and_preserves_existing_servers()
    {
        // Seed ~/.claude.json with an unrelated server + a scalar key.
        File.WriteAllText(
            Path.Combine(_home, ".claude.json"),
            """{ "numStartups": 7, "mcpServers": { "gitnexus": { "command": "/x/gn", "args": ["mcp"] } } }""");

        var env = Run("plugin", "install", "--host", "claude");
        Assert.True(env.RootElement.GetProperty("ok").GetBoolean());

        using var claude = JsonDocument.Parse(File.ReadAllText(Path.Combine(_home, ".claude.json")));
        var servers = claude.RootElement.GetProperty("mcpServers");
        Assert.Equal(7, claude.RootElement.GetProperty("numStartups").GetInt32());
        Assert.True(servers.TryGetProperty("gitnexus", out _), "existing server must be preserved");
        var aioffice = servers.GetProperty("aioffice");
        Assert.Equal("stdio", aioffice.GetProperty("type").GetString());
        Assert.Equal(_exe, aioffice.GetProperty("command").GetString());

        Assert.True(File.Exists(Path.Combine(_home, ".claude", "skills", "aioffice", "SKILL.md")));
        Assert.True(File.Exists(Path.Combine(_home, ".claude", "commands", "aioffice.md")));
    }

    [Fact]
    public void TonoBraid_plugin_digest_matches_reference_algorithm()
    {
        var env = Run("plugin", "install", "--host", "tonobraid");
        Assert.True(env.RootElement.GetProperty("ok").GetBoolean());

        var pluginDir = Path.Combine(_home, ".tonoagent", "plugins", "aioffice");
        Assert.True(File.Exists(Path.Combine(pluginDir, ".claude-plugin", "plugin.json")));
        Assert.True(File.Exists(Path.Combine(pluginDir, ".mcp.json")));
        Assert.True(File.Exists(Path.Combine(pluginDir, "skills", "aioffice", "SKILL.md")));

        using var registry = JsonDocument.Parse(File.ReadAllText(Path.Combine(_home, ".tonoagent", "plugins.json")));
        var record = registry.RootElement.GetProperty("plugins").EnumerateArray()
            .Single(p => p.GetProperty("id").GetString() == "aioffice");
        Assert.Equal("cc", record.GetProperty("format").GetString());
        Assert.Equal("trusted", record.GetProperty("trust").GetString());

        // The recorded digest MUST equal what TonoBraid recomputes on load, else
        // it silently downgrades the plugin to "untrusted".
        Assert.Equal(ReferenceTreeDigest(pluginDir), record.GetProperty("digest").GetString());
    }

    [Fact]
    public void Install_is_idempotent_and_uninstall_is_surgical()
    {
        // Seed neighbors on every host to prove they survive uninstall.
        File.WriteAllText(
            Path.Combine(_home, ".claude.json"),
            """{ "mcpServers": { "gitnexus": { "command": "/x/gn" } } }""");
        File.WriteAllText(
            Path.Combine(_home, ".codex", "config.toml"),
            "[mcp_servers.pencil]\ncommand = \"/x/pencil\"\n");
        File.WriteAllText(
            Path.Combine(_home, ".config", "opencode", "opencode.json"),
            """{ "$schema": "https://opencode.ai/config.json", "mcp": { "pencil": { "type": "local", "command": ["/x/pencil"] } } }""");

        Run("plugin", "install", "--host", "all");

        // Second install is a no-op: every host reports up-to-date.
        var second = Run("plugin", "install", "--host", "all");
        foreach (var host in second.RootElement.GetProperty("data").GetProperty("hosts").EnumerateArray())
        {
            Assert.Equal("up-to-date", host.GetProperty("status").GetString());
        }

        Run("plugin", "uninstall", "--host", "all");

        using var claude = JsonDocument.Parse(File.ReadAllText(Path.Combine(_home, ".claude.json")));
        var servers = claude.RootElement.GetProperty("mcpServers");
        Assert.True(servers.TryGetProperty("gitnexus", out _), "neighbor preserved");
        Assert.False(servers.TryGetProperty("aioffice", out _), "aioffice removed");

        var codex = File.ReadAllText(Path.Combine(_home, ".codex", "config.toml"));
        Assert.Contains("[mcp_servers.pencil]", codex);
        Assert.DoesNotContain("aioffice", codex);

        using var oc = JsonDocument.Parse(File.ReadAllText(Path.Combine(_home, ".config", "opencode", "opencode.json")));
        var mcp = oc.RootElement.GetProperty("mcp");
        Assert.True(mcp.TryGetProperty("pencil", out _), "neighbor preserved");
        Assert.False(mcp.TryGetProperty("aioffice", out _), "aioffice removed");

        Assert.False(Directory.Exists(Path.Combine(_home, ".tonoagent", "plugins", "aioffice")));
        Assert.False(File.Exists(Path.Combine(_home, ".claude", "skills", "aioffice", "SKILL.md")));
    }

    [Fact]
    public void Dry_run_writes_nothing()
    {
        var env = Run("plugin", "install", "--host", "all", "--dry-run");
        Assert.True(env.RootElement.GetProperty("ok").GetBoolean());
        Assert.True(env.RootElement.GetProperty("data").GetProperty("dryRun").GetBoolean());

        Assert.False(File.Exists(Path.Combine(_home, ".claude.json")));
        Assert.False(File.Exists(Path.Combine(_home, ".codex", "config.toml")));
        Assert.False(Directory.Exists(Path.Combine(_home, ".tonoagent", "plugins", "aioffice")));
        Assert.False(File.Exists(Path.Combine(_home, ".claude", "skills", "aioffice", "SKILL.md")));
    }

    [Fact]
    public void Refuses_to_overwrite_an_unparseable_or_non_object_config()
    {
        // An array-root file (or any non-object / invalid JSON) must be left
        // byte-for-byte untouched, and the host reported as "error" — never
        // silently clobbered into an empty object.
        var ocPath = Path.Combine(_home, ".config", "opencode", "opencode.json");
        const string original = """[{"keep":"me"},{"and":"me"}]""";
        File.WriteAllText(ocPath, original);

        var env = Run("plugin", "install", "--host", "opencode");
        var host = env.RootElement.GetProperty("data").GetProperty("hosts").EnumerateArray().Single();
        Assert.Equal("error", host.GetProperty("status").GetString());
        Assert.Equal(original, File.ReadAllText(ocPath));
    }

    [Fact]
    public void One_bad_config_does_not_block_the_other_hosts()
    {
        // A corrupt ~/.claude.json must not abort the whole --host all run.
        File.WriteAllText(Path.Combine(_home, ".claude.json"), """{ "mcpServers": { "x": """);

        var env = Run("plugin", "install", "--host", "all");
        Assert.True(env.RootElement.GetProperty("ok").GetBoolean());

        var statuses = env.RootElement.GetProperty("data").GetProperty("hosts").EnumerateArray()
            .ToDictionary(h => h.GetProperty("host").GetString()!, h => h.GetProperty("status").GetString());
        Assert.Equal("error", statuses["claude"]);
        Assert.Equal("installed", statuses["codex"]);
        Assert.Equal("installed", statuses["tonobraid"]);
        Assert.True(Directory.Exists(Path.Combine(_home, ".tonoagent", "plugins", "aioffice")));
    }

    // ---- harness ----------------------------------------------------------

    private JsonDocument Run(params string[] args)
    {
        Assert.True(File.Exists(_dll), $"CLI binary not found next to tests: {_dll}");

        var psi = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
        };
        psi.ArgumentList.Add(_dll);
        foreach (var a in args)
        {
            psi.ArgumentList.Add(a);
        }

        // Isolated environment: redirect every host home into the temp dir.
        psi.Environment["HOME"] = _home;
        psi.Environment["USERPROFILE"] = _home;
        psi.Environment["CODEX_HOME"] = Path.Combine(_home, ".codex");
        psi.Environment.Remove("XDG_CONFIG_HOME");
        psi.Environment.Remove("AIOFFICE_WORKSPACE");
        psi.Environment["AIOFFICE_EXE"] = _exe;

        using var proc = Process.Start(psi)!;
        var stdout = proc.StandardOutput.ReadToEnd();
        var stderr = proc.StandardError.ReadToEnd();
        proc.WaitForExit(60_000);

        Assert.False(string.IsNullOrWhiteSpace(stdout), $"no envelope on stdout. stderr: {stderr}");
        return JsonDocument.Parse(stdout);
    }

    /// <summary>Independent re-implementation of TonoBraid's computeTreeDigest, for parity assertion.</summary>
    private static string ReferenceTreeDigest(string root)
    {
        var files = Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)
            .Select(p => (Rel: Path.GetRelativePath(root, p).Replace('\\', '/'), Abs: p))
            .Where(t => !t.Rel.Split('/').Any(s => s is ".git" or "node_modules"))
            .OrderBy(t => t.Rel, StringComparer.Ordinal);

        using var sha = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        foreach (var (rel, abs) in files)
        {
            sha.AppendData(Encoding.UTF8.GetBytes(rel));
            sha.AppendData([0]);
            sha.AppendData(File.ReadAllBytes(abs));
            sha.AppendData([0]);
        }

        return Convert.ToHexString(sha.GetHashAndReset()).ToLowerInvariant();
    }
}
